"""Workbook-level operations over a byte-preserving package.

Mirrors ``xlsxedit.Workbook`` where the released version has the API and
pre-implements the direction upstream has merged but not yet released:
worksheet-part duplication backs ``copy_worksheet`` (upstream commit
``3271d365``), and sheet-rename following in formulas extends
``rename_worksheet``. Sheet naming itself (the ``<sheet>`` elements in
``xl/workbook.xml``) stays with the caller, whose raw attribute text must
survive verbatim.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, BinaryIO

from xlsxedit.opc.constants import CT
from xlsxedit.opc.constants import RT as _RT

from _patched_xlsxedit._drawing import WSDR_CLOSE, WSDR_OPEN, anchor_xml
from _patched_xlsxedit._package import Package
from _patched_xlsxedit._remap import rename_sheets_in_formulas
from _patched_xlsxedit._xmltext import parse_attrs, unescape

if TYPE_CHECKING:
    from pathlib import Path

    from _patched_xlsxedit._content_types import ContentTypes
    from _patched_xlsxedit._drawing import PicturePlacement
    from _patched_xlsxedit._rels import Relationships


class RT(_RT):
    """``xlsxedit``'s relationship types, plus the ones it lacks."""

    COMMENTS = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    )


_WORKBOOK_MEMBER = "xl/workbook.xml"
_SHEET_TAG_RE = re.compile(r"<sheet((?:\s[^>]*)?)></sheet>")
_SHEET_MEMBER_RE = re.compile(r"^xl/worksheets/sheet(\d+)\.xml$")


@dataclass
class _SheetCounters:
    """Allocation baselines for new worksheet parts."""

    number: int
    sheet_id: int


class Workbook:
    """An open workbook whose parts are raw bytes."""

    def __init__(self, package: Package) -> None:
        """Wrap an open ``package``."""
        self._package = package
        self._workbook_rels: Relationships | None = None
        self._content_types: ContentTypes | None = None
        self._sheet_counters: _SheetCounters | None = None
        self._media: dict[bytes, str] = {}

    @classmethod
    def open(cls, source: str | Path | BinaryIO) -> Workbook:
        """Open an ``.xlsx`` file, path, or binary stream."""
        return cls(Package.open(source))

    def save(self, target: BinaryIO) -> None:
        """Flush pending bookkeeping edits and write the archive."""
        self.flush()
        self._package.save(target)

    def flush(self) -> None:
        """Write pending workbook-rels / content-types edits back."""
        if self._workbook_rels is not None:
            self._workbook_rels.save()
        if self._content_types is not None:
            self._content_types.save()

    @property
    def package(self) -> Package:
        """The underlying byte-preserving package."""
        return self._package

    def workbook_rels(self) -> Relationships:
        """The workbook's relationships, opened once per workbook."""
        if self._workbook_rels is None:
            self._workbook_rels = self._package.rels(_WORKBOOK_MEMBER)
        return self._workbook_rels

    def content_types(self) -> ContentTypes:
        """The content-types editor, opened once per workbook."""
        if self._content_types is None:
            self._content_types = self._package.content_types()
        return self._content_types

    def sheet_members(self) -> list[tuple[str, str]]:
        """Return ``(sheet name, member name)`` in workbook order.

        Parsed from the current ``xl/workbook.xml`` and workbook
        relationships bytes.
        """
        raw = self._package.read(_WORKBOOK_MEMBER).decode()
        targets = {
            rel.rId: rel.target_partname.membername for rel in self.workbook_rels()
        }
        out: list[tuple[str, str]] = []
        for tag in _SHEET_TAG_RE.findall(raw):
            attrs = dict(parse_attrs(tag))
            out.append((unescape(attrs["name"]), targets[attrs["r:id"]]))
        return out

    # ------------------------------------------------------- sheet structure

    def _counters(self) -> _SheetCounters:
        """Capture allocation baselines before any add or remove.

        Numbers of parts removed later in the same run are not reused,
        so allocations stay stable regardless of operation order.
        """
        if self._sheet_counters is None:
            number = max(
                (
                    int(m.group(1))
                    for name in self._package.member_names
                    if (m := _SHEET_MEMBER_RE.match(name))
                ),
                default=0,
            )
            raw = self._package.read(_WORKBOOK_MEMBER).decode()
            sheet_id = max(
                (
                    int(dict(parse_attrs(tag))["sheetId"])
                    for tag in _SHEET_TAG_RE.findall(raw)
                ),
                default=0,
            )
            self._sheet_counters = _SheetCounters(number, sheet_id)
        return self._sheet_counters

    def next_sheet_id(self) -> int:
        """Allocate the next unused ``sheetId``."""
        counters = self._counters()
        counters.sheet_id += 1
        return counters.sheet_id

    def add_worksheet_part(self, data: bytes) -> tuple[str, str, int]:
        """Register ``data`` as a new worksheet part.

        Allocates the part name, workbook ``rId``, and ``sheetId``, and
        registers the relationship and content-type override. Returns
        ``(member, rid, sheet_id)``; appending the ``<sheet>`` element is
        the caller's. This is the structural core of upstream's
        ``copy_worksheet``.
        """
        counters = self._counters()
        counters.number += 1
        member = f"xl/worksheets/sheet{counters.number}.xml"
        rid = self.workbook_rels().next_rId()
        sheet_id = self.next_sheet_id()
        self._package.write(member, data)
        self.workbook_rels().add_relationship(RT.WORKSHEET, member, rid)
        self.content_types().add_override(f"/{member}", CT.WORKSHEET)
        return member, rid, sheet_id

    def remove_worksheet_part(self, member: str, rid: str) -> None:
        """Drop a worksheet part with its relationship and override.

        The structural core of ``xlsxedit.Workbook.remove_worksheet``.
        """
        self._counters()
        self._package.remove(member)
        self.workbook_rels().remove(rid)
        self.content_types().remove_override(f"/{member}")

    # ------------------------------------------------------ pictures / media

    def add_media(self, data: bytes) -> str:
        """Store a PNG in ``xl/media``, deduplicated by content."""
        member = self._media.get(data)
        if member is None:
            member = self._package.next_partname("/xl/media/image%d.png").membername
            self._package.write(member, data)
            self._media[data] = member
            self.content_types().ensure_default("png", CT.PNG)
        return member

    def add_pictures(self, pictures: list[PicturePlacement]) -> str:
        """Create a drawing part anchoring ``pictures``; return its member.

        The byte-preserving counterpart of ``Worksheet.add_image``:
        media parts, the drawing part with its relationships, and the
        content-type override are all registered. Relating the drawing to
        its worksheet is the caller's.
        """
        drawing_member = self._package.next_partname(
            "/xl/drawings/drawing%d.xml"
        ).membername
        rels = self._package.rels(drawing_member)
        media_rids: dict[str, str] = {}
        anchors: list[str] = []
        for pic_id, image in enumerate(pictures, start=1):
            media_member = self.add_media(image.data)
            rid = media_rids.get(media_member)
            if rid is None:
                rid = f"rId{len(media_rids) + 1}"
                media_rids[media_member] = rid
                rels.add_relationship(RT.IMAGE, media_member, rid)
            anchors.append(anchor_xml(image, pic_id=pic_id, embed_rid=rid))
        self._package.write(
            drawing_member, (WSDR_OPEN + "".join(anchors) + WSDR_CLOSE).encode()
        )
        rels.save()
        self.content_types().add_override(f"/{drawing_member}", CT.DRAWING)
        return drawing_member

    # ------------------------------------------------------------- renames

    def propagate_sheet_renames(
        self, renames: list[tuple[str, str]], members: list[str]
    ) -> None:
        """Follow sheet renames through formulas.

        Rewrites sheet-qualified references in the ``<f>`` nodes of
        ``members`` and of the chart parts reachable through their
        drawings. The formula-following extension of
        ``xlsxedit.Workbook.rename_worksheet``.
        """
        renames = [(old, new) for old, new in renames if old != new]
        if not renames:
            return
        scan = list(members)
        for member in members:
            for rel in self._package.rels(member):
                if not rel.is_external and rel.reltype == RT.DRAWING:
                    drawing = rel.target_partname.membername
                    scan.extend(
                        r.target_partname.membername
                        for r in self._package.rels(drawing)
                        if not r.is_external and r.reltype == RT.CHART
                    )
        for member in scan:
            raw = self._package.read(member).decode()
            updated = rename_sheets_in_formulas(raw, renames)
            if updated != raw:
                self._package.write(member, updated.encode())
