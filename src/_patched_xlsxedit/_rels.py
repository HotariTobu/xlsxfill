"""Byte-preserving OPC relationships.

``xlsxedit.opc.rel.Relationships`` models the same collection but
re-serializes it through lxml, which loses the source attribute order and
element form. This variant keeps the part text verbatim and splices edits
into it. The API mirrors the upstream class (``add_relationship`` /
``next_rId``) so it can be proposed to xlsxedit as-is; ``remove`` is an
addition upstream does not have yet.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from xlsxedit.opc.packuri import PackURI

from _patched_xlsxedit._xmltext import escape_attr, parse_attrs, unescape

if TYPE_CHECKING:
    from collections.abc import Iterator

    from _patched_xlsxedit._package import Package

_SKELETON = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    "<Relationships xmlns="
    '"http://schemas.openxmlformats.org/package/2006/relationships">'
    "</Relationships>"
)
_REL_TAG_RE = re.compile(r"<Relationship\s([^>]*?)/?>")
_RID_NUM_RE = re.compile(r"^rId(\d+)$")


class Relationship:
    """One relationship of a part, as parsed from the raw text."""

    def __init__(
        self,
        rId: str,  # noqa: N803 (mirrors xlsxedit.opc.rel.Relationship)
        reltype: str,
        target_ref: str,
        target_mode: str,
        base_uri: str,
    ) -> None:
        """Wrap one parsed ``<Relationship>`` element."""
        self.rId = rId
        self.reltype = reltype
        self.target_ref = target_ref
        self.target_mode = target_mode
        self._base_uri = base_uri

    @property
    def is_external(self) -> bool:
        """Whether the relationship targets outside the package."""
        return self.target_mode == "External"

    @property
    def target_partname(self) -> PackURI:
        """The absolute pack URI of an internal target."""
        if self.is_external:
            raise ValueError("External relationships have no target partname")
        return PackURI.from_rel_ref(self._base_uri, self.target_ref)


class Relationships:
    """The relationships of one part, edited at the text level.

    Iterating yields the parsed [`Relationship`][_patched_xlsxedit._rels.Relationship]
    entries; edits splice into the raw text and [`save`][_patched_xlsxedit._rels.Relationships.save]
    writes the text back only when it changed. A part without a
    relationships member starts from the canonical empty skeleton.
    """

    def __init__(self, pkg: Package, member: str) -> None:
        """Open the relationships of part ``member`` in ``pkg``."""
        source = PackURI(f"/{member}")
        self._pkg = pkg
        self._member = source.rels_uri.membername
        self._base_uri = source.baseURI
        if self._member in pkg:
            self._text = pkg.read(self._member).decode()
        else:
            self._text = _SKELETON
        self._original = self._text
        self._rels: list[Relationship] = []
        self._max_rid = 0
        for m in _REL_TAG_RE.finditer(self._text):
            attrs = {name: unescape(value) for name, value in parse_attrs(m.group(1))}
            rel = Relationship(
                attrs["Id"],
                attrs["Type"],
                attrs["Target"],
                attrs.get("TargetMode", "Internal"),
                self._base_uri,
            )
            self._rels.append(rel)
            self._bump_max(rel.rId)

    def _bump_max(self, rid: str) -> None:
        m = _RID_NUM_RE.match(rid)
        if m is not None:
            self._max_rid = max(self._max_rid, int(m.group(1)))

    def __iter__(self) -> Iterator[Relationship]:
        """Iterate over the current relationships."""
        return iter(list(self._rels))

    def next_rId(self) -> str:  # noqa: N802 (mirrors xlsxedit.opc.rel)
        """Allocate the next unused ``rId``.

        Ids removed during this editing session are not reused, so
        allocations stay stable regardless of interleaved removals.
        """
        self._max_rid += 1
        return f"rId{self._max_rid}"

    def add_relationship(
        self,
        reltype: str,
        target: str,
        rId: str,  # noqa: N803 (mirrors xlsxedit.opc.rel)
        target_mode: str = "Internal",
    ) -> Relationship:
        """Append a relationship to ``target``.

        ``target`` is a member name for internal relationships (the
        relative reference is derived from the source part) or the raw
        URL for ``target_mode="External"``.
        """
        if target_mode == "External":
            target_ref = target
        else:
            target_ref = PackURI(f"/{target}").relative_ref(self._base_uri)
        external = ' TargetMode="External"' if target_mode == "External" else ""
        element = (
            f'<Relationship Id="{rId}" Target="{escape_attr(target_ref)}"'
            f'{external} Type="{reltype}"></Relationship>'
        )
        self._text = self._text.replace(
            "</Relationships>", f"{element}</Relationships>", 1
        )
        rel = Relationship(rId, reltype, target_ref, target_mode, self._base_uri)
        self._rels.append(rel)
        self._bump_max(rId)
        return rel

    def remove(self, rid: str) -> None:
        """Delete the relationship with id ``rid``."""
        self._text = re.sub(
            rf'<Relationship Id="{re.escape(rid)}"[^>]*></Relationship>',
            "",
            self._text,
            count=1,
        )
        self._rels = [rel for rel in self._rels if rel.rId != rid]

    def save(self) -> None:
        """Write the relationships member back if it changed."""
        if self._text != self._original:
            self._pkg.write(self._member, self._text.encode())
            self._original = self._text
