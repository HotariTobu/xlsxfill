"""Byte-preserving OPC package.

``xlsxedit.opc.package.OpcPackage`` parses every part into lxml trees and
re-serializes them on save, which does not round-trip authored XML
byte-for-byte. This ``Package`` keeps every member as its raw bytes:
a part serves its original bytes until they are explicitly replaced, and
saving writes the members back without reformatting anything.
"""

from __future__ import annotations

import re
import zipfile
from io import BytesIO
from pathlib import Path
from typing import BinaryIO

from xlsxedit.opc.packuri import PackURI

from _patched_xlsxedit._content_types import ContentTypes
from _patched_xlsxedit._rels import Relationships


def _read_source(source: str | Path | BinaryIO) -> bytes:
    if isinstance(source, str | Path):
        return Path(source).read_bytes()
    return source.read()


class Package:
    """An open ``.xlsx`` package whose members are raw bytes."""

    def __init__(self, members: dict[str, bytes]) -> None:
        """Wrap raw ZIP ``members``, keyed by member name."""
        self._members = members

    @classmethod
    def open(cls, source: str | Path | BinaryIO) -> Package:
        """Open an ``.xlsx`` file, path, or binary stream."""
        data = _read_source(source)
        with zipfile.ZipFile(BytesIO(data)) as zipf:
            members = {
                name: zipf.read(name)
                for name in zipf.namelist()
                if not name.endswith("/")
            }
        return cls(members)

    def save(self, target: BinaryIO) -> None:
        """Write the members as a ZIP archive, in their held order."""
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zipf:
            for name, data in self._members.items():
                zipf.writestr(name, data)

    @property
    def member_names(self) -> list[str]:
        """The current member names."""
        return list(self._members)

    def __contains__(self, name: str) -> bool:
        """Return whether member ``name`` exists."""
        return name in self._members

    def read(self, name: str) -> bytes:
        """Return the current bytes of member ``name``."""
        return self._members[name]

    def write(self, name: str, data: bytes) -> None:
        """Replace (or add) member ``name`` with ``data``."""
        self._members[name] = data

    def remove(self, name: str) -> None:
        """Delete member ``name``."""
        del self._members[name]

    def next_partname(self, tmpl: str) -> PackURI:
        """Allocate the next partname for a ``%d`` template.

        Mirrors ``xlsxedit``'s ``OpcPackage.next_partname``: ``tmpl`` is a
        ``/``-prefixed pack URI template such as
        ``"/xl/worksheets/sheet%d.xml"``.
        """
        pattern = re.compile(
            "^" + re.escape(tmpl.removeprefix("/")).replace("%d", r"(\d+)") + "$"
        )
        used = [int(m.group(1)) for name in self._members if (m := pattern.match(name))]
        return PackURI(tmpl % (max(used, default=0) + 1))

    def rels_member_for(self, name: str) -> str:
        """Return the relationships member name for member ``name``."""
        return PackURI(f"/{name}").rels_uri.membername

    def rels(self, name: str) -> Relationships:
        """Open the relationships of member ``name`` for reading or editing."""
        return Relationships(self, name)

    def content_types(self) -> ContentTypes:
        """Open ``[Content_Types].xml`` for editing."""
        return ContentTypes(self)
