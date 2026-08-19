"""Byte-preserving ``[Content_Types].xml`` editing.

xlsxedit manages content types implicitly inside its package layer; this
editor is the explicit, text-splicing counterpart. It is a candidate for
an upstream public API.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from _patched_xlsxedit._package import Package

_MEMBER = "[Content_Types].xml"


class ContentTypes:
    """The content-types part of a package, edited at the text level."""

    def __init__(self, pkg: Package) -> None:
        """Open the content-types part of ``pkg``."""
        self._pkg = pkg
        self._text = pkg.read(_MEMBER).decode()
        self._original = self._text

    def ensure_default(self, extension: str, content_type: str) -> None:
        """Add a ``<Default>`` for ``extension`` unless one exists."""
        if f'Extension="{extension}"' in self._text:
            return
        element = (
            f'<Default Extension="{extension}" ContentType="{content_type}"></Default>'
        )
        last = self._text.rindex("</Default>") + len("</Default>")
        self._text = self._text[:last] + element + self._text[last:]

    def add_override(self, partname: str, content_type: str) -> None:
        """Append an ``<Override>`` for the ``/``-prefixed ``partname``."""
        element = (
            f'<Override PartName="{partname}" ContentType="{content_type}"></Override>'
        )
        self._text = self._text.replace("</Types>", f"{element}</Types>", 1)

    def remove_override(self, partname: str) -> None:
        """Delete the ``<Override>`` for the ``/``-prefixed ``partname``."""
        self._text = re.sub(
            rf'<Override PartName="{re.escape(partname)}"[^>]*></Override>',
            "",
            self._text,
            count=1,
        )

    def save(self) -> None:
        """Write the content-types member back if it changed."""
        if self._text != self._original:
            self._pkg.write(_MEMBER, self._text.encode())
            self._original = self._text
