"""Shared string table editing at the text level.

The byte-preserving counterpart of ``xlsxedit.shared_strings``. New
strings are appended at the end in first-use order and deduplicated
against every existing entry. Entries that lose all their references are
kept so no later index moves. ``count`` / ``uniqueCount`` are recomputed
from the final workbook when the table is emitted.
"""

from __future__ import annotations

import re

from _patched_xlsxedit._xmltext import escape_text, unescape

_SI_RE = re.compile(r"<si>.*?</si>", re.DOTALL)
_T_RE = re.compile(r"<t(?:\s[^>]*)?>(.*?)</t>", re.DOTALL)
_SST_TAG_RE = re.compile(r"<sst[^>]*>")


def _needs_preserve(text: str) -> bool:
    if not text:
        return False
    return text != text.strip() or "\n" in text


class SharedStringTable:
    """The mutable shared string table of one workbook."""

    def __init__(self, raw: bytes) -> None:
        """Parse the raw bytes of ``xl/sharedStrings.xml``."""
        self._raw = raw.decode()
        body_start = self._raw.index(">", self._raw.index("<sst")) + 1
        body_end = self._raw.rindex("</sst>")
        self._head = self._raw[:body_start]
        self._tail = self._raw[body_end:]
        self._entries: list[str] = _SI_RE.findall(
            self._raw[body_start:body_end],
        )
        self._index_by_text: dict[str, int] = {}
        for index in range(len(self._entries)):
            self._index_by_text.setdefault(self.text_of(index), index)
        # Entry texts stay parsed here, keyed by index, so raw is not rescanned.
        self._texts = [self.text_of(index) for index in range(len(self._entries))]

    def text_of(self, index: int) -> str:
        """Return the display text of entry ``index`` (plain or rich)."""
        return "".join(unescape(t) for t in _T_RE.findall(self._entries[index]))

    def text(self, index: int) -> str:
        """Return the cached display text of entry ``index``."""
        return self._texts[index]

    def index_for(self, text: str) -> int:
        """Return the index for ``text``, appending a new entry if needed."""
        existing = self._index_by_text.get(text)
        if existing is not None:
            return existing
        attr = ' xml:space="preserve"' if _needs_preserve(text) else ""
        self._entries.append(f"<si><t{attr}>{escape_text(text)}</t></si>")
        index = len(self._entries) - 1
        self._index_by_text[text] = index
        self._texts.append(text)
        return index

    def emit(self, reference_count: int) -> bytes:
        """Serialize the table with ``count`` set to ``reference_count``."""
        head = re.sub(
            r'((?<![A-Za-z])count=")\d*(")',
            lambda m: f"{m.group(1)}{reference_count}{m.group(2)}",
            self._head,
            count=1,
        )
        head = re.sub(
            r'(uniqueCount=")\d*(")',
            lambda m: f"{m.group(1)}{len(self._entries)}{m.group(2)}",
            head,
            count=1,
        )
        return (head + "".join(self._entries) + self._tail).encode()
