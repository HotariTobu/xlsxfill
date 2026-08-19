"""Worksheet child-element insertion in schema order.

The byte-preserving counterpart of
``xlsxedit.hyperlinks.insert_worksheet_child``: a new child element is
placed before the first present element that follows it in the
``CT_Worksheet`` sequence.
"""

from __future__ import annotations

_CHILD_ORDER = (
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "drawing",
    "legacyDrawing",
    "picture",
    "oleObjects",
    "tableParts",
)


def _insert_before_any(raw: str, element: str, followers: tuple[str, ...]) -> str:
    position = min(
        (index for tag in followers if (index := raw.find(tag)) != -1),
        default=len(raw),
    )
    return raw[:position] + element + raw[position:]


def insert_worksheet_child(raw: str, element: str, tag: str) -> str:
    """Insert ``element`` (a ``<tag …>`` child) at its schema position."""
    index = _CHILD_ORDER.index(tag)
    followers = (
        *(f"<{name}" for name in _CHILD_ORDER[index + 1 :]),
        "</worksheet>",
    )
    return _insert_before_any(raw, element, followers)


def add_hyperlinks(raw: str, tags: list[str]) -> str:
    """Append ``<hyperlink>`` tags, creating the block if needed."""
    body = "".join(tags)
    if "</hyperlinks>" in raw:
        return raw.replace("</hyperlinks>", f"{body}</hyperlinks>", 1)
    return insert_worksheet_child(raw, f"<hyperlinks>{body}</hyperlinks>", "hyperlinks")
