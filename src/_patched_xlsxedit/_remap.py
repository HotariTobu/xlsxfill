"""Reference following through arbitrary row/column mappings.

``xlsxedit.row_shift`` follows references (ranges, sqrefs, defined names,
table parts) through a uniform "insert n at row r" shift. These are the
same followers generalized to arbitrary endpoint mappings — a uniform
shift is the special case — plus formula rebasing, drawing-anchor
duplication, and sheet renames in formulas. The generalization is a
candidate upstream request.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Protocol

from xlsxedit.merge import parse_range
from xlsxedit.oxml.address import col_to_index, index_to_col

from _patched_xlsxedit._xmltext import escape_attr, escape_go_text, unescape

if TYPE_CHECKING:
    from collections.abc import Callable, Mapping


class PointMapper(Protocol):
    """Maps one template coordinate to the output coordinate."""

    def __call__(self, coord: int, *, absolute: bool, end: bool) -> int:
        """Map ``coord``; ``end`` marks the closing endpoint of a range."""
        ...


class EndpointMapper(Protocol):
    """Maps one range endpoint to the output coordinate."""

    def __call__(self, coord: int, *, end: bool) -> int:
        """Map ``coord``; ``end`` marks the closing endpoint of a range."""
        ...


# ------------------------------------------------------------- range refs

_RANGE_RE = re.compile(
    r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?$",
)


def remap_range_ref(ref: str, map_row: EndpointMapper, map_col: EndpointMapper) -> str:
    """Remap an A1-style range (``$`` markers preserved)."""
    m = _RANGE_RE.match(ref)
    if m is None:
        return ref
    c1a, c1, r1a, r1, c2a, c2, r2a, r2 = m.groups()
    single = c2 is None
    if single:
        c2a, c2, r2a, r2 = c1a, c1, r1a, r1
    new_c1 = index_to_col(map_col(col_to_index(c1), end=False))
    new_r1 = map_row(int(r1), end=False)
    new_c2 = index_to_col(map_col(col_to_index(c2), end=True))
    new_r2 = map_row(int(r2), end=True)
    first = f"{c1a}{new_c1}{r1a}{new_r1}"
    if single:
        return first
    return f"{first}:{c2a}{new_c2}{r2a}{new_r2}"


def remap_sqref(sqref: str, map_row: EndpointMapper, map_col: EndpointMapper) -> str:
    """Remap a space-separated list of ranges."""
    return " ".join(remap_range_ref(part, map_row, map_col) for part in sqref.split())


_DEFINED_NAME_REF_RE = re.compile(
    r"('(?:[^']|'')+'|[A-Za-z0-9_.]+)!"
    r"(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)",
)


def remap_defined_name(
    text: str,
    mappers_for_sheet: Mapping[str, tuple[EndpointMapper, EndpointMapper]],
) -> str:
    """Remap sheet-qualified ranges in a defined-name value.

    ``mappers_for_sheet`` maps a sheet name to its ``(map_row, map_col)``
    pair; sheets without an entry are left untouched.
    """

    def _sub(m: re.Match[str]) -> str:
        sheet_ref, ref = m.group(1), m.group(2)
        name = sheet_ref
        if name.startswith("'"):
            name = name[1:-1].replace("''", "'")
        mappers = mappers_for_sheet.get(name)
        if mappers is None:
            return m.group(0)
        return f"{sheet_ref}!{remap_range_ref(ref, mappers[0], mappers[1])}"

    return _DEFINED_NAME_REF_RE.sub(_sub, text)


# -------------------------------------------------------- formula rebasing

_REF_RE = re.compile(
    r"""
    (?P<ref>
        (?P<c1>\$?)(?P<col1>[A-Za-z]{1,3})(?P<r1>\$?)(?P<row1>\d+)
        (?: : (?P<c2>\$?)(?P<col2>[A-Za-z]{1,3})(?P<r2>\$?)(?P<row2>\d+))?
    )
    """,
    re.VERBOSE,
)
_BOUNDARY_BEFORE = set(
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_.$!'\""
)
_BOUNDARY_AFTER = set(
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_.("
)


def rebase_formula(text: str, map_col: PointMapper, map_row: PointMapper) -> str:
    """Rewrite the cell references of one formula.

    Double-quoted strings and sheet-qualified references are left as they
    are; every bare A1-style reference or range is mapped through
    ``map_col`` / ``map_row``.
    """
    out: list[str] = []
    pos = 0
    while pos < len(text):
        char = text[pos]
        if char == '"':
            end = text.find('"', pos + 1)
            end = len(text) - 1 if end == -1 else end
            out.append(text[pos : end + 1])
            pos = end + 1
            continue
        if char == "'":
            end = text.find("'", pos + 1)
            end = len(text) - 1 if end == -1 else end
            out.append(text[pos : end + 1])
            pos = end + 1
            continue
        m = _REF_RE.match(text, pos)
        if m is not None and _is_bare(text, m):
            out.append(_map_match(m, map_col, map_row))
            pos = m.end()
            continue
        out.append(char)
        pos += 1
    return "".join(out)


def _is_bare(text: str, m: re.Match[str]) -> bool:
    before = text[m.start() - 1] if m.start() > 0 else ""
    after = text[m.end()] if m.end() < len(text) else ""
    return before not in _BOUNDARY_BEFORE and after not in _BOUNDARY_AFTER


def _map_match(m: re.Match[str], map_col: PointMapper, map_row: PointMapper) -> str:
    is_range = m.group("col2") is not None
    first = _map_single(
        m.group("c1"),
        m.group("col1"),
        m.group("r1"),
        m.group("row1"),
        map_col,
        map_row,
        end=False,
    )
    if not is_range:
        return first
    second = _map_single(
        m.group("c2"),
        m.group("col2"),
        m.group("r2"),
        m.group("row2"),
        map_col,
        map_row,
        end=True,
    )
    return f"{first}:{second}"


def _map_single(
    col_abs: str,
    col: str,
    row_abs: str,
    row: str,
    map_col: PointMapper,
    map_row: PointMapper,
    *,
    end: bool,
) -> str:
    new_col = map_col(col_to_index(col), absolute=col_abs == "$", end=end)
    new_row = map_row(int(row), absolute=row_abs == "$", end=end)
    return f"{col_abs}{index_to_col(new_col)}{row_abs}{new_row}"


# ------------------------------------------------------------ table parts


def remap_table(
    raw: str,
    map_range: Callable[[str], str],
    header_text: Callable[[int, int], str],
) -> str | None:
    """Follow a table part through a range mapping.

    The table and auto-filter refs go through ``map_range``; the columns
    are rebuilt from ``header_text(row, col)`` over the new header row.
    Returns ``None`` when the part has no table ref to follow.
    """
    table_tag = re.search(r'(<table [^>]*?ref=")([^"]*)(")', raw)
    if table_tag is None:
        return None
    new_ref = map_range(table_tag.group(2))
    raw = raw.replace(
        table_tag.group(0),
        f"{table_tag.group(1)}{new_ref}{table_tag.group(3)}",
        1,
    )
    raw = re.sub(
        r'(<autoFilter ref=")([^"]*)(")',
        lambda m: f"{m.group(1)}{map_range(m.group(2))}{m.group(3)}",
        raw,
        count=1,
    )
    c1, r1, c2, _ = parse_range(new_ref)
    names = [
        header_text(r1, col) for col in range(col_to_index(c1), col_to_index(c2) + 1)
    ]
    columns = "".join(
        f'<tableColumn id="{i}" name="{escape_attr(name)}"></tableColumn>'
        for i, name in enumerate(names, start=1)
    )
    return re.sub(
        r'<tableColumns count="\d+">.*?</tableColumns>',
        f'<tableColumns count="{len(names)}">{columns}</tableColumns>',
        raw,
        count=1,
        flags=re.DOTALL,
    )


# --------------------------------------------------------- drawing anchors

_ANCHOR_RE = re.compile(
    r"<xdr:(twoCellAnchor|oneCellAnchor)((?:\s[^>]*)?)>.*?</xdr:\1>",
    re.DOTALL,
)
_CNVPR_RE = re.compile(r'(<xdr:cNvPr id=")(\d+)(" name=")([^"]*)(")')
_FROM_TO_RE = re.compile(
    r"(<xdr:(?:from|to)>)(.*?)(</xdr:(?:from|to)>)",
    re.DOTALL,
)


def _shift_anchor(block: str, delta_row: int, delta_col: int) -> str:
    if delta_row == 0 and delta_col == 0:
        return block

    def _shift(m: re.Match[str]) -> str:
        inner = m.group(2)
        inner = re.sub(
            r"(<xdr:col>)(\d+)(</xdr:col>)",
            lambda c: f"{c.group(1)}{int(c.group(2)) + delta_col}{c.group(3)}",
            inner,
        )
        inner = re.sub(
            r"(<xdr:row>)(\d+)(</xdr:row>)",
            lambda c: f"{c.group(1)}{int(c.group(2)) + delta_row}{c.group(3)}",
            inner,
        )
        return f"{m.group(1)}{inner}{m.group(3)}"

    return _FROM_TO_RE.sub(_shift, block)


def _renumber_anchor(block: str, new_id: int) -> str:
    def _sub(m: re.Match[str]) -> str:
        name = re.sub(r"\d+$", str(new_id), m.group(4))
        return f"{m.group(1)}{new_id}{m.group(3)}{name}{m.group(5)}"

    return _CNVPR_RE.sub(_sub, block, count=1)


def expand_anchors(
    raw: str,
    placements: Callable[[int, int], list[tuple[int, int]]],
) -> str:
    """Duplicate cell-anchored shapes to their mapped placements.

    ``placements(row, col)`` receives a shape's 1-based anchor row and
    0-based anchor column and returns the output ``(row, col)`` positions
    in duplication order; an empty list drops the shape. Anchors with
    ``editAs="absolute"`` are left untouched.
    """
    max_id = max(
        (int(n) for _, n, *_ in (m.groups() for m in _CNVPR_RE.finditer(raw))),
        default=0,
    )
    out: list[str] = []
    last_end = 0
    next_id = max_id + 1
    for m in _ANCHOR_RE.finditer(raw):
        out.append(raw[last_end : m.start()])
        last_end = m.end()
        block = m.group(0)
        if 'editAs="absolute"' in m.group(2):
            out.append(block)
            continue
        from_m = re.search(
            r"<xdr:from><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>",
            block,
            re.DOTALL,
        )
        if from_m is None:
            out.append(block)
            continue
        from_col = int(from_m.group(1))
        from_row = int(from_m.group(2)) + 1
        targets = placements(from_row, from_col)
        if not targets:
            continue
        for index, (new_row, new_col) in enumerate(targets):
            shifted = _shift_anchor(block, new_row - from_row, new_col - from_col)
            if index == 0:
                out.append(shifted)
            else:
                out.append(_renumber_anchor(shifted, next_id))
                next_id += 1
    out.append(raw[last_end:])
    return "".join(out)


# --------------------------------------------------- sheet renames (formulas)


def quote_sheet_name(name: str) -> str:
    """Quote a sheet name for use in a formula reference prefix."""
    if name.replace("_", "").isalnum():
        return name
    return "'" + name.replace("'", "''") + "'"


def sheet_prefix_patterns(old: str) -> list[str]:
    """The formula prefixes (decoded) that reference sheet ``old``."""
    patterns = [f"'{old.replace(chr(39), chr(39) * 2)}'!"]
    if quote_sheet_name(old) == old:
        patterns.append(f"{old}!")
    return patterns


def rename_in_formula_text(text: str, renames: list[tuple[str, str]]) -> str:
    """Apply sheet renames to one decoded formula string."""
    for old, new in renames:
        if old == new:
            continue
        replacement = f"{quote_sheet_name(new)}!"
        for pattern in sheet_prefix_patterns(old):
            text = text.replace(pattern, replacement)
    return text


_F_NODE_RE = re.compile(r"(<f(?:\s[^>]*)?>)(.*?)(</f>)", re.DOTALL)


def rename_sheets_in_formulas(raw: str, renames: list[tuple[str, str]]) -> str:
    """Apply sheet renames to every ``<f>`` node of one part's raw text."""

    def _sub(m: re.Match[str]) -> str:
        text = unescape(m.group(2))
        renamed = rename_in_formula_text(text, renames)
        if renamed == text:
            return m.group(0)
        return f"{m.group(1)}{escape_go_text(renamed)}{m.group(3)}"

    return _F_NODE_RE.sub(_sub, raw)
