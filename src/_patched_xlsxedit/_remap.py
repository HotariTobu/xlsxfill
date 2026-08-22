from __future__ import annotations

import re
from copy import deepcopy
from typing import TYPE_CHECKING, Protocol

from xlsxedit.merge import parse_range
from xlsxedit.oxml.address import col_to_index, index_to_col

if TYPE_CHECKING:
    from collections.abc import Callable, Iterable, Mapping

    from lxml.etree import _Element
    from xlsxedit.opc.part import Part
    from xlsxedit.workbook import Workbook

SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

_F = f"{{{SML_NS}}}f"
_MERGE_CELLS = f"{{{SML_NS}}}mergeCells"
_MERGE_CELL = f"{{{SML_NS}}}mergeCell"
_CONDITIONAL_FORMATTING = f"{{{SML_NS}}}conditionalFormatting"
_DATA_VALIDATION = f"{{{SML_NS}}}dataValidation"
_HYPERLINK = f"{{{SML_NS}}}hyperlink"
_AUTO_FILTER = f"{{{SML_NS}}}autoFilter"
_TABLE_COLUMNS = f"{{{SML_NS}}}tableColumns"
_TABLE_COLUMN = f"{{{SML_NS}}}tableColumn"
_DEFINED_NAMES = f"{{{SML_NS}}}definedNames"
_DEFINED_NAME = f"{{{SML_NS}}}definedName"

_TWO_CELL = f"{{{XDR_NS}}}twoCellAnchor"
_ONE_CELL = f"{{{XDR_NS}}}oneCellAnchor"
_FROM = f"{{{XDR_NS}}}from"
_TO = f"{{{XDR_NS}}}to"
_COL = f"{{{XDR_NS}}}col"
_ROW = f"{{{XDR_NS}}}row"
_CNVPR = f"{{{XDR_NS}}}cNvPr"


class PointMapper(Protocol):
    def __call__(self, coord: int, *, absolute: bool, end: bool) -> int: ...


class EndpointMapper(Protocol):
    def __call__(self, coord: int, *, end: bool) -> int: ...


_RANGE_RE = re.compile(
    r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?$",
)


def remap_range_ref(ref: str, map_row: EndpointMapper, map_col: EndpointMapper) -> str:
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
    return " ".join(remap_range_ref(part, map_row, map_col) for part in sqref.split())


def remap_sheet_references(
    ws_element: _Element,
    map_row: EndpointMapper,
    map_col: EndpointMapper,
) -> None:
    block = ws_element.find(_MERGE_CELLS)
    if block is not None:
        for merge in block.findall(_MERGE_CELL):
            ref = merge.get("ref")
            if ref:
                merge.set("ref", remap_range_ref(ref, map_row, map_col))
    for element in ws_element.iter(_CONDITIONAL_FORMATTING, _DATA_VALIDATION):
        sqref = element.get("sqref")
        if sqref:
            element.set("sqref", remap_sqref(sqref, map_row, map_col))
    for element in ws_element.iter(_HYPERLINK, _AUTO_FILTER):
        ref = element.get("ref")
        if ref:
            element.set("ref", remap_range_ref(ref, map_row, map_col))


def set_merged_refs(ws_element: _Element, refs: list[str]) -> None:
    block = ws_element.find(_MERGE_CELLS)
    if block is None:
        return
    if not refs:
        ws_element.remove(block)
        return
    for child in list(block):
        block.remove(child)
    for ref in refs:
        block.makeelement(_MERGE_CELL, {})
        merge = block.makeelement(_MERGE_CELL, {"ref": ref})
        block.append(merge)
    block.set("count", str(len(refs)))


def merged_refs(ws_element: _Element) -> list[str]:
    block = ws_element.find(_MERGE_CELLS)
    if block is None:
        return []
    return [m.get("ref", "") for m in block.findall(_MERGE_CELL) if m.get("ref")]


_DEFINED_NAME_REF_RE = re.compile(
    r"('(?:[^']|'')+'|[A-Za-z0-9_.]+)!"
    r"(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)",
)


def remap_defined_name(
    text: str,
    mappers_for_sheet: Mapping[str, tuple[EndpointMapper, EndpointMapper]],
) -> str:

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


def remap_defined_names(
    workbook_element: _Element,
    mappers_for_sheet: Mapping[str, tuple[EndpointMapper, EndpointMapper]],
) -> None:
    block = workbook_element.find(_DEFINED_NAMES)
    if block is None:
        return
    for element in block.findall(_DEFINED_NAME):
        text = element.text
        if not text:
            continue
        element.text = remap_defined_name(text, mappers_for_sheet)


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
    out: list[str] = []
    pos = 0
    while pos < len(text):
        char = text[pos]
        if char in ('"', "'"):
            end = text.find(char, pos + 1)
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


def rebase_formulas(
    row_element: _Element,
    map_col: PointMapper,
    map_row: PointMapper,
) -> None:
    for f_elm in row_element.iter(_F):
        text = f_elm.text
        if not text:
            continue
        f_elm.text = rebase_formula(text, map_col, map_row)


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


def remap_table(
    table_element: _Element,
    map_range: Callable[[str], str],
    header_text: Callable[[int, int], str],
) -> None:
    ref = table_element.get("ref")
    if not ref:
        return
    new_ref = map_range(ref)
    table_element.set("ref", new_ref)
    auto = table_element.find(_AUTO_FILTER)
    if auto is not None and auto.get("ref"):
        auto.set("ref", map_range(auto.get("ref", "")))
    columns = table_element.find(_TABLE_COLUMNS)
    if columns is None:
        return
    c1, r1, c2, _ = parse_range(new_ref)
    names = [
        header_text(r1, col) for col in range(col_to_index(c1), col_to_index(c2) + 1)
    ]
    for child in list(columns):
        columns.remove(child)
    columns.set("count", str(len(names)))
    for i, name in enumerate(names, start=1):
        column = columns.makeelement(_TABLE_COLUMN, {"id": str(i), "name": name})
        columns.append(column)


_TWO_CELL = f"{{{XDR_NS}}}twoCellAnchor"
_ONE_CELL = f"{{{XDR_NS}}}oneCellAnchor"
_FROM = f"{{{XDR_NS}}}from"
_TO = f"{{{XDR_NS}}}to"
_XDR_COL = f"{{{XDR_NS}}}col"
_XDR_ROW = f"{{{XDR_NS}}}row"
_CNVPR = f"{{{XDR_NS}}}cNvPr"


def expand_anchors(
    part: Part, placements: Callable[[int, int], list[tuple[int, int]]]
) -> None:
    from xlsxedit.oxml.parser import parse_xml, serialize_xml

    root = parse_xml(part.blob)
    changed = False
    for anchor in list(root):
        if anchor.tag not in (_TWO_CELL, _ONE_CELL):
            continue
        if anchor.get("editAs") == "absolute":
            continue
        corner = anchor.find(_FROM)
        if corner is None:
            continue
        column = int(corner.findtext(_XDR_COL, "0") or "0")
        row = int(corner.findtext(_XDR_ROW, "0") or "0") + 1
        targets = placements(row, column)
        if targets == [(row, column)]:
            continue
        changed = True
        index = list(root).index(anchor)
        if not targets:
            root.remove(anchor)
            continue
        _shift_anchor(anchor, targets[0][0] - row, targets[0][1] - column)
        for offset, (target_row, target_column) in enumerate(targets[1:], start=1):
            copy = deepcopy(anchor)
            _shift_anchor(
                copy, target_row - targets[0][0], target_column - targets[0][1]
            )
            root.insert(index + offset, copy)
    if changed:
        _renumber(root)
        part._blob = serialize_xml(root)


def _shift_anchor(anchor: _Element, rows: int, columns: int) -> None:
    if rows == 0 and columns == 0:
        return
    for corner in (anchor.find(_FROM), anchor.find(_TO)):
        if corner is None:
            continue
        column = corner.find(_XDR_COL)
        if column is not None:
            column.text = str(int(column.text or "0") + columns)
        row = corner.find(_XDR_ROW)
        if row is not None:
            row.text = str(int(row.text or "0") + rows)


def _renumber(root: _Element) -> None:
    names = list(root.iter(_CNVPR))
    if not names:
        return
    first = min(int(name.get("id", "0")) for name in names)
    for offset, name in enumerate(names):
        new_id = first + offset
        name.set("id", str(new_id))
        label = name.get("name")
        if label:
            name.set("name", re.sub(r"\d+$", str(new_id), label))


def quote_sheet_name(name: str) -> str:
    if name.replace("_", "").isalnum():
        return name
    return "'" + name.replace("'", "''") + "'"


def sheet_prefix_patterns(old: str) -> list[str]:
    patterns = [f"'{old.replace(chr(39), chr(39) * 2)}'!"]
    if quote_sheet_name(old) == old:
        patterns.append(f"{old}!")
    return patterns


def rename_in_formula_text(text: str, renames: list[tuple[str, str]]) -> str:
    for old, new in renames:
        if old == new:
            continue
        replacement = f"{quote_sheet_name(new)}!"
        for pattern in sheet_prefix_patterns(old):
            text = text.replace(pattern, replacement)
    return text


def rename_sheets_in_formulas(
    elements: Iterable[_Element],
    renames: list[tuple[str, str]],
) -> None:
    pairs = [(old, new) for old, new in renames if old != new]
    if not pairs:
        return
    for element in elements:
        for f_elm in element.iter(_F):
            text = f_elm.text
            if not text:
                continue
            f_elm.text = rename_in_formula_text(text, pairs)


def rename_sheets_in_charts(workbook: Workbook, renames: list[tuple[str, str]]) -> None:
    pairs = [(old, new) for old, new in renames if old != new]
    if not pairs:
        return
    from xlsxedit.drawing import drawing_parts_for_worksheet
    from xlsxedit.oxml.parser import parse_xml, serialize_xml

    chart_rel = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    )
    chart_formula = "{http://schemas.openxmlformats.org/drawingml/2006/chart}f"
    for worksheet in workbook.worksheets:
        for drawing in drawing_parts_for_worksheet(worksheet):
            for rel in drawing.rels:
                if rel.is_external or rel.reltype != chart_rel:
                    continue
                part = rel.target_part
                root = parse_xml(part.blob)
                touched = False
                for element in root.iter(chart_formula):
                    text = element.text
                    if not text:
                        continue
                    renamed = rename_in_formula_text(text, pairs)
                    if renamed != text:
                        element.text = renamed
                        touched = True
                if touched:
                    part._blob = serialize_xml(root)
