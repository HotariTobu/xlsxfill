from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING

from xlsxedit.opc.constants import SML_NS
from xlsxedit.oxml.address import col_to_index

from _patched_xlsxedit._clone import cell_column, copy_columns, copy_rows
from _patched_xlsxedit._remap import (
    expand_anchors,
    merged_refs,
    rebase_formulas,
    remap_defined_names,
    remap_range_ref,
    remap_sheet_references,
    set_merged_refs,
)
from _patched_xlsxedit._tables import remap_tables

if TYPE_CHECKING:
    from lxml.etree import _Element
    from xlsxedit.workbook import Workbook
    from xlsxedit.worksheet import Worksheet

    from _patched_xlsxedit._remap import EndpointMapper, PointMapper

_ROW = f"{{{SML_NS}}}row"
_C = f"{{{SML_NS}}}c"
_SHEET_DATA = f"{{{SML_NS}}}sheetData"
_COLS = f"{{{SML_NS}}}cols"
_COL = f"{{{SML_NS}}}col"
_CONDITIONAL_FORMATTING = f"{{{SML_NS}}}conditionalFormatting"
_DATA_VALIDATIONS = f"{{{SML_NS}}}dataValidations"
_DATA_VALIDATION = f"{{{SML_NS}}}dataValidation"


class _Span:
    def __init__(self, lo: int, hi: int, copies: int) -> None:
        self.lo = lo
        self.hi = hi
        self.copies = copies

    @property
    def size(self) -> int:
        return self.hi - self.lo + 1

    @property
    def added(self) -> int:
        return self.copies * self.size

    def copy_of(self, out: int) -> int | None:
        if out < self.lo or out > self.hi + self.added:
            return None
        return (out - self.lo) // self.size

    def point(self, coord: int, *, at: int, absolute: bool) -> int:
        moved = coord if coord <= self.hi else coord + self.added
        if absolute:
            return moved
        copy = self.copy_of(at)
        return moved if copy is None else moved + copy * self.size

    def sources(self, last: int) -> list[int]:
        head = list(range(1, self.lo))
        block = list(range(self.lo, self.hi + 1))
        tail = list(range(self.hi + 1, last + 1))
        return head + block * (self.copies + 1) + tail

    def placements(self, coord: int) -> list[int]:
        if self.lo <= coord <= self.hi:
            return [coord + k * self.size for k in range(self.copies + 1)]
        return [self.point(coord, at=coord, absolute=True)]

    def intervals(self, lo: int, hi: int) -> list[tuple[int, int]]:
        out: list[tuple[int, int]] = []
        if lo < self.lo:
            out.append((lo, min(hi, self.lo - 1)))
        start, stop = max(lo, self.lo), min(hi, self.hi)
        if start <= stop:
            out.extend(
                (start + k * self.size, stop + k * self.size)
                for k in range(self.copies + 1)
            )
        if hi > self.hi:
            out.append((max(lo, self.hi + 1) + self.added, hi + self.added))
        return out


class _Gap:
    def __init__(self, lo: int, hi: int) -> None:
        self.lo = lo
        self.hi = hi

    @property
    def size(self) -> int:
        return self.hi - self.lo + 1

    def point(self, coord: int, *, end: bool) -> int:
        if coord < self.lo:
            return coord
        if coord > self.hi:
            return coord - self.size
        return self.lo - 1 if end else self.lo

    def sources(self, last: int) -> list[int]:
        return list(range(1, self.lo)) + list(range(self.hi + 1, last + 1))

    def placements(self, coord: int) -> list[int]:
        if self.lo <= coord <= self.hi:
            return []
        return [self.point(coord, end=False)]

    def intervals(self, lo: int, hi: int) -> list[tuple[int, int]]:
        out: list[tuple[int, int]] = []
        if lo < self.lo:
            out.append((lo, min(hi, self.lo - 1)))
        if hi > self.hi:
            out.append((max(lo, self.hi + 1) - self.size, hi - self.size))
        return out


type _Region = _Span | _Gap


def duplicate_rows(
    workbook: Workbook, worksheet: Worksheet, top: int, bottom: int, copies: int
) -> list[tuple[int, int]]:
    if copies <= 0:
        return []
    _apply(workbook, worksheet, _Span(top, bottom, copies), vertical=True)
    size = bottom - top + 1
    return [
        (bottom + (k - 1) * size + 1, bottom + k * size) for k in range(1, copies + 1)
    ]


def duplicate_columns(
    workbook: Workbook, worksheet: Worksheet, left: int, right: int, copies: int
) -> list[tuple[int, int]]:
    if copies <= 0:
        return []
    _apply(workbook, worksheet, _Span(left, right, copies), vertical=False)
    size = right - left + 1
    return [
        (right + (k - 1) * size + 1, right + k * size) for k in range(1, copies + 1)
    ]


def delete_rows(
    workbook: Workbook, worksheet: Worksheet, top: int, bottom: int
) -> None:
    _apply(workbook, worksheet, _Gap(top, bottom), vertical=True)


def delete_columns(
    workbook: Workbook, worksheet: Worksheet, left: int, right: int
) -> None:
    _apply(workbook, worksheet, _Gap(left, right), vertical=False)


def _sheet_data(worksheet: Worksheet) -> _Element:
    element = worksheet._part.element
    sheet_data = element.find(_SHEET_DATA)
    if sheet_data is None:
        message = "worksheet has no sheetData"
        raise ValueError(message)
    return sheet_data


def _last_line(sheet_data: _Element, *, vertical: bool) -> int:
    rows = sheet_data.findall(_ROW)
    if vertical:
        return max((int(r.get("r", "0")) for r in rows), default=0)
    return max(
        (cell_column(c) + 1 for r in rows for c in r.findall(_C)),
        default=0,
    )


def _apply(
    workbook: Workbook, worksheet: Worksheet, region: _Region, *, vertical: bool
) -> None:
    sheet_data = _sheet_data(worksheet)
    last = _last_line(sheet_data, vertical=vertical)
    sources = region.sources(max(last, region.hi))
    if vertical:
        copy_rows(sheet_data, sources, base=1)
    else:
        copy_columns(sheet_data, [c - 1 for c in sources], base=0)
        _rebuild_cols(worksheet._part.element, region)

    _follow(workbook, worksheet, region, vertical=vertical)
    worksheet._invalidate_merge_map()
    worksheet._invalidate_bulk_indexes()


def _rebuild_cols(ws_element: _Element, region: _Region) -> None:
    block = ws_element.find(_COLS)
    if block is None:
        return
    pieces: list[tuple[int, int, _Element]] = []
    for entry in list(block.findall(_COL)):
        block.remove(entry)
        lo = int(entry.get("min", "0"))
        hi = int(entry.get("max", "0"))
        pieces.extend(
            (start, stop, entry)
            for start, stop in region.intervals(lo, hi)
            if start <= stop
        )
    for start, stop, entry in _joined(sorted(pieces, key=lambda piece: piece[0])):
        moved = deepcopy(entry)
        moved.set("min", str(start))
        moved.set("max", str(stop))
        block.append(moved)
    if len(block) == 0:
        ws_element.remove(block)


def _joined(
    pieces: list[tuple[int, int, _Element]],
) -> list[tuple[int, int, _Element]]:
    out: list[tuple[int, int, _Element]] = []
    for start, stop, entry in pieces:
        if out and out[-1][1] + 1 == start and _same(out[-1][2], entry):
            out[-1] = (out[-1][0], stop, out[-1][2])
            continue
        out.append((start, stop, entry))
    return out


def _same(one: _Element, other: _Element) -> bool:
    keep = ("min", "max")
    return {k: v for k, v in one.attrib.items() if k not in keep} == {
        k: v for k, v in other.attrib.items() if k not in keep
    }


def _endpoint(region: _Region, *, offset: int = 0) -> EndpointMapper:

    def mapper(coord: int, *, end: bool) -> int:
        line = coord + offset
        if isinstance(region, _Gap):
            return region.point(line, end=end) - offset
        return region.point(line, at=line, absolute=True) - offset

    return mapper


def _identity(coord: int, *, end: bool) -> int:  # noqa: ARG001 - EndpointMapper conformance
    return coord


def _point(region: _Region, *, at: int, offset: int = 0) -> PointMapper:

    def mapper(coord: int, *, absolute: bool, end: bool) -> int:
        line = coord + offset
        if isinstance(region, _Gap):
            return region.point(line, end=end) - offset
        return region.point(line, at=at, absolute=absolute) - offset

    return mapper


def _identity_point(coord: int, *, absolute: bool, end: bool) -> int:  # noqa: ARG001 - PointMapper conformance
    return coord


def _follow(
    workbook: Workbook, worksheet: Worksheet, region: _Region, *, vertical: bool
) -> None:
    element = worksheet._part.element
    map_row = _endpoint(region) if vertical else _identity
    map_col = _identity if vertical else _endpoint(region, offset=1)

    refs = merged_refs(element)
    sqrefs = [
        (marked, marked.get("sqref", ""))
        for marked in element.iter(_CONDITIONAL_FORMATTING, _DATA_VALIDATION)
    ]
    remap_sheet_references(element, map_row, map_col)
    if refs:
        set_merged_refs(
            element,
            [
                ref
                for old in refs
                for ref in _range_copies(old, region, vertical=vertical)
            ],
        )
    _copy_sqrefs(sqrefs, region, vertical=vertical)

    for row_elm in _sheet_data(worksheet).findall(_ROW):
        row_num = int(row_elm.get("r", "0"))
        for c_elm in row_elm.findall(_C):
            at = row_num if vertical else cell_column(c_elm) + 1
            point = _point(region, at=at, offset=0 if vertical else 1)
            rebase_formulas(
                c_elm,
                _identity_point if vertical else point,
                point if vertical else _identity_point,
            )

    remap_tables(worksheet, lambda ref: remap_range_ref(ref, map_row, map_col))
    remap_defined_names(
        workbook._workbook_part.element,
        {worksheet.name: (map_row, map_col)},
    )
    _follow_anchors(worksheet, region, vertical=vertical)


def _copy_sqrefs(
    marked: list[tuple[_Element, str]], region: _Region, *, vertical: bool
) -> None:
    for element, old in marked:
        ranges = [
            ref
            for part in old.split()
            for ref in _range_copies(part, region, vertical=vertical)
        ]
        if ranges:
            element.set("sqref", " ".join(ranges))
            continue
        _drop(element)


def _drop(element: _Element) -> None:
    parent = element.getparent()
    if parent is None:
        return
    parent.remove(element)
    if parent.tag != _DATA_VALIDATIONS:
        return
    if len(parent):
        parent.set("count", str(len(parent)))
        return
    block = parent.getparent()
    if block is not None:
        block.remove(parent)


def _range_copies(ref: str, region: _Region, *, vertical: bool) -> list[str]:
    from xlsxedit.merge import parse_range
    from xlsxedit.oxml.address import index_to_col

    c1, r1, c2, r2 = parse_range(ref)
    lo, hi = (r1, r2) if vertical else (col_to_index(c1) + 1, col_to_index(c2) + 1)
    if not (region.lo <= lo and hi <= region.hi):
        return [remap_range_ref(ref, *_pair(region, vertical=vertical))]
    starts = region.placements(lo)
    ends = region.placements(hi)
    out: list[str] = []
    for start, end in zip(starts, ends, strict=False):
        if vertical:
            out.append(_range(c1, start, c2, end))
        else:
            out.append(_range(index_to_col(start - 1), r1, index_to_col(end - 1), r2))
    return out


def _range(c1: str, r1: int, c2: str, r2: int) -> str:
    first = f"{c1}{r1}"
    second = f"{c2}{r2}"
    return first if first == second else f"{first}:{second}"


def _pair(region: _Region, *, vertical: bool) -> tuple[EndpointMapper, EndpointMapper]:
    if vertical:
        return (_endpoint(region), _identity)
    return (_identity, _endpoint(region, offset=1))


def _follow_anchors(worksheet: Worksheet, region: _Region, *, vertical: bool) -> None:
    from xlsxedit.drawing import drawing_parts_for_worksheet

    def placements(row: int, col: int) -> list[tuple[int, int]]:
        if vertical:
            return [(line, col) for line in region.placements(row)]
        return [(row, line - 1) for line in region.placements(col + 1)]

    for part in drawing_parts_for_worksheet(worksheet):
        expand_anchors(part, placements)
