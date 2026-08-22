from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING

from xlsxedit.oxml.address import (
    col_to_index,
    index_to_col,
    join_address,
    split_address,
)

if TYPE_CHECKING:
    from lxml.etree import _Element

SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_ROW = f"{{{SML_NS}}}row"
_C = f"{{{SML_NS}}}c"


def cell_column(c_elm: _Element) -> int:
    letters, _ = split_address(c_elm.get("r", "A1"))
    return col_to_index(letters)


def set_row_number(row_elm: _Element, num: int) -> None:
    row_elm.set("r", str(num))
    for c_elm in row_elm.findall(_C):
        letters, _ = split_address(c_elm.get("r", "A1"))
        c_elm.set("r", join_address(letters, num))


def set_cell_column(c_elm: _Element, col: int, row: int) -> None:
    c_elm.set("r", join_address(index_to_col(col), row))


def copy_rows(
    sheet_data: _Element,
    sources: list[int],
    base: int = 1,
) -> list[tuple[int, _Element]]:
    by_num = {int(r.get("r", "0")): r for r in sheet_data.findall(_ROW)}
    produced: list[tuple[int, _Element]] = []
    for child in list(sheet_data):
        sheet_data.remove(child)
    for index, src in enumerate(sources):
        source = by_num.get(src)
        if source is None:
            continue
        row_elm = deepcopy(source)
        set_row_number(row_elm, base + index)
        sheet_data.append(row_elm)
        produced.append((index, row_elm))
    return produced


def copy_columns(
    sheet_data: _Element,
    sources: list[int],
    base: int = 0,
) -> list[tuple[int, _Element]]:
    produced: list[tuple[int, _Element]] = []
    for row_elm in sheet_data.findall(_ROW):
        row_num = int(row_elm.get("r", "0"))
        by_col = {cell_column(c): c for c in row_elm.findall(_C)}
        for child in list(row_elm):
            if child.tag == _C:
                row_elm.remove(child)
        for index, src in enumerate(sources):
            source = by_col.get(src)
            if source is None:
                continue
            c_elm = deepcopy(source)
            set_cell_column(c_elm, base + index, row_num)
            row_elm.append(c_elm)
            produced.append((index, c_elm))
        if len(row_elm) == 0 and set(row_elm.attrib) <= {"r"}:
            sheet_data.remove(row_elm)
    return produced
