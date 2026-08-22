from __future__ import annotations

from typing import TYPE_CHECKING

from xlsxedit.merge import parse_range
from xlsxedit.opc.constants import SML_NS
from xlsxedit.oxml.address import col_to_index, index_to_col, join_address
from xlsxedit.oxml.parser import serialize_xml

if TYPE_CHECKING:
    from collections.abc import Callable

    from xlsxedit.drawing import Table
    from xlsxedit.workbook import Workbook
    from xlsxedit.worksheet import Worksheet

_TABLE_COLUMNS = f"{{{SML_NS}}}tableColumns"
_TABLE_COLUMN = f"{{{SML_NS}}}tableColumn"


def remap_tables(worksheet: Worksheet, map_range: Callable[[str], str]) -> None:
    for table in worksheet.tables:
        if not table.ref:
            continue
        moved = map_range(table.ref)
        if moved != table.ref:
            table.resize(moved)


def refresh_table_columns(workbook: Workbook) -> None:
    for worksheet in workbook.worksheets:
        for table in worksheet.tables:
            if not table.ref:
                continue
            _rename(table, _headers(worksheet, table.ref))


def _rename(table: Table, names: list[str]) -> None:
    element = table._element
    columns = element.find(_TABLE_COLUMNS)
    if columns is None or names == [c.get("name", "") for c in columns]:
        return
    for child in list(columns):
        columns.remove(child)
    columns.set("count", str(len(names)))
    for index, name in enumerate(names, start=1):
        columns.append(
            columns.makeelement(_TABLE_COLUMN, {"id": str(index), "name": name})
        )
    if table._part is not None:
        table._part._blob = serialize_xml(element)


def _headers(worksheet: Worksheet, ref: str) -> list[str]:
    c1, r1, c2, _ = parse_range(ref)
    return [
        _header(worksheet, r1, col)
        for col in range(col_to_index(c1), col_to_index(c2) + 1)
    ]


def _header(worksheet: Worksheet, row: int, col: int) -> str:
    value = worksheet[join_address(index_to_col(col), row)].value
    return "" if value is None else str(value)
