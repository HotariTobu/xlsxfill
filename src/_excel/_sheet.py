from __future__ import annotations

from typing import TYPE_CHECKING

from _patched_xlsxedit import (
    delete_columns,
    delete_rows,
    duplicate_columns,
    duplicate_rows,
)

if TYPE_CHECKING:
    from xlsxedit.parts import WorksheetPart
    from xlsxedit.workbook import Workbook
    from xlsxedit.worksheet import Worksheet


class Sheet:
    def __init__(self, workbook: Workbook, part: WorksheetPart) -> None:
        self._workbook = workbook
        self._part = part

    def __repr__(self) -> str:
        return f"<Sheet {self.name!r}>"

    @property
    def _sheet(self) -> Worksheet:
        for sheet in self._workbook.worksheets:
            if sheet._part is self._part:
                return sheet
        message = "the sheet is no longer in the workbook"
        raise ValueError(message)

    @property
    def name(self) -> str:
        return self._sheet.name

    @property
    def merged_ranges(self) -> list[tuple[int, int, int, int]]:
        from xlsxedit.merge import parse_range
        from xlsxedit.oxml.address import col_to_index

        blocks = []
        for ref in self._sheet.merged_ranges:
            left, top, right, bottom = parse_range(ref)
            blocks.append(
                (top, col_to_index(left) + 1, bottom, col_to_index(right) + 1)
            )
        return blocks

    def duplicate_rows(
        self, top: int, bottom: int, *, copies: int
    ) -> list[tuple[int, int]]:
        return duplicate_rows(self._workbook, self._sheet, top, bottom, copies)

    def duplicate_columns(
        self, left: int, right: int, *, copies: int
    ) -> list[tuple[int, int]]:
        return duplicate_columns(self._workbook, self._sheet, left, right, copies)

    def delete_rows(self, top: int, bottom: int) -> None:
        delete_rows(self._workbook, self._sheet, top, bottom)

    def delete_columns(self, left: int, right: int) -> None:
        delete_columns(self._workbook, self._sheet, left, right)
