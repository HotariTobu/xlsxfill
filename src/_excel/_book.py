from __future__ import annotations

import re
from typing import TYPE_CHECKING, overload

from xlsxedit.oxml.address import index_to_col, join_address
from xlsxedit.workbook import Workbook

from _excel._sheet import Sheet
from _excel._where import BookText, Cell, CellText, Found, SheetText
from _patched_xlsxedit import (
    Writer,
    add_link,
    add_picture,
    book_texts,
    copy_worksheet,
    refresh_shared_string_count,
    refresh_table_columns,
    sheet_texts,
)

if TYPE_CHECKING:
    from collections.abc import Sequence
    from pathlib import Path
    from typing import BinaryIO

    from xlsxedit.cell import Cell as _Cell
    from xlsxedit.opc.part import Part
    from xlsxedit.parts import WorksheetPart
    from xlsxedit.worksheet import Worksheet

    from _excel._image import Image
    from _excel._where import CellValue, Where

_DEFAULT_ROW_HEIGHT = 15.0
_PIXELS_PER_POINT = 96 / 72
_CELL_PADDING_PIXELS = 5
_PIXELS_PER_CHARACTER = 7


class Book:
    def __init__(self, workbook: Workbook) -> None:
        self._workbook = workbook
        self._writer = Writer(workbook)
        self._media: dict[bytes, Part] = {}
        self._sheets: dict[int, Sheet] = {}

    def __repr__(self) -> str:
        return f"<Book {self._workbook.sheetnames!r}>"

    @classmethod
    def open(cls, source: str | Path | BinaryIO) -> Book:
        return cls(Workbook.open(source))

    def save(self, target: str | Path | BinaryIO) -> None:
        refresh_table_columns(self._workbook)
        refresh_shared_string_count(self._workbook)
        self._workbook.save(target)

    def find(self, pattern: str | re.Pattern[str] | None = None) -> list[Found]:
        matches = re.compile(pattern) if isinstance(pattern, str) else pattern
        found: list[Found] = []
        for worksheet in self._workbook.worksheets:
            sheet = self._sheet_for(worksheet._part)
            for row, cells in worksheet.iter_rows():
                for cell in cells:
                    if cell.has_formula:
                        continue
                    value = cell.value
                    if not isinstance(value, str) or _misses(matches, value):
                        continue
                    column = cell._element.get("r", "")
                    found.append(Found(Cell(sheet, row, _column_of(column)), value))
            for ref in sheet_texts(self._workbook, worksheet):
                text = ref.read()
                if _misses(matches, text):
                    continue
                place = (
                    CellText(sheet, ref)
                    if ref.row is not None
                    else SheetText(sheet, ref)
                )
                found.append(Found(place, text))
        for ref in book_texts(self._workbook):
            text = ref.read()
            if not _misses(matches, text):
                found.append(Found(BookText(ref), text))
        return found

    @overload
    def set(self, where: Cell, value: CellValue) -> None: ...

    @overload
    def set(self, where: CellText | SheetText | BookText, value: str) -> None: ...

    def set(self, where: Where, value: CellValue) -> None:
        if isinstance(where, Cell):
            self._writer.write(self._cell(where), value)
            return
        where._ref.write("" if value is None else str(value))

    def cell_size(self, cell: Cell) -> tuple[int, int]:
        worksheet = cell.sheet._sheet
        characters = worksheet.column_dimensions[index_to_col(cell.column - 1)].width
        points = worksheet.row_dimensions[cell.row].height
        if points is None:
            points = _default_row_height(worksheet)
        return (
            round(characters * _PIXELS_PER_CHARACTER) + _CELL_PADDING_PIXELS,
            round(points * _PIXELS_PER_POINT),
        )

    def insert_link(self, cell: Cell, url: str) -> None:
        add_link(
            cell.sheet._sheet,
            join_address(index_to_col(cell.column - 1), cell.row),
            url,
        )

    def insert_image(
        self,
        cell: Cell,
        image: Image,
        *,
        width: float,
        height: float,
        offset_x: float = 0,
        offset_y: float = 0,
        crop: tuple[float, float, float, float] | None = None,
        alt: str | None = None,
    ) -> None:
        add_picture(
            self._workbook,
            cell.sheet._sheet,
            cell.row,
            cell.column,
            image.data,
            width=width,
            height=height,
            offset_x=offset_x,
            offset_y=offset_y,
            crop=crop,
            alt=alt,
            media=self._media,
        )

    def duplicate_sheet(self, sheet: Sheet, names: Sequence[str]) -> list[Sheet]:
        copies: list[Sheet] = []
        after = sheet.name
        for name in names:
            made = copy_worksheet(self._workbook, sheet.name, name, after=after)
            copies.append(self._sheet_for(made._part))
            after = name
        return copies

    def delete_sheet(self, sheet: Sheet) -> None:
        self._workbook.remove_worksheet(sheet.name)

    def _sheet_for(self, part: WorksheetPart) -> Sheet:
        found = self._sheets.get(id(part))
        if found is None:
            found = Sheet(self._workbook, part)
            self._sheets[id(part)] = found
        return found

    def _cell(self, cell: Cell) -> _Cell:
        worksheet = cell.sheet._sheet
        return worksheet[join_address(index_to_col(cell.column - 1), cell.row)]


def _misses(matches: re.Pattern[str] | None, text: str) -> bool:
    return matches is not None and matches.search(text) is None


def _column_of(address: str) -> int:
    from xlsxedit.oxml.address import col_to_index, split_address

    letters, _ = split_address(address)
    return col_to_index(letters) + 1


def _default_row_height(worksheet: Worksheet) -> float:
    from xlsxedit.opc.constants import SML_NS

    element = worksheet._part.element.find(f"{{{SML_NS}}}sheetFormatPr")
    if element is not None and element.get("defaultRowHeight"):
        return float(element.get("defaultRowHeight", _DEFAULT_ROW_HEIGHT))
    return _DEFAULT_ROW_HEIGHT
