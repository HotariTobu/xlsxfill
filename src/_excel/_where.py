from __future__ import annotations

from datetime import date, datetime, time
from typing import TYPE_CHECKING, Literal, cast

if TYPE_CHECKING:
    from _excel._sheet import Sheet
    from _patched_xlsxedit import TextRef

type CellValue = str | int | float | bool | date | time | datetime | None


class Cell:
    def __init__(self, sheet: Sheet, row: int, column: int) -> None:
        self._sheet = sheet
        self._row = row
        self._column = column

    def __repr__(self) -> str:
        return f"<Cell {self._sheet.name!r} r{self._row}c{self._column}>"

    @property
    def sheet(self) -> Sheet:
        return self._sheet

    @property
    def row(self) -> int:
        return self._row

    @property
    def column(self) -> int:
        return self._column

    @property
    def part(self) -> Literal["cell"]:
        return "cell"


class CellText:
    def __init__(self, sheet: Sheet, ref: TextRef) -> None:
        self._sheet = sheet
        self._ref = ref

    def __repr__(self) -> str:
        return f"<CellText {self._sheet.name!r} {self._ref.part} at {self._ref.at}>"

    @property
    def sheet(self) -> Sheet:
        return self._sheet

    @property
    def row(self) -> int:
        return self._ref.row or 0

    @property
    def column(self) -> int:
        return self._ref.column or 0

    @property
    def part(self) -> Literal["comment", "tooltip", "validation"]:
        return cast('Literal["comment", "tooltip", "validation"]', self._ref.part)


class SheetText:
    def __init__(self, sheet: Sheet, ref: TextRef) -> None:
        self._sheet = sheet
        self._ref = ref

    def __repr__(self) -> str:
        return f"<SheetText {self._sheet.name!r} {self._ref.part} at {self._ref.at}>"

    @property
    def sheet(self) -> Sheet:
        return self._sheet

    @property
    def part(self) -> Literal["name", "header", "footer", "shape", "chart"]:
        return cast(
            'Literal["name", "header", "footer", "shape", "chart"]', self._ref.part
        )


class BookText:
    def __init__(self, ref: TextRef) -> None:
        self._ref = ref

    def __repr__(self) -> str:
        return f"<BookText at {self._ref.at}>"

    @property
    def part(self) -> Literal["properties"]:
        return "properties"


type Where = Cell | CellText | SheetText | BookText


class Found:
    def __init__(self, where: Where, text: str) -> None:
        self._where = where
        self._text = text

    def __repr__(self) -> str:
        return f"<Found {self._where!r} {self._text!r}>"

    @property
    def where(self) -> Where:
        return self._where

    @property
    def text(self) -> str:
        return self._text
