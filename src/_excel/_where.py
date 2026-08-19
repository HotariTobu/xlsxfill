"""Where a piece of text lives in a workbook.

Places come from [`find`][_excel.Book.find] and go back to
[`set`][_excel.Book.set]. Nothing else makes one: a place is the
library's account of where it read something, not a coordinate the
caller can name on its own. What a place tells you is what the caller
has been shown it needs — which sheet, which cell, which kind of place
— and how the library finds its way back there is not part of that.

Places are positions, not handles. Duplicating or deleting rows moves
what was there, so a place found before a structural change does not
survive it.
"""

from __future__ import annotations

from datetime import date, datetime, time
from typing import TYPE_CHECKING, Literal, cast

if TYPE_CHECKING:
    from _excel._sheet import Sheet
    from _patched_xlsxedit import TextRef

type CellValue = str | int | float | bool | date | time | datetime | None
"""What a cell can hold. ``None`` empties the cell but keeps its format."""


class Cell:
    """A cell's own content."""

    def __init__(self, sheet: Sheet, row: int, column: int) -> None:
        """Point at the cell at ``row`` and ``column`` of ``sheet``."""
        self._sheet = sheet
        self._row = row
        self._column = column

    def __repr__(self) -> str:
        """Show the sheet and the cell."""
        return f"<Cell {self._sheet.name!r} r{self._row}c{self._column}>"

    @property
    def sheet(self) -> Sheet:
        """The sheet it is on."""
        return self._sheet

    @property
    def row(self) -> int:
        """Its 1-based row, as in Excel's own R1C1 notation."""
        return self._row

    @property
    def column(self) -> int:
        """Its 1-based column."""
        return self._column

    @property
    def part(self) -> Literal["cell"]:
        """What holds the text."""
        return "cell"


class CellText:
    """Text a cell carries besides its content."""

    def __init__(self, sheet: Sheet, ref: TextRef) -> None:
        """Point at what ``ref`` names, on ``sheet``."""
        self._sheet = sheet
        self._ref = ref

    def __repr__(self) -> str:
        """Show the sheet, the cell, and the kind."""
        return f"<CellText {self._sheet.name!r} {self._ref.part} at {self._ref.at}>"

    @property
    def sheet(self) -> Sheet:
        """The sheet it is on."""
        return self._sheet

    @property
    def row(self) -> int:
        """The cell's 1-based row."""
        return self._ref.row or 0

    @property
    def column(self) -> int:
        """The cell's 1-based column."""
        return self._ref.column or 0

    @property
    def part(self) -> Literal["comment", "tooltip", "validation"]:
        """What holds the text."""
        return cast('Literal["comment", "tooltip", "validation"]', self._ref.part)


class SheetText:
    """Text a sheet carries outside its cells."""

    def __init__(self, sheet: Sheet, ref: TextRef) -> None:
        """Point at what ``ref`` names, on ``sheet``."""
        self._sheet = sheet
        self._ref = ref

    def __repr__(self) -> str:
        """Show the sheet and the kind."""
        return f"<SheetText {self._sheet.name!r} {self._ref.part} at {self._ref.at}>"

    @property
    def sheet(self) -> Sheet:
        """The sheet it belongs to."""
        return self._sheet

    @property
    def part(self) -> Literal["name", "header", "footer", "shape", "chart"]:
        """What holds the text."""
        return cast(
            'Literal["name", "header", "footer", "shape", "chart"]', self._ref.part
        )


class BookText:
    """One of the workbook's document properties."""

    def __init__(self, ref: TextRef) -> None:
        """Point at what ``ref`` names."""
        self._ref = ref

    def __repr__(self) -> str:
        """Show which property."""
        return f"<BookText at {self._ref.at}>"

    @property
    def part(self) -> Literal["properties"]:
        """What holds the text."""
        return "properties"


type Where = Cell | CellText | SheetText | BookText
"""Any place a piece of text lives."""


class Found:
    """A search hit."""

    def __init__(self, where: Where, text: str) -> None:
        """Record that ``text`` was read at ``where``."""
        self._where = where
        self._text = text

    def __repr__(self) -> str:
        """Show the place and the text."""
        return f"<Found {self._where!r} {self._text!r}>"

    @property
    def where(self) -> Where:
        """Where the text lives."""
        return self._where

    @property
    def text(self) -> str:
        """The text as it reads now."""
        return self._text
