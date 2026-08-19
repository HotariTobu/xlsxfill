"""What a person can do to a workbook in Excel.

Everything here has a counterpart a person can reach in Excel: open and
save a file, copy or delete a sheet, insert copied rows, delete rows,
search, type into a place, insert a link, insert a picture. Nothing here
exposes how an xlsx file is put together — no parts, no shared strings,
no XML, no coordinate mappings — because none of that is anything a
person does in Excel.

That is the whole rule for what belongs here. An operation with no
counterpart in Excel does not go in, and an operation that has one
behaves the way Excel behaves, down to what happens to the formulas.

The rule governs the operations, not the places they act on. A place is
whatever the caller has to tell apart, which is why
[`find`][_excel.Book.find] reaches text Excel's own search cannot.
"""

from _excel._book import Book
from _excel._image import Image
from _excel._sheet import Sheet
from _excel._where import (
    BookText,
    Cell,
    CellText,
    CellValue,
    Found,
    SheetText,
    Where,
)

__all__ = [
    "Book",
    "BookText",
    "Cell",
    "CellText",
    "CellValue",
    "Found",
    "Image",
    "Sheet",
    "SheetText",
    "Where",
]
