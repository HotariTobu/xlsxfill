"""What xlsxedit does not do yet, written the way xlsxedit would do it.

Every name here takes and returns xlsxedit's own things — a ``Workbook``,
a ``Worksheet``, a ``Cell`` — and fills a gap upstream leaves open:

- **Duplicating and deleting rows and columns.** ``insert_rows`` writes
  new data and pushes the rest down; it never copies cell content, and
  ``insert_columns`` says outright that it rewrites neither formula text
  nor drawing anchors.
- **Following references through it.** ``row_shift`` follows a uniform
  shift, and nothing follows formula text or drawing anchors at all.
- **Text outside cells.** ``Workbook.replace`` reaches cells only; tab
  names, headers, comments, shapes, charts, validation messages,
  tooltips and document properties are out of reach.
- **Writing values in bulk.** ``Cell.value`` never reuses a shared
  string, edits entries other cells may share, and costs a full-workbook
  scan per write.
- **Saying how many cells the shared strings serve.** ``sst/@count`` is
  a count of references, and the only thing that maintains it counts
  table entries instead, so it is wrong the moment two cells share a
  string or a cell is cleared.
- **Copying a sheet to a chosen position.** ``copy_worksheet`` landed
  upstream after 1.0.1 and always appends.
- **Pictures from bytes at a chosen size.** ``add_image`` takes a path
  and picks the size itself.

Each of these is a candidate upstream request, so each can be deleted
here the day upstream has it.
"""

from _patched_xlsxedit._copy_worksheet import copy_worksheet, move_worksheet
from _patched_xlsxedit._duplicate import (
    delete_columns,
    delete_rows,
    duplicate_columns,
    duplicate_rows,
)
from _patched_xlsxedit._image import add_picture, read_size
from _patched_xlsxedit._link import add_link
from _patched_xlsxedit._tables import refresh_table_columns
from _patched_xlsxedit._text import TextRef, book_texts, sheet_texts
from _patched_xlsxedit._write import Writer, refresh_shared_string_count, serial

__all__ = [
    "TextRef",
    "Writer",
    "add_link",
    "add_picture",
    "book_texts",
    "copy_worksheet",
    "delete_columns",
    "delete_rows",
    "duplicate_columns",
    "duplicate_rows",
    "move_worksheet",
    "read_size",
    "refresh_shared_string_count",
    "refresh_table_columns",
    "serial",
    "sheet_texts",
]
