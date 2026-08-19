"""Writing cell values without wrecking the shared-string table.

``Cell.value`` has three problems as a bulk writer. It never reuses an
existing shared string, so writing the same word into a thousand cells
grows the table a thousand times. When the entry it is replacing looks
unreferenced it edits that entry in place, which rewrites a table other
cells may yet come to share. And deciding that takes a scan of every cell
in the workbook, so writing n cells costs n squared.

A writer keeps the index it needs, built once and kept up to date.

The table's own ``count`` needs restating afterwards, for a reason that
has nothing to do with this writer; see
[`refresh_shared_string_count`][_patched_xlsxedit.refresh_shared_string_count].
"""

from __future__ import annotations

from datetime import date, datetime, time
from typing import TYPE_CHECKING

from lxml import etree
from xlsxedit.opc.constants import SML_NS

if TYPE_CHECKING:
    from xlsxedit.cell import Cell
    from xlsxedit.workbook import Workbook

_V = f"{{{SML_NS}}}v"
_T = f"{{{SML_NS}}}t"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_F = f"{{{SML_NS}}}f"
_IS = f"{{{SML_NS}}}is"

_EPOCH = date(1899, 12, 30)

type Value = str | int | float | bool | date | time | datetime | None


def serial(value: date | time | datetime) -> float | int:
    """The Excel serial number of a date, a time, or both.

    Raises:
        ValueError: The datetime carries a time zone, which a serial
            number cannot express.
    """
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            message = "a datetime with a time zone has no serial number"
            raise ValueError(message)
        return _days(value.date()) + _fraction(value.time())
    if isinstance(value, date):
        return _days(value)
    return _fraction(value)


def _days(value: date) -> int:
    return (value - _EPOCH).days


def _fraction(value: time) -> float:
    seconds = value.hour * 3600 + value.minute * 60 + value.second
    return (seconds + value.microsecond / 1_000_000) / 86400


class Writer:
    """Writes cell values into one workbook, reusing shared strings."""

    def __init__(self, workbook: Workbook) -> None:
        """Index the shared strings ``workbook`` already has."""
        self._workbook = workbook
        self._index: dict[str, int] = {}
        table = workbook.shared_strings
        if table is not None:
            for i in range(len(table)):
                self._index.setdefault(table.text(i), i)

    def write(self, cell: Cell, value: Value) -> None:
        """Put ``value`` in ``cell``, keeping its format.

        Raises:
            ValueError: The value cannot be written, such as a datetime
                carrying a time zone.
        """
        element = cell._element
        if isinstance(value, bool):
            self._clear(cell)
            element.set("t", "b")
            self._set_v(element, "1" if value else "0")
            return
        if value is None or value == "":
            self._blank(cell)
            return
        if isinstance(value, str):
            self._clear(cell)
            element.set("t", "s")
            self._set_v(element, str(self.intern(value)))
            return
        self._clear(cell)
        element.attrib.pop("t", None)
        number = serial(value) if isinstance(value, date | time) else value
        self._set_v(element, repr(number))

    def intern(self, text: str) -> int:
        """The shared-string index of ``text``, adding it if it is new."""
        found = self._index.get(text)
        if found is not None:
            return found
        table = self._workbook.shared_strings
        if table is None:
            message = "workbook has no shared strings"
            raise ValueError(message)
        index = table.add(text)
        _mark_space(table.get(index))
        self._index[text] = index
        return index

    def _blank(self, cell: Cell) -> None:
        element = cell._element
        self._clear(cell)
        element.attrib.pop("t", None)
        if element.get("s") is not None:
            return
        row = element.getparent()
        if row is None:
            return
        row.remove(element)
        if len(row) == 0 and set(row.attrib) <= {"r"}:
            sheet_data = row.getparent()
            if sheet_data is not None:
                sheet_data.remove(row)

    def _clear(self, cell: Cell) -> None:
        element = cell._element
        if element.find(_F) is not None:
            self._workbook._on_formula_removed(cell._worksheet, element)
        for child in list(element):
            if child.tag in (_V, _F, _IS):
                element.remove(child)

    @staticmethod
    def _set_v(element: etree._Element, text: str) -> None:
        etree.SubElement(element, _V).text = text


def refresh_shared_string_count(workbook: Workbook) -> None:
    """Restate how many cells the shared string table actually serves.

    ``sst/@count`` is the total number of string references in the
    workbook, not the size of the table -- ECMA-376 §18.4.9. The only
    thing that maintains it is ``SharedStringTable.add``, which adds one
    per new entry, so the attribute tracks the table instead. Write the
    same string into two cells and the count is one short; clear a cell
    and it stays too high.

    Nothing can keep a running total, because the references move: rows
    and columns get copied, cells get cleared. Counting them is the only
    answer, and it has to happen once the sheets say what they finally
    say.
    """
    table = workbook.shared_strings
    if table is None:
        return
    references = 0
    for worksheet in workbook.worksheets:
        for cell in worksheet.iter_cells():
            element = cell._element
            if element.get("t") == "s" and element.find(_V) is not None:
                references += 1
    table._element.set("count", str(references))
    table._element.set("uniqueCount", str(len(table)))


def _mark_space(si: etree._Element) -> None:
    """Mark an entry whose whitespace a reader would otherwise be free to fold.

    ``SharedStringTable.add`` decides this on ``" " in text``, which looks
    at spaces and nothing else. A string broken across lines carries no
    space, so it goes in unmarked and the line breaks can be folded away.
    Tabs go the same way. What has to be marked is whitespace at either
    end, or a line break or a tab anywhere.

    Only ever marks. An entry marked where it did not need to be is
    harmless, and unmarking it would be a matter of taste, not of the
    text surviving.
    """
    for node in si.iter(_T):
        text = node.text or ""
        if text != text.strip() or any(c in text for c in "\n\r\t"):
            node.set(_XML_SPACE, "preserve")
