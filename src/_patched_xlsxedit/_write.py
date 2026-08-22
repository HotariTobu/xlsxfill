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
    def __init__(self, workbook: Workbook) -> None:
        self._workbook = workbook
        self._index: dict[str, int] = {}
        table = workbook.shared_strings
        if table is not None:
            for i in range(len(table)):
                self._index.setdefault(table.text(i), i)

    def write(self, cell: Cell, value: Value) -> None:
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
    for node in si.iter(_T):
        text = node.text or ""
        if text != text.strip() or any(c in text for c in "\n\r\t"):
            node.set(_XML_SPACE, "preserve")
