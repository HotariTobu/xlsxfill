from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from xlsxedit.opc.constants import SML_NS
from xlsxedit.oxml.address import col_to_index, split_address
from xlsxedit.oxml.parser import parse_xml, serialize_xml

from _patched_xlsxedit._remap import (
    rename_sheets_in_charts,
    rename_sheets_in_formulas,
)

if TYPE_CHECKING:
    from collections.abc import Callable, Iterator

    from lxml.etree import _Element
    from xlsxedit.opc.part import Part
    from xlsxedit.workbook import Workbook
    from xlsxedit.worksheet import Worksheet

COMMENTS_RELTYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)

_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_CHART_RELTYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
)

_HEADER_FOOTER = f"{{{SML_NS}}}headerFooter"
_DATA_VALIDATION = f"{{{SML_NS}}}dataValidation"
_HYPERLINK = f"{{{SML_NS}}}hyperlink"
_COMMENT = f"{{{SML_NS}}}comment"
_SML_T = f"{{{SML_NS}}}t"
_DRAWING_T = f"{{{_DRAWING_NS}}}t"

_HEADER_TAGS = (
    ("oddHeader", "header"),
    ("evenHeader", "header"),
    ("firstHeader", "header"),
    ("oddFooter", "footer"),
    ("evenFooter", "footer"),
    ("firstFooter", "footer"),
)
_VALIDATION_ATTRS = ("errorTitle", "error", "promptTitle", "prompt")
_PROPERTY_TAGS = (
    "{http://purl.org/dc/elements/1.1/}title",
    "{http://purl.org/dc/elements/1.1/}subject",
    "{http://purl.org/dc/elements/1.1/}creator",
    "{http://purl.org/dc/elements/1.1/}description",
    "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}keywords",
    "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}category",
    (
        "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
        "lastModifiedBy"
    ),
)


@dataclass
class TextRef:
    sheet: Worksheet | None
    part: str
    at: int
    row: int | None = None
    column: int | None = None
    read: Callable[[], str] = field(default=str, repr=False)
    write: Callable[[str], None] = field(default=lambda _: None, repr=False)


def sheet_texts(workbook: Workbook, worksheet: Worksheet) -> Iterator[TextRef]:
    yield _name_ref(workbook, worksheet)
    yield from _header_refs(worksheet)
    yield from _validation_refs(worksheet)
    yield from _tooltip_refs(worksheet)
    yield from _comment_refs(worksheet)
    yield from _shape_refs(worksheet)
    yield from _chart_refs(worksheet)


def book_texts(workbook: Workbook) -> Iterator[TextRef]:
    element = workbook.properties._element
    for at, tag in enumerate(_PROPERTY_TAGS):
        for found in element.iter(tag):
            yield TextRef(
                None, "properties", at, read=_reader(found), write=_writer(found)
            )


def _reader(element: _Element) -> Callable[[], str]:
    def read() -> str:
        return element.text or ""

    return read


def _writer(
    element: _Element, on_change: Callable[[], None] | None = None
) -> Callable[[str], None]:
    def write(value: str) -> None:
        element.text = value
        if on_change is not None:
            on_change()

    return write


def _attr_reader(element: _Element, name: str) -> Callable[[], str]:
    def read() -> str:
        return element.get(name) or ""

    return read


def _attr_writer(element: _Element, name: str) -> Callable[[str], None]:
    def write(value: str) -> None:
        element.set(name, value)

    return write


def _name_ref(workbook: Workbook, worksheet: Worksheet) -> TextRef:
    def read() -> str:
        return worksheet.name

    def write(value: str) -> None:
        old = worksheet.name
        if value == old:
            return
        workbook.rename_worksheet(old, value)
        worksheet._name = value
        rename_sheets_in_formulas(
            [ws._part.element for ws in workbook.worksheets], [(old, value)]
        )
        rename_sheets_in_charts(workbook, [(old, value)])

    return TextRef(worksheet, "name", 0, read=read, write=write)


def _header_refs(worksheet: Worksheet) -> Iterator[TextRef]:
    block = worksheet._part.element.find(_HEADER_FOOTER)
    if block is None:
        return
    seen: dict[str, int] = {"header": 0, "footer": 0}
    for tag, part in _HEADER_TAGS:
        element = block.find(f"{{{SML_NS}}}{tag}")
        if element is None:
            continue
        at = seen[part]
        seen[part] += 1
        yield TextRef(
            worksheet, part, at, read=_reader(element), write=_writer(element)
        )


def _validation_refs(worksheet: Worksheet) -> Iterator[TextRef]:
    for index, element in enumerate(worksheet._part.element.iter(_DATA_VALIDATION)):
        row, column = _first_cell(element.get("sqref", ""))
        for at, name in enumerate(_VALIDATION_ATTRS):
            if element.get(name) is None:
                continue
            yield TextRef(
                worksheet,
                "validation",
                index * len(_VALIDATION_ATTRS) + at,
                row,
                column,
                read=_attr_reader(element, name),
                write=_attr_writer(element, name),
            )


def _tooltip_refs(worksheet: Worksheet) -> Iterator[TextRef]:
    for at, element in enumerate(worksheet._part.element.iter(_HYPERLINK)):
        if element.get("tooltip") is None:
            continue
        row, column = _first_cell(element.get("ref", ""))
        yield TextRef(
            worksheet,
            "tooltip",
            at,
            row,
            column,
            read=_attr_reader(element, "tooltip"),
            write=_attr_writer(element, "tooltip"),
        )


def _first_cell(ref: str) -> tuple[int | None, int | None]:
    first = ref.split()[0].split(":")[0] if ref.strip() else ""
    if not first:
        return (None, None)
    try:
        letters, row = split_address(first.replace("$", ""))
    except (ValueError, IndexError):
        return (None, None)
    return (row, col_to_index(letters) + 1)


def _blob_refs(part: Part, tag: str) -> Iterator[tuple[_Element, Callable[[], None]]]:
    root = parse_xml(part.blob)

    def flush() -> None:
        part._blob = serialize_xml(root)

    for element in root.iter(tag):
        yield element, flush


def _comment_refs(worksheet: Worksheet) -> Iterator[TextRef]:
    for rel in worksheet._part.rels:
        if rel.is_external or rel.reltype != COMMENTS_RELTYPE:
            continue
        part = rel.target_part
        root = parse_xml(part.blob)
        cells = _comment_cells(root)

        def flush(part: Part = part, root: _Element = root) -> None:
            part._blob = serialize_xml(root)

        for at, element in enumerate(root.iter(_SML_T)):
            row, column = cells.get(id(element), (None, None))
            yield TextRef(
                worksheet,
                "comment",
                at,
                row,
                column,
                read=_reader(element),
                write=_writer(element, flush),
            )


def _comment_cells(root: _Element) -> dict[int, tuple[int | None, int | None]]:
    out: dict[int, tuple[int | None, int | None]] = {}
    for comment in root.iter(_COMMENT):
        cell = _first_cell(comment.get("ref", ""))
        for element in comment.iter(_SML_T):
            out[id(element)] = cell
    return out


def _shape_refs(worksheet: Worksheet) -> Iterator[TextRef]:
    from xlsxedit.drawing import drawing_parts_for_worksheet

    at = 0
    for part in drawing_parts_for_worksheet(worksheet):
        for element, flush in _blob_refs(part, _DRAWING_T):
            yield TextRef(
                worksheet,
                "shape",
                at,
                read=_reader(element),
                write=_writer(element, flush),
            )
            at += 1


def _chart_refs(worksheet: Worksheet) -> Iterator[TextRef]:
    from xlsxedit.drawing import drawing_parts_for_worksheet

    at = 0
    for drawing in drawing_parts_for_worksheet(worksheet):
        for rel in drawing.rels:
            if rel.is_external or rel.reltype != _CHART_RELTYPE:
                continue
            for element, flush in _blob_refs(rel.target_part, _DRAWING_T):
                yield TextRef(
                    worksheet,
                    "chart",
                    at,
                    read=_reader(element),
                    write=_writer(element, flush),
                )
                at += 1
