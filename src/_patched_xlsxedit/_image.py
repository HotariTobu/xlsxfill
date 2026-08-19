"""Pictures placed from bytes, at a size the caller decides.

``Worksheet.add_image`` takes a filesystem path and fits the picture into
a pixel box of its own choosing. Substitution has the bytes in hand and
has already worked out where the picture should go, so this takes both.

Reading the header for the picture's own size is the same thing Excel
does when you insert one — it is how it can offer to keep the aspect
ratio — so the size is available before anything is placed.

Sizes arrive in pixels and may be fractional, because the format counts
in units far finer than a pixel and rounding on the way in would show.
"""

from __future__ import annotations

import struct
from typing import TYPE_CHECKING

from lxml import etree
from xlsxedit.opc.constants import CT, OFFICE_REL_NS, RT, SML_NS
from xlsxedit.opc.part import Part
from xlsxedit.oxml.parser import serialize_xml
from xlsxedit.worksheet_order import insert_worksheet_child

if TYPE_CHECKING:
    from lxml.etree import _Element
    from xlsxedit.workbook import Workbook
    from xlsxedit.worksheet import Worksheet

EMU_PER_PIXEL = 9525

XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

_NSMAP = {"xdr": XDR_NS, "a": A_NS, "r": OFFICE_REL_NS}
_WS_DRAWING = f"{{{SML_NS}}}drawing"

_PNG_MAGIC = b"\x89PNG\r\n\x1a\n"
_GIF_MAGIC = (b"GIF87a", b"GIF89a")
_JPEG_MAGIC = b"\xff\xd8"
_GIF_CONTENT_TYPE = "image/gif"
_JPEG_SIZE_MARKERS = frozenset(range(0xC0, 0xD0)) - {0xC4, 0xC8, 0xCC}


def read_size(data: bytes) -> tuple[int, int]:
    """The picture's own width and height in pixels.

    Raises:
        ValueError: The bytes are not a picture this can read.
    """
    if data.startswith(_PNG_MAGIC):
        return _png_size(data)
    if data.startswith(_GIF_MAGIC):
        return _gif_size(data)
    if data.startswith(_JPEG_MAGIC):
        return _jpeg_size(data)
    message = "not a PNG, JPEG, or GIF image"
    raise ValueError(message)


def content_type(data: bytes) -> tuple[str, str]:
    """The media type and file extension the bytes should be stored under."""
    if data.startswith(_PNG_MAGIC):
        return CT.PNG, "png"
    if data.startswith(_GIF_MAGIC):
        return _GIF_CONTENT_TYPE, "gif"
    return CT.JPEG, "jpeg"


def _png_size(data: bytes) -> tuple[int, int]:
    if len(data) < 24:
        message = "truncated PNG image"
        raise ValueError(message)
    width, height = struct.unpack(">II", data[16:24])
    return width, height


def _gif_size(data: bytes) -> tuple[int, int]:
    if len(data) < 10:
        message = "truncated GIF image"
        raise ValueError(message)
    width, height = struct.unpack("<HH", data[6:10])
    return width, height


def _jpeg_size(data: bytes) -> tuple[int, int]:
    pos = 2
    while pos + 9 < len(data):
        if data[pos] != 0xFF:
            pos += 1
            continue
        marker = data[pos + 1]
        if marker in _JPEG_SIZE_MARKERS:
            height, width = struct.unpack(">HH", data[pos + 5 : pos + 9])
            return width, height
        (length,) = struct.unpack(">H", data[pos + 2 : pos + 4])
        pos += 2 + length
    message = "no size in JPEG image"
    raise ValueError(message)


def add_picture(
    workbook: Workbook,
    worksheet: Worksheet,
    row: int,
    column: int,
    data: bytes,
    *,
    width: float,
    height: float,
    offset_x: float = 0,
    offset_y: float = 0,
    crop: tuple[float, float, float, float] | None = None,
    alt: str | None = None,
    media: dict[bytes, Part] | None = None,
) -> None:
    """Anchor a picture at a cell, drawn at the given pixel size.

    Args:
        workbook: The workbook to add the media part to.
        worksheet: The sheet to draw on.
        row: 1-based anchor row.
        column: 1-based anchor column.
        data: The picture file's bytes.
        width: How wide to draw it, in pixels.
        height: How tall to draw it, in pixels.
        offset_x: Pixels right of the anchor cell's left edge.
        offset_y: Pixels below the anchor cell's top edge.
        crop: The part of the picture to show, as ``(left, top, right,
            bottom)`` in its own pixels.
        alt: Alternative text.
        media: Media parts already added, keyed by their bytes, so the
            same picture used twice is stored once.
    """
    package = workbook._package
    drawing, root = _drawing_part(workbook, worksheet)

    store = media if media is not None else {}
    part = store.get(data)
    if part is None:
        media_type, extension = content_type(data)
        partname = package.next_partname(f"/xl/media/image%d.{extension}")
        part = Part(partname, media_type, data, package)
        package._add_part(part)
        store[data] = part

    root.append(
        _anchor(
            row=row,
            column=column,
            width=width,
            height=height,
            offset_x=offset_x,
            offset_y=offset_y,
            crop=_crop_edges(crop, read_size(data)),
            alt=alt,
            pic_id=len(list(root.iter(f"{{{XDR_NS}}}cNvPr"))) + 1,
            embed=drawing.relate_to(part, RT.IMAGE),
        )
    )
    drawing._blob = serialize_xml(root)


def _crop_edges(
    crop: tuple[float, float, float, float] | None, size: tuple[int, int]
) -> tuple[int, int, int, int]:
    """The four edges as the hundred-thousandths the format counts in."""
    if crop is None:
        return (0, 0, 0, 0)
    left, top, right, bottom = crop
    width, height = size
    return (
        int(left / width * 100000),
        int(top / height * 100000),
        int((width - right) / width * 100000),
        int((height - bottom) / height * 100000),
    )


def _drawing_part(workbook: Workbook, worksheet: Worksheet) -> tuple[Part, _Element]:
    from xlsxedit.drawing import drawing_parts_for_worksheet
    from xlsxedit.oxml.parser import parse_xml

    existing = drawing_parts_for_worksheet(worksheet)
    if existing:
        return existing[0], parse_xml(existing[0].blob)
    package = workbook._package
    partname = package.next_partname("/xl/drawings/drawing%d.xml")
    part = Part(partname, CT.DRAWING, b"", package)
    package._add_part(part)
    rid = worksheet._part.relate_to(part, RT.DRAWING)
    element = etree.Element(_WS_DRAWING)
    element.set(f"{{{OFFICE_REL_NS}}}id", rid)
    insert_worksheet_child(worksheet._part.element, element)
    return part, etree.Element(f"{{{XDR_NS}}}wsDr", nsmap=_NSMAP)


def _anchor(
    *,
    row: int,
    column: int,
    width: float,
    height: float,
    offset_x: float,
    offset_y: float,
    crop: tuple[int, int, int, int],
    alt: str | None,
    pic_id: int,
    embed: str,
) -> _Element:
    anchor = etree.Element(f"{{{XDR_NS}}}oneCellAnchor", nsmap=_NSMAP)
    origin = etree.SubElement(anchor, f"{{{XDR_NS}}}from")
    etree.SubElement(origin, f"{{{XDR_NS}}}col").text = str(column - 1)
    etree.SubElement(origin, f"{{{XDR_NS}}}colOff").text = str(
        int(offset_x * EMU_PER_PIXEL)
    )
    etree.SubElement(origin, f"{{{XDR_NS}}}row").text = str(row - 1)
    etree.SubElement(origin, f"{{{XDR_NS}}}rowOff").text = str(
        int(offset_y * EMU_PER_PIXEL)
    )
    etree.SubElement(
        anchor,
        f"{{{XDR_NS}}}ext",
        cx=str(int(width * EMU_PER_PIXEL)),
        cy=str(int(height * EMU_PER_PIXEL)),
    )

    pic = etree.SubElement(anchor, f"{{{XDR_NS}}}pic")
    names = etree.SubElement(pic, f"{{{XDR_NS}}}nvPicPr")
    name = etree.SubElement(names, f"{{{XDR_NS}}}cNvPr")
    name.set("id", str(pic_id))
    name.set("name", f"Picture {pic_id}")
    if alt is not None:
        name.set("descr", alt)
    etree.SubElement(names, f"{{{XDR_NS}}}cNvPicPr")

    fill = etree.SubElement(pic, f"{{{XDR_NS}}}blipFill")
    blip = etree.SubElement(fill, f"{{{A_NS}}}blip")
    blip.set(f"{{{OFFICE_REL_NS}}}embed", embed)
    left, top, right, bottom = crop
    if any(crop):
        rect = etree.SubElement(fill, f"{{{A_NS}}}srcRect")
        for edge, amount in (("l", left), ("t", top), ("r", right), ("b", bottom)):
            if amount:
                rect.set(edge, str(amount))
    stretch = etree.SubElement(fill, f"{{{A_NS}}}stretch")
    etree.SubElement(stretch, f"{{{A_NS}}}fillRect")

    shape = etree.SubElement(pic, f"{{{XDR_NS}}}spPr")
    geometry = etree.SubElement(shape, f"{{{A_NS}}}prstGeom", prst="rect")
    etree.SubElement(geometry, f"{{{A_NS}}}avLst")

    etree.SubElement(anchor, f"{{{XDR_NS}}}clientData")
    return anchor
