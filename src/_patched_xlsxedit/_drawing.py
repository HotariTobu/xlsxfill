"""Cell-anchored pictures: PNG sizing, cell-fit geometry, and anchors.

The byte-preserving counterpart of ``xlsxedit.Worksheet.add_image``.
Upstream takes an image path; taking raw bytes here follows the same
direction as upstream's file-like support (its issue #1) and is a
candidate upstream request.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass

from _patched_xlsxedit._xmltext import escape_attr

EMU_PER_PIXEL = 9525
_PNG_MAGIC = b"\x89PNG\r\n\x1a\n"


@dataclass(frozen=True)
class PicturePlacement:
    """One picture to be anchored at a cell."""

    col: int
    row: int
    descr: str | None
    data: bytes
    fit: str
    cell_width_chars: float
    cell_height_points: float


def png_size(data: bytes) -> tuple[int, int]:
    """Width and height in pixels of a PNG image."""
    if not data.startswith(_PNG_MAGIC) or len(data) < 24:
        raise ValueError("not a PNG image")
    width, height = struct.unpack(">II", data[16:24])
    return width, height


def column_width_px(width_chars: float) -> int:
    """Column width in pixels from the character width."""
    return round(width_chars * 7) + 5


def row_height_emu(height_points: float) -> int:
    """Row height in EMU from points."""
    return int(height_points * 12700)


def anchor_xml(image: PicturePlacement, *, pic_id: int, embed_rid: str) -> str:
    """Build one ``<xdr:oneCellAnchor>`` block for an image in a cell."""
    cell_width_emu = column_width_px(image.cell_width_chars) * EMU_PER_PIXEL
    cell_height_emu = row_height_emu(image.cell_height_points)
    px_w, px_h = png_size(image.data)
    img_w, img_h = px_w * EMU_PER_PIXEL, px_h * EMU_PER_PIXEL
    col_off = 0
    row_off = 0
    src_rect = ""
    if image.fit == "fill":
        ext_w, ext_h = cell_width_emu, cell_height_emu
    elif image.fit == "cover":
        ext_w, ext_h = cell_width_emu, cell_height_emu
        scale = max(cell_width_emu / img_w, cell_height_emu / img_h)
        visible_w = cell_width_emu / scale
        visible_h = cell_height_emu / scale
        crop_x = int((1 - visible_w / img_w) / 2 * 100000)
        crop_y = int((1 - visible_h / img_h) / 2 * 100000)
        parts = []
        if crop_x:
            parts.append(f'l="{crop_x}" r="{crop_x}"')
        if crop_y:
            parts.append(f't="{crop_y}" b="{crop_y}"')
        if parts:
            src_rect = f"<a:srcRect {' '.join(parts)}/>"
    else:  # contain
        scale = min(cell_width_emu / img_w, cell_height_emu / img_h)
        ext_w = int(img_w * scale)
        ext_h = int(img_h * scale)
        col_off = (cell_width_emu - ext_w) // 2
        row_off = (cell_height_emu - ext_h) // 2
    descr = f' descr="{escape_attr(image.descr)}"' if image.descr is not None else ""
    return (
        "<xdr:oneCellAnchor>"
        "<xdr:from>"
        f"<xdr:col>{image.col}</xdr:col><xdr:colOff>{col_off}</xdr:colOff>"
        f"<xdr:row>{image.row - 1}</xdr:row><xdr:rowOff>{row_off}</xdr:rowOff>"
        "</xdr:from>"
        f'<xdr:ext cx="{ext_w}" cy="{ext_h}"/>'
        "<xdr:pic>"
        "<xdr:nvPicPr>"
        f'<xdr:cNvPr id="{pic_id}" name="Picture {pic_id}"{descr}/>'
        "<xdr:cNvPicPr/>"
        "</xdr:nvPicPr>"
        "<xdr:blipFill>"
        f'<a:blip r:embed="{embed_rid}"/>'
        f"{src_rect}"
        "<a:stretch><a:fillRect/></a:stretch>"
        "</xdr:blipFill>"
        "<xdr:spPr>"
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        "</xdr:spPr>"
        "</xdr:pic>"
        "<xdr:clientData/>"
        "</xdr:oneCellAnchor>"
    )


WSDR_OPEN = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/'
    'spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/'
    'drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships">'
)
WSDR_CLOSE = "</xdr:wsDr>"
