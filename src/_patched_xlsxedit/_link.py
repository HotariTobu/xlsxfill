"""Attaching a link to a cell.

``Hyperlink.url`` reuses an existing external relationship whenever the
target matches, so ten cells linking to the same address end up sharing
one relationship. Excel gives each link its own, and sharing means
removing one link takes the others with it.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree
from xlsxedit.opc.constants import OFFICE_REL_NS, RT, SML_NS
from xlsxedit.worksheet_order import insert_worksheet_child, reposition_worksheet_child

if TYPE_CHECKING:
    from xlsxedit.worksheet import Worksheet

_HYPERLINKS = f"{{{SML_NS}}}hyperlinks"
_HYPERLINK = f"{{{SML_NS}}}hyperlink"
_R_ID = f"{{{OFFICE_REL_NS}}}id"


def add_link(worksheet: Worksheet, address: str, url: str) -> None:
    """Point the cell at ``address`` to ``url``, on a link of its own.

    Args:
        worksheet: The sheet the cell is on.
        address: The cell in A1 notation.
        url: The link target.
    """
    part = worksheet._part
    block = part.element.find(_HYPERLINKS)
    if block is None:
        block = etree.Element(_HYPERLINKS)
        insert_worksheet_child(part.element, block)
    else:
        reposition_worksheet_child(part.element, block)

    element = next(
        (link for link in block.findall(_HYPERLINK) if link.get("ref") == address),
        None,
    )
    if element is None:
        element = etree.SubElement(block, _HYPERLINK)
        element.set("ref", address)
    element.attrib.pop("location", None)

    rels = part.rels
    r_id = rels.next_rId()
    rels.add_relationship(RT.HYPERLINK, url, r_id, "External")
    element.set(_R_ID, r_id)
