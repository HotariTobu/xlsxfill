"""Externally patched xlsxedit.

xlsxedit re-serializes XML parts through lxml on save, which does not
round-trip authored XML byte-for-byte. xlsxfill's output contract is
byte-exact, so this package patches xlsxedit from the outside with
byte-preserving counterparts of its APIs:

- [`Package`][_patched_xlsxedit.Package] keeps every part as raw bytes
  (vs. ``xlsxedit.opc.package.OpcPackage``).
- [`Relationships`][_patched_xlsxedit._rels.Relationships] and
  [`ContentTypes`][_patched_xlsxedit._content_types.ContentTypes] splice
  bookkeeping edits into the raw text (vs. ``xlsxedit.opc.rel``).
- [`Workbook`][_patched_xlsxedit.Workbook] carries the worksheet-part
  structure operations behind upstream's ``copy_worksheet`` /
  ``remove_worksheet`` / ``rename_worksheet``, pictures behind
  ``add_image``, and [`SharedStringTable`][_patched_xlsxedit.SharedStringTable]
  (vs. ``xlsxedit.shared_strings``).
- The remap module generalizes ``xlsxedit.row_shift`` from uniform shifts
  to arbitrary endpoint mappings.

Each patched API is written as a preview of xlsxedit's own direction so
its use case can be proposed upstream as-is. xlsxedit itself is never
imported by xlsxfill; everything goes through this package.
"""

from xlsxedit.merge import parse_range
from xlsxedit.oxml.address import col_to_index, index_to_col, split_address

from _patched_xlsxedit._drawing import PicturePlacement
from _patched_xlsxedit._package import Package
from _patched_xlsxedit._remap import (
    EndpointMapper,
    PointMapper,
    expand_anchors,
    rebase_formula,
    remap_defined_name,
    remap_range_ref,
    remap_sqref,
    remap_table,
)
from _patched_xlsxedit._shared_strings import SharedStringTable
from _patched_xlsxedit._workbook import RT, Workbook
from _patched_xlsxedit._worksheet import add_hyperlinks, insert_worksheet_child
from _patched_xlsxedit._xmltext import (
    emit_attrs,
    escape_attr,
    escape_go_text,
    escape_text,
    parse_attrs,
    unescape,
)

__all__ = [
    "RT",
    "EndpointMapper",
    "Package",
    "PicturePlacement",
    "PointMapper",
    "SharedStringTable",
    "Workbook",
    "add_hyperlinks",
    "col_to_index",
    "emit_attrs",
    "escape_attr",
    "escape_go_text",
    "escape_text",
    "expand_anchors",
    "index_to_col",
    "insert_worksheet_child",
    "parse_attrs",
    "parse_range",
    "rebase_formula",
    "remap_defined_name",
    "remap_range_ref",
    "remap_sqref",
    "remap_table",
    "split_address",
    "unescape",
]
