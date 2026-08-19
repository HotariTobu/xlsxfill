"""``Workbook.copy_worksheet`` for xlsxedit releases that lack it.

Merged upstream (commit ``3271d365``, "added copy_worksheet Fixes #2")
but not in 1.0.1, which ships neither the method nor the
``xlsxedit.sheet_clone`` module it needs. This is a port of that merged
code, kept behaviour-identical so it can be deleted the day upstream
releases.

[`copy_worksheet`][_patched_xlsxedit.copy_worksheet] calls upstream's own
method when the installed version has it, so the port stops running by
itself.
"""

from __future__ import annotations

import re
from contextlib import suppress
from copy import deepcopy
from typing import TYPE_CHECKING

from xlsxedit.exceptions import DuplicateWorksheetError, WorksheetNotFoundError
from xlsxedit.opc.constants import CT, RT, SML_NS
from xlsxedit.opc.part import Part, XmlPart
from xlsxedit.oxml.parser import parse_xml, serialize_xml
from xlsxedit.parts import WorksheetPart

if TYPE_CHECKING:
    from xlsxedit.opc.package import OpcPackage
    from xlsxedit.opc.packuri import PackURI
    from xlsxedit.opc.rel import Relationship
    from xlsxedit.workbook import Workbook
    from xlsxedit.worksheet import Worksheet

_DEFINED_NAMES = f"{{{SML_NS}}}definedNames"
_DEFINED_NAME = f"{{{SML_NS}}}definedName"
_SHEETS = f"{{{SML_NS}}}sheets"
_SHEET = f"{{{SML_NS}}}sheet"
_SHEET_VIEW = f"{{{SML_NS}}}sheetView"
_DIGITS_RE = re.compile(r"(\d+)(\.[^/]+)?$")

_SHARE_RELTYPES = {RT.IMAGE}
_SHARE_CONTENT_TYPES = {CT.STYLES, CT.SHARED_STRINGS, CT.THEME, CT.WORKBOOK}


def _partname_template(partname: PackURI) -> str:
    s = str(partname)
    m = _DIGITS_RE.search(s)
    if m:
        return s[: m.start(1)] + "%d" + (m.group(2) or "")
    stem, ext = s.rsplit(".", 1) if "." in s.rsplit("/", 1)[-1] else (s, "")
    if ext:
        return f"{stem}%d.{ext}"
    return s + "%d"


def _should_share(rel: Relationship) -> bool:
    if rel.is_external:
        return False
    if rel.reltype in _SHARE_RELTYPES:
        return True
    part = rel.target_part
    ct = getattr(part, "content_type", "") or ""
    return ct in _SHARE_CONTENT_TYPES or ct.startswith("image/")


def _clone_part(
    package: OpcPackage, source: Part, cloned: dict[int, Part], workbook: Workbook
) -> Part:
    key = id(source)
    if key in cloned:
        return cloned[key]
    partname = package.next_partname(_partname_template(source.partname))
    if isinstance(source, XmlPart):
        dest = type(source)(
            partname, source.content_type, deepcopy(source.element), package
        )
    else:
        dest = Part(partname, source.content_type, source.blob, package)
    package._add_part(dest)
    cloned[key] = dest
    _copy_relationships(source, dest, package, cloned, workbook)
    if source.content_type == CT.TABLE:
        _uniquify_table(dest, workbook)
    return dest


def _copy_relationships(
    source: Part,
    dest: Part,
    package: OpcPackage,
    cloned: dict[int, Part],
    workbook: Workbook,
) -> None:
    for rel in list(source.rels):
        if rel.is_external:
            dest.rels.add_relationship(rel.reltype, rel.target_ref, rel.rId, "External")
            continue
        if _should_share(rel):
            dest.rels.add_relationship(rel.reltype, rel.target_part, rel.rId)
            continue
        target = _clone_part(package, rel.target_part, cloned, workbook)
        dest.rels.add_relationship(rel.reltype, target, rel.rId)


def _uniquify_table(part: Part, workbook: Workbook) -> None:
    elm = parse_xml(part.blob)
    used_ids: set[int] = set()
    used_names: set[str] = set()
    for ws in workbook.worksheets:
        for tbl in ws.tables:
            with suppress(TypeError, ValueError):
                used_ids.add(int(tbl._element.get("id", "0")))
            if tbl.name:
                used_names.add(tbl.name)
            disp = tbl._element.get("displayName")
            if disp:
                used_names.add(disp)
    new_id = max(used_ids, default=0) + 1
    elm.set("id", str(new_id))
    name = f"Table{new_id}"
    n = new_id
    while name in used_names:
        n += 1
        name = f"Table{n}"
    elm.set("name", name)
    elm.set("displayName", name)
    part._blob = serialize_xml(elm)


def clone_sheet_relationships(
    source_part: WorksheetPart, dest_part: WorksheetPart, workbook: Workbook
) -> None:
    """Copy sheet rels onto ``dest_part`` (same rIds; share images, clone the rest)."""
    cloned: dict[int, Part] = {}
    _copy_relationships(source_part, dest_part, workbook._package, cloned, workbook)


def copy_local_defined_names(
    workbook: Workbook, source_index: int, dest_index: int
) -> None:
    """Duplicate definedNames with ``localSheetId`` equal to ``source_index``."""
    wb_elm = workbook._workbook_part.element
    block = wb_elm.find(_DEFINED_NAMES)
    if block is None:
        return
    for elm in list(block.findall(_DEFINED_NAME)):
        local_id = elm.get("localSheetId")
        if local_id is None or int(local_id) != source_index:
            continue
        clone = deepcopy(elm)
        clone.set("localSheetId", str(dest_index))
        block.append(clone)


def copy_worksheet(
    workbook: Workbook, name: str, new_name: str, *, after: str | None = None
) -> Worksheet:
    """Duplicate a worksheet and place the copy where ``after`` says.

    Upstream's own method appends the copy to the end of the workbook and
    offers no way to say otherwise, so this repositions afterwards. It
    calls upstream when the installed version has the method, and runs
    the port of it when it does not.

    Args:
        workbook: The workbook to work in.
        name: The sheet to copy.
        new_name: The copy's tab name.
        after: The sheet the copy goes directly after; ``None`` leaves it
            at the end.

    Returns:
        The copy.
    """
    upstream = getattr(type(workbook), "copy_worksheet", None)
    copy = (
        upstream(workbook, name, new_name)
        if upstream is not None
        else _copy_worksheet(workbook, name, new_name)
    )
    _deselect(copy)
    if after is not None:
        move_worksheet(workbook, new_name, after=after)
        copy = workbook[new_name]
    return copy


def _deselect(worksheet: Worksheet) -> None:
    """Take the selection off a copy.

    Copying a sheet copies its ``<sheetView>`` with it, selection and
    all, so a workbook comes out claiming two tabs are the one in front.
    """
    for view in worksheet._part.element.iter(_SHEET_VIEW):
        view.attrib.pop("tabSelected", None)


def move_worksheet(workbook: Workbook, name: str, *, after: str) -> None:
    """Move a sheet's tab so it sits directly after ``after``."""
    workbook_part = workbook._workbook_part
    sheets = workbook_part.element.find(_SHEETS)
    if sheets is None:
        return
    moving = anchor = None
    for elm in sheets.findall(_SHEET):
        if elm.get("name") == name:
            moving = elm
        if elm.get("name") == after:
            anchor = elm
    if moving is None or anchor is None or moving is anchor:
        return
    sheets.remove(moving)
    sheets.insert(list(sheets).index(anchor) + 1, moving)
    workbook._refresh_sheets()


def _copy_worksheet(workbook: Workbook, name: str, new_name: str) -> Worksheet:
    """Duplicate an existing worksheet (cells, styles, merges, sheet-owned parts)."""
    sheets_by_name = workbook._sheets_by_name
    if new_name in sheets_by_name:
        raise DuplicateWorksheetError(f"worksheet {new_name!r} already exists")
    if name not in sheets_by_name:
        raise WorksheetNotFoundError(f"worksheet {name!r} not found")

    package = workbook._package
    workbook_part = workbook._workbook_part
    source = sheets_by_name[name]
    source_index = list(workbook.sheetnames).index(name)
    element = deepcopy(source._part.element)
    partname = package.next_partname("/xl/worksheets/sheet%d.xml")
    part = WorksheetPart(
        partname,
        source._part.content_type,
        element,
        package,
    )
    package._add_part(part)
    clone_sheet_relationships(source._part, part, workbook)

    r_id = workbook_part.relate_to(part, RT.WORKSHEET)
    sheet_id = workbook_part.next_sheet_id()
    workbook_part.append_sheet_element(new_name, sheet_id, r_id)
    copy_local_defined_names(workbook, source_index, len(workbook._sheets))
    workbook._refresh_sheets()
    return workbook._sheets_by_name[new_name]
