"""Workbook-level orchestration of the fill run."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, cast

from _patched_xlsxedit import (
    RT,
    PicturePlacement,
    SharedStringTable,
    add_hyperlinks,
    emit_attrs,
    escape_attr,
    escape_go_text,
    escape_text,
    expand_anchors,
    insert_worksheet_child,
    parse_attrs,
    remap_defined_name,
    remap_table,
    unescape,
)
from xlsxfill._containers import substitute_container, substitute_text_nodes
from xlsxfill._engine import SheetEngine, instances_of, root_mapper
from xlsxfill._problems import BookProblem, CellProblem, Problem, SheetProblem
from xlsxfill._resolve import collection_length
from xlsxfill._sheetxml import SheetDoc
from xlsxfill._syntax import Image, IndexStep, Link, ParsedText, ValueRef, tokenize

if TYPE_CHECKING:
    from collections.abc import Callable, Mapping
    from typing import Literal

    from _patched_xlsxedit import EndpointMapper, Workbook
    from xlsxfill._syntax import PathStep
    from xlsxfill._values import Value

_SHEET_TAG_RE = re.compile(r"<sheet((?:\s[^>]*)?)></sheet>")
_SHEETS_RE = re.compile(r"(<sheets>).*(</sheets>)", re.DOTALL)
_CORE_TEXT_TAGS = (
    "dc:title",
    "dc:subject",
    "dc:creator",
    "dc:description",
    "cp:keywords",
    "cp:category",
    "lastModifiedBy",
)

_SST_MEMBER = "xl/sharedStrings.xml"
_WORKBOOK_MEMBER = "xl/workbook.xml"
_CORE_MEMBER = "docProps/core.xml"


@dataclass
class SheetInstance:
    """One output sheet: an original, or one copy of a repeated sheet."""

    name: str
    member: str
    bindings: dict[str, int] = field(default_factory=dict)


def _name_steps(name_raw: str) -> list[object]:
    steps: list[object] = []
    for segment in tokenize(unescape(name_raw)).constructs:
        path = getattr(segment, "path", None)
        if path:
            steps.extend(path)
    return steps


def _strip_tab_selected(raw: bytes) -> bytes:
    return raw.replace(b' tabSelected="true"', b"", 1)


class Book:
    """One fill run over an open workbook."""

    def __init__(self, wb: Workbook, data: Mapping[str, Value]) -> None:
        """Prepare the run against ``wb`` with input ``data``."""
        self.wb = wb
        self.pkg = wb.package
        self.data = data
        self.sst = SharedStringTable(self.pkg.read(_SST_MEMBER))
        self.problems: list[Problem] = []
        self.renames: list[tuple[str, str]] = []
        self._parsed_cache: dict[int, ParsedText] = {}
        self._template_names: dict[str, str] = {}

    # -------------------------------------------------------------- caches

    def parsed_sst(self, index: int) -> ParsedText:
        """Tokenized text of shared-string entry ``index`` (cached)."""
        cached = self._parsed_cache.get(index)
        if cached is None:
            cached = tokenize(self.sst.text(index))
            self._parsed_cache[index] = cached
        return cached

    # ------------------------------------------------------------ problems

    def report_cell(
        self,
        kind: str,
        construct: str,
        reason: str,
        sheet: str,
        cell: str,
    ) -> str:
        """Record a cell problem; return the message to embed."""
        problem = CellProblem(
            kind=cast("Literal['syntax', 'data']", kind),
            construct=construct,
            reason=reason,
            sheet=sheet,
            cell=cell,
            part="cell",
        )
        self.problems.append(problem)
        return problem.message

    def sheet_reporter(
        self, sheet: str, part: str = "header_footer"
    ) -> Callable[[str, str, str], str]:
        """Build a problem sink for sheet-level string containers."""

        def report(kind: str, construct: str, reason: str) -> str:
            problem = SheetProblem(
                kind=cast("Literal['syntax', 'data']", kind),
                construct=construct,
                reason=reason,
                sheet=sheet,
                part=cast(
                    "Literal['sheet_name', 'header_footer', 'shape', 'chart']",
                    part,
                ),
            )
            self.problems.append(problem)
            return problem.message

        return report

    def book_reporter(self) -> Callable[[str, str, str], str]:
        """Build a problem sink for workbook-level string containers."""

        def report(kind: str, construct: str, reason: str) -> str:
            problem = BookProblem(
                kind=cast("Literal['syntax', 'data']", kind),
                construct=construct,
                reason=reason,
                part="doc_props",
            )
            self.problems.append(problem)
            return problem.message

        return report

    # -------------------------------------------------------- link / image

    def _attach_links_images(self, engine: SheetEngine) -> None:
        """Materialize collected hyperlinks and images for one sheet."""
        if not engine.pending_links and not engine.pending_images:
            return
        rels = self.pkg.rels(engine.member)

        hyperlink_tags = []
        for addr, url in engine.pending_links:
            rid = rels.next_rId()
            rels.add_relationship(RT.HYPERLINK, url, rid, target_mode="External")
            hyperlink_tags.append(f'<hyperlink ref="{addr}" r:id="{rid}"></hyperlink>')
        if hyperlink_tags:
            engine.doc.post = add_hyperlinks(engine.doc.post, hyperlink_tags)

        if engine.pending_images:
            drawing_member = self.wb.add_pictures(
                [
                    PicturePlacement(
                        image.col,
                        image.row,
                        image.alt,
                        image.data,
                        image.fit,
                        engine.column_width(image.col),
                        engine.row_height(image.row),
                    )
                    for image in engine.pending_images
                ]
            )
            rid = rels.next_rId()
            rels.add_relationship(RT.DRAWING, drawing_member, rid)
            engine.doc.post = insert_worksheet_child(
                engine.doc.post, f'<drawing r:id="{rid}"></drawing>', "drawing"
            )
        rels.save()

    # ------------------------------------------------------------ phases

    def run(self) -> list[Problem]:
        """Execute the whole fill and return the collected problems."""
        instances = self._plan_sheets()

        engines = []
        for instance in instances:
            engine = SheetEngine(self, instance.name, instance.member)
            engine.base_bindings = dict(instance.bindings)
            engines.append(engine)
        for engine in engines:
            engine.process()
        for engine in engines:
            self._attach_links_images(engine)
            self._update_tables(engine)
            self._process_shape_drawings(engine)
        self._update_defined_names(engines)
        reference_count = sum(e.string_reference_count() for e in engines)
        for engine in engines:
            self.pkg.write(engine.member, engine.emit())
        self.pkg.write(_SST_MEMBER, self.sst.emit(reference_count))

        sheets = [(i.name, i.member) for i in instances]
        self._substitute_related_parts(sheets)
        self._substitute_core()
        self.wb.propagate_sheet_renames(self.renames, [member for _, member in sheets])
        return self.problems

    # ------------------------------------------------------------ followers

    def _update_tables(self, engine: SheetEngine) -> None:
        if not engine.axes_changed:
            return
        for rel in self.pkg.rels(engine.member):
            if rel.is_external or rel.reltype != RT.TABLE:
                continue
            member = rel.target_partname.membername
            raw = self.pkg.read(member).decode()
            updated = remap_table(raw, engine.remap_ref, engine.cell_text)
            if updated is None:
                continue
            self.pkg.write(member, updated.encode())

    def _update_defined_names(self, engines: list[SheetEngine]) -> None:
        raw = self.pkg.read(_WORKBOOK_MEMBER).decode()
        if "<definedName" not in raw:
            return
        mappers_for: dict[str, tuple[EndpointMapper, EndpointMapper]] = {}
        for engine in engines:
            mappers = (root_mapper(engine.vaxis), root_mapper(engine.haxis))
            mappers_for[engine.sheet_name] = mappers
            template = self._template_names.get(engine.member)
            if template is not None:
                mappers_for.setdefault(template, mappers)

        def _sub(m: re.Match[str]) -> str:
            text = unescape(m.group(2))
            remapped = remap_defined_name(text, mappers_for)
            if remapped == text:
                return m.group(0)
            return f"{m.group(1)}{escape_go_text(remapped)}{m.group(3)}"

        updated = re.sub(
            r"(<definedName(?:\s[^>]*)?>)(.*?)(</definedName>)",
            _sub,
            raw,
            flags=re.DOTALL,
        )
        if updated != raw:
            self.pkg.write(_WORKBOOK_MEMBER, updated.encode())

    def _process_shape_drawings(self, engine: SheetEngine) -> None:
        if not engine.axes_changed:
            return

        def placements(from_row: int, from_col: int) -> list[tuple[int, int]]:
            vertical = instances_of(engine.vaxis, from_row, from_row)
            horizontal = instances_of(engine.haxis, from_col, from_col)
            pairs = (
                [(v, h) for v in vertical for h in horizontal]
                if engine.outer_vertical_major()
                else [(v, h) for h in horizontal for v in vertical]
            )
            return [(v[0], h[0]) for v, h in pairs]

        for rel in self.pkg.rels(engine.member):
            if rel.is_external or rel.reltype != RT.DRAWING:
                continue
            member = rel.target_partname.membername
            raw = self.pkg.read(member).decode()
            updated = expand_anchors(raw, placements)
            if updated != raw:
                self.pkg.write(member, updated.encode())

    # ---------------------------------------------------- sheet repetition

    def _sheet_paths(self, member: str, name_raw: str) -> list[tuple[PathStep, ...]]:
        """All placeholder paths on a sheet (tab name, cells, sections)."""
        paths: list[tuple[PathStep, ...]] = []
        doc = SheetDoc(self.pkg.read(member))
        parsed_texts = [tokenize(unescape(name_raw))]
        for row in doc.rows:
            for cell in row.cells:
                index = cell.sst_index
                if index is not None:
                    parsed_texts.append(self.parsed_sst(index))
        parsed_texts.append(tokenize(doc.pre))
        parsed_texts.append(tokenize(doc.post))
        for parsed in parsed_texts:
            for segment in parsed.constructs:
                candidates: list[tuple[PathStep, ...] | None]
                if isinstance(segment, Link):
                    candidates = [segment.label_path, segment.url_path]
                elif isinstance(segment, Image):
                    candidates = [segment.alt_path, segment.data_path]
                elif isinstance(segment, ValueRef):
                    candidates = [segment.path]
                else:
                    candidates = []
                paths.extend(p for p in candidates if p is not None)
        return paths

    def _plan_sheets(self) -> list[SheetInstance]:
        workbook_raw = self.pkg.read(_WORKBOOK_MEMBER).decode()
        members = [member for _, member in self.wb.sheet_members()]
        tags = _SHEET_TAG_RE.findall(workbook_raw)

        instances: list[SheetInstance] = []
        out_tags: list[str] = []
        for tag, member in zip(tags, members, strict=True):
            attrs = parse_attrs(tag)
            attr_map = dict(attrs)
            name_raw = attr_map["name"]
            self._template_names[member] = unescape(name_raw)
            paths = self._sheet_paths(member, name_raw)
            uses_s = any(
                isinstance(step, IndexStep) and step.is_sheet
                for path in paths
                for step in path  # type: ignore[attr-defined]
            )
            report = self.sheet_reporter(unescape(name_raw), part="sheet_name")
            if not uses_s:
                replaced = substitute_container(
                    name_raw, self.data, {}, escape_attr, report
                )
                out_attrs = attrs
                if replaced is not None:
                    self.renames.append((unescape(name_raw), unescape(replaced)))
                    out_attrs = [(n, replaced if n == "name" else v) for n, v in attrs]
                out_tags.append(f"<sheet{emit_attrs(out_attrs)}></sheet>")
                instances.append(
                    SheetInstance(unescape(dict(out_attrs)["name"]), member)
                )
                continue

            count = self._sheet_count(paths)
            if count == 0:
                self.wb.remove_worksheet_part(member, attr_map["r:id"])
                continue
            name_has_s = any(
                isinstance(step, IndexStep) and step.is_sheet
                for step in _name_steps(name_raw)
            )
            source_bytes = self.pkg.read(member)
            for i in range(count):
                if name_has_s:
                    new_name_raw = (
                        substitute_container(
                            name_raw, self.data, {"s": i}, escape_attr, report
                        )
                        or name_raw
                    )
                else:
                    base_raw = (
                        substitute_container(
                            name_raw, self.data, {"s": i}, escape_attr, report
                        )
                        or name_raw
                    )
                    new_name_raw = f"{base_raw} ({i + 1})"
                if i == 0:
                    instance_member = member
                    sheet_id = attr_map["sheetId"]
                    rid = attr_map["r:id"]
                else:
                    instance_member, rid, sheet_id_n = self.wb.add_worksheet_part(
                        _strip_tab_selected(source_bytes)
                    )
                    sheet_id = str(sheet_id_n)
                out_attrs = [
                    (
                        n,
                        new_name_raw
                        if n == "name"
                        else sheet_id
                        if n == "sheetId"
                        else rid
                        if n == "r:id"
                        else v,
                    )
                    for n, v in attrs
                ]
                out_tags.append(f"<sheet{emit_attrs(out_attrs)}></sheet>")
                instances.append(
                    SheetInstance(unescape(new_name_raw), instance_member, {"s": i})
                )

        updated = _SHEETS_RE.sub(
            lambda m: f"{m.group(1)}{''.join(out_tags)}{m.group(2)}",
            workbook_raw,
            count=1,
        )
        if updated != workbook_raw:
            self.pkg.write(_WORKBOOK_MEMBER, updated.encode())
        return instances

    def _sheet_count(self, paths: list[tuple[PathStep, ...]]) -> int:
        lengths = [
            length
            for path in paths
            if (length := collection_length(path, "s", self.data, {})) is not None
        ]
        return min(lengths) if lengths else 0

    # --------------------------------------------------- related substitution

    def _substitute_member_text_nodes(
        self,
        member: str,
        tags: tuple[str, ...],
        report,  # noqa: ANN001
    ) -> None:
        raw = self.pkg.read(member).decode()
        updated = substitute_text_nodes(raw, tags, self.data, {}, escape_text, report)
        if updated != raw:
            self.pkg.write(member, updated.encode())

    def _substitute_related_parts(self, sheets: list[tuple[str, str]]) -> None:
        for sheet_name, member in sheets:
            rels = self.pkg.rels(member)
            for rel in rels:
                if rel.is_external:
                    continue
                target = rel.target_partname.membername
                if rel.reltype == RT.COMMENTS:
                    self._substitute_member_text_nodes(
                        target,
                        ("t",),
                        self.sheet_reporter(sheet_name, part="shape"),
                    )
                elif rel.reltype == RT.DRAWING:
                    self._substitute_member_text_nodes(
                        target,
                        ("a:t",),
                        self.sheet_reporter(sheet_name, part="shape"),
                    )
                    for drawing_rel in self.pkg.rels(target):
                        if (
                            not drawing_rel.is_external
                            and drawing_rel.reltype == RT.CHART
                        ):
                            self._substitute_member_text_nodes(
                                drawing_rel.target_partname.membername,
                                ("a:t",),
                                self.sheet_reporter(sheet_name, part="chart"),
                            )

    def _substitute_core(self) -> None:
        if _CORE_MEMBER not in self.pkg:
            return
        self._substitute_member_text_nodes(
            _CORE_MEMBER, _CORE_TEXT_TAGS, self.book_reporter()
        )
