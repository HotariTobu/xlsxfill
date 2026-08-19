"""Per-sheet substitution engine: markers, bands, and cell values."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import TYPE_CHECKING, cast

from _patched_xlsxedit import (
    col_to_index,
    escape_attr,
    escape_go_text,
    escape_text,
    index_to_col,
    parse_range,
    rebase_formula,
    remap_range_ref,
    remap_sqref,
    unescape,
)
from xlsxfill._bands import Band, BandLayout, MarkerAt, build_layout, mark_unused
from xlsxfill._containers import substitute_attrs, substitute_text_nodes
from xlsxfill._resolve import (
    ResolveError,
    check_assertion,
    check_scalar,
    collection_length,
    concat_text,
    render_number,
    resolve,
)
from xlsxfill._sheetxml import SheetDoc
from xlsxfill._syntax import (
    Image,
    IndexStep,
    Link,
    Literal,
    Marker,
    ParsedText,
    ValueRef,
)

if TYPE_CHECKING:
    from collections.abc import Mapping

    from _patched_xlsxedit import EndpointMapper, PointMapper
    from xlsxfill._book import Book
    from xlsxfill._sheetxml import Cell, Row
    from xlsxfill._syntax import PathStep
    from xlsxfill._values import Value

_HEADER_FOOTER_TAGS = (
    "oddHeader",
    "oddFooter",
    "evenHeader",
    "evenFooter",
    "firstHeader",
    "firstFooter",
)
_VALIDATION_ATTRS = ("error", "errorTitle", "prompt", "promptTitle")


@dataclass(frozen=True)
class Slot:
    """One output line (row or column) of an expansion."""

    src: int
    bindings: dict[str, int]


@dataclass
class Instance:
    """One instantiated template region (a band block or the whole sheet)."""

    lo: int
    hi: int
    start: int
    stop: int
    children: list[Instance] = field(default_factory=list)


class Axis:
    """The expansion result along one axis (rows or columns).

    Maps template coordinates to output coordinates with Excel-compatible
    insert/copy/delete semantics: relative references resolve inside the
    innermost band block containing the referring cell, absolute
    references resolve globally.
    """

    def __init__(self, slots: list[Slot], root: Instance, base: int) -> None:
        """Wrap the ``slots`` and instance tree ``root``; ``base`` offsets positions."""
        self.slots = slots
        self.root = root
        self.base = base

    def chain_for(self, position: int) -> list[Instance]:
        """The instance path (outermost first) containing a slot position."""
        chain = [self.root]
        while True:
            child = next(
                (c for c in chain[-1].children if c.start <= position < c.stop),
                None,
            )
            if child is None:
                return chain
            chain.append(child)

    def map_point(
        self,
        coord: int,
        chain: list[Instance],
        *,
        absolute: bool,
        end: bool,
    ) -> int:
        """Map a template coordinate referenced from a cell in ``chain``."""
        instances = [self.root] if absolute else list(reversed(chain))
        for instance in instances:
            if instance.lo <= coord <= instance.hi:
                return self.base + self._find(instance, coord, end=end)
        if coord > self.root.hi:
            covered = self.root.hi - self.root.lo + 1
            return coord + len(self.slots) - covered
        return coord

    def _find(self, instance: Instance, coord: int, *, end: bool) -> int:
        positions = range(instance.start, instance.stop)
        hits = [p for p in positions if self.slots[p].src == coord]
        if hits:
            # Copies are INSERTED at the removed marker line, so references
            # to a surviving template line always keep its first copy.
            return hits[0]
        if end:
            below = [p for p in positions if self.slots[p].src < coord]
            return below[-1] if below else instance.start - 1
        above = [p for p in positions if self.slots[p].src > coord]
        return above[0] if above else instance.stop

    @property
    def changed(self) -> bool:
        """Whether the axis differs from the identity mapping."""
        return (
            any(
                slot.src != self.base + position
                for position, slot in enumerate(self.slots)
            )
            or self.base + len(self.slots) != self.root.hi + 1
        )


def map_root_point(axis: Axis | None, coord: int, *, end: bool) -> int:
    """Map one coordinate through an axis with insert/delete semantics."""
    if axis is None:
        return coord
    return axis.map_point(coord, [axis.root], absolute=True, end=end)


def root_mapper(axis: Axis | None) -> EndpointMapper:
    """Wrap an axis as an endpoint mapper for range remapping."""

    def mapper(coord: int, *, end: bool) -> int:
        return map_root_point(axis, coord, end=end)

    return mapper


def instances_of(axis: Axis | None, lo: int, hi: int) -> list[tuple[int, int]]:
    """All output copies of the template span ``[lo, hi]`` on one axis.

    Returns ``(start, end)`` output coordinates in generation order. The
    span must not straddle a band boundary.
    """
    if axis is None:
        return [(lo, hi)]
    result: list[tuple[int, int]] = []

    def visit(instance: Instance) -> None:
        containing = [
            child for child in instance.children if child.lo <= lo and hi <= child.hi
        ]
        if containing:
            for child in containing:
                visit(child)
            return
        child_ranges = [(c.start, c.stop) for c in instance.children]

        def in_child(position: int) -> bool:
            return any(start <= position < stop for start, stop in child_ranges)

        start_pos = end_pos = None
        for position in range(instance.start, instance.stop):
            if in_child(position):
                continue
            if axis.slots[position].src == lo and start_pos is None:
                start_pos = position
            if axis.slots[position].src == hi:
                end_pos = position
        if start_pos is not None and end_pos is not None:
            result.append((axis.base + start_pos, axis.base + end_pos))

    visit(axis.root)
    return result


@dataclass(frozen=True)
class PendingImage:
    """One image to be anchored at a cell."""

    col: int
    row: int
    alt: str | None
    data: bytes
    fit: str


class SheetEngine:
    """Processes one worksheet part."""

    def __init__(self, book: Book, sheet_name: str, member: str) -> None:
        """Prepare the engine for worksheet ``member`` named ``sheet_name``."""
        self.book = book
        self.sheet_name = sheet_name
        self.member = member
        self.doc = SheetDoc(book.pkg.read(member))
        self.layout = BandLayout(
            roots_by_direction={True: [], False: []},
            marker_reasons={},
            bands=[],
        )
        self.vaxis: Axis | None = None
        self.haxis: Axis | None = None
        self._count_memo: dict[tuple[int, tuple[tuple[str, int], ...]], int | None] = {}
        self._band_refs: dict[int, list[tuple[tuple[PathStep, ...], int, int]]] = {}
        self.base_bindings: dict[str, int] = {}
        self.pending_links: list[tuple[str, str]] = []
        self.pending_images: list[PendingImage] = []

    # ------------------------------------------------------------- parsing

    def _parsed(self, cell: Cell) -> ParsedText | None:
        index = cell.sst_index
        if index is None:
            return None
        return self.book.parsed_sst(index)

    def _iter_cells(self) -> list[tuple[Row, Cell]]:
        return [(row, cell) for row in self.doc.rows for cell in row.cells]

    # ------------------------------------------------------- band analysis

    def _collect_markers(self) -> list[MarkerAt]:
        markers: list[MarkerAt] = []
        for row, cell in self._iter_cells():
            parsed = self._parsed(cell)
            if parsed is None:
                continue
            for seg, segment in enumerate(parsed.segments):
                if isinstance(segment, Marker):
                    markers.append(MarkerAt(segment, row.num, cell.col, seg))
        return markers

    def _collect_band_refs(self, layout: BandLayout) -> dict[int, bool]:
        """Index value references by band; return per-band used flags."""
        self._band_refs = {id(band): [] for band in layout.bands}
        used: dict[int, bool] = {}
        seen: set[str] = set()
        for row, cell in self._iter_cells():
            parsed = self._parsed(cell)
            if parsed is None:
                continue
            for segment in parsed.constructs:
                for path in _paths_of(segment):
                    for step in path:
                        if (
                            not isinstance(step, IndexStep)
                            or step.is_fixed
                            or step.is_sheet
                        ):
                            continue
                        band = layout.band_for(
                            step.symbol, row.num, cell.col, seen=seen
                        )
                        if band is None:
                            continue
                        used[id(band)] = True
                        self._band_refs[id(band)].append((path, row.num, cell.col))
        return used

    def _annotate_markers(self, layout: BandLayout) -> None:
        """Store marker outcomes on the cells before expansion moves them."""
        for row, cell in self._iter_cells():
            parsed = self._parsed(cell)
            if parsed is None:
                continue
            for seg, segment in enumerate(parsed.segments):
                if isinstance(segment, Marker):
                    reason = layout.marker_reasons.get((row.num, cell.col, seg))
                    cell.notes[seg] = reason

    def _check_merge_straddle(self) -> None:
        """Invalidate bands whose boundary is straddled by a merged cell."""
        refs = re.findall(r'<mergeCell ref="([^"]*)"></mergeCell>', self.doc.post)
        if not refs:
            return
        dropped: set[int] = set()
        for ref in refs:
            c1, r1, c2, r2 = parse_range(ref)
            spans = {
                True: (r1, r2),
                False: (col_to_index(c1), col_to_index(c2)),
            }
            for band in self.layout.bands:
                lo, hi = spans[band.vertical]
                intersects = not (hi < band.start or lo > band.stop - 1)
                contained = band.start <= lo and hi <= band.stop - 1
                if intersects and not contained:
                    dropped.add(id(band))
        self.layout.drop(dropped, "merged cell straddles the band boundary")

    # ------------------------------------------------------------ counting

    def band_count(self, band: Band, bindings: Mapping[str, int]) -> int | None:
        """Iteration count of ``band`` under ``bindings``.

        Returns ``None`` when the band never instantiates because a
        crossing parent has zero blocks; the band is then left untouched.
        """
        key = (id(band), tuple(sorted(bindings.items())))
        if key in self._count_memo:
            return self._count_memo[key]
        refs = self._band_refs.get(id(band), [])
        contexts = self._contexts(band, refs, dict(bindings))
        result: int | None = None
        for context in contexts:
            lengths = [
                length
                for path, _, _ in refs
                if (
                    length := collection_length(
                        path, band.name, self.book.data, context
                    )
                )
                is not None
            ]
            if lengths:
                block = min(lengths)
                result = block if result is None else max(result, block)
        self._count_memo[key] = result
        return result

    def _contexts(
        self,
        band: Band,
        refs: list[tuple[tuple[PathStep, ...], int, int]],
        bindings: dict[str, int],
    ) -> list[dict[str, int]]:
        """Binding contexts to evaluate: crossing bands enumerate to max."""
        needed: list[Band] = []
        seen: set[str] = set()
        for path, row, col in refs:
            for step in path:
                if isinstance(step, IndexStep) and step.symbol == band.name:
                    break
                if (
                    isinstance(step, IndexStep)
                    and not step.is_fixed
                    and not step.is_sheet
                    and step.symbol not in bindings
                ):
                    other = self.layout.band_for(step.symbol, row, col, seen=seen)
                    if other is not None and other not in needed:
                        needed.append(other)
        contexts = [bindings]
        for other in needed:
            expanded: list[dict[str, int]] = []
            for context in contexts:
                count = self.band_count(other, context)
                if not count:
                    # A crossing band without blocks binds nothing; the
                    # dependent references simply fail to resolve.
                    expanded.append(context)
                    continue
                expanded.extend({**context, other.name: i} for i in range(count))
            contexts = expanded
        return contexts

    # ----------------------------------------------------------- expansion

    def _make_axis(self, roots: list[Band], lo: int, hi: int) -> Axis:
        slots: list[Slot] = []

        def emit(
            lo: int, hi: int, bands: list[Band], bindings: dict[str, int]
        ) -> Instance:
            instance = Instance(lo, hi, len(slots), 0)
            pos = lo
            for band in sorted(bands, key=lambda b: b.start):
                slots.extend(Slot(x, bindings) for x in range(pos, band.start))
                count = self.band_count(band, bindings)
                if count is None:
                    slots.extend(
                        Slot(x, bindings) for x in range(band.start, band.marker_stop)
                    )
                else:
                    for i in range(count):
                        instance.children.append(
                            emit(
                                band.start,
                                band.stop - 1,
                                band.children,
                                {**bindings, band.name: i},
                            )
                        )
                pos = band.marker_stop
            slots.extend(Slot(x, bindings) for x in range(pos, hi + 1))
            instance.stop = len(slots)
            return instance

        root = emit(lo, hi, roots, dict(self.base_bindings))
        return Axis(slots, root, lo)

    def _apply_vertical(self, axis: Axis) -> None:
        rows_by_num = {row.num: row for row in self.doc.rows}
        new_rows: list[Row] = []
        for index, slot in enumerate(axis.slots):
            source = rows_by_num.get(slot.src)
            if source is None:
                continue
            row = source.copy()
            row.set_num(axis.base + index)
            for cell in row.cells:
                cell.bindings.update(slot.bindings)
            new_rows.append(row)
        self.doc.rows = new_rows

    def _apply_horizontal(self, axis: Axis) -> None:
        for row in self.doc.rows:
            by_col = {cell.col: cell for cell in row.cells}
            cells = []
            for index, slot in enumerate(axis.slots):
                source = by_col.get(slot.src)
                if source is None:
                    continue
                cell = source.copy()
                cell.set_address(axis.base + index, row.num)
                cell.bindings.update(slot.bindings)
                cells.append(cell)
            row.cells = cells

    def _expand_cols(self, axis: Axis) -> None:
        cols = self.doc.cols()
        if not cols:
            return
        props: dict[int, dict[str, str]] = {}
        for attrs in cols:
            mapping = dict(attrs)
            core = {
                name: value
                for name, value in mapping.items()
                if name not in ("min", "max")
            }
            for col in range(int(mapping["min"]) - 1, int(mapping["max"])):
                props[col] = core
        entries: list[tuple[dict[str, str], int, int]] = []
        max_covered = max(props, default=-1)
        sources = [slot.src for slot in axis.slots]
        sources.extend(
            range(self.root_hi(axis) + 1, max_covered + 1),
        )
        for index, src in enumerate(sources):
            core = props.get(src)
            if core is None:
                continue
            excel_col = index + 1
            if entries and entries[-1][0] == core and entries[-1][2] == excel_col - 1:
                entries[-1] = (core, entries[-1][1], excel_col)
            else:
                entries.append((core, excel_col, excel_col))
        out = [
            sorted(
                [*core.items(), ("max", str(hi)), ("min", str(lo))],
            )
            for core, lo, hi in entries
        ]
        self.doc.set_cols(out)

    @staticmethod
    def root_hi(axis: Axis) -> int:
        """The last template coordinate the axis covered."""
        return axis.root.hi

    # ----------------------------------------------------- formula rebasing

    @staticmethod
    def _axis_mapper(axis: Axis | None, chain: list[Instance] | None) -> PointMapper:
        def mapper(coord: int, *, absolute: bool, end: bool) -> int:
            if axis is None or chain is None:
                return coord
            return axis.map_point(coord, chain, absolute=absolute, end=end)

        return mapper

    def _rebase_formulas(self) -> None:
        if not self.axes_changed:
            return
        for row in self.doc.rows:
            vchain = (
                self.vaxis.chain_for(row.num - self.vaxis.base) if self.vaxis else None
            )
            for cell in row.cells:
                raw = cell.formula_text
                if raw is None:
                    continue
                hchain = (
                    self.haxis.chain_for(cell.col - self.haxis.base)
                    if self.haxis
                    else None
                )
                map_row = self._axis_mapper(self.vaxis, vchain)
                map_col = self._axis_mapper(self.haxis, hchain)
                text = unescape(raw)
                rebased = rebase_formula(text, map_col, map_row)
                if rebased != text:
                    cell.set_formula_text(escape_go_text(rebased))

    # -------------------------------------------------------- substitution

    def _report_cell(
        self, kind: str, construct: str, reason: str, cell_addr: str
    ) -> str:
        return self.book.report_cell(
            kind, construct, reason, self.sheet_name, cell_addr
        )

    def _eval_ref(self, ref: ValueRef, cell: Cell, addr: str) -> tuple[str, object]:
        if ref.error is not None:
            return ("message", self._report_cell("syntax", ref.src, ref.error, addr))
        for step in ref.path:
            if (
                isinstance(step, IndexStep)
                and not step.is_fixed
                and not step.is_sheet
                and step.symbol not in cell.bindings
                and step.symbol not in self.base_bindings
            ):
                reason = self._unbound_reason(step.symbol)
                return (
                    "message",
                    self._report_cell("syntax", ref.src, reason, addr),
                )
        bindings = {**self.base_bindings, **cell.bindings}
        try:
            value = resolve(ref.path, self.book.data, bindings)
            if value is not None:
                if ref.assert_type is not None:
                    check_assertion(value, ref.assert_type)
                check_scalar(value)
                if isinstance(value, datetime) and value.tzinfo is not None:
                    render_number(value)
        except ResolveError as error:
            return (
                "message",
                self._report_cell("data", ref.src, error.reason, addr),
            )
        return ("value", value)

    def _cell_is_starved(self, cell: Cell) -> bool:
        """Whether a crossing-band block leaves this cell empty."""
        for name, index in cell.bindings.items():
            candidates = [b for b in self.layout.bands if b.name == name]
            if len(candidates) != 1:
                continue
            context = {k: v for k, v in cell.bindings.items() if k != name}
            count = self.band_count(candidates[0], context)
            if count is not None and index >= count:
                return True
        return False

    def _substitute_cells(self) -> None:
        removed: set[int] = set()
        for row in self.doc.rows:
            for cell in row.cells:
                if not self._substitute_cell(row, cell):
                    removed.add(id(cell))
        if removed:
            for row in self.doc.rows:
                row.cells = [c for c in row.cells if id(c) not in removed]

    def _resolve_link_part(
        self,
        path: tuple[PathStep, ...],
        cell: Cell,
        src: str,
        addr: str,
    ) -> tuple[bool, object]:
        """Resolve one link/image sub-path; returns ``(ok, value-or-message)``."""
        for step in path:
            if (
                isinstance(step, IndexStep)
                and not step.is_fixed
                and not step.is_sheet
                and step.symbol not in cell.bindings
                and step.symbol not in self.base_bindings
            ):
                reason = self._unbound_reason(step.symbol)
                return (False, self._report_cell("syntax", src, reason, addr))
        bindings = {**self.base_bindings, **cell.bindings}
        try:
            return (True, resolve(path, self.book.data, bindings))
        except ResolveError as error:
            return (
                False,
                self._report_cell("data", src, error.reason, addr),
            )

    def _unbound_reason(self, token: str) -> str:
        if token in self.layout.names:
            return f'band "{token}" used outside its band'
        return f'band "{token}" is not declared'

    def _eval_link(self, link: Link, cell: Cell, addr: str) -> tuple[str, object]:
        if link.error is not None:
            return (
                "message",
                self._report_cell("syntax", link.src, link.error, addr),
            )
        ok, url = self._resolve_link_part(link.url_path, cell, link.src, addr)
        if not ok:
            return ("message", url)
        label_piece: tuple[str, object] = ("skip", None)
        if link.label_path is not None:
            ok, label = self._resolve_link_part(link.label_path, cell, link.src, addr)
            if not ok:
                return ("message", label)
            label_piece = ("text", concat_text(cast("Value", label)))
        if isinstance(url, str) and url != "":
            self.pending_links.append((addr, url))
        return label_piece

    def _eval_image(
        self, image: Image, cell: Cell, row_num: int, addr: str
    ) -> tuple[str, object]:
        if image.error is not None:
            return (
                "message",
                self._report_cell("syntax", image.src, image.error, addr),
            )
        ok, data = self._resolve_link_part(image.data_path, cell, image.src, addr)
        if not ok:
            return ("message", data)
        alt: str | None = None
        if image.alt_path is not None:
            ok, alt_value = self._resolve_link_part(
                image.alt_path, cell, image.src, addr
            )
            if not ok:
                return ("message", alt_value)
            alt = concat_text(cast("Value", alt_value))
        if isinstance(data, bytes) and data:
            self.pending_images.append(
                PendingImage(cell.col, row_num, alt, data, image.fit)
            )
        return ("skip", None)

    def _substitute_cell(self, row: Row, cell: Cell) -> bool:
        """Substitute one cell; return False when the cell disappears."""
        if cell.bindings and self._cell_is_starved(cell):
            return self._apply_blank(cell)
        parsed = self._parsed(cell)
        if parsed is None or parsed.is_static:
            return True
        addr = f"{cell.get('r')}"
        links = [s for s in parsed.constructs if isinstance(s, Link)]
        too_many_links = len(links) >= 2
        pieces: list[tuple[str, object]] = []
        for seg, segment in enumerate(parsed.segments):
            if isinstance(segment, Literal):
                if segment.raw:
                    pieces.append(("text", segment.raw))
            elif isinstance(segment, Marker):
                reason = cell.notes.get(seg)
                if reason is not None:
                    pieces.append(
                        (
                            "message",
                            self._report_cell("syntax", segment.src, reason, addr),
                        )
                    )
            elif isinstance(segment, Link):
                if too_many_links:
                    reason = "two or more links in one cell"
                    pieces.append(
                        (
                            "message",
                            self._report_cell("syntax", segment.src, reason, addr),
                        )
                    )
                else:
                    pieces.append(self._eval_link(segment, cell, addr))
            elif isinstance(segment, Image):
                pieces.append(self._eval_image(segment, cell, row.num, addr))
            else:
                pieces.append(self._eval_ref(segment, cell, addr))
        pieces = [p for p in pieces if p[0] != "skip"]
        if len(pieces) == 1 and pieces[0][0] == "value":
            return self._apply_value(cell, pieces[0][1])
        parts: list[str] = []
        for kind, payload in pieces:
            if kind in ("text", "message"):
                parts.append(str(payload))
            else:
                parts.append(concat_text(cast("Value", payload)))
        text = "".join(parts)
        if text == "":
            return self._apply_blank(cell)
        return self._apply_string(cell, text)

    def _apply_value(self, cell: Cell, value: object) -> bool:
        if value is None:
            return self._apply_blank(cell)
        if isinstance(value, bool):
            cell.set("t", "b")
            cell.inner = f"<v>{'1' if value else '0'}</v>"
            return True
        if isinstance(value, str):
            if value == "":
                return self._apply_blank(cell)
            return self._apply_string(cell, value)
        cell.unset("t")
        cell.inner = f"<v>{render_number(cast('Value', value))}</v>"
        return True

    def _apply_string(self, cell: Cell, text: str) -> bool:
        cell.set("t", "s")
        cell.inner = f"<v>{self.book.sst.index_for(text)}</v>"
        return True

    def _apply_blank(self, cell: Cell) -> bool:
        if cell.get("s") is None:
            return False
        cell.unset("t")
        cell.inner = ""
        return True

    # ----------------------------------------------------------- externals

    @property
    def axes_changed(self) -> bool:
        """Whether any expansion changed the sheet geometry."""
        return (self.vaxis is not None and self.vaxis.changed) or (
            self.haxis is not None and self.haxis.changed
        )

    def outer_vertical_major(self) -> bool:
        """Whether the vertical direction is the outer one for duplication."""
        for band in self.layout.bands:
            if band.vertical:
                continue
            if any(
                other.vertical and other.covers(band.start_at.row)
                for other in self.layout.bands
            ):
                return True
        for band in self.layout.bands:
            if not band.vertical:
                continue
            if any(
                not other.vertical and other.covers(band.start_at.col)
                for other in self.layout.bands
            ):
                return False
        return True

    def merge_copies(self, ref: str) -> list[str]:
        """All output copies of one merged range, in duplication order."""
        c1, r1, c2, r2 = parse_range(ref)
        vertical = instances_of(self.vaxis, r1, r2)
        horizontal = instances_of(self.haxis, col_to_index(c1), col_to_index(c2))
        pairs = (
            [(v, h) for v in vertical for h in horizontal]
            if self.outer_vertical_major()
            else [(v, h) for h in horizontal for v in vertical]
        )
        out = []
        for (nr1, nr2), (nc1, nc2) in pairs:
            first = f"{index_to_col(nc1)}{nr1}"
            second = f"{index_to_col(nc2)}{nr2}"
            out.append(first if first == second else f"{first}:{second}")
        return out

    def remap_ref(self, ref: str) -> str:
        """Remap an A1-style range through the sheet's axes."""
        return remap_range_ref(ref, root_mapper(self.vaxis), root_mapper(self.haxis))

    def _remap_sections(self) -> None:
        if not self.axes_changed:
            return
        post = self.doc.post
        post = re.sub(
            r'(<autoFilter ref=")([^"]*)(")',
            lambda m: f"{m.group(1)}{self.remap_ref(m.group(2))}{m.group(3)}",
            post,
        )
        post = re.sub(
            r'(sqref=")([^"]*)(")',
            lambda m: (
                f"{m.group(1)}"
                f"{remap_sqref(m.group(2), root_mapper(self.vaxis), root_mapper(self.haxis))}"
                f"{m.group(3)}"
            ),
            post,
        )
        post = re.sub(
            r'(<hyperlink ref=")([^"]*)(")',
            lambda m: f"{m.group(1)}{self.remap_ref(m.group(2))}{m.group(3)}",
            post,
        )
        merge_block = re.search(
            r'<mergeCells count="\d+">.*?</mergeCells>', post, re.DOTALL
        )
        if merge_block is not None:
            refs = re.findall(
                r'<mergeCell ref="([^"]*)"></mergeCell>', merge_block.group(0)
            )
            copies = [copy for ref in refs for copy in self.merge_copies(ref)]
            if copies:
                body = "".join(f'<mergeCell ref="{ref}"></mergeCell>' for ref in copies)
                replacement = f'<mergeCells count="{len(copies)}">{body}</mergeCells>'
            else:
                replacement = ""
            post = post.replace(merge_block.group(0), replacement, 1)
        self.doc.post = post

    def column_width(self, col: int) -> float:
        """Width (in characters) of 0-based column ``col``."""
        for attrs in self.doc.cols():
            mapping = dict(attrs)
            if int(mapping["min"]) - 1 <= col <= int(mapping["max"]) - 1:
                return float(mapping["width"])
        return 8.43

    def row_height(self, row_num: int) -> float:
        """Height (in points) of 1-based row ``row_num``."""
        for row in self.doc.rows:
            if row.num == row_num:
                mapping = dict(row.attrs)
                if "ht" in mapping:
                    return float(mapping["ht"])
                break
        m = re.search(r'defaultRowHeight="([0-9.]+)"', self.doc.pre)
        return float(m.group(1)) if m else 15.0

    def cell_text(self, row_num: int, col: int) -> str:
        """Display text of the output cell at 1-based ``row_num`` / 0-based ``col``."""
        for row in self.doc.rows:
            if row.num != row_num:
                continue
            for cell in row.cells:
                if cell.col != col:
                    continue
                index = cell.sst_index
                if index is not None:
                    return self.book.sst.text(index)
                value = cell.value_text
                return unescape(value) if value is not None else ""
        return ""

    def _substitute_sections(self) -> None:
        report = self.book.sheet_reporter(self.sheet_name)
        declared = self.layout.names
        for source in ("pre", "post"):
            raw = getattr(self.doc, source)
            raw = substitute_text_nodes(
                raw,
                _HEADER_FOOTER_TAGS,
                self.book.data,
                self.base_bindings,
                escape_text,
                report,
                declared,
            )
            raw = substitute_attrs(
                raw,
                "dataValidation",
                _VALIDATION_ATTRS,
                self.book.data,
                self.base_bindings,
                escape_attr,
                report,
                declared,
            )
            raw = substitute_attrs(
                raw,
                "hyperlink",
                ("tooltip",),
                self.book.data,
                self.base_bindings,
                escape_attr,
                report,
                declared,
            )
            setattr(self.doc, source, raw)

    # ---------------------------------------------------------------- main

    def process(self) -> None:
        """Run marker analysis, expansion, and substitution on the sheet."""
        markers = self._collect_markers()
        self.layout = build_layout(markers)
        self._check_merge_straddle()
        used = self._collect_band_refs(self.layout)
        mark_unused(self.layout, used)
        self._collect_band_refs(self.layout)
        self._annotate_markers(self.layout)

        vertical_roots = self.layout.roots_by_direction[True]
        horizontal_roots = self.layout.roots_by_direction[False]
        if vertical_roots and self.doc.rows:
            hi = max(row.num for row in self.doc.rows)
            self.vaxis = self._make_axis(vertical_roots, 1, hi)
            self._apply_vertical(self.vaxis)
        if horizontal_roots and self.doc.rows:
            hi = max(
                (cell.col for row in self.doc.rows for cell in row.cells),
                default=-1,
            )
            self.haxis = self._make_axis(horizontal_roots, 0, hi)
            self._apply_horizontal(self.haxis)
            self._expand_cols(self.haxis)

        self._rebase_formulas()
        self._remap_sections()
        self._substitute_cells()
        self._substitute_sections()

    def emit(self) -> bytes:
        """Serialize the processed worksheet."""
        return self.doc.emit()

    def string_reference_count(self) -> int:
        """Count ``t="s"`` cells for the shared-string ``count`` attribute."""
        return sum(
            1 for row in self.doc.rows for cell in row.cells if cell.get("t") == "s"
        )


def _paths_of(segment: object) -> list[tuple[PathStep, ...]]:
    if isinstance(segment, ValueRef):
        return [segment.path]
    if isinstance(segment, Link):
        return [p for p in (segment.label_path, segment.url_path) if p]
    if isinstance(segment, Image):
        return [p for p in (segment.alt_path, segment.data_path) if p]
    return []
