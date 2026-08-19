"""Per-sheet substitution: markers, bands, and cell values.

Everything a sheet needs done is said in Excel's own terms — copy these
rows, delete those, put this in that cell. Which rows to copy is what a
band declaration works out to, and that is the whole of what lives here.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import TYPE_CHECKING, cast

from _excel import Cell
from _excel import Image as Picture
from xlsxfill._bands import Band, BandLayout, MarkerAt, build_layout, mark_unused
from xlsxfill._containers import substitute_container
from xlsxfill._resolve import (
    ResolveError,
    check_assertion,
    check_scalar,
    collection_length,
    concat_text,
    render_number,
    resolve,
)
from xlsxfill._syntax import (
    Image,
    IndexStep,
    Link,
    Literal,
    Marker,
    ParsedText,
    ValueRef,
    tokenize,
)

if TYPE_CHECKING:
    from collections.abc import Mapping

    from _excel import CellValue, Found, Sheet
    from xlsxfill._book import Book
    from xlsxfill._syntax import PathStep
    from xlsxfill._values import Value


@dataclass(frozen=True)
class _Line:
    """Where one output line came from and what it is an iteration of."""

    source: int
    bindings: dict[str, int]


class SheetEngine:
    """Processes one worksheet."""

    def __init__(self, book: Book, sheet: Sheet, bindings: dict[str, int]) -> None:
        """Prepare the engine for ``sheet`` under sheet-level ``bindings``."""
        self.book = book
        self.sheet = sheet
        self.base_bindings = bindings
        self.name = sheet.name
        self.layout = BandLayout(
            roots_by_direction={True: [], False: []},
            marker_reasons={},
            bands=[],
        )
        self.rows: dict[int, _Line] = {}
        self.columns: dict[int, _Line] = {}
        self._texts: dict[tuple[int, int], str] = {}
        self._parsed: dict[str, ParsedText] = {}
        self._band_refs: dict[int, list[tuple[tuple[PathStep, ...], int, int]]] = {}
        self._counts: dict[tuple[int, tuple[tuple[str, int], ...]], int | None] = {}
        self._links: list[tuple[Cell, str]] = []
        self._pictures: list[tuple[Cell, Picture, str, str | None]] = []

    # ---------------------------------------------------------------- text

    def parse(self, text: str) -> ParsedText:
        """Tokenize ``text``, remembering the result."""
        found = self._parsed.get(text)
        if found is None:
            found = tokenize(text)
            self._parsed[text] = found
        return found

    def load(self, hits: list[Found]) -> None:
        """Take the sheet's cell texts as the template to work from."""
        self._texts = {
            (hit.where.row, hit.where.column): hit.text
            for hit in hits
            if isinstance(hit.where, Cell)
        }

    # -------------------------------------------------------- band analysis

    def analyse(self) -> None:
        """Work out the sheet's bands and which of them are used."""
        markers = [
            MarkerAt(segment, row, column, seg)
            for (row, column), text in self._texts.items()
            for seg, segment in enumerate(self.parse(text).segments)
            if isinstance(segment, Marker)
        ]
        self.layout = build_layout(markers)
        self._drop_straddled()
        used = self._collect_refs()
        mark_unused(self.layout, used)
        self._collect_refs()

    def _drop_straddled(self) -> None:
        """Invalidate a band a merged block reaches across.

        A band that repeats has to be able to stand alone. A merged block
        that starts inside one and ends outside cannot be copied without
        being torn, so the declaration is refused rather than honoured
        half way.
        """
        dropped: set[int] = set()
        for top, left, bottom, right in self.sheet.merged_ranges:
            for band in self.layout.bands:
                lo, hi = (top, bottom) if band.vertical else (left, right)
                last = band.stop - 1
                touches = not (hi < band.start or lo > last)
                inside = band.start <= lo and hi <= last
                if touches and not inside:
                    dropped.add(id(band))
        self.layout.drop(dropped, "merged cell straddles the band boundary")

    def _collect_refs(self) -> dict[int, bool]:
        self._band_refs = {id(band): [] for band in self.layout.bands}
        used: dict[int, bool] = {}
        seen: set[str] = set()
        for (row, column), text in self._texts.items():
            for segment in self.parse(text).constructs:
                for path in _paths_of(segment):
                    for step in path:
                        if (
                            not isinstance(step, IndexStep)
                            or step.is_fixed
                            or step.is_sheet
                        ):
                            continue
                        band = self.layout.band_for(step.symbol, row, column, seen=seen)
                        if band is None:
                            continue
                        used[id(band)] = True
                        self._band_refs[id(band)].append((path, row, column))
        return used

    # ------------------------------------------------------------ counting

    def count(self, band: Band, bindings: Mapping[str, int]) -> int | None:
        """How many times ``band`` repeats under ``bindings``.

        ``None`` means it never comes round at all, because a band
        crossing it has no blocks; it is then left as it stands.
        """
        key = (id(band), tuple(sorted(bindings.items())))
        if key in self._counts:
            return self._counts[key]
        refs = self._band_refs.get(id(band), [])
        result: int | None = None
        for context in self._contexts(band, refs, dict(bindings)):
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
        self._counts[key] = result
        return result

    def _contexts(
        self,
        band: Band,
        refs: list[tuple[tuple[PathStep, ...], int, int]],
        bindings: dict[str, int],
    ) -> list[dict[str, int]]:
        needed: list[Band] = []
        seen: set[str] = set()
        for path, row, column in refs:
            for step in path:
                if isinstance(step, IndexStep) and step.symbol == band.name:
                    break
                if (
                    isinstance(step, IndexStep)
                    and not step.is_fixed
                    and not step.is_sheet
                    and step.symbol not in bindings
                ):
                    other = self.layout.band_for(step.symbol, row, column, seen=seen)
                    if other is not None and other not in needed:
                        needed.append(other)
        contexts = [bindings]
        for other in needed:
            widened: list[dict[str, int]] = []
            for context in contexts:
                count = self.count(other, context)
                if not count:
                    widened.append(context)
                    continue
                widened.extend({**context, other.name: i} for i in range(count))
            contexts = widened
        return contexts

    # ----------------------------------------------------------- expansion

    def expand(self) -> None:
        """Copy and delete rows and columns until the bands have played out."""
        rows = [row for row, _ in self._texts]
        columns = [column for _, column in self._texts]
        if not rows:
            return
        if _vertical_first(self.layout):
            self._expand_rows(max(rows))
            self._expand_columns(max(columns))
        else:
            self._expand_columns(max(columns))
            self._expand_rows(max(rows))

    def _expand_rows(self, last: int) -> None:
        roots = self.layout.roots_by_direction[True]
        self.rows = {}
        self._walk(
            1, last, roots, dict(self.base_bindings), 0, self.rows, vertical=True
        )

    def _expand_columns(self, last: int) -> None:
        roots = self.layout.roots_by_direction[False]
        self.columns = {}
        self._walk(
            1, last, roots, dict(self.base_bindings), 0, self.columns, vertical=False
        )

    def _walk(
        self,
        lo: int,
        hi: int,
        bands: list[Band],
        bindings: dict[str, int],
        offset: int,
        lines: dict[int, _Line],
        *,
        vertical: bool,
    ) -> int:
        """Play out ``bands`` over template lines ``lo``..``hi``.

        ``offset`` is how far the region has already moved. Returns how
        many lines this level added, negative when it removed some.
        """
        added = 0
        position = lo
        for band in sorted(bands, key=lambda band: band.start):
            for line in range(position, band.start):
                lines[line + offset + added] = _Line(line, dict(bindings))
            count = self.count(band, bindings)
            base = band.start + offset + added
            if count is None:
                for line in range(band.start, band.marker_stop):
                    lines[line + offset + added] = _Line(line, dict(bindings))
            elif count == 0:
                self._delete(base, base + band.period, vertical=vertical)
                added -= band.period + 1
            else:
                added += self._repeat(
                    band, base, count, bindings, lines, vertical=vertical
                )
            position = band.marker_stop
        for line in range(position, hi + 1):
            lines[line + offset + added] = _Line(line, dict(bindings))
        return added

    def _repeat(
        self,
        band: Band,
        base: int,
        count: int,
        bindings: dict[str, int],
        lines: dict[int, _Line],
        *,
        vertical: bool,
    ) -> int:
        """Leave ``count`` copies of one band's block where one stood.

        The copies are made before the declaration line goes, so that a
        reference to that line still says where it was when the copies
        were laid out, and closes onto the last of them rather than the
        first.
        """
        added = 0
        if count > 1:
            self._copy(base, base + band.period - 1, count - 1, vertical=vertical)
            added += (count - 1) * band.period
        marker = base + band.period + added
        self._delete(marker, marker, vertical=vertical)
        added -= 1
        grown = 0
        for index in range(count):
            top = base + index * band.period + grown
            grown += self._walk(
                band.start,
                band.stop - 1,
                band.children,
                {**bindings, band.name: index},
                top - band.start,
                lines,
                vertical=vertical,
            )
        return added + grown

    def _delete(self, first: int, last: int, *, vertical: bool) -> None:
        if vertical:
            self.sheet.delete_rows(first, last)
        else:
            self.sheet.delete_columns(first, last)

    def _copy(self, first: int, last: int, copies: int, *, vertical: bool) -> None:
        if vertical:
            self.sheet.duplicate_rows(first, last, copies=copies)
        else:
            self.sheet.duplicate_columns(first, last, copies=copies)

    # -------------------------------------------------------- substitution

    def bindings_at(self, row: int, column: int) -> dict[str, int]:
        """The iteration indices in force at one output cell."""
        bindings = dict(self.base_bindings)
        line = self.rows.get(row)
        if line is not None:
            bindings.update(line.bindings)
        line = self.columns.get(column)
        if line is not None:
            bindings.update(line.bindings)
        return bindings

    def origin(self, row: int, column: int) -> tuple[int, int]:
        """The template cell one output cell was copied from."""
        vertical = self.rows.get(row)
        horizontal = self.columns.get(column)
        return (
            vertical.source if vertical is not None else row,
            horizontal.source if horizontal is not None else column,
        )

    def substitute(self, hits: list[Found]) -> None:
        """Resolve every placeholder the sheet is left holding."""
        for hit in hits:
            where = hit.where
            if isinstance(where, Cell):
                self._cell(where, hit.text)
            elif where.part != "name":
                self._container(hit)
        self._attach()

    def _attach(self) -> None:
        """Hang the links and pictures the cells asked for.

        Left until the cells are settled, so a picture is sized against
        the cell as it finally reads and the parts land in the workbook
        in the order they were asked for.
        """
        for where, url in self._links:
            self.book.excel.insert_link(where, url)
        for where, picture, fit, alt in self._pictures:
            self._place(where, picture, fit, alt)

    def _container(self, hit: Found) -> None:
        where = hit.where
        report = self.book.sheet_reporter(self.name, where)
        replaced = substitute_container(
            hit.text,
            self.book.data,
            self.bindings_at(getattr(where, "row", 0), getattr(where, "column", 0)),
            report,
            self.layout.names,
        )
        if replaced is not None and replaced != hit.text:
            self.book.excel.set(where, replaced)

    def _cell(self, where: Cell, text: str) -> None:
        parsed = self.parse(text)
        if parsed.is_static:
            return
        own = self.bindings_at(where.row, where.column)
        if self._starved(own):
            self.book.excel.set(where, None)
            return
        pieces = self._pieces(where, parsed, own)
        if len(pieces) == 1 and pieces[0][0] == "value":
            self.book.excel.set(where, cast("CellValue", pieces[0][1]))
            return
        rendered = "".join(
            str(payload) if kind in ("text", "message") else concat_text(payload)
            for kind, payload in pieces
        )
        self.book.excel.set(where, rendered or None)

    def _pieces(
        self, where: Cell, parsed: ParsedText, bindings: dict[str, int]
    ) -> list[tuple[str, Value]]:
        address = _address(where.row, where.column)
        source = self.origin(where.row, where.column)
        links = [s for s in parsed.constructs if isinstance(s, Link)]
        crowded = len(links) >= 2
        pieces: list[tuple[str, Value]] = []
        for seg, segment in enumerate(parsed.segments):
            if isinstance(segment, Literal):
                if segment.raw:
                    pieces.append(("text", segment.raw))
            elif isinstance(segment, Marker):
                reason = self.layout.marker_reasons.get((*source, seg))
                if reason is not None:
                    pieces.append(
                        (
                            "message",
                            self._report("syntax", segment.src, reason, address),
                        )
                    )
            elif isinstance(segment, Link):
                if crowded:
                    pieces.append(
                        (
                            "message",
                            self._report(
                                "syntax",
                                segment.src,
                                "two or more links in one cell",
                                address,
                            ),
                        )
                    )
                else:
                    pieces.append(self._link(segment, where, bindings, address))
            elif isinstance(segment, Image):
                pieces.append(self._image(segment, where, bindings, address))
            else:
                pieces.append(self._value(segment, bindings, address))
        return [piece for piece in pieces if piece[0] != "skip"]

    # ----------------------------------------------------------- resolving

    def _report(self, kind: str, construct: str, reason: str, address: str) -> str:
        return self.book.report_cell(kind, construct, reason, self.name, address)

    def _bound(
        self, path: tuple[PathStep, ...], bindings: Mapping[str, int]
    ) -> str | None:
        for step in path:
            if (
                isinstance(step, IndexStep)
                and not step.is_fixed
                and not step.is_sheet
                and step.symbol not in bindings
            ):
                if step.symbol in self.layout.names:
                    return f'band "{step.symbol}" used outside its band'
                return f'band "{step.symbol}" is not declared'
        return None

    def _value(
        self, ref: ValueRef, bindings: Mapping[str, int], address: str
    ) -> tuple[str, Value]:
        if ref.error is not None:
            return ("message", self._report("syntax", ref.src, ref.error, address))
        unbound = self._bound(ref.path, bindings)
        if unbound is not None:
            return ("message", self._report("syntax", ref.src, unbound, address))
        try:
            value = resolve(ref.path, self.book.data, bindings)
            if value is not None:
                if ref.assert_type is not None:
                    check_assertion(value, ref.assert_type)
                check_scalar(value)
                if isinstance(value, datetime) and value.tzinfo is not None:
                    render_number(value)
        except ResolveError as error:
            return ("message", self._report("data", ref.src, error.reason, address))
        return ("value", value)

    def _part(
        self,
        path: tuple[PathStep, ...],
        bindings: Mapping[str, int],
        src: str,
        address: str,
    ) -> tuple[bool, Value]:
        unbound = self._bound(path, bindings)
        if unbound is not None:
            return (False, self._report("syntax", src, unbound, address))
        try:
            return (True, resolve(path, self.book.data, bindings))
        except ResolveError as error:
            return (False, self._report("data", src, error.reason, address))

    def _link(
        self, link: Link, where: Cell, bindings: Mapping[str, int], address: str
    ) -> tuple[str, Value]:
        if link.error is not None:
            return ("message", self._report("syntax", link.src, link.error, address))
        ok, url = self._part(link.url_path, bindings, link.src, address)
        if not ok:
            return ("message", url)
        label: tuple[str, Value] = ("skip", None)
        if link.label_path is not None:
            ok, text = self._part(link.label_path, bindings, link.src, address)
            if not ok:
                return ("message", text)
            label = ("text", concat_text(text))
        if isinstance(url, str) and url != "":
            self._links.append((where, url))
        return label

    def _image(
        self, image: Image, where: Cell, bindings: Mapping[str, int], address: str
    ) -> tuple[str, Value]:
        if image.error is not None:
            return ("message", self._report("syntax", image.src, image.error, address))
        ok, data = self._part(image.data_path, bindings, image.src, address)
        if not ok:
            return ("message", data)
        alt: str | None = None
        if image.alt_path is not None:
            ok, text = self._part(image.alt_path, bindings, image.src, address)
            if not ok:
                return ("message", text)
            alt = concat_text(text)
        if isinstance(data, bytes) and data:
            try:
                self._pictures.append((where, Picture(data), image.fit, alt))
            except ValueError as error:
                return ("message", self._report("data", image.src, str(error), address))
        return ("skip", None)

    def _place(self, where: Cell, picture: Picture, fit: str, alt: str | None) -> None:
        frame_width, frame_height = self.book.excel.cell_size(where)
        width, height, offset_x, offset_y, crop = _fitted(
            picture, fit, frame_width, frame_height
        )
        self.book.excel.insert_image(
            where,
            picture,
            width=width,
            height=height,
            offset_x=offset_x,
            offset_y=offset_y,
            crop=crop,
            alt=alt,
        )

    def _starved(self, bindings: Mapping[str, int]) -> bool:
        for name, index in bindings.items():
            candidates = [band for band in self.layout.bands if band.name == name]
            if len(candidates) != 1:
                continue
            context = {k: v for k, v in bindings.items() if k != name}
            count = self.count(candidates[0], context)
            if count is not None and index >= count:
                return True
        return False


def _fitted(
    picture: Picture, fit: str, frame_width: int, frame_height: int
) -> tuple[float, float, float, float, tuple[float, float, float, float] | None]:
    """How big to draw a picture in a cell, and what of it to show."""
    if fit == "fill":
        return (frame_width, frame_height, 0, 0, None)
    if fit == "cover":
        scale = max(frame_width / picture.width, frame_height / picture.height)
        keep_width = frame_width / scale
        keep_height = frame_height / scale
        left = (picture.width - keep_width) / 2
        top = (picture.height - keep_height) / 2
        return (
            frame_width,
            frame_height,
            0,
            0,
            (left, top, left + keep_width, top + keep_height),
        )
    scale = min(frame_width / picture.width, frame_height / picture.height)
    width = round(picture.width * scale)
    height = round(picture.height * scale)
    return (
        width,
        height,
        (frame_width - width) // 2,
        (frame_height - height) // 2,
        None,
    )


def _vertical_first(layout: BandLayout) -> bool:
    """Whether rows are played out before columns.

    The band that encloses the other goes first, so its copies already
    carry the inner declaration when the inner one comes round.
    """
    for band in layout.bands:
        if band.vertical:
            continue
        if any(
            other.vertical and other.covers(band.start_at.row) for other in layout.bands
        ):
            return True
    for band in layout.bands:
        if not band.vertical:
            continue
        if any(
            not other.vertical and other.covers(band.start_at.col)
            for other in layout.bands
        ):
            return False
    return True


def _address(row: int, column: int) -> str:
    letters = ""
    index = column
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return f"{letters}{row}"


def _paths_of(segment: object) -> list[tuple[PathStep, ...]]:
    if isinstance(segment, ValueRef):
        return [segment.path]
    if isinstance(segment, Link):
        return [path for path in (segment.label_path, segment.url_path) if path]
    if isinstance(segment, Image):
        return [path for path in (segment.alt_path, segment.data_path) if path]
    return []
