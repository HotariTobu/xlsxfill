from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlsxfill._syntax import Marker


@dataclass(frozen=True)
class MarkerAt:
    marker: Marker
    row: int
    col: int
    seg: int

    def pos(self, *, vertical: bool) -> int:
        return self.row if vertical else self.col

    def cross(self, *, vertical: bool) -> int:
        return self.col if vertical else self.row


@dataclass
class Band:
    name: str
    vertical: bool
    start: int
    period: int
    start_at: MarkerAt
    plus_at: MarkerAt
    children: list[Band] = field(default_factory=list)

    @property
    def stop(self) -> int:
        return self.start + self.period

    @property
    def marker_stop(self) -> int:
        return self.stop + 1

    def contains(self, other: Band) -> bool:
        return self.start <= other.start and other.marker_stop <= self.marker_stop

    def covers(self, coordinate: int) -> bool:
        return self.start <= coordinate < self.stop

    def reset_children(self) -> Band:
        self.children = []
        return self


@dataclass
class BandLayout:
    roots_by_direction: dict[bool, list[Band]]
    marker_reasons: dict[tuple[int, int, int], str]
    bands: list[Band]

    def drop(self, band_ids: set[int], reason: str) -> None:
        if not band_ids:
            return
        for band in self.bands:
            if id(band) in band_ids:
                self.marker_reasons[_key(band.start_at)] = reason
                self.marker_reasons[_key(band.plus_at)] = reason
        self.bands = [b for b in self.bands if id(b) not in band_ids]
        for direction in (True, False):
            self.roots_by_direction[direction] = _nest(
                [b.reset_children() for b in self.bands if b.vertical is direction],
            )

    @property
    def names(self) -> frozenset[str]:
        return frozenset(band.name for band in self.bands)

    def band_for(self, name: str, row: int, col: int, *, seen: set[str]) -> Band | None:
        found: Band | None = None
        for band in self.bands:
            if band.name != name:
                continue
            seen.add(name)
            if band.covers(row if band.vertical else col) and (
                found is None or band.period <= found.period
            ):
                found = band
        return found


def _pair_markers(
    markers: list[MarkerAt],
    *,
    vertical: bool,
    reasons: dict[tuple[int, int, int], str],
) -> list[Band]:
    starts = sorted(
        (m for m in markers if not m.marker.plus),
        key=lambda m: m.pos(vertical=vertical),
    )
    pluses = sorted(
        (m for m in markers if m.marker.plus),
        key=lambda m: m.pos(vertical=vertical),
    )
    bands: list[Band] = []
    used_pluses: set[MarkerAt] = set()
    for i, start in enumerate(starts):
        next_start = (
            starts[i + 1].pos(vertical=vertical) if i + 1 < len(starts) else None
        )
        candidate = next(
            (
                p
                for p in pluses
                if p not in used_pluses
                and p.pos(vertical=vertical) >= start.pos(vertical=vertical)
                and (next_start is None or p.pos(vertical=vertical) <= next_start)
            ),
            None,
        )
        if candidate is None:
            reasons[_key(start)] = f"missing #{{{start.marker.name}+1}}"
            continue
        used_pluses.add(candidate)
        band = _validate_pair(start, candidate, vertical=vertical, reasons=reasons)
        if band is not None:
            bands.append(band)
    for plus in pluses:
        if plus not in used_pluses:
            reasons[_key(plus)] = f"missing #{{{plus.marker.name}}}"
    return bands


def _key(marker_at: MarkerAt) -> tuple[int, int, int]:
    return (marker_at.row, marker_at.col, marker_at.seg)


def _validate_pair(
    start: MarkerAt,
    plus: MarkerAt,
    *,
    vertical: bool,
    reasons: dict[tuple[int, int, int], str],
) -> Band | None:
    if start.cross(vertical=vertical) != plus.cross(vertical=vertical):
        line = "column" if vertical else "row"
        reason = f'start and "+1" markers are not in the same {line}'
        reasons[_key(start)] = reason
        reasons[_key(plus)] = reason
        return None
    period = plus.pos(vertical=vertical) - start.pos(vertical=vertical)
    if period == 0:
        reason = "band height is 0" if vertical else "band width is 0"
        reasons[_key(start)] = reason
        reasons[_key(plus)] = reason
        return None
    return Band(
        name=start.marker.name,
        vertical=vertical,
        start=start.pos(vertical=vertical),
        period=period,
        start_at=start,
        plus_at=plus,
    )


def _check_overlaps(
    bands: list[Band],
    reasons: dict[tuple[int, int, int], str],
) -> list[Band]:
    bad: set[int] = set()
    for i, a in enumerate(bands):
        for j in range(i + 1, len(bands)):
            b = bands[j]
            if a.vertical != b.vertical:
                continue
            if a.marker_stop <= b.start or b.marker_stop <= a.start:
                continue
            if a.contains(b) or b.contains(a):
                continue
            bad.update((i, j))
    for index in sorted(bad):
        band = bands[index]
        for marker_at in (band.start_at, band.plus_at):
            reasons[_key(marker_at)] = "bands overlap"
    return [band for i, band in enumerate(bands) if i not in bad]


def _nest(bands: list[Band]) -> list[Band]:
    ordered = sorted(bands, key=lambda b: (b.start, -b.period))
    roots: list[Band] = []
    stack: list[Band] = []
    for band in ordered:
        while stack and not stack[-1].contains(band):
            stack.pop()
        if stack:
            stack[-1].children.append(band)
        else:
            roots.append(band)
        stack.append(band)
    return roots


def build_layout(marker_ats: list[MarkerAt]) -> BandLayout:
    reasons: dict[tuple[int, int, int], str] = {}
    valid = [m for m in marker_ats if m.marker.error is None]
    for marker_at in marker_ats:
        if marker_at.marker.error is not None:
            reasons[_key(marker_at)] = marker_at.marker.error

    bands: list[Band] = []
    by_name_dir: dict[tuple[str, bool], list[MarkerAt]] = {}
    for marker_at in valid:
        direction = marker_at.marker.name[0] == "r"
        by_name_dir.setdefault((marker_at.marker.name, direction), []).append(marker_at)
    for (_, direction), group in sorted(by_name_dir.items()):
        bands.extend(_pair_markers(group, vertical=direction, reasons=reasons))

    bands = _check_overlaps(bands, reasons)

    roots = {
        True: _nest([b for b in bands if b.vertical]),
        False: _nest([b for b in bands if not b.vertical]),
    }
    return BandLayout(roots_by_direction=roots, marker_reasons=reasons, bands=bands)


def mark_unused(
    layout: BandLayout,
    is_used: dict[int, bool],
) -> None:
    dropped = {id(band) for band in layout.bands if not is_used.get(id(band), False)}
    layout.drop(dropped, "band is never used")
