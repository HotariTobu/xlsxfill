"""Workbook-level orchestration of the fill run.

The whole run is said in Excel's own terms: search the workbook, copy or
delete sheets, then put values where the search found placeholders. What
a placeholder means, and how many times a band comes round, is all that
belongs to this package.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from _excel import BookText, Cell, CellText, SheetText
from xlsxfill._containers import substitute_container
from xlsxfill._engine import SheetEngine
from xlsxfill._problems import BookProblem, CellProblem, Problem, SheetProblem
from xlsxfill._resolve import collection_length
from xlsxfill._syntax import Image, IndexStep, Link, ValueRef, tokenize

if TYPE_CHECKING:
    from collections.abc import Callable, Mapping
    from typing import Literal

    from _excel import Book as Excel
    from _excel import Found, Sheet, Where
    from xlsxfill._syntax import PathStep
    from xlsxfill._values import Value

_SHEET_PARTS: dict[str, str] = {
    "name": "sheet_name",
    "header": "header_footer",
    "footer": "header_footer",
    "shape": "shape",
    "chart": "chart",
}


class Book:
    """One fill run over an open workbook."""

    def __init__(self, excel: Excel, data: Mapping[str, Value]) -> None:
        """Prepare the run against ``excel`` with input ``data``."""
        self.excel = excel
        self.data = data
        self.problems: list[Problem] = []

    # ------------------------------------------------------------ problems

    def report_cell(
        self, kind: str, construct: str, reason: str, sheet: str, cell: str
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
        self, sheet: str, where: Where
    ) -> Callable[[str, str, str], str]:
        """Build a problem sink for one string container."""

        def report(kind: str, construct: str, reason: str) -> str:
            problem = _problem_for(where, sheet, kind, construct, reason)
            self.problems.append(problem)
            return problem.message

        return report

    def book_reporter(self) -> Callable[[str, str, str], str]:
        """Build a problem sink for the document properties."""

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

    # --------------------------------------------------------------- phases

    def run(self) -> list[Problem]:
        """Execute the whole fill and return the collected problems."""
        engines = []
        for sheet, bindings in self._plan_sheets():
            engine = SheetEngine(self, sheet, bindings)
            engine.load(_on(self.excel.find(), sheet))
            engine.analyse()
            engine.expand()
            engines.append(engine)
        for engine in engines:
            engine.substitute(_on(self.excel.find(), engine.sheet))
        self._substitute_properties()
        return self.problems

    def _substitute_properties(self) -> None:
        report = self.book_reporter()
        for hit in self.excel.find():
            if not isinstance(hit.where, BookText):
                continue
            replaced = substitute_container(hit.text, self.data, {}, report)
            if replaced is not None and replaced != hit.text:
                self.excel.set(hit.where, replaced)

    # ---------------------------------------------------- sheet repetition

    def _plan_sheets(self) -> list[tuple[Sheet, dict[str, int]]]:
        plan: list[tuple[Sheet, dict[str, int]]] = []
        for sheet in _sheets_of(self.excel.find()):
            plan.extend(self._plan_sheet(sheet))
        return plan

    def _plan_sheet(self, sheet: Sheet) -> list[tuple[Sheet, dict[str, int]]]:
        hits = _on(self.excel.find(), sheet)
        name_hit = _name_of(hits)
        name = sheet.name
        paths = [path for hit in hits for path in _paths_in(hit.text)]
        if not any(
            isinstance(step, IndexStep) and step.is_sheet
            for path in paths
            for step in path
        ):
            self._rename(name_hit, name, {})
            return [(sheet, {})]

        count = self._sheet_count(paths)
        if count == 0:
            self.excel.delete_sheet(sheet)
            return []
        named = any(
            isinstance(step, IndexStep) and step.is_sheet
            for path in _paths_in(name)
            for step in path
        )
        titles = [self._title(name, index, named=named) for index in range(count)]
        self._rename(name_hit, name, {"s": 0}, titles[0])
        copies = self.excel.duplicate_sheet(sheet, titles[1:])
        return [(sheet, {"s": 0})] + [
            (copy, {"s": index + 1}) for index, copy in enumerate(copies)
        ]

    def _title(self, name: str, index: int, *, named: bool) -> str:
        report = self.sheet_reporter(name, _NO_WHERE)
        base = substitute_container(name, self.data, {"s": index}, report) or name
        return base if named else f"{base} ({index + 1})"

    def _rename(
        self,
        name_hit: Found | None,
        name: str,
        bindings: dict[str, int],
        title: str | None = None,
    ) -> None:
        if name_hit is None:
            return
        if title is None:
            report = self.sheet_reporter(name, name_hit.where)
            title = substitute_container(name, self.data, bindings, report)
        if title is not None and title != name:
            self.excel.set(name_hit.where, title)

    def _sheet_count(self, paths: list[tuple[PathStep, ...]]) -> int:
        lengths = [
            length
            for path in paths
            if (length := collection_length(path, "s", self.data, {})) is not None
        ]
        return min(lengths) if lengths else 0


def _problem_for(
    where: Where, sheet: str, kind: str, construct: str, reason: str
) -> Problem:
    known = cast("Literal['syntax', 'data']", kind)
    if isinstance(where, Cell | CellText):
        return CellProblem(
            kind=known,
            construct=construct,
            reason=reason,
            sheet=sheet,
            cell="",
            part=where.part,
        )
    part = _SHEET_PARTS.get(getattr(where, "part", "shape"), "shape")
    return SheetProblem(
        kind=known,
        construct=construct,
        reason=reason,
        sheet=sheet,
        part=cast("Literal['sheet_name', 'header_footer', 'shape', 'chart']", part),
    )


class _NoWhere:
    """Stands in for the tab name of a sheet not yet renamed."""

    @property
    def part(self) -> str:
        """Always the tab name."""
        return "name"


_NO_WHERE = cast("Where", _NoWhere())


def _sheets_of(hits: list[Found]) -> list[Sheet]:
    sheets: list[Sheet] = []
    for hit in hits:
        sheet = getattr(hit.where, "sheet", None)
        if sheet is not None and sheet not in sheets:
            sheets.append(sheet)
    return sheets


def _on(hits: list[Found], sheet: Sheet) -> list[Found]:
    return [hit for hit in hits if getattr(hit.where, "sheet", None) is sheet]


def _name_of(hits: list[Found]) -> Found | None:
    for hit in hits:
        if isinstance(hit.where, SheetText) and hit.where.part == "name":
            return hit
    return None


def _paths_in(text: str) -> list[tuple[PathStep, ...]]:
    paths: list[tuple[PathStep, ...]] = []
    for segment in tokenize(text).constructs:
        if isinstance(segment, Link):
            candidates = [segment.label_path, segment.url_path]
        elif isinstance(segment, Image):
            candidates = [segment.alt_path, segment.data_path]
        elif isinstance(segment, ValueRef):
            candidates = [segment.path]
        else:
            candidates = []
        paths.extend(path for path in candidates if path)
    return paths
