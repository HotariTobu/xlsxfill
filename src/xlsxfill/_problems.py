"""Errors reported in the output while processing continues."""

from dataclasses import dataclass
from typing import Literal


@dataclass(frozen=True)
class BaseProblem:
    """An error reported in the output while processing continues."""

    kind: Literal["syntax", "data"]
    construct: str
    reason: str

    @property
    def message(self) -> str:
        """The string embedded in the output.

        ``#SYNTAX! <construct>: <reason>`` or ``#DATA! <construct>: <reason>``.
        """
        return f"#{self.kind.upper()}! {self.construct}: {self.reason}"


@dataclass(frozen=True)
class BookProblem(BaseProblem):
    """A problem in a workbook-level string container."""

    part: Literal["doc_props"]


@dataclass(frozen=True)
class SheetProblem(BaseProblem):
    """A problem in a sheet-level string container."""

    sheet: str
    part: Literal["sheet_name", "header_footer", "shape", "chart"]


@dataclass(frozen=True)
class CellProblem(BaseProblem):
    """A problem in a cell or in a string container attached to a cell."""

    sheet: str
    cell: str
    part: Literal["cell", "comment", "tooltip", "validation"]


type Problem = BookProblem | SheetProblem | CellProblem
"""A problem reported in the output."""
