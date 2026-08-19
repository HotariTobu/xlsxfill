"""One worksheet."""

from __future__ import annotations

from typing import TYPE_CHECKING

from _patched_xlsxedit import (
    delete_columns,
    delete_rows,
    duplicate_columns,
    duplicate_rows,
)

if TYPE_CHECKING:
    from xlsxedit.parts import WorksheetPart
    from xlsxedit.workbook import Workbook
    from xlsxedit.worksheet import Worksheet


class Sheet:
    """A worksheet, addressed the way Excel addresses one.

    Rows and columns are 1-based, as in Excel's own R1C1 notation.
    """

    def __init__(self, workbook: Workbook, part: WorksheetPart) -> None:
        """Wrap the sheet backed by ``part`` in ``workbook``."""
        self._workbook = workbook
        self._part = part

    def __repr__(self) -> str:
        """Show the tab name."""
        return f"<Sheet {self.name!r}>"

    @property
    def _sheet(self) -> Worksheet:
        """The live worksheet, which renaming replaces behind our back."""
        for sheet in self._workbook.worksheets:
            if sheet._part is self._part:
                return sheet
        message = "the sheet is no longer in the workbook"
        raise ValueError(message)

    @property
    def name(self) -> str:
        """The sheet's tab name.

        Read-only. Renaming goes through [`Book.set`][_excel.Book.set] on
        the name's own place, so that formulas referring to the sheet
        follow it.
        """
        return self._sheet.name

    @property
    def merged_ranges(self) -> list[tuple[int, int, int, int]]:
        """Every merged block on the sheet.

        Returns:
            ``(top, left, bottom, right)`` of each, all 1-based.
        """
        from xlsxedit.merge import parse_range
        from xlsxedit.oxml.address import col_to_index

        blocks = []
        for ref in self._sheet.merged_ranges:
            left, top, right, bottom = parse_range(ref)
            blocks.append(
                (top, col_to_index(left) + 1, bottom, col_to_index(right) + 1)
            )
        return blocks

    def duplicate_rows(
        self, top: int, bottom: int, *, copies: int
    ) -> list[tuple[int, int]]:
        """Copy rows ``top`` through ``bottom`` and insert the copies below.

        The copies carry the cells, their formats and their formulas;
        rows below move down; and everything that referred to the copied
        rows follows — formulas, merged cells, conditional formats,
        validation, links, shapes, tables, defined names.

        What a formula ends up pointing at is what Excel's "Insert
        Copied Cells" would leave it pointing at: a relative reference in
        a copy keeps the distance it had, an absolute one stays put, and
        a range reaching over the copied rows stretches over all of them.

        Args:
            top: 1-based first row to copy.
            bottom: 1-based last row to copy, inclusive.
            copies: How many copies to make. ``0`` changes nothing.

        Returns:
            ``(top, bottom)`` of each copy, top to bottom. The original
            block is not included.
        """
        return duplicate_rows(self._workbook, self._sheet, top, bottom, copies)

    def duplicate_columns(
        self, left: int, right: int, *, copies: int
    ) -> list[tuple[int, int]]:
        """Copy columns ``left`` through ``right`` and insert the copies right.

        The column-wise counterpart of
        [`duplicate_rows`][_excel.Sheet.duplicate_rows], with the same
        copy rule.

        Args:
            left: 1-based first column to copy.
            right: 1-based last column to copy, inclusive.
            copies: How many copies to make. ``0`` changes nothing.

        Returns:
            ``(left, right)`` of each copy, left to right. The original
            block is not included.
        """
        return duplicate_columns(self._workbook, self._sheet, left, right, copies)

    def delete_rows(self, top: int, bottom: int) -> None:
        """Delete rows ``top`` through ``bottom``.

        Rows below move up and references follow. A range that straddled
        the deleted rows closes up over them.

        Args:
            top: 1-based first row to delete.
            bottom: 1-based last row to delete, inclusive.
        """
        delete_rows(self._workbook, self._sheet, top, bottom)

    def delete_columns(self, left: int, right: int) -> None:
        """Delete columns ``left`` through ``right``.

        Args:
            left: 1-based first column to delete.
            right: 1-based last column to delete, inclusive.
        """
        delete_columns(self._workbook, self._sheet, left, right)
