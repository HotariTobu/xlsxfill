"""The entry point."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, BinaryIO

from _patched_xlsxedit import Workbook
from xlsxfill._book import Book
from xlsxfill._validate import validate

if TYPE_CHECKING:
    from collections.abc import Mapping

    from xlsxfill._problems import Problem
    from xlsxfill._values import Value


def fill(
    template: str | Path | BinaryIO,
    data: Mapping[str, Value],
    output: str | Path | BinaryIO,
) -> list[Problem]:
    """Merge a template xlsx with data and write the result.

    Args:
        template: The template xlsx.
        data: The data to merge into the template.
        output: Where the resulting xlsx is written.

    Returns:
        The ``#SYNTAX!`` / ``#DATA!`` errors embedded in the output.

    Raises:
        DataError: The input data as a whole is unusable.
    """
    validate(data)
    wb = Workbook.open(template)
    problems = Book(wb, data).run()
    if isinstance(output, str | Path):
        with Path(output).open("wb") as stream:
            wb.save(stream)
    else:
        wb.save(output)
    return problems
