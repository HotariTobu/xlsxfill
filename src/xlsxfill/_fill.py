from __future__ import annotations

from typing import TYPE_CHECKING, BinaryIO

from _excel import Book as Excel
from xlsxfill._book import Book
from xlsxfill._validate import validate

if TYPE_CHECKING:
    from collections.abc import Mapping
    from pathlib import Path

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
        One problem per [message][xlsxfill.BaseProblem.message] embedded in the
        output.

    Raises:
        DataError: ``data`` is unusable as a whole.
    """
    validate(data)
    book = Excel.open(template)
    problems = Book(book, data).run()
    book.save(output)
    return problems
