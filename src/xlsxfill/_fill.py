"""The entry point."""

from collections.abc import Mapping
from pathlib import Path
from typing import BinaryIO

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
    raise NotImplementedError
