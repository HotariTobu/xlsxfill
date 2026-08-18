"""Golden tests: every member of the output must match expected.xlsx exactly."""

import zipfile
from io import BytesIO
from typing import cast

import pytest
from conftest import DATA_GOLDEN, load_input

from xlsxfill import Value, fill

BOOKS = sorted(
    path.name.removesuffix(".template.xlsx")
    for path in DATA_GOLDEN.glob("*.template.xlsx")
)


def _members(source: BytesIO) -> dict[str, bytes]:
    with zipfile.ZipFile(source) as zipf:
        return {
            name: zipf.read(name) for name in zipf.namelist() if not name.endswith("/")
        }


def assert_books_equal(actual_io: BytesIO, expected_io: BytesIO) -> None:
    """Assert that two workbooks have identical members with identical bytes."""
    actual = _members(actual_io)
    expected = _members(expected_io)
    assert sorted(actual) == sorted(expected)
    for name in expected:
        assert actual[name] == expected[name], name


@pytest.mark.parametrize("book", BOOKS)
def test_golden(book: str) -> None:
    """Fill the template with the input and compare against the expected."""
    data = cast("dict[str, Value]", load_input(DATA_GOLDEN / f"{book}.input.json"))
    output = BytesIO()
    fill(DATA_GOLDEN / f"{book}.template.xlsx", data, output)
    output.seek(0)
    assert_books_equal(
        output,
        BytesIO((DATA_GOLDEN / f"{book}.expected.xlsx").read_bytes()),
    )
