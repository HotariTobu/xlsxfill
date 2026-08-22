import json
from io import BytesIO
from typing import TYPE_CHECKING, cast

import pytest
from conftest import DATA_RAISE, load_input

import xlsxfill
from xlsxfill import Value, fill

if TYPE_CHECKING:
    from collections.abc import Mapping

CASES = sorted(
    path.name.removesuffix(".template.xlsx")
    for path in DATA_RAISE.glob("*.template.xlsx")
)


@pytest.mark.parametrize("case", CASES)
def test_raise(case: str) -> None:
    error = json.loads((DATA_RAISE / f"{case}.error.json").read_text(encoding="utf-8"))
    exception_type = getattr(xlsxfill, error["type"])
    data = cast("Mapping[str, Value]", load_input(DATA_RAISE / f"{case}.input.json"))
    with pytest.raises(exception_type) as excinfo:
        fill(DATA_RAISE / f"{case}.template.xlsx", data, BytesIO())
    assert str(excinfo.value) == error["message"]
