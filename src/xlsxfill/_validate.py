"""Input-data validation before a fill run."""

from __future__ import annotations

from collections.abc import Mapping

from xlsxfill._exceptions import DataError
from xlsxfill._resolve import kind_of

_FORBIDDEN_NAME_CHARS = ".#:()[]!${}"


def _root_kind(data: object) -> str:
    kind = kind_of(data)
    return "array" if kind == "list" else kind


def _validate_names(value: object) -> None:
    if isinstance(value, Mapping):
        for name, child in value.items():
            for char in name:
                if char in _FORBIDDEN_NAME_CHARS:
                    raise DataError(
                        f'property name "{name}" contains a forbidden '
                        f'character: "{char}"',
                    )
            _validate_names(child)
    elif isinstance(value, list):
        for child in value:
            _validate_names(child)


def validate(data: object) -> None:
    """Reject input whose shape or property names are unusable.

    Raises:
        DataError: The input data as a whole is unusable.
    """
    if not isinstance(data, Mapping):
        raise DataError(f"input root must be an object, got {_root_kind(data)}")
    _validate_names(data)
