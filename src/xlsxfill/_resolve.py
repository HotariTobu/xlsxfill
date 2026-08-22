from __future__ import annotations

from collections.abc import Mapping, Sequence
from datetime import date, datetime, time
from typing import TYPE_CHECKING

from xlsxfill._syntax import IndexStep, PropStep

if TYPE_CHECKING:
    from xlsxfill._syntax import PathStep
    from xlsxfill._values import Value

_EPOCH = date(1899, 12, 30)


def kind_of(value: object) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, int | float):
        return "num"
    if isinstance(value, str):
        return "str"
    if isinstance(value, datetime):
        return "datetime"
    if isinstance(value, date):
        return "date"
    if isinstance(value, time):
        return "time"
    if isinstance(value, bytes):
        return "bytes"
    if isinstance(value, Mapping):
        return "dict"
    if isinstance(value, Sequence):
        return "list"
    return type(value).__name__


class ResolveError(Exception):
    def __init__(self, reason: str) -> None:
        super().__init__(reason)
        self.reason = reason


def resolve(
    path: tuple[PathStep, ...],
    data: Mapping[str, Value],
    bindings: Mapping[str, int],
) -> Value:
    if len(path) == 1 and isinstance(path[0], IndexStep) and not path[0].is_fixed:
        return bindings[path[0].symbol] + 1

    value: Value = dict(data)
    prev_name = ""
    for step in path:
        if isinstance(step, PropStep):
            if value is None:
                raise ResolveError(f'"{prev_name}" is null')
            if not isinstance(value, Mapping):
                raise ResolveError(f'"{prev_name}" is not an object')
            value = value.get(step.name)
            prev_name = step.name
        else:
            if value is None:
                raise ResolveError(f'"{prev_name}" is null')
            if not isinstance(value, list):
                raise ResolveError(f'"{prev_name}" is not a list')
            if step.is_fixed:
                index = int(step.symbol)
                if index >= len(value):
                    raise ResolveError(
                        f"index {index} out of range ({len(value)} items)",
                    )
            else:
                index = bindings[step.symbol]
                if index >= len(value):
                    return None
            value = value[index]
    return value


def collection_length(
    path: tuple[PathStep, ...],
    band: str,
    data: Mapping[str, Value],
    bindings: Mapping[str, int],
) -> int | None:
    prefix: list[PathStep] = []
    for step in path:
        if isinstance(step, IndexStep) and step.symbol == band:
            break
        prefix.append(step)
    else:
        return None
    if not prefix:
        return None
    try:
        value = resolve(tuple(prefix), data, bindings)
    except (ResolveError, KeyError):
        return None
    if isinstance(value, list):
        return len(value)
    return None


def check_assertion(value: Value, expected: str) -> None:
    actual = kind_of(value)
    if actual != expected:
        raise ResolveError(
            f"type assertion failed: expected {expected}, got {actual}",
        )


def date_serial(value: date) -> int:
    return (value - _EPOCH).days


def time_fraction(value: time) -> float:
    seconds = value.hour * 3600 + value.minute * 60 + value.second
    return (seconds + value.microsecond / 1_000_000) / 86400


def render_number(value: Value) -> str:
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            raise ResolveError("tz-aware datetime is not supported (make it naive)")
        return repr(date_serial(value.date()) + time_fraction(value.time()))
    if isinstance(value, date):
        return repr(date_serial(value))
    if isinstance(value, time):
        return repr(time_fraction(value))
    return repr(value)


def check_scalar(value: Value) -> None:
    actual = kind_of(value)
    if actual in ("dict", "list"):
        raise ResolveError(f"referenced value is not a scalar ({actual})")


def concat_text(value: Value) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return render_number(value)
