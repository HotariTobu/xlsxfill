from datetime import date, datetime, time

type Value = (
    str
    | int
    | float
    | bool
    | date
    | time
    | datetime
    | bytes
    | list[Value]
    | dict[str, Value]
    | None
)
"""A value in the input data."""
