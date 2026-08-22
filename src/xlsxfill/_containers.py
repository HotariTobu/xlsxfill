from __future__ import annotations

from typing import TYPE_CHECKING, Protocol

from xlsxfill._resolve import ResolveError, check_scalar, concat_text, resolve
from xlsxfill._syntax import (
    Image,
    IndexStep,
    Link,
    Literal,
    Marker,
    ValueRef,
    tokenize,
)

if TYPE_CHECKING:
    from collections.abc import Mapping

    from xlsxfill._values import Value


class ProblemSink(Protocol):
    def __call__(self, kind: str, construct: str, reason: str, /) -> str: ...


def substitute_container(
    text: str,
    data: Mapping[str, Value],
    bindings: Mapping[str, int],
    report: ProblemSink,
    declared_bands: frozenset[str] = frozenset(),
) -> str | None:
    parsed = tokenize(text)
    if parsed.is_static:
        return None
    pieces: list[str] = []
    for segment in parsed.segments:
        if isinstance(segment, Literal):
            pieces.append(segment.raw)
        elif isinstance(segment, Marker):
            reason = segment.error or "band marker outside a cell"
            pieces.append(report("syntax", segment.src, reason))
        elif isinstance(segment, Link | Image):
            kind = "link" if isinstance(segment, Link) else "image"
            reason = segment.error or f"{kind} outside a cell"
            pieces.append(report("syntax", segment.src, reason))
        else:
            pieces.append(_value_text(segment, data, bindings, report, declared_bands))
    return "".join(pieces)


def _value_text(
    ref: ValueRef,
    data: Mapping[str, Value],
    bindings: Mapping[str, int],
    report: ProblemSink,
    declared_bands: frozenset[str],
) -> str:
    if ref.error is not None:
        return report("syntax", ref.src, ref.error)
    if ref.assert_type is not None:
        return report("syntax", ref.src, "type assertion outside a cell")
    for step in ref.path:
        if (
            isinstance(step, IndexStep)
            and not step.is_fixed
            and step.symbol not in bindings
        ):
            if step.symbol in declared_bands:
                reason = f'band "{step.symbol}" used outside its band'
            else:
                reason = f'band "{step.symbol}" is not declared'
            return report("syntax", ref.src, reason)
    try:
        value = resolve(ref.path, data, bindings)
        check_scalar(value)
        return concat_text(value)
    except ResolveError as error:
        return report("data", ref.src, error.reason)
