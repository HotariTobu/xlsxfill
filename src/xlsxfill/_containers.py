"""Substitution in non-cell string containers.

Sheet names, headers/footers, shape and chart text, comments, hyperlink
tooltips, data-validation messages, and document properties are plain
string containers: values concatenate as text, and type assertions,
links, images, and band markers are syntax errors there.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Protocol

from _patched_xlsxedit import emit_attrs, parse_attrs, unescape
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
    from collections.abc import Callable, Mapping

    from xlsxfill._values import Value


class ProblemSink(Protocol):
    """Callback receiving ``(kind, construct, reason)`` for each failure."""

    def __call__(self, kind: str, construct: str, reason: str, /) -> str:
        """Record the problem and return the message to embed."""
        ...


def substitute_container(
    raw: str,
    data: Mapping[str, Value],
    bindings: Mapping[str, int],
    escape: Callable[[str], str],
    report: ProblemSink,
    declared_bands: frozenset[str] = frozenset(),
) -> str | None:
    """Substitute placeholders in raw (escaped) container text.

    ``declared_bands`` are the valid band names of the enclosing sheet,
    used to distinguish a band name used outside its band from an
    undeclared one. Returns the new raw text, or ``None`` when the text
    is static so the caller can leave the original bytes untouched.
    """
    parsed = tokenize(raw)
    if parsed.is_static:
        return None
    pieces: list[str] = []
    for segment in parsed.segments:
        if isinstance(segment, Literal):
            pieces.append(segment.raw)
            continue
        if isinstance(segment, Marker):
            reason = segment.error or "band marker outside a cell"
            pieces.append(escape(report("syntax", segment.src, reason)))
            continue
        if isinstance(segment, Link | Image):
            kind = "link" if isinstance(segment, Link) else "image"
            reason = segment.error or f"{kind} outside a cell"
            pieces.append(escape(report("syntax", segment.src, reason)))
            continue
        pieces.append(
            escape(_value_text(segment, data, bindings, report, declared_bands))
        )
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
        return report(
            "syntax",
            ref.src,
            "type assertion outside a cell",
        )
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


def substitute_text_nodes(
    raw: str,
    node_tags: tuple[str, ...],
    data: Mapping[str, Value],
    bindings: Mapping[str, int],
    escape: Callable[[str], str],
    report: ProblemSink,
    declared_bands: frozenset[str] = frozenset(),
) -> str:
    """Substitute placeholders inside ``<tag>…</tag>`` text nodes.

    ``node_tags`` are literal tag names (``"t"``, ``"a:t"``); the nodes
    must not contain child elements.
    """
    out = raw
    for tag in node_tags:
        out = _substitute_tag(out, tag, data, bindings, escape, report, declared_bands)
    return out


def _substitute_tag(
    raw: str,
    tag: str,
    data: Mapping[str, Value],
    bindings: Mapping[str, int],
    escape: Callable[[str], str],
    report: ProblemSink,
    declared_bands: frozenset[str],
) -> str:
    pattern = re.compile(
        rf"(<{re.escape(tag)}(?:\s[^>]*)?>)(.*?)(</{re.escape(tag)}>)",
        re.DOTALL,
    )

    def _sub(m: re.Match[str]) -> str:
        replaced = substitute_container(
            m.group(2), data, bindings, escape, report, declared_bands
        )
        if replaced is None:
            return m.group(0)
        return f"{m.group(1)}{replaced}{m.group(3)}"

    return pattern.sub(_sub, raw)


def substitute_attrs(
    raw: str,
    tag: str,
    attr_names: tuple[str, ...],
    data: Mapping[str, Value],
    bindings: Mapping[str, int],
    escape: Callable[[str], str],
    report: ProblemSink,
    declared_bands: frozenset[str] = frozenset(),
) -> str:
    """Substitute placeholders in selected attributes of ``<tag …>`` tags."""
    pattern = re.compile(rf"<{re.escape(tag)}((?:\s[^>]*)?)>")

    def _sub(m: re.Match[str]) -> str:
        attrs = parse_attrs(m.group(1))
        changed = False
        out: list[tuple[str, str]] = []
        for name, value in attrs:
            if name in attr_names:
                replaced = substitute_container(
                    value, data, bindings, escape, report, declared_bands
                )
                if replaced is not None:
                    out.append((name, replaced))
                    changed = True
                    continue
            out.append((name, value))
        if not changed:
            return m.group(0)
        return f"<{tag}{emit_attrs(out)}>"

    return pattern.sub(_sub, raw)


def decoded_equals(raw: str, text: str) -> bool:
    """Whether raw XML text decodes to ``text``."""
    return unescape(raw) == text
