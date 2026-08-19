"""Raw XML text helpers.

Parts are edited at the text level so untouched bytes survive verbatim.
These helpers decode entity-escaped runs and re-encode inserted values
with the same minimal escaping excelize (Go ``encoding/xml``) emits, so
edited parts stay byte-consistent with authored ones.
"""

from __future__ import annotations

import re

_ENTITY_RE = re.compile(
    r"&(?:#[xX](?P<hex>[0-9A-Fa-f]+)|#(?P<dec>[0-9]+)"
    r"|(?P<named>amp|lt|gt|quot|apos));",
)
_NAMED = {"amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'"}


def unescape(raw: str) -> str:
    """Decode XML character entities in ``raw``."""

    def _sub(m: re.Match[str]) -> str:
        if m.group("hex"):
            return chr(int(m.group("hex"), 16))
        if m.group("dec"):
            return chr(int(m.group("dec")))
        return _NAMED[m.group("named")]

    return _ENTITY_RE.sub(_sub, raw)


def escape_text(text: str) -> str:
    """Encode ``text`` for an XML text node (minimal escaping)."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def escape_attr(text: str) -> str:
    """Encode ``text`` for a double-quoted XML attribute value."""
    return (
        escape_text(text)
        .replace('"', "&quot;")
        .replace("\t", "&#x9;")
        .replace("\n", "&#xA;")
        .replace("\r", "&#xD;")
    )


def escape_go_text(text: str) -> str:
    """Encode ``text`` the way Go's ``encoding/xml`` escapes text nodes.

    excelize (Go) escapes the five XML specials plus whitespace controls;
    matching it keeps re-encoded template text byte-identical.
    """
    return (
        escape_text(text)
        .replace("'", "&#39;")
        .replace('"', "&#34;")
        .replace("\t", "&#x9;")
        .replace("\n", "&#xA;")
        .replace("\r", "&#xD;")
    )


def parse_attrs(tag: str) -> list[tuple[str, str]]:
    """Parse ``name="value"`` pairs from an XML start tag, in source order."""
    return re.findall(r'([^\s=/>"]+)="([^"]*)"', tag)


def emit_attrs(attrs: list[tuple[str, str]]) -> str:
    """Serialize attribute pairs back into start-tag text."""
    return "".join(f' {name}="{value}"' for name, value in attrs)
