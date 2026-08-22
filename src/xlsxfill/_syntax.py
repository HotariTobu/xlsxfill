from __future__ import annotations

import re
from dataclasses import dataclass, field

TYPE_NAMES = ("num", "str", "bool", "date", "time", "datetime")

_LINK_IMAGE_RE = re.compile(
    r"^\s*(?P<bang>!?)\[(?P<label>[^]]*)\]\((?P<target>[^)]*)\)(?P<rest>.*)$",
    re.DOTALL,
)
_ASSERT_RE = re.compile(
    rf"^(?P<path>.*?)\s*:\s*(?P<type>{'|'.join(TYPE_NAMES)})\s*$",
    re.DOTALL,
)
_FIT_NAMES = ("contain", "cover", "fill")


@dataclass(frozen=True)
class PropStep:
    name: str


@dataclass(frozen=True)
class IndexStep:
    symbol: str

    @property
    def is_fixed(self) -> bool:
        return self.symbol.isdigit()

    @property
    def is_sheet(self) -> bool:
        return self.symbol == "s"


type PathStep = PropStep | IndexStep


@dataclass(frozen=True)
class Literal:
    raw: str


@dataclass(frozen=True)
class ValueRef:
    src: str
    path: tuple[PathStep, ...] = ()
    assert_type: str | None = None
    error: str | None = None


@dataclass(frozen=True)
class Link:
    src: str
    label_path: tuple[PathStep, ...] | None = None
    url_path: tuple[PathStep, ...] = ()
    error: str | None = None


@dataclass(frozen=True)
class Image:
    src: str
    alt_path: tuple[PathStep, ...] | None = None
    data_path: tuple[PathStep, ...] = ()
    fit: str = "contain"
    error: str | None = None


@dataclass(frozen=True)
class Marker:
    src: str
    name: str = ""
    plus: bool = False
    error: str | None = None


type Construct = ValueRef | Link | Image | Marker
type Segment = Literal | Construct


@dataclass
class ParsedText:
    segments: list[Segment] = field(default_factory=list)

    @property
    def constructs(self) -> list[Construct]:
        return [seg for seg in self.segments if not isinstance(seg, Literal)]

    @property
    def is_static(self) -> bool:
        return not self.constructs


def parse_path(text: str) -> tuple[PathStep, ...] | None:
    steps: list[PathStep] = []
    rest = text
    first = True
    while True:
        m = re.match(r"^\s*([^.#]*?)\s*([.#]|$)", rest, re.DOTALL)
        if m is None:
            return None
        name, sep = m.group(1), m.group(2)
        if first and not name and sep == "#":
            pass  # leading "#index" (e.g. "${#r}")
        elif not name:
            return None
        else:
            steps.append(PropStep(name))
        rest = rest[m.end() :]
        while sep == "#":
            m = re.match(r"^\s*([^.#]*?)\s*([.#]|$)", rest, re.DOTALL)
            if m is None or not m.group(1):
                return None
            steps.append(IndexStep(m.group(1)))
            sep = m.group(2)
            rest = rest[m.end() :]
        if sep == "":
            return tuple(steps)
        first = False


def parse_marker(src: str, inner: str) -> Marker:
    body = inner.strip()
    if not body:
        return Marker(src, error="empty band marker")
    name, _, suffix = body.partition("+")
    name = name.strip()
    plus = False
    if suffix or "+" in body:
        if suffix.strip() != "1":
            return Marker(src, error='only "+1" is allowed')
        plus = True
    if name == "s":
        return Marker(src, error="sheets have no repetition declaration")
    if not name or name[0] not in "rc":
        return Marker(src, error='band name must start with "r" or "c"')
    return Marker(src, name=name, plus=plus)


def parse_placeholder(src: str, inner: str) -> ValueRef | Link | Image:
    if not inner.strip():
        return ValueRef(src, error="empty placeholder")

    link_match = _LINK_IMAGE_RE.match(inner)
    if link_match is not None:
        return _parse_link_or_image(src, link_match)

    assert_type: str | None = None
    body = inner
    assert_match = _ASSERT_RE.match(inner)
    if assert_match is not None:
        assert_type = assert_match.group("type")
        body = assert_match.group("path")
    path = parse_path(body)
    if path is None:
        return ValueRef(src, error="malformed path")
    return ValueRef(src, path=path, assert_type=assert_type)


def _parse_link_or_image(src: str, m: re.Match[str]) -> Link | Image:
    is_image = m.group("bang") == "!"
    label_text = m.group("label")
    target_text = m.group("target")
    rest = m.group("rest").strip()

    label_path: tuple[PathStep, ...] | None = None
    if label_text.strip():
        label_path = parse_path(label_text)
        if label_path is None:
            return _link_image_error(src, is_image=is_image, error="malformed path")
    target_path = parse_path(target_text)

    fit = "contain"
    if is_image and rest in _FIT_NAMES:
        fit, rest = rest, ""

    error: str | None = None
    if rest.startswith(":"):
        kind = "an image" if is_image else "a link"
        error = f"type assertion on {kind}"
    elif rest or target_path is None:
        error = "malformed path"

    if is_image:
        return Image(
            src,
            alt_path=label_path,
            data_path=target_path or (),
            fit=fit,
            error=error,
        )
    return Link(src, label_path=label_path, url_path=target_path or (), error=error)


def _link_image_error(src: str, *, is_image: bool, error: str) -> Link | Image:
    if is_image:
        return Image(src, error=error)
    return Link(src, error=error)


def tokenize(raw: str) -> ParsedText:
    segments: list[Segment] = []
    pos = 0
    while True:
        dollar = raw.find("${", pos)
        hash_ = raw.find("#{", pos)
        starts = [s for s in (dollar, hash_) if s != -1]
        if not starts:
            break
        start = min(starts)
        end = raw.find("}", start + 2)
        if end == -1:
            break
        if start > pos:
            segments.append(Literal(raw[pos:start]))
        src = raw[start : end + 1]
        inner = raw[start + 2 : end]
        if raw[start] == "$":
            segments.append(parse_placeholder(src, inner))
        else:
            segments.append(parse_marker(src, inner))
        pos = end + 1
    if pos < len(raw) or not segments:
        segments.append(Literal(raw[pos:]))
    return ParsedText(segments)
