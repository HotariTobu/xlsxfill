from collections.abc import Iterable, Mapping
from pathlib import Path
from re import Match, sub
from typing import Protocol, cast

INDEX_URI = "llms.txt"
FULL_URI = "llms-full.txt"

API_URI = "api.md"
API_PACKAGE = "xlsxfill"
API_HANDLER = "python"
API_PLUGIN = "mkdocstrings"

KIND_FUNCTION = "function"
KIND_CLASS = "class"
KIND_ALIAS = "type alias"
KIND_ATTRIBUTE = "attribute"

PAGES = {
    "index.md": (
        "Home",
        "What the library does, how the output is produced, and a worked example.",
    ),
    "syntax.md": (
        "Syntax",
        "The placeholder notation in full, and the message each construct embeds when it fails.",
    ),
    "api.md": ("API", "The signature of every public symbol, with its docstring."),
}

LINK = r"\]\((?P<uri>[\w./-]+\.md)(?P<anchor>#[\w.-]*)?\)"
CROSSREF = r"\[(?P<label>[^\]]+)\]\[(?P<identifier>[\w.]+)\]"

_URLS: dict[str, str] = {}
_MARKDOWN: dict[str, str] = {}


class _Docstring(Protocol):
    value: str


class _Kind(Protocol):
    value: str


class _Parameter(Protocol):
    name: str
    annotation: object


class _Member(Protocol):
    name: str
    kind: _Kind
    docstring: _Docstring | None
    annotation: object


class _Object(Protocol):
    name: str
    kind: _Kind
    docstring: _Docstring | None
    exports: list[str] | None
    members: Mapping[str, "_Object"]
    is_alias: bool
    final_target: "_Object"
    bases: list[object]
    parameters: Iterable[_Parameter]
    returns: object
    value: object


class _Handler(Protocol):
    def get_options(self, local_options: Mapping[str, object]) -> object: ...
    def collect(self, identifier: str, options: object) -> _Object: ...


class _Plugin(Protocol):
    def get_handler(self, handler_name: str) -> _Handler: ...


class _File(Protocol):
    src_uri: str
    url: str


class _Page(Protocol):
    file: _File


def _fence(body: str) -> str:
    return f"```python\n{body}\n```"


def _absolute(markdown: str) -> str:
    def replace(match: Match[str]) -> str:
        url = _URLS.get(match["uri"])
        if url is None:
            return match[0]
        return f"]({url}{match['anchor'] or ''})"

    return sub(LINK, replace, markdown)


def _crossrefs(text: str) -> str:
    def replace(match: Match[str]) -> str:
        return f"[{match['label']}]({_URLS[API_URI]}#{match['identifier']})"

    return sub(CROSSREF, replace, text)


def _docstring(obj: _Object) -> list[str]:
    if obj.docstring is None:
        return []
    return [_crossrefs(obj.docstring.value.strip()), ""]


def _signature(obj: _Object) -> str:
    parameters = ", ".join(f"{p.name}: {p.annotation}" for p in obj.parameters)
    return f"{obj.name}({parameters}) -> {obj.returns}"


def _declaration(obj: _Object) -> str:
    if obj.kind.value == KIND_FUNCTION:
        return _signature(obj)
    if obj.kind.value == KIND_ALIAS:
        return f"{obj.name} = {obj.value}"
    bases = ", ".join(str(base) for base in obj.bases)
    return f"class {obj.name}({bases})" if bases else f"class {obj.name}"


def _attributes(obj: _Object) -> list[str]:
    if obj.kind.value != KIND_CLASS:
        return []

    lines = []
    for name, member in obj.members.items():
        attribute = cast("_Member", member)
        if name.startswith("_") or attribute.kind.value != KIND_ATTRIBUTE:
            continue
        entry = f"- `{name}: {attribute.annotation}`"
        if attribute.docstring is not None:
            summary = " ".join(attribute.docstring.value.split())
            entry += f" — {_crossrefs(summary)}"
        lines.append(entry)

    return [*lines, ""] if lines else []


def _api(module: _Object) -> str:
    sections = ["# API", ""]
    for name in module.exports or []:
        member = module.members[name]
        obj = member.final_target if member.is_alias else member
        sections.extend(
            (
                f"## {name}",
                "",
                _fence(_declaration(obj)),
                "",
                *_docstring(obj),
                *_attributes(obj),
            )
        )
    return "\n".join(sections).rstrip()


def _index(name: str, summary: str, base: str) -> str:
    lines = [f"# {name}", "", f"> {summary}", "", "## Docs", ""]
    for uri, (title, description) in PAGES.items():
        lines.append(f"- [{title}]({_URLS[uri]}): {description}")
    lines.extend(
        (
            "",
            "## Optional",
            "",
            (
                f"- [{FULL_URI}]({base}{FULL_URI}): the pages above, "
                "concatenated into one file."
            ),
        )
    )
    return "\n".join(lines)


def _tail(file: _File) -> str:
    if file.src_uri in PAGES:
        return file.src_uri
    return "" if file.url == "./" else file.url


def _full() -> str:
    return "\n\n".join(_MARKDOWN[uri] for uri in PAGES)


def on_files(
    files: Iterable[_File], *, config: Mapping[str, object], **_kwargs: object
) -> Iterable[_File]:
    site_url = config["site_url"]
    if not isinstance(site_url, str):
        msg = "site_url must be set to build the llms.txt files"
        raise TypeError(msg)

    base = site_url.rstrip("/") + "/"
    _URLS.clear()
    _URLS.update(
        (file.src_uri, base + _tail(file))
        for file in files
        if file.src_uri.endswith(".md")
    )

    plugin = cast(
        "_Plugin", cast("Mapping[str, object]", config["plugins"])[API_PLUGIN]
    )
    handler = plugin.get_handler(API_HANDLER)
    module = handler.collect(API_PACKAGE, handler.get_options({}))

    _MARKDOWN.clear()
    _MARKDOWN[API_URI] = _api(module)
    return files


def on_page_markdown(markdown: str, *, page: _Page, **_kwargs: object) -> str:
    uri = page.file.src_uri
    if uri in PAGES and uri != API_URI:
        _MARKDOWN[uri] = _absolute(markdown.strip())
    return markdown


def on_post_build(*, config: Mapping[str, object], **_kwargs: object) -> None:
    site_dir = Path(str(config["site_dir"]))
    site_url = str(config["site_url"]).rstrip("/") + "/"

    pages = {
        INDEX_URI: _index(
            str(config["site_name"]), str(config["site_description"]), site_url
        ),
        FULL_URI: _full(),
        **_MARKDOWN,
    }
    for uri, text in pages.items():
        (site_dir / uri).write_text(text + "\n", encoding="utf-8")
