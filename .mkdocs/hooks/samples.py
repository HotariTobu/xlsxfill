import json
from collections.abc import Iterable, Mapping
from textwrap import indent
from typing import Protocol, cast
from urllib.parse import quote

MARKER = "<!-- SAMPLES -->"
VIEWER = "https://view.officeapps.live.com/op/embed.aspx?src="

API_PATH = "api.md"
API_PACKAGE = "xlsxfill"

ASSET_PATH = "samples/"
TEMPLATE_SUFFIX = ".template.xlsx"
INPUT_SUFFIX = ".input.json"
EXPECTED_SUFFIX = ".expected.xlsx"
ERROR_SUFFIX = ".error.json"
TAB_INDENT = "    "


class _Inclusion(Protocol):
    def is_included(self) -> bool: ...


class _Asset(Protocol):
    src_uri: str
    dest_uri: str
    content_string: str
    inclusion: _Inclusion


def _viewer(url: str) -> str:
    src = VIEWER + quote(url, safe="")
    return (
        f'<iframe src="{src}" width="100%" height="480" '
        'style="border:0" loading="lazy"></iframe>'
    )


def _fence(language: str, body: str) -> str:
    return f"```{language}\n{body}\n```"


def _reference(name: str) -> str:
    return f"[{name}]({API_PATH}#{API_PACKAGE}.{name})"


def _outcome(stem: str, base: str, assets: Mapping[str, _Asset]) -> tuple[str, str]:
    expected = assets.get(stem + EXPECTED_SUFFIX)
    if expected is not None:
        return "Expected", _viewer(base + expected.dest_uri)

    error = cast(
        "Mapping[str, str]", json.loads(assets[stem + ERROR_SUFFIX].content_string)
    )
    message = indent(_fence("text", error["message"]), TAB_INDENT)
    return "Error", f'!!! failure "{_reference(error["type"])}"\n\n{message}'


def _block(stem: str, name: str, base: str, assets: Mapping[str, _Asset]) -> str:
    data = assets[stem + INPUT_SUFFIX].content_string.strip()
    template = base + assets[stem + TEMPLATE_SUFFIX].dest_uri
    label, outcome = _outcome(stem, base, assets)
    return "\n".join(
        (
            f"### {name}",
            "",
            '=== "Template"',
            "",
            indent(_viewer(template), TAB_INDENT),
            "",
            '===+ "Input"',
            "",
            indent(_fence("json", data), TAB_INDENT),
            "",
            f'=== "{label}"',
            "",
            indent(outcome, TAB_INDENT),
            "",
        )
    )


def _body(base: str, assets: Mapping[str, _Asset]) -> str:
    groups: dict[str, list[str]] = {}
    for uri in assets:
        if not uri.endswith(TEMPLATE_SUFFIX):
            continue
        group, _, name = (
            uri.removeprefix(ASSET_PATH).removesuffix(TEMPLATE_SUFFIX).rpartition("/")
        )
        groups.setdefault(group, []).append(name)

    sections = []
    for group in sorted(groups):
        sections.append(f"## {group.capitalize()}\n")
        stem = f"{ASSET_PATH}{group}/" if group else ASSET_PATH
        sections.extend(
            _block(stem + name, name, base, assets) for name in sorted(groups[group])
        )
    return "\n".join(sections)


def on_page_markdown(
    markdown: str,
    *,
    config: Mapping[str, object],
    files: Iterable[_Asset],
    **_kwargs: object,
) -> str:
    if MARKER not in markdown:
        return markdown

    site_url = config["site_url"]
    if not isinstance(site_url, str):
        msg = "site_url must be set to build the samples page"
        raise TypeError(msg)

    assets = {
        file.src_uri: file
        for file in files
        if file.src_uri.startswith(ASSET_PATH) and file.inclusion.is_included()
    }
    return markdown.replace(MARKER, _body(site_url.rstrip("/") + "/", assets))
