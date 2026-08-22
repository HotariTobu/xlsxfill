from __future__ import annotations

import argparse
import sys
import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree

SML = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
SST = "xl/sharedStrings.xml"
XMLISH = (".xml", ".rels", ".vml")

PACKAGE_PARTS = ("[Content_Types].xml", ".rels")


class RefusedError(Exception):
    pass


def members(raw: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(BytesIO(raw)) as zipf:
        return {
            info.filename: zipf.read(info.filename)
            for info in zipf.infolist()
            if not info.filename.endswith("/")
        }


def repack(raw: bytes, new: dict[str, bytes]) -> bytes:
    out = BytesIO()
    with zipfile.ZipFile(BytesIO(raw)) as zipf, zipfile.ZipFile(out, "w") as dest:
        for info in zipf.infolist():
            if info.filename.endswith("/"):
                continue
            copy = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            copy.compress_type = info.compress_type
            copy.external_attr = info.external_attr
            copy.create_system = info.create_system
            dest.writestr(copy, new[info.filename])
    return out.getvalue()


def restamp(original: bytes, made: bytes) -> bytes:
    with zipfile.ZipFile(BytesIO(original)) as before:
        stamps = {
            info.filename: info
            for info in before.infolist()
            if not info.filename.endswith("/")
        }
    out = BytesIO()
    with zipfile.ZipFile(BytesIO(made)) as after, zipfile.ZipFile(out, "w") as dest:
        for info in after.infolist():
            if info.filename.endswith("/"):
                continue
            was = stamps.get(info.filename, info)
            copy = zipfile.ZipInfo(info.filename, date_time=was.date_time)
            copy.compress_type = was.compress_type
            copy.external_attr = was.external_attr
            copy.create_system = was.create_system
            dest.writestr(copy, after.read(info.filename))
    return out.getvalue()


def flatten(raw: bytes) -> list[str]:
    lines: list[str] = []

    def walk(node: etree._Element, path: str) -> None:
        if not isinstance(node.tag, str):
            return
        attrs = sorted(node.attrib.items())
        lines.append(f"{path}|{node.tag}|{attrs}|{node.text!r}|{node.tail!r}")
        seen: dict[str, int] = {}
        for child in node:
            if not isinstance(child.tag, str):
                continue
            index = seen.get(child.tag, 0)
            seen[child.tag] = index + 1
            walk(child, f"{path}/{child.tag}[{index}]")

    walk(etree.fromstring(raw), "")
    return lines


def is_package_part(name: str) -> bool:
    return name.endswith(PACKAGE_PARTS)


def package(raw: bytes) -> tuple[bytes, list[str]]:
    from xlsxedit.workbook import Workbook

    out = BytesIO()
    Workbook.open(BytesIO(raw)).save(out)
    made = out.getvalue()

    before, after = members(raw), members(made)
    lost = sorted(set(before) - set(after))
    gained = sorted(set(after) - set(before))
    if lost or gained:
        message = f"saving through xlsxedit lost {lost} and gained {gained}"
        raise RefusedError(message)
    for name, blob in before.items():
        if is_package_part(name):
            continue
        if not name.lower().endswith(XMLISH):
            if after[name] != blob:
                message = f"saving through xlsxedit rewrote the binary part {name}"
                raise RefusedError(message)
            continue
        if flatten(after[name]) != flatten(blob):
            message = f"saving through xlsxedit changed what {name} holds"
            raise RefusedError(message)

    changed = sorted(name for name in before if after[name] != before[name])
    return made, changed


def serialize(raw: bytes) -> tuple[bytes, list[str]]:
    old = members(raw)
    new: dict[str, bytes] = {}
    for name, blob in old.items():
        if not name.lower().endswith(XMLISH):
            new[name] = blob
            continue
        made = etree.tostring(
            etree.fromstring(blob),
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        if flatten(made) != flatten(blob):
            message = f"re-serializing {name} would change what it holds"
            raise RefusedError(message)
        new[name] = made
    changed = sorted(name for name in old if new[name] != old[name])
    return repack(raw, new), changed


def references(parts: dict[str, bytes]) -> int:
    total = 0
    for name, blob in parts.items():
        if not (name.startswith("xl/worksheets/") and name.endswith(".xml")):
            continue
        for cell in etree.fromstring(blob).iter(f"{SML}c"):
            if cell.get("t") == "s" and cell.find(f"{SML}v") is not None:
                total += 1
    return total


def count(raw: bytes) -> tuple[bytes, str | None]:
    parts = members(raw)
    if SST not in parts:
        return raw, None

    before = etree.fromstring(parts[SST])
    old_count = before.get("count")
    old_unique = before.get("uniqueCount")
    new_count = str(references(parts))
    new_unique = str(len(before.findall(f"{SML}si")))
    if (old_count, old_unique) == (new_count, new_unique):
        return raw, None

    after = etree.fromstring(parts[SST])
    after.set("count", new_count)
    after.set("uniqueCount", new_unique)
    made = etree.tostring(
        after, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    check = etree.fromstring(made)
    for attribute, was in (("count", old_count), ("uniqueCount", old_unique)):
        if was is None:
            check.attrib.pop(attribute, None)
        else:
            check.set(attribute, was)
    if flatten(etree.tostring(check)) != flatten(parts[SST]):
        message = "restating the shared string counts would change an entry"
        raise RefusedError(message)

    parts[SST] = made
    told = f"count {old_count} -> {new_count}, uniqueCount {old_unique} -> {new_unique}"
    return repack(raw, parts), told


def canonize(raw: bytes) -> tuple[bytes, list[str]]:
    told: list[str] = []
    raw, repackaged = package(raw)
    if repackaged:
        told.append(f"package form:  {', '.join(repackaged)}")
    raw, reserialized = serialize(raw)
    if reserialized:
        told.append(f"serialization: {len(reserialized)} parts")
    raw, recounted = count(raw)
    if recounted:
        told.append(f"shared strings: {recounted}")
    return raw, told


def settled(raw: bytes) -> tuple[bytes, list[str]]:
    made, told = canonize(raw)
    return restamp(raw, made), told


def run(path: Path) -> list[str]:
    if not path.name.endswith(".template.xlsx"):
        message = (
            f"{path.name} is not a template. Canonizing an expected would"
            " re-sort [Content_Types].xml away from what a fill produces"
        )
        raise RefusedError(message)
    raw = path.read_bytes()
    made, told = settled(raw)
    again, _ = settled(made)
    if again != made:
        message = f"{path.name}: canonizing is not settled after one pass"
        raise RefusedError(message)
    if made == raw:
        return []
    path.write_bytes(made)
    return told or ["stored differently, holding the same thing"]


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Put generated templates into the form the fixture suite keeps.",
    )
    parser.add_argument(
        "paths",
        nargs="+",
        type=Path,
        metavar="TEMPLATE",
        help="a *.template.xlsx to canonize in place",
    )
    for path in parser.parse_args(argv).paths:
        try:
            told = run(path)
        except RefusedError as refused:
            print(f"{path}\n  refused: {refused}")
            return 1
        print(path)
        for line in told or ["already canonical"]:
            print(f"  {line}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
