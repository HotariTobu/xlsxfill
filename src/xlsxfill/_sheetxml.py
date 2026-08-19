"""Text-level worksheet document model.

A worksheet part is split into the text before ``<sheetData>``, the parsed
rows/cells, and the text after ``</sheetData>``. Rows and cells keep their
attributes in source order with raw values, so re-emitting an untouched
sheet reproduces the original bytes.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from _patched_xlsxedit import (
    col_to_index,
    emit_attrs,
    index_to_col,
    parse_attrs,
    split_address,
)

_SHEET_DATA_RE = re.compile(r"(<sheetData(?:\s[^>]*)?>)(.*)(</sheetData>)", re.DOTALL)
_ROW_RE = re.compile(r"<row((?:\s[^>]*)?)>(.*?)</row>", re.DOTALL)
_CELL_RE = re.compile(r"<c((?:\s[^>]*)?)>(.*?)</c>", re.DOTALL)
_V_RE = re.compile(r"<v>(.*?)</v>", re.DOTALL)
_F_RE = re.compile(r"<f(?:\s[^>]*)?>(.*?)</f>", re.DOTALL)
_COL_RE = re.compile(r"<col((?:\s[^>]*)?)></col>", re.DOTALL)
_COLS_RE = re.compile(r"<cols>.*?</cols>", re.DOTALL)


@dataclass
class Cell:
    """One ``<c>`` element."""

    col: int
    attrs: list[tuple[str, str]]
    inner: str
    bindings: dict[str, int] = field(default_factory=dict)
    notes: dict[int, str | None] = field(default_factory=dict)

    def get(self, name: str) -> str | None:
        """Return the raw value of attribute ``name``, or ``None``."""
        for attr, value in self.attrs:
            if attr == name:
                return value
        return None

    def set(self, name: str, value: str) -> None:
        """Set attribute ``name``, keeping its source position if present.

        A new attribute is appended at the end (after ``r`` / ``s``), which
        is the canonical position for ``t``.
        """
        for i, (attr, _) in enumerate(self.attrs):
            if attr == name:
                self.attrs[i] = (name, value)
                return
        self.attrs.append((name, value))

    def unset(self, name: str) -> None:
        """Remove attribute ``name`` if present."""
        self.attrs = [(attr, value) for attr, value in self.attrs if attr != name]

    @property
    def value_text(self) -> str | None:
        """The raw text of the ``<v>`` child, or ``None``."""
        m = _V_RE.search(self.inner)
        return m.group(1) if m else None

    @property
    def formula_text(self) -> str | None:
        """The raw text of the ``<f>`` child, or ``None``."""
        m = _F_RE.search(self.inner)
        return m.group(1) if m else None

    def set_formula_text(self, raw: str) -> None:
        """Replace the raw text of the ``<f>`` child."""
        self.inner = _F_RE.sub(
            lambda m: m.group(0).replace(m.group(1), raw, 1), self.inner, count=1
        )

    @property
    def sst_index(self) -> int | None:
        """The shared-string index when the cell is a ``t="s"`` cell."""
        if self.get("t") != "s":
            return None
        value = self.value_text
        return int(value) if value is not None else None

    def set_address(self, col: int, row: int) -> None:
        """Move the cell to 0-based column ``col`` in 1-based ``row``."""
        self.col = col
        self.set("r", f"{index_to_col(col)}{row}")

    def emit(self) -> str:
        """Serialize the cell."""
        return f"<c{emit_attrs(self.attrs)}>{self.inner}</c>"

    def copy(self) -> Cell:
        """Return an independent copy."""
        return Cell(
            self.col,
            list(self.attrs),
            self.inner,
            dict(self.bindings),
            dict(self.notes),
        )


@dataclass
class Row:
    """One ``<row>`` element."""

    num: int
    attrs: list[tuple[str, str]]
    cells: list[Cell]

    def set_num(self, num: int) -> None:
        """Renumber the row and its cells."""
        self.num = num
        self.attrs = [
            (name, str(num) if name == "r" else value) for name, value in self.attrs
        ]
        for cell in self.cells:
            cell.set_address(cell.col, num)

    def emit(self) -> str:
        """Serialize the row."""
        cells = "".join(cell.emit() for cell in self.cells)
        return f"<row{emit_attrs(self.attrs)}>{cells}</row>"

    def copy(self) -> Row:
        """Return an independent copy."""
        return Row(self.num, list(self.attrs), [cell.copy() for cell in self.cells])


class SheetDoc:
    """A worksheet part opened for text-level editing."""

    def __init__(self, raw: bytes) -> None:
        """Parse the raw bytes of a ``xl/worksheets/sheetN.xml`` part."""
        text = raw.decode()
        m = _SHEET_DATA_RE.search(text)
        if m is None:
            raise ValueError("worksheet part has no sheetData element")
        self.pre = text[: m.start()] + m.group(1)
        self.post = m.group(3) + text[m.end() :]
        self.rows: list[Row] = [
            _parse_row(row_match) for row_match in _ROW_RE.finditer(m.group(2))
        ]

    def emit(self) -> bytes:
        """Serialize the worksheet part."""
        rows = "".join(row.emit() for row in self.rows if row.cells)
        return (self.pre + rows + self.post).encode()

    def cols(self) -> list[list[tuple[str, str]]]:
        """Return the ``<col>`` attribute lists inside ``<cols>``."""
        m = _COLS_RE.search(self.pre)
        if m is None:
            return []
        return [parse_attrs(col.group(1)) for col in _COL_RE.finditer(m.group(0))]

    def set_cols(self, cols: list[list[tuple[str, str]]]) -> None:
        """Replace the ``<cols>`` element content."""
        body = "".join(f"<col{emit_attrs(attrs)}></col>" for attrs in cols)
        replacement = f"<cols>{body}</cols>" if cols else ""
        self.pre = _COLS_RE.sub(lambda _: replacement, self.pre, count=1)


def _parse_row(row_match: re.Match[str]) -> Row:
    attrs = parse_attrs(row_match.group(1))
    num = int(dict(attrs)["r"])
    cells = []
    for cell_match in _CELL_RE.finditer(row_match.group(2)):
        cell_attrs = parse_attrs(cell_match.group(1))
        address = dict(cell_attrs)["r"]
        col_letters, _ = split_address(address)
        cells.append(Cell(col_to_index(col_letters), cell_attrs, cell_match.group(2)))
    return Row(num, attrs, cells)
