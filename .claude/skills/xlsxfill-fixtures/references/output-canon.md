# Output canon

The exact XML forms the expected files use. The expected fixtures are the only place this spec is recorded — every new fixture must follow these forms so the suite stays self-consistent.

## Strings (sharedStrings.xml)

- Strings produced by substitution are appended at the end in first-use order (sheet order, then cell order)
- Existing entries that lose all references are **kept** — removing one would renumber every later reference
- `count` = actual reference total across all sheets; `uniqueCount` = actual entry total
- Strings with leading/trailing whitespace or embedded newlines get `xml:space="preserve"`; newlines are literal newline characters

## Cell values

- **Numbers** (including date/time/datetime serials): Python `repr` shortest round-trip form — `46251`, `0.3958333333333333`, `46251.395833333336`, `36.0`. Serials: `(date - 1899-12-30).days`
- **Booleans**: `<c t="b"><v>1</v></c>` / `<v>0</v>`
- **null → blank cell**: a cell with no style loses its `<c>` element entirely; a row whose cells all disappear loses its `<row>` element
- **Formula cells**: keep `t="str"` and no cached value; copies get their references rebased with Excel-compatible insert/delete/copy semantics

## Sheet parts

- **dimension**: left untouched (the `A1` excelize wrote stays)
- **cols**: adjacent identical entries collapse via `min`/`max`; a deleted column loses its width entry
- **Hyperlinks**: one rel per cell (never shared, even for the same URL); sheet rels use alphabetical attribute order `Id, Target, TargetMode, Type`

## Drawings

- **Images**: oneCellAnchor. Cell size in EMU: column width → px = round(w×7)+5, row height → EMU = pt×12700. contain = centered; cover = cropped via `<a:srcRect>` (units of 1/100000); fill = the whole cell. Add the `png` Default and the drawing Override to Content_Types
- **Shape copies**: duplicates with anchors shifted to each iteration position, appended right after the original; `cNvPr` id/name continue sequentially from the existing maximum

## Book-level parts

- **Sheet duplication**: the new part is `sheetN+1.xml`, its rId is appended at the end of the rels, sheetId = max+1. Copies drop `tabSelected` (only one tab may stay selected). Content_Types Overrides are appended at the end of Types. Sheet deletion is the reverse: remove the part, rel, Override, and entry — do not renumber
- **Tables**: when the range changes, update `ref` and the inner autoFilter `ref` together; when columns are added, update tableColumns with the actual header values
- **Non-cell containers** (headers/footers, shape text, charts, comments, tooltips, data validation, docProps, sheet names): their text lives inline in each part, not in sharedStrings — edit it there
