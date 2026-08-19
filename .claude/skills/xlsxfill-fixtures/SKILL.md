---
name: xlsxfill-fixtures
description: Create or update xlsxfill test fixtures under tests/data_golden/ and tests/data_raise/ (template.xlsx / input.json / expected.xlsx triples). Use this skill whenever the user asks to add, fix, regenerate, or verify a test fixture, a template workbook, a golden/expected xlsx, or a raise case — even if they only name a book (values, expansion, errors, …) or a single case. Templates are generated with excelize (Go); expected files are authored by hand-editing the template's XML parts.
---

# xlsxfill fixture authoring

## File layout

```
tests/data_golden/<book>.template.xlsx / <book>.input.json / <book>.expected.xlsx
tests/data_raise/<case>.template.xlsx / <case>.input.json / <case>.error.json
```

- data_golden: the generated output must match expected.xlsx **exactly**. Embedded-error cases (`#SYNTAX!` / `#DATA!`) go here too
- data_raise: raising cases. `error.json` is `{"type": "DataError", "message": "<exact>"}`
- Anchor every sheet at B2 (row 1 and column A stay empty, column A width 3). Give structural cases one sheet per case (sheet name = case name)

## Workflow

1. Generate the template with a Go program (excelize)
2. Canonize the template
3. Write input.json
4. Author the expected by editing the template's extracted XML parts
5. Verify: full dump + part diff
6. Only then consider the fixture done

## 1. Generate the template (excelize / Go)

Write one generator function per book with excelize v2.11 and run it. Use excelize rather than authoring template XML by hand because it emits single-line XML with a stable element/attribute order — that is what makes step 4's targeted `Edit` calls and the byte-identity invariant workable.

Helper pattern that keeps generators short:

```go
import (
    _ "image/png" // AddPictureFromBytes needs the caller to register the decoder
    "github.com/xuri/excelize/v2"
)

type cellDef struct{ ref, val string }

// One case sheet, anchored at B2: col A width 3, listed cols width 12.
func bandSheet(f *excelize.File, name, wideCols string, cells []cellDef) { ... }
```

Gotchas learned the hard way:

- Pass formulas to `SetCellFormula` **without** the leading `=` (`C2*D2`); excelize writes `<c t="str"><f>C2*D2</f></c>`
- Chart titles / axis labels take `excelize.ChartTitle{Paragraph: []excelize.RichTextRun{...}}`
- Tooltips: `SetCellHyperLink(..., excelize.HyperlinkOpts{Tooltip: &tip})`; data validation: `NewDataValidation` + `SetInput`/`SetError`
- Print areas are defined names (`_xlnm.Print_Area`, Scope = sheet name)
- A shape's "don't move or size with cells" is `Shape.Format.Positioning = "absolute"`
- Dates can be written as serial values into a formatted cell (`SetCellValue(sheet, "D6", 46251)`)
- Make the first case sheet by renaming the one `NewFile` already gives you (`SetSheetName("Sheet1", …)`). Deleting it instead leaves the parts numbered from `sheet2.xml`

## 2. Canonize the template

```bash
uv run python .claude/skills/xlsxfill-fixtures/scripts/canonize.py tests/data_golden/<book>.template.xlsx
```

Run it the moment the template is generated, before writing anything against it. What excelize writes is a perfectly good xlsx and not the one the suite is written against, in three ways that would otherwise surface as differences between a template and its expected — differences the substitution never made:

- **The package.** `[Content_Types].xml` and the relationship parts come out in excelize's arrangement; xlsxedit rebuilds both when it saves
- **The serialization.** Parts xlsxedit keeps as blobs — charts, drawings, tables, comments — stay as excelize wrote them, while the parts it parses come back in lxml's form
- **`sst/@count`.** excelize writes the number of entries; ECMA-376 §18.4.9 asks for the number of references

Templates only. It refuses anything else, because a fill *appends* what it adds to `[Content_Types].xml` while a save *sorts* the whole list — canonizing an expected would move it away from the very run it records.

Every step checks its own work and refuses rather than changing what a workbook holds, so a run that prints `already canonical` for each of the existing templates is the standing proof that it and the suite still agree.

## 3. input.json

- One file per book (merged input). Keys must not collide between cases; sharing a key is fine when the value is identical
- Type tags (the runner converts them): `{"$date": "2026-08-17"}` / `{"$time": "09:30"}` / `{"$datetime": "2026-08-17T09:30"}` / `{"$bytes": "<base64>"}`

## 4. Author the expected (hand-edit XML)

1. `unzip` the template.xlsx into a working directory
2. Edit only the parts substitution touches, **using the Edit tool** — command-driven edits (sed, python, heredocs) hide the diff from the user and are forbidden
3. `zip -q -X` the parts back **in the template's part order**. Order and `-X` matter because the fixtures' value comes from being stable, reviewable artifacts

**Invariant: parts untouched by substitution stay byte-identical to the template.** This is the property the whole fixture suite exists to pin down — if an unrelated part changed, the expected is wrong, not the invariant.

Before editing any part, read `references/output-canon.md`. It defines the exact XML forms the expected files use (sharedStrings bookkeeping, number formatting, part-level conventions). Deviating from it makes fixtures inconsistent with each other, and the expected files are the only place this spec lives.

**Example — substituting one cell** (`${customer.name}` → `Alice`):

sheet1.xml, before → after:
```xml
<c r="C2" t="s"><v>1</v></c>
<c r="C2" t="s"><v>34</v></c>
```

sharedStrings.xml: append `<si><t>Alice</t></si>` at the end (index 34), keep entry 1, and set `count` to the actual reference total and `uniqueCount` to the actual entry total.

## 5. Verify (every book, every time)

1. **Full dump** — `scripts/dumpbook/` is a bundled Go probe. First use: `cd scripts/dumpbook && go mod tidy`. Then:
   ```bash
   go run . path/to/book.expected.xlsx
   ```
   It prints one sorted fact per line, in two layers.

   **Resolved** — what a reader of xlsx understands the workbook to hold: cells (type, raw value, displayed value, formula, expanded style), merges, hyperlinks, data validations, conditional formats, tables, defined names, pictures (sha256 of the image bytes, alt text, anchor, offsets, positioning), comments, sheet visibility, sheet view, sheet properties, page layout and margins, row/column geometry.

   **Structural** — `part` lines flatten every remaining part node by node, with no model of what it means, plus a `blob` hash for each binary part. This is the only coverage charts, VML, themes, `docProps`, `[Content_Types].xml` and the `.rels` get: excelize has no reader for any of them, so a fixture whose chart title or axis label is wrong would otherwise pass unnoticed.

   Read both against the case expectations line by line — this catches wrong indices and geometry that a visual skim misses.

   The output is canonical, not the file's bytes, so it also answers "are these two books the same workbook?":
   ```bash
   diff -u <(go run . a.xlsx) <(go run . b.xlsx)
   ```
   How the XML is written — the declaration's quoting, `<x/>` against `<x></x>`, which of two prefixes bound to one namespace a tag uses — never reaches the dump. Element *order* does reach the `part` lines, so a set-like part such as `[Content_Types].xml` shows a difference when its entries are merely rearranged. That is wanted here — fixtures are pinned to their exact form — but it is a difference in form, not in what Excel would show.

   Every cell appears twice, and the pair is the point:

   - `cell` — what excelize resolves: the displayed value, the shared string looked up, the style expanded
   - `raw-cell` — what the sheet actually stores

   They differ where excelize interprets. Asked about a cell inside a merged block it answers with the block's value, so a cell storing nothing and a cell storing the same text look alike through `cell` alone; `raw-cell` tells them apart. `raw-cell` is emitted for every sheet, always.

   What it will not do:

   - **Skip anything silently.** A sheet excelize refuses to address — a tab name over 31 characters, which is what an embedded `#SYNTAX!` in a sheet name produces — is read straight out of the package instead and announced with a `book fallback` line
   - **Miss a cell.** The cells to report come from the sheet's own XML, not from `GetRows`, which drops the empty cells at the end of a row — and a cell holding nothing but a style is exactly that: empty, and still data
   - **Swallow an error.** A failed call prints a `*-error` fact, so a missing line always means missing data rather than a call that quietly returned nothing
   - **Print a pointer.** Structs go through JSON. `%+v` prints addresses that change between runs and read as differences

   Row and column geometry also comes from the XML, never through excelize: `GetRowVisible` reports a row that does not exist as hidden, and indexes rows by position rather than by their `r` attribute.
2. **Part diff** — unzip template and expected separately and `diff -rq`. The differences must be exactly: the parts substitution touches, plus deliberately added/removed parts. Anything else is a bug in the expected
3. **Recount** `count` / `uniqueCount` after editing and confirm they match what you wrote
