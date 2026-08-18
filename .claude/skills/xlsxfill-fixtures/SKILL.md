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
2. Write input.json
3. Author the expected by editing the template's extracted XML parts
4. Verify: full dump + part diff
5. Only then consider the fixture done

## 1. Generate the template (excelize / Go)

Write one generator function per book with excelize v2.11 and run it. Use excelize rather than authoring template XML by hand because it emits single-line XML with a stable element/attribute order — that is what makes step 3's targeted `Edit` calls and the byte-identity invariant workable.

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

## 2. input.json

- One file per book (merged input). Keys must not collide between cases; sharing a key is fine when the value is identical
- Type tags (the runner converts them): `{"$date": "2026-08-17"}` / `{"$time": "09:30"}` / `{"$datetime": "2026-08-17T09:30"}` / `{"$bytes": "<base64>"}`

## 3. Author the expected (hand-edit XML)

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

## 4. Verify (every book, every time)

1. **Full dump** — `scripts/dumpbook/` is a bundled Go probe. First use: `cd scripts/dumpbook && go mod tidy`. Then:
   ```bash
   go run . path/to/book.expected.xlsx
   ```
   It prints every cell's value/type, formulas, column widths, hyperlinks, and picture counts. Check the output against the case expectations line by line — this catches wrong indices and geometry that a visual skim misses. It skips sheets excelize cannot read (e.g. tab names over 31 characters); verify those tabs directly in workbook.xml
2. **Part diff** — unzip template and expected separately and `diff -rq`. The differences must be exactly: the parts substitution touches, plus deliberately added/removed parts. Anything else is a bug in the expected
3. **Recount** `count` / `uniqueCount` after editing and confirm they match what you wrote
