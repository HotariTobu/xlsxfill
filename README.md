# xlsxfill

Generate an xlsx by merging a template xlsx with data.

The template is an ordinary Excel workbook. Placeholders are plain strings typed
through the Excel UI, so the people who own the layout can keep owning it —
formatting, formulas, charts, images and print settings all survive untouched.

- **Placeholders, not code.** `${customer.name}` for a value, `#{rItem}` to mark
  a repeating band.
- **The data's type decides everything.** Type specifiers such as `:num` or
  `:date` are assertions, not conversions.
- **Excel keeps its own job.** Arithmetic stays in formulas, formatting stays in
  number formats, conditionals stay in `=IF()`.

## Install

```bash
pip install xlsxfill
```

## Usage

Given a template whose first sheet looks like this:

|   | A | B | C | D |
|---|---|---|---|---|
| 1 | `${customer.name}` | | | |
| 2 | `#{r}${items#r.name}` | `${items#r.qty:num}` | `${items#r.price:num}` | `=B2*C2` |
| 3 | `#{r+1}` | | | |

```python
from xlsxfill import fill

problems = fill(
    "template.xlsx",
    {
        "customer": {"name": "Acme Inc."},
        "items": [
            {"name": "Widget", "qty": 3, "price": 1200},
            {"name": "Gadget", "qty": 1, "price": 4800},
        ],
    },
    "output.xlsx",
)
```

Row 2 is repeated once per item, rows below it shift down, and `=B2*C2` is
rebased for every copy.

`template` and `output` accept a path or a binary file object, so a workbook can
be built entirely in memory:

```python
from io import BytesIO

buffer = BytesIO()
fill("template.xlsx", data, buffer)
```

## Errors

Two kinds of failure are distinguished.

**Unusable input** raises. The output is never written.

```python
from xlsxfill import DataError

try:
    fill("template.xlsx", data, "output.xlsx")
except DataError:
    ...
```

**A bad placeholder or a bad value** does not stop the run. The offending cell
receives `#SYNTAX! ...` or `#DATA! ...`, and `fill` returns a list of `Problem`
describing each one:

```python
problems = fill("template.xlsx", data, "output.xlsx")
for problem in problems:
    print(problem.message)
```

This way one broken field never costs you the whole report.

## Documentation

- [Placeholder syntax](https://hotaritobu.github.io/xlsxfill/syntax/)
- [API reference](https://hotaritobu.github.io/xlsxfill/api/)

## Supported environments

- Python ≥ 3.12
- Linux, macOS, Windows

## License

[CC0 1.0 Universal](LICENSE) — public domain dedication.
