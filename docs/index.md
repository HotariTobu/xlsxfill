# xlsxfill

Generate an xlsx by merging a template xlsx with data.

## Install

```bash
pip install xlsxfill
```

Python 3.12 or later.

## How the output is produced

xlsxfill does not rebuild the workbook from a model of its own. It opens the
template with [xlsxedit](https://github.com/jonas-kupferschmid/xlsxedit), edits
the parts the placeholders touch, and writes the package back out. The elements
of the original template are not destroyed along the way.

## Filling a template

Given `template.xlsx`, whose first sheet looks like this:

|   | A | B | C | D |
|---|---|---|---|---|
| 1 | `${customer.name}` | | | |
| 2 | `#{r}${items#r.name}` | `${items#r.qty:num}` | `${items#r.price:num}` | `=B2*C2` |
| 3 | `#{r+1}` | | | |

```python
from xlsxfill import fill

fill(
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

`output.xlsx`:

|   | A | B | C | D |
|---|---|---|---|---|
| 1 | Acme Inc. | | | |
| 2 | Widget | 3 | 1200 | `=B2*C2` |
| 3 | Gadget | 1 | 4800 | `=B3*C3` |

Row 2 is duplicated once per item, the row holding `#{r+1}` is removed, and the
formula is rebased for each copy.

The notation itself is documented in [Syntax](syntax.md).

## Errors

Failures split in two. `DataError` is raised when the input data as a whole is
unusable, and nothing is written. A bad placeholder or a bad value does not stop
the run: the offending construct is replaced with a message, and `fill` returns
one `Problem` per message.

Which constructs and values fall on which side is listed in
[Syntax](syntax.md#errors); the types are described in [API](api.md).
