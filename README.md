# xlsxfill

Generate an xlsx by merging a template xlsx with data.

The output is built through xlsxedit, so the elements of the original template
are not destroyed along the way.

## Install

```bash
pip install xlsxfill
```

## Usage

`template.xlsx`, whose first sheet looks like this:

|   | A |
|---|---|
| 1 | `${customer.name}` |

```python
from xlsxfill import fill

fill(
    "template.xlsx",
    {"customer": {"name": "Acme Inc."}},
    "output.xlsx",
)
```

`output.xlsx`:

|   | A |
|---|---|
| 1 | Acme Inc. |

The placeholder syntax is documented in
[Syntax](https://hotaritobu.github.io/xlsxfill/latest/syntax/); real workbooks
run through the library are shown in
[Samples](https://hotaritobu.github.io/xlsxfill/latest/samples/).

## Acknowledgments

[xlsxedit](https://github.com/jonas-kupferschmid/xlsxedit)
