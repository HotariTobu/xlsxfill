# xlsxfill

[![PyPI](https://img.shields.io/pypi/v/xlsxfill)](https://pypi.org/project/xlsxfill/)
[![Python](https://img.shields.io/pypi/pyversions/xlsxfill)](https://pypi.org/project/xlsxfill/)
[![Preflight & Test](https://github.com/HotariTobu/xlsxfill/actions/workflows/preflight-and-test.yml/badge.svg)](https://github.com/HotariTobu/xlsxfill/actions/workflows/preflight-and-test.yml)
[![License](https://img.shields.io/pypi/l/xlsxfill)](LICENSE)

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

For language models, the documentation is indexed in
[llms.txt](https://hotaritobu.github.io/xlsxfill/latest/llms.txt), and the pages
it lists are concatenated into
[llms-full.txt](https://hotaritobu.github.io/xlsxfill/latest/llms-full.txt).

## Acknowledgments

[xlsxedit](https://github.com/jonas-kupferschmid/xlsxedit)
