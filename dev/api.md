# API

## BaseProblem

```python
class BaseProblem
```

An error reported in the output while processing continues.

- `kind: Literal['syntax', 'data']`
- `construct: str`
- `reason: str`
- `message: str` — The string embedded in the output. ``#SYNTAX! <construct>: <reason>`` or ``#DATA! <construct>: <reason>``.

## BookProblem

```python
class BookProblem(BaseProblem)
```

A problem in a workbook-level string container.

- `part: Literal['doc_props']`

## CellProblem

```python
class CellProblem(BaseProblem)
```

A problem in a cell or in a string container attached to a cell.

- `sheet: str`
- `cell: str`
- `part: Literal['cell', 'comment', 'tooltip', 'validation']`

## DataError

```python
class DataError(XlsxfillError)
```

The input data as a whole is unusable; nothing is written.

## Problem

```python
Problem = BookProblem | SheetProblem | CellProblem
```

A problem reported in the output.

One of [BookProblem](https://hotaritobu.github.io/xlsxfill/dev/api.md#xlsxfill.BookProblem),
[SheetProblem](https://hotaritobu.github.io/xlsxfill/dev/api.md#xlsxfill.SheetProblem) or [CellProblem](https://hotaritobu.github.io/xlsxfill/dev/api.md#xlsxfill.CellProblem).

## SheetProblem

```python
class SheetProblem(BaseProblem)
```

A problem in a sheet-level string container.

- `sheet: str`
- `part: Literal['sheet_name', 'header_footer', 'shape', 'chart']`

## Value

```python
Value = str | int | float | bool | date | time | datetime | bytes | list[Value] | dict[str, Value] | None
```

A value in the input data.

## XlsxfillError

```python
class XlsxfillError(Exception)
```

Base class for all exceptions raised by xlsxfill.

## fill

```python
fill(template: str | Path | BinaryIO, data: Mapping[str, Value], output: str | Path | BinaryIO) -> list[Problem]
```

Merge a template xlsx with data and write the result.

Args:
    template: The template xlsx.
    data: The data to merge into the template.
    output: Where the resulting xlsx is written.

Returns:
    One problem per [message](https://hotaritobu.github.io/xlsxfill/dev/api.md#xlsxfill.BaseProblem.message) embedded in the
    output.

Raises:
    DataError: ``data`` is unusable as a whole.
