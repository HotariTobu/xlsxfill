# xlsxfill tests

The specification is defined here, exhaustively, as data. Extending it means
editing the data, not the Python files.

- **data_golden / test_golden.py** — a template, the data merged into it, and
  the workbook the run must write. Every member of the package is compared byte
  for byte, so a book states not only what was filled in but that nothing else
  moved.

- **data_raise / test_raise.py** — a template, the data, and the exception
  raised in place of a workbook. Nothing is written, so the message is matched
  in full.
