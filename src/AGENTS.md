# xlsxfill src

Dependency direction: `xlsxedit` (PyPI) ← `_patched_xlsxedit` ← `_excel` ← `xlsxfill`.
Each layer imports only the one below it.

- **xlsxedit** — the package on PyPI that reads and writes xlsx. Not ours.

- **_patched_xlsxedit** — what xlsxedit does not do yet, written the way
  xlsxedit would do it. Every name takes and returns xlsxedit's own things,
  and each is a candidate upstream request, so each can be deleted the day
  upstream has it.

- **_excel** — what a person can do to a workbook in Excel. Nothing here
  exposes how an xlsx file is put together.

- **xlsxfill** — the template language. It says what to do in Excel's terms
  and never learns anything about the file format.
