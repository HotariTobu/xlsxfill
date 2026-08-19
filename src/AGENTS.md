# xlsxedit / _patched_xlsxedit / xlsxfill

Dependency direction: `xlsxedit` (PyPI) ← `_patched_xlsxedit` ← `xlsxfill`.

- **xlsxedit** re-serializes XML parts through lxml on save, which does not
  round-trip authored XML byte-for-byte. xlsxfill's output contract is
  byte-exact, so xlsxedit's editing APIs cannot be used directly.
- **_patched_xlsxedit** patches xlsxedit from the outside: byte-preserving
  counterparts of its APIs (`Package`, `Relationships`, `ContentTypes`,
  `Workbook`, `SharedStringTable`, `row_shift` generalized to arbitrary
  mappings). It imports xlsxedit (address utilities, `PackURI`, `CT`/`RT`
  constants) and never imports xlsxfill.
- **xlsxfill** holds the template language (syntax, data resolution, band
  expansion, orchestration) and imports xlsxedit only through
  `_patched_xlsxedit`, never directly.

Boundary rule: knowledge of the xlsx format belongs in `_patched_xlsxedit`;
knowledge of the template language belongs in `xlsxfill`.

Each patched API is a preview of xlsxedit's own direction: it mirrors an
existing upstream API, or one merged upstream but unreleased (e.g.
`copy_worksheet`, the `insert_rows` reference following), so its use case
can be requested upstream as-is. Once upstream ships a feature, the patch
is dropped in favor of the upstream API.
