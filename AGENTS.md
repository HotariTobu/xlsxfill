# xlsxfill

## Language

All documentation, code comments, commit messages, and issues must be written in **English**.

## Layout

- `/src`
- `/tests`
- `/docs` — the documentation site.

## Code comments

Don't write comments. The reader can read the implementation, and a comment is
believed without it being read — the next implementation is then written on what
the comment said rather than on what the code does.

The exception is the public API: its docstrings are what a library user consults
instead of the source, and they are rendered into the reference site. Keep them,
and prefer forms a tool can check — cross-references fail the strict docs build
when the symbol they name is gone; prose does not.

## Lint suppressions

Don't add or expand lint suppressions without explicit user approval. This covers:

- `[tool.ruff.lint] ignore` / `extend-ignore`
- `[tool.ruff.lint.per-file-ignores]` (new entries, expanded rule lists, new file globs)
- `# noqa: <code>` inline comments
- `# type: ignore` / `# ty: ignore` inline comments
- `[tool.ty.rules]` severity downgrades

Resolving a lint violation by suppression vs. by fixing code is a user judgment call. Surface the violation and wait for direction.

## Commands

@Makefile
