.PHONY: autofix inspection typecheck preflight test docs\:serve docs\:build

autofix:
	uv run ruff check --fix
	uv run ruff format

inspection:
	uv run ruff check
	uv run ruff format --check

typecheck:
	uv run ty check

preflight: inspection typecheck

test:
	uv run pytest

docs\:serve:
	uv run --group docs mkdocs serve

docs\:build:
	uv run --group docs mkdocs build --strict
