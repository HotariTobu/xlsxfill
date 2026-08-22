import base64
import json
from datetime import date, datetime, time
from pathlib import Path

from xlsxfill import Value

DATA_GOLDEN = Path(__file__).parent / "data_golden"
DATA_RAISE = Path(__file__).parent / "data_raise"


def decode(node: object) -> Value:
    if isinstance(node, dict):
        if len(node) == 1:
            ((key, raw),) = node.items()
            match key:
                case "$date":
                    return date.fromisoformat(raw)
                case "$time":
                    return time.fromisoformat(raw)
                case "$datetime":
                    return datetime.fromisoformat(raw)
                case "$bytes":
                    return base64.b64decode(raw)
        return {key: decode(value) for key, value in node.items()}
    if isinstance(node, list):
        return [decode(value) for value in node]
    if node is None or isinstance(node, str | int | float | bool):
        return node
    raise TypeError(f"unexpected node type: {type(node)!r}")


def load_input(path: Path) -> Value:
    return decode(json.loads(path.read_text(encoding="utf-8")))
