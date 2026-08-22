from __future__ import annotations

from _patched_xlsxedit import read_size


class Image:
    def __init__(self, data: bytes) -> None:
        self._data = data
        self._width, self._height = read_size(data)

    def __repr__(self) -> str:
        return f"<Image {self._width}x{self._height}>"

    @property
    def data(self) -> bytes:
        return self._data

    @property
    def width(self) -> int:
        return self._width

    @property
    def height(self) -> int:
        return self._height
