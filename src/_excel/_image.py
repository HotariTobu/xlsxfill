"""An image, before it is placed in a workbook."""

from __future__ import annotations

from _patched_xlsxedit import read_size


class Image:
    """An image file, read far enough to know how big it is.

    Excel reads an image when you insert it — that is how it can offer
    to keep the aspect ratio and how it knows what "reset to original
    size" means. This does the same. The bytes are stored in the
    workbook exactly as they arrived; only the header is looked at.
    """

    def __init__(self, data: bytes) -> None:
        """Read ``data``.

        Args:
            data: The image file's bytes.

        Raises:
            ValueError: The bytes are not an image this can read.
        """
        self._data = data
        self._width, self._height = read_size(data)

    def __repr__(self) -> str:
        """Show its own size."""
        return f"<Image {self._width}x{self._height}>"

    @property
    def data(self) -> bytes:
        """The bytes it was made from."""
        return self._data

    @property
    def width(self) -> int:
        """Its own width in pixels, before any placing."""
        return self._width

    @property
    def height(self) -> int:
        """Its own height in pixels, before any placing."""
        return self._height
