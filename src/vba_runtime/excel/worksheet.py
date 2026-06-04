"""VBA Worksheet wrapper around openpyxl."""

from __future__ import annotations

from typing import Any, Union

from vba_runtime.excel.range_ import VBARange


class VBAWorksheet:
    """Thin wrapper around openpyxl.Worksheet providing VBA-like interface.

    Usage:
        ws.cells(1, 1).value = "Hello"
        val = ws.cells(2, 3).value
        rng = ws.range("A1:B5")
    """

    def __init__(self, ws: Any) -> None:
        self._ws = ws

    @property
    def name(self) -> str:
        """Worksheet name (title)."""
        return self._ws.title

    @name.setter
    def name(self, value: str) -> None:
        self._ws.title = value

    def cells(self, row: int, col: int) -> VBARange:
        """VBA Cells(row, col) -- both 1-based, matches openpyxl."""
        cell = self._ws.cell(row=row, column=col)
        return VBARange(cell)

    def range(self, ref: str) -> VBARange:
        """VBA Range("A1") or Range("A1:B5").

        Returns a VBARange wrapping the cell or range of cells.
        """
        result = self._ws[ref]
        return VBARange(result)

    @property
    def used_range(self) -> str:
        """Returns the dimensions string of the used range (e.g., 'A1:F10')."""
        return self._ws.dimensions

    @property
    def rows(self) -> Any:
        """Iterator over rows in the worksheet."""
        return self._ws.iter_rows()

    @property
    def max_row(self) -> int:
        """Highest row number with data."""
        return self._ws.max_row

    @property
    def max_column(self) -> int:
        """Highest column number with data."""
        return self._ws.max_column
