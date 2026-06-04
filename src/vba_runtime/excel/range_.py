"""VBA Range wrapper around openpyxl cell/range."""

from __future__ import annotations

from typing import Any, Optional, Union


class VBARange:
    """Wraps openpyxl cell or range for VBA-like access.

    Handles both single cells and multi-cell ranges transparently.

    Usage:
        # Single cell
        rng = VBARange(ws_cell)
        rng.value = 42
        print(rng.value)

        # Multi-cell range
        rng = VBARange(range_tuple)
        sub = rng.cells(1, 2)  # row 1, col 2 within the range
    """

    def __init__(self, ref: Any) -> None:
        self._ref = ref
        self._is_single = not isinstance(ref, tuple)

    @property
    def value(self) -> Any:
        """Get value of cell (single cell) or tuple of tuples of values (range)."""
        if self._is_single:
            return self._ref.value
        # Multi-cell range: return nested tuple of values
        return tuple(
            tuple(cell.value for cell in row)
            for row in self._ref
        )

    @value.setter
    def value(self, val: Any) -> None:
        """Set value. Only works for single cells."""
        if self._is_single:
            self._ref.value = val
        else:
            raise AttributeError(
                "Cannot set value on a multi-cell range directly. "
                "Use .cells(row, col).value = val instead."
            )

    def cells(self, row: int, col: int) -> "VBARange":
        """Access a cell within this range by 1-based row and column offset.

        For a single cell, row=1, col=1 returns itself.
        For a multi-cell range, accesses the cell at (row, col) within the range.
        """
        if self._is_single:
            if row == 1 and col == 1:
                return self
            raise IndexError("Single cell range: only (1,1) is valid")

        # Multi-cell range: _ref is a tuple of row-tuples
        row_idx = row - 1
        col_idx = col - 1
        if row_idx < 0 or row_idx >= len(self._ref):
            raise IndexError(f"Row {row} out of range")
        row_data = self._ref[row_idx]
        if col_idx < 0 or col_idx >= len(row_data):
            raise IndexError(f"Column {col} out of range")
        return VBARange(row_data[col_idx])

    @property
    def row(self) -> int:
        """Row number of the cell."""
        if self._is_single:
            return self._ref.row
        return self._ref[0][0].row

    @property
    def column(self) -> int:
        """Column number of the cell."""
        if self._is_single:
            return self._ref.column
        return self._ref[0][0].column

    @property
    def rows_count(self) -> int:
        """Number of rows in the range."""
        if self._is_single:
            return 1
        return len(self._ref)

    @property
    def columns_count(self) -> int:
        """Number of columns in the range."""
        if self._is_single:
            return 1
        return len(self._ref[0]) if self._ref else 0

    def __repr__(self) -> str:
        if self._is_single:
            return f"VBARange({self._ref})"
        return f"VBARange(<{self.rows_count}x{self.columns_count} range>)"
