"""VBA error handling support."""

from __future__ import annotations

import traceback
from typing import Any, Callable, Optional


class VBARuntimeError(Exception):
    """Exception raised by VBA Err.Raise."""

    def __init__(self, number: int, description: str = "") -> None:
        self.number = number
        self.description = description
        super().__init__(f"VBA Error {number}: {description}")


class VBAError:
    """Emulates the VBA Err object.

    Usage:
        err = VBAError()
        err.raise_error(1004, "Application-defined or object-defined error")
    """

    def __init__(self) -> None:
        self.number: int = 0
        self.description: str = ""
        self.source: str = ""

    def clear(self) -> None:
        """Err.Clear -- reset all properties."""
        self.number = 0
        self.description = ""
        self.source = ""

    def raise_error(
        self,
        number: int,
        description: str = "",
        source: str = "",
    ) -> None:
        """Err.Raise -- set error properties and raise VBARuntimeError."""
        self.number = number
        self.description = description
        self.source = source
        raise VBARuntimeError(number, description)


# Standard VBA error number mappings
_PYTHON_TO_VBA_ERROR: dict[type, tuple[int, str]] = {
    ZeroDivisionError: (11, "Division by zero"),
    TypeError: (13, "Type mismatch"),
    OverflowError: (6, "Overflow"),
    FileNotFoundError: (53, "File not found"),
    PermissionError: (70, "Permission denied"),
    IndexError: (9, "Subscript out of range"),
    KeyError: (9, "Subscript out of range"),
    ValueError: (13, "Type mismatch"),
    OSError: (75, "Path/File access error"),
}


def _map_exception(exc: Exception) -> tuple[int, str]:
    """Map a Python exception to a VBA error number and description."""
    if isinstance(exc, VBARuntimeError):
        return exc.number, exc.description

    for exc_type, (number, desc) in _PYTHON_TO_VBA_ERROR.items():
        if isinstance(exc, exc_type):
            return number, f"{desc}: {exc}"

    # Generic runtime error
    return 1004, f"Application-defined or object-defined error: {exc}"


class OnErrorResumeNext:
    """Context manager for On Error Resume Next blocks.

    Catches exceptions from individual statements and records them
    in the VBAError object instead of propagating.

    Usage:
        err = VBAError()
        handler = OnErrorResumeNext(err)

        # Wrap each potentially-failing statement:
        handler.try_exec(lambda: risky_operation())
        handler.try_exec(lambda: another_operation())

        # Or use as context manager for a block where each
        # statement is wrapped via try_exec:
        with OnErrorResumeNext(err) as on_error:
            on_error.try_exec(lambda: something())
            if err.number != 0:
                # handle error
                err.clear()
    """

    def __init__(self, err: VBAError) -> None:
        self._err = err

    def __enter__(self) -> "OnErrorResumeNext":
        self._err.clear()
        return self

    def __exit__(
        self,
        exc_type: Optional[type],
        exc_val: Optional[BaseException],
        exc_tb: Any,
    ) -> bool:
        if exc_val is not None and isinstance(exc_val, Exception):
            number, description = _map_exception(exc_val)
            self._err.number = number
            self._err.description = description
            return True  # Suppress the exception
        return False

    def try_exec(self, func: Callable[[], Any]) -> Any:
        """Execute a single statement, catching errors into err.

        Returns the result of func() on success, or None on error.
        """
        try:
            self._err.clear()
            return func()
        except Exception as exc:
            number, description = _map_exception(exc)
            self._err.number = number
            self._err.description = description
            return None
