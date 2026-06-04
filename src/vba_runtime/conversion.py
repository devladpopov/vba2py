"""VBA type conversion functions."""

from __future__ import annotations

import math
import re
from datetime import datetime
from typing import Any, Optional


def vba_cstr(v: Any) -> str:
    """CStr(expression) -- convert to string."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "True" if v else "False"
    return str(v)


def vba_cint(v: Any) -> int:
    """CInt(expression) -- convert to Integer (banker's rounding to int, range -32768..32767)."""
    if v is None or v == "":
        return 0
    if isinstance(v, bool):
        return -1 if v else 0
    if isinstance(v, str):
        v = float(v)
    # VBA CInt uses banker's rounding
    result = int(_bankers_round(float(v), 0))
    if result < -32768 or result > 32767:
        raise OverflowError("Overflow: CInt result out of range (-32768 to 32767)")
    return result


def vba_clng(v: Any) -> int:
    """CLng(expression) -- convert to Long (banker's rounding to int, range -2147483648..2147483647)."""
    if v is None or v == "":
        return 0
    if isinstance(v, bool):
        return -1 if v else 0
    if isinstance(v, str):
        v = float(v)
    result = int(_bankers_round(float(v), 0))
    if result < -2147483648 or result > 2147483647:
        raise OverflowError("Overflow: CLng result out of range")
    return result


def vba_cdbl(v: Any) -> float:
    """CDbl(expression) -- convert to Double."""
    if v is None or v == "":
        return 0.0
    if isinstance(v, bool):
        return -1.0 if v else 0.0
    return float(v)


def vba_cbool(v: Any) -> bool:
    """CBool(expression) -- convert to Boolean. Any nonzero = True."""
    if v is None or v == "":
        return False
    if isinstance(v, str):
        low = v.strip().lower()
        if low == "true":
            return True
        if low == "false":
            return False
        return float(v) != 0
    if isinstance(v, bool):
        return v
    return float(v) != 0


def vba_cdate(v: Any) -> datetime:
    """CDate(expression) -- convert to Date."""
    if v is None:
        raise ValueError("Invalid use of Null")
    if isinstance(v, datetime):
        return v
    if isinstance(v, (int, float)):
        # VBA serial date: day 1 = 1899-12-31 (date serial 1)
        base = datetime(1899, 12, 30)
        from datetime import timedelta

        return base + timedelta(days=float(v))
    if isinstance(v, str):
        # Try common date formats
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(v.strip(), fmt)
            except ValueError:
                continue
        raise ValueError(f"Type mismatch: cannot convert '{v}' to Date")
    raise TypeError(f"Type mismatch: cannot convert {type(v).__name__} to Date")


def vba_val(s: Optional[str]) -> float:
    """Val(string) -- reads numeric prefix from string.

    Skips leading whitespace, reads digits/decimal/sign until non-numeric found.
    """
    if s is None:
        return 0.0
    s = str(s).lstrip()
    if not s:
        return 0.0
    match = re.match(r"[+-]?\d*\.?\d*", s)
    if match and match.group():
        text = match.group()
        if text in ("", "+", "-", "."):
            return 0.0
        return float(text)
    return 0.0


def vba_int_func(n: Any) -> int:
    """Int(number) -- floor toward negative infinity.

    Int(6.3) = 6, Int(-6.3) = -7
    """
    if n is None:
        return 0
    return math.floor(float(n))


def vba_fix(n: Any) -> int:
    """Fix(number) -- truncate toward zero.

    Fix(6.3) = 6, Fix(-6.3) = -6
    """
    if n is None:
        return 0
    return math.trunc(float(n))


def vba_isnumeric(v: Any) -> bool:
    """IsNumeric(expression) -- True if value can be interpreted as a number."""
    if v is None:
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, bool):
        return True
    if isinstance(v, str):
        try:
            float(v)
            return True
        except (ValueError, TypeError):
            return False
    return False


def vba_isdate(v: Any) -> bool:
    """IsDate(expression) -- True if value can be converted to a date."""
    if isinstance(v, datetime):
        return True
    if isinstance(v, str):
        try:
            vba_cdate(v)
            return True
        except (ValueError, TypeError):
            return False
    return False


def vba_isempty(v: Any) -> bool:
    """IsEmpty(expression) -- True if variable is Empty (uninitialized)."""
    # Import here to avoid circular imports
    try:
        from vba_runtime.types import EMPTY

        return v is EMPTY
    except ImportError:
        return v is None


def vba_isnull(v: Any) -> bool:
    """IsNull(expression) -- True if expression is Null."""
    try:
        from vba_runtime.types import NULL

        return v is NULL
    except ImportError:
        return v is None


def vba_typename(v: Any) -> str:
    """TypeName(varname) -- return string describing data type."""
    if v is None:
        return "Nothing"
    try:
        from vba_runtime.types import EMPTY, NULL

        if v is EMPTY:
            return "Empty"
        if v is NULL:
            return "Null"
    except ImportError:
        pass
    if isinstance(v, bool):
        return "Boolean"
    if isinstance(v, int):
        return "Long"
    if isinstance(v, float):
        return "Double"
    if isinstance(v, str):
        return "String"
    if isinstance(v, datetime):
        return "Date"
    return type(v).__name__


def _bankers_round(value: float, decimals: int) -> float:
    """Banker's rounding (round half to even)."""
    import decimal

    d = decimal.Decimal(str(value))
    rounded = d.quantize(
        decimal.Decimal(10) ** -decimals,
        rounding=decimal.ROUND_HALF_EVEN,
    )
    return float(rounded)
