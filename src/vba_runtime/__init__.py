"""VBA Runtime Library for Python -- provides VBA-compatible functions and types.

Usage:
    from vba_runtime import vba_left, vba_mid, vba_cint, VBAWorkbook, EMPTY, NULL
"""

# String functions
from vba_runtime.strings import (
    vba_asc,
    vba_chr,
    vba_instr,
    vba_join,
    vba_lcase,
    vba_left,
    vba_len,
    vba_ltrim,
    vba_mid,
    vba_replace,
    vba_right,
    vba_rtrim,
    vba_space,
    vba_split,
    vba_string_func,
    vba_strreverse,
    vba_trim,
    vba_ucase,
)

# Conversion functions
from vba_runtime.conversion import (
    vba_cbool,
    vba_cdate,
    vba_cdbl,
    vba_cint,
    vba_clng,
    vba_cstr,
    vba_fix,
    vba_int_func,
    vba_isdate,
    vba_isempty,
    vba_isnull,
    vba_isnumeric,
    vba_typename,
    vba_val,
)

# Math functions
from vba_runtime.math_funcs import (
    vba_abs,
    vba_round,
    vba_rnd,
    vba_sgn,
    vba_sqr,
)

# Error handling
from vba_runtime.errors import (
    OnErrorResumeNext,
    VBAError,
    VBARuntimeError,
)

# Types
from vba_runtime.types import (
    EMPTY,
    NULL,
    VBAArray,
    VBACollection,
)

# Excel support (lazy -- only import when used to avoid openpyxl dependency)
def __getattr__(name: str):
    """Lazy import for Excel classes to avoid requiring openpyxl at import time."""
    _excel_names = {"VBAWorkbook", "VBAWorksheet", "VBARange"}
    if name in _excel_names:
        from vba_runtime import excel
        return getattr(excel, name)
    raise AttributeError(f"module 'vba_runtime' has no attribute {name!r}")


__all__ = [
    # Strings
    "vba_left", "vba_right", "vba_mid", "vba_instr", "vba_replace",
    "vba_len", "vba_trim", "vba_ltrim", "vba_rtrim",
    "vba_ucase", "vba_lcase", "vba_space", "vba_string_func",
    "vba_asc", "vba_chr", "vba_split", "vba_join", "vba_strreverse",
    # Conversion
    "vba_cstr", "vba_cint", "vba_clng", "vba_cdbl", "vba_cbool", "vba_cdate",
    "vba_val", "vba_int_func", "vba_fix",
    "vba_isnumeric", "vba_isdate", "vba_isempty", "vba_isnull", "vba_typename",
    # Math
    "vba_abs", "vba_sgn", "vba_sqr", "vba_rnd", "vba_round",
    # Errors
    "VBAError", "VBARuntimeError", "OnErrorResumeNext",
    # Types
    "EMPTY", "NULL", "VBACollection", "VBAArray",
    # Excel (lazy)
    "VBAWorkbook", "VBAWorksheet", "VBARange",
]
