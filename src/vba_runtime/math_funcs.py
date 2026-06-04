"""VBA math functions."""

from __future__ import annotations

import decimal
import math
import random
from typing import Any, Optional


def vba_abs(n: Any) -> float:
    """Abs(number) -- absolute value."""
    if n is None:
        return 0
    return abs(float(n))


def vba_sgn(n: Any) -> int:
    """Sgn(number) -- returns -1, 0, or 1."""
    if n is None:
        return 0
    val = float(n)
    if val > 0:
        return 1
    elif val < 0:
        return -1
    return 0


def vba_sqr(n: Any) -> float:
    """Sqr(number) -- square root (VBA's Sqr = sqrt)."""
    if n is None:
        raise ValueError("Invalid use of Null")
    val = float(n)
    if val < 0:
        raise ValueError("Invalid procedure call or argument: Sqr of negative number")
    return math.sqrt(val)


# Module-level RNG state for VBA-compatible Rnd behavior
_rng = random.Random()


def vba_rnd(seed: Optional[float] = None) -> float:
    """Rnd([seed]) -- VBA-compatible random number.

    - seed < 0: Reinitialize with seed, return first value
    - seed = 0: Return last generated number (approximated: re-seed same state)
    - seed > 0 or omitted: Next random number in sequence
    Returns float in [0.0, 1.0)
    """
    global _rng

    if seed is not None:
        if seed < 0:
            _rng = random.Random(seed)
            return _rng.random()
        elif seed == 0:
            # VBA returns the most recently generated number
            # We approximate by using a fixed state peek
            state = _rng.getstate()
            val = _rng.random()
            _rng.setstate(state)
            return val

    return _rng.random()


def vba_round(n: Any, decimals: int = 0) -> float:
    """Round(number, [decimals]) -- VBA uses Banker's rounding (round half to even)."""
    if n is None:
        return 0
    d = decimal.Decimal(str(float(n)))
    rounded = d.quantize(
        decimal.Decimal(10) ** -decimals,
        rounding=decimal.ROUND_HALF_EVEN,
    )
    result = float(rounded)
    # Return int-like value when decimals=0 for cleaner output
    if decimals == 0:
        return int(result)
    return result
