"""VBA string functions with 1-based indexing."""

from __future__ import annotations

from typing import Optional, Union


def _safe_str(s: Optional[str]) -> str:
    """Convert None/Empty to empty string."""
    if s is None:
        return ""
    return str(s)


def vba_left(s: Optional[str], n: int) -> str:
    """Left(s, n) -- return first n characters."""
    s = _safe_str(s)
    if n < 0:
        raise ValueError("Invalid procedure call or argument")
    return s[:n]


def vba_right(s: Optional[str], n: int) -> str:
    """Right(s, n) -- return last n characters."""
    s = _safe_str(s)
    if n < 0:
        raise ValueError("Invalid procedure call or argument")
    if n == 0:
        return ""
    return s[-n:]


def vba_mid(s: Optional[str], start: int, length: Optional[int] = None) -> str:
    """Mid(s, start, [length]) -- 1-based start position."""
    s = _safe_str(s)
    if start < 1:
        raise ValueError("Invalid procedure call or argument")
    idx = start - 1  # convert to 0-based
    if length is None:
        return s[idx:]
    if length < 0:
        raise ValueError("Invalid procedure call or argument")
    return s[idx : idx + length]


def vba_instr(
    start_or_s: Union[int, str, None],
    s_or_find: Optional[str] = None,
    find: Optional[str] = None,
) -> int:
    """InStr([start], string, find) -- returns 1-based position, 0 if not found.

    Two calling conventions:
        vba_instr(string, find)        -- start defaults to 1
        vba_instr(start, string, find)
    """
    if find is not None:
        # Three-arg form: start, string, find
        start = int(start_or_s)  # type: ignore[arg-type]
        s = _safe_str(s_or_find)
        find_str = _safe_str(find)
    elif s_or_find is not None:
        # Two-arg form: string, find
        start = 1
        s = _safe_str(start_or_s)
        find_str = _safe_str(s_or_find)
    else:
        raise TypeError("InStr requires at least 2 arguments")

    if start < 1:
        raise ValueError("Invalid procedure call or argument")

    if find_str == "":
        return start

    idx = s.find(find_str, start - 1)
    return idx + 1 if idx >= 0 else 0


def vba_replace(
    s: Optional[str],
    find: str,
    replacement: str,
    start: int = 1,
    count: int = -1,
) -> str:
    """Replace(expression, find, replace, [start], [count])."""
    s = _safe_str(s)
    if start < 1:
        raise ValueError("Invalid procedure call or argument")

    # VBA Replace returns the string starting from `start` with replacements applied
    result = s[start - 1 :]
    if count < 0:
        result = result.replace(find, replacement)
    else:
        result = result.replace(find, replacement, count)
    return result


def vba_len(s: Optional[str]) -> int:
    """Len(s) -- return length."""
    if s is None:
        return 0
    return len(str(s))


def vba_trim(s: Optional[str]) -> str:
    """Trim(s) -- remove leading and trailing spaces."""
    return _safe_str(s).strip()


def vba_ltrim(s: Optional[str]) -> str:
    """LTrim(s) -- remove leading spaces."""
    return _safe_str(s).lstrip()


def vba_rtrim(s: Optional[str]) -> str:
    """RTrim(s) -- remove trailing spaces."""
    return _safe_str(s).rstrip()


def vba_ucase(s: Optional[str]) -> str:
    """UCase(s) -- convert to uppercase."""
    return _safe_str(s).upper()


def vba_lcase(s: Optional[str]) -> str:
    """LCase(s) -- convert to lowercase."""
    return _safe_str(s).lower()


def vba_space(n: int) -> str:
    """Space(n) -- return n spaces."""
    if n < 0:
        raise ValueError("Invalid procedure call or argument")
    return " " * n


def vba_string_func(n: int, char: Union[str, int]) -> str:
    """String(n, char) -- return char repeated n times."""
    if n < 0:
        raise ValueError("Invalid procedure call or argument")
    if isinstance(char, int):
        char = chr(char)
    return char[0] * n if char else ""


def vba_asc(s: Optional[str]) -> int:
    """Asc(s) -- return ASCII value of first character."""
    s = _safe_str(s)
    if not s:
        raise ValueError("Invalid procedure call or argument")
    return ord(s[0])


def vba_chr(n: int) -> str:
    """Chr(n) -- return character for ASCII value."""
    return chr(n)


def vba_split(
    s: Optional[str], delimiter: str = " ", limit: int = -1
) -> list[str]:
    """Split(expression, [delimiter], [limit])."""
    s = _safe_str(s)
    if limit < 0:
        return s.split(delimiter)
    # VBA limit = max number of substrings; Python maxsplit = limit - 1
    return s.split(delimiter, limit - 1) if limit > 0 else [s]


def vba_join(arr: list, delimiter: str = " ") -> str:
    """Join(sourcearray, [delimiter])."""
    return delimiter.join(str(item) for item in arr)


def vba_strreverse(s: Optional[str]) -> str:
    """StrReverse(s) -- reverse a string."""
    return _safe_str(s)[::-1]


def vba_like(s: Optional[str], pattern: Optional[str]) -> bool:
    """VBA ``Like`` operator: wildcard pattern matching.

    Supports: ``*`` (any chars), ``?`` (single char), ``#`` (single digit),
    ``[charlist]`` and ``[!charlist]`` (character classes).
    """
    import re

    s = _safe_str(s)
    pattern = _safe_str(pattern)
    regex_parts: list[str] = []
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "*":
            regex_parts.append(".*")
        elif ch == "?":
            regex_parts.append(".")
        elif ch == "#":
            regex_parts.append(r"\d")
        elif ch == "[":
            end = pattern.find("]", i + 1)
            if end == -1:
                regex_parts.append(re.escape(ch))
            else:
                inner = pattern[i + 1 : end]
                if inner.startswith("!"):
                    inner = "^" + inner[1:]
                regex_parts.append("[" + inner + "]")
                i = end
        else:
            regex_parts.append(re.escape(ch))
        i += 1
    return re.fullmatch("".join(regex_parts), s) is not None
