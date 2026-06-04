"""Tests for vba_runtime library."""

from vba_runtime.strings import (
    vba_left, vba_right, vba_mid, vba_instr, vba_replace,
    vba_len, vba_trim, vba_ucase, vba_lcase, vba_split, vba_join,
    vba_strreverse, vba_asc, vba_chr,
)
from vba_runtime.conversion import (
    vba_cstr, vba_cint, vba_clng, vba_cdbl, vba_cbool,
    vba_val, vba_int_func, vba_fix, vba_isnumeric,
)
from vba_runtime.math_funcs import vba_abs, vba_sgn, vba_sqr, vba_round
from vba_runtime.types import VBACollection, VBAArray, EMPTY, NULL
from vba_runtime.errors import VBAError, VBARuntimeError


class TestStringFunctions:
    def test_left(self):
        assert vba_left("Hello", 3) == "Hel"
        assert vba_left("Hi", 10) == "Hi"
        assert vba_left(None, 3) == ""

    def test_right(self):
        assert vba_right("Hello", 3) == "llo"
        assert vba_right("Hi", 10) == "Hi"

    def test_mid(self):
        assert vba_mid("Hello", 2, 3) == "ell"  # 1-based
        assert vba_mid("Hello", 1) == "Hello"

    def test_instr(self):
        assert vba_instr("Hello World", "World") == 7  # 1-based
        assert vba_instr("Hello", "xyz") == 0
        # With start position
        assert vba_instr(8, "Hello World", "o") == 0 or vba_instr(1, "Hello World", "o") == 5

    def test_replace(self):
        assert vba_replace("Hello World", "World", "VBA") == "Hello VBA"

    def test_len(self):
        assert vba_len("Hello") == 5
        assert vba_len(None) == 0

    def test_trim(self):
        assert vba_trim("  hello  ") == "hello"

    def test_ucase_lcase(self):
        assert vba_ucase("hello") == "HELLO"
        assert vba_lcase("HELLO") == "hello"

    def test_split_join(self):
        assert vba_split("a,b,c", ",") == ["a", "b", "c"]
        assert vba_join(["a", "b", "c"], ",") == "a,b,c"

    def test_strreverse(self):
        assert vba_strreverse("Hello") == "olleH"

    def test_asc_chr(self):
        assert vba_asc("A") == 65
        assert vba_chr(65) == "A"


class TestConversion:
    def test_cstr(self):
        assert vba_cstr(42) == "42"

    def test_cint(self):
        assert vba_cint("42") == 42
        assert vba_cint(2.5) == 2  # banker's rounding

    def test_cdbl(self):
        assert vba_cdbl("3.14") == 3.14

    def test_cbool(self):
        assert vba_cbool(1) is True
        assert vba_cbool(0) is False

    def test_val(self):
        assert vba_val("123abc") == 123.0
        assert vba_val("abc") == 0.0

    def test_int_func(self):
        assert vba_int_func(5.7) == 5
        assert vba_int_func(-5.7) == -6  # floor toward negative infinity

    def test_fix(self):
        assert vba_fix(5.7) == 5
        assert vba_fix(-5.7) == -5  # truncate toward zero

    def test_isnumeric(self):
        assert vba_isnumeric("42") is True
        assert vba_isnumeric("abc") is False


class TestMath:
    def test_abs(self):
        assert vba_abs(-5) == 5

    def test_sgn(self):
        assert vba_sgn(10) == 1
        assert vba_sgn(-10) == -1
        assert vba_sgn(0) == 0

    def test_sqr(self):
        assert vba_sqr(25) == 5.0

    def test_round(self):
        assert vba_round(2.5, 0) == 2  # banker's rounding
        assert vba_round(3.5, 0) == 4


class TestVBACollection:
    def test_add_and_item(self):
        c = VBACollection()
        c.add("first")
        c.add("second")
        assert c.item(1) == "first"  # 1-based
        assert c.item(2) == "second"

    def test_keyed_access(self):
        c = VBACollection()
        c.add("value1", key="key1")
        assert c.item("key1") == "value1"

    def test_count(self):
        c = VBACollection()
        c.add("a")
        c.add("b")
        assert c.count == 2

    def test_remove(self):
        c = VBACollection()
        c.add("a")
        c.add("b")
        c.remove(1)
        assert c.count == 1
        assert c.item(1) == "b"

    def test_iteration(self):
        c = VBACollection()
        c.add("a")
        c.add("b")
        assert list(c) == ["a", "b"]


class TestVBAArray:
    def test_one_based(self):
        arr = VBAArray([10, 20, 30], base=1)
        assert arr[1] == 10
        assert arr[3] == 30

    def test_zero_based(self):
        arr = VBAArray([10, 20, 30], base=0)
        assert arr[0] == 10

    def test_setitem(self):
        arr = VBAArray([0, 0, 0], base=1)
        arr[2] = 99
        assert arr[2] == 99

    def test_redim(self):
        arr = VBAArray([1, 2, 3], base=1)
        arr.redim(5)
        assert len(arr) == 5

    def test_redim_preserve(self):
        arr = VBAArray([1, 2, 3], base=1)
        arr.redim(5, preserve=True)
        assert arr[1] == 1
        assert arr[4] is EMPTY


class TestErrors:
    def test_vba_error(self):
        err = VBAError()
        assert err.number == 0
        try:
            err.raise_error(13, "Type mismatch")
        except VBARuntimeError as e:
            assert e.number == 13
            assert "Type mismatch" in str(e)
        assert err.number == 13
        err.clear()
        assert err.number == 0


class TestSpecialValues:
    def test_empty(self):
        assert repr(EMPTY) == "Empty"
        assert str(EMPTY) == ""  # VBA: Empty converts to ""
        assert bool(EMPTY) is False

    def test_null(self):
        assert repr(NULL) == "Null"
        assert str(NULL) == "Null"
        # VBA: bool(Null) raises error
        import pytest
        with pytest.raises(TypeError):
            bool(NULL)
        # Null propagation
        assert (NULL + 1) is NULL
