"""Tests for VBA extraction from Office files."""

from pathlib import Path

import pytest

from vba2py.extract import (
    ExtractionError, extract_vba, is_office_file, _strip_attributes,
)


class TestIsOfficeFile:
    def test_xlsm(self):
        assert is_office_file(Path("book.xlsm"))
        assert is_office_file(Path("BOOK.XLSM"))

    def test_xls_docm(self):
        assert is_office_file(Path("legacy.xls"))
        assert is_office_file(Path("doc.docm"))

    def test_plain_sources(self):
        assert not is_office_file(Path("module.bas"))
        assert not is_office_file(Path("module.vba"))
        assert not is_office_file(Path("module.txt"))


class TestStripAttributes:
    def test_strips_vb_attributes(self):
        code = 'Attribute VB_Name = "Module1"\nSub Foo()\nEnd Sub'
        assert _strip_attributes(code) == "Sub Foo()\nEnd Sub"

    def test_keeps_normal_code(self):
        code = "Sub Foo()\nx = 1\nEnd Sub"
        assert _strip_attributes(code) == code


class TestExtractVba:
    def test_no_macros_raises(self, tmp_path):
        openpyxl = pytest.importorskip("openpyxl")
        xlsm = tmp_path / "empty.xlsm"
        wb = openpyxl.Workbook()
        wb.save(xlsm)
        with pytest.raises(ExtractionError):
            extract_vba(xlsm)

    def test_missing_file_raises(self, tmp_path):
        with pytest.raises(Exception):
            extract_vba(tmp_path / "nope.xlsm")
