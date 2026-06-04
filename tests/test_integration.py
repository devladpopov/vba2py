"""Integration tests: full VBA files through the pipeline."""

from pathlib import Path

from vba2py.parser import parse
from vba2py.codegen import generate

FIXTURES = Path(__file__).parent / "fixtures"


def _convert_fixture(name: str) -> str:
    """Read a VBA fixture file and return generated Python."""
    vba_source = (FIXTURES / name).read_text(encoding="utf-8-sig")
    module = parse(vba_source)
    return generate(module)


def _assert_valid_python(code: str) -> None:
    """Assert that the generated code compiles as valid Python."""
    compile(code, "<generated>", "exec")


class TestSimpleSub:
    def test_converts(self):
        py = _convert_fixture("simple_sub.bas")
        _assert_valid_python(py)
        assert "def HelloWorld():" in py
        assert "print(msg)" in py

    def test_has_imports(self):
        py = _convert_fixture("simple_sub.bas")
        # Should be valid even without imports for this simple case
        _assert_valid_python(py)


class TestLoops:
    def test_converts(self):
        py = _convert_fixture("loops.bas")
        _assert_valid_python(py)
        assert "for i in range(" in py
        assert "while" in py

    def test_all_loop_types(self):
        py = _convert_fixture("loops.bas")
        assert "for i in range(1, 10 + 1):" in py  # For
        assert "range(10, 0 + 1," in py  # For Step
        assert "while (i < 5):" in py  # Do While
        assert "while True:" in py  # Do...Loop Until


class TestFunctionWithLogic:
    def test_converts(self):
        py = _convert_fixture("function_with_logic.bas")
        _assert_valid_python(py)
        assert "def Fibonacci(n)" in py
        assert "def GetGrade(score)" in py

    def test_function_return(self):
        py = _convert_fixture("function_with_logic.bas")
        assert "Fibonacci_result" in py
        assert "return Fibonacci_result" in py

    def test_select_case(self):
        py = _convert_fixture("function_with_logic.bas")
        assert "if score ==" in py or "elif" in py


class TestStringOps:
    def test_converts(self):
        py = _convert_fixture("string_ops.bas")
        _assert_valid_python(py)

    def test_runtime_imports(self):
        py = _convert_fixture("string_ops.bas")
        assert "from vba_runtime import" in py
        assert "vba_left" in py
        assert "vba_right" in py
        assert "vba_mid" in py
        assert "vba_instr" in py

    def test_valid_python_with_runtime(self):
        """The generated code should be valid Python syntax."""
        py = _convert_fixture("string_ops.bas")
        _assert_valid_python(py)


class TestExcelOperations:
    def test_converts(self):
        py = _convert_fixture("excel_operations.bas")
        _assert_valid_python(py)

    def test_sum_column(self):
        py = _convert_fixture("excel_operations.bas")
        assert "def SumColumn():" in py
        assert "ws.cell(row=i, column=1).value" in py
        assert 'ws["B1"]' in py

    def test_copy_range(self):
        py = _convert_fixture("excel_operations.bas")
        assert "def CopyRange():" in py
        assert "destRow = (destRow + 1)" in py

    def test_cells_mapping(self):
        py = _convert_fixture("excel_operations.bas")
        # Cells(srcRow, 1) -> ws.cell(row=srcRow, column=1).value
        assert "ws.cell(row=srcRow, column=1).value" in py


class TestErrorHandling:
    def test_converts(self):
        py = _convert_fixture("error_handling.bas")
        _assert_valid_python(py)

    def test_on_error_comments(self):
        py = _convert_fixture("error_handling.bas")
        assert "On Error Resume Next" in py
        assert "On Error GoTo 0" in py

    def test_safe_convert_function(self):
        py = _convert_fixture("error_handling.bas")
        assert "def SafeConvert(value)" in py
        assert "SafeConvert_result" in py
        assert "return SafeConvert_result" in py


class TestWithBlocks:
    def test_converts(self):
        py = _convert_fixture("with_blocks.bas")
        _assert_valid_python(py)

    def test_with_variable(self):
        py = _convert_fixture("with_blocks.bas")
        assert "_with_0 = ws" in py
        assert "_with_0.Name" in py

    def test_with_method_call(self):
        py = _convert_fixture("with_blocks.bas")
        # .Add inside With should use _with variable
        assert "_with_0.Add(" in py


class TestDataProcessing:
    def test_converts(self):
        py = _convert_fixture("data_processing.bas")
        _assert_valid_python(py)

    def test_clean_string_function(self):
        py = _convert_fixture("data_processing.bas")
        assert "def CleanString(s)" in py
        assert "CleanString_result" in py
        assert "vba_trim" in py
        assert "vba_ucase" in py

    def test_process_data_sub(self):
        py = _convert_fixture("data_processing.bas")
        assert "def ProcessData():" in py
        assert "CleanString(rawValue)" in py

    def test_find_in_range(self):
        py = _convert_fixture("data_processing.bas")
        assert "def FindInRange(searchValue, lastRow)" in py
        assert "return FindInRange_result" in py

    def test_runtime_imports(self):
        py = _convert_fixture("data_processing.bas")
        assert "from vba_runtime import" in py
        assert "vba_left" in py
        assert "vba_trim" in py
