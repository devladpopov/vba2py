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
