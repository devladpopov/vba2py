"""Tests for code generator: VBA source -> Python source."""

from vba2py.parser import parse
from vba2py.codegen import generate


def vba_to_py(vba: str) -> str:
    """Helper: parse VBA and generate Python."""
    return generate(parse(vba))


class TestSimpleSub:
    def test_empty_sub(self):
        py = vba_to_py("Sub Foo()\nEnd Sub")
        assert "def Foo():" in py
        assert "pass" in py

    def test_hello_world(self):
        py = vba_to_py('Sub Hello()\nDebug.Print "Hello"\nEnd Sub')
        assert "def Hello():" in py
        assert 'print("Hello")' in py

    def test_dim_and_assign(self):
        py = vba_to_py("Sub T()\nDim x As Integer\nx = 42\nEnd Sub")
        assert "x: int = None" in py
        assert "x = 42" in py


class TestFunction:
    def test_function_return(self):
        py = vba_to_py("Function Double(n As Integer) As Integer\nDouble = n * 2\nEnd Function")
        assert "def Double(n)" in py
        assert "Double_result = (n * 2)" in py
        assert "return Double_result" in py

    def test_exit_function(self):
        py = vba_to_py("Function F() As Integer\nIf True Then\nF = 1\nExit Function\nEnd If\nEnd Function")
        assert "return F_result" in py


class TestLoops:
    def test_for_loop(self):
        py = vba_to_py("Sub T()\nFor i = 1 To 10\nx = i\nNext i\nEnd Sub")
        assert "for i in range(1, 10 + 1):" in py

    def test_for_step(self):
        py = vba_to_py("Sub T()\nFor i = 10 To 0 Step -2\nNext\nEnd Sub")
        assert "range(10, 0 + 1, -2)" in py or "range(10, 0 + 1, (-2))" in py

    def test_for_each(self):
        py = vba_to_py("Sub T()\nFor Each item In coll\nNext\nEnd Sub")
        assert "for item in coll:" in py

    def test_while(self):
        py = vba_to_py("Sub T()\nWhile x > 0\nx = x - 1\nWend\nEnd Sub")
        assert "while (x > 0):" in py

    def test_do_while(self):
        py = vba_to_py("Sub T()\nDo While x < 10\nx = x + 1\nLoop\nEnd Sub")
        assert "while (x < 10):" in py

    def test_do_loop_until(self):
        py = vba_to_py("Sub T()\nDo\nx = x + 1\nLoop Until x >= 10\nEnd Sub")
        assert "while True:" in py
        assert "break" in py


class TestConditionals:
    def test_if_else(self):
        py = vba_to_py("Sub T()\nIf x > 0 Then\ny = 1\nElse\ny = 0\nEnd If\nEnd Sub")
        assert "if (x > 0):" in py
        assert "else:" in py

    def test_elseif(self):
        py = vba_to_py("Sub T()\nIf x > 0 Then\ny = 1\nElseIf x = 0 Then\ny = 0\nEnd If\nEnd Sub")
        assert "elif" in py

    def test_select_case(self):
        py = vba_to_py("Sub T()\nSelect Case x\nCase 1\ny = 1\nCase Else\ny = 0\nEnd Select\nEnd Sub")
        assert "if x == 1:" in py
        assert "else:" in py


class TestStringFunctions:
    def test_left(self):
        py = vba_to_py('Sub T()\nx = Left(s, 5)\nEnd Sub')
        assert "vba_left" in py
        assert "from vba_runtime import" in py

    def test_mid(self):
        py = vba_to_py('Sub T()\nx = Mid(s, 1, 3)\nEnd Sub')
        assert "vba_mid" in py

    def test_instr(self):
        py = vba_to_py('Sub T()\nx = InStr(s, "abc")\nEnd Sub')
        assert "vba_instr" in py

    def test_ucase(self):
        py = vba_to_py('Sub T()\nx = UCase(s)\nEnd Sub')
        assert "vba_ucase" in py


class TestExcelObjects:
    def test_cells(self):
        py = vba_to_py("Sub T()\nx = Cells(1, 2)\nEnd Sub")
        assert "ws.cell(row=1, column=2).value" in py

    def test_range(self):
        py = vba_to_py('Sub T()\nx = Range("A1")\nEnd Sub')
        assert 'ws["A1"]' in py


class TestConcatenation:
    def test_ampersand(self):
        py = vba_to_py('Sub T()\nx = "Hello " & name\nEnd Sub')
        assert "str(" in py
        assert "+" in py


class TestWith:
    def test_with_block(self):
        py = vba_to_py("Sub T()\nWith obj\n.Name = 1\nEnd With\nEnd Sub")
        assert "_with_0 = obj" in py
        assert "_with_0.Name = 1" in py


class TestOnError:
    def test_resume_next(self):
        py = vba_to_py("Sub T()\nOn Error Resume Next\nEnd Sub")
        assert "Resume Next" in py or "resume" in py.lower()

    def test_goto_label(self):
        py = vba_to_py("Sub T()\nOn Error GoTo ErrHandler\nEnd Sub")
        assert "GoTo ErrHandler" in py

    def test_goto_zero(self):
        py = vba_to_py("Sub T()\nOn Error GoTo 0\nEnd Sub")
        assert "GoTo 0" in py


class TestReDim:
    def test_redim_basic(self):
        py = vba_to_py("Sub T()\nReDim arr(10)\nEnd Sub")
        assert "arr = [None] * (10 + 1)" in py

    def test_redim_preserve(self):
        py = vba_to_py("Sub T()\nReDim Preserve arr(20)\nEnd Sub")
        assert "extend" in py
        assert "arr" in py


class TestGoTo:
    def test_goto(self):
        py = vba_to_py("Sub T()\nGoTo ErrorHandler\nEnd Sub")
        assert "GoTo ErrorHandler" in py
        assert "TODO" in py


class TestErase:
    def test_erase_array(self):
        py = vba_to_py("Sub T()\nErase arr\nEnd Sub")
        assert "arr = []" in py


class TestTypeHints:
    def test_dim_with_type(self):
        py = vba_to_py("Sub T()\nDim x As Integer\nEnd Sub")
        assert "x: int = None" in py

    def test_dim_string(self):
        py = vba_to_py("Sub T()\nDim s As String\nEnd Sub")
        assert "s: str = None" in py

    def test_dim_double(self):
        py = vba_to_py("Sub T()\nDim d As Double\nEnd Sub")
        assert "d: float = None" in py

    def test_dim_boolean(self):
        py = vba_to_py("Sub T()\nDim b As Boolean\nEnd Sub")
        assert "b: bool = None" in py

    def test_dim_no_type(self):
        py = vba_to_py("Sub T()\nDim x\nEnd Sub")
        assert "x = None" in py
        assert ": " not in py.split("x = None")[0].split("\n")[-1]

    def test_dim_new_object(self):
        py = vba_to_py("Sub T()\nDim c As New Collection\nEnd Sub")
        assert "c = Collection()" in py

    def test_dim_array(self):
        py = vba_to_py("Sub T()\nDim arr(10) As Integer\nEnd Sub")
        assert "arr = []" in py


class TestConstant:
    def test_const_string(self):
        py = vba_to_py('Sub T()\nConst NAME = "Hello"\nEnd Sub')
        assert 'NAME = "Hello"' in py

    def test_const_number(self):
        py = vba_to_py("Sub T()\nConst MAX_SIZE = 100\nEnd Sub")
        assert "MAX_SIZE = 100" in py


class TestWithMethodCalls:
    def test_with_dot_call(self):
        py = vba_to_py('Sub T()\nWith obj\n.Add "item"\nEnd With\nEnd Sub')
        assert "_with_0.Add(" in py

    def test_with_dot_assign(self):
        py = vba_to_py('Sub T()\nWith obj\n.Name = "test"\nEnd With\nEnd Sub')
        assert '_with_0.Name = "test"' in py


class TestNested:
    def test_nested_if_in_for(self):
        py = vba_to_py("Sub T()\nFor i = 1 To 10\nIf i > 5 Then\nx = i\nEnd If\nNext i\nEnd Sub")
        assert "for i in range(1, 10 + 1):" in py
        assert "if (i > 5):" in py

    def test_for_with_exit(self):
        py = vba_to_py("Sub T()\nFor i = 1 To 10\nIf i = 5 Then\nExit For\nEnd If\nNext i\nEnd Sub")
        assert "break" in py


class TestLikeOperator:
    def test_like_generates_runtime_call(self):
        py = vba_to_py('Sub T()\nIf s Like "a*" Then\nx = 1\nEnd If\nEnd Sub')
        assert 'vba_like(s, "a*")' in py
        assert "from vba_runtime import" in py
        assert "vba_like" in py
        compile(py, "<gen>", "exec")
