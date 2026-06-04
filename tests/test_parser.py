"""Tests for VBA parser."""

from vba2py.parser import parse
from vba2py.ast_nodes import (
    Module, Sub, Function, ForStmt, IfStmt, DimStmt,
    CallStmt, AssignStmt, WhileStmt, DoWhileStmt, DoUntilStmt,
    SelectCaseStmt, ForEachStmt, WithStmt, ExitStmt, OnErrorStmt,
    SetStmt, LabelStmt, ReDimStmt, GoToStmt, EraseStmt, OptionStmt,
)


class TestSubParsing:
    def test_empty_sub(self):
        mod = parse("Sub Foo()\nEnd Sub")
        assert len(mod.procedures) == 1
        assert isinstance(mod.procedures[0], Sub)
        assert mod.procedures[0].name == "Foo"
        assert mod.procedures[0].body == []

    def test_sub_with_params(self):
        mod = parse("Sub Bar(x As Integer, ByVal y As String)\nEnd Sub")
        sub = mod.procedures[0]
        assert len(sub.params) == 2
        assert sub.params[0].name == "x"
        assert sub.params[0].param_type == "Integer"
        assert sub.params[1].passing == "ByVal"

    def test_private_sub(self):
        mod = parse("Private Sub Secret()\nEnd Sub")
        assert mod.procedures[0].access == "Private"


class TestFunctionParsing:
    def test_function_with_return_type(self):
        mod = parse("Function Add(a As Integer, b As Integer) As Integer\nAdd = a + b\nEnd Function")
        func = mod.procedures[0]
        assert isinstance(func, Function)
        assert func.return_type == "Integer"
        assert len(func.body) == 1


class TestStatements:
    def test_dim(self):
        mod = parse("Sub T()\nDim x As Integer\nEnd Sub")
        body = mod.procedures[0].body
        assert len(body) == 1
        assert isinstance(body[0], DimStmt)
        assert body[0].declarations[0].name == "x"

    def test_assignment(self):
        mod = parse("Sub T()\nx = 42\nEnd Sub")
        body = mod.procedures[0].body
        assert isinstance(body[0], AssignStmt)

    def test_set(self):
        mod = parse("Sub T()\nSet obj = New Collection\nEnd Sub")
        body = mod.procedures[0].body
        assert isinstance(body[0], SetStmt)

    def test_call_explicit(self):
        mod = parse("Sub T()\nCall Foo(1, 2)\nEnd Sub")
        body = mod.procedures[0].body
        assert isinstance(body[0], CallStmt)
        assert body[0].name == "Foo"
        assert len(body[0].args) == 2

    def test_debug_print(self):
        mod = parse('Sub T()\nDebug.Print "hello"\nEnd Sub')
        body = mod.procedures[0].body
        assert isinstance(body[0], CallStmt)
        assert body[0].name == "Debug.Print"

    def test_exit_sub(self):
        mod = parse("Sub T()\nExit Sub\nEnd Sub")
        body = mod.procedures[0].body
        assert isinstance(body[0], ExitStmt)
        assert body[0].kind == "Sub"

    def test_label(self):
        mod = parse("Sub T()\nErrorHandler:\nx = 1\nEnd Sub")
        body = mod.procedures[0].body
        assert isinstance(body[0], LabelStmt)
        assert body[0].name == "ErrorHandler"


class TestControlFlow:
    def test_if_then_else(self):
        code = "Sub T()\nIf x > 0 Then\ny = 1\nElseIf x = 0 Then\ny = 0\nElse\ny = -1\nEnd If\nEnd Sub"
        mod = parse(code)
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, IfStmt)
        assert len(stmt.elseif_clauses) == 1
        assert stmt.else_body is not None

    def test_single_line_if(self):
        mod = parse("Sub T()\nIf x > 0 Then y = 1\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, IfStmt)
        assert len(stmt.then_body) == 1

    def test_for_loop(self):
        mod = parse("Sub T()\nFor i = 1 To 10\nx = i\nNext i\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, ForStmt)
        assert stmt.var == "i"

    def test_for_step(self):
        mod = parse("Sub T()\nFor i = 10 To 0 Step -1\nNext\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, ForStmt)
        assert stmt.step is not None

    def test_for_each(self):
        mod = parse("Sub T()\nFor Each item In collection\nNext\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, ForEachStmt)

    def test_do_while(self):
        mod = parse("Sub T()\nDo While x > 0\nx = x - 1\nLoop\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, DoWhileStmt)
        assert stmt.check_post is False

    def test_do_loop_until(self):
        mod = parse("Sub T()\nDo\nx = x + 1\nLoop Until x >= 10\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, DoUntilStmt)
        assert stmt.check_post is True

    def test_while_wend(self):
        mod = parse("Sub T()\nWhile x > 0\nx = x - 1\nWend\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, WhileStmt)

    def test_select_case(self):
        code = "Sub T()\nSelect Case x\nCase 1\ny = 1\nCase 2, 3\ny = 2\nCase Else\ny = 0\nEnd Select\nEnd Sub"
        mod = parse(code)
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, SelectCaseStmt)
        assert len(stmt.cases) == 2
        assert stmt.else_body is not None

    def test_with(self):
        mod = parse("Sub T()\nWith obj\n.Name = 1\nEnd With\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, WithStmt)

    def test_on_error_resume_next(self):
        mod = parse("Sub T()\nOn Error Resume Next\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, OnErrorStmt)
        assert stmt.action == "ResumeNext"

    def test_on_error_goto(self):
        mod = parse("Sub T()\nOn Error GoTo ErrHandler\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, OnErrorStmt)
        assert stmt.label == "ErrHandler"


class TestModuleLevel:
    def test_multiple_procedures(self):
        code = "Sub A()\nEnd Sub\n\nSub B()\nEnd Sub\n\nFunction C() As Integer\nC = 1\nEnd Function"
        mod = parse(code)
        assert len(mod.procedures) == 3

    def test_module_dim(self):
        code = "Dim x As Integer\nSub A()\nEnd Sub"
        mod = parse(code)
        assert len(mod.declarations) == 1
        assert len(mod.procedures) == 1


class TestNewConstructs:
    def test_redim(self):
        mod = parse("Sub T()\nReDim arr(10)\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, ReDimStmt)
        assert stmt.name == "arr"
        assert not stmt.preserve

    def test_redim_preserve(self):
        mod = parse("Sub T()\nReDim Preserve arr(20)\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, ReDimStmt)
        assert stmt.preserve
        assert stmt.name == "arr"

    def test_goto(self):
        mod = parse("Sub T()\nGoTo ErrorHandler\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, GoToStmt)
        assert stmt.label == "ErrorHandler"

    def test_erase(self):
        mod = parse("Sub T()\nErase arr\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, EraseStmt)
        assert "arr" in stmt.arrays

    def test_option_explicit(self):
        mod = parse("Option Explicit\nSub T()\nEnd Sub")
        assert len(mod.declarations) == 1
        assert isinstance(mod.declarations[0], OptionStmt)

    def test_dim_multiple(self):
        mod = parse("Sub T()\nDim x As Integer, y As String, z\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, DimStmt)
        assert len(stmt.declarations) == 3
        assert stmt.declarations[0].var_type == "Integer"
        assert stmt.declarations[1].var_type == "String"
        assert stmt.declarations[2].var_type is None

    def test_dim_array(self):
        mod = parse("Sub T()\nDim arr(10) As Integer\nEnd Sub")
        stmt = mod.procedures[0].body[0]
        assert isinstance(stmt, DimStmt)
        assert stmt.declarations[0].is_array

    def test_optional_param(self):
        mod = parse("Sub T(Optional x As Integer = 5)\nEnd Sub")
        param = mod.procedures[0].params[0]
        assert param.optional
        assert param.default is not None

    def test_byval_param(self):
        mod = parse("Function F(ByVal x As String) As String\nF = x\nEnd Function")
        param = mod.procedures[0].params[0]
        assert param.passing == "ByVal"

    def test_nested_if(self):
        code = "Sub T()\nIf x > 0 Then\nIf y > 0 Then\nz = 1\nEnd If\nEnd If\nEnd Sub"
        mod = parse(code)
        outer_if = mod.procedures[0].body[0]
        assert isinstance(outer_if, IfStmt)
        inner_if = outer_if.then_body[0]
        assert isinstance(inner_if, IfStmt)

    def test_nested_for(self):
        code = "Sub T()\nFor i = 1 To 5\nFor j = 1 To 5\nx = i * j\nNext j\nNext i\nEnd Sub"
        mod = parse(code)
        outer_for = mod.procedures[0].body[0]
        assert isinstance(outer_for, ForStmt)
        inner_for = outer_for.body[0]
        assert isinstance(inner_for, ForStmt)
