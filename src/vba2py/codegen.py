"""Generate Python source code from a VBA AST.

Strategy: walk VBA AST nodes and produce Python source lines directly.
We avoid Python's ``ast`` module to keep the output readable and allow
VBA-specific comments/annotations.
"""

from __future__ import annotations

from vba2py.ast_nodes import (
    Module, Sub, Function, Param, VarDecl, ConstDecl,
    AssignStmt, SetStmt, CallStmt, IfStmt, SelectCaseStmt, CaseClause,
    ForStmt, ForEachStmt, DoWhileStmt, DoUntilStmt, WhileStmt,
    WithStmt, DimStmt, ExitStmt, ReturnStmt, OnErrorStmt, LabelStmt,
    Literal, Identifier, MemberAccess, IndexAccess, BinaryOp, UnaryOp,
    NewExpr, DotAccess, Expression, Statement, ASTNode,
)

# VBA built-in functions that map to vba_runtime helpers
_RUNTIME_FUNCS: dict[str, str] = {
    "left": "vba_left", "right": "vba_right", "mid": "vba_mid",
    "instr": "vba_instr", "replace": "vba_replace",
    "len": "vba_len", "trim": "vba_trim", "ltrim": "vba_ltrim",
    "rtrim": "vba_rtrim", "ucase": "vba_ucase", "lcase": "vba_lcase",
    "space": "vba_space", "asc": "vba_asc", "chr": "vba_chr",
    "split": "vba_split", "join": "vba_join", "strreverse": "vba_strreverse",
    "cstr": "vba_cstr", "cint": "vba_cint", "clng": "vba_clng",
    "cdbl": "vba_cdbl", "cbool": "vba_cbool",
    "val": "vba_val", "int": "vba_int_func", "fix": "vba_fix",
    "abs": "abs", "sgn": "vba_sgn", "sqr": "vba_sqr",
    "isnumeric": "vba_isnumeric", "isempty": "vba_isempty",
    "isnull": "vba_isnull", "isdate": "vba_isdate",
    "typename": "vba_typename",
    "round": "vba_round", "rnd": "vba_rnd",
    "msgbox": "print", "inputbox": "input",
    "string": "vba_string_func",
    "array": "list",
}

# VBA type names -> Python type hints
_TYPE_MAP: dict[str, str] = {
    "integer": "int", "long": "int", "byte": "int",
    "single": "float", "double": "float", "currency": "float",
    "string": "str", "boolean": "bool", "variant": "Any",
    "object": "Any", "date": "Any",
}

# Binary operator mapping
_OP_MAP: dict[str, str] = {
    "+": "+", "-": "-", "*": "*", "/": "/",
    "\\": "//", "^": "**", "&": "+",
    "=": "==", "<>": "!=", "<": "<", ">": ">",
    "<=": "<=", ">=": ">=",
    "And": "and", "Or": "or", "Not": "not",
    "Xor": "^", "Mod": "%",
    "Is": "is", "Like": "like",  # 'like' needs runtime
}


class CodeGenerator:
    """Walk a VBA AST and emit Python source code."""

    def __init__(self) -> None:
        self._lines: list[str] = []
        self._indent = 0
        self._with_stack: list[str] = []
        self._current_func: str | None = None
        self._runtime_imports: set[str] = set()
        self._needs_openpyxl = False

    def generate(self, module: Module) -> str:
        """Generate Python source for an entire module."""
        for decl in module.declarations:
            self._emit_node(decl)
        for proc in module.procedures:
            self._emit_blank()
            self._emit_blank()
            self._emit_node(proc)

        header = self._build_header()
        body = "\n".join(self._lines)
        if header and body:
            return header + "\n\n" + body + "\n"
        return (header or body) + "\n"

    # -- Header / imports ---------------------------------------------------

    def _build_header(self) -> str:
        lines: list[str] = []
        if self._runtime_imports:
            names = sorted(self._runtime_imports)
            lines.append(f"from vba_runtime import {', '.join(names)}")
        if self._needs_openpyxl:
            lines.append("import openpyxl")
        return "\n".join(lines)

    # -- Output helpers -----------------------------------------------------

    def _emit(self, text: str) -> None:
        self._lines.append("    " * self._indent + text)

    def _emit_blank(self) -> None:
        self._lines.append("")

    # -- Dispatch -----------------------------------------------------------

    def _emit_node(self, node: ASTNode) -> None:
        method = f"_emit_{type(node).__name__}"
        handler = getattr(self, method, None)
        if handler:
            handler(node)
        else:
            self._emit(f"# TODO: unsupported {type(node).__name__}")

    # -- Procedures ---------------------------------------------------------

    def _emit_Sub(self, node: Sub) -> None:
        params = ", ".join(self._param_str(p) for p in node.params)
        self._emit(f"def {self._py_name(node.name)}({params}):")
        self._indent += 1
        self._emit_body(node.body)
        self._indent -= 1

    def _emit_Function(self, node: Function) -> None:
        self._current_func = node.name
        params = ", ".join(self._param_str(p) for p in node.params)
        ret = ""
        if node.return_type:
            py_type = _TYPE_MAP.get(node.return_type.lower(), node.return_type)
            ret = f" -> {py_type}"
        self._emit(f"def {self._py_name(node.name)}({params}){ret}:")
        self._indent += 1
        # Initialize return variable
        self._emit(f"{self._py_name(node.name)}_result = None")
        self._emit_body(node.body)
        self._emit(f"return {self._py_name(node.name)}_result")
        self._indent -= 1
        self._current_func = None

    def _param_str(self, p: Param) -> str:
        name = self._py_name(p.name)
        if p.default:
            return f"{name}={self._expr(p.default)}"
        if p.optional:
            return f"{name}=None"
        return name

    # -- Statements ---------------------------------------------------------

    def _emit_DimStmt(self, node: DimStmt) -> None:
        for decl in node.declarations:
            name = self._py_name(decl.name)
            if decl.is_array:
                self._emit(f"{name} = []")
            elif decl.var_type and decl.var_type.startswith("New "):
                cls = decl.var_type[4:]
                self._emit(f"{name} = {cls}()")
            else:
                self._emit(f"{name} = None")

    def _emit_ConstDecl(self, node: ConstDecl) -> None:
        name = self._py_name(node.name)
        val = self._expr(node.value) if node.value else "None"
        self._emit(f"{name} = {val}")

    def _emit_VarDecl(self, node: VarDecl) -> None:
        name = self._py_name(node.name)
        self._emit(f"{name} = None")

    def _emit_AssignStmt(self, node: AssignStmt) -> None:
        target_str = self._expr(node.target)
        value_str = self._expr(node.value)

        # Check if this is a function return assignment: FuncName = value
        if (self._current_func
                and isinstance(node.target, Identifier)
                and node.target.name.lower() == self._current_func.lower()):
            self._emit(f"{self._py_name(self._current_func)}_result = {value_str}")
            return

        # Handle Range/Cells .Value assignment
        self._emit(f"{target_str} = {value_str}")

    def _emit_SetStmt(self, node: SetStmt) -> None:
        self._emit(f"{self._expr(node.target)} = {self._expr(node.value)}")

    def _emit_CallStmt(self, node: CallStmt) -> None:
        name = node.name
        args_str = ", ".join(self._expr(a) for a in node.args)

        # Debug.Print -> print()
        if name == "Debug.Print":
            self._emit(f"print({args_str})")
            return

        # Check for runtime function mapping
        low = name.lower()
        if low in _RUNTIME_FUNCS:
            py_name = _RUNTIME_FUNCS[low]
            if py_name.startswith("vba_"):
                self._runtime_imports.add(py_name)
            self._emit(f"{py_name}({args_str})")
            return

        # Member call: obj.Method(args)
        if "." in name:
            parts = name.split(".")
            py_parts = [self._py_name(p) for p in parts]
            self._emit(f"{'.'.join(py_parts)}({args_str})")
            return

        self._emit(f"{self._py_name(name)}({args_str})")

    def _emit_IfStmt(self, node: IfStmt) -> None:
        cond = self._expr(node.condition)
        self._emit(f"if {cond}:")
        self._indent += 1
        self._emit_body(node.then_body)
        self._indent -= 1

        for ei_cond, ei_body in node.elseif_clauses:
            self._emit(f"elif {self._expr(ei_cond)}:")
            self._indent += 1
            self._emit_body(ei_body)
            self._indent -= 1

        if node.else_body:
            self._emit("else:")
            self._indent += 1
            self._emit_body(node.else_body)
            self._indent -= 1

    def _emit_SelectCaseStmt(self, node: SelectCaseStmt) -> None:
        var = self._expr(node.expr)
        for i, case in enumerate(node.cases):
            vals = ", ".join(self._expr(e) for e in case.expressions)
            keyword = "if" if i == 0 else "elif"
            if len(case.expressions) == 1:
                self._emit(f"{keyword} {var} == {vals}:")
            else:
                self._emit(f"{keyword} {var} in ({vals}):")
            self._indent += 1
            self._emit_body(case.body)
            self._indent -= 1

        if node.else_body:
            self._emit("else:")
            self._indent += 1
            self._emit_body(node.else_body)
            self._indent -= 1

    def _emit_ForStmt(self, node: ForStmt) -> None:
        var = self._py_name(node.var)
        start = self._expr(node.start)
        end = self._expr(node.end)
        if node.step:
            step = self._expr(node.step)
            self._emit(f"for {var} in range({start}, {end} + 1, {step}):")
        else:
            self._emit(f"for {var} in range({start}, {end} + 1):")
        self._indent += 1
        self._emit_body(node.body)
        self._indent -= 1

    def _emit_ForEachStmt(self, node: ForEachStmt) -> None:
        var = self._py_name(node.var)
        coll = self._expr(node.collection)
        self._emit(f"for {var} in {coll}:")
        self._indent += 1
        self._emit_body(node.body)
        self._indent -= 1

    def _emit_DoWhileStmt(self, node: DoWhileStmt) -> None:
        cond = self._expr(node.condition)
        if node.check_post:
            self._emit("while True:")
            self._indent += 1
            self._emit_body(node.body)
            self._emit(f"if not ({cond}):")
            self._indent += 1
            self._emit("break")
            self._indent -= 1
            self._indent -= 1
        else:
            self._emit(f"while {cond}:")
            self._indent += 1
            self._emit_body(node.body)
            self._indent -= 1

    def _emit_DoUntilStmt(self, node: DoUntilStmt) -> None:
        cond = self._expr(node.condition)
        if node.check_post:
            self._emit("while True:")
            self._indent += 1
            self._emit_body(node.body)
            self._emit(f"if {cond}:")
            self._indent += 1
            self._emit("break")
            self._indent -= 1
            self._indent -= 1
        else:
            self._emit(f"while not ({cond}):")
            self._indent += 1
            self._emit_body(node.body)
            self._indent -= 1

    def _emit_WhileStmt(self, node: WhileStmt) -> None:
        cond = self._expr(node.condition)
        self._emit(f"while {cond}:")
        self._indent += 1
        self._emit_body(node.body)
        self._indent -= 1

    def _emit_WithStmt(self, node: WithStmt) -> None:
        target = self._expr(node.target)
        with_var = f"_with_{len(self._with_stack)}"
        self._with_stack.append(with_var)
        self._emit(f"{with_var} = {target}")
        for stmt in node.body:
            self._emit_node(stmt)
        self._with_stack.pop()

    def _emit_ExitStmt(self, node: ExitStmt) -> None:
        if node.kind in ("For", "Do"):
            self._emit("break")
        elif node.kind == "Sub":
            self._emit("return")
        elif node.kind == "Function":
            if self._current_func:
                self._emit(f"return {self._py_name(self._current_func)}_result")
            else:
                self._emit("return")

    def _emit_ReturnStmt(self, node: ReturnStmt) -> None:
        if node.value:
            self._emit(f"return {self._expr(node.value)}")
        else:
            self._emit("return")

    def _emit_OnErrorStmt(self, node: OnErrorStmt) -> None:
        if node.action == "ResumeNext":
            self._emit("# On Error Resume Next: errors will be silently caught")
            self._emit("# TODO: wrap subsequent statements in try/except")
        elif node.action == "GoTo0":
            self._emit("# On Error GoTo 0: reset error handling")
        elif node.action == "GoTo" and node.label:
            self._emit(f"# On Error GoTo {node.label}: error handling")
            self._emit("# TODO: restructure with try/except")

    def _emit_LabelStmt(self, node: LabelStmt) -> None:
        self._emit(f"# label: {node.name}")

    # -- Expression rendering -----------------------------------------------

    def _expr(self, node: Expression | None) -> str:
        if node is None:
            return "None"

        if isinstance(node, Literal):
            return self._literal(node)
        if isinstance(node, Identifier):
            return self._identifier(node)
        if isinstance(node, MemberAccess):
            return f"{self._expr(node.target)}.{self._py_name(node.member)}"
        if isinstance(node, IndexAccess):
            return self._index_access(node)
        if isinstance(node, BinaryOp):
            return self._binary_op(node)
        if isinstance(node, UnaryOp):
            return self._unary_op(node)
        if isinstance(node, NewExpr):
            return f"{node.type_name}()"
        if isinstance(node, DotAccess):
            return self._dot_access(node)
        return f"# unsupported expr: {type(node).__name__}"

    def _literal(self, node: Literal) -> str:
        if node.lit_type == "string":
            escaped = node.value.replace("\\", "\\\\").replace('"', '\\"')
            return f'"{escaped}"'
        if node.lit_type == "boolean":
            return "True" if node.value else "False"
        if node.lit_type == "nothing":
            return "None"
        if node.lit_type == "date":
            return f'"{node.value}"  # VBA Date literal'
        return str(node.value)

    def _identifier(self, node: Identifier) -> str:
        low = node.name.lower()
        if low in _RUNTIME_FUNCS:
            py = _RUNTIME_FUNCS[low]
            if py.startswith("vba_"):
                self._runtime_imports.add(py)
            return py
        return self._py_name(node.name)

    def _index_access(self, node: IndexAccess) -> str:
        target = node.target
        args_str = ", ".join(self._expr(a) for a in node.args)

        # Check if target is a known runtime function
        if isinstance(target, Identifier):
            low = target.name.lower()
            if low in _RUNTIME_FUNCS:
                py_name = _RUNTIME_FUNCS[low]
                if py_name.startswith("vba_"):
                    self._runtime_imports.add(py_name)
                return f"{py_name}({args_str})"

        target_str = self._expr(target)

        # Cells(r, c) -> ws.cell(row=r, column=c).value
        if isinstance(target, Identifier) and target.name.lower() == "cells":
            if len(node.args) == 2:
                return f"ws.cell(row={self._expr(node.args[0])}, column={self._expr(node.args[1])}).value"

        # Range("A1") -> ws["A1"]
        if isinstance(target, Identifier) and target.name.lower() == "range":
            if len(node.args) == 1:
                return f"ws[{self._expr(node.args[0])}]"

        # Generic function call or array access
        return f"{target_str}({args_str})"

    def _binary_op(self, node: BinaryOp) -> str:
        left = self._expr(node.left)
        right = self._expr(node.right)
        op = _OP_MAP.get(node.op, node.op)

        # String concatenation: & -> str() + str()
        if node.op == "&":
            return f"str({left}) + str({right})"

        return f"({left} {op} {right})"

    def _unary_op(self, node: UnaryOp) -> str:
        operand = self._expr(node.operand)
        if node.op == "Not":
            return f"not ({operand})"
        return f"({node.op}{operand})"

    def _dot_access(self, node: DotAccess) -> str:
        if self._with_stack:
            return f"{self._with_stack[-1]}.{self._py_name(node.member)}"
        return f"self.{self._py_name(node.member)}"

    # -- Helpers ------------------------------------------------------------

    def _emit_body(self, stmts: list[Statement]) -> None:
        if not stmts:
            self._emit("pass")
            return
        for s in stmts:
            self._emit_node(s)

    @staticmethod
    def _py_name(name: str) -> str:
        """Convert VBA name to Python-safe name (snake_case for known patterns)."""
        # Reserved words
        _reserved = {"print", "type", "input", "list", "range", "format",
                      "next", "id", "dir", "open", "close", "len"}
        result = name
        if result.lower() in _reserved:
            result = result + "_"
        return result


# -- Public API -------------------------------------------------------------

def generate(module: Module) -> str:
    """Generate Python source code from a VBA AST Module."""
    return CodeGenerator().generate(module)
