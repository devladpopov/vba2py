"""Typed VBA AST node hierarchy using Python dataclasses."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


# ---------------------------------------------------------------------------
# Base classes
# ---------------------------------------------------------------------------

@dataclass
class ASTNode:
    """Root of every AST node.  ``lineno`` is patched by the parser."""
    lineno: int = field(default=0, repr=False, compare=False)


@dataclass
class Statement(ASTNode):
    """Abstract base for all statements."""


@dataclass
class Expression(ASTNode):
    """Abstract base for all expressions."""


# ---------------------------------------------------------------------------
# Top-level
# ---------------------------------------------------------------------------

@dataclass
class Module(ASTNode):
    name: str = ""
    declarations: list[ASTNode] = field(default_factory=list)
    procedures: list[ASTNode] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Declarations
# ---------------------------------------------------------------------------

@dataclass
class VarDecl(ASTNode):
    name: str = ""
    var_type: str | None = None
    is_array: bool = False


@dataclass
class ConstDecl(ASTNode):
    name: str = ""
    var_type: str | None = None
    value: Expression | None = None


# ---------------------------------------------------------------------------
# Procedures & parameters
# ---------------------------------------------------------------------------

@dataclass
class Param(ASTNode):
    name: str = ""
    param_type: str | None = None
    passing: str = "ByRef"
    optional: bool = False
    default: Expression | None = None


@dataclass
class Sub(ASTNode):
    name: str = ""
    params: list[Param] = field(default_factory=list)
    body: list[Statement] = field(default_factory=list)
    access: str = "Public"


@dataclass
class Function(ASTNode):
    name: str = ""
    params: list[Param] = field(default_factory=list)
    body: list[Statement] = field(default_factory=list)
    return_type: str | None = None
    access: str = "Public"


# ---------------------------------------------------------------------------
# Statements
# ---------------------------------------------------------------------------

@dataclass
class AssignStmt(Statement):
    target: Expression | None = None
    value: Expression | None = None


@dataclass
class SetStmt(Statement):
    target: Expression | None = None
    value: Expression | None = None


@dataclass
class CallStmt(Statement):
    name: str = ""
    args: list[Expression] = field(default_factory=list)


@dataclass
class IfStmt(Statement):
    condition: Expression | None = None
    then_body: list[Statement] = field(default_factory=list)
    elseif_clauses: list[tuple[Expression, list[Statement]]] = field(default_factory=list)
    else_body: list[Statement] | None = None


@dataclass
class CaseClause(ASTNode):
    expressions: list[Expression] = field(default_factory=list)
    body: list[Statement] = field(default_factory=list)


@dataclass
class SelectCaseStmt(Statement):
    expr: Expression | None = None
    cases: list[CaseClause] = field(default_factory=list)
    else_body: list[Statement] | None = None


@dataclass
class ForStmt(Statement):
    var: str = ""
    start: Expression | None = None
    end: Expression | None = None
    step: Expression | None = None
    body: list[Statement] = field(default_factory=list)


@dataclass
class ForEachStmt(Statement):
    var: str = ""
    collection: Expression | None = None
    body: list[Statement] = field(default_factory=list)


@dataclass
class DoWhileStmt(Statement):
    condition: Expression | None = None
    body: list[Statement] = field(default_factory=list)
    check_post: bool = False


@dataclass
class DoUntilStmt(Statement):
    condition: Expression | None = None
    body: list[Statement] = field(default_factory=list)
    check_post: bool = False


@dataclass
class WhileStmt(Statement):
    condition: Expression | None = None
    body: list[Statement] = field(default_factory=list)


@dataclass
class WithStmt(Statement):
    target: Expression | None = None
    body: list[Statement] = field(default_factory=list)


@dataclass
class DimStmt(Statement):
    declarations: list[VarDecl] = field(default_factory=list)


@dataclass
class ExitStmt(Statement):
    kind: str = ""  # "Sub", "Function", "For", "Do"


@dataclass
class ReturnStmt(Statement):
    """Represents ``FunctionName = expr`` (Function return via assignment)."""
    value: Expression | None = None


@dataclass
class OnErrorStmt(Statement):
    action: str = ""       # "GoTo", "ResumeNext", "GoTo0"
    label: str | None = None


@dataclass
class LabelStmt(Statement):
    name: str = ""


# ---------------------------------------------------------------------------
# Expressions
# ---------------------------------------------------------------------------

@dataclass
class Literal(Expression):
    value: Any = None
    lit_type: str = ""  # "integer", "float", "string", "boolean", "date", "nothing"


@dataclass
class Identifier(Expression):
    name: str = ""


@dataclass
class MemberAccess(Expression):
    target: Expression | None = None
    member: str = ""


@dataclass
class IndexAccess(Expression):
    """Array/dictionary access **and** function-call syntax (VBA is ambiguous)."""
    target: Expression | None = None
    args: list[Expression] = field(default_factory=list)


@dataclass
class BinaryOp(Expression):
    left: Expression | None = None
    op: str = ""
    right: Expression | None = None


@dataclass
class UnaryOp(Expression):
    op: str = ""
    operand: Expression | None = None


@dataclass
class NewExpr(Expression):
    type_name: str = ""


@dataclass
class TypeOfExpr(Expression):
    expr: Expression | None = None
    type_name: str = ""


@dataclass
class DotAccess(Expression):
    """Shorthand member access inside ``With`` blocks (e.g. ``.Name``)."""
    member: str = ""


# ---------------------------------------------------------------------------
# Additional statements
# ---------------------------------------------------------------------------

@dataclass
class ReDimStmt(Statement):
    name: str = ""
    dimensions: list[Expression] = field(default_factory=list)
    preserve: bool = False


@dataclass
class GoToStmt(Statement):
    label: str = ""


@dataclass
class EraseStmt(Statement):
    arrays: list[str] = field(default_factory=list)


@dataclass
class OptionStmt(Statement):
    option: str = ""  # "Explicit", "Base 0", "Base 1", "Compare Text", etc.
