"""Recursive descent parser for VBA source code.

Consumes tokens produced by :func:`vba2py.lexer.tokenize` and builds a typed
AST using the node classes from :mod:`vba2py.ast_nodes`.
"""

from __future__ import annotations

from vba2py.lexer import Token, TokenType, tokenize
from vba2py.ast_nodes import (
    Module, VarDecl, ConstDecl, Sub, Function, Param,
    AssignStmt, SetStmt, CallStmt, IfStmt, CaseClause, CaseRange, CaseIs, SelectCaseStmt,
    ForStmt, ForEachStmt, DoWhileStmt, DoUntilStmt, WhileStmt,
    WithStmt, DimStmt, ExitStmt, ReturnStmt, OnErrorStmt, LabelStmt,
    ReDimStmt, GoToStmt, EraseStmt, OptionStmt,
    Literal, Identifier, MemberAccess, IndexAccess, BinaryOp, UnaryOp,
    NewExpr, DotAccess, Expression, Statement,
)

# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------


class ParseError(Exception):
    """Raised when the parser encounters unexpected input."""

    def __init__(self, message: str, token: Token) -> None:
        self.token = token
        super().__init__(
            f"Line {token.lineno}, Col {token.col}: {message} "
            f"(got {token.type.name} {token.value!r})"
        )


# ---------------------------------------------------------------------------
# Access modifier helpers
# ---------------------------------------------------------------------------

_ACCESS_KEYWORDS = {TokenType.KW_PUBLIC, TokenType.KW_PRIVATE, TokenType.KW_STATIC}


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------


class Parser:
    """Recursive-descent parser for VBA source code."""

    def __init__(self, tokens: list[Token]) -> None:
        self.tokens = tokens
        self.pos = 0

    # -- Helpers ------------------------------------------------------------

    def peek(self) -> Token:
        """Return current token without consuming it."""
        return self.tokens[self.pos]

    def advance(self) -> Token:
        """Consume and return current token."""
        tok = self.tokens[self.pos]
        if tok.type is not TokenType.EOF:
            self.pos += 1
        return tok

    def expect(self, tt: TokenType) -> Token:
        """Consume a token of type *tt* or raise :class:`ParseError`."""
        tok = self.peek()
        if tok.type is not tt:
            raise ParseError(f"Expected {tt.name}", tok)
        return self.advance()

    def match(self, *types: TokenType) -> Token | None:
        """If the current token matches any of *types*, consume and return it."""
        if self.peek().type in types:
            return self.advance()
        return None

    def at(self, *types: TokenType) -> bool:
        """Return ``True`` if the current token is one of *types*."""
        return self.peek().type in types

    def skip_newlines(self) -> None:
        """Skip zero or more NEWLINE tokens."""
        while self.peek().type is TokenType.NEWLINE:
            self.advance()

    def expect_end_of_statement(self) -> None:
        """Expect NEWLINE, COLON, or EOF (statement terminator)."""
        if self.at(TokenType.NEWLINE, TokenType.COLON, TokenType.EOF):
            if self.at(TokenType.NEWLINE, TokenType.COLON):
                self.advance()
            return
        raise ParseError("Expected end of statement", self.peek())

    def _lookahead(self, offset: int = 1) -> Token:
        """Return a token *offset* positions ahead (0 = current)."""
        idx = self.pos + offset
        if idx >= len(self.tokens):
            return self.tokens[-1]  # EOF
        return self.tokens[idx]

    # -- Entry point --------------------------------------------------------

    def parse(self) -> Module:
        """Parse the full token stream into a :class:`Module`."""
        mod = Module()
        self.skip_newlines()

        while not self.at(TokenType.EOF):
            # Access modifier
            if self.at(*_ACCESS_KEYWORDS):
                next_tok = self._lookahead()
                if next_tok.type is TokenType.KW_SUB:
                    mod.procedures.append(self.parse_sub())
                elif next_tok.type is TokenType.KW_FUNCTION:
                    mod.procedures.append(self.parse_function())
                elif next_tok.type is TokenType.KW_STATIC:
                    # e.g. Public Static Sub ...
                    mod.procedures.append(self._parse_procedure_with_access())
                else:
                    # Module-level Dim/Const with access modifier
                    stmt = self.parse_statement()
                    mod.declarations.append(stmt)
            elif self.at(TokenType.KW_SUB):
                mod.procedures.append(self.parse_sub())
            elif self.at(TokenType.KW_FUNCTION):
                mod.procedures.append(self.parse_function())
            elif self.at(TokenType.KW_DIM, TokenType.KW_CONST):
                mod.declarations.append(self.parse_statement())
            else:
                # Other module-level statements (Option Explicit, etc.)
                mod.declarations.append(self.parse_statement())
            self.skip_newlines()

        return mod

    # -- Procedures ---------------------------------------------------------

    def _parse_access(self) -> tuple[str, bool]:
        """Consume optional access modifier and Static keyword.

        Returns ``(access, is_static)`` where *access* is one of
        ``"Public"`` / ``"Private"`` and *is_static* is a bool.
        """
        access = "Public"
        is_static = False
        if self.match(TokenType.KW_PUBLIC):
            access = "Public"
        elif self.match(TokenType.KW_PRIVATE):
            access = "Private"
        if self.match(TokenType.KW_STATIC):
            is_static = True  # noqa: F841  – reserved for future use
        return access, is_static

    def _parse_procedure_with_access(self) -> Sub | Function:
        """Handle cases like ``Public Static Sub ...``."""
        if self.at(TokenType.KW_SUB):
            return self.parse_sub()
        elif self.at(TokenType.KW_FUNCTION):
            return self.parse_function()
        # Access keyword already consumed by caller; grab next
        access, _ = self._parse_access()
        if self.at(TokenType.KW_SUB):
            return self.parse_sub()
        return self.parse_function()

    def parse_sub(self) -> Sub:
        """Parse a ``Sub … End Sub`` block."""
        lineno = self.peek().lineno
        access, _ = self._parse_access()
        self.expect(TokenType.KW_SUB)
        name_tok = self.expect(TokenType.IDENTIFIER)
        params = self._parse_param_list()
        self.expect_end_of_statement()
        self.skip_newlines()
        body = self._parse_body(end_check=self._is_end_sub)
        # Consume "End Sub"
        self.expect(TokenType.KW_END)
        self.expect(TokenType.KW_SUB)
        node = Sub(name=name_tok.value, params=params, body=body, access=access)
        node.lineno = lineno
        return node

    def parse_function(self) -> Function:
        """Parse a ``Function … End Function`` block."""
        lineno = self.peek().lineno
        access, _ = self._parse_access()
        self.expect(TokenType.KW_FUNCTION)
        name_tok = self.expect(TokenType.IDENTIFIER)
        params = self._parse_param_list()
        return_type: str | None = None
        if self.match(TokenType.KW_AS):
            return_type = self.expect(TokenType.IDENTIFIER).value
        self.expect_end_of_statement()
        self.skip_newlines()
        body = self._parse_body(end_check=self._is_end_function)
        self.expect(TokenType.KW_END)
        self.expect(TokenType.KW_FUNCTION)
        node = Function(
            name=name_tok.value, params=params, body=body,
            return_type=return_type, access=access,
        )
        node.lineno = lineno
        return node

    # -- Parameters ---------------------------------------------------------

    def _parse_param_list(self) -> list[Param]:
        """Parse ``(param, param, …)`` or nothing if no parens."""
        if not self.match(TokenType.LPAREN):
            return []
        params: list[Param] = []
        if not self.at(TokenType.RPAREN):
            params.append(self._parse_param())
            while self.match(TokenType.COMMA):
                params.append(self._parse_param())
        self.expect(TokenType.RPAREN)
        return params

    def _parse_param(self) -> Param:
        lineno = self.peek().lineno
        optional = bool(self.match(TokenType.KW_OPTIONAL))
        passing = "ByRef"
        if self.match(TokenType.KW_BYVAL):
            passing = "ByVal"
        elif self.match(TokenType.KW_BYREF):
            passing = "ByRef"
        name_tok = self.expect(TokenType.IDENTIFIER)
        param_type: str | None = None
        if self.match(TokenType.KW_AS):
            param_type = self.expect(TokenType.IDENTIFIER).value
        default: Expression | None = None
        if self.match(TokenType.EQ):
            default = self.parse_expression()
        p = Param(
            name=name_tok.value, param_type=param_type,
            passing=passing, optional=optional, default=default,
        )
        p.lineno = lineno
        return p

    # -- Body parsing helpers -----------------------------------------------

    def _is_end_sub(self) -> bool:
        return (
            self.at(TokenType.KW_END)
            and self._lookahead().type is TokenType.KW_SUB
        )

    def _is_end_function(self) -> bool:
        return (
            self.at(TokenType.KW_END)
            and self._lookahead().type is TokenType.KW_FUNCTION
        )

    def _is_end_if(self) -> bool:
        return (
            (self.at(TokenType.KW_END) and self._lookahead().type is TokenType.KW_IF)
            or self.at(TokenType.KW_ELSE, TokenType.KW_ELSEIF)
        )

    def _is_end_select(self) -> bool:
        return (
            (self.at(TokenType.KW_END) and self._lookahead().type is TokenType.KW_SELECT)
            or self.at(TokenType.KW_CASE)
        )

    def _is_end_with(self) -> bool:
        return (
            self.at(TokenType.KW_END)
            and self._lookahead().type is TokenType.KW_WITH
        )

    def _parse_body(self, end_check) -> list[Statement]:  # noqa: ANN001
        """Parse statements until *end_check()* returns True or EOF."""
        stmts: list[Statement] = []
        while not self.at(TokenType.EOF) and not end_check():
            self.skip_newlines()
            if self.at(TokenType.EOF) or end_check():
                break
            stmts.append(self.parse_statement())
            # Consume statement separator (newline or colon) if present
            if self.at(TokenType.NEWLINE, TokenType.COLON):
                self.advance()
        self.skip_newlines()
        return stmts

    # -- Statements ---------------------------------------------------------

    def parse_statement(self) -> Statement:
        """Dispatch to the appropriate statement parser."""
        tok = self.peek()
        tt = tok.type

        if tt is TokenType.KW_DIM:
            return self.parse_dim()
        if tt is TokenType.KW_CONST:
            return self._parse_const()
        if tt is TokenType.KW_IF:
            return self.parse_if()
        if tt is TokenType.KW_SELECT:
            return self.parse_select_case()
        if tt is TokenType.KW_FOR:
            return self.parse_for()
        if tt is TokenType.KW_DO:
            return self.parse_do()
        if tt is TokenType.KW_WHILE:
            return self.parse_while()
        if tt is TokenType.KW_WITH:
            return self.parse_with()
        if tt is TokenType.KW_SET:
            return self._parse_set()
        if tt is TokenType.KW_LET:
            return self._parse_let()
        if tt is TokenType.KW_CALL:
            return self._parse_call()
        if tt is TokenType.KW_EXIT:
            return self.parse_exit()
        if tt is TokenType.KW_REDIM:
            return self._parse_redim()
        if tt is TokenType.KW_GOTO:
            return self._parse_goto()
        if tt is TokenType.KW_ERASE:
            return self._parse_erase()
        if tt is TokenType.KW_ON:
            return self.parse_on_error()
        if tt is TokenType.KW_DEBUG:
            return self._parse_debug()
        if tt in _ACCESS_KEYWORDS:
            # Module-level access modifier before Dim/Const
            self.advance()  # consume access keyword
            if self.at(TokenType.KW_CONST):
                return self._parse_const()
            if self.at(TokenType.KW_DIM):
                return self.parse_dim()
            raise ParseError("Expected Dim or Const after access modifier", self.peek())
        if tt is TokenType.DOT:
            # .Member inside With block — treat as assignment or call
            return self._parse_dot_statement()
        if tt is TokenType.IDENTIFIER:
            return self._parse_identifier_statement()

        raise ParseError("Unexpected token at start of statement", tok)

    # -- Dim ----------------------------------------------------------------

    def parse_dim(self) -> DimStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_DIM)
        decls: list[VarDecl] = [self._parse_var_decl()]
        while self.match(TokenType.COMMA):
            decls.append(self._parse_var_decl())
        node = DimStmt(declarations=decls)
        node.lineno = lineno
        return node

    def _parse_var_decl(self) -> VarDecl:
        lineno = self.peek().lineno
        name_tok = self.expect(TokenType.IDENTIFIER)
        is_array = False
        if self.match(TokenType.LPAREN):
            # Array declaration — skip dimension specifiers
            if not self.at(TokenType.RPAREN):
                self.parse_expression()
                while self.match(TokenType.COMMA):
                    self.parse_expression()
            self.expect(TokenType.RPAREN)
            is_array = True
        var_type: str | None = None
        if self.match(TokenType.KW_AS):
            # Handle "New ClassName"
            if self.match(TokenType.KW_NEW):
                var_type = "New " + self.expect(TokenType.IDENTIFIER).value
            else:
                var_type = self.expect(TokenType.IDENTIFIER).value
        vd = VarDecl(name=name_tok.value, var_type=var_type, is_array=is_array)
        vd.lineno = lineno
        return vd

    # -- Const --------------------------------------------------------------

    def _parse_const(self) -> ConstDecl:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_CONST)
        name_tok = self.expect(TokenType.IDENTIFIER)
        var_type: str | None = None
        if self.match(TokenType.KW_AS):
            var_type = self.expect(TokenType.IDENTIFIER).value
        self.expect(TokenType.EQ)
        value = self.parse_expression()
        node = ConstDecl(name=name_tok.value, var_type=var_type, value=value)
        node.lineno = lineno
        return node

    # -- If -----------------------------------------------------------------

    def parse_if(self) -> IfStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_IF)
        condition = self.parse_expression()
        self.expect(TokenType.KW_THEN)

        # Single-line vs multi-line: if THEN is followed by NEWLINE → multi
        if self.at(TokenType.NEWLINE, TokenType.EOF):
            return self._parse_multiline_if(condition, lineno)
        return self._parse_singleline_if(condition, lineno)

    def _parse_singleline_if(self, condition: Expression, lineno: int) -> IfStmt:
        then_body = [self.parse_statement()]
        else_body: list[Statement] | None = None
        if self.match(TokenType.KW_ELSE):
            else_body = [self.parse_statement()]
        node = IfStmt(
            condition=condition, then_body=then_body,
            else_body=else_body,
        )
        node.lineno = lineno
        return node

    def _parse_multiline_if(self, condition: Expression, lineno: int) -> IfStmt:
        self.skip_newlines()
        then_body = self._parse_body(end_check=self._is_end_if)
        elseif_clauses: list[tuple[Expression, list[Statement]]] = []
        else_body: list[Statement] | None = None

        while self.match(TokenType.KW_ELSEIF):
            ei_cond = self.parse_expression()
            self.expect(TokenType.KW_THEN)
            self.skip_newlines()
            ei_body = self._parse_body(end_check=self._is_end_if)
            elseif_clauses.append((ei_cond, ei_body))

        if self.match(TokenType.KW_ELSE):
            self.skip_newlines()
            else_body = self._parse_body(
                end_check=lambda: (
                    self.at(TokenType.KW_END)
                    and self._lookahead().type is TokenType.KW_IF
                )
            )

        self.expect(TokenType.KW_END)
        self.expect(TokenType.KW_IF)
        node = IfStmt(
            condition=condition, then_body=then_body,
            elseif_clauses=elseif_clauses, else_body=else_body,
        )
        node.lineno = lineno
        return node

    # -- Select Case --------------------------------------------------------

    def parse_select_case(self) -> SelectCaseStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_SELECT)
        self.expect(TokenType.KW_CASE)
        expr = self.parse_expression()
        self.expect_end_of_statement()
        self.skip_newlines()

        cases: list[CaseClause] = []
        else_body: list[Statement] | None = None

        while self.at(TokenType.KW_CASE):
            self.advance()  # consume Case
            # Case Else
            if self.match(TokenType.KW_ELSE):
                self.skip_newlines()
                else_body = self._parse_body(end_check=self._is_end_select)
                break
            # Regular Case item, item, … (expr | expr To expr | Is <op> expr)
            exprs: list = [self._parse_case_item()]
            while self.match(TokenType.COMMA):
                exprs.append(self._parse_case_item())
            self.expect_end_of_statement()
            self.skip_newlines()
            body = self._parse_body(end_check=self._is_end_select)
            clause = CaseClause(expressions=exprs, body=body)
            clause.lineno = lineno
            cases.append(clause)

        self.expect(TokenType.KW_END)
        self.expect(TokenType.KW_SELECT)
        node = SelectCaseStmt(expr=expr, cases=cases, else_body=else_body)
        node.lineno = lineno
        return node

    _CASE_IS_OPS = {
        TokenType.EQ: "=",
        TokenType.NEQ: "<>",
        TokenType.LT: "<",
        TokenType.GT: ">",
        TokenType.LTE: "<=",
        TokenType.GTE: ">=",
    }

    def _parse_case_item(self):
        """Parse one item of a Case clause: expr | expr To expr | Is <op> expr."""
        if self.at(TokenType.KW_IS):
            self.advance()
            tok = self.peek()
            op = self._CASE_IS_OPS.get(tok.type)
            if op is None:
                raise ParseError("Expected comparison operator after 'Case Is'", tok)
            self.advance()
            return CaseIs(op=op, value=self.parse_expression())
        expr = self.parse_expression()
        if self.match(TokenType.KW_TO):
            return CaseRange(lo=expr, hi=self.parse_expression())
        return expr

    # -- For / For Each -----------------------------------------------------

    def parse_for(self) -> ForStmt | ForEachStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_FOR)

        if self.match(TokenType.KW_EACH):
            return self._parse_for_each(lineno)
        return self._parse_for_next(lineno)

    def _parse_for_each(self, lineno: int) -> ForEachStmt:
        var = self.expect(TokenType.IDENTIFIER).value
        self.expect(TokenType.KW_IN)
        collection = self.parse_expression()
        self.expect_end_of_statement()
        self.skip_newlines()
        body = self._parse_body(
            end_check=lambda: self.at(TokenType.KW_NEXT)
        )
        self.expect(TokenType.KW_NEXT)
        self.match(TokenType.IDENTIFIER)  # optional variable after Next
        node = ForEachStmt(var=var, collection=collection, body=body)
        node.lineno = lineno
        return node

    def _parse_for_next(self, lineno: int) -> ForStmt:
        var = self.expect(TokenType.IDENTIFIER).value
        self.expect(TokenType.EQ)
        start = self.parse_expression()
        self.expect(TokenType.KW_TO)
        end = self.parse_expression()
        step: Expression | None = None
        if self.match(TokenType.KW_STEP):
            step = self.parse_expression()
        self.expect_end_of_statement()
        self.skip_newlines()
        body = self._parse_body(
            end_check=lambda: self.at(TokenType.KW_NEXT)
        )
        self.expect(TokenType.KW_NEXT)
        self.match(TokenType.IDENTIFIER)  # optional variable after Next
        node = ForStmt(var=var, start=start, end=end, step=step, body=body)
        node.lineno = lineno
        return node

    # -- Do -----------------------------------------------------------------

    def parse_do(self) -> DoWhileStmt | DoUntilStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_DO)

        # Pre-check: Do While / Do Until
        if self.match(TokenType.KW_WHILE):
            cond = self.parse_expression()
            self.expect_end_of_statement()
            self.skip_newlines()
            body = self._parse_body(end_check=lambda: self.at(TokenType.KW_LOOP))
            self.expect(TokenType.KW_LOOP)
            node = DoWhileStmt(condition=cond, body=body, check_post=False)
            node.lineno = lineno
            return node

        if self.match(TokenType.KW_UNTIL):
            cond = self.parse_expression()
            self.expect_end_of_statement()
            self.skip_newlines()
            body = self._parse_body(end_check=lambda: self.at(TokenType.KW_LOOP))
            self.expect(TokenType.KW_LOOP)
            node = DoUntilStmt(condition=cond, body=body, check_post=False)
            node.lineno = lineno
            return node

        # No pre-check → post-check or infinite loop
        self.expect_end_of_statement()
        self.skip_newlines()
        body = self._parse_body(end_check=lambda: self.at(TokenType.KW_LOOP))
        self.expect(TokenType.KW_LOOP)

        if self.match(TokenType.KW_WHILE):
            cond = self.parse_expression()
            node = DoWhileStmt(condition=cond, body=body, check_post=True)
            node.lineno = lineno
            return node
        if self.match(TokenType.KW_UNTIL):
            cond = self.parse_expression()
            node = DoUntilStmt(condition=cond, body=body, check_post=True)
            node.lineno = lineno
            return node

        # Do ... Loop (infinite)
        node = DoWhileStmt(condition=Literal(True, "boolean"), body=body, check_post=False)
        node.lineno = lineno
        return node

    # -- While / Wend -------------------------------------------------------

    def parse_while(self) -> WhileStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_WHILE)
        cond = self.parse_expression()
        self.expect_end_of_statement()
        self.skip_newlines()
        body = self._parse_body(end_check=lambda: self.at(TokenType.KW_WEND))
        self.expect(TokenType.KW_WEND)
        node = WhileStmt(condition=cond, body=body)
        node.lineno = lineno
        return node

    # -- With ---------------------------------------------------------------

    def parse_with(self) -> WithStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_WITH)
        target = self.parse_expression()
        self.expect_end_of_statement()
        self.skip_newlines()
        body = self._parse_body(end_check=self._is_end_with)
        self.expect(TokenType.KW_END)
        self.expect(TokenType.KW_WITH)
        node = WithStmt(target=target, body=body)
        node.lineno = lineno
        return node

    # -- Set ----------------------------------------------------------------

    def _parse_set(self) -> SetStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_SET)
        target = self._parse_lhs_expression()
        self.expect(TokenType.EQ)
        value = self.parse_expression()
        node = SetStmt(target=target, value=value)
        node.lineno = lineno
        return node

    # -- Let ----------------------------------------------------------------

    def _parse_let(self) -> AssignStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_LET)
        target = self.parse_expression()
        self.expect(TokenType.EQ)
        value = self.parse_expression()
        node = AssignStmt(target=target, value=value)
        node.lineno = lineno
        return node

    # -- Call ---------------------------------------------------------------

    def _parse_call(self) -> CallStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_CALL)
        name_tok = self.expect(TokenType.IDENTIFIER)
        name = name_tok.value
        # Handle member access: Call obj.Method(...)
        while self.match(TokenType.DOT):
            member = self.expect(TokenType.IDENTIFIER)
            name = name + "." + member.value
        args: list[Expression] = []
        if self.match(TokenType.LPAREN):
            if not self.at(TokenType.RPAREN):
                args.append(self.parse_expression())
                while self.match(TokenType.COMMA):
                    args.append(self.parse_expression())
            self.expect(TokenType.RPAREN)
        node = CallStmt(name=name, args=args)
        node.lineno = lineno
        return node

    # -- Exit ---------------------------------------------------------------

    def parse_exit(self) -> ExitStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_EXIT)
        tok = self.peek()
        if tok.type is TokenType.KW_SUB:
            self.advance()
            kind = "Sub"
        elif tok.type is TokenType.KW_FUNCTION:
            self.advance()
            kind = "Function"
        elif tok.type is TokenType.KW_FOR:
            self.advance()
            kind = "For"
        elif tok.type is TokenType.KW_DO:
            self.advance()
            kind = "Do"
        else:
            raise ParseError("Expected Sub, Function, For, or Do after Exit", tok)
        node = ExitStmt(kind=kind)
        node.lineno = lineno
        return node

    # -- On Error -----------------------------------------------------------

    def parse_on_error(self) -> OnErrorStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_ON)
        self.expect(TokenType.KW_ERROR)

        if self.match(TokenType.KW_RESUME):
            # On Error Resume Next
            self.expect(TokenType.KW_NEXT)
            node = OnErrorStmt(action="ResumeNext")
            node.lineno = lineno
            return node

        self.expect(TokenType.KW_GOTO)
        label_tok = self.peek()
        if label_tok.type is TokenType.INTEGER and label_tok.value == "0":
            self.advance()
            node = OnErrorStmt(action="GoTo0")
            node.lineno = lineno
            return node

        # On Error GoTo <label>
        label = self.expect(TokenType.IDENTIFIER).value
        node = OnErrorStmt(action="GoTo", label=label)
        node.lineno = lineno
        return node

    # -- Debug.Print --------------------------------------------------------

    def _parse_debug(self) -> CallStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_DEBUG)
        self.expect(TokenType.DOT)
        self.expect(TokenType.KW_PRINT)
        args: list[Expression] = []
        if not self.at(TokenType.NEWLINE, TokenType.COLON, TokenType.EOF):
            args.append(self.parse_expression())
            while self.match(TokenType.COMMA, TokenType.SEMICOLON):
                if self.at(TokenType.NEWLINE, TokenType.COLON, TokenType.EOF):
                    break
                args.append(self.parse_expression())
        node = CallStmt(name="Debug.Print", args=args)
        node.lineno = lineno
        return node

    # -- ReDim --------------------------------------------------------------

    def _parse_redim(self) -> ReDimStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_REDIM)
        preserve = bool(self.match(TokenType.KW_PRESERVE))
        name_tok = self.expect(TokenType.IDENTIFIER)
        dims: list[Expression] = []
        if self.match(TokenType.LPAREN):
            if not self.at(TokenType.RPAREN):
                dims.append(self.parse_expression())
                while self.match(TokenType.COMMA):
                    dims.append(self.parse_expression())
            self.expect(TokenType.RPAREN)
        node = ReDimStmt(name=name_tok.value, dimensions=dims, preserve=preserve)
        node.lineno = lineno
        return node

    # -- GoTo ---------------------------------------------------------------

    def _parse_goto(self) -> GoToStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_GOTO)
        label_tok = self.expect(TokenType.IDENTIFIER)
        node = GoToStmt(label=label_tok.value)
        node.lineno = lineno
        return node

    # -- Erase --------------------------------------------------------------

    def _parse_erase(self) -> EraseStmt:
        lineno = self.peek().lineno
        self.expect(TokenType.KW_ERASE)
        arrays: list[str] = [self.expect(TokenType.IDENTIFIER).value]
        while self.match(TokenType.COMMA):
            arrays.append(self.expect(TokenType.IDENTIFIER).value)
        node = EraseStmt(arrays=arrays)
        node.lineno = lineno
        return node

    # -- Identifier-led statements ------------------------------------------

    def _parse_lhs_expression(self) -> Expression:
        """Parse left-hand side of an assignment or implicit call.

        This is like ``parse_expression`` but stops *before* consuming ``=``
        so that the caller can distinguish assignment from comparison.
        We parse up through concatenation (``&``) but not comparisons.
        """
        return self._parse_concat()

    def _parse_identifier_statement(self) -> Statement:
        """Parse a statement that begins with an IDENTIFIER.

        Could be assignment, implicit call, or label.
        """
        lineno = self.peek().lineno

        # Check for label:  name: (followed by NEWLINE / stmt / EOF)
        if (
            self._lookahead().type is TokenType.COLON
            and self._lookahead(2).type
            in (TokenType.NEWLINE, TokenType.EOF, TokenType.IDENTIFIER,
                TokenType.KW_DIM, TokenType.KW_IF, TokenType.KW_FOR,
                TokenType.KW_DO, TokenType.KW_WHILE, TokenType.KW_SELECT,
                TokenType.KW_SET, TokenType.KW_CALL, TokenType.KW_EXIT,
                TokenType.KW_ON, TokenType.KW_WITH, TokenType.KW_DEBUG)
        ):
            name_tok = self.advance()
            self.advance()  # consume colon
            node = LabelStmt(name=name_tok.value)
            node.lineno = lineno
            return node

        # Option Explicit / Option Base 0|1 / Option Compare Text|Binary
        if self.peek().value.lower() == "option":
            self.advance()  # consume "Option"
            parts: list[str] = []
            while not self.at(TokenType.NEWLINE, TokenType.COLON, TokenType.EOF):
                parts.append(self.advance().value)
            node = OptionStmt(option=" ".join(parts))
            node.lineno = lineno
            return node

        # Parse the left-hand side — stop before '=' so we can treat
        # it as assignment rather than comparison.
        target = self._parse_lhs_expression()

        # Assignment: target = value
        if self.match(TokenType.EQ):
            value = self.parse_expression()
            node = AssignStmt(target=target, value=value)
            node.lineno = lineno
            return node

        # If the expression is just an Identifier or MemberAccess, treat as
        # implicit call with remaining tokens as arguments.
        if isinstance(target, (Identifier, MemberAccess)):
            name = self._expr_to_call_name(target)
            args: list[Expression] = []
            if not self.at(TokenType.NEWLINE, TokenType.COLON, TokenType.EOF):
                args.append(self.parse_expression())
                while self.match(TokenType.COMMA):
                    args.append(self.parse_expression())
            node = CallStmt(name=name, args=args)
            node.lineno = lineno
            return node

        # IndexAccess that was parsed (e.g. ``MyFunc(1, 2)``) → CallStmt
        if isinstance(target, IndexAccess):
            name = self._expr_to_call_name(target.target)
            node = CallStmt(name=name, args=target.args)
            node.lineno = lineno
            return node

        raise ParseError("Cannot interpret identifier statement", self.peek())

    def _parse_dot_statement(self) -> Statement:
        """Parse a statement starting with ``.Member`` (inside With block)."""
        lineno = self.peek().lineno
        target = self._parse_lhs_expression()
        if self.match(TokenType.EQ):
            value = self.parse_expression()
            node = AssignStmt(target=target, value=value)
            node.lineno = lineno
            return node
        # Implicit call: .Method arg1, arg2
        if isinstance(target, DotAccess):
            name = "." + target.member
            args: list[Expression] = []
            if not self.at(TokenType.NEWLINE, TokenType.COLON, TokenType.EOF):
                args.append(self.parse_expression())
                while self.match(TokenType.COMMA):
                    args.append(self.parse_expression())
            node = CallStmt(name=name, args=args)
            node.lineno = lineno
            return node
        raise ParseError("Cannot interpret dot statement", self.peek())

    @staticmethod
    def _expr_to_call_name(expr: Expression) -> str:
        """Convert an Identifier or MemberAccess chain to a dotted name string."""
        if isinstance(expr, Identifier):
            return expr.name
        if isinstance(expr, MemberAccess):
            parts: list[str] = []
            node = expr
            while isinstance(node, MemberAccess):
                parts.append(node.member)
                node = node.target
            if isinstance(node, Identifier):
                parts.append(node.name)
            parts.reverse()
            return ".".join(parts)
        return str(expr)

    # ===================================================================
    # EXPRESSION PARSING — precedence climbing via recursive functions
    # ===================================================================

    def parse_expression(self) -> Expression:
        """Top-level expression entry point."""
        return self._parse_xor()

    # -- Xor ----------------------------------------------------------------
    def _parse_xor(self) -> Expression:
        left = self._parse_or()
        while self.match(TokenType.KW_XOR):
            right = self._parse_or()
            left = BinaryOp(left=left, op="Xor", right=right)
        return left

    # -- Or -----------------------------------------------------------------
    def _parse_or(self) -> Expression:
        left = self._parse_and()
        while self.match(TokenType.KW_OR):
            right = self._parse_and()
            left = BinaryOp(left=left, op="Or", right=right)
        return left

    # -- And ----------------------------------------------------------------
    def _parse_and(self) -> Expression:
        left = self._parse_not()
        while self.match(TokenType.KW_AND):
            right = self._parse_not()
            left = BinaryOp(left=left, op="And", right=right)
        return left

    # -- Not (unary) --------------------------------------------------------
    def _parse_not(self) -> Expression:
        if self.match(TokenType.KW_NOT):
            operand = self._parse_not()
            return UnaryOp(op="Not", operand=operand)
        return self._parse_comparison()

    # -- Comparisons --------------------------------------------------------
    _CMP_OPS = {
        TokenType.EQ: "=",
        TokenType.NEQ: "<>",
        TokenType.LT: "<",
        TokenType.GT: ">",
        TokenType.LTE: "<=",
        TokenType.GTE: ">=",
        TokenType.KW_IS: "Is",
        TokenType.KW_LIKE: "Like",
    }

    def _parse_comparison(self) -> Expression:
        left = self._parse_concat()
        while self.peek().type in self._CMP_OPS:
            op = self._CMP_OPS[self.advance().type]
            right = self._parse_concat()
            left = BinaryOp(left=left, op=op, right=right)
        return left

    # -- Concatenation (&) --------------------------------------------------
    def _parse_concat(self) -> Expression:
        left = self._parse_additive()
        while self.match(TokenType.AMPERSAND):
            right = self._parse_additive()
            left = BinaryOp(left=left, op="&", right=right)
        return left

    # -- Additive (+, -) ----------------------------------------------------
    def _parse_additive(self) -> Expression:
        left = self._parse_mod()
        while self.at(TokenType.PLUS, TokenType.MINUS):
            op = self.advance().value
            right = self._parse_mod()
            left = BinaryOp(left=left, op=op, right=right)
        return left

    # -- Mod ----------------------------------------------------------------
    def _parse_mod(self) -> Expression:
        left = self._parse_int_div()
        while self.match(TokenType.KW_MOD):
            right = self._parse_int_div()
            left = BinaryOp(left=left, op="Mod", right=right)
        return left

    # -- Integer division (\) -----------------------------------------------
    def _parse_int_div(self) -> Expression:
        left = self._parse_multiplicative()
        while self.match(TokenType.BACKSLASH):
            right = self._parse_multiplicative()
            left = BinaryOp(left=left, op="\\", right=right)
        return left

    # -- Multiplicative (*, /) ----------------------------------------------
    def _parse_multiplicative(self) -> Expression:
        left = self._parse_unary()
        while self.at(TokenType.STAR, TokenType.SLASH):
            op = self.advance().value
            right = self._parse_unary()
            left = BinaryOp(left=left, op=op, right=right)
        return left

    # -- Unary (-, +) -------------------------------------------------------
    def _parse_unary(self) -> Expression:
        if self.at(TokenType.MINUS, TokenType.PLUS):
            op = self.advance().value
            operand = self._parse_unary()
            return UnaryOp(op=op, operand=operand)
        return self._parse_exponent()

    # -- Exponentiation (^) -------------------------------------------------
    def _parse_exponent(self) -> Expression:
        base = self._parse_postfix()
        if self.match(TokenType.CARET):
            # Right-associative
            exp = self._parse_unary()
            return BinaryOp(left=base, op="^", right=exp)
        return base

    # -- Postfix: member access, index, bang --------------------------------
    def _parse_postfix(self) -> Expression:
        expr = self._parse_primary()
        while True:
            if self.match(TokenType.DOT):
                member_tok = self.expect(TokenType.IDENTIFIER)
                expr = MemberAccess(target=expr, member=member_tok.value)
            elif self.at(TokenType.LPAREN):
                self.advance()
                args: list[Expression] = []
                if not self.at(TokenType.RPAREN):
                    args.append(self.parse_expression())
                    while self.match(TokenType.COMMA):
                        args.append(self.parse_expression())
                self.expect(TokenType.RPAREN)
                expr = IndexAccess(target=expr, args=args)
            elif self.match(TokenType.BANG):
                member_tok = self.expect(TokenType.IDENTIFIER)
                expr = MemberAccess(target=expr, member=member_tok.value)
            else:
                break
        return expr

    # -- Primary expressions ------------------------------------------------
    def _parse_primary(self) -> Expression:
        tok = self.peek()

        if tok.type is TokenType.INTEGER:
            self.advance()
            val = tok.value
            # Handle hex/octal literals
            if val.startswith("&") or val.startswith("&"):
                low = val.lower()
                if "h" in low:
                    num = int(val.lstrip("&").lstrip("Hh").rstrip("&"), 16)
                else:
                    num = int(val.lstrip("&").lstrip("Oo").rstrip("&"), 8)
            else:
                num = int(val)
            node = Literal(value=num, lit_type="integer")
            node.lineno = tok.lineno
            return node

        if tok.type is TokenType.FLOAT:
            self.advance()
            node = Literal(value=float(tok.value), lit_type="float")
            node.lineno = tok.lineno
            return node

        if tok.type is TokenType.STRING:
            self.advance()
            node = Literal(value=tok.value, lit_type="string")
            node.lineno = tok.lineno
            return node

        if tok.type is TokenType.DATE_LITERAL:
            self.advance()
            node = Literal(value=tok.value, lit_type="date")
            node.lineno = tok.lineno
            return node

        if tok.type is TokenType.KW_TRUE:
            self.advance()
            node = Literal(value=True, lit_type="boolean")
            node.lineno = tok.lineno
            return node

        if tok.type is TokenType.KW_FALSE:
            self.advance()
            node = Literal(value=False, lit_type="boolean")
            node.lineno = tok.lineno
            return node

        if tok.type is TokenType.KW_NOTHING:
            self.advance()
            node = Literal(value=None, lit_type="nothing")
            node.lineno = tok.lineno
            return node

        if tok.type is TokenType.KW_NEW:
            self.advance()
            type_name = self.expect(TokenType.IDENTIFIER).value
            node = NewExpr(type_name=type_name)
            node.lineno = tok.lineno
            return node

        if tok.type is TokenType.KW_ME:
            self.advance()
            node = Identifier(name="Me")
            node.lineno = tok.lineno
            return node

        if tok.type is TokenType.DOT:
            # .Member inside With block
            self.advance()
            member_tok = self.expect(TokenType.IDENTIFIER)
            node = DotAccess(member=member_tok.value)
            node.lineno = tok.lineno
            return node

        if tok.type is TokenType.LPAREN:
            self.advance()
            expr = self.parse_expression()
            self.expect(TokenType.RPAREN)
            return expr

        if tok.type is TokenType.IDENTIFIER:
            self.advance()
            node = Identifier(name=tok.value)
            node.lineno = tok.lineno
            return node

        raise ParseError("Expected expression", tok)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def parse(source: str) -> Module:
    """Parse VBA source code and return an AST Module."""
    tokens = tokenize(source)
    return Parser(tokens).parse()
