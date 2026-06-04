"""Tests for VBA lexer."""

from vba2py.lexer import tokenize, TokenType


def tok_types(source: str) -> list[str]:
    return [t.type.name for t in tokenize(source) if t.type != TokenType.EOF]


def tok_values(source: str) -> list[str]:
    return [t.value for t in tokenize(source) if t.type != TokenType.EOF]


class TestKeywords:
    def test_sub(self):
        types = tok_types("Sub Foo()")
        assert types == ["KW_SUB", "IDENTIFIER", "LPAREN", "RPAREN"]

    def test_case_insensitive(self):
        types = tok_types("DIM x AS Integer")
        assert types == ["KW_DIM", "IDENTIFIER", "KW_AS", "IDENTIFIER"]

    def test_if_then_else(self):
        types = tok_types("If x Then y Else z")
        assert "KW_IF" in types
        assert "KW_THEN" in types
        assert "KW_ELSE" in types


class TestLiterals:
    def test_string(self):
        toks = tokenize('"Hello, World!"')
        assert toks[0].type == TokenType.STRING
        assert toks[0].value == "Hello, World!"

    def test_string_escaped_quote(self):
        toks = tokenize('"He said ""hi""!"')
        assert toks[0].value == 'He said "hi"!'

    def test_integer(self):
        toks = tokenize("42")
        assert toks[0].type == TokenType.INTEGER
        assert toks[0].value == "42"

    def test_float(self):
        toks = tokenize("3.14")
        assert toks[0].type == TokenType.FLOAT

    def test_hex(self):
        toks = tokenize("&HFF")
        assert toks[0].type == TokenType.INTEGER
        assert toks[0].value == "&HFF"

    def test_date(self):
        toks = tokenize("#12/31/2025#")
        assert toks[0].type == TokenType.DATE_LITERAL
        assert toks[0].value == "12/31/2025"


class TestOperators:
    def test_comparison(self):
        types = tok_types("x <> y")
        assert "NEQ" in types

    def test_all_ops(self):
        types = tok_types("+ - * / \\ ^ & = <> < > <= >=")
        assert "PLUS" in types
        assert "BACKSLASH" in types
        assert "CARET" in types


class TestNewlines:
    def test_collapse(self):
        types = tok_types("x\n\n\ny")
        assert types.count("NEWLINE") == 1

    def test_line_continuation(self):
        types = tok_types("x = 1 _\n  + 2")
        # The _ joins lines so no NEWLINE between 1 and +
        assert "NEWLINE" not in types

    def test_comment_skipped(self):
        vals = tok_values("x ' this is a comment")
        assert vals == ["x"]

    def test_colon_separator(self):
        types = tok_types("x = 1 : y = 2")
        assert "COLON" in types
