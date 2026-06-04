"""VBA tokenizer -- case-insensitive keyword matching, no external deps."""

from __future__ import annotations

import enum
import re
from dataclasses import dataclass


# ---------------------------------------------------------------------------
# Token types
# ---------------------------------------------------------------------------

class TokenType(enum.Enum):
    # --- Keywords (alpha-sorted for quick lookup) --------------------------
    KW_AND = "And"
    KW_AS = "As"
    KW_BYREF = "ByRef"
    KW_BYVAL = "ByVal"
    KW_CALL = "Call"
    KW_CASE = "Case"
    KW_CONST = "Const"
    KW_DEBUG = "Debug"
    KW_DIM = "Dim"
    KW_DO = "Do"
    KW_EACH = "Each"
    KW_ELSE = "Else"
    KW_ELSEIF = "ElseIf"
    KW_END = "End"
    KW_EQV = "Eqv"
    KW_ERASE = "Erase"
    KW_ERROR = "Error"
    KW_EXIT = "Exit"
    KW_FALSE = "False"
    KW_FOR = "For"
    KW_FRIEND = "Friend"
    KW_FUNCTION = "Function"
    KW_GOTO = "GoTo"
    KW_IF = "If"
    KW_IMP = "Imp"
    KW_IN = "In"
    KW_IS = "Is"
    KW_LET = "Let"
    KW_LIKE = "Like"
    KW_LOOP = "Loop"
    KW_ME = "Me"
    KW_MOD = "Mod"
    KW_NEW = "New"
    KW_NEXT = "Next"
    KW_NOT = "Not"
    KW_NOTHING = "Nothing"
    KW_ON = "On"
    KW_OPTIONAL = "Optional"
    KW_OR = "Or"
    KW_PARAMARRAY = "ParamArray"
    KW_PRESERVE = "Preserve"
    KW_PRINT = "Print"
    KW_PRIVATE = "Private"
    KW_PUBLIC = "Public"
    KW_REDIM = "ReDim"
    KW_REM = "Rem"
    KW_RESUME = "Resume"
    KW_SELECT = "Select"
    KW_SET = "Set"
    KW_STATIC = "Static"
    KW_STEP = "Step"
    KW_SUB = "Sub"
    KW_THEN = "Then"
    KW_TO = "To"
    KW_TRUE = "True"
    KW_TYPEOF = "TypeOf"
    KW_UNTIL = "Until"
    KW_WEND = "Wend"
    KW_WHILE = "While"
    KW_WITH = "With"
    KW_XOR = "Xor"

    # --- Literals / identifiers --------------------------------------------
    IDENTIFIER = "IDENTIFIER"
    INTEGER = "INTEGER"
    FLOAT = "FLOAT"
    STRING = "STRING"
    DATE_LITERAL = "DATE_LITERAL"

    # --- Structural ---------------------------------------------------------
    NEWLINE = "NEWLINE"
    EOF = "EOF"

    # --- Delimiters ---------------------------------------------------------
    LPAREN = "("
    RPAREN = ")"
    COMMA = ","
    COLON = ":"
    SEMICOLON = ";"
    DOT = "."
    BANG = "!"

    # --- Operators ----------------------------------------------------------
    PLUS = "+"
    MINUS = "-"
    STAR = "*"
    SLASH = "/"
    BACKSLASH = "\\"
    CARET = "^"
    AMPERSAND = "&"
    EQ = "="
    NEQ = "<>"
    LT = "<"
    GT = ">"
    LTE = "<="
    GTE = ">="
    COLON_EQ = ":="


# Build a case-insensitive keyword map: lowered text -> TokenType
_KEYWORD_MAP: dict[str, TokenType] = {
    tt.value.lower(): tt
    for tt in TokenType
    if tt.name.startswith("KW_")
}


# ---------------------------------------------------------------------------
# Token
# ---------------------------------------------------------------------------

@dataclass(slots=True)
class Token:
    type: TokenType
    value: str
    lineno: int
    col: int

    def __repr__(self) -> str:
        return f"Token({self.type.name}, {self.value!r}, L{self.lineno}:{self.col})"


# ---------------------------------------------------------------------------
# Lexer errors
# ---------------------------------------------------------------------------

class LexerError(Exception):
    def __init__(self, message: str, lineno: int, col: int) -> None:
        self.lineno = lineno
        self.col = col
        super().__init__(f"Line {lineno}, Col {col}: {message}")


# ---------------------------------------------------------------------------
# Tokenizer
# ---------------------------------------------------------------------------

# Regex for identifiers and keywords (VBA allows underscores and digits after first char).
_IDENT_RE = re.compile(r"[A-Za-z_]\w*[\$%&#!]?")
# Hex literal  &H...   (optional trailing &)
_HEX_RE = re.compile(r"&[Hh][0-9A-Fa-f]+&?")
# Octal literal &O...  (optional trailing &)
_OCT_RE = re.compile(r"&[Oo][0-7]+&?")
# Numeric (float or int) -- we match float first to avoid partial matches
_NUMBER_RE = re.compile(
    r"(?:\d+\.\d*(?:[Ee][+-]?\d+)?)"   # 1.0, 1., 1.0E2
    r"|(?:\.\d+(?:[Ee][+-]?\d+)?)"      # .5, .5E-1
    r"|(?:\d+[Ee][+-]?\d+)"             # 1E5
    r"|(?:\d+)"                          # 123
)


def tokenize(source: str) -> list[Token]:
    """Tokenize VBA *source* and return a flat list of ``Token`` objects.

    * Whitespace (except newlines) is skipped.
    * Multiple consecutive newlines are collapsed into one NEWLINE token.
    * Line continuations (`` _`` at end of line) join physical lines.
    * Comments (``'`` or ``Rem``) are discarded.
    """
    tokens: list[Token] = []
    pos = 0
    length = len(source)
    lineno = 1
    line_start = 0  # index of the first char of the current line

    def _col() -> int:
        return pos - line_start + 1

    def _add(tt: TokenType, val: str, col: int) -> None:
        tokens.append(Token(tt, val, lineno, col))

    while pos < length:
        ch = source[pos]

        # --- Whitespace (not newline) --------------------------------------
        if ch in (" ", "\t"):
            pos += 1
            continue

        # --- Line continuation: space(s) + _ + optional \r + \n ------------
        if ch == "_" and pos > 0 and source[pos - 1] in (" ", "\t"):
            # Look ahead for newline
            ahead = pos + 1
            while ahead < length and source[ahead] in (" ", "\t"):
                ahead += 1
            if ahead < length and source[ahead] in ("\r", "\n"):
                # Consume the _ , any whitespace, and the newline
                if source[ahead] == "\r" and ahead + 1 < length and source[ahead + 1] == "\n":
                    ahead += 2
                else:
                    ahead += 1
                lineno += 1
                line_start = ahead
                pos = ahead
                continue
            # Otherwise fall through -- treat _ as part of an identifier

        # --- Newlines (\r\n or \n or \r) -----------------------------------
        if ch in ("\r", "\n"):
            col = _col()
            if ch == "\r" and pos + 1 < length and source[pos + 1] == "\n":
                pos += 2
            else:
                pos += 1
            lineno += 1
            line_start = pos
            # Collapse: emit only if previous token isn't already NEWLINE
            if not tokens or tokens[-1].type is not TokenType.NEWLINE:
                _add(TokenType.NEWLINE, "\\n", col)
            continue

        # --- Comments: '  or Rem (at statement boundary) -------------------
        if ch == "'":
            # Skip until end of line
            while pos < length and source[pos] not in ("\r", "\n"):
                pos += 1
            continue

        # --- String literal: "..." with "" as embedded quote ---------------
        if ch == '"':
            col = _col()
            pos += 1  # skip opening quote
            buf: list[str] = []
            while pos < length:
                c = source[pos]
                if c == '"':
                    if pos + 1 < length and source[pos + 1] == '"':
                        buf.append('"')
                        pos += 2
                    else:
                        pos += 1  # closing quote
                        break
                else:
                    buf.append(c)
                    pos += 1
            else:
                raise LexerError("Unterminated string literal", lineno, col)
            _add(TokenType.STRING, "".join(buf), col)
            continue

        # --- Date literal: #...# ------------------------------------------
        if ch == "#":
            col = _col()
            pos += 1
            start = pos
            while pos < length and source[pos] != "#":
                if source[pos] in ("\r", "\n"):
                    raise LexerError("Unterminated date literal", lineno, col)
                pos += 1
            if pos >= length:
                raise LexerError("Unterminated date literal", lineno, col)
            date_val = source[start:pos]
            pos += 1  # skip closing #
            _add(TokenType.DATE_LITERAL, date_val, col)
            continue

        # --- Hex / Octal literals ------------------------------------------
        if ch == "&" and pos + 1 < length and source[pos + 1] in ("H", "h", "O", "o"):
            col = _col()
            if source[pos + 1] in ("H", "h"):
                m = _HEX_RE.match(source, pos)
            else:
                m = _OCT_RE.match(source, pos)
            if m:
                _add(TokenType.INTEGER, m.group(), col)
                pos = m.end()
                continue
            # Fall through to & operator if not a valid hex/oct literal

        # --- Numeric literals (decimal) ------------------------------------
        if ch.isdigit() or (ch == "." and pos + 1 < length and source[pos + 1].isdigit()):
            col = _col()
            m = _NUMBER_RE.match(source, pos)
            if m:
                val = m.group()
                pos = m.end()
                if "." in val or "e" in val.lower():
                    _add(TokenType.FLOAT, val, col)
                else:
                    _add(TokenType.INTEGER, val, col)
                continue

        # --- Multi-char operators (order matters) --------------------------
        two = source[pos: pos + 2]
        if two == "<>":
            _add(TokenType.NEQ, two, _col()); pos += 2; continue
        if two == "<=":
            _add(TokenType.LTE, two, _col()); pos += 2; continue
        if two == ">=":
            _add(TokenType.GTE, two, _col()); pos += 2; continue
        if two == ":=":
            _add(TokenType.COLON_EQ, two, _col()); pos += 2; continue

        # --- Single-char operators / delimiters ----------------------------
        _SINGLE: dict[str, TokenType] = {
            "(": TokenType.LPAREN,
            ")": TokenType.RPAREN,
            ",": TokenType.COMMA,
            ":": TokenType.COLON,
            ";": TokenType.SEMICOLON,
            ".": TokenType.DOT,
            "!": TokenType.BANG,
            "+": TokenType.PLUS,
            "-": TokenType.MINUS,
            "*": TokenType.STAR,
            "/": TokenType.SLASH,
            "\\": TokenType.BACKSLASH,
            "^": TokenType.CARET,
            "&": TokenType.AMPERSAND,
            "=": TokenType.EQ,
            "<": TokenType.LT,
            ">": TokenType.GT,
        }
        if ch in _SINGLE:
            _add(_SINGLE[ch], ch, _col())
            pos += 1
            continue

        # --- Identifiers / keywords ----------------------------------------
        if ch.isalpha() or ch == "_":
            col = _col()
            m = _IDENT_RE.match(source, pos)
            if m:
                word = m.group()
                pos = m.end()
                lower = word.lower()

                # Handle Rem as a comment keyword (only at statement start)
                if lower == "rem":
                    # Consume rest of line as comment
                    while pos < length and source[pos] not in ("\r", "\n"):
                        pos += 1
                    continue

                kw = _KEYWORD_MAP.get(lower)
                if kw is not None:
                    _add(kw, word, col)
                else:
                    _add(TokenType.IDENTIFIER, word, col)
                continue

        # --- Unknown character ---------------------------------------------
        raise LexerError(f"Unexpected character: {ch!r}", lineno, _col())

    # Ensure the stream ends with EOF (not a dangling NEWLINE)
    if tokens and tokens[-1].type is TokenType.NEWLINE:
        tokens.pop()
    tokens.append(Token(TokenType.EOF, "", lineno, _col()))
    return tokens
