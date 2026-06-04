# vba2py

**AST-based VBA to Python transpiler.** Deterministic, no AI hallucinations.

Converts Excel VBA macros to clean Python code using openpyxl/pandas.

> **Status:** Early development. Research phase complete, implementation starting.

## Why

- 345M+ Microsoft 365 users, millions of legacy VBA macros
- No open-source AST-based VBA-to-Python converter exists
- AI approaches produce <40% correct code on complex tasks
- VBScript deprecated (removal post-2027), Office 2016/2019 EOL
- Enterprise pays $5K-250K+ for manual migration

## Approach

```
.xlsm → oletools/olevba → ANTLR4 Parser → VBA AST → Analyzer → Python IR → ast.unparse() + black → Python + vba_runtime
```

- **Parser:** ANTLR4 with battle-tested VBA grammar
- **IR:** Python `ast` module nodes for guaranteed valid output
- **Runtime:** `vba_runtime` package for VBA built-in functions and Excel object model wrappers

## Scope (v1)

**Supported:**
- Loops, conditionals, Select Case
- Functions/Subs → `def`
- Range/Cell read/write via openpyxl
- VBA built-in string functions (Left, Right, Mid, InStr)
- Error handling (On Error → try/except)
- Variable declarations, type hints

**Not supported (v1):**
- COM Automation (CreateObject)
- UserForms
- Events / callbacks
- Windows API declarations

## License

MIT
