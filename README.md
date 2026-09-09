# vba2py

**AST-based VBA to Python transpiler.** Deterministic, no AI, no cloud.

Converts Excel VBA macros to clean Python code using openpyxl for Excel operations.

## Install

```bash
pip install vba2py
```

## Usage

```bash
# Convert a VBA file to Python
vba2py macro.bas -o output.py

# Extract and convert macros straight from an Excel workbook
vba2py workbook.xlsm -o output.py

# Print the AST instead
vba2py macro.bas --ast
```

```python
from vba2py.parser import parse
from vba2py.codegen import generate

vba_code = """
Sub HelloWorld()
    Dim name As String
    name = "World"
    Debug.Print "Hello " & name
End Sub
"""
module = parse(vba_code)
python_code = generate(module)
print(python_code)
```

## Example

**VBA input:**
```vba
Function SumRange(lastRow As Long) As Double
    Dim total As Double
    Dim i As Long
    total = 0
    For i = 1 To lastRow
        If Cells(i, 1).Value <> "" Then
            total = total + Cells(i, 1).Value
        End If
    Next i
    SumRange = total
End Function
```

**Python output:**
```python
from vba_runtime import vba_len

def SumRange(lastRow) -> float:
    SumRange_result = None
    total: float = None
    i: int = None
    total = 0
    for i in range(1, lastRow + 1):
        if (ws.cell(row=i, column=1).value != ""):
            total = (total + ws.cell(row=i, column=1).value)
    SumRange_result = total
    return SumRange_result
```

## Supported VBA Constructs

| Category | VBA | Python |
|----------|-----|--------|
| **Procedures** | `Sub`, `Function` | `def` with type hints and return values |
| **Variables** | `Dim x As Integer` | `x: int = None` |
| **Constants** | `Const X = 5` | `X = 5` |
| **For loop** | `For i = 1 To 10 Step 2` | `for i in range(1, 11, 2):` |
| **For Each** | `For Each x In coll` | `for x in coll:` |
| **Do While** | `Do While x > 0 ... Loop` | `while (x > 0):` |
| **Do Until** | `Do ... Loop Until x = 0` | `while True: ... if x == 0: break` |
| **While/Wend** | `While x > 0 ... Wend` | `while (x > 0):` |
| **If/ElseIf** | `If ... ElseIf ... Else ... End If` | `if ... elif ... else:` |
| **Single-line If** | `If x > 0 Then y = 1` | `if (x > 0): y = 1` |
| **Select Case** | `Select Case x ... End Select` | `if x == ...: elif ...` |
| **Case To / Is** | `Case 1 To 10`, `Case Is > 100` | `if 1 <= x <= 10:`, `elif x > 100:` |
| **With blocks** | `With obj ... .Name = 1 ... End With` | `_with_0 = obj; _with_0.Name = 1` |
| **Set** | `Set obj = New Collection` | `obj = Collection()` |
| **Exit** | `Exit For`, `Exit Sub`, `Exit Function` | `break`, `return` |
| **On Error** | `On Error Resume Next` | `# TODO comment` |
| **GoTo** | `GoTo label` | `# TODO comment` |
| **ReDim** | `ReDim Preserve arr(20)` | `arr.extend(...)` |
| **Erase** | `Erase arr` | `arr = []` |
| **Labels** | `ErrorHandler:` | `# label: ErrorHandler` |
| **Debug.Print** | `Debug.Print "hello"` | `print("hello")` |
| **Cells** | `Cells(r, c).Value` | `ws.cell(row=r, column=c).value` |
| **Range** | `Range("A1").Value` | `ws["A1"]` |
| **String concat** | `"Hello " & name` | `str("Hello ") + str(name)` |
| **Operators** | `And`, `Or`, `Not`, `Mod`, `\`, `^` | `and`, `or`, `not`, `%`, `//`, `**` |
| **Comparisons** | `=`, `<>`, `<`, `>`, `<=`, `>=` | `==`, `!=`, `<`, `>`, `<=`, `>=` |

## Runtime Library (vba_runtime)

The transpiler generates code that imports from `vba_runtime`, a library that emulates VBA-specific behavior:

| Category | Functions |
|----------|-----------|
| **Strings** | `vba_left`, `vba_right`, `vba_mid`, `vba_instr`, `vba_replace`, `vba_len`, `vba_trim`, `vba_ucase`, `vba_lcase`, `vba_split`, `vba_join`, `vba_strreverse`, `vba_asc`, `vba_chr`, `vba_space` |
| **Conversion** | `vba_cstr`, `vba_cint`, `vba_clng`, `vba_cdbl`, `vba_cbool`, `vba_val`, `vba_int_func`, `vba_fix` |
| **Math** | `vba_abs`, `vba_sgn`, `vba_sqr`, `vba_round`, `vba_rnd` |
| **Type checking** | `vba_isnumeric`, `vba_isempty`, `vba_isnull`, `vba_isdate` |
| **Special values** | `EMPTY` (falsy, equals 0 and ""), `NULL` (propagates through arithmetic) |
| **Collections** | `VBACollection` (1-based, string-keyed), `VBAArray` (ReDim Preserve) |
| **Errors** | `VBAError` (Err object), `OnErrorResumeNext` (context manager) |
| **Excel** | `VBAWorkbook`, `VBAWorksheet`, `VBARange` (openpyxl wrappers) |

## Not Supported (v1)

These VBA features cannot be automatically transpiled:

- COM Automation (`CreateObject`, late binding)
- UserForms and ActiveX controls
- Windows API declarations (`Declare Function ... Lib`)
- `Application.*` methods (screen updating, calculation modes)
- `SendKeys`
- Events and callbacks

## Architecture

```
VBA Source
    |
  Lexer (57 keywords, case-insensitive)
    |
  Recursive Descent Parser
    |
  Typed AST (30+ node types)
    |
  Code Generator -> Python Source + vba_runtime imports
```

No external parser dependencies. No ANTLR, no Java. Hand-written recursive descent parser with full operator precedence.

## Tests

```bash
pip install -e ".[dev]"
pytest tests/ -v
```

154 tests covering lexer, parser, codegen, runtime library, and full-pipeline integration.

## Why Not AI?

LLMs produce less than 40% correct code on real benchmarks. A deterministic parser gives you the same output every time, and you can verify correctness with tests. This tool handles the mechanical 60-70% of VBA migration so you can focus on the parts that need domain knowledge.

## License

MIT
