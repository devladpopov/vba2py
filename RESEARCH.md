# VBA to Python Transpiler: Market & Technical Research

Date: 2026-06-03

---

## 1. Existing Converters & Competitors

### Open-Source (GitHub)

```
Tool                Stars  Approach        Last Update  Status
vb2py (reingart)    39     AST-based       2014         Dead (Jul 2024)
xl_vb2py (ceprio)   4      AST + xlwings   2021         Dead
py2vba (brotchie)   36     Reverse (Py->VBA) 2021       Dead
xlsliberator        6      LLM + regex     Active       Claude API hybrid
MCP-Server-VBA      1      FastMCP/AI      Dec 2025     Minimal
```

Вывод: ни одного живого AST-based VBA->Python конвертера в open-source. Поле пустое.

### Коммерческие / SaaS

| Продукт | Подход | Модель | Ограничения |
|---------|--------|--------|-------------|
| **VBAtoPython.com** | Детерминированный rule engine | $5 файл, $19 day pass, $49 week, $149/мес | 5000 строк макс, no company info |
| **CodeConvert AI** | LLM-based | Freemium (2 конверсии/день) | 50K char лимит, hallucinations |
| **CodePorting.ai** | AI + YAML config | SaaS | Snippet-level |
| **Taskade** | AI-based | SaaS | Generic, не специализирован |
| **AWS Transform Custom** | Enterprise AI + auto-testing | Enterprise (AWS pricing) | Cloud-locked, opaque, дорого |
| **CodingFleet** | AI-based | SaaS | Специфично для Access VBA |

Главный прямой конкурент: **VBAtoPython.com**. Deterministic rule engine + optional AI refinement. Closed-source. Free tier до 100 строк, Pro: $5/файл, $19/day pass, $49/week, $149/мес. Лимит 5000 строк / 250K символов на модуль. Покрывает loops, conditionals, cell ops, string functions. Слабо на COM, UserForms, ActiveX.

**AWS Transform Custom** (май 2026): самый серьёзный enterprise-конкурент. AI с intelligent chunking, автогенерация тестов, regression testing, рефакторинг procedural->OOP. Доступен в Kiro и VS Code. Но AWS lock-in и enterprise pricing.

### AI/LLM подходы

- ChatGPT/Claude: ~85-95% functional accuracy на coding tasks, но нет бенчмарков именно для VBA->Python
- Лучшие LLM реализуют <40% кода корректно (ResearchCodeBench, Gemini-2.5-Pro: 37.3%)
- xlsliberator: 4-фазный пайплайн (Claude API + AST validation), claims 90%+ success
- VBAssistant AI: Microsoft Store, AI co-pilot для VBA Editor (не конвертер)
- Проблема AI: hallucinations, непоследовательность, нет гарантий корректности

### Вывод по конкурентам

Ниша AST-based open-source конвертера **полностью свободна**. VBAtoPython.com закрыт и дорогой. AI-подходы ненадёжны. Это реальная возможность для 3-10K stars проекта.

---

## 2. Смежные инструменты

### VBA Parsers

```
Tool                  Stars  Tech         Use Case
ViperMonkey           1122   pyparsing    Malware analysis, VBA emulation
oletools              3347   Python       OLE extraction (olevba)
pcodedmp              488    Python       P-code disassembly
pcode2code            115    Python       P-code decompilation
vbSparkle             --     ANTLR/C#     VBA deobfuscation (Airbus CERT)
Rubberduck VBA        2100   ANTLR4/C#    VBA IDE, static analysis (archived Mar 2026)
```

### ANTLR Grammars для VBA

```
Source                    Quality     Notes
grammars-v4/vba6          Basic       70+ keywords, module structure
Rubberduck VBALexer/Parser Complete    Split lexer/parser, case-insensitive, best quality
Vba.Language (rossknudsen) Outdated    C#, archived 2015
antlr4-vba-parser          PoC        Python, 3 stars, incomplete
```

**Лучший вариант для парсера: грамматика Rubberduck** (наиболее полная, покрывает implicit/explicit calls, events, annotations, file I/O). Но написана под C#, нужна адаптация для Python target.

### Python-Excel bridges

```
Tool       Stars  License    Role
xlwings    3400   BSD        Bidirectional Python-Excel (requires Excel)
openpyxl   --     MIT        Read/write .xlsx (no formula eval)
PyXLL      --     Commercial Python as Excel add-in
xlSlim     --     Commercial Python bundled in Excel
pycel      604    --         Compiles Excel spreadsheets to Python
formulas   416    --         Excel formula interpreter
xlcalculator --   --         Formula network evaluation
```

---

## 3. Рынок

### Размер

- 345M подписчиков Microsoft 365
- 100M+ пользователей указывают Excel как навык в LinkedIn
- VBA используется в: финансах, бухгалтерии, страховании, производстве, госструктурах
- 4M нехватка разработчиков к 2025, VBA-специалистов становится всё меньше

### Тренд миграции

- Microsoft блокирует интернет-макросы по умолчанию (с 2022-2023)
- VBScript deprecated: Phase 1 (2024-25, Feature on Demand), Phase 2 (2026-27, disabled by default), Phase 3 (post-2027, полное удаление)
- Office 2016/2019 end-of-life: 14 Oct 2025, нет security updates для legacy VBA
- Python in Excel (GA 2024): =PY() в ячейке, pandas/matplotlib/scikit-learn. Работает в Microsoft Cloud (не локально). Дополняет VBA, не заменяет: нельзя делать automation, file exports, макросы
- Microsoft инвестирует в Office Scripts (TypeScript, cloud-native), Power Automate
- VBA "still supported" но экосистемное давление очевидно
- Реальный бенчмарк: Python (pandas) 20x быстрее VBA на 180MB финансовом датасете
- Прогноз: 10-15 лет переходный период с длинным хвостом. Новые проекты на Python, legacy VBA остаётся

### Боли пользователей

1. **Безопасность**: VBA макросы как главный вектор малвари, аудит-проблемы
2. **Талант**: VBA-разработчики уходят, новые не приходят
3. **Legacy code**: тысячи макросов, никто не знает что они делают
4. **Отсутствие инструментов**: нет автоматической миграции, всё руками
5. **1-based vs 0-based**: массивы, индексы, Off-by-one ошибки
6. **GoTo/On Error Resume Next**: нет прямого аналога в Python
7. **COM/ActiveX зависимости**: не портируются
8. **26% аналитиков не рекомендуют Excel** (опрос), только 35% времени FP&A на высокоценные задачи

### Стоимость миграции

- Enterprise: $5K-$250K+ за проект
- Консультанты: $67-200/час, проекты длятся недели/месяцы
- Инструмент, автоматизирующий даже 40% конверсии vs 95%, даёт 3-4x разницу в стоимости проекта
- Главные блокеры: объём legacy кода, password-protected VBA projects ("останавливает 90% конверсионных проектов"), бизнес-риск (финансовые модели mission-critical)

---

## 4. Технические вызовы

### VBA Grammar

- Официальная спецификация: MS-VBAL rev 2.4 (May 2025), ~400 стр., ABNF-формат, не machine-parseable
- Case-insensitive (lexer обрабатывает через fragment rules)
- Optional parentheses в вызовах: `MsgBox "hello"` vs `MsgBox("hello")` имеют разную семантику (ByRef vs ByVal)
- Line continuations (`_` с пробелом перед ним)
- Implicit typing (Variant по умолчанию, со специфичной coercion: `Null + 1 = Null`, `"5" + 3 = 8`)
- With blocks (контекстный `.Property` доступ, нужен scope-stack)
- Colon как statement separator: `x = 1 : y = 2`
- Default properties: `Range("A1") = 5` это `Range("A1").Value = 5` (implicit COM concept)
- Bang operator: `rs!FieldName` = `rs.Fields("FieldName").Value`
- Single-line If vs multi-line If (два разных parse rules)
- VBA7: `PtrSafe`, `LongPtr` для 64-bit. Структурно ~идентичен VBA6

### Parsing quirks (критичные для реализации)

1. **Optional parens**: без скобок аргументы ByRef, со скобками вокруг единственного аргумента ByVal
2. **Default properties**: COM concept без аналога в Python
3. **Variant coercion**: `Empty + 1 = 1`, `Null + anything = Null`, `"5" + 3 = 8` vs `"5" & 3 = "53"`
4. **ByRef primitives**: Python не поддерживает pass-by-reference для примитивов, нужны wrapper objects или return tuples

### Маппинг Excel Object Model -> openpyxl

```
VBA                        openpyxl                         Сложность
Workbooks.Open(path)       openpyxl.load_workbook(path)     Easy
Workbook.ActiveSheet       wb.active                        Easy
Sheets("name")             wb["name"]                       Easy
Range("A1").Value          ws["A1"].value                   Easy
Cells(r, c)                ws.cell(row=r, column=c)         Easy (both 1-based)
Range("A1:B5")             ws["A1:B5"]                      Easy (tuple of tuples)
UsedRange                  ws.dimensions                    Easy (returns string)
Worksheet.Name             ws.title                         Easy
Merge                      ws.merge_cells()                 Easy
Save                       wb.save(filename)                Easy (must specify name)
Range.Font.Bold            cell.font = Font(bold=True)      Medium (style objects)
Range.NumberFormat          cell.number_format               Easy
Range.Offset(1,2)          coordinate math                  Medium
Range.Find/Replace         нет в openpyxl, итерация         Medium
Range.Sort/AutoFilter      ws.auto_filter (partial)         Medium
Range.Copy/Paste           decompose to read+write          Medium
ActiveCell/Selection        --                              Hard (stateful, runtime)
Application.*              --                              N/A (no-op or logging)
Events                     --                              Impossible
UserForms                  --                              Impossible
COM/API Declare            --                              Impossible
```

### pandas mapping (для data-processing макросов)

```
VBA Pattern               pandas Equivalent
Read range into array     pd.read_excel(path, sheet_name=...)
Loop through rows         df.iterrows() or vectorized ops
Filter rows               df[df["col"] > value]
Sort                      df.sort_values(by=...)
VLOOKUP                   df.merge() or df.set_index().loc[]
Aggregate/sum             df.groupby().sum()
Write back                df.to_excel(path)
```

Выбор openpyxl vs pandas зависит от задачи: openpyxl для cell-level манипуляции и форматирования, pandas для data transformation.

### Transpilation difficulty matrix

**Tier 1: Direct 1:1 (~60-70% типичных макросов)**

```
VBA                            Python                         Complexity
If/Then/Else/ElseIf            if/elif/else                   Trivial
For i = 1 To 10 Step 2         for i in range(1, 11, 2)       Trivial (adjust bounds)
For Each x In collection       for x in collection            Trivial
Do While/Until...Loop          while condition:               Trivial
Select Case x                  match x: (3.10+) or if/elif   Low
Dim x As Integer               x: int (type hint)             Low
Sub/Function                   def                            Low
Exit Sub/Function/For/Do       return / break                 Low
x & y (string concat)          str(x) + str(y)                Low
Not/And/Or                     not/and/or                     Low
Nothing                        None                           Trivial
```

**Tier 2: Requires runtime library**

```
VBA                            Strategy                       Complexity
Left/Right/Mid                 s[:n] / s[-n:] / s[start-1:]   Medium (1-based)
InStr(start, s, find)          s.find(find, start-1) + 1      Medium
CStr/CInt/CLng/CDbl            str()/int()/float()            Low
IsNumeric/IsDate/IsEmpty       Runtime helpers                Medium
DateSerial/Now/Date            datetime module                Medium
Format(x, "0.00")             f-string or runtime mapping    Medium
Array(1,2,3)                   VBAArray([1,2,3])              Low
Err.Number/Description         Custom VBAError class          Medium
MsgBox/InputBox                print/input or GUI             Low
```

**Tier 3: Significant emulation**

```
VBA                            Challenge                      Complexity
On Error GoTo label            Restructure to try/except      High
On Error Resume Next           Wrap every statement in try    High
GoTo label                     No goto in Python, state machine High
GoSub/Return                   Extract into separate function High
With obj / .Prop               Temp variable or context mgr   Medium
Variant                        VBAVariant coercion rules      High
ByRef parameters               Wrapper objects / return tuples High
1-based arrays (Option Base)   VBAArray wrapper class         Medium
ReDim Preserve                 list.extend()                  Medium
Class modules                  Python classes, @property      Medium
Collections (1-based, keyed)   Custom VBACollection           Medium
```

**Tier 4: Practically impossible**

```
VBA                            Why
Declare Function ... Lib       Win API / DLL calls, platform-specific
UserForms                      GUI framework, need tkinter/PyQt per control
CreateObject("...")            COM automation, need win32com
Application.Run/OnTime         Excel scheduler
SendKeys                       UI automation
ActiveX controls               Deeply Excel GUI-integrated
```

### Рекомендуемая архитектура

```
.xlsm/.xlsb file
    |
[oletools/olevba] -- Extract VBA source per module
    |
[ANTLR4 Lexer] (VBALexer.g4 -> Python target)
    |
[ANTLR4 Parser] (VBAParser.g4 -> Python target)
    |
[Concrete Parse Tree] (ANTLR4 ParseTree)
    |
[AST Builder] (Visitor pattern -> typed AST nodes)
    |
[VBA AST] (clean, typed node hierarchy)
    |
[Analyzer / Transformer]
  - Type inference pass
  - With-block resolution
  - Default property resolution
  - ByRef parameter detection
  - Error handling restructuring (GoTo -> try/except)
  - Array bounds analysis
    |
[Python IR] (Python ast.Module nodes)
    |
[ast.unparse()] + black formatter
    |
[Python source] + [vba_runtime imports]
```

### Architectural Decision Records

**ADR-001: ANTLR4 с грамматикой Rubberduck (GPL v3)**
Самая полная и battle-tested VBA грамматика. 14,898 commits, production IDE. ANTLR4 генерирует Python parser из .g4 файлов.
Fallback: grammars-v4/vba6 (MIT) если GPL проблема.
Downside: ANTLR4 runtime dependency, Java для компиляции грамматики.

**ADR-002: Two-Stage AST (ParseTree -> Typed AST)**
ANTLR4 parse tree generic (string tokens). Visitor конвертирует в typed AST. Чистое разделение parsing/semantics. Grammar changes affect only visitor.

**ADR-003: Python ast module как IR**
Генерируем `ast.Module` nodes, используем `ast.unparse()` (Python 3.9+). Гарантированно валидный Python. Black как post-formatter.

**ADR-004: Runtime library (vba_runtime)**
Generated code делает `from vba_runtime import Mid, VBACollection`. Тестируется независимо.

**ADR-005: Scope v1: Excel Data Macros**
Target: 80% use case (loops, conditionals, string manipulation, Range/Cell read/write). Exclude: UserForms, COM, API Declare, Events. Чёткие error messages для unsupported constructs.

### Runtime library structure

```
vba_runtime/
  __init__.py
  types.py          -- VBAVariant, VBACollection, VBAArray, VBADate, VBACurrency
  errors.py         -- VBAError, OnErrorResumeNextContext
  strings.py        -- Mid, InStr, Replace, Left, Right, Format (1-based)
  math_funcs.py     -- Int, Fix, Rnd, Sgn, Abs (VBA-compatible)
  conversion.py     -- CStr, CInt, CLng, CDbl, CBool, CDate, Val
  datetime_funcs.py -- DateSerial, DateAdd, DateDiff, Now, Timer
  file_io.py        -- Open/Close/Print/Input# (sequential file I/O)
  interaction.py    -- MsgBox, InputBox (console or GUI mode)
  excel/
    workbook.py     -- VBAWorkbook wrapping openpyxl.Workbook
    worksheet.py    -- VBAWorksheet wrapping openpyxl.Worksheet
    range_.py       -- VBARange wrapping cell/range operations
    application.py  -- VBAApplication (global state, no-ops)
```

### Стратегия тестирования

1. **Unit tests per AST node**: parse VBA snippet, verify AST structure
2. **Golden file tests**: VBA input -> expected Python output, `assert transpile(vba) == expected`
3. **Behavioral equivalence**: run VBA in Excel (recorded results), run transpiled Python, compare outputs
4. **Incremental corpus**: single statements -> real-world macros
5. **Regression suite**: публичные VBA макросы, transpile, verify valid Python

### Выбор парсера: сравнение

```
Parser                         Completeness  License  Reusable?
Rubberduck VBAParser+Lexer.g4  Best          GPL v3   Best candidate
grammars-v4 vba6/vba.g4        Good (~90%)   MIT      Good fallback
ViperMonkey (pyparsing)        Partial       BSD      No (too coupled)
antlr4-vba-parser (PyPI)       = grammars-v4 BSD      PoC only (4 commits)
oletools/olevba                Extraction    BSD      Input stage only
```

---

## 5. Целевая аудитория

```
Сегмент                     Размер     Боль                          Бюджет          Срочность    Канал
Финансовые аналитики        Большой    Legacy модели 5-50K строк     $500-5K/проект  Средняя      LinkedIn, CFA/CFI
Enterprise IT               Средний    Security, compliance, аудит   $10K-250K       Высокая      Enterprise sales, Ignite
Data engineers              Средний    ETL модернизация              Средний         Средняя      r/dataengineering, dbt
Бухгалтеры                  Огромный   Макросы 10+ лет, автор ушёл   $5-50           Низкая       r/accounting
Индивидуальные разработчики Огромный   Хочу Python, знаю VBA         $0-20           Низкая       YouTube, SO, r/learnpython
```

---

## 6. Бизнес-модель

### Рекомендация: Open-Source Core + SaaS

**Open-source (MIT/Apache-2.0)**:
- CLI конвертер (pip install vba2python)
- VS Code extension
- GitHub Action для CI
- Привлечение community, contributions, stars

**SaaS (vba2python.dev или подобное)**:
- Web-интерфейс: вставь VBA, получи Python
- Batch conversion (директории/проекты)
- Визуализация AST
- Enterprise API

**Pricing (ориентир на VBAtoPython.com, но дешевле)**:
- Free: до 200 строк, одиночные файлы (SEO + user acquisition)
- Pro $20-50/мес: unlimited files, batch, API access
- Enterprise $500-2000/мес: on-prem, audit trails, team features, priority support

### VS Code Extension

- Marketplace: бесплатная base-версия
- Highlight VBA, кнопка "Convert to Python"
- Preview diff
- Freemium upsell к SaaS

### Revenue path

Free open-source CLI + web converter (acquisition) -> Pro SaaS для power users -> Enterprise licensing для IT teams. Enterprise сегмент ($10K-250K/проект) это основной revenue, но bottom-up adoption от разработчиков и аналитиков это воронка.

### Timing

Окно открывается сейчас (2026-2028): VBScript Phase 2, Office EOL, macro blocking сходятся. Early mover advantage значителен: enterprise migration projects длятся месяцы и организации зафиксируют tooling рано.

---

## 7. SEO

### Ключевые запросы

```
Keyword                           Est. Monthly (Global)  Competition  Intent
vba to python                     2000-5000              Medium       Info/Transactional
convert vba to python             500-1500               Low-Medium   Transactional
python vs vba                     1000-3000              Medium       Informational
excel macro to python             500-1000               Low-Medium   Info/Transactional
excel vba deprecated              500-1500               Low          Info (fear-driven)
vba to python converter           200-500                Low          High-intent
python in excel vs vba            300-800                Low-Medium   Informational
vba to python cheat sheet         200-500                Low          Informational
vba migration tool                50-200                 Very Low     High-intent
vba parser                        100-300                Very Low     Developer
vba to python transpiler          <50                    Very Low     Developer
```

### Long-tail (низкая конкуренция, высокий intent)

- "convert vba macro to python script"
- "vba to python migration guide finance"
- "replace vba with python openpyxl"
- "vba on error goto python equivalent"
- "excel vba to pandas dataframe"
- "vba array to python list conversion"
- "how to convert vba class module to python"

### Content gaps (чего нет в интернете)

1. Comprehensive VBA-to-Python migration playbook для enterprise
2. Industry-specific guides: "VBA to Python for Financial Modeling"
3. Before/after performance benchmarks с реальными данными
4. Видео live-конверсии сложных VBA проектов
5. "VBA audit" tool: сканирует .xlsm и оценивает сложность миграции
6. Comparison content: "vbatopython.com vs ChatGPT for VBA conversion"
7. Regulatory angle: VBA risks + SOX, GDPR compliance
8. Case studies: "How [Company] migrated 50K lines of VBA to Python"

### Content-стратегия

1. Landing page: "The Open-Source VBA to Python Converter"
2. Blog: "Why migrate from VBA to Python in 2026"
3. Docs: mapping tables (VBA function -> Python equivalent)
4. Tutorials: "Convert your first Excel macro to Python"
5. Comparison: "VBA vs Python for Excel automation"
6. SEO pages: по каждому long-tail keyword

---

## 8. Ключевые выводы

1. **Ниша свободна**: нет ни одного живого open-source AST-based конвертера
2. **Рынок огромен**: 345M пользователей Excel, миллионы VBA макросов
3. **Боль реальна**: security, talent shortage, legacy code
4. **Технически реализуемо**: грамматики существуют (Rubberduck), openpyxl mapping покрывает 70%+ use cases
5. **Парсер**: ANTLR4 с грамматикой Rubberduck (адаптация на Python target)
6. **Scope v1**: базовые макросы (циклы, условия, Range access, built-in функции), без COM/UserForms/Events
7. **Прямой конкурент**: VBAtoPython.com (closed-source, дорогой)
8. **AI не конкурент**: ненадёжен для production migration, наш детерминированный подход дополняет AI
9. **Потенциал**: 3-10K stars реалистичен при хорошем README, docs, и реальной полезности
10. **Timing**: окно 2026-2028, VBScript Phase 2 + Office EOL + macro blocking сходятся
11. **AWS Transform Custom**: единственный серьёзный enterprise конкурент (май 2026), но AWS lock-in
12. **Defensibility**: детерминированный AST-parser с глубоким Excel Object Model knowledge сложнее реплицировать чем GPT-wrapper
13. **Building blocks**: ANTLR grammars-v4/vba (канонический), oletools (extraction), xlwings (target runtime), formulas/pycel (formula eval)
