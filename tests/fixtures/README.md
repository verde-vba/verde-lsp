# VBA Language Server Test Fixtures

A comprehensive corpus of VBA source code for testing the verde-lsp parser, lexer, symbol resolution, and completion features.

## Directory Structure

| Directory | Files | Description |
|-----------|-------|-------------|
| `basic/` | 10 | Fundamental constructs: Sub, Function, Property, Dim, Const, Type, Enum, access modifiers, scope |
| `control_flow/` | 8 | Control flow: If/ElseIf/Else, For/Next, For Each, Do Loop, While/Wend, Select Case, With, GoTo/GoSub |
| `oop/` | 11 | OOP patterns: Class modules, Implements, Events, RaiseEvent, WithEvents, constructors, IEnumVARIANT |
| `error_handling/` | 3 | Error handling: On Error GoTo, On Error Resume Next, Resume, Erl |
| `advanced/` | 11 | Advanced features: ReDim Preserve, Declare (API calls), ParamArray, Optional, ByVal/ByRef, DefType, file I/O, Like operator, collections |
| `real_world/` | 5 | Complex real-world modules: financial calculations, Excel utilities, clipboard API, keyboard shortcuts |
| `edge_cases/` | 11 | Parser edge cases: line continuations, multi-statement lines, Attribute lines, conditional compilation, form modules (.frm), literals, type hints, nested structures, empty procedures |
| `excel_objects/` | 6 | Excel Object Model usage: Worksheet events, ThisWorkbook, Range operations, Workbook management, Charts, PivotTables |

## File Extensions

- `.bas` -- Standard modules
- `.cls` -- Class modules
- `.frm` -- Form modules (UserForm with designer section)

## Coverage Summary

### Language Constructs Covered

- **Declarations**: Sub, Function, Property Get/Let/Set, Dim, ReDim, Const, Type, Enum, Declare
- **Access modifiers**: Public, Private, Friend, Static
- **Control flow**: If/ElseIf/Else, For/Next, For Each/Next, Do While/Until/Loop, While/Wend, Select Case, With, GoTo, GoSub/Return, Exit, On...GoTo, On...GoSub
- **Error handling**: On Error GoTo, On Error Resume Next, On Error GoTo 0, Resume, Resume Next, Err object
- **Operators**: Like, AddressOf, Is, TypeOf...Is
- **Statements**: Let, Set, Call, Open/Close, Print, Input, Line Input, Get, Put, Seek, Lock/Unlock, Width, Write, Mid (statement), LSet, RSet, Name, Kill, MkDir, RmDir, ChDir, ChDrive, FileCopy, Erase, Randomize, Stop, Load/Unload, AppActivate, SendKeys, SetAttr, SaveSetting/DeleteSetting, SavePicture, Beep
- **Parameters**: ByVal, ByRef, Optional, ParamArray
- **Options**: Option Explicit, Option Compare Binary/Text, Option Base, Option Private Module
- **Compiler directives**: #If/#ElseIf/#Else/#End If, #Const
- **Attributes**: VB_Name, VB_GlobalNameSpace, VB_Creatable, VB_PredeclaredId, VB_Exposed, VB_Description, VB_UserMemId, VB_MemberFlags
- **OOP**: Implements, Event, RaiseEvent, WithEvents, Class_Initialize, Class_Terminate, IEnumVARIANT (For Each support), default members
- **Type hints**: %, &, !, #, @, $ suffix characters
- **Literals**: Integer, Long, Hex (&H), Octal (&O), Float, String (with escaping), Boolean, Date (#...#), Nothing, Null, Empty
- **Edge cases**: Line continuations (_), multi-statement lines (:), Rem comments, line numbers, nested structures, empty procedures

### Excel Object Model Coverage

- Workbook, Worksheet, Range, Cells, Selection, Application
- Chart, ChartObject, PivotTable, PivotCache, PivotField
- WorkbookConnection, ListObject, Names
- Events: Worksheet_Change, Worksheet_SelectionChange, Workbook_Open, Workbook_BeforeClose, Workbook_BeforeSave, etc.
- Methods: Find, Sort, AutoFilter, SpecialCells, Intersect, Union, Copy, PasteSpecial

## Sources and Attribution

### Primary Sources (test suites from VBA parser/LSP projects)

| Repository | License | What was used |
|-----------|---------|---------------|
| [tmepple/tree-sitter-vba](https://github.com/tmepple/tree-sitter-vba) | -- | Example `.bas`/`.cls` files and test corpus patterns (control flow, declarations, error handling, literals, procedures) |
| [uwol/proleap-vb6-parser](https://github.com/uwol/proleap-vb6-parser) | MIT | MSDN statement test fixtures (80+ files covering every VB6/VBA statement), ASG integration tests (class modules, properties, calls, enums) |
| [Beakerboy/VBA-Linter](https://github.com/Beakerboy/VBA-Linter) | MIT | Test VBA files for linter rules |
| [SSlinky/VBA-LanguageServer](https://github.com/SSlinky/VBA-LanguageServer) | -- | Reference for VBA LSP feature coverage |

### VBA Code Libraries (real-world modules)

| Repository | License | What was used |
|-----------|---------|---------------|
| [nylen/vba-common-library](https://github.com/nylen/vba-common-library) | MIT | ExcelUtils, ArrayUtils, StringUtils, FileUtils patterns |
| [austinleedavis/VBA-utilities](https://github.com/austinleedavis/VBA-utilities) | MIT | Collections, Arrays, interfaces (ICollection, IList, IVariantComparator) |
| [codetuner/VBA-Library](https://github.com/codetuner/VBA-Library) | -- | Clipboard API, Dictionary, Stack, JSON converter patterns |
| [ReneNyffenegger/about-VBA](https://github.com/ReneNyffenegger/about-VBA) | -- | Interface examples, class patterns, error handling, conditional compilation, arrays, operators, data types, IEnumVARIANT |

### Other References

| Repository | Notes |
|-----------|-------|
| [rubberduck-vba/Rubberduck](https://github.com/rubberduck-vba/Rubberduck) | GPL-3.0. VBA parser grammar (ANTLR) and test patterns referenced for coverage completeness |
| [Beakerboy/antlr4-vba](https://github.com/Beakerboy/antlr4-vba) | ANTLR4 VBA grammar reference |
| [rossknudsen/Vba.Language](https://github.com/rossknudsen/Vba.Language) | ANTLR VBA grammar project reference |
| [serkonda7/vscode-vba](https://github.com/serkonda7/vscode-vba) | TextMate grammar for VBA syntax reference |

## Notes

- All fixtures are synthetic, composed from MSDN documentation examples, open-source VBA code, and purpose-built test cases.
- Code from external sources has been adapted, simplified, or restructured to serve as parser test cases.
- The proleap-vb6-parser MSDN examples are derived from the official Microsoft VB6/VBA language reference.
- Some files contain intentionally unusual but valid VBA constructs to stress-test parser edge cases.
