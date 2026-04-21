# Changelog

All notable changes to **verde-lsp** (VBA Language Server) are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-04-21

First tagged release. Covers the full MVP scope plus Phase 2/3 LSP features.

### Added

#### Parser / analysis
- `logos`-based VBA lexer (case-insensitive keywords).
- Recursive-descent parser producing an arena-allocated AST (`la-arena`).
- Symbol table construction with module / procedure / variable / constant / UDT member scopes.
- Name resolution and diagnostics, including `Option Explicit` undeclared-variable detection.
- UDT (`Type` block) parsing and member symbol registration.

#### LSP features
- `textDocument/completion` — keywords, built-in functions, local / module symbols,
  Excel Object Model members (Range, Worksheet, Workbook, Application, PivotTable, Chart, Shape),
  and sheet / table / named-range identifiers sourced from `workbook-context.json`.
- `textDocument/hover` — variable types, procedure signatures, built-in descriptions.
- `textDocument/signatureHelp` — parameter hints for procedures and built-ins.
- `textDocument/definition` — goto-definition for procedures, variables, constants, types.
- `textDocument/references` — find all references within a module / workspace.
- `textDocument/rename` — symbol rename across a module.
- `textDocument/documentHighlight` — highlight occurrences of the symbol under the cursor.
- `textDocument/documentSymbol` — outline view of modules, procedures, and fields.
- `textDocument/foldingRange` — fold `Sub` / `Function` / `If` / `For` / `With` / `Type` blocks.
- `textDocument/codeAction` — quick fix to insert a `Dim` statement for undeclared variables
  flagged under `Option Explicit`.
- `textDocument/formatting` — indent normalization (depth tracking for
  `Sub` / `If` / `For` / `With` / `Select` / `Do` / `While` / `Type`,
  with `ElseIf` / `Else` / `Case` aligned at `depth - 1`),
  keyword case normalization, and trailing-whitespace trimming.
- `textDocument/inlayHint` — inline type labels for `Dim` variables and constants
  (reuses `Symbol.type_name`, falls back to `Variant`).
- `textDocument/prepareCallHierarchy` + `callHierarchy/incomingCalls` + `callHierarchy/outgoingCalls` —
  text-based call hierarchy using procedure body ranges.
- `workspace/symbol` — workspace-wide symbol search (public procedures, module-level names).
- `Me` keyword completion / hover / goto-def for the current class module
  (`.cls` header validation).

#### Class modules
- `.cls` file support, including the `VERSION 1.0 CLASS` header block and
  `Attribute VB_Name` / `VB_Exposed` parsing.
- Instance variable / method resolution via the `Me` keyword.

#### Tooling / packaging
- stdio-based LSP transport (`tower-lsp`).
- UTF-16 position encoding advertised in `ServerCapabilities`.
- Release artifacts for Windows, Linux, and macOS (Release workflow on `v*` tag push).
- CI on Ubuntu, Windows, and macOS (`cargo test --all`,
  `cargo clippy -- -D warnings`, `cargo fmt --check`).

### Known limitations

- Excel Object Model coverage is an MVP subset (the types listed above); less common
  members fall through to generic `Variant`.
- Call hierarchy is text-based, so a type reference like `Dim x As Foo` can be reported
  as a call to `Foo`. An AST-based call-site detector is tracked as a future improvement.

[0.1.0]: https://github.com/verde-vba/verde-lsp/releases/tag/v0.1.0
