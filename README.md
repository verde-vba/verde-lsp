# verde-lsp

VBA Language Server for [Verde](https://github.com/verde-vba/verde).

Rust-based LSP implementation that parses `.bas` / `.cls` / `.frm` and speaks the
Language Server Protocol over stdio. Designed for the Verde desktop app and usable
from any LSP-aware editor.

## Features

### Navigation & editing
- **Go to Definition** — procedures, variables, constants, user-defined types
- **Find References** — module- / workspace-wide occurrences
- **Rename** — safe symbol rename across a module
- **Document Highlight** — highlight the symbol under the cursor
- **Document Symbol** — outline view of modules, procedures, fields
- **Workspace Symbol** — workspace-wide symbol search
- **Folding Range** — `Sub` / `Function` / `If` / `For` / `With` / `Type` blocks
- **Call Hierarchy** — `prepareCallHierarchy` + incoming / outgoing calls

### Assistance while typing
- **Completion** — keywords, built-in functions, local / module symbols, and
  Excel Object Model members (Range, Worksheet, Workbook, Application,
  PivotTable, Chart, Shape). Sheet / table / named-range names are sourced from
  `workbook-context.json`.
- **Hover** — variable types, procedure signatures, built-in descriptions
- **Signature Help** — parameter hints for procedures and built-ins
- **Inlay Hint** — inline type labels for `Dim` variables and constants

### Correctness
- **Diagnostics** — `Option Explicit` undeclared-variable detection, parse errors
- **Code Action** — quick fix to insert a `Dim` statement for undeclared variables
- **Formatting** — indent normalization, keyword case normalization,
  trailing-whitespace trimming

### Class modules
- `.cls` support, including the `VERSION 1.0 CLASS` header and
  `Attribute VB_Name` / `VB_Exposed`
- `Me` keyword — completion / hover / goto-def on the current class

## Usage

`verde-lsp` communicates over stdio following the LSP spec.

```bash
mise install          # installs Rust + just
just build            # cargo build --release
just run              # start the server on stdio
just test             # cargo test
just clippy           # cargo clippy -- -D warnings
just fmt              # cargo fmt
just check            # cargo check
```

The release binary is at `target/release/verde-lsp` (or `verde-lsp.exe` on Windows).
Prebuilt binaries for Windows and Linux are published on each `v*` tag — see
[Releases](https://github.com/verde-vba/verde-lsp/releases).

## Architecture

```
src/
├── main.rs              # Entry point (stdio LSP server)
├── server.rs            # LSP protocol handler
├── lib.rs               # Crate root (module graph)
├── parser/
│   ├── lexer.rs         # logos-based VBA tokenizer
│   ├── ast.rs           # AST node definitions (la-arena)
│   └── parse.rs         # Recursive descent parser
├── analysis/
│   ├── symbols.rs       # Symbol table construction
│   ├── resolve.rs       # Name resolution
│   └── diagnostics.rs   # Diagnostic computation
├── excel_model/
│   ├── types.rs         # Excel Object Model type definitions
│   └── context.rs       # workbook-context.json reader
├── completion.rs        # textDocument/completion
├── hover.rs             # textDocument/hover
├── signature_help.rs    # textDocument/signatureHelp
├── definition.rs        # textDocument/definition
├── references.rs        # textDocument/references
├── rename.rs            # textDocument/rename
├── document_highlight.rs# textDocument/documentHighlight
├── document_symbol.rs   # textDocument/documentSymbol
├── workspace_symbol.rs  # workspace/symbol
├── folding_range.rs     # textDocument/foldingRange
├── code_action.rs       # textDocument/codeAction
├── formatting.rs        # textDocument/formatting
├── inlay_hint.rs        # textDocument/inlayHint
├── call_hierarchy.rs    # textDocument/prepareCallHierarchy + callHierarchy/*
└── vba_builtins.rs      # VBA keywords and built-in functions
```

Tech stack: `tower-lsp`, `logos`, `la-arena`, `smol_str`, `dashmap`, `tokio`.

## Design notes

- **Case-insensitive** — identifiers and keywords compare case-insensitively
  (standard VBA semantics).
- **Option Explicit** — tracked per module; controls undeclared-variable warnings.
- **Excel Object Model** — hardcoded MVP types (see Completion above). Unknown
  members fall through to generic `Variant`.
- **Position encoding** — UTF-16, advertised in `ServerCapabilities`.

## License

[MIT](./LICENSE)
