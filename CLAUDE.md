# verde-lsp — VBA Language Server

## Overview

Rust-based LSP for VBA, used by Verde desktop app and (future) VS Code extension.
Communicates via stdio. Parses VBA source files (.bas, .cls, .frm).

## Tech Stack

- `tower-lsp` — LSP protocol
- `logos` — lexer generator
- `la-arena` — arena-allocated AST nodes
- `smol_str` — interned strings
- `dashmap` — concurrent file state

## Commands

```bash
mise install          # Install Rust + just
just check            # cargo check
just test             # cargo test
just build            # cargo build --release
just fmt              # cargo fmt
just clippy           # cargo clippy
```

## Architecture

- Parser: lexer (logos) → tokens → recursive descent → AST (la-arena)
- Analysis: AST → symbol table → resolution/diagnostics
- LSP features call into analysis via AnalysisHost

## Key Design

- Case-insensitive (VBA is case-insensitive)
- Option Explicit: per-module, controls undeclared variable warnings
- Excel Object Model: hardcoded MVP types (Range, Worksheet, Workbook, Application)
- workbook-context.json: provides sheet/table/named range info for completion
