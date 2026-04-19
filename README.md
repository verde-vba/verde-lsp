# verde-lsp

VBA Language Server for [Verde](https://github.com/verde-vba/verde).

## Features (MVP)

- **Go to Definition** — procedures, variables, types
- **Completion** — keywords, built-in functions, symbols, sheet/table names
- **Hover** — variable types, procedure signatures
- **Diagnostics** — undeclared variables (Option Explicit), parse errors
- **Rename** — symbol renaming across module

## Usage

Communicates via stdio (LSP protocol).

```bash
# Build
just build

# Run
just run

# Test
just test
```

## Architecture

```
src/
├── main.rs              # Entry point (stdio LSP server)
├── server.rs            # LSP protocol handler
├── parser/
│   ├── lexer.rs         # Logos-based VBA tokenizer
│   ├── ast.rs           # AST node definitions (la-arena)
│   └── parse.rs         # Recursive descent parser
├── analysis/
│   ├── symbols.rs       # Symbol table construction
│   ├── resolve.rs       # Name resolution
│   └── diagnostics.rs   # Diagnostic computation
├── completion.rs        # Completion provider
├── hover.rs             # Hover provider
├── definition.rs        # Go-to-definition
├── rename.rs            # Rename provider
├── excel_model/
│   ├── types.rs         # Excel Object Model type definitions
│   └── context.rs       # workbook-context.json reader
└── vba_builtins.rs      # VBA keywords and built-in functions
```

## License

MIT
