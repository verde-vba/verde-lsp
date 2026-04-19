# verde-lsp — VBA Language Server

# Run all checks
check:
    cargo check

# Run tests
test:
    cargo test

# Format code
fmt:
    cargo fmt

# Lint
clippy:
    cargo clippy -- -D warnings

# Build release
build:
    cargo build --release

# Run LSP server (stdio)
run:
    cargo run
