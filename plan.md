# verde-lsp バックログ

> 最終更新: 2026-04-20  
> 現在ブランチ: main (最新: a20b0cf)  
> テスト基準: 67 green, cargo clippy -D warnings 0 件

---

## 次 Sprint 推奨 (Sprint N+8)

**新規 PBI を検討** — PBI-06+07 完了により rename が call site まで追跡できるようになった。  
候補: (C) シンボルスコープ（手続き境界）を考慮した completion フィルタリング（SymbolTable に proc_scope を追加が必要）、または新規 PBI 検討。

---

## プロダクトバックログ

### PBI-01 — ローカル変数をシンボルテーブルに登録する

| 項目 | 内容 |
|------|------|
| **目的** | `Dim x As String` など手続き内 LocalDeclaration を `SymbolTable` に登録し、hover/completion/definition がその識別子を認識できるようにする。現状 diagnostics 層だけが知っており、LSP 機能として露出されていない。 |
| **受入基準** | (1) 手続き内 Dim 変数へのカーソル hover が型名を返す。(2) 同手続き内の completion 候補にその変数が現れる。(3) 既存 36 tests 以上 green, clippy 0 |
| **見積サイズ** | M (symbols.rs + hover.rs + completion.rs 改修) |
| **依存** | なし |
| **調査メモ** | `src/analysis/symbols.rs:137-138` の `_ => {}` が LocalDeclarationNode を無視。diagnostics.rs の `local_declared` 収集ロジックを参考に同等処理を symbol 登録として移植できる可能性あり |

---

### PBI-02 — 手続きパラメータを hover シグネチャに表示する (Small)

| 項目 | 内容 |
|------|------|
| **目的** | `SymbolDetail::Procedure.params` は parser が収集済みだが `src/analysis/symbols.rs:80` で常に空ベクタになっている。これを実値で埋め、hover が `Sub Foo(x As Long, y As String)` を表示できるようにする。 |
| **受入基準** | (1) 引数を持つ手続きへの hover が引数名+型を含むシグネチャを返す。(2) cargo test 36+ green, clippy 0 |
| **見積サイズ** | S (symbols.rs 単体、`ProcedureNode.params` フィールドを `ParameterNode` から読む) |
| **依存** | なし (PBI-01 と並行可能) |
| **調査メモ** | `src/parser/ast.rs` の `ProcedureNode.params: Vec<NodeId>` → `ParameterNode { name, type_name, passing, ... }` が利用可能。symbols.rs の populate_procedure_symbol 相当箇所に 5-10 行追加で済む見込み |

---

### PBI-03 — rename の位置取得バグを修正する (Small/Bug)

| 項目 | 内容 |
|------|------|
| **目的** | `src/rename.rs:14` が `find_word_at_position` に `""` を渡しているため rename が実質的に broken。正しいソース文字列を渡し、カーソル位置の識別子を対象にリネームできるようにする。 |
| **受入基準** | (1) カーソルを手続き名または変数名に置いて rename を実行すると、同一ファイル内の全出現箇所が新名称に置換されたワークスペース編集が返る。(2) cargo test 36+ green, clippy 0 |
| **見積サイズ** | S (rename.rs 単体修正。ロジックは既存 hover/definition と対称) |
| **依存** | なし。PBI-01 完了後は手続き内変数の rename も自動で恩恵を受ける |
| **調査メモ** | `src/rename.rs:14` の空文字列渡しを `source` 変数に置き換えるだけ。テストは既存 option_explicit パターンを踏襲 |

---

### PBI-04 — Call 文からの Go-to-Definition を実装する

| 項目 | 内容 |
|------|------|
| **目的** | `definition.rs` はカーソル位置を無視し module-level 線形検索だけ行う。Call 文上の識別子がシンボルテーブルの手続き定義に飛べるようにし、コードナビを実用化する。 |
| **受入基準** | (1) Call 文 (`Call Foo` / `Foo arg`) の手続き名上で Go-to-Definition を実行すると対応する Sub/Function 宣言行にジャンプする。(2) cargo test 36+ green, clippy 0 |
| **見積サイズ** | M (definition.rs + analysis 層: 位置→トークン→識別子解決パイプライン) |
| **依存** | PBI-02 (パラメータ整備でシンボル精度向上) が完了していると実装品質が上がる |
| **調査メモ** | `StatementNode::Call(CallStatementNode { tokens, .. })` のトークン列からカーソル位置に対応する Identifier を特定し `find_symbol_by_name` に渡す。position→byte-offset 変換が必要 |

---

~~**PBI-05b** — Do / ReDim StatementNode 化 → Sprint N+5 完了~~

---

## 完了済み (Sprint N+7)

| コミット | 内容 |
|----------|------|
| `deca407` | feat: rename が call site も WorkspaceEdit に含むようになった (find_all_word_occurrences) |
| `a20b0cf` | refactor: eq_ignore_ascii_case で文字列確保を除去 |

---

## 完了済み (Sprint N+6)

| コミット | 内容 |
|----------|------|
| `4bb8c62` | refactor: ExitStatementNode / GoToStatementNode / OnErrorStatementNode 追加 (Tidy First) |
| `f52f25b` | feat: Exit Sub/Function/For/Do → StatementNode::Exit |
| `53ac2dd` | feat: GoTo → StatementNode::GoTo |
| `1f48120` | feat: On Error → StatementNode::OnError |
| `1525c2e` | refactor: classify_and_parse_statement doc comment 更新 |

---

## 完了済み (Sprint N+5)

| コミット | 内容 |
|----------|------|
| `4e8aa3c` | refactor: DoStatementNode + RedimStatementNode 追加 (Tidy First) |
| `89af46c` | feat: PBI-05b — Do/ReDim の StatementNode 化、DeclKind::ReDim 除去 |

## 完了済み (Sprint N+4)

| コミット | 内容 |
|----------|------|
| `daa1508` | refactor: WhileStatementNode + StatementNode::While 追加 (Tidy First) |
| `e32c5ae` | feat: PBI-05 (While) — WhileStatementNode パース + Option Explicit 診断対応 |
| `d933f8a` | refactor: classify_and_parse_statement doc comment 更新 |

## 完了済み (Sprint N+3)

| コミット | 内容 |
|----------|------|
| `8c5e416` | refactor: LocalDeclarationNode に per-name identifier span を追加 (Tidy First) |
| `1e7bc30` | feat: PBI-04 — Call 文・ベア呼び出し・ローカル変数への Go-to-Definition |

## 完了済み (Sprint N+2)

| コミット | 内容 |
|----------|------|
| `563024d` | refactor: LocalDeclarationNode に per-name 型情報を追加 (Tidy First) |
| `11ac2ba` | feat: PBI-01 — ローカル変数を SymbolTable に登録 (hover/completion 対応) |
| `d999975` | test: PBI-01 受入基準テスト (hover 型名表示・completion 候補露出) |

## 完了済み (Sprint N+1)

| コミット | 内容 |
|----------|------|
| `a77b011` + `d76bb0b` + `8ed4d9a` + `a9247df` | PBI-02+03: procedure params in hover, with_source fix for hover/definition/rename, name_span added to ProcedureNode |

## 完了済み (Sprint N)

| コミット | 内容 |
|----------|------|
| `33d29cf` | refactor: clippy 2 件除去 (new_without_default / needless_lifetimes) |
| `7ccfb89` | feat: If/For/With/Select/Call/Set ヘッダの undeclared identifier 検出 |
| `9ee51a7` | test: Set 文 RHS の undeclared identifier カバレッジ |
| `572cbb7` | refactor: dead field ProcedureNode.body_range 削除 |
| `12b1307` | test: For ループ上限式の undeclared identifier カバレッジ |
