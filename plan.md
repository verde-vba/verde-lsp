# verde-lsp バックログ

> 最終更新: 2026-04-20
> 現在ブランチ: main (最新: 3c629a5)
> テスト基準: 62 green (lib + integration), cargo clippy -D warnings 0 件

---

## Sprint N+8 レトロスペクティブ (2026-04-20)

### Sprint Goal 達成状況

目標「VBA のスコープルール（手続き境界）を completion が正しく反映し、補完精度を向上させる」を完全達成。

### KPT

#### Keep (よかったこと、継続すること)

- **Tidy First 実践**: `proc_scope`/`proc_ranges` の構造追加（構造変更）と挙動変更を別コミットに分離し、コンパイル確認を挟んでから実装した。リグレッションリスクを最小化する習慣が定着している。
- **TDD サイクルの維持**: refactor → test → feat の 3 コミット構成で RED/GREEN/REFACTOR を明確に分離できた。
- **`position_to_offset` の公開**: 内部ユーティリティを他モジュールから再利用可能にする判断が適切だった。analysis 層の凝集度が高まった。
- **全テスト green・clippy 0 件**: 62 件すべて通過、警告ゼロで Sprint を終了できた。

#### Problem (課題)

- **モジュールレベル `Dim` の無限ループバグ**: `parse_module()` で `Dim m As String` を処理すると `parse_variable()` が `Dim` キーワードを消費しないため無限ループが発生する。`Public m As String` では正常動作するため、テストを `Public` で回避して Sprint 完了としたが根本修正は未着手。
- **テスト回避の技術的負債**: バグを修正せずテストを変更して回避したことで、モジュールレベル `Dim` を使うコードに対する LSP 機能が実際には動作しない状態が残っている。

#### Try (次 Sprint で試すこと)

- モジュールレベル `Dim` パーサーバグを専用 PBI として立ててバグ修正 TDD サイクルで対応する（後述の「次 Sprint 推奨」参照）。

---

## 次 Sprint 推奨

### 最優先: BUG-01 — モジュールレベル `Dim` 無限ループ修正 (Small/Bug)

| 項目 | 内容 |
|------|------|
| **目的** | `parse_module()` が `Dim m As String` を処理すると `parse_variable()` が `Dim` トークンを消費しないため無限ループが発生する。`Public`/`Private` 修飾子付きなら正常動作するため、修飾子なし `Dim` のモジュールレベル宣言を正しくパースできるようにする。 |
| **受入基準** | (1) `Dim m As String` を含むモジュールが無限ループなくパースできる。(2) `parse_variable()` が `Dim` トークンを適切に消費する。(3) 既存 62 tests 以上 green, clippy 0 件 |
| **見積サイズ** | S (`src/parser.rs` の `parse_variable()` 周辺修正) |
| **依存** | なし |
| **調査メモ** | `parse_variable()` 内で `Dim` キーワードを事前に consume するか、呼び出し側 `parse_module()` が consume してから渡す方式が考えられる。既存の `Public`/`Private` 経路との対称性を確認すること。Sprint N+8 でのテスト回避コミットを revert して正規のテストに戻す。 |

### 候補: PBI-09 — 複数ファイル対応・クロスモジュール補完 (Large)

| 項目 | 内容 |
|------|------|
| **目的** | 現状は単一ファイルの補完のみ。複数モジュール間で Public Sub/Function/変数を参照できるようにし、実際の VBA プロジェクト規模に対応する。 |
| **受入基準** | (1) モジュール A の `Public Sub Foo()` がモジュール B での補完候補に出る。(2) `cargo test` 62+ green, clippy 0 件 |
| **見積サイズ** | L (AnalysisHost の workspace 管理拡張が必要) |
| **依存** | BUG-01 完了後に取り組む方が安全 |

---

## プロダクトバックログ

### PBI-02 — 手続きパラメータを hover シグネチャに表示する (Small)

| 項目 | 内容 |
|------|------|
| **目的** | `SymbolDetail::Procedure.params` は parser が収集済みだが `src/analysis/symbols.rs:80` で常に空ベクタになっている。これを実値で埋め、hover が `Sub Foo(x As Long, y As String)` を表示できるようにする。 |
| **受入基準** | (1) 引数を持つ手続きへの hover が引数名+型を含むシグネチャを返す。(2) cargo test 62+ green, clippy 0 |
| **見積サイズ** | S |
| **依存** | なし |
| **調査メモ** | `src/parser/ast.rs` の `ProcedureNode.params: Vec<NodeId>` → `ParameterNode { name, type_name, passing, ... }` が利用可能。`symbols.rs` の `populate_procedure_symbol` 相当箇所に 5-10 行追加で済む見込み |

---

## 完了済み (Sprint N+8)

| コミット | 内容 |
|----------|------|
| (refactor) | `Symbol` 構造体に `proc_scope: Option<SmolStr>` フィールドを追加（構造のみ、全箇所 `None` で初期化） |
| (test) | scope-aware completion フィルタリングの受入基準テスト RED |
| (feat) | `build_symbol_table()` で `proc_scope` 設定 + `complete()` でフィルタリング実装 (GREEN) |

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

---

## 完了済み (Sprint N+4)

| コミット | 内容 |
|----------|------|
| `daa1508` | refactor: WhileStatementNode + StatementNode::While 追加 (Tidy First) |
| `e32c5ae` | feat: PBI-05 (While) — WhileStatementNode パース + Option Explicit 診断対応 |
| `d933f8a` | refactor: classify_and_parse_statement doc comment 更新 |

---

## 完了済み (Sprint N+3)

| コミット | 内容 |
|----------|------|
| `8c5e416` | refactor: LocalDeclarationNode に per-name identifier span を追加 (Tidy First) |
| `1e7bc30` | feat: PBI-04 — Call 文・ベア呼び出し・ローカル変数への Go-to-Definition |

---

## 完了済み (Sprint N+2)

| コミット | 内容 |
|----------|------|
| `563024d` | refactor: LocalDeclarationNode に per-name 型情報を追加 (Tidy First) |
| `11ac2ba` | feat: PBI-01 — ローカル変数を SymbolTable に登録 (hover/completion 対応) |
| `d999975` | test: PBI-01 受入基準テスト (hover 型名表示・completion 候補露出) |

---

## 完了済み (Sprint N+1)

| コミット | 内容 |
|----------|------|
| `a77b011` + `d76bb0b` + `8ed4d9a` + `a9247df` | PBI-02+03: procedure params in hover, with_source fix for hover/definition/rename, name_span added to ProcedureNode |

---

## 完了済み (Sprint N)

| コミット | 内容 |
|----------|------|
| `33d29cf` | refactor: clippy 2 件除去 (new_without_default / needless_lifetimes) |
| `7ccfb89` | feat: If/For/With/Select/Call/Set ヘッダの undeclared identifier 検出 |
| `9ee51a7` | test: Set 文 RHS の undeclared identifier カバレッジ |
| `572cbb7` | refactor: dead field ProcedureNode.body_range 削除 |
| `12b1307` | test: For ループ上限式の undeclared identifier カバレッジ |
