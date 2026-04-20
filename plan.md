# verde-lsp バックログ

> 最終更新: 2026-04-21 (Sprint N+16 完了)
> 現在ブランチ: main
> テスト基準: 75 green (lib 36 + integration 39), cargo clippy -D warnings 0 件

---

## 次 Sprint 推奨 (Sprint N+17)

**Sprint Goal 候補**: 新規 PBI が Refinement 後に Ready になり次第実行

### PBI-10 — For Each ループ変数の undeclared 誤検出除外 ✅ Won't Do (Already Working)

`does_not_warn_on_for_each_with_declared_items` が既に green のため実作業不要。
`True`/`False`/`Nothing`/`Null`/`Empty` は lexer が専用トークン (`Token::True` 等) として処理するため
`scan_expression_tokens` の Identifier チェックをバイパス — builtins への追加も不要。

### PBI-11 — workbook-context.json シート名補完 (Medium) 🔲 Backlog (Not Ready)

| 項目 | 内容 |
|------|------|
| **目的** | workbook-context.json のシート名・テーブル名・名前付き範囲を補完候補に追加。 |
| **背景** | CLAUDE.md に「workbook-context.json: provides sheet/table/named range info for completion」と記載あり。未実装。 |
| **Not Ready 理由** | `VbaLanguageServer` が workspace root (`InitializeParams.root_uri`) を未取得。JSON 読み込み前に root_uri フィールド追加が前提作業として必要。 |
| **受入基準** | workbook-context.json からシート名が補完候補に現れること。68+ green, clippy 0。 |
| **見積サイズ** | M |
| **依存** | workspace root 取得 (Tidy First コミット要) |

### PBI-12 — 修飾呼び出し `ModuleA.Foo` の `ModuleA` undeclared 誤検出除外 (Small) ✅ Ready

| 項目 | 内容 |
|------|------|
| **目的** | `Call ModuleA.Foo` のように module 名で修飾した呼び出しで `ModuleA` が undeclared として誤検出されないようにする。 |
| **背景** | `scan_expression_tokens` の `after_dot` は `Foo` をスキップするが、`ModuleA` 自体は identifier として検査対象になる。モジュール名は symbol table になく URI から取得できる。 |
| **実装方針** | `AnalysisHost::diagnostics` で `self.files.keys()` を走査し、URI の最終セグメントから拡張子を除いたモジュール名 (`uri.path_segments().next_back()?.split('.').next()`) を lowercase で `cross_module_names` に追加。別途 `collect_other_module_names(&self, current_uri: &Url)` ヘルパーを追加して責務を分離。 |
| **受入基準** | (1) 2 ファイル workspace で `ModuleA.Foo` 呼び出しが `ModuleA` undeclared 警告を出さない。(2) 本来の未宣言識別子は引き続き検出。(3) 70+ green, clippy 0 件。 |
| **見積サイズ** | S |
| **依存** | PBI-09c (完了済み) |

---

## Sprint N+12 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-09c クロスモジュール diagnostics を届ける」を完全達成。

### KPT

#### Keep
- `diagnostics::compute` への `cross_module_names: &HashSet<String>` 追加という最小シグネチャ変更が効果的だった。`all_public_symbols_from_other_files()` を lowercase 化して渡す 3 行のみで GREEN。
- REFACTOR 評価で「3 呼び出し元がそれぞれ異なる型を要求するため helper 不要」を迷わず判断できた。
- RED → GREEN → 変更なし (REFACTOR) → 変更なし (Tidy After) の 2 コミット構成が Sprint N+11 の "Tidy After 変更不要" パターンの再確認となった。

#### Problem
- `check_option_explicit` の引数が 4 個になり、将来さらに増える場合は構造体化を検討すべき。

#### Try
- `check_option_explicit` の引数が 5 個を超えた時点で `DiagnosticsContext` 構造体を導入する。

---

## Sprint N+16 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-14 workbook-context.json 自動再読み込み」を完全達成。

### KPT

#### Keep
- `reload_workbook_context_from_path` を同期ヘルパーに統一したことで `initialized` の `tokio::fs::read_to_string` + `.await` が不要になり、`RwLockReadGuard` 問題も回避できた。
- `register_capability(vec![...])` が `RegistrationParams` wrapper 不要だと API 確認で即発見。コンパイルエラー 1 回で解決。

#### Problem
- `did_change_watched_files` ハンドラのサーバー統合テストは未実装。`reload_workbook_context_from_path` の単体テストのみ。

#### Try
- サーバー統合テストが必要な場合は mock client を使う tower-lsp のテストユーティリティを調査する。

---

## 完了済み (Sprint N+16)

| コミット | 内容 |
|----------|------|
| `b664801` | docs(scrum): PBI-14 Refinement |
| `e434b2d` | test: reload_workbook_context_from_path RED テスト (PBI-14) |
| `6f59b1a` | feat: workbook-context.json 自動再読み込み — reload_workbook_context_from_path + didChangeWatchedFiles (PBI-14) |

---

## Sprint N+15 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-13 workbook-context.json tables/named_ranges 補完拡張」を完全達成。

### KPT

#### Keep
- REFACTOR で `push_named_items` ヘルパーを抽出。27 行 → 3 行呼び出しに圧縮。「3 ループ同パターン = 抽出対象」の判断が迅速だった。
- `..Default::default()` を既存テストに追加するだけで後方互換を維持できた。`Default` derive の価値を再確認。

#### Problem
- `workbook-context.json` の変更を LSP 再起動なしに反映する方法がない。`workspace/didChangeWatchedFiles` 対応が次の課題。

#### Try
- PBI-14 として `workspace/didChangeWatchedFiles` で JSON 再読み込みを実装する。

---

## 完了済み (Sprint N+15)

| コミット | 内容 |
|----------|------|
| `7c8c098` | docs(scrum): PBI-13 Refinement |
| `2e594c2` | test: workbook tables/named_ranges 補完 RED テスト (PBI-13) |
| `1e1836b` | feat: workbook-context.json tables/named_ranges を completion 候補に追加 (PBI-13) |
| `1c946e2` | refactor: push_named_items ヘルパー抽出 (PBI-13 REFACTOR) |

---

## Sprint N+14 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-11 workbook-context.json シート名補完」を完全達成。

### KPT

#### Keep
- Tidy First (構造追加) → RED → GREEN の 4 コミット構成が Medium PBI を安全に完遂する手順として有効。
- `std::sync::RwLockReadGuard` を `.await` 前に `clone()` して drop するパターン (`let root = self.root_uri.read().unwrap().clone();`) がコンパイルエラー 1 件で即座に解決できた。コメントで理由を残し将来のメンテに備えた。
- `WorkbookContext` に `Default` derive を付けたことで `set_workbook_context` を呼ばない既存テストがそのまま動作した。

#### Problem
- `workbook-context.json` の `tables`/`named_ranges` フィールドは定義済みだが未使用。MVP 完了後に補完拡張が必要。

#### Try
- PBI-13 として tables/named_ranges を completion に追加する Small スプリントを計画する。

---

## 完了済み (Sprint N+14)

| コミット | 内容 |
|----------|------|
| `4b4bfba` | refactor: WorkbookContext 構造体追加 + AnalysisHost/VbaLanguageServer 構造準備 (Tidy First PBI-11) |
| `9b47ead` | test: workbook-context.json シート名補完 RED テスト (PBI-11) |
| `75afdda` | feat: workbook-context.json シート名を completion 候補に追加 (PBI-11) |

---

## Sprint N+13 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-12 ModuleA.Foo 修飾呼び出しの ModuleA undeclared 誤検出除外」を完全達成。

### KPT

#### Keep
- `collect_other_module_names` という専用ヘルパーで「URI → モジュール名抽出」の責務を分離。`diagnostics` が膨らまず読みやすい。
- `filter_map` で `path_segments().next_back().split('.').next()` を繋ぐ Rust イディオムが URI 操作を 1 行で表現できた。
- REFACTOR 評価: `files.iter()` を 2 回走査するが型が異なるため統合は過剰と判断 → 変更なし。

#### Problem
- `path_segments()` は `file://` URI でのみ正しく動作する。`untitled:` や `vscode-vfs:` スキームでは意図しない結果になりうる。

#### Try
- URI スキームが `file` でない場合は module 名抽出をスキップするガードを将来追加する。

---

## 完了済み (Sprint N+13)

| コミット | 内容 |
|----------|------|
| `ade8233` | test: ModuleA.Foo 修飾呼び出し RED テスト 2 件 (PBI-12) |
| `30e7a03` | feat: ModuleA.Foo 修飾呼び出しの ModuleA undeclared 誤検出除外 (PBI-12) |

---

## 完了済み (Sprint N+12)

| コミット | 内容 |
|----------|------|
| `8656345` | test: クロスモジュール Option Explicit RED テスト 2 件 (PBI-09c) |
| `95b5d03` | feat: クロスモジュール diagnostics — 他モジュール Public シンボルを undeclared 検出から除外 (PBI-09c) |

---

## Sprint N+11 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-09b クロスモジュール hover/goto-def を届ける」を完全達成。

### KPT

#### Keep
- `find_public_symbol_in_other_files()` の設計判断: URI を返す必要があるため `all_public_symbols_from_other_files()` と分離した。2 つのメソッドが責務で明確に分かれた（全件 vs 名前検索）。
- `symbol_to_hover` ヘルパー抽出が GREEN フェーズで自然に発生し、REFACTOR フェーズのコミット不要に繋がった。Tidy First/After より GREEN 時点でのヘルパー抽出が効率的なケースがある。
- Phase 5/6 で「変更不要」の判断を迷わず下せた — 「3 行以上の同一ロジック」という明確な基準が効いた。

#### Problem
- RED テスト 2 件の word 抽出位置（col=4, col=9）を試行錯誤せずに確定できなかった可能性。テスト設計時に source 文字列のオフセットを手計算する手間が発生した。

#### Try
- RED テスト記述時に position の根拠（"Call Foo" の "F" は col=9）をコメントで残す習慣をつける。

---

## 完了済み (Sprint N+11)

| コミット | 内容 |
|----------|------|
| `c217de0` | test: クロスモジュール hover/goto-def RED テスト 2 件 (PBI-09b) |
| `73c81f0` | feat: クロスモジュール hover/goto-def — find_public_symbol_in_other_files + fallback 追加 (PBI-09b) |

---

## Sprint N+10 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-09a クロスモジュール補完 MVP を届ける」を完全達成。

### KPT

#### Keep
- `AnalysisHost.files` が既に `DashMap<Url, FileAnalysis>` であることをリファインメント段階で確認し、Large → Small への見積修正が正確だった。
- `proc_scope.is_none()` フィルタでモジュールレベルの Public シンボルのみを横断対象にする設計判断が明快。
- TDD サイクル: RED test → GREEN (all_public_symbols_from_other_files) → REFACTOR (symbol_kind_to_completion_kind 抽出) → Tidy After (_value デストラクチャ化) の 4 コミット構成が整然としていた。

#### Problem
- `symbol_kind_to_completion_kind` の match ブロックが一時的に 2 箇所に存在した (Phase 4 → Phase 5 で解消)。GREEN フェーズで最初からヘルパーを意識すれば 1 コミット削減できた。

#### Try
- GREEN 実装時に既存の match/map パターンを再利用するか、即座にヘルパー化するかを意識する。

---

## 完了済み (Sprint N+10)

| コミット | 内容 |
|----------|------|
| `c15d5ea` | refactor: let _value で EnumMember 値を明示的に未使用とマーク (Tidy First) |
| `1fa0c8e` | test: completion_includes_public_symbols_from_other_files (RED) |
| `9628f1a` | feat: クロスモジュール補完 — all_public_symbols_from_other_files + complete() 連結 (PBI-09a) |
| `3e474f6` | refactor: symbol_kind_to_completion_kind ヘルパー抽出で重複除去 |
| `b19e09e` | refactor: _value をデストラクチャパターンに移動し dead statement 除去 |

---

## Sprint N+9 レトロスペクティブ (2026-04-20)

### Sprint Goal 達成状況

目標「裸の `Dim` 宣言をモジュールレベルで正しくパースし、無限ループバグを根絶する」を完全達成。

### KPT

#### Keep
- `#[ignore]` 付きテストで RED を安全にコミットし、その後 fix + `#[ignore]` 削除で GREEN とする 2 コミット戦略が機能した。
- 修正は `parse_variable()` 内 5 行以内。既存の `Const` consume パターンとの対称設計で影響範囲が明確。

#### Problem
- Sprint N+8 の技術的負債（`Public` 回避）がそのまま次スプリントの作業に繋がった。より早い段階でバグを分離・記録すべきだった。

#### Try
- バグ発見時は即座に `#[ignore]` 付き RED テストを書いてコミットし、後続 Sprint で修正する流れを定着させる。

---

## 完了済み (Sprint N+9)

| コミット | 内容 |
|----------|------|
| `085e99b` | test: BUG-01 module-level bare Dim parse test (red, #[ignore]) |
| `5e5eebb` | fix: parse_variable が Dim キーワードを consume するよう修正 |
| `e579320` | test: completion_module_var_visible_everywhere を Dim 正規形式に戻す |

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

### 候補: PBI-09 — 複数ファイル対応・クロスモジュール補完 (→ PBI-09a に Small 分割済み)

元 Large PBI。`AnalysisHost.files` が既に `DashMap<Url, FileAnalysis>` であることが判明し、Sprint N+10 でのリファインメントにより PBI-09a として Small 化。残余分（hover/goto-def への拡張など）は PBI-09b 以降で対応。

---

## プロダクトバックログ

| PBI | タイトル | サイズ | 状態 |
|-----|----------|--------|------|
| PBI-09c | クロスモジュール diagnostics (undeclared 誤検出除外) | S | **Done** |
| PBI-10 | For Each ループ変数 undeclared 誤検出除外 | S | **Won't Do** (already working) |
| PBI-11 | workbook-context.json シート名補完 | M | **Done** |
| PBI-13 | workbook-context.json tables/named_ranges 補完拡張 | XS | **Done** |
| PBI-14 | workbook-context.json 自動再読み込み (didChangeWatchedFiles) | S | **Done** |
| PBI-12 | 修飾呼び出し ModuleA.Foo の ModuleA undeclared 誤検出除外 | S | **Done** |

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
