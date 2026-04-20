# verde-lsp バックログ

> 最終更新: 2026-04-21 (Sprint N+20 完了)
> 現在ブランチ: main
> テスト基準: 83 green (lib 36 + integration 47), cargo clippy -D warnings 0 件

---

## 次 Sprint 推奨 (Sprint N+21)

**Sprint Goal 候補**: 新規 PBI を Refinement 後に実行

---

## プロダクトバックログ

| PBI | タイトル | サイズ | 状態 |
|-----|----------|--------|------|
| PBI-09a | クロスモジュール補完 MVP | S | **Done** |
| PBI-09b | クロスモジュール hover/goto-def | S | **Done** |
| PBI-09c | クロスモジュール diagnostics (undeclared 誤検出除外) | S | **Done** |
| PBI-10 | For Each ループ変数 undeclared 誤検出除外 | S | **Won't Do** (already working) |
| PBI-11 | workbook-context.json シート名補完 | M | **Done** |
| PBI-12 | 修飾呼び出し ModuleA.Foo の ModuleA undeclared 誤検出除外 | S | **Done** |
| PBI-13 | workbook-context.json tables/named_ranges 補完拡張 | XS | **Done** |
| PBI-14 | workbook-context.json 自動再読み込み (didChangeWatchedFiles) | S | **Done** |
| PBI-15 | textDocument/references プロバイダ実装 | XS | **Done** |
| PBI-16 | textDocument/references クロスファイル拡張 | S | **Done** |
| PBI-17 | textDocument/rename クロスファイル拡張 | S | **Done** |
| PBI-18 | rename guard cross-module フォールバック | S | **Done** |

---

## Sprint N+20 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-18 rename guard cross-module フォールバック」を完全達成。

### KPT

#### Keep
- `found_locally || found_cross` のシンプルな論理和で guard 拡張。6 行の変更でテスト通過。
- PBI-16 → PBI-17 → PBI-18 と小さな incremental 拡張を重ねた結果、クロスモジュール LSP 機能 (completion/hover/goto-def/diagnostics/references/rename) が全て揃った。

#### Problem
- guard の cross-module チェックが `find_public_symbol_in_other_files` (Public シンボルのみ) に限定されている。Private シンボルの cross-file rename は未対応。

#### Try
- Private シンボルが cross-file で参照されることは VBA では稀なので、現時点では許容範囲と判断。

---

## Sprint N+19 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-17 textDocument/rename クロスファイル拡張」を完全達成。

### KPT

#### Keep
- `all_file_sources()` が `references` (PBI-16) と `rename` (PBI-17) の両方で活躍。PBI-16 の Tidy First がここで回収された。
- word 取得 + guard を現在ファイルのみで行い、テキスト検索のみ全ファイルに拡張するという責務分割が clean。

#### Problem
- rename の guard (`find_symbol_by_name`) は現在ファイルの symbol のみ確認する。他ファイルで定義されたシンボルを他ファイルの call site から rename できない → PBI-18 で解決。

#### Try
- `find_references` の結果を `rename` でも活用できるか検討。

---

## Sprint N+18 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-16 textDocument/references クロスファイル拡張」を完全達成。

### KPT

#### Keep
- `all_file_sources()` という汎用ヘルパーで references 以外の将来用途にも使える API を定義。
- PBI-15 の変更差分が `src/references.rs` 内に閉じており、`find_references` の差分は +11/-14 行と小さかった。

#### Problem
- `rename` はまだ single-file のみ → PBI-17 で解決。

---

## Sprint N+17 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-15 textDocument/references プロバイダ実装」を完全達成。

### KPT

#### Keep
- `rename.rs` のパターン (word → occurrences → lsp_range) をそのまま `references.rs` に転用し 20 行以内で完結。XS 見積が正確だった。

#### Problem
- `find_references` はシングルファイルのみ → PBI-16 で解決。

---

## Sprint N+16 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-14 workbook-context.json 自動再読み込み」を完全達成。

### KPT

#### Keep
- `reload_workbook_context_from_path` を同期ヘルパーに統一したことで `initialized` の `tokio::fs::read_to_string` + `.await` が不要になり、`RwLockReadGuard` 問題も回避できた。
- `register_capability(vec![...])` が `RegistrationParams` wrapper 不要だと API 確認で即発見。

#### Problem
- `did_change_watched_files` ハンドラのサーバー統合テストは未実装。`reload_workbook_context_from_path` の単体テストのみ。

---

## Sprint N+15 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-13 workbook-context.json tables/named_ranges 補完拡張」を完全達成。

### KPT

#### Keep
- REFACTOR で `push_named_items` ヘルパーを抽出。27 行 → 3 行呼び出しに圧縮。
- `..Default::default()` を既存テストに追加するだけで後方互換を維持できた。

---

## Sprint N+14 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-11 workbook-context.json シート名補完」を完全達成。

### KPT

#### Keep
- Tidy First (構造追加) → RED → GREEN の 4 コミット構成が Medium PBI を安全に完遂。
- `std::sync::RwLockReadGuard` を `.await` 前に `clone()` して drop するパターンがコンパイルエラー 1 件で解決。

#### Problem
- `workbook-context.json` の `tables`/`named_ranges` フィールドは未使用 → PBI-13 で解決。

---

## Sprint N+13 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-12 ModuleA.Foo 修飾呼び出しの ModuleA undeclared 誤検出除外」を完全達成。

### KPT

#### Keep
- `collect_other_module_names` という専用ヘルパーで「URI → モジュール名抽出」の責務を分離。
- `filter_map` で URI 操作を 1 行で表現。

#### Problem
- `path_segments()` は `file://` URI でのみ正しく動作。他スキームでの guard が将来課題。

---

## Sprint N+12 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-09c クロスモジュール diagnostics を届ける」を完全達成。

### KPT

#### Keep
- `diagnostics::compute` への `cross_module_names: &HashSet<String>` 追加という最小シグネチャ変更が効果的。
- REFACTOR 評価で「3 呼び出し元がそれぞれ異なる型を要求するため helper 不要」を迷わず判断できた。

#### Problem
- `check_option_explicit` の引数が 4 個になり、5 個超で `DiagnosticsContext` 構造体を導入する。

---

## 完了済み Sprint 要約 (N 〜 N+11)

| Sprint | PBI | 主要コミット |
|--------|-----|-------------|
| N | If/For/With/Select/Call/Set の undeclared 検出 | 7ccfb89 |
| N+1 | PBI-02+03: procedure params hover, name_span | a77b011 |
| N+2 | PBI-01: ローカル変数 SymbolTable 登録 | 11ac2ba |
| N+3 | PBI-04: Call/bare/local goto-def | 1e7bc30 |
| N+4 | PBI-05 While: WhileStatementNode | e32c5ae |
| N+5 | PBI-05b Do/ReDim: StatementNode 化 | 89af46c |
| N+6 | Exit/GoTo/OnError: StatementNode 追加 | f52f25b |
| N+7 | rename call site 対応 | deca407 |
| N+8 | scope-aware completion (proc_scope) | build_symbol_table |
| N+9 | BUG-01 module-level Dim パーサー修正 | 5e5eebb |
| N+10 | PBI-09a クロスモジュール補完 | 9628f1a |
| N+11 | PBI-09b クロスモジュール hover/goto-def | 73c81f0 |
