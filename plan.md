# verde-lsp バックログ

> 最終更新: 2026-04-21 (Sprint N+24 完了)
> 現在ブランチ: main
> テスト基準: 94 green (lib 36 + integration 58), cargo clippy -D warnings 0 件

---

## Sprint N+22 (2026-04-21)

### Sprint Goal
PBI-20: Private シンボルの cross-file rename 抑止 — Private Sub / local variable が他ファイルに rename 伝播しないよう visibility チェックを追加する

### Path Chosen
Option (A) — `rename.rs` の guard を `(word, cross_file_eligible)` に拡張。`cross_file_eligible = is_public_module_level || found_cross` で分岐。

### Scope
- `tests/rename.rs` に 2 テスト追加 (RED)
  - `rename_private_sub_stays_in_single_file`
  - `rename_local_variable_stays_in_single_file`
- `src/rename.rs` で visibility チェックを追加 (GREEN)
  - `Visibility::Public && proc_scope.is_none()` の場合のみ cross-file
  - それ以外は current file のみ

### Acceptance Criteria
1. `Private Sub Foo()` を rename しても他ファイルの `Private Sub Foo()` は変更されない
2. `Dim x` (local variable) を rename しても他ファイルの `x` は変更されない
3. `Public Sub Foo()` の cross-file rename は引き続き動作する (回帰なし)
4. cargo test 88 → 90 green, clippy -D warnings 0 件

---

## Sprint N+22 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-20 Private シンボルの cross-file rename 抑止」を完全達成。

### KPT

#### Keep
- `(word, cross_file_eligible)` という最小タプル変更で guard ロジックを拡張。既存の公開 API は一切変えず。
- `Visibility::Public && proc_scope.is_none()` という条件が VBA モジュール公開シンボルの定義と正確に一致し、追加の型定義ゼロ。

#### Problem
- rename はまだ text-based (find_all_word_occurrences)。同一ファイル内でも別 procedure の同名ローカル変数が rename される可能性がある。

#### Try
- intra-file scope-aware rename: cursor が proc_scope=Some(X) のシンボル上にある場合、同 procedure 内の occurrences のみ rename する (PBI-21 候補)。

---

## Sprint N+23 (2026-04-21)

### Sprint Goal
PBI-21: intra-file scope-aware rename — 同一ファイル内でも proc_scope を尊重。cursor が proc_scope=Some(X) のシンボル上にある場合、同 procedure 内の occurrences のみ rename する。

### Path Chosen
Option (A) — `rename.rs` の closure に `proc_constraint: Option<TextRange>` を追加。2段階で決定:
1. `find_symbol_at_position` で cursor symbol の proc_scope を確認 (declaration site)
2. use site の場合は `position_to_offset` + `proc_ranges` で containing proc を特定、word がそこのローカルシンボルか確認してから constraint を設定

### Scope
- `tests/rename.rs` に 2 テスト追加 (RED)
  - `rename_local_var_stays_within_its_own_procedure`
  - `rename_from_use_site_stays_within_its_procedure`
- `src/rename.rs` で proc_constraint 計算と occurrences フィルタを追加 (GREEN)

### Acceptance Criteria
1. 同一ファイル内の別 procedure に同名ローカル変数がある場合、cursor の procedure のみ rename される
2. declaration site (Dim x) と use site (x = 1) の両方でスコープ制限が動作する
3. Public module-level シンボルの cross-file rename は影響なし (proc_constraint=None)
4. cargo test 90 → 92 green, clippy -D warnings 0 件

---

## Sprint N+23 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-21 intra-file scope-aware rename」を完全達成。

### KPT

#### Keep
- 2段階の proc_constraint 決定ロジック (declaration site は symbol at cursor、use site は containing proc + local check) が宣言/使用どちらのカーソル位置でも正確に動作。
- `proc_constraint = None` のケース (Public module-level、cross-file) では既存動作が完全に保たれる。フィルタの `match proc_constraint { Some(c) if file_uri == *uri => ... }` パターンが clean。
- `TextRange: Copy` を活かして `proc_ranges` から `*r` でコピー取得、追加アロケーションゼロ。

#### Problem
- cargo test の並列実行時に >60s 警告が散発することがある (既知、高負荷環境の issue)。

#### Try
- `cargo test` の並列度制御 (`-- --test-threads=N`) を Follow-up として記録 (優先度低)。

---

## Follow-ups (優先度低)

- cargo test 並列化チューニング: >60s 警告散発対策として `-- --test-threads=4` など設定を検討。CI では影響なし。

---

## Sprint N+24 (2026-04-21)

### Sprint Goal
PBI-22: rename のパラメータスコープ対応 — procedure params も proc_constraint で絞り込み、別 procedure の同名 parameter を rename しない

### Path Chosen
既存の proc_constraint 機構を確認テストで仕様として固定。`ParameterNode.span` が full parameter span を持ち `find_symbol_at_position` が宣言サイトで param を発見できること、および `proc_scope: Some(proc.name)` が use site Step 2 ロジックを通じて正しく機能することを 2 テストで証明。

### Scope
- `tests/rename.rs` に 2 テスト追加
  - `rename_parameter_stays_within_its_procedure`
  - `rename_parameter_from_use_site_stays_within_its_procedure`
- `src/rename.rs` の変更ゼロ (実装は PBI-21 時点で既に正しかった)

### Acceptance Criteria
1. 別 Sub に同名パラメータがある場合、cursor Sub 内のみ rename される (declaration site)
2. use site (x = 1) からの rename も同一制約で動作する
3. cargo test 92 → 94 green, clippy -D warnings 0 件

---

## Sprint N+24 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-22 パラメータスコープ対応 rename」を完全達成。ただし実装変更はゼロ — 既存の proc_constraint 機構が既にカバーしていた。

### KPT

#### Keep
- テストを書いて即 GREEN になった場合でも「確認テスト」として価値がある。パラメータが proc_constraint を正しく通過するという保証がコードベースに残った。
- `ParameterNode.span` が full parameter span (modifier〜型名まで) を持つ設計が、`find_symbol_at_position` の汎用性を担保している。

#### Problem
- RED フェーズで想定どおりのテスト失敗が起きなかった。PBI-21 の実装範囲の見積が「ローカル変数のみ」と言いながら実際はパラメータも包含していた (見積誤差ではなく設計の良さ)。

#### Try
- 次 PBI で「既に動く可能性が高い XS タスク」は事前に probe を走らせて実装状況を確認してから Sprint に組み込む。

---

## 次 Sprint 推奨 (Sprint N+25)

**Sprint Goal 候補**:
1. symbol kind 対応 (completion/hover での種別表示改善) — S
2. goto-def for parameters (パラメータの定義ジャンプ) — XS

---

## Sprint N+21 (2026-04-21)

### Sprint Goal
PBI-19: `textDocument/documentSymbol` プロバイダ実装 — Module 内 procedure/variable/type を階層で返す

### Path Chosen
Option (A) — 新規 LSP API、既存 `SymbolTable` を再利用

### Scope
- `src/document_symbol.rs` 新規作成
- `SymbolTable.symbols` → `DocumentSymbol` 階層変換
  - `proc_scope=None` → トップレベル
  - `proc_scope=Some(name)` → 対応 Procedure の children
- `server.rs` に `document_symbol` ハンドラ追加
- `InitializeResult` に `document_symbol_provider` 追加

### Probes
- `proc_ranges` で Procedure の full range を取得 → `DocumentSymbol.range`
- `Symbol.span` は name span → `DocumentSymbol.selection_range`
- LSP `SymbolKind`: Procedure→FUNCTION, Variable→VARIABLE, TypeDef→STRUCT, EnumDef→ENUM

### Acceptance Criteria
1. `Sub Foo()` が kind=FUNCTION のトップレベルシンボルとして返される
2. `Foo` のパラメータが children に含まれる
3. Procedure 内の `Dim x` が children に含まれる
4. モジュールレベル `Dim y` がトップレベルシンボルとして返される
5. cargo test 80→84 green, clippy -D warnings 0 件

---

## 次 Sprint 推奨 (Sprint N+22)

**Sprint Goal 候補**: Private 修飾子 cross-file rename 抑止 または symbol kind 対応

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
| PBI-19 | textDocument/documentSymbol プロバイダ | S | **Done** |
| PBI-20 | Private シンボルの cross-file rename 抑止 | XS | **Done** |
| PBI-21 | intra-file scope-aware rename (proc_scope 尊重) | XS | **Done** |
| PBI-22 | rename パラメータスコープ対応 (proc_constraint 確認) | XS | **Done** |

---

## Sprint N+21 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-19 textDocument/documentSymbol プロバイダ実装」を完全達成。

### KPT

#### Keep
- `proc_scope=None/Some(name)` という既存フィールドを階層構造にそのままマップ。新規データ構造ゼロ。
- `proc_ranges` を procedure の full range に再利用し、`selection_range` = name span の LSP 慣例を正確に実装。
- `collect_children` を独立関数に切り出しで、`build_hierarchy` の責務を明確化。

#### Problem
- EnumMember の span が EnumDef の span と同一になっている (build_symbol_table の既知制限)。

#### Try
- EnumMember の個別 name_span を追跡するか、現状の制限をコメントで明示。

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
