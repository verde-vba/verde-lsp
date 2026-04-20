# verde-lsp バックログ

> 最終更新: 2026-04-21 (Sprint N+29 完了)
> 現在ブランチ: main
> テスト基準: 101 green (lib 39 + integration 62), cargo clippy -D warnings 0 件

---

## Sprint N+29 (2026-04-21)

### Sprint Goal
PBI-27: Enum value parsing 対応 — `Enum X\n    A = 1\n    B = 2\nEnd Enum` 形式の integer literal 右辺を読み `Option<i64>` として `EnumDefNode.members` / `SymbolDetail::EnumMember.value` に格納する。これにより Sprint N+28 の retrospective Try 項目を解消し、hover の `EnumName.Member = N` レンダリングが実データで機能する。

### Path Chosen
Option (A) — `parse_enum_def` のループで Identifier を消費した直後に `try_parse_enum_member_value()` helper を呼び、`Eq` + `NumberLiteral` の 2 トークンを lookahead で読んで `i64::parse` する。負数 (`Minus + NumberLiteral`)・16 進 (`HexLiteral`)・定数式・他メンバ参照は全て `None` のまま (MVP スコープ維持)。

### Scope
- `src/parser/parse.rs` の `parse_enum_def` ループ拡張 (GREEN)
- `src/parser/parse.rs` に `try_parse_enum_member_value` helper 追加 (GREEN)
- `src/analysis/symbols.rs` のテストモジュールに `enum_member_with_explicit_value_captures_integer_literal` 追加 (RED → GREEN)

### Acceptance Criteria
1. `Enum Color\n    Red = 1\n    Green = 2\nEnd Enum` を parse すると Red.value == Some(1), Green.value == Some(2) になる
2. 既存 `enum_member_symbol_has_enum_member_detail` (value 省略ケース) は回帰なし
3. 負数・16 進・定数式は `None` のまま (既存動作維持、新規テスト不要)
4. cargo test 100 → 101 green, clippy -D warnings 0 件, cargo fmt --check pass

---

## Sprint N+29 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-27 Enum value parsing 対応」を完全達成。Sprint N+28 retrospective で記録した「parser の既知制限」が hover の `EnumMember` レンダリング実装と組み合わさり、実際の enum 宣言で value 表示が有効化された。

### KPT

#### Keep
- PBI-26 で `SymbolDetail::EnumMember.value: Option<i64>` の拡張点を先に用意しておいたため、今回は parser 側 1 ファイルの変更のみで完結。変更差分は 22 行。
- `try_parse_enum_member_value` を独立 helper に切り出しで、負数・16 進などの将来拡張時に編集点が局所化された。
- `matches!(self.peek(), Some(t) if t.token == Token::Eq)` で Eq の lookahead を副作用なく判定し、Eq が無い場合は `pos` を進めない設計 — 既存の identifier-only ループと共存。

#### Problem
- 負数 (`Red = -1`) と 16 進 (`Red = &H10`) は実 VBA コードで頻出するが未対応。
- enum member に値を持たないケース (VBA 仕様では前メンバ + 1 で implicit value) の計算は行わず `None` のまま。

#### Try
- PBI-28 候補: 負数 (`Minus + NumberLiteral`) と 16 進 (`HexLiteral`) の対応。helper 内の match を拡張するだけなので XS 見積。
- PBI-29 候補: implicit value 計算 (`A` → 0, `B` → 1 など、前メンバ + 1)。parser か build_symbol_table のどちらで計算するかは別途検討。

---

## Sprint N+28 (2026-04-21)

### Sprint Goal
PBI-26: `SymbolDetail::None` 完全廃止 — EnumMember 用の専用バリアントを追加し、`None` を enum から削除する。全 `SymbolDetail` variant が対応する `SymbolKind` に紐付く設計に揃える。

### Path Chosen
Option (A) — `SymbolDetail::EnumMember { parent_enum: SmolStr, value: Option<i64> }` を新設。`build_symbol_table` の EnumMember 登録を `None → EnumMember` に変更。`hover.rs` の `None` ブランチを `EnumMember` 専用レンダリング (`"Color.Red"` / `"Color.Red = 0"`) に置換。`SymbolDetail::None` を enum 定義から削除。

### Scope
- `src/analysis/symbols.rs` から `SymbolDetail::None` を削除し、`EnumMember` を追加 (RED→GREEN)
- `src/hover.rs` の `None` ブランチを `EnumMember` ブランチに置換
- テスト 1 件追加: `enum_member_symbol_has_enum_member_detail`

### Acceptance Criteria
1. `SymbolDetail::None` バリアントが削除されている
2. `SymbolDetail::EnumMember` が存在し、EnumMember シンボルに使われる
3. hover で EnumMember が `EnumName.MemberName` 形式で表示される (value がある場合は `= N` 付き)
4. cargo test 99 → 100 green, clippy -D warnings 0 件

---

## Sprint N+28 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-26 SymbolDetail::None 完全廃止」を完全達成。PBI-25 と合わせて `SymbolDetail` 6 variant (None 含む) → 5 variant + 全て semantic に対応する SymbolKind を持つ設計へ改善。

### KPT

#### Keep
- PBI-25 の延長線上で設計方向性が既に定まっていたため、RED→GREEN のサイクルが variant 追加/廃止のみに集中できた。REFACTOR 不要。
- Rust の `match` exhaustiveness により、`None` 削除時に `hover.rs` のみが影響範囲と即座に特定できた。型システムが移行を安全にガイド。
- `SymbolDetail::EnumMember` に `value` フィールドを持たせたことで、将来の parser 改善 (enum value parsing) がそのまま hover 表示に反映される拡張点が用意された。

#### Problem
- RED フェーズで `value: Some(0)` を期待したが、パーサーが Enum の `= N` を読まない既知制限のため `None` でしか返らず、テストを弱める調整が必要になった (probe 不足)。
- Enum value parsing は別課題 — `src/parser/parse.rs` line 800 の `members.push((text, None))` が hardcoded None。

#### Try
- PBI-27 候補: Enum value parsing 対応 (`Enum X; A = 1; B = 2; End Enum` の右辺を読み `Option<i64>` に格納)。hover の `EnumMember` レンダリングが即座に活用される。
- RED テストを書く前に parser 側の挙動を軽く probe する習慣 (特に既存機能の前提を使う場合)。

---

## Sprint N+27 (2026-04-21)

### Sprint Goal
PBI-25: パラメータ symbol 登録の設計改善 — `SymbolDetail::Parameter` バリアントを追加し、パラメータが `SymbolDetail::None` で登録される設計上の非対称を解消する

### Path Chosen
Option (A) — `SymbolDetail::Parameter { type_name, passing, is_optional }` を新設。`build_symbol_table` のパラメータ登録を `None → Parameter` に変更。`hover.rs` に `Parameter` ブランチを追加し `ByVal/ByRef/Optional` を含む表示を実装。`SymbolDetail::None` は `EnumMember` 専用として残存。

### Scope
- `src/analysis/symbols.rs` に `SymbolDetail::Parameter` バリアント追加 + 登録変更 (RED→GREEN)
- `src/hover.rs` に `SymbolDetail::Parameter` ブランチ追加 (GREEN)
- テスト 1 件追加: `parameter_symbol_has_parameter_detail`

### Acceptance Criteria
1. `SymbolDetail::Parameter` バリアントが存在し、パラメータシンボルに使われる
2. hover でパラメータが `ByVal x As Integer` 形式で表示される
3. `SymbolDetail::None` は EnumMember にのみ使用される
4. cargo test 98 → 99 green, clippy -D warnings 0 件

---

## Sprint N+27 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-25 SymbolDetail::Parameter 追加」を完全達成。設計の非対称 (`SymbolKind::Parameter` あり ← `SymbolDetail` に対応バリアントなし) を解消。

### KPT

#### Keep
- `Symbol.type_name` フィールドも引き続き設定することで、`type_name` を直接参照している既存コードへの影響ゼロ。非破壊的な拡張。
- `passing` / `is_optional` を `SymbolDetail::Parameter` に持たせたことで、将来の signature help / completion など詳細表示への拡張点が整備された。

#### Problem
- `SymbolDetail::None` は `EnumMember` にのみ残存。EnumMember にも専用バリアント (`SymbolDetail::EnumMember`) を追加すれば `None` を完全廃止できる。

#### Try
- `SymbolDetail::None` の完全廃止を PBI-26 候補として記録 (優先度低)。

---

## Sprint N+26 (2026-04-21)

### Sprint Goal
PBI-24: hover scope-aware — 同名パラメータを持つ複数 procedure がある場合に、cursor がある procedure のパラメータ型を hover で正しく表示する

### Path Chosen
`hover.rs` の `matches.first()` を definition.rs と同形の scope-aware 選択に変更。`position_to_offset` + `proc_ranges` で containing_proc を特定し、`proc_scope` が一致するシンボルを優先。fallback として先頭シンボル (既存動作) を維持。

### Scope
- `tests/hover.rs` に 2 テスト追加 (RED)
  - `hover_parameter_in_first_proc_shows_its_type`
  - `hover_parameter_in_second_proc_shows_its_type`
- `src/hover.rs` の `hover` 関数を scope-aware に変更 (GREEN)

### Acceptance Criteria
1. Sub A(x As Integer) / Sub B(x As String) 構造で Sub A 内カーソルの hover → "Integer"
2. Sub B 内カーソルの hover → "String"
3. 既存 3 tests (local variable / cross-module / sub signature) は回帰なし
4. cargo test 96 → 98 green, clippy -D warnings 0 件

---

## Sprint N+26 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-24 hover scope-aware」を完全達成。2つの修正を同時適用: scope-aware 選択 + SymbolDetail::None の type_name 利用。

### KPT

#### Keep
- definition.rs と同形の `containing_proc + proc_scope` 優先ロジックを hover.rs に転用。パターンの再利用が機能した。
- `SymbolDetail::None` でも `type_name` フィールドは正しく設定されていた。hover のレンダリング層だけ修正すれば済んだ。

#### Problem
- パラメータが `detail: SymbolDetail::None` で登録される設計は、hover や将来の機能拡張で再び同様の見落としを起こす可能性がある。

#### Try
- パラメータを `SymbolDetail::Variable { is_static: false }` で登録するか、専用の `SymbolDetail::Parameter` を追加することを検討 (PBI-25 候補)。

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

## 次 Sprint 推奨 (Sprint N+26)

**Sprint Goal 候補**:
1. hover scope-aware 対応 (hover.rs の同名シンボル選択を proc_scope 優先に) — XS
2. symbol kind 対応 (completion/hover での種別表示改善) — S

---

## Sprint N+25 レトロスペクティブ (2026-04-21)

### Sprint Goal 達成状況

目標「PBI-23 scope-aware parameter goto-def」を完全達成。

### KPT

#### Keep
- Test 1 (Sub A 内カーソル) は `.first()` の挿入順で偶然 PASS — Test 2 (Sub B 内カーソル) が正しく RED になり、実装の不備を明示した。対称的なテスト対を追加する習慣が有効。
- `proc_ranges` + `proc_scope` の組み合わせは rename / goto-def で共通パターン化。同じ 3-5 行で containing_proc 特定が書ける。

#### Problem
- `definition.rs` の `find_symbol_by_name().first()` は PBI-23 まで scope-aware でなかった。同様の問題が `hover.rs` などにも潜在する可能性がある。

#### Try
- hover.rs の symbol 選択も同名ローカル変数/パラメータで正しく scope-aware か確認する (次 Sprint 候補)。

---

## Sprint N+25 (2026-04-21)

### Sprint Goal
PBI-23: goto-def for parameters — 同名パラメータを持つ複数 procedure がある場合に、cursor がある procedure のパラメータ宣言へ正しくジャンプする (scope-aware goto-def)

### Path Chosen
`definition.rs` の `find_symbol_by_name(...).first()` を scope-aware 選択に変更。`position_to_offset` + `proc_ranges` で containing_proc を特定し、`proc_scope` が一致するシンボルを優先。fallback として先頭シンボル (既存動作) を維持。rename.rs の proc_constraint 決定ロジックと対称な実装。

### Scope
- `tests/definition.rs` に 2 テスト追加 (RED)
  - `goto_def_parameter_from_use_site_jumps_to_owning_proc_param`
  - `goto_def_parameter_in_second_proc_jumps_to_its_own_param`
- `src/definition.rs` の `goto_definition` を scope-aware に変更 (GREEN)

### Acceptance Criteria
1. 同名パラメータを持つ 2 つの Sub で、Sub A 内 use site から goto-def → Sub A のパラメータ宣言 (line 0, col 6)
2. Sub B 内 use site から goto-def → Sub B のパラメータ宣言 (line 3, col 6)
3. 既存 4 tests (call site / bare call / local variable / cross-module) は回帰なし
4. cargo test 94 → 96 green, clippy -D warnings 0 件

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
| PBI-23 | goto-def scope-aware (同名パラメータを含む procedure で正しくジャンプ) | XS | **Done** |
| PBI-24 | hover scope-aware (同名パラメータを含む procedure で正しい型を表示) | XS | **Done** |
| PBI-25 | SymbolDetail::Parameter 追加 — パラメータ登録の設計改善 | XS | **Done** |
| PBI-26 | SymbolDetail::None 完全廃止 — EnumMember バリアント追加 | XS | **Done** |
| PBI-27 | Enum value parsing 対応 — integer literal 右辺を i64 として格納 | XS | **Done** |

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
