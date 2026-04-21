# verde-lsp バックログ

> 最終更新: 2026-04-21 (Sprint N+47 完了 — PBI-44 Me. 補完 + .cls ヘッダー検証)
> 現在ブランチ: main
> テスト基準: 143 green (lib 55 + integration 88), cargo clippy -D warnings 0 件

---

## 詳細ロードマップ (2026-04-21 策定)

### 前提軸

| 軸 | 決定 | 影響 |
|---|---|---|
| 主要利用シナリオ | **Verde desktop 組み込み最優先** | VS Code 拡張は後回し可。配布は Verde 同梱経路 |
| ユーザー像 | **既存 VBA 開発者** (補完/rename/refactor 重視) | signature help / workspace symbol / format を高優先 |
| 非機能 | Windows 対応は**今すぐ** (Windows 専用開発) | CI/配布を Phase 0 に昇格 |

### Phase 0 — Windows 基盤 & 重大バグ潰し (最優先 / 1-2 Sprint)

| PBI | タイトル | サイズ | 根拠 |
|---|---|---|---|
| PBI-31 | `position_to_offset` を UTF-16 対応 (LSP 準拠) | S | **Done (Sprint N+33)** |
| ~~PBI-32~~ | ~~parser の `.expect()` 除去~~ | ~~XS~~ | **Cancel (Sprint N+34)** — production の panic 経路が存在しないことが判明 |
| PBI-33 | Windows CI 追加 (GitHub Actions `windows-latest`) | S | build/test/clippy matrix、`file:///C:/...` URI 経路確認 |
| PBI-34 | リリースバイナリ自動配布 (tag → `verde-lsp.exe`) | S | Verde desktop 同梱経路の実現 |
| PBI-32b | `std::sync::RwLock.unwrap()` の poison 経路を `expect` コメント化 | XS | `src/server.rs:46,70` / `src/analysis.rs:49,65,69,73` の 2 次 panic を抑止 |
| PBI-32c | `positionEncoding` capability negotiation — initialize で UTF-16 を明示宣言 | XS | LSP 3.17 仕様。将来の UTF-8 client 互換を担保 |

### Phase 1 — 既存開発者の日常機能 (2-3 Sprint)

| PBI | タイトル | サイズ |
|---|---|---|
| PBI-35 | `textDocument/signatureHelp` — 関数呼び出し中のパラメータ表示 | M |
| PBI-36 | `workspace/symbol` — プロジェクト横断シンボル検索 | S |
| PBI-37 | `textDocument/documentHighlight` — 同名シンボルハイライト | XS |
| PBI-38 | `textDocument/foldingRange` — Sub/Function/With ブロック折りたたみ | S |
| PBI-39 | `textDocument/codeAction` — "Dim を追加" quick fix (Option Explicit 連携) | M |

### Phase 2 — Verde desktop 統合検証 (1 Sprint)

| PBI | タイトル | サイズ |
|---|---|---|
| PBI-40 | Verde からの stdio 起動 E2E テスト | S |
| PBI-41 | `workbook-context.json` 書き出し経路の検証 (サーバー側は受信のみ) | XS |
| PBI-42 | ログ出力方針統一 (`env_logger` → stderr、Verde 側でピックアップ可能に) | XS |

### Phase 3 — Symbol 精度強化 (2-3 Sprint / 重い)

| PBI | タイトル | サイズ |
|---|---|---|
| PBI-43 | UDT (`Type` ブロック) メンバー解決 — `foo.bar` の `.bar` completion/hover | L |
| PBI-44 | Class module (`.cls`) サポート — `Me` / インスタンス変数 | L |
| PBI-45 | Excel Object Model 拡充 (PivotTable, Chart, Shape) | M |

### Phase 4 — Polish & リファクタ系キラー機能

| PBI | タイトル | サイズ |
|---|---|---|
| PBI-46 | `textDocument/formatting` — indent / 識別子 case 正規化 | M |
| PBI-47 | Extract Sub/Function リファクタ (code action) | L |
| PBI-48 | `textDocument/inlayHint` — 暗黙 `As Variant` 等の型ヒント | S |
| PBI-49 | call hierarchy (呼び出し元/先ツリー) | M |

### 判断メモ

- Phase 0 の UTF-16 バグ (PBI-31) は**実害が出る前に**潰す。現状 BMP 日本語 (ひらがな/漢字) は偶然動くが、文字列リテラルの emoji 等で破綻
- Class module (PBI-44) の優先度は Verde 利用プロジェクトに `.cls` がどれだけ含まれるかで変動 — 要データ収集
- formatting (PBI-46) は「既存開発者」軸なら Phase 1 に繰り上げる余地あり

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
| PBI-28 | Enum member value で負数・16 進 literal 対応 | XS | **Done** |
| PBI-29 | Enum member implicit value 計算（前メンバ + 1） | XS | **Done** |
| PBI-30b | SymbolDetail::EnumMember.value を Option<i64> → i64（Tidy First 型変更） | XS | **Done** |
| PBI-31 | `position_to_offset` を UTF-16 対応 (LSP 準拠) | S | **Done (Sprint N+33)** |
| ~~PBI-32~~ | ~~parser の `.expect()` 除去~~ | ~~XS~~ | **Cancel (Sprint N+34)** |
| PBI-32b | `RwLock.unwrap()` poison 経路の `expect` 明示化 | XS | **Done (Sprint N+37)** |
| PBI-32c | `positionEncoding` capability negotiation | XS | **Done (Sprint N+34)** |
| PBI-33 | Windows CI 追加 (`windows-latest` matrix) | S | **Done (Sprint N+35)** |
| PBI-34 | リリースバイナリ自動配布 (tag → `verde-lsp.exe`) | S | **Done (Sprint N+36)** |
| PBI-35 | `textDocument/signatureHelp` 実装 | M | **Done (Sprint N+38)** |
| PBI-36 | `workspace/symbol` — プロジェクト横断検索 | S | **Done (Sprint N+40)** |
| PBI-37 | `textDocument/documentHighlight` — 同名ハイライト | XS | **Done (Sprint N+39)** |
| PBI-38 | `textDocument/foldingRange` — ブロック折りたたみ | S | **Done (Sprint N+41)** |
| PBI-39 | `textDocument/codeAction` — "Dim を追加" quick fix | M | **Done (Sprint N+42)** |
| PBI-40 | Verde からの stdio 起動 E2E テスト | S | **Done (Sprint N+44)** |
| PBI-41 | `workbook-context.json` 書き出し経路検証 | XS | **Done (Sprint N+43)** — tests/workbook.rs 既存カバレッジで完了 |
| PBI-42 | ログ出力方針統一 (`env_logger` → stderr) | XS | **Done (Sprint N+43)** |
| PBI-43 | UDT メンバー解決 (`foo.bar` completion/hover) | L | **Done (Sprint N+46)** |
| PBI-44 | Class module (`.cls`) サポート | L | **Done (Sprint N+47)** |
| PBI-45 | Excel Object Model 拡充 (PivotTable/Chart/Shape) | M | Backlog (Phase 3) |
| PBI-46 | `textDocument/formatting` — indent/case 正規化 | M | Backlog (Phase 4) |
| PBI-47 | Extract Sub/Function リファクタ (code action) | L | Backlog (Phase 4) |
| PBI-48 | `textDocument/inlayHint` — 暗黙 `As Variant` 表示 | S | Backlog (Phase 4) |
| PBI-49 | call hierarchy プロバイダ | M | Backlog (Phase 4) |

---

## 完了済み Sprint 要約 (N 〜 N+34)

| Sprint | PBI | 主要成果 |
|--------|-----|----------|
| N | — | If/For/With/Select/Call/Set の undeclared 検出 |
| N+1 | PBI-02+03 | procedure params hover, name_span |
| N+2 | PBI-01 | ローカル変数 SymbolTable 登録 |
| N+3 | PBI-04 | Call/bare/local goto-def |
| N+4 | PBI-05 | While: WhileStatementNode |
| N+5 | PBI-05b | Do/ReDim: StatementNode 化 |
| N+6 | — | Exit/GoTo/OnError: StatementNode 追加 |
| N+7 | — | rename call site 対応 |
| N+8 | — | scope-aware completion (proc_scope) |
| N+9 | BUG-01 | module-level Dim パーサー修正 |
| N+10 | PBI-09a | クロスモジュール補完 |
| N+11 | PBI-09b | クロスモジュール hover/goto-def |
| N+12 | PBI-09c | クロスモジュール diagnostics (undeclared 誤検出除外) |
| N+13 | PBI-12 | ModuleA.Foo 修飾呼び出しの ModuleA undeclared 誤検出除外 |
| N+14 | PBI-11 | workbook-context.json シート名補完 |
| N+15 | PBI-13 | workbook-context.json tables/named_ranges 補完拡張 |
| N+16 | PBI-14 | workbook-context.json 自動再読み込み (didChangeWatchedFiles) |
| N+17 | PBI-15 | textDocument/references プロバイダ実装 |
| N+18 | PBI-16 | textDocument/references クロスファイル拡張 |
| N+19 | PBI-17 | textDocument/rename クロスファイル拡張 |
| N+20 | PBI-18 | rename guard cross-module フォールバック |
| N+21 | PBI-19 | textDocument/documentSymbol プロバイダ実装 |
| N+22 | PBI-20 | Private シンボルの cross-file rename 抑止 |
| N+23 | PBI-21 | intra-file scope-aware rename (proc_scope 尊重) |
| N+24 | PBI-22 | rename パラメータスコープ対応 (proc_constraint 確認テスト) |
| N+25 | PBI-23 | goto-def scope-aware (同名パラメータを含む procedure で正しくジャンプ) |
| N+26 | PBI-24 | hover scope-aware (同名パラメータを含む procedure で正しい型を表示) |
| N+27 | PBI-25 | SymbolDetail::Parameter 追加 — パラメータ登録の設計改善 |
| N+28 | PBI-26 | SymbolDetail::None 完全廃止 — EnumMember バリアント追加 |
| N+29 | PBI-27 | Enum value parsing 対応 — integer literal 右辺を i64 として格納 |
| N+30 | PBI-28 | Enum member 負数・16 進 literal 対応 |
| N+31 | PBI-29 | Enum member implicit value 計算（前メンバ + 1） |
| N+32 | PBI-30b | SymbolDetail::EnumMember.value を Option<i64> → i64（Tidy First 型変更） |
| N+33 | PBI-31 | position_to_offset / offset_to_position を UTF-16 code unit ベースに変更 |
| N+34 | PBI-32c | PBI-32 調査 → Cancel 確定 / positionEncoding UTF-16 明示宣言 / PBI-32b 新設 |
| N+35 | PBI-33 | Windows CI (`windows-latest` matrix) 追加 |
| N+36 | PBI-34 | リリースバイナリ自動配布 (tag → GitHub Release) |
| N+37 | PBI-32b | RwLock poison 経路を `.expect()` コメント化 |
| N+38 | PBI-35 | `textDocument/signatureHelp` 実装 |
| N+39 | PBI-37 | `textDocument/documentHighlight` 実装 |
| N+40 | PBI-36 | `workspace/symbol` 実装 |
| N+41 | PBI-38 | `textDocument/foldingRange` 実装 |
| N+42 | PBI-39 | `textDocument/codeAction` "Dim を追加" quick fix |
| N+43 | PBI-41/42 | workbook-context 書き出し検証 / env_logger → stderr |
| N+44 | PBI-40 | Verde stdio 起動 E2E テスト / Phase 2 完結 |
| N+45 | PBI-43 (partial) | Type ブロック parser 実装 + UdtMember シンボル登録 (Parser/SymbolTable 層) |
| N+46 | PBI-43 (完結) | dot-access 補完 / hover / goto-def 実装 — PBI-43 全完了 |
| N+47 | PBI-44 | Me. dot-access 補完 + .cls ヘッダー検証 — PBI-44 完了 |

---

## Sprint N+35 レトロスペクティブ (PBI-33 Windows CI)

PBI-33 完全達成。`url` クレートの `path_segments()` が Windows ドライブレター (`C:`) を正しく扱うため既存コードは無変更。smoke test をクロスプラットフォーム実装にしたことで macOS CI でも常に検証可能。111 green / clippy 0 / fmt pass。

**Keep**: `path_segments()` 挙動確認後にテストを書いた順序 / `fail-fast: false` の意識的な選択 / smoke test を `#[cfg(windows)]` skip にしなかった判断
**Problem**: Windows runner 実際動作は push までは不明 / `cargo fmt --check` の CRLF 問題の可能性
**Try**: PBI-34 (リリースバイナリ自動配布) を次 Sprint に / CRLF 問題発生時は `.editorconfig` 対応

---

## Sprint N+36 レトロスペクティブ (PBI-34 リリースバイナリ自動配布)

PBI-34 完全達成。`.github/workflows/release.yml` を新規作成。`v*` タグで 3 OS (windows/ubuntu/macos) matrix ビルドが走り、`verde-lsp-windows.exe` / `verde-lsp-linux` / `verde-lsp-macos` として GitHub Release に自動アップロード。111 green 維持。

**Keep**: 3 OS マルチプラットフォーム対応 / `permissions: contents: write` の明示
**Problem**: `softprops/action-gh-release@v2` と Windows `shell: bash` + `cp` の実際動作は未検証 / release.yml に cargo test/clippy を含めていない
**Try**: `v0.1.0` タグ push で動作確認 / release.yml に `cargo test --release` 追加を検討

---

## Sprint N+37 レトロスペクティブ (PBI-32b RwLock poison defensive hardening)

PBI-32b 完全達成。Phase 0 の全 PBI が Done。`server.rs` / `analysis.rs` の RwLock 上の `.unwrap()` 6 件を `.expect("... poisoned: ...")` に変換。Option B (`.expect()` コメント化) を選択 — `tokio::sync::RwLock` 置換は sync メソッドの async 化が必要で XS を超えるため。111 green / clippy 0 / fmt pass。

**Keep**: XS 見積通り Option B で完了 / Phase 0 クリーンアップの順序 (CI → Release → Defensive) が自然
**Problem**: `workbook_context` 読み取り `.expect()` が 3 箇所で同じ文字列 / std::sync::RwLock が async コンテキストで使われ続ける
**Try**: Phase 1 PBI-35 (signatureHelp) に着手

---

## Sprint N+38 レトロスペクティブ (PBI-35 signatureHelp)

PBI-35 完全達成。Phase 1 の最初の機能 PBI を 1 Sprint で消化。後方テキストスキャン方式 — AST に依存せずカーソルオフセットから後方スキャンして `(` と関数名を探す。`clamped` で active_parameter を上限クリップし ParamArray ケースに対応。114 green (lib 47 / integration 67) / clippy 0 / fmt pass。

**Keep**: 後方スキャン方式による入力中のコードへの耐性 / `hover.rs` の `format_params` パターン再利用 / `clamped` による ParamArray 対応
**Problem**: Named arguments (`Foo x:=1, y:=2`) は未対応 — VBA での使用頻度は低い
**Try**: PBI-36 (workspace/symbol) または PBI-37 (documentHighlight XS) を次 Sprint に

---

## Sprint N+39 レトロスペクティブ (PBI-37 documentHighlight)

PBI-37 完全達成。XS 見積通り実装 29 行。`references.rs` のシングルファイル版として `document_highlight.rs` を新規作成。`find_all_word_occurrences` + `text_range_to_lsp_range` を再利用。116 green (lib 47 / integration 69) / clippy 0 / fmt pass。

**Keep**: XS PBI を 1 Sprint で消化するテンポ / `find_word_at_position` の動作を理解したテスト修正 (空白位置が隣接ワードを返す仕様を再確認)
**Problem**: 現在ファイルのみ対応 (cross-file highlight は VBA では通常不要) / `DocumentHighlightKind::TEXT` のみで READ/WRITE 区別なし
**Try**: PBI-36 (workspace/symbol, S) または PBI-38 (foldingRange, S) を次 Sprint に

---

## PBI-35 リファインメント (参考 — 実装済み)

### 技術的前提

| 項目 | 現状 |
|---|---|
| `SymbolDetail::Procedure { params: Vec<ParameterInfo> }` | 既存。`ParameterInfo { name, type_name, passing, is_optional }` がフル情報を保持 |
| cross-module 解決 | `find_public_symbol_in_other_files` / `all_public_symbols_from_other_files` で対応済み |
| cursor 位置 → procedure 特定 | `proc_ranges` + `position_to_offset` の確立済みパターン |
| 新規ファイル | `src/signature_help.rs` を新規作成 |

### MVP スコープ

- trigger characters: `(` と `,`
- cross-module 関数: 対象
- Named arguments (`Foo x:=1`): **対象外**
- VBA ビルトイン関数のシグネチャ: **対象外**

### 設計上の判断

- **後方スキャン方式**: AST より直接テキストスキャンの方が、入力中の不完全なコードに耐性がある
- **括弧ネスト対応**: `Foo(Bar(|))` では最内側 `Bar` のシグネチャを返す (depth tracking)

---

## Sprint N+44 レトロスペクティブ (PBI-40 stdio E2E テスト / Phase 2 完結)

PBI-40 完全達成。Phase 2 全 PBI (PBI-40/41/42) が Done となり、Phase 3 (UDT / Class module) への前提が整った。`tests/e2e_stdio.rs` に `stdio_lifecycle_completes_gracefully` を新規追加。`CARGO_BIN_EXE_verde-lsp` で実バイナリを起動し、initialize → initialized → didOpen → completion → shutdown → exit の完全 LSP ハンドシェイクを wire protocol レベルで verify。実行時間 0.35s / 124 green (lib 47 + integration 77) / clippy 0 / fmt pass。

**Keep**: `recv()` helper が notification をスキップしてリクエスト応答だけを返す設計 (publishDiagnostics を無視できる) / `drop(stdin)` で EOF → serve() 終了のパターンが明快 / tokio::timeout の二段構え (lifecycle 5s + wait 5s)
**Problem**: `exit` notification 後のプロセス終了は stdin EOF に依存 — tower-lsp が `std::process::exit()` を呼ばない実装のため `drop(stdin)` が必須 (落とし穴)
**Try**: Phase 3 PBI-43 (UDT, L) / PBI-44 (Class module, L) — Sprint N+45 で design-weight を評価してから実装 Sprint を切るか直接着手するか決定

### 次 Sprint 候補 (Sprint N+45)

| 候補 | タイプ | 推奨 |
|---|---|---|
| PBI-43: UDT メンバー解決 (`foo.bar` completion/hover) | Planning Sprint (design-weight 評価) | ★ 推奨 |
| PBI-44: Class module (`.cls`) サポート | Planning Sprint (Verde `.cls` 普及度確認後) | 条件付き |
| Follow-up: E2E test を Windows CI matrix に追加 | XS | 追加可能 |

PBI-43/44 は **L サイズ**かつ AST 変更を伴うため、Sprint N+45 は Planning Sprint として設計方針を先に確定させることを推奨。

---

## PBI-43 Backlog Refinement (Sprint N+45 Planning 材料)

### 概要

UDT (`Type` ブロック) のメンバーを `foo.bar` の形式で解決する。補完・hover・goto-def の 3 機能が対象。

### L 見積の根拠

| レイヤー | 変更内容 | コスト |
|---|---|---|
| Parser / AST | `Type ... End Type` ブロックを `TypeBlockNode` として認識、メンバーを `la-arena` に格納 | M |
| SymbolTable | `SymbolDetail::UdtMember { type_name: SmolStr }` を追加、変数の型を `SmolStr` で保持 | S |
| Dot-access 解析 | `foo.bar` の `foo` 型を lookup → UDT 定義を引いて `.bar` の候補を生成 | M |
| テスト | parser / symbol / completion / hover の各層で TDD red-green | S |

合計: L (4-6 Sprint 日相当)。既存 `Enum` ブロック実装 (`src/parser.rs` + `SymbolDetail::EnumMember`) が参考パターン。

### TDD 可能性

**高い。** 各レイヤーを独立してテスト可能:
1. `parser` 単体: `Type Foo\n  x As Long\nEnd Type` → `TypeBlockNode { members: [("x", "Long")] }`
2. `symbol_table` 単体: UDT 定義がシンボルテーブルに登録されること
3. `completion` 統合: `Dim f As Foo\nf.` でメンバー `x` が補完候補に現れること

### `Type` ブロック parser 変更の影響範囲

- `src/parser.rs`: `parse_type_block()` 関数追加。`parse_module_level_statement()` のマッチアームに `Token::Type` を追加。
- `src/ast.rs`: `ModuleLevelStatement::TypeBlock(TypeBlockNode)` バリアント追加。
- `src/analysis.rs`: `collect_symbols()` に `TypeBlock` アームを追加し UDT 定義を登録。
- `src/completion.rs`: dot-access 検出ロジック追加 (`foo.` のプレフィックスで `foo` 型を解決)。
- **既存テストへの影響**: `Type` キーワードをパースできない現状では、`Type` ブロックを含むファイルが `parse_error` 診断を出す可能性。実際の .bas ファイルに `Type` ブロックが含まれているか確認が必要。

### Planning Sprint 必要性の判断

**直接実装を推奨 (Planning Sprint 不要)**。理由:
- `Enum` ブロック (PBI-26/27/28/29) の実装パターンが確立しており、`Type` ブロックは同様のアプローチで実装可能。
- dot-access 解析は新規だが、スコープが明確 (UDT のみ、Excel Object Model の dot 解析とは別経路)。
- TDD で段階的に進められるため設計の不確実性が低い。

**ただし Sprint 分割を推奨**:
- Sprint N+45: Parser + SymbolTable 層 (AST 変更 + UDT 定義登録)
- Sprint N+46: Dot-access 解析 + completion/hover/goto-def

---

## Sprint N+45 レトロスペクティブ (PBI-43 partial — Parser/SymbolTable 層)

PBI-43 の Parser + SymbolTable 層を完全達成。`parse_type_def()` は既存骨格が EndType まで skip するだけだったため、Enum パターンの `skip_newlines()` + identifier ループを踏襲して最小変更で完成。`TypeDefNode.members` を `Vec<NodeId>` → `Vec<(SmolStr, Option<SmolStr>)>` に Tidy First 変更したことで `build_symbol_table` TypeDef arm が Enum arm と完全対称になった。`SymbolKind::UdtMember` / `SymbolDetail::UdtMember{type_name}` を追加し、exhaustive match を completion/document_symbol/workspace_symbol/hover の 4 ファイルで更新。124 → 130 green (新規 6: parser 4 + symbol_table 2) / clippy 0 / fmt pass。

**Keep**: Tidy First → RED → GREEN の順守 / `cargo check` で exhaustive match の欠損を一括検出できた / 既存 TypeDefNode 骨格の再利用でAST変更ゼロ
**Problem**: `TypeDefNode.members: Vec<NodeId>` が常に空のまま Sprint N+44 まで放置されていた (骨格だけあって実装なし)
**Try**: Sprint N+46 で dot-access 解析 + completion/hover/goto-def を実装して PBI-43 を完結させる

---

## Sprint N+46 レトロスペクティブ (PBI-43 完結 — dot-access + completion/hover/goto-def)

PBI-43 を完全達成。Tidy First 2 件 (TypeDefNode.members に per-member span 追加 / SymbolDetail::UdtMember に parent_type 追加) で構造変更を先行し、既存 130 tests を維持したまま dot-access 補完・hover・goto-def をゼロ追加ロジックで GREEN にできた。hover と goto-def は Tidy First の恩恵で実装変更不要。completion のみ `parse_dot_access_at` + `complete_dot_access` を追加。130 → 136 green (新規 6: completion 4 + hover 1 + goto-def 1) / clippy 0 / fmt pass。

**Keep**: Tidy First の per-member span がそのまま goto-def の jump 位置精度を解決した — 構造変更が振る舞い変更を代替した好例
**Problem**: `cargo fmt` 後に `#[test]` 属性が重複/消失するバグが発生 — テスト挿入位置に注意が必要
**Try**: Phase 3 継続: PBI-44 (Class module) または PBI-45 (Excel Object Model 拡充)

### Sprint N+46 Sprint Goal 草案 (PBI-43 完結 — dot-access + completion/hover/goto-def)

`Dim f As MyType` のような UDT 型変数を宣言した後、`f.` と入力した時点で `MyType` のメンバー候補 (x, name 等) が補完に現れ、hover で型情報 `x As Long` が表示され、goto-def で `Type MyType` ブロックのメンバー行にジャンプできること。

**推奨手順**:
1. `foo.` の dot-access prefix を検出する補完ロジックを `src/completion.rs` に追加
2. 変数名 `foo` → 型名 `MyType` の解決: `SymbolTable.symbols` から `foo` を検索し `type_name` を取得
3. 型名 `MyType` → `SymbolKind::UdtMember` かつ `proc_scope == None` なシンボルをフィルタ
4. hover: `f.x` の dot-access を解析して `UdtMember` シンボルを返す
5. goto-def: `UdtMember.span` を `TypeDefNode.span` と紐付ける (現状は TypeDef 全体 span のみ)

---

## Sprint N+47 レトロスペクティブ (PBI-44 — Me. 補完 / .cls ヘッダー検証)

PBI-44 を完全達成。`complete_dot_access` に `Me` 特殊ケースを追加し、`Me.` でカレントモジュールの module-level シンボル (Procedure/Function/Property/Variable/Constant) を補完候補として返す。`.cls` ヘッダー (`VERSION 1.0 CLASS` / `BEGIN`/`END`) は既存の unknown-token スキップ機構で既に処理されていることをパーステストで文書化。hover/goto-def は `find_word_at_position` が `Me.` プレフィックスを自動除外するため実装変更ゼロで動作確認。136 → 143 green (新規 7) / clippy 0 / fmt pass。

**Keep**: `parse_dot_access_at` の `(var_name, partial)` インタフェースが `Me` 特殊ケースの追加点として機能した / テスト → 実装 → 全 suite の TDD 順守
**Problem**: `.cls` ヘッダーの `END` (単体) が `EndSub` 等の複合トークンと競合する可能性を事前確認すべきだった (実際は問題なし)
**Try**: PBI-45 (Excel Object Model 拡充) または PBI-46 (formatting) を次 Sprint に

---

## Sprint N+48 レトロスペクティブ (PBI-45 — Excel Object Model 拡充)

PBI-45 を完全達成。`src/excel_model/types.rs` に PivotTable / Chart / Shape の `ExcelObjectType` 定義を追加し、`complete_dot_access` に Excel builtin type フォールバック (UdtMember lookup 失敗時に `load_builtin_types()` を参照) を実装。Tidy First 順序 (型定義追加 → completion fallback 実装) を厳守。143 → 148 green (新規 5: PivotTable/Chart/Shape テスト 3 + regression 1) / clippy 0 / fmt pass。

**Keep**: `load_builtin_types()` フォールバックが UDT 解決パスに触れずに Excel 型を吸収した設計 / Tidy First で `types.rs` 先行変更後に `completion.rs` を変更した順序 / regression テスト `existing_range_dot_completion_still_works` が既存補完の劣化を即検出
**Problem**: PBI-43 (UDT) で確立した dot-access パスと Excel builtin フォールバックが同一 `complete_dot_access` 内に混在し、将来分岐が増えると複雑化するリスクがある
**Try**: PBI-46 (textDocument/formatting) はスコープが広いため Planning Sprint (N+49) で見積もり・分割方針を先に確定する

---

## PBI-46 Backlog Refinement (Sprint N+49 Planning材料)

### 概要

`textDocument/formatting` (全文書整形) を実装し、VBA ファイルの indent と識別子 case を自動正規化する。

### 整形ルール候補と MVP 選定

| ルール | 難度 | AST 必要 | MVP |
|---|---|---|---|
| キーワード case 正規化 (`dim` → `Dim`) | XS | 不要 (token-based) | ★ |
| Indent 正規化 (4 space per nesting level) | M | 不要 (depth tracking) | ★ |
| 演算子スペース (`x=y` → `x = y`) | S-M | 必要 (`=` 曖昧性) | × (defer) |
| 行末空白除去 | XS | 不要 | ★ (α に同梱) |
| 手続き間 blank line 1 行統一 | XS | 不要 | 条件付き |

**MVP = キーワード case 正規化 + 行末空白除去 + indent 正規化** (演算子スペースは差し戻し)。

### 実装方針

**token-based アプローチ** (AST 不使用):
- 既存 `lex()` が `SpannedToken { token, span, text }` を返す
- `Token` enum variant の canonical form (例: `Token::Sub` → `"Sub"`) と `text` を比較し差分を `TextEdit` に変換
- キーワード case は `SpannedToken.text` vs canonical form の文字列比較のみ
- `End Sub` / `End Function` 等の複合トークンも単一 `Token` バリアントのため canonical 表記が確定

**Indent 計算** (depth tracking):
- 行頭から字句トークンを走査し、`Sub/Function/Property/If/For/With/Select/Do/While/Type` でネストを +1、対応する `End *` / `Next` / `Loop` / `Wend` で -1
- `ElseIf` / `Else` / `Case` はネスト depth を一時的に -1 してインデントを揃える (VBA 慣習)
- 各行の現在インデント量を `leading_spaces()` で計測し、`depth * 4` と比較して TextEdit を生成

**server.rs 変更**:
- `server_capabilities()` に `document_formatting_provider: Some(OneOf::Left(true))` を追加
- `formatting()` handler を実装 (`src/formatting.rs` 新規作成)

### TDD 可能性

**高い。** golden-string 比較で各ルールを独立テスト可能:

```rust
// tests/formatting.rs
fn format_keyword_case_normalizes_dim() {
    let input = "dim x as integer\nsub Foo()\nend sub";
    let result = apply_formatting(input);
    assert_eq!(result, "Dim x As Integer\nSub Foo()\nEnd Sub");
}

fn format_indent_normalizes_nested_if() {
    let input = "Sub Foo()\nIf True Then\nx = 1\nEnd If\nEnd Sub";
    let result = apply_formatting(input);
    assert_eq!(result, "Sub Foo()\n    If True Then\n        x = 1\n    End If\nEnd Sub");
}
```

`apply_formatting(src: &str) -> String` を pure function で実装するため、LSP 層と分離でき単体テストが容易。

### 見積と Sprint 分割方針

| Sprint | 内容 | サイズ |
|---|---|---|
| N+50 (α) | `src/formatting.rs` 新規 + keyword case + 行末空白 + `server.rs` capability 追加 | S |
| N+51 (β) | indent 正規化 (depth tracking) | M |

合計: M (2 Sprint)。PBI-43 (L→2 Sprint) の前例と同等。α は S のため N+50 で完結可能性高い。

### 代替案 / リスク

| 項目 | 評価 |
|---|---|
| 外部 VBA formatter crate | 存在しない — 実装必須 |
| `rustfmt` 流用 | Rust 専用、不可 |
| `prettier` VBA plugin | JS エコシステム依存、LSP サーバーへの組み込み不適 |
| `=` 演算子スペース | `x = y` / `If x = y` / `As Long = 5` の文脈区別に AST 必要 — **defer** |
| `ElseIf/Else/Case` indent | VBA 慣習 (Sub 本体と同深度) と `If` ボディの +1 の境界が複雑 — β で慎重に扱う |

### 二値判断 (Sprint N+50 採択基準)

- **採択**: α (S) で keyword case + 行末空白 → 既存テスト 148 green 維持を確認
- **差し戻し**: `lex()` のスパンが UTF-16 TextEdit 座標と整合しない問題が発覚した場合 (要: β 前に確認)

---

## Follow-ups (優先度低)

- ~~cargo test 並列化チューニング: >60s 警告散発対策として `-- --test-threads=4` など設定を検討。~~ **Done** (.cargo/config.toml に test-threads = 4 を設定)
- `v0.1.0` タグ push で `release.yml` の動作検証 (3 OS matrix + artifact 添付確認)。
- Named arguments (`Foo x:=1`) の signatureHelp 対応 (VBA での使用頻度は低い、将来課題)。
- ~~`workbook_context` 読み取り helper 抽出 (3 箇所の `.expect()` 文字列重複解消、優先度低)。~~ **Done** (f31ab09)
- ~~E2E stdio テスト (`stdio_lifecycle_completes_gracefully`) を Windows CI matrix に追加 (release.yml または ci.yml)。~~ **Done** (ci.yml に windows-latest + cargo test --all が既に設定済み)
