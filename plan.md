# verde-lsp バックログ

> 最終更新: 2026-04-21 (Sprint N+43 完了 — PBI-41/42 Phase 2 ログ・workbook 経路確認)
> 現在ブランチ: main
> テスト基準: 123 green (lib 47 + integration 76), cargo clippy -D warnings 0 件

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
| PBI-40 | Verde からの stdio 起動 E2E テスト | S | Backlog (Phase 2) |
| PBI-41 | `workbook-context.json` 書き出し経路検証 | XS | **Done (Sprint N+43)** — tests/workbook.rs 既存カバレッジで完了 |
| PBI-42 | ログ出力方針統一 (`env_logger` → stderr) | XS | **Done (Sprint N+43)** |
| PBI-43 | UDT メンバー解決 (`foo.bar` completion/hover) | L | Backlog (Phase 3) |
| PBI-44 | Class module (`.cls`) サポート | L | Backlog (Phase 3) |
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

## Follow-ups (優先度低)

- cargo test 並列化チューニング: >60s 警告散発対策として `-- --test-threads=4` など設定を検討。CI では影響なし。
- `v0.1.0` タグ push で `release.yml` の動作検証 (3 OS matrix + artifact 添付確認)。
- Named arguments (`Foo x:=1`) の signatureHelp 対応 (VBA での使用頻度は低い、将来課題)。
- `workbook_context` 読み取り helper 抽出 (3 箇所の `.expect()` 文字列重複解消、優先度低)。
