# verde-lsp LSP 機能ギャップ分析

> 作成日: 2026-04-22
> 対象: verde-lsp v0.1.0 (Sprint N+47 完了時点)
> テスト基準: 148 green / clippy 0

---

## 1. 現状サマリ

### 実装済み LSP 機能

| 機能 | プロトコル | 状態 |
|---|---|---|
| テキスト同期 | `textDocument/didOpen`, `didChange`, `didClose` | Full sync |
| 補完 | `textDocument/completion` | `.` トリガー |
| ホバー | `textDocument/hover` | Markdown 形式 |
| シグネチャヘルプ | `textDocument/signatureHelp` | `(`, `,` トリガー |
| 定義ジャンプ | `textDocument/definition` | クロスモジュール対応 |
| 参照検索 | `textDocument/references` | クロスファイル対応 |
| リネーム | `textDocument/rename` | クロスファイル対応 |
| コードアクション | `textDocument/codeAction` | QuickFix のみ |
| 折りたたみ | `textDocument/foldingRange` | Proc/Type/Enum |
| ワークスペースシンボル | `workspace/symbol` | 横断検索 |
| ドキュメントハイライト | `textDocument/documentHighlight` | 同名ハイライト |
| ドキュメントシンボル | `textDocument/documentSymbol` | 階層表示 |
| フォーマット | `textDocument/formatting` | keyword case + indent |
| インレイヒント | `textDocument/inlayHint` | 暗黙型表示 |
| コールハイエラルキー | `textDocument/callHierarchy` | incoming/outgoing |

---

## 2. ドット補完の未対応機能

現在 `complete_dot_access` (`completion.rs:118-232`) が対応しているのは:
- `Me.` (クラスモジュールの自メンバー)
- UDT メンバー (`Dim f As MyType` → `f.x`)
- Excel ビルトイン型 (`Range`, `Worksheet`, `Workbook`, `Application`, `PivotTable`, `Chart`, `Shape`)

### P1 (高優先) --- 日常的な VBA 開発に直結

| ID | 機能 | 説明 | 実装コスト | 根拠 |
|----|------|------|-----------|------|
| DOT-01 | Enum ドットアクセス | `Color.Red` — `SymbolDetail::EnumMember { parent_enum }` のデータは既存。`complete_dot_access` に Enum 分岐を追加するだけ | XS | VBA で Enum は頻出。データ既存のため最小工数 |
| DOT-02 | モジュール名ドットアクセス | `Module1.Foo` — 他ファイルのモジュール名をプレフィックスとして認識し、そのモジュールの Public シンボルを補完 | S | `all_public_symbols_from_other_files` と `collect_other_module_names` が既存 |
| DOT-03 | With ブロック内ドット | `.Value` (先頭ドット) — `With rng ... End With` のコンテキストを追跡し、先頭ドット時に `With` 対象のメンバーを補完 | M | VBA 開発者が最も多用するパターンの一つ。パーサー側で `With` 対象の型追跡が必要 |

### P2 (中優先) --- 開発体験の向上

| ID | 機能 | 説明 | 実装コスト | 根拠 |
|----|------|------|-----------|------|
| DOT-04 | チェーンドットアクセス | `rng.Font.Bold` — `parse_dot_access_at` が1段階のみ。型の再帰的解決ロジックが必要 | M | ネストしたオブジェクト操作は一般的だが、With ブロックで回避されることも多い |
| DOT-05 | 関数戻り値ドットアクセス | `GetRange().Value` — 関数呼び出しの戻り値型を解決。`SymbolDetail::Procedure.return_type` のデータは存在する | M | パーサー側で `()` 直前の関数名を認識するロジックが必要 |
| DOT-06 | VBA ランタイム型 | `Collection`, `Dictionary`, `Scripting.FileSystemObject` 等のメンバー補完 | M | 型定義データの追加が主なコスト。CreateObject パターンとの連携も必要 |

### P3 (低優先) --- あると嬉しい

| ID | 機能 | 説明 | 実装コスト | 根拠 |
|----|------|------|-----------|------|
| DOT-07 | 部分入力フィルタリング | `rng.Va` → `Value` のみに絞り込み。`parse_dot_access_at` の `member_partial` は未使用 | XS | 多くの LSP クライアントがクライアント側でフィルタするため優先度低 |

---

## 3. 未実装の LSP プロトコル機能

### P1 (高優先) --- エディタ体験に大きく影響

| ID | LSP メソッド | 説明 | 実装コスト | 根拠 |
|----|-------------|------|-----------|------|
| LSP-01 | `textDocument/semanticTokens` | 構文ハイライトをサーバーが提供。変数/プロパティ/メソッド/定数を正確に色分け。TextMate 文法では不可能な精度 | M | エディタの見た目に直結。既存シンボルテーブルのデータで大部分は実装可能 |
| LSP-02 | `textDocument/typeDefinition` | `Dim x As MyType` の `x` から `MyType` の定義にジャンプ。goto-def は `x` の宣言へ飛ぶが型定義へは飛べない | S | `symbol.type_name` → `find_symbol_by_name` で型を検索するだけ。データ既存 |
| LSP-03 | `textDocument/prepareRename` | リネーム前にバリデーション。キーワードや読み取り専用シンボルのリネームを事前拒否 | XS | 現在はリネーム実行後にエラーが出るパターンがある。UX 改善 |

### P2 (中優先) --- 特定ワークフローで有用

| ID | LSP メソッド | 説明 | 実装コスト | 根拠 |
|----|-------------|------|-----------|------|
| LSP-04 | `textDocument/selectionRange` | カーソルから段階的にスコープを広げる選択 (式 → 文 → ブロック → プロシージャ → モジュール) | S | AST のノード階層を利用可能。VS Code の Expand Selection に対応 |
| LSP-05 | `textDocument/onTypeFormatting` | `End Sub` 入力時の自動インデント調整、`:` 入力後の行分割等 | S | 既存 formatting ロジックの部分適用 |
| LSP-06 | `textDocument/implementation` | `Implements IFoo` から実装クラスへのジャンプ。VBA のインターフェースパターンで有用 | M | `Implements` ステートメントのパース + クロスモジュール検索が必要 |
| LSP-07 | `textDocument/rangeFormatting` | 選択範囲のみフォーマット。現在はファイル全体のみ | XS | 既存 `apply_formatting` を範囲限定で適用 |

### P3 (低優先) --- 将来的な拡張

| ID | LSP メソッド | 説明 | 実装コスト | 根拠 |
|----|-------------|------|-----------|------|
| LSP-08 | `textDocument/codeLens` | 関数上に「N references」等のインライン情報表示 | S | references 機能が既存。表示データの集約のみ |
| LSP-09 | `textDocument/documentLink` | コメント内の URL やファイルパスをクリック可能リンク化 | XS | VBA での利用場面は限定的 |
| LSP-10 | `textDocument/linkedEditingRange` | `Sub Foo` と `End Sub` を連動編集 | S | VBA では構造上の恩恵が限定的 |

---

## 4. 既存機能の制限事項

### P1 (高優先)

| ID | 機能 | 現状の制限 | 改善内容 | 実装コスト |
|----|------|-----------|---------|-----------|
| LIM-01 | Signature Help | ユーザー定義関数のみ | VBA ビルトイン関数 (`MsgBox`, `InStr`, `Format` 等) の引数ヘルプ追加 | S |
| LIM-02 | Diagnostics | パースエラー + Option Explicit 未宣言 + 未使用変数 | 引数の数チェック、型ミスマッチ警告、到達不能コード検出 | L |
| LIM-03 | Goto Definition | ドットアクセスの右辺未対応 | `obj.Method` の `Method` にカーソルを置いて定義ジャンプ | S |
| LIM-04 | Hover | ドットアクセスのメンバー未対応 | `obj.Method` の `Method` にホバーで型情報表示。Excel ビルトイン型のドキュメントも表示 | S |

### P2 (中優先)

| ID | 機能 | 現状の制限 | 改善内容 | 実装コスト |
|----|------|-----------|---------|-----------|
| LIM-05 | Code Action | "Dim を追加" QuickFix のみ | Extract Sub/Function、Add `Option Explicit`、未使用変数の削除、`Implements` メンバー自動生成 | M-L |
| LIM-06 | Rename | テキストベースの名前置換 | スコープを正確に考慮した安全なリネーム。同名の別スコープシンボルを巻き込まないことを保証 | M |
| LIM-07 | References | 識別子名の文字列一致検索 | ドットアクセスのメンバー参照も検出。スコープ考慮による精度向上 | M |
| LIM-08 | Folding Range | Proc/Type/Enum のみ | `If`, `For`, `With`, `Select Case`, `Do`, コメントブロックの折りたたみ追加 | S |
| LIM-09 | Formatting | keyword case + indent | 演算子周りのスペース正規化、空行の統一、`With` 内インデント対応 | M |
| LIM-10 | Completion (一般) | キーワード・変数・ビルトインが一律で全て出る | コンテキストに応じたフィルタリング (ステートメント開始位置ではキーワード優先、式中では変数優先等) | M |

### P3 (低優先)

| ID | 機能 | 現状の制限 | 改善内容 | 実装コスト |
|----|------|-----------|---------|-----------|
| LIM-11 | Document Highlight | `DocumentHighlightKind::TEXT` のみ | READ/WRITE の区別 (代入先 vs 参照) | S |
| LIM-12 | Completion | スニペットなし | `For...Next`, `If...End If`, `Select Case...End Select` 等の構文スニペット | S |

---

## 5. VBA 固有の未実装機能

### P1 (高優先)

| ID | 機能 | 説明 | 実装コスト | 根拠 |
|----|------|------|-----------|------|
| VBA-01 | Application 暗黙グローバル補完 | `Range()`, `Cells()`, `ActiveSheet`, `Selection` 等を通常補完に追加。`application_globals()` は `types.rs` に既存 | XS | VBA で最も頻繁に使う識別子群。データは既に存在 |
| VBA-02 | 条件コンパイル | `#If VBA7 Then` / `#Else` / `#End If` のパース・折りたたみ・`VBA7`, `Win64` 等の定数補完 | M | 64bit 対応 VBA コードで頻出 |

### P2 (中優先)

| ID | 機能 | 説明 | 実装コスト | 根拠 |
|----|------|------|-----------|------|
| VBA-03 | `Declare` ステートメント | Win32 API 宣言のパース、ホバー、シグネチャヘルプ | M | レガシーコードで頻出 |
| VBA-04 | `WithEvents` イベント補完 | `Private WithEvents btn As CommandButton` → `btn_Click()` 等のイベントプロシージャ自動補完 | M | フォーム/クラスモジュールで必須 |
| VBA-05 | `Implements` メンバー自動生成 | `Implements IMyInterface` 宣言後、未実装メソッドのスタブ生成 (Code Action) | M | インターフェースパターン使用時に必須 |

### P3 (低優先)

| ID | 機能 | 説明 | 実装コスト | 根拠 |
|----|------|------|-----------|------|
| VBA-06 | `Attribute` ステートメント解析 | `VB_Description`, `VB_PredeclaredId` 等のメタデータをホバー/シンボルに反映 | S | エクスポート済み .bas/.cls で有用 |
| VBA-07 | 配列型の解析 | `Dim arr() As Long`, `ReDim arr(1 To 10)` の型情報表示、`UBound`/`LBound` の連携 | S | 配列操作時の型安全性 |

---

## 6. 推奨実装ロードマップ

既存 plan.md の Phase 体系を継続し、Phase 5 以降として整理する。

### Phase 3 残り (現在のバックログ)

| 優先 | ID | PBI 候補 | サイズ |
|------|-------|---------|--------|
| ★★★ | DOT-01 | PBI-51: Enum ドットアクセス補完 | XS |
| ★★★ | DOT-02 | PBI-52: モジュール名ドットアクセス補完 | S |
| ★★★ | VBA-01 | PBI-53: Application 暗黙グローバル補完 | XS |
| ★★☆ | LIM-01 | PBI-54: ビルトイン関数 Signature Help | S |
| ★★☆ | LIM-03 | PBI-55: ドットアクセスの goto-def 対応 | S |
| ★★☆ | LIM-04 | PBI-56: ドットアクセスの hover 対応 | S |

### Phase 5 --- 高精度補完 & セマンティック機能

| 優先 | ID | PBI 候補 | サイズ |
|------|-------|---------|--------|
| ★★★ | DOT-03 | PBI-57: With ブロック内ドット補完 | M |
| ★★★ | LSP-01 | PBI-58: Semantic Tokens プロバイダ | M |
| ★★☆ | LSP-02 | PBI-59: Type Definition プロバイダ | S |
| ★★☆ | DOT-04 | PBI-60: チェーンドットアクセス | M |
| ★★☆ | LIM-08 | PBI-61: 制御構造の Folding Range 追加 | S |
| ★☆☆ | DOT-05 | PBI-62: 関数戻り値ドットアクセス | M |

### Phase 6 --- Diagnostics 強化 & リファクタリング

| 優先 | ID | PBI 候補 | サイズ |
|------|-------|---------|--------|
| ★★☆ | LIM-02 | PBI-63: Diagnostics 拡張 (引数数チェック等) | L |
| ★★☆ | LIM-05 | PBI-64: Code Action 拡張 (Extract Sub 等) | L |
| ★★☆ | VBA-02 | PBI-65: 条件コンパイル対応 | M |
| ★☆☆ | LIM-06 | PBI-66: スコープ安全 Rename | M |
| ★☆☆ | LSP-03 | PBI-67: Prepare Rename | XS |

### Phase 7 --- VBA 固有 & 仕上げ

| 優先 | ID | PBI 候補 | サイズ |
|------|-------|---------|--------|
| ★★☆ | VBA-03 | PBI-68: Declare ステートメント対応 | M |
| ★★☆ | VBA-04 | PBI-69: WithEvents イベント補完 | M |
| ★★☆ | VBA-05 | PBI-70: Implements メンバー自動生成 | M |
| ★☆☆ | DOT-06 | PBI-71: VBA ランタイム型補完 | M |
| ★☆☆ | LSP-04 | PBI-72: Selection Range | S |
| ★☆☆ | LSP-05 | PBI-73: On Type Formatting | S |
| ★☆☆ | LIM-12 | PBI-74: 構文スニペット補完 | S |

---

## 7. 実装コスト凡例

| サイズ | 目安 |
|--------|------|
| XS | 1 Sprint 未満。既存データの活用で完結 |
| S | 1 Sprint。新規ファイル 1 つ程度 |
| M | 2 Sprint。パーサー拡張や複数モジュール変更を伴う |
| L | 3+ Sprint。設計判断を含む大規模変更 |

## 8. 優先度の判断基準

| 優先度 | 基準 |
|--------|------|
| ★★★ | VBA 開発者の日常ワークフローに直結。未実装だとユーザーが「動かない」と感じる |
| ★★☆ | 開発体験を明確に向上させる。なくても作業はできるが、あると効率が上がる |
| ★☆☆ | あると嬉しい。特定のユースケースで役立つ |
