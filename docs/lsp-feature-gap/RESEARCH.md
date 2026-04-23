# verde-lsp LSP 機能ギャップ — コード調査レポート

> 作成日: 2026-04-22
> 対象: verde-lsp v0.1.0 (Sprint N+47 / 148 green)
> ベース: [docs/lsp-feature-gap.md](../lsp-feature-gap.md)

---

## 目次

1. [パーサー・AST 基盤の調査](#1-パーサーast-基盤の調査)
2. [ドット補完の詳細調査](#2-ドット補完の詳細調査)
3. [未実装 LSP プロトコルの実現可能性](#3-未実装-lsp-プロトコルの実現可能性)
4. [既存機能の制限事項の詳細](#4-既存機能の制限事項の詳細)
5. [VBA 固有機能の調査](#5-vba-固有機能の調査)
6. [横断的な気づき・設計上の懸念](#6-横断的な気づき設計上の懸念)

---

## 1. パーサー・AST 基盤の調査

### 1.1 パーサーアーキテクチャ

**tree-sitter ではない。** 完全自前実装の 3 層構成:

| 層 | クレート | ファイル | 役割 |
|---|---|---|---|
| Lexer | `logos` | `src/parser/lexer.rs` | 正規表現ベースのトークン生成 |
| Parser | 手書き | `src/parser/parse.rs` | 再帰下降パーサー (バックトラックなし) |
| AST | `la-arena` | `src/parser/ast.rs` | アリーナアロケーション AST |

パーサーは `Parser<'a>` 構造体 (`parse.rs:34-39`) がトークンスライスを左から右へ走査し、
位置カウンタ (`pos`) を進めながら AST ノードを構築する。

### 1.2 AST ノード構造

```
AstNode (ast.rs:34-42)
├── Module(ModuleNode)        — 現在未使用 (root が直接子ノードを保持)
├── Procedure(ProcedureNode)  — Sub/Function/Property Get/Let/Set
├── Variable(VariableNode)    — モジュールレベル宣言
├── Parameter(ParameterNode)  — プロシージャ引数
├── TypeDef(TypeDefNode)      — Type...End Type
├── EnumDef(EnumDefNode)      — Enum...End Enum
└── Statement(StatementNode)  — プロシージャ本文の各ステートメント
```

**StatementNode のバリアント** (全 14 種, `ast.rs:124-139`):

| バリアント | 内部データ | 備考 |
|---|---|---|
| LocalDeclaration | `DeclKind`, `names: Vec<(SmolStr, Option<SmolStr>, TextRange)>` | Dim/Static/Const。名前・型・スパンをタプルで保持 |
| Expression | `tokens, span` | 代入・関数呼び出しのキャッチオール |
| If | `tokens, span` | **ヘッダー行のみ**。本体・Else・End If は別ステートメント |
| For | `tokens, span` | **ヘッダー行のみ**。Next は別ステートメント |
| With | `tokens, span` | **ヘッダー行のみ**。End With は別ステートメント |
| Select | `tokens, span` | **ヘッダー行のみ** |
| Call | `tokens, span` | `Call Foo(...)` |
| Set | `tokens, span` | `Set obj = ...` |
| While | `tokens, span` | ヘッダー行のみ |
| Do | `tokens, span` | ヘッダー行のみ |
| Redim | `tokens, span` | `ReDim [Preserve] arr(n)` |
| Exit | `tokens, span` | `Exit Sub/Function/For/Do` |
| GoTo | `tokens, span` | `GoTo label` |
| OnError | `tokens, span` | `On Error GoTo/Resume` |

> **気づき**: If/For/With/Select/Do/While は**ヘッダー行のみ**が AST ノードになり、
> 対応する End トークンは別のステートメントとして処理される。
> つまり**ブロック構造がAST上で表現されていない**。
> これは With ブロックのコンテキスト追跡 (DOT-03) や
> 制御構造の折りたたみ (LIM-08) の実装に大きく影響する。

### 1.3 スパン情報の網羅性

全 AST ノードに `span: TextRange` (`ast.rs:270-282`) が付与されている。
`TextRange` は `start: u32, end: u32` のバイトオフセット。

| ノード | スパン範囲 | 追加のスパン |
|---|---|---|
| ProcedureNode | Sub〜End Sub 全体 (`span`) | 名前のみ (`name_span`) |
| VariableNode | 変数名のみ | — |
| ParameterNode | パラメータ名のみ | — |
| TypeDefNode | Type 全体 | メンバーごとの `TextRange` (メンバー名のスパン) |
| EnumDefNode | Enum 全体 | — (メンバー個別スパンなし) |
| StatementNode 各種 | ステートメント全体 | 内部の `tokens` に個々の `SpannedToken.span` |

> **気づき**: EnumDef のメンバーには個別スパンがない (`Vec<(SmolStr, Option<i64>)>`)。
> Enum メンバーへの goto-def や rename で正確な位置ジャンプができない可能性がある。
> TypeDef のメンバーは Sprint N+45 で per-member span が追加済みなので、
> Enum にも同様の対応が将来必要。

### 1.4 モジュールレベルのパース対応表

`parse_module()` (`parse.rs:93-122`) のトークン分岐:

| トークン | 処理 | 状態 |
|---|---|---|
| `Option` | `parse_option()` → Option Explicit 検出 | **対応済み** |
| `Attribute` | `skip_attribute_line()` → スキップ | **対応済み** (スキップのみ) |
| `Public/Private/Friend` | `parse_declaration_with_visibility()` | **対応済み** |
| `Sub` | `parse_procedure()` | **対応済み** |
| `Function` | `parse_procedure()` | **対応済み** |
| `Property` | `parse_property()` | **対応済み** |
| `Dim/Const` | `parse_variable()` | **対応済み** |
| `Type` | `parse_type_def()` | **対応済み** |
| `Enum` | `parse_enum_def()` | **対応済み** |
| `Implements` | **`_ => self.pos += 1`** (スキップ) | **未対応** — トークンは定義済みだがパース分岐なし |
| `Event` | **`_ => self.pos += 1`** (スキップ) | **未対応** |
| `WithEvents` | **`_ => self.pos += 1`** (スキップ) | **未対応** |
| `Declare` | **`_ => self.pos += 1`** (スキップ) | **未対応** — トークンすら未定義 (Identifier に吸収) |
| `#If` | トークン化されない | **完全に不在** |

> **気づき**: `Implements` のトークン (`Token::Implements`) は `lexer.rs:117` で定義済みだが、
> パーサーのモジュールレベル分岐 (`parse.rs:103-120`) には `Token::Implements` のアームがない。
> `_ =>` キャッチオールで**無言スキップ**される。エラーも出ない。
> つまりユーザーが `Implements IFoo` と書いても何も起こらない。

> **気づき**: `Declare` はさらに深刻で、レキサーにトークンすら定義されていない。
> `Declare Sub Foo Lib "kernel32"` と書くと、`Declare` は `Identifier` トークンになり、
> 次の `Sub` でプロシージャパースが始まってしまう。壊れたASTが生成される可能性がある。

### 1.5 エラー回復戦略

パーサーのエラー回復は**サイレントスキップ方式**:

- モジュールレベル: 未認識トークンは `self.pos += 1` で 1 トークンずつスキップ (`parse.rs:117-119`)
- プロシージャ内: End トークンか EOF まで消費を続行
- 唯一のエラー: 未終了プロシージャ (`parse.rs:271-274`)

この設計は不完全なコードへの耐性が高い反面、構文エラーの報告が最小限。

### 1.6 行継続 (`_`) のサポート

レキサーが `_[ \t]*\n` を単一の `LineContinuation` トークンとして認識 (`lexer.rs:232-233`)。
パーサーには 3 種の行継続スキップ関数:

| 関数 | 用途 | 動作 |
|---|---|---|
| `skip_line_continuations()` | パラメータリスト内 | `LineContinuation` と `Newline` の両方をスキップ |
| `skip_line_continuations_preserving_newline()` | プロシージャシグネチャ | `LineContinuation` のみスキップ、`Newline` は残す |
| `skip_statement_separators()` | ステートメント間 | `Newline`, `Colon`, `LineContinuation`, `Comment` をスキップ |

> **気づき**: `signature_help.rs` の `find_call_context()` は行継続を認識しない。
> `b'\n' | b'\r' => return None` (`signature_help.rs:88`) で改行即打ち切り。
> VBA でよくある複数行にまたがる関数呼び出しでシグネチャヘルプが効かない。

---

## 2. ドット補完の詳細調査

### 2.1 `parse_dot_access_at` のアルゴリズム詳細

**場所**: `resolve.rs:127-156`

後方バイトスキャン方式で、カーソル位置からドットアクセスのコンテキストを検出する:

```
Step 1: カーソルから後方走査 → member_start を特定 (識別子文字以外で停止)
Step 2: member_start-1 が '.' でなければ None
Step 3: ドット位置から後方走査 → var_start を特定
Step 4: (var_name, member_partial) を返す
```

識別子文字の判定 (`resolve.rs:158-160`):
```rust
fn is_ident_char(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_'
}
```

### 2.2 各ケースのバイトレベルトレース

#### `a.b.c` (チェーンドット — DOT-04)

```
ソース: "a.b.c"  カーソル: offset=5 (末尾)

Step 1: 'c' は ident → member_start=4
Step 2: bytes[3]='.' ✓ → dot_pos=3, member_partial="c"
Step 3: 'b' は ident → var_start=2
        bytes[1]='.' → ident ではない → 停止

結果: ("b", "c") を返す
```

**問題**: `"a"` が完全に無視される。`b` は変数ではなく `a` のメンバーなので、
シンボルテーブルに `b` という変数は存在しない → 型解決失敗 → 空の補完。

#### `.Value` (With ブロック内先頭ドット — DOT-03)

```
ソース: "    .Value"  カーソル: offset=10 (末尾)

Step 1: "Value" を後方走査 → member_start=5
Step 2: bytes[4]='.' ✓ → dot_pos=4, member_partial="Value"
Step 3: bytes[3]=' ' → ident ではない → var_start=4
        var_start == dot_pos → return None

結果: None (ドットの前に識別子がないため)
```

**問題**: `parse_dot_access_at` は「ドットの前に識別子があること」を前提としている (`resolve.rs:150`)。
With ブロックのコンテキスト (With 対象の型) を知る仕組みが完全に不在。

#### `GetRange().Value` (関数戻り値ドット — DOT-05)

```
ソース: "GetRange().Value"  カーソル: offset=16 (末尾)

Step 1: "Value" を後方走査 → member_start=11
Step 2: bytes[10]='.' ✓ → dot_pos=10, member_partial="Value"
Step 3: bytes[9]=')' → is_ident_char(')') = false → var_start=10
        var_start == dot_pos → return None

結果: None (括弧はident文字ではないため)
```

**問題**: 括弧の前の関数名を認識する機能がない。
関数の戻り値型を解決するには式レベルの型推論が必要。

### 2.3 `complete_dot_access` の分岐フロー

**場所**: `completion.rs:118-232`

```
complete_dot_access(symbols, source, position)
    │
    ├── offset = position_to_offset() ── None → return None
    ├── (var_name, _) = parse_dot_access_at() ── None → return None
    │
    ├── [1] var_name == "Me" ?
    │   └── YES → module-level の Proc/Func/Prop/Var/Const を返す
    │            (completion.rs:127-150)
    │
    ├── [2] 変数の型を解決
    │   ├── proc_scope 内の Variable/Parameter を検索 (優先)
    │   └── module-level の Variable/Parameter を検索 (フォールバック)
    │   └── type_name が None → return None
    │
    ├── [3] UDT メンバー検索
    │   └── SymbolKind::UdtMember で parent_type == type_name のものをフィルタ
    │   └── 見つかった → return Some(members)
    │
    ├── [**ここに Enum 分岐が必要** — DOT-01]
    │
    ├── [4] Excel ビルトイン型フォールバック
    │   └── load_builtin_types() から type_name に一致する型を検索
    │   └── 見つかった → properties + methods を返す
    │
    └── return Some(members)  ← members が空の場合、空リストを返す
```

> **気づき**: フロー [2] の型解決は `Variable` と `Parameter` のみを検索する
> (`completion.rs:162-178`)。つまり:
> - `Dim x As Color` の `x.` → `Color` (型名) を取得 → [3] UDT 検索 → 失敗 →
>   [4] Excel ビルトイン → 失敗 → 空リスト
> - Enum は [3] でも [4] でもヒットしない。[3] と [4] の間に Enum 分岐を挿入すべき。

### 2.4 DOT-01: Enum ドットアクセスの実装可能性

**結論: データは完全に揃っている。追加コードは約10行。**

既存データ (`symbols.rs:238-249`):
```rust
// build_symbol_table() 内の Enum 処理
symbols.push(Symbol {
    name: member_name.clone(),
    kind: SymbolKind::EnumMember,
    type_name: Some(ed.name.clone()),  // ← Enum名が type_name に入る
    detail: SymbolDetail::EnumMember {
        parent_enum: ed.name.clone(),   // ← parent_enum で親Enumを特定可能
        value: resolved,
    },
    ...
});
```

**ただし注意点が2つ**:

1. **直接アクセスパターン** (`Color.Red`): `Color` は型名であり変数名ではない。
   現在の `complete_dot_access` は変数/パラメータのみを検索するため (`completion.rs:162-178`)、
   `Color` がヒットしない。**型解決の前に `SymbolKind::EnumDef` で名前一致を試す分岐**が必要。

2. **変数経由パターン** (`Dim c As Color` → `c.`): `c` の `type_name` は `"Color"`。
   [3] UDT 検索は失敗するが、Enum 分岐を [3] と [4] の間に追加すれば動作する。

> **気づき**: パターン1 (直接アクセス) は現在の型解決フローに乗らない。
> `Me.` と同様に、`parse_dot_access_at` の返す `var_name` を
> **変数名としてではなく型名/Enum名としても**検索する分岐が必要。
> これは DOT-02 (モジュール名ドットアクセス) にも共通する課題。

### 2.5 DOT-02: モジュール名ドットアクセスの実装可能性

**現状のモジュール名認識** (`analysis.rs:123-135`):

`collect_other_module_names()` は URI のファイル名部分を抽出して小文字化:
```
file:///path/to/ModuleA.bas → "modulea"
```

**問題点**:

1. `collect_other_module_names()` は **diagnostics 専用** (`analysis.rs:103-107`)。
   補完側からは呼ばれていない。

2. 返すのは `Vec<String>` (名前だけ) で、**URI やシンボルとの紐付けがない**。
   `Module1.` で補完するには「Module1」から当該ファイルの Public シンボルを取得する必要があるが、
   **モジュール名 → URI のマッピングが存在しない**。

3. `Symbol` 構造体にモジュール名/URI フィールドがない (`symbols.rs:14-23`)。
   `all_public_symbols_from_other_files()` はフラットなリストを返すだけ (`analysis.rs:158-173`)。

**必要な変更**:

```
Option A: AnalysisHost に新メソッド追加
  pub fn public_symbols_from_module(&self, module_name: &str) -> Vec<Symbol>
  → files を iterate し、URI のファイル名部分が module_name に一致するエントリの
    Public/module-level シンボルを返す

Option B: complete_dot_access に分岐追加
  var_name が変数にも型にもヒットしない場合、
  collect_other_module_names 相当のロジックで var_name をモジュール名として照合
```

> **気づき**: Option A が望ましい。`complete_dot_access` 内でのみモジュール解決するのではなく、
> `AnalysisHost` レベルで提供すれば、hover/goto-def の `ModuleA.Foo` 対応にも流用できる。
> `collect_other_module_names` のロジック (URI → filename → stem) を共通化すべき。

### 2.6 DOT-03: With ブロック補完の実装課題

**WithStatementNode の構造** (`ast.rs:185-191`):

```rust
pub struct WithStatementNode {
    pub tokens: Vec<SpannedToken>,  // With キーワード + 対象式の生トークン列
    pub span: TextRange,
}
```

`With rng` → `tokens = [Token::With, Token::Identifier("rng")]`
`With ws.Range("A1")` → `tokens = [Token::With, Token::Identifier("ws"), Token::Dot, ...]`

**課題の分解**:

| 課題 | 難易度 | 説明 |
|---|---|---|
| With 対象式の抽出 | 低 | tokens から With の次のトークンを取り出すだけ (単純識別子の場合) |
| With 対象の型解決 | 中 | 抽出した識別子の型を SymbolTable から検索 |
| カーソルの With スコープ判定 | **高** | AST上で With ブロックの範囲が明示されていない |
| ネストした With 対応 | **高** | `With a ... With b ... .x ... End With ... End With` の最内 With を特定 |

> **気づき**: 最大の障壁は「AST 上で With ブロックが構造化されていない」こと。
> `WithStatementNode` はヘッダー行のみで、End With は別の（Expression 相当の）ステートメント。
> カーソルが With ブロック内にあるかを判定するには:
>
> 1. プロシージャ本体のステートメントを走査
> 2. `StatementNode::With` を見つけたら「With 開始」
> 3. 対応する `End With` トークンを見つけたら「With 終了」
> 4. カーソルオフセットがその範囲内かチェック
> 5. ネストに対応するにはスタックが必要
>
> これは `proc_ranges` と同様の「With 範囲」をシンボルテーブル構築時に
> 事前計算して保持するアプローチが有効。

### 2.7 DOT-07: `member_partial` の未使用

`parse_dot_access_at` は `(var_name, member_partial)` を返す (`resolve.rs:155`):

```rust
Some((var_name, member_partial))
```

しかし `complete_dot_access` は `member_partial` を捨てている (`completion.rs:124`):

```rust
let (var_name, _) = parse_dot_access_at(source, offset)?;
//               ^ 未使用
```

**影響**: `rng.Va` と入力しても `Value`, `VarType`, ... が全件返る。
LSP クライアント (VS Code 等) がクライアント側でフィルタリングするため実害は小さいが、
大量の候補がある場合にネットワーク/処理効率に影響する可能性がある。

---

## 3. 未実装 LSP プロトコルの実現可能性

### 3.1 LSP-01: Semantic Tokens

**実現可能性: 高**

既存データで対応可能な SemanticTokenType マッピング:

| SymbolKind | → SemanticTokenType | データ元 |
|---|---|---|
| Procedure | `method` | `symbols.rs:27` |
| Function | `function` | `symbols.rs:28` |
| Property | `property` | `symbols.rs:29` |
| Variable | `variable` | `symbols.rs:30` |
| Constant | `variable` + modifier `readonly` | `symbols.rs:31` |
| Parameter | `parameter` | `symbols.rs:32` |
| TypeDef | `struct` | `symbols.rs:33` |
| EnumDef | `enum` | `symbols.rs:34` |
| EnumMember | `enumMember` | `symbols.rs:35` |
| UdtMember | `property` | `symbols.rs:36` |

レキサーの `SpannedToken` が正確なバイトオフセット (`span: Range<usize>`) を提供するため、
トークン位置の計算は容易。

**実装アプローチ**:

1. シンボルテーブルの各シンボルを走査し、`span` を使って対応する行・列を算出
2. キーワードはレキサーの `Token` enum バリアントから直接マッピング
3. コメント (`Token::Comment`, `lexer.rs:250`) と文字列 (`Token::StringLiteral`, `lexer.rs:239`) もトークンから取得可能

> **気づき**: シンボルテーブルのスパンは**宣言位置のみ**を指す。
> 参照箇所 (使用箇所) のセマンティックトークンを提供するには、
> `find_all_word_occurrences` 相当の処理が必要。
> 宣言のみのハイライトなら XS、参照込みなら M の作業量。

### 3.2 LSP-02: Type Definition

**実現可能性: 高**

既存の `goto_definition` (`definition.rs:6-66`) とほぼ同じ構造で実装可能。

```
goto_definition: カーソルの単語 → find_symbol_by_name → シンボルの span にジャンプ
type_definition: カーソルの単語 → find_symbol_by_name → symbol.type_name を取得
                → type_name で再度 find_symbol_by_name → 型シンボルの span にジャンプ
```

**注意点**:

- `symbol.type_name` は `Option<SmolStr>` (`symbols.rs:17`)。`None` の場合 (暗黙 Variant) はジャンプ先なし。
- ビルトイン型 (`Long`, `String` 等) はソースコード上に定義がないためジャンプ不可。
- `TypeDef` と `EnumDef` は `SymbolTable` に登録されているのでジャンプ可能。
- クロスモジュールの型は `find_public_symbol_in_other_files` で検索可能。

### 3.3 LSP-03: Prepare Rename

**実現可能性: 高 (XS)**

現在の `rename.rs` にはバリデーションがない:

```rust
pub fn rename(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
    new_name: &str,  // ← バリデーションなし
) -> Option<WorkspaceEdit>
```

> **気づき (重要)**: 現在、VBA キーワードへのリネームを**防止する仕組みがない**。
> `x` を `Sub` にリネームすることが可能。これはバグと言えるレベルの欠陥。
> `vba_builtins::KEYWORDS` (72 語) と `BUILTIN_TYPES` (13 語) との照合チェックを
> `prepareRename` で行うのが最小コストの修正。

```rust
// 追加すべきチェック (数行で実装可能):
if vba_builtins::KEYWORDS.iter().any(|kw| kw.eq_ignore_ascii_case(new_name))
    || vba_builtins::BUILTIN_TYPES.iter().any(|t| t.eq_ignore_ascii_case(new_name))
{
    return None; // リネーム拒否
}
```

### 3.4 LSP-04: Selection Range

**実現可能性: 高**

AST の階層構造がそのまま Selection Range に使える:

```
文字 "x" → 識別子 "x" (name span)
         → ステートメント "Dim x As Long" (statement span)
         → プロシージャ本体 (proc body range)
         → プロシージャ全体 (proc span)
         → ファイル全体
```

`proc_ranges` (`symbols.rs:10`) が既にプロシージャのバイト範囲を提供。
ステートメントの `span` フィールドも利用可能。

**制約**: ブロック構造 (If...End If の範囲) が AST で表現されていないため、
ステートメント → プロシージャの間の中間レベル (ブロックレベル) の選択は不可。

### 3.5 LSP-05: On Type Formatting

**実現可能性: 中**

既存の `apply_formatting` (`formatting.rs:3-47`) は 2 フェーズ構成:

1. **Phase 1**: レキサートークンを走査し、キーワードの case を正規化 (`formatting.rs:5-17`)
2. **Phase 2**: 行ごとにインデント深度を計算し、正規化 (`formatting.rs:19-46`)

`calculate_line_indents()` (`formatting.rs:53-88`) はブロック開閉トークンを走査して
各行のインデント深度を独立に計算する。

**On Type Formatting への適用**:
- `End Sub` 入力時: `End Sub` がある行のインデントを depth=0 に設定
- `End If` 入力時: depth を 1 減らす
- `:` 入力時: ステートメント分割 (VBA の `:` は行内ステートメント区切り)

> **気づき**: `calculate_line_indents` は各行のインデントを独立に計算する設計のため、
> 「特定の行のみ再計算」が比較的容易。全文書フォーマットのロジックを
> 行範囲指定で切り出す形で実装できそう。

### 3.6 LSP-06: Implementation (Implements)

**実現可能性: 低 — パーサー拡張が必須**

現状:
- `Token::Implements` はレキサーで定義済み (`lexer.rs:117-118`)
- パーサーの分岐に `Token::Implements` のアームがない → サイレントスキップ
- AST に `ImplementsNode` が存在しない
- シンボルテーブルにインターフェース情報なし

**必要な作業量**:

| 作業 | サイズ |
|---|---|
| `ImplementsNode` を AST に追加 | XS |
| `parse_module()` に Implements 分岐を追加 | XS |
| シンボルテーブルに ImplementsInfo を追加 | S |
| クロスモジュールのインターフェース解決 | M |
| Implementation プロバイダ (goto implementation) | S |
| Code Action: 未実装メンバーのスタブ生成 | M |

合計: **L サイズ**。パーサー変更 → シンボル → 解析 → LSP の全レイヤーに波及。

---

## 4. 既存機能の制限事項の詳細

### 4.1 LIM-01: Signature Help のビルトイン関数未対応

**原因**: `signature_help.rs:18-32` でシンボルテーブルのみを検索。
ビルトイン関数のシグネチャデータが存在しない。

`vba_builtins.rs` は関数**名**のリスト (`BUILTIN_FUNCTIONS`, 88 関数) のみで、
パラメータ情報を持たない:

```rust
pub const BUILTIN_FUNCTIONS: &[&str] = &[
    "Abs", "Array", "Asc", ...  // 名前だけ、シグネチャなし
];
```

**必要なデータ構造**:
```rust
pub struct BuiltinSignature {
    pub name: &'static str,
    pub params: &'static [(&'static str, &'static str, bool)], // (name, type, optional)
    pub return_type: Option<&'static str>,
}
```

**追加のバグ — 行継続未対応**:

`find_call_context()` (`signature_help.rs:52-93`) は改行で即座に `None` を返す:
```rust
b'\n' | b'\r' => return None,  // signature_help.rs:88
```

VBA で一般的な複数行呼び出し:
```vba
MsgBox "Hello", _
    vbInformation, _
    "Title"
```
→ 2行目以降でシグネチャヘルプが消える。

### 4.2 LIM-02: Diagnostics の現状

**場所**: `diagnostics.rs`

現在の診断項目:

| 診断 | 重要度 | 実装 |
|---|---|---|
| パースエラー | ERROR | `parse_result.errors` をそのまま報告 |
| Option Explicit 違反 | WARNING | トークンベースの未宣言検出 |
| 未使用変数 | — | **未実装** |
| 引数の数チェック | — | **未実装** |
| 型ミスマッチ | — | **未実装** |

**Option Explicit のトークン走査** (`diagnostics.rs:113-287`):
- プロシージャ本体のステートメントを走査
- 各ステートメントの `tokens` フィールドからトークンを 1 つずつチェック
- `.` の後のトークンはスキップ (メンバーアクセスの RHS)
- `As` の後のトークンはスキップ (型名位置)
- それ以外の `Identifier` が `declared` セットになければ警告

> **気づき**: この走査は**非常にフラット**。式の構造を理解していない。
> `For Each x In collection` の `x` はループ変数として「宣言済み」と扱うが、
> これはステートメント種別ごとのハードコードで対応している。
> 引数の数チェックを追加するには、呼び出し式を構造的にパースする必要があり、
> 現在のトークンフラット走査では不十分。L サイズの作業。

### 4.3 LIM-05: Code Action の現状

`code_action.rs` には**1種類のみ**:
- `Option Explicit` 違反に対する `Dim varName As Variant` 挿入 QuickFix

**パターン** (`code_action.rs:10-50`):
1. 受け取った `diagnostics` をフィルタ (メッセージに特定文字列を含むか)
2. 診断メッセージから変数名を抽出
3. 挿入位置を計算 (プロシージャヘッダーの次の行)
4. `TextEdit` + `CodeAction` を構築

> **気づき**: このパターンは拡張しやすい設計。新しい CodeAction を追加するには:
> 1. 新しい診断メッセージのフィルタ文字列を定義
> 2. `code_actions()` のループ内に新しい分岐を追加
> 3. 対応する `TextEdit` を生成
>
> ただし「Extract Sub」のような大規模リファクタリングは、
> ソースコードの範囲選択 + 引数解析が必要で、現在のアーキテクチャでは困難。

### 4.4 LIM-08: Folding Range の現状

`folding_range.rs` はプロシージャのみ折りたたみ対応:

```rust
// proc_ranges から FoldingRange を生成
symbols.proc_ranges.iter().map(|(_, range)| {
    FoldingRange { start_line, end_line, kind: FoldingRangeKind::Region, ... }
})
```

**折りたたまれないブロック**:
- If...End If
- For...Next
- With...End With
- Select Case...End Select
- Do...Loop
- Type...End Type (← 既にシンボルテーブルにスパンがあるので追加は容易)
- Enum...End Enum (← 同上)
- コメントブロック

> **気づき**: Type/Enum は `SymbolTable` に `span` があるため、
> `proc_ranges` と同じパターンで即座に追加できる (XS)。
> 制御構造のブロック範囲はAST上で表現されていないため、
> トークンベースの開閉マッチングが必要 (S)。

### 4.5 LIM-07: References のスコープ非考慮

`references.rs:8-24` は純粋なテキスト検索:

```rust
let word = resolve::find_word_at_position(source, position)?;
// ...
for (file_uri, source) in host.all_file_sources() {
    let occurrences = resolve::find_all_word_occurrences(&source, &word);
    // ...
}
```

`find_all_word_occurrences` (`resolve.rs:30-48`) は単語境界チェック付きの
テキストマッチングのみ:

```rust
// bytes[i-1] がident文字でない AND 対象文字列と一致 AND bytes[i+len] がident文字でない
```

> **気づき**: Sub A と Sub B に同名の変数 `x` がある場合、両方の `x` が参照として返される。
> `definition.rs` はスコープ考慮しているのに `references.rs` はしていない、という非対称性がある。
> `document_highlight.rs` も同じ問題を抱えている (同一のロジック)。

### 4.6 LIM-10: Completion のコンテキスト非認識

`completion.rs:35-112` の通常補完 (非ドットアクセス時) は**全候補を返す**:

```
1. 全 72 キーワード
2. 全 88 ビルトイン関数
3. スコープフィルタ済みシンボル
4. 他ファイルの Public シンボル
5. Workbook コンテキスト (シート名等)
```

**コンテキスト区別なし**:

| カーソル位置 | 期待される候補 | 実際 |
|---|---|---|
| ステートメント開始 (`\n    `) | キーワード (`Dim`, `If`, `For` 等) | 全候補 |
| `As ` の後 | 型名 (`Long`, `String`, `MyType`) | 全候補 |
| `=` の後 | 変数、関数、定数 | 全候補 |
| `Call ` の後 | プロシージャ名 | 全候補 |

> **気づき**: カーソルの「直前トークン」を `parse_dot_access_at` 的に
> 後方スキャンで特定し、コンテキストに応じてフィルタする方式が有効。
> ただし LSP クライアントもフィルタリングするため、優先度は低め。

---

## 5. VBA 固有機能の調査

### 5.1 VBA-01: Application 暗黙グローバルの非対称性

`application_globals()` (`types.rs:222-235`) は以下を返す:
```
ActiveWorkbook, ActiveSheet, ActiveCell, Selection, ThisWorkbook,
Workbooks, ScreenUpdating, DisplayAlerts, EnableEvents, StatusBar,
Calculation, Calculate, Run, InputBox
```

**使われている場所** — diagnostics のみ (`diagnostics.rs:105`):
```rust
for name in crate::excel_model::types::application_globals() {
    declared.insert(name.to_ascii_lowercase());
}
```

**使われていない場所** — completion:

`completion.rs:35-112` の通常補完に `application_globals()` の呼び出しがない。
つまり `Range`, `Cells`, `ActiveSheet` 等が**補完候補に出ない**。

> **気づき (重要)**: diagnostics では `Range` を「宣言済み」として扱い未宣言警告を出さないのに、
> completion では `Range` を補完候補として出さない、という**矛盾した動作**。
> 修正は数行:
> ```rust
> // completion.rs の complete() 内に追加
> for name in crate::excel_model::types::application_globals() {
>     items.push(CompletionItem {
>         label: name,
>         kind: Some(CompletionItemKind::PROPERTY),
>         detail: Some("Application".to_string()),
>         ..Default::default()
>     });
> }
> ```

### 5.2 VBA-02: 条件コンパイルの完全不在

レキサーに `#If`, `#Else`, `#End If`, `#Const` のトークン定義がない。

`#` はどのトークンルールにもマッチしないため、
`#If VBA7 Then` はレキサーエラー → スキップされる。

**影響**: `#If VBA7 Then ... #Else ... #End If` を含むファイルで:
- `#If` 行がスキップされ、以降のコードが通常コードとして解析される
- 条件分岐の両方のブランチが同時にパースされる可能性
- `#End If` もスキップされるため、構文エラーにはならない (サイレントスキップ)

### 5.3 VBA-03: Declare ステートメントの影響

`Declare` はレキサーにトークンが定義されていない。
`Declare Sub Foo Lib "kernel32" (...)` は以下のように誤トークン化される:

```
Identifier("Declare") → Sub → Identifier("Foo") → Identifier("Lib") → StringLiteral("kernel32") → ...
```

`parse_module()` が `Identifier("Declare")` を見てスキップした後、
`Sub` トークンを見てプロシージャパースを開始する。
結果として `Foo Lib "kernel32"` がプロシージャ本体として解釈される可能性がある。

> **気づき**: `Declare` の欠如は単なる機能不足ではなく、
> **Declare を含むファイルで AST が壊れる**リスクがある。
> 最低限 `Token::Declare` を追加し、`parse_module` で
> 行末までスキップする処理を入れるべき (XS の防御的修正)。

### 5.4 `load_builtin_types()` のパフォーマンス

`load_builtin_types()` は呼び出しのたびに `Vec<ExcelObjectType>` を新規構築する (`types.rs:34`)。

**呼び出し箇所**:
1. `complete_dot_access()` — 毎回のドット補完リクエスト (`completion.rs:207`)
2. `application_globals()` — diagnostics 計算時 (`types.rs:223`)

つまり**ファイル変更ごとに 2 回**、7 型 × 約 25 メンバーのベクタを構築している。

> **気づき**: 現時点では微小なコスト (ハードコード値の Vec 構築) だが、
> 将来 JSON ファイルからの読み込みに移行した場合 (types.rs:35 のコメント参照)、
> キャッシュ機構が必要になる。`once_cell::sync::Lazy` または
> `std::sync::LazyLock` での初期化が適切。

---

## 6. 横断的な気づき・設計上の懸念

### 6.1 アーキテクチャ上のボトルネック

1. **ブロック構造の不在**: If/For/With/Select/Do がヘッダー行のみの AST ノードで、
   ブロック範囲がない。With コンテキスト、制御構造折りたたみ、Selection Range の
   中間レベルなど、複数の機能が同じ制約に阻まれている。
   **根本対策**: ブロック範囲を `proc_ranges` と同様に事前計算して保持する。

2. **Symbol にモジュール情報がない**: クロスモジュール機能 (モジュール名ドットアクセス、
   モジュール修飾リネーム等) が Symbol → ファイル の逆引きなしでは実装困難。
   **根本対策**: `Symbol` に `module_uri: Option<Url>` を追加するか、
   `AnalysisHost` にモジュール名 → URI のインデックスを追加する。

3. **式の型推論がない**: 変数の `type_name` は宣言時の `As Type` から取るだけで、
   代入や関数戻り値からの推論はない。チェーンドットや関数戻り値ドットアクセスの
   前提条件。**影響範囲**: DOT-04, DOT-05, LIM-02 の一部。

### 6.2 即効性の高い修正 (XS サイズ)

| 項目 | 対象ファイル | 行数目安 | 効果 |
|---|---|---|---|
| Application グローバル補完追加 (VBA-01) | `completion.rs` | ~8 行 | `Range`, `ActiveSheet` 等が補完に出る |
| Declare トークン追加 + スキップ | `lexer.rs`, `parse.rs` | ~10 行 | Declare を含むファイルの AST 破壊を防止 |
| Rename のキーワードチェック | `rename.rs` | ~5 行 | キーワードへのリネームを防止 |
| Type/Enum の FoldingRange 追加 | `folding_range.rs` | ~15 行 | Type/Enum ブロックが折りたためる |
| Enum ドットアクセス (パターン 2: 変数経由) | `completion.rs` | ~15 行 | `Dim c As Color` → `c.Red` が補完に出る |

### 6.3 再利用可能なパターン

コードベースで確立された実装パターン:

| パターン | 使用例 | 再利用先 |
|---|---|---|
| シンボル名検索 | `find_symbol_by_name` | 全ての新規シンボル解決 |
| スコープ考慮検索 | `definition.rs:19-42` の proc_scope マッチング | references, rename のスコープ改善 |
| ショートサーキット + フォールバック | `signature_help.rs:18-32` | 新規の cross-module 解決 |
| 全ファイル走査 | `references.rs:15-21` の `all_file_sources()` | semantic tokens, workspace diagnostics |
| トークン分類マッチング | `formatting.rs` の `is_open_token/is_close_token` | folding range のブロック検出 |
| SymbolDetail パターンマッチ | UDT メンバーフィルタ (`completion.rs:182-200`) | Enum メンバーフィルタ |

### 6.4 優先度の再評価

コード調査の結果、`lsp-feature-gap.md` からの優先度変更を推奨:

| 項目 | 元の優先度 | 推奨変更 | 理由 |
|---|---|---|---|
| VBA-01 (Application グローバル) | P1 ★★★ | **最優先** | 矛盾した動作の修正。8 行で完了 |
| Declare トークン追加 | 未掲載 | **P0 (防御的修正)** | Declare 含むファイルで AST 破壊のリスク |
| LSP-03 (Prepare Rename) | P1 ★★★ | **最優先** | キーワードリネーム防止。5 行で完了 |
| DOT-01 (Enum ドット) | P1 ★★★ | ★★★ 据え置き | 実装コスト XS 確認済み |
| DOT-03 (With ブロック) | P1 ★★★ | **★★☆ に降格** | ブロック構造不在により M ではなく L に近い |
| LSP-01 (Semantic Tokens) | P1 ★★★ | ★★★ 据え置き | データ既存で実現可能性が確認できた |
