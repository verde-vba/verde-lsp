// ============================================================
// Dashboard Data (AI edits this section)
// ============================================================

const userStoryRoles = ["VBA開発者"] as const satisfies readonly string[];

const scrum: ScrumDashboard = {
  product_goal: {
    statement: "VBA開発者が補完・hover・goto-def・rename 等の LSP 機能を通じて、Verde desktop 上で効率よく VBA コードを記述できる",
    success_metrics: [
      { metric: "LSP 機能カバレッジ (Phase 1-3 PBI 完了率)", target: "100%" },
      { metric: "cargo test green", target: "全テスト pass" },
      { metric: "Windows / macOS / Linux CI", target: "matrix 全 pass" },
    ],
  },

  product_backlog: [
    { id: "PBI-44", story: { role: "VBA開発者", capability: "Me キーワードとインスタンス変数 補完/hover/goto-def", benefit: "クラスベース VBA 開発でも IDE 支援" }, acceptance_criteria: [], status: "done" },
    { id: "PBI-45", story: { role: "VBA開発者", capability: "PivotTable/Chart/Shape 補完候補", benefit: "API 名タイプミス防止" }, acceptance_criteria: [], status: "done" },
    {
      id: "PBI-46",
      story: {
        role: "VBA開発者",
        capability: "VBA ファイルの indent と識別子 case をフォーマッタで自動整形する",
        benefit: "コードスタイルを統一でき、レビューの負担が減る",
      },
      acceptance_criteria: [
        {
          criterion: "textDocument/formatting で indent 整形が機能する",
          verification: "formatting テストが green (既存 8 + 新規 indent tests)",
        },
        {
          criterion: "ElseIf/Else/Case は depth-1 (VBA 慣習) — テストで文書化",
          verification: "format_indent_else_if_aligned_with_if / format_indent_select_case テストが green",
        },
        {
          criterion: "cargo clippy -D warnings 0 件 / cargo fmt pass",
          verification: "just clippy && just fmt",
        },
      ],
      status: "done",
    },
    {
      id: "PBI-47",
      story: {
        role: "VBA開発者",
        capability: "cargo test の並列実行スレッド数を制限し、>60s タイムアウト警告を抑止する",
        benefit: "CI フィードバックの信頼性が上がる",
      },
      acceptance_criteria: [
        {
          criterion: ".cargo/config.toml に test.test-threads 設定が追加される",
          verification: "just test が >60s 警告なしで完走",
        },
        {
          criterion: "ci.yml の Ubuntu/Windows 両 runner で cargo test --all が pass",
          verification: "GitHub Actions CI green",
        },
      ],
      status: "done",
    },
    {
      id: "PBI-48",
      story: {
        role: "VBA開発者",
        capability: "textDocument/inlayHint で Dim 宣言変数・定数の型を inline 表示する",
        benefit: "型宣言を読まずにカーソル近くで型を把握でき、コードの可読性が上がる",
      },
      acceptance_criteria: [
        {
          criterion: "Dim x As String の変数 x の名前末尾に ': String' ヒントが返る",
          verification: "inlay_hint_shows_dim_variable_type テストが green",
        },
        {
          criterion: "型宣言なし (Dim x) の変数には ': Variant' ヒントが返る",
          verification: "inlay_hint_shows_variant_for_untyped_dim テストが green",
        },
        {
          criterion: "Const PI As Double = 3.14 の定数 PI に ': Double' ヒントが返る",
          verification: "inlay_hint_shows_const_type テストが green",
        },
        {
          criterion: "inlayHintProvider が server capabilities に宣言される",
          verification: "server_capabilities_declares_inlay_hint_provider テストが green",
        },
        {
          criterion: "cargo clippy -D warnings 0 件 / cargo fmt pass",
          verification: "just clippy && just fmt",
        },
      ],
      status: "done",
    },
    {
      id: "PBI-49",
      story: {
        role: "VBA開発者",
        capability: "textDocument/prepareCallHierarchy + callHierarchy/incomingCalls + outgoingCalls で関数呼び出しツリーをナビゲートする",
        benefit: "呼び出し元・呼び出し先を視覚的に辿れ、大規模 VBA モジュールのリファクタが容易になる",
      },
      acceptance_criteria: [
        {
          criterion: "Sub/Function/Property 上でカーソルを置いたとき prepareCallHierarchy が CallHierarchyItem を返す",
          verification: "prepare_call_hierarchy_returns_item_for_sub テストが green",
        },
        {
          criterion: "incomingCalls が呼び出し元手続きを返す (宣言行は除外)",
          verification: "incoming_calls_returns_callers テストが green",
        },
        {
          criterion: "outgoingCalls が手続き本体内で呼び出している手続き名を返す",
          verification: "outgoing_calls_returns_callees テストが green",
        },
        {
          criterion: "クロスファイルの呼び出し元を incomingCalls が返す",
          verification: "incoming_calls_cross_file テストが green",
        },
        {
          criterion: "call_hierarchy_provider が server capabilities に宣言される",
          verification: "server_capabilities_declares_call_hierarchy_provider テストが green",
        },
      ],
      status: "done",
    },
  ],

  sprint: null,

  definition_of_done: {
    checks: [
      { name: "Tests pass", run: "just test" },
      { name: "Clippy passes", run: "just clippy" },
      { name: "Format passes", run: "cargo fmt --check" },
    ],
  },

  completed: [
    {
      number: 54,
      pbi_id: "PBI-49",
      goal: "textDocument/prepareCallHierarchy + callHierarchy/incomingCalls + outgoingCalls — find_all_word_occurrences + proc_ranges を活用したテキストベース call hierarchy",
      status: "done",
      subtasks: [],
    },
    {
      number: 53,
      pbi_id: "PBI-48",
      goal: "textDocument/inlayHint — Dim 変数・定数の型を inline 表示 (Symbol.type_name 再利用、Variant fallback あり)",
      status: "done",
      subtasks: [],
    },
    {
      number: 52,
      pbi_id: "PBI-47",
      goal: "cargo test 並列実行スレッド数チューニング + Follow-up (A) クローズ — .cargo/config.toml に test-threads = 4 を設定し >60s 警告を抑止",
      status: "done",
      subtasks: [],
    },
    {
      number: 51,
      pbi_id: "PBI-46 (β)",
      goal: "indent 正規化 (depth tracking) — Sub/If/For/With/Select/Do/While/Type でネスト、ElseIf/Else/Case は depth-1 例外、Public/Private 修飾子スキップ",
      status: "done",
      subtasks: [],
    },
    {
      number: 50,
      pbi_id: "PBI-46 (α)",
      goal: "textDocument/formatting — keyword case 正規化 + 行末空白除去 + LSP provider 配線",
      status: "done",
      subtasks: [],
    },
    {
      number: 49,
      pbi_id: "PBI-46 (planning)",
      goal: "PBI-46 実装見積・TDD方針・Sprint分割を plan.md に文書化 — N+50 二値判断可能状態",
      status: "done",
      subtasks: [],
    },
    {
      number: 48,
      pbi_id: "PBI-45",
      goal: "PivotTable / Chart / Shape の dot-access 補完が動作し、既存 Range/Worksheet/Workbook/Application の補完が引き続き green であること",
      status: "done",
      subtasks: [],
    },
    {
      number: 45,
      pbi_id: "PBI-43 (partial)",
      goal: "Type ブロック parser 実装 + UdtMember シンボル登録",
      status: "done",
      subtasks: [],
    },
    {
      number: 46,
      pbi_id: "PBI-43",
      goal: "dot-access 補完 / hover / goto-def 実装 — PBI-43 全完了",
      status: "done",
      subtasks: [],
    },
    {
      number: 47,
      pbi_id: "PBI-44",
      goal: "Me. でカレントクラスモジュールのメンバー補完 + .cls ヘッダー検証",
      status: "done",
      subtasks: [],
    },
  ],

  retrospectives: [
    {
      sprint: 54,
      improvements: [
        {
          action: "proc_ranges (body span) + find_all_word_occurrences + decl_spans 除外の 3 要素が call hierarchy の incoming/outgoing 両方を支えた — 新 LSP 機能追加前に proc_ranges の活用可否を確認するパターンを定着させる",
          timing: "sprint",
          status: "completed",
          outcome: "prepare + incomingCalls (cross-file) + outgoingCalls を src/call_hierarchy.rs 1 ファイルで S サイズ完結 (170 → 176 green)",
        },
        {
          action: "テキストベーススキャンは型名 (Dim x As Foo) を誤って呼び出しとして含む可能性がある — 精度改善が必要な場合は AST ベースの call-site 検出 PBI を将来追加する",
          timing: "product",
          status: "active",
          outcome: null,
        },
      ],
    },
    {
      sprint: 53,
      improvements: [
        {
          action: "Symbol.type_name の既存格納を Probe で確認してから実装着手 — 新機能追加前に SymbolTable の既存フィールドを確認する習慣が実装コストを正確に見積もる鍵",
          timing: "sprint",
          status: "completed",
          outcome: "新規ファイル 1 本 + server.rs 配線のみで PBI-48 を S で完結 (165 → 170 green)",
        },
        {
          action: "offset_to_position (UTF-16 対応済み) を inlay_hint.rs で再利用 — LSP 座標変換は既存ユーティリティを探してから実装すること",
          timing: "product",
          status: "completed",
          outcome: "encode_utf16().count() の重複実装を回避、UTF-16 座標精度を保証",
        },
      ],
    },
    {
      sprint: 52,
      improvements: [
        {
          action: "Probe-first アプローチ: ci.yml を読む前に実装コストを見積もらず、まずコードを確認したことで (A) が already-done と即判明 — 「実装前に現状確認」を次 PBI でも踏襲する",
          timing: "sprint",
          status: "completed",
          outcome: "Follow-up (A) 実装ゼロ達成 / Follow-up (B) XS で完結 / 合計コスト最小",
        },
        {
          action: ".cargo/config.toml の test-threads 設定は CI と local の両方で効く — 今後 >60s 警告が再発した場合は thread 数を 1 に絞って直列化デバッグする選択肢も持つ",
          timing: "product",
          status: "completed",
          outcome: "165 tests green, >60s 警告なし",
        },
      ],
    },
    {
      sprint: 51,
      improvements: [
        {
          action: "calculate_line_indents pure helper 先行 (Tidy First) → apply_formatting 配線の順序が Sprint β でも機能した — 構造変更と振る舞い変更の分離パターンを次 PBI でも踏襲する",
          timing: "sprint",
          status: "completed",
          outcome: "165 green (前 155 + 新規 10) / clippy 0 / fmt pass",
        },
        {
          action: "ElseIf/Else/Case depth-1 例外を実装前にテストで文書化 (Sprint N+50 KPT Try 達成) — 仕様曖昧さを解消してから実装できた",
          timing: "sprint",
          status: "completed",
          outcome: "format_indent_else_if_aligned_with_if + format_indent_select_case_aligned_with_select テストが仕様書として機能",
        },
        {
          action: "first_block_token で Public/Private/Friend/Static をスキップする設計 — Declare Function のような外部宣言を誤って open token 扱いしない安全設計",
          timing: "sprint",
          status: "completed",
          outcome: "format_indent_public_sub_open_token テストで検証済み",
        },
        {
          action: "Follow-up (A): E2E stdio テストを Windows CI matrix に追加する後続タスク — ci.yml の os: [ubuntu-latest, windows-latest] + cargo test --all が既に設定済み、tests/e2e_stdio.rs の stdio_lifecycle_completes_gracefully は Sprint N+44 で追加済みであり達成済みと確認",
          timing: "sprint",
          status: "completed",
          outcome: "実装変更不要。CI matrix と E2E テストは既に整合していることを Probe で確認済み (2026-04-21)",
        },
      ],
    },
    {
      sprint: 50,
      improvements: [
        {
          action: "[compact] PBI-46 α/β 完了: pure function 先行 → handler 配線の Tidy First 順序、ElseIf/Else/Case depth 例外をテストで先行文書化、UTF-16 座標計算は handler 側に委譲",
          timing: "sprint",
          status: "completed",
          outcome: "Sprint N+51 で全パターン達成確認済み",
        },
      ],
    },
  ],
};

// ============================================================
// Type Definitions (DO NOT MODIFY - request human review for schema changes)
// ============================================================

type PBIStatus = "draft" | "refining" | "ready" | "done";
type SprintStatus = "planning" | "in_progress" | "review" | "done" | "cancelled";
type SubtaskStatus = "pending" | "red" | "green" | "refactoring" | "completed";
type SubtaskType = "behavioral" | "structural";
type CommitPhase = "green" | "refactoring";
type ImprovementTiming = "immediate" | "sprint" | "product";
type ImprovementStatus = "active" | "completed" | "abandoned";

interface SuccessMetric { metric: string; target: string; }
interface ProductGoal { statement: string; success_metrics: SuccessMetric[]; }
interface AcceptanceCriterion { criterion: string; verification: string; }
interface UserStory { role: (typeof userStoryRoles)[number]; capability: string; benefit: string; }
interface PBI { id: string; story: UserStory; acceptance_criteria: AcceptanceCriterion[]; status: PBIStatus; }
interface Commit { hash: string; message: string; phase: CommitPhase; }
interface Subtask { test: string; implementation: string; type: SubtaskType; status: SubtaskStatus; commits: Commit[]; notes: string[]; }
interface Sprint { number: number; pbi_id: string; goal: string; status: SprintStatus; subtasks: Subtask[]; }
interface DoDCheck { name: string; run: string; }
interface DefinitionOfDone { checks: DoDCheck[]; }
interface Improvement { action: string; timing: ImprovementTiming; status: ImprovementStatus; outcome: string | null; }
interface Retrospective { sprint: number; improvements: Improvement[]; }
interface ScrumDashboard {
  product_goal: ProductGoal;
  product_backlog: PBI[];
  sprint: Sprint | null;
  definition_of_done: DefinitionOfDone;
  completed: Sprint[];
  retrospectives: Retrospective[];
}

console.log(JSON.stringify(scrum, null, 2));
