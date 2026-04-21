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
  ],

  sprint: {
    number: 52,
    pbi_id: "PBI-47",
    goal: "cargo test 並列実行スレッド数チューニング + Follow-up (A) クローズ — .cargo/config.toml に test-threads 設定を追加し >60s 警告を抑止",
    status: "done",
    subtasks: [
      {
        test: ".cargo/config.toml に [test] test-threads = 4 が設定されること",
        implementation: ".cargo/config.toml 新規作成 or 更新",
        type: "structural",
        status: "completed",
        commits: [],
        notes: [],
      },
      {
        test: "just test が warning なしで完走すること (手動検証)",
        implementation: "cargo test --all 実行確認",
        type: "behavioral",
        status: "completed",
        commits: [],
        notes: ["165 tests green (lib 55 + integration 110), >60s warning なし"],
      },
      {
        test: "plan.md Follow-up (A) が完了マーク済みであること",
        implementation: "plan.md Follow-up セクション更新",
        type: "structural",
        status: "completed",
        commits: [],
        notes: [],
      },
    ],
  },

  definition_of_done: {
    checks: [
      { name: "Tests pass", run: "just test" },
      { name: "Clippy passes", run: "just clippy" },
      { name: "Format passes", run: "cargo fmt --check" },
    ],
  },

  completed: [
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
          action: "apply_formatting の pure function 先行 → LSP handler 配線の Tidy First 順序を Sprint β (N+51) でも維持する",
          timing: "sprint",
          status: "completed",
          outcome: "Sprint N+51 で calculate_line_indents pure helper 先行 → apply_formatting 配線の順序で実現",
        },
        {
          action: "indent 正規化 (β) では ElseIf/Else/Case の depth 例外処理を先にテストで文書化し、実装前に仕様を確定する",
          timing: "sprint",
          status: "completed",
          outcome: "Sprint N+51 で format_indent_else_if_aligned_with_if + format_indent_select_case テストで文書化達成",
        },
        {
          action: "document_end_position の UTF-16 encode_utf16().count() パターンを indent 正規化後も維持 (文字列リテラル内 Unicode 対応)",
          timing: "sprint",
          status: "completed",
          outcome: "Sprint N+51: indent 正規化は行頭ホワイトスペースのみ置換 — UTF-16 座標計算は LSP handler 側 (変更なし) が担当",
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
