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
    {
      id: "PBI-44",
      story: {
        role: "VBA開発者",
        capability: "Class module (.cls) の Me キーワードとインスタンス変数に対して補完・hover・goto-def を使う",
        benefit: "クラスベース VBA 開発でも IDE 支援が得られ生産性が向上する",
      },
      acceptance_criteria: [
        {
          criterion: "Class モジュールの Property/Method 定義が SymbolTable に登録される",
          verification: "cargo test でシンボル登録テストが green",
        },
        {
          criterion: "`Me.` で当該クラスのメンバーが補完候補に現れる",
          verification: "completion テストが green",
        },
        {
          criterion: "インスタンス変数に hover すると型情報が表示される",
          verification: "hover テストが green",
        },
        {
          criterion: "goto-def でメンバー定義行にジャンプできる",
          verification: "definition テストが green",
        },
        {
          criterion: "cargo clippy -D warnings 0 件 / cargo fmt pass",
          verification: "just clippy && just fmt",
        },
      ],
      status: "done",
    },
    {
      id: "PBI-45",
      story: {
        role: "VBA開発者",
        capability: "PivotTable / Chart / Shape 等の Excel オブジェクトモデルを補完候補で選択できる",
        benefit: "Excel 高度操作の API 名をタイプミスせず入力できる",
      },
      acceptance_criteria: [
        {
          criterion: "`Dim pt As PivotTable\\npt.` と入力した時点で PivotTable の主要プロパティ・メソッドが補完候補に現れる",
          verification: "pt_dot_completion_returns_pivottable_members テストが green",
        },
        {
          criterion: "`Dim ch As Chart\\nch.` と入力した時点で Chart の主要プロパティ・メソッドが補完候補に現れる",
          verification: "chart_dot_completion_returns_chart_members テストが green",
        },
        {
          criterion: "`Dim sh As Shape\\nsh.` と入力した時点で Shape の主要プロパティ・メソッドが補完候補に現れる",
          verification: "shape_dot_completion_returns_shape_members テストが green",
        },
        {
          criterion: "既存 Range / Worksheet / Workbook / Application の dot-access 補完が引き続き動作する",
          verification: "既存 completion テスト群が green のまま維持される",
        },
        {
          criterion: "cargo clippy -D warnings 0 件 / cargo fmt pass",
          verification: "just clippy && just fmt",
        },
      ],
      status: "done",
    },
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
          verification: "formatting テストが green",
        },
      ],
      status: "draft",
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
      sprint: 46,
      improvements: [
        {
          action: "Tidy First の per-member span が goto-def 精度を解決した — 構造変更が振る舞い変更を代替するパターンを踏襲する",
          timing: "sprint",
          status: "completed",
          outcome: "Sprint N+47 で hover/goto-def が実装変更ゼロで動作することを確認",
        },
        {
          action: "cargo fmt 後の #[test] 属性重複/消失バグを避けるためテスト挿入位置に注意する",
          timing: "sprint",
          status: "completed",
          outcome: "Sprint N+47 で cat >> でテスト追記後に cargo fmt を適用して解消",
        },
      ],
    },
    {
      sprint: 47,
      improvements: [
        {
          action: "Me. 特殊ケースを追加するだけで Class module の補完が実現できた — 既存の dot-access 設計が拡張点として機能した",
          timing: "sprint",
          status: "completed",
          outcome: "Sprint N+48 で同パターンを PivotTable/Chart/Shape へ適用し、builtin type fallback として一般化に成功",
        },
        {
          action: "テスト追記は cat >> より Edit ツールを使う方が fmt 差分を事前に制御しやすい",
          timing: "sprint",
          status: "completed",
          outcome: "Sprint N+48 で Edit ツールをテスト追記に使用し、fmt 差分ゼロを維持",
        },
      ],
    },
    {
      sprint: 48,
      improvements: [
        {
          action: "excel_model/types.rs への型定義追加 → completion.rs の fallback 実装という Tidy First 順序を維持する",
          timing: "sprint",
          status: "active",
          outcome: null,
        },
        {
          action: "regression テスト (existing_range_dot_completion_still_works) を各 PBI で必ず追加し、既存補完の劣化を即検出できる体制を保つ",
          timing: "sprint",
          status: "active",
          outcome: null,
        },
        {
          action: "PBI-46 (textDocument/formatting) 着手前にスコープ見積もりを行い、indent 整形のみ / case 整形のみ に分割できるか検討する",
          timing: "sprint",
          status: "active",
          outcome: null,
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
