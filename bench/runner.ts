// Bench orchestration (spec §11). Pure: every step that touches Excel, the network or the core
// request path comes in through RunnerDeps, so the bench panel can wire in the real adapter and
// core modules, and tests can use fakes. Records hold no cell values, prompts or formulas.
import { quoteSheet } from "../src/core/a1";
import { tokenize } from "../src/core/formulaGate";
import type { CellValue, GateCode, GateResult, Mode, ModelReply } from "../src/core/types";
import { CATEGORIES, compare as defaultCompare, type BenchTask, type Category, type Check, type CompareResult } from "./compare";

export type Condition = "raw" | "substituted" | "structure_only";
export const CONDITIONS: readonly Condition[] = ["raw", "substituted", "structure_only"];
export const DEFAULT_RUNS = 3;

export type Stage = "ok" | "audit" | "send" | "parse" | "clarify" | "gate" | "mismatch";
export const STAGES: readonly Stage[] = ["ok", "audit", "send", "parse", "clarify", "gate", "mismatch"];

/** A request ready to send. `handle` carries whatever the deps need later (audited outbound, map, context). */
export interface Prepared {
  body: string;
  mode: Mode;
  bytes: number;
  handle?: unknown;
}

/** The auditor (or, for raw runs, the marker check) stopped the request. Reasons are codes, not values. */
export interface Blocked {
  blocked: true;
  reasons: string[];
  bytes?: number;
}

export type PrepareResult = Prepared | Blocked;

export type SendOutcome =
  | { ok: true; raw: string; latencyMs: number; tokens?: { prompt: number; completion: number } }
  | { ok: false; message: string; latencyMs?: number; status?: number };

export type ParseOutcome =
  | { ok: true; reply: ModelReply }
  | { ok: false; error: string; /** The reply content, for the one correction retry. */ content?: string | null };

export interface Placement {
  /** Sheet row to write the formula in: the model's placement row, or the first data row. */
  row: number;
  fillDown: boolean;
  firstDataRow: number;
  lastDataRow: number;
}

/**
 * What the runner needs from the host. The bench panel wires these to the real modules:
 * - prepare: raw -> bench/rawSender buildRawBody (marker checked); substituted / structure_only ->
 *   schema, transform, substituteUserText, payload, then the auditor (a block becomes Blocked).
 * - send: raw -> sendRaw; others -> provider.send with the AuditedOutbound kept in `handle`.
 * - parse: reply.parseReply (raw runs parse like substituted ones).
 * - gate: formulaGate.check with the task's sheet context; `phase` says model or restored formula.
 * - restore: restore.restoreFormula with the session map (identity for raw runs).
 * - writeAndRead: clear _nymform_scratch, write the formula at column A of `placement.row`, fill down
 *   to lastDataRow when fillDown is set and the formula doesn't spill, then read back the values of the
 *   written or spilled range. References in the formula are already qualified with the task's sheet.
 * - correct (optional): the §7.9 retry, i.e. payload.buildCorrection plus the auditor.
 */
export interface RunnerDeps {
  prepare(task: BenchTask, condition: Condition, run: number): Promise<PrepareResult> | PrepareResult;
  send(prepared: Prepared, task: BenchTask, condition: Condition, signal?: AbortSignal): Promise<SendOutcome>;
  parse(raw: string, condition: Condition, prepared: Prepared): ParseOutcome;
  gate(formula: string, task: BenchTask, phase: "model" | "restored", prepared: Prepared): GateResult;
  restore(formula: string, prepared: Prepared): string | { formula: string; unknownTokens?: string[] };
  writeAndRead(formula: string, placement: Placement, task: BenchTask): Promise<CellValue[][]>;
  correct?(prepared: Prepared, invalidReply: string, task: BenchTask, condition: Condition): Promise<PrepareResult> | PrepareResult;
  compare?(actual: CellValue[][], check: Check): CompareResult;
  /** Called after each run, e.g. to show progress. */
  onRecord?(record: RunRecord): void;
  signal?: AbortSignal;
  /** Model ID, copied into each record. */
  model?: string;
}

export interface RunRecord {
  task: string;
  category: Category;
  workbook: string;
  condition: Condition;
  run: number;
  pass: boolean;
  stage: Stage;
  /**
   * Plain explanation for a failure: the runner's own text, a compare() detail (positions and
   * counts, never values) or the host's error message.
   */
  detail?: string;
  bytes: number | null;
  promptTokens: number | null;
  completionTokens: number | null;
  latencyMs: number | null;
  gateReasons: GateCode[];
  retries: number;
  model?: string;
}

// ---------------------------------------------------------------------------------------------
// Helpers

function isBlocked(p: PrepareResult): p is Blocked {
  return (p as Blocked).blocked === true;
}

function message(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

/** Row numbers of the task's range: header row 1, data rows after it. */
export function taskRows(task: BenchTask): { headerRow: number; firstDataRow: number; lastDataRow: number } {
  const m = /^[A-Z]+(\d+):[A-Z]+(\d+)$/.exec(task.range);
  if (!m) throw new Error(`Task ${task.id} has an unreadable range.`);
  const top = Number(m[1]);
  return { headerRow: top, firstDataRow: top + 1, lastDataRow: Number(m[2]) };
}

/**
 * Prefixes every unqualified cell or range reference with the sheet, so a formula written on the
 * scratch sheet still computes on the task's data. Everything else stays byte-identical.
 * Throws the tokenizer's error on a formula it can't read (the gate rejects those anyway).
 */
export function qualifyFormula(formula: string, sheet: string): string {
  const prefix = `${quoteSheet(sheet)}!`;
  let out = "";
  let last = 0;
  for (const t of tokenize(formula)) {
    if (t.type !== "ref" || t.text.includes("!")) continue;
    out += formula.slice(last, t.start) + prefix;
    last = t.start;
  }
  return out + formula.slice(last);
}

function placementFor(reply: ModelReply, task: BenchTask): Placement {
  const rows = taskRows(task);
  const m = reply.placement ? /^\$?[A-Za-z]{1,3}\$?(\d+)$/.exec(reply.placement.cell.trim()) : null;
  const row = m ? Math.max(1, Number(m[1])) : rows.firstDataRow;
  return { row, fillDown: reply.placement?.fill_down ?? false, firstDataRow: rows.firstDataRow, lastDataRow: rows.lastDataRow };
}

// ---------------------------------------------------------------------------------------------
// One run

/** Runs one task under one condition once. Never throws: a failing step becomes the record's stage. */
export async function runTask(task: BenchTask, condition: Condition, run: number, deps: RunnerDeps): Promise<RunRecord> {
  const rec: RunRecord = {
    task: task.id,
    category: task.category,
    workbook: task.workbook,
    condition,
    run,
    pass: false,
    stage: "mismatch",
    bytes: null,
    promptTokens: null,
    completionTokens: null,
    latencyMs: null,
    gateReasons: [],
    retries: 0,
  };
  if (deps.model !== undefined) rec.model = deps.model;
  const fail = (stage: Stage, detail: string): RunRecord => {
    rec.stage = stage;
    rec.pass = false;
    rec.detail = detail;
    return rec;
  };
  const addSend = (s: SendOutcome, bytes: number) => {
    rec.bytes = (rec.bytes ?? 0) + bytes;
    if (s.latencyMs !== undefined) rec.latencyMs = (rec.latencyMs ?? 0) + s.latencyMs;
    if (s.ok && s.tokens) {
      rec.promptTokens = (rec.promptTokens ?? 0) + s.tokens.prompt;
      rec.completionTokens = (rec.completionTokens ?? 0) + s.tokens.completion;
    }
  };

  // Prepare (payload + audit, or the raw body behind the marker check).
  let first: PrepareResult;
  try {
    first = await deps.prepare(task, condition, run);
  } catch (e) {
    return fail("audit", `Couldn't prepare the request: ${message(e)}`);
  }
  if (isBlocked(first)) {
    if (first.bytes !== undefined) rec.bytes = first.bytes;
    return fail("audit", `Blocked before sending${first.reasons.length ? `: ${first.reasons.join(", ")}` : ""}.`);
  }
  let current: Prepared = first;

  // Send.
  const sendOnce = async (p: Prepared): Promise<SendOutcome> => {
    let s: SendOutcome;
    try {
      s = await deps.send(p, task, condition, deps.signal);
    } catch (e) {
      s = { ok: false, message: message(e) };
    }
    addSend(s, p.bytes);
    return s;
  };
  let sent = await sendOnce(current);
  if (!sent.ok) return fail("send", sent.message);

  // Parse, with the one correction retry (spec §7.9) when the host supports it.
  const parseSafe = (raw: string, p: Prepared): ParseOutcome => {
    try {
      return deps.parse(raw, condition, p);
    } catch (e) {
      return { ok: false, error: message(e) };
    }
  };
  let parsed = parseSafe(sent.raw, current);
  if (!parsed.ok && deps.correct) {
    rec.retries = 1;
    let retry: PrepareResult;
    try {
      retry = await deps.correct(current, parsed.content ?? sent.raw, task, condition);
    } catch (e) {
      return fail("parse", `The reply couldn't be read, and the retry couldn't be prepared: ${message(e)}`);
    }
    if (isBlocked(retry)) return fail("audit", "The correction retry was blocked before sending.");
    current = retry;
    sent = await sendOnce(current);
    if (!sent.ok) return fail("send", sent.message);
    parsed = parseSafe(sent.raw, current);
  }
  if (!parsed.ok) return fail("parse", `The reply couldn't be read: ${parsed.error}`);
  const reply = parsed.reply;
  if (reply.kind === "clarify") return fail("clarify", "The model asked a clarifying question.");
  if (reply.kind !== "formula" || !reply.formula) return fail("parse", "The reply held no formula.");
  const prepared = current;

  // Gate, restore, gate again.
  const gateSafe = (formula: string, phase: "model" | "restored"): GateResult => {
    try {
      return deps.gate(formula, task, phase, prepared);
    } catch {
      return { ok: false, reasons: [{ code: "G-PARSE", detail: "The gate stopped on an unexpected problem." }] };
    }
  };
  const before = gateSafe(reply.formula, "model");
  if (!before.ok) {
    rec.gateReasons = [...new Set(before.reasons.map((r) => r.code))];
    return fail("gate", "The model's formula was rejected by the gate.");
  }
  let restored: string;
  try {
    const r = deps.restore(reply.formula, prepared);
    restored = typeof r === "string" ? r : r.formula;
  } catch (e) {
    return fail("gate", `Couldn't restore the formula: ${message(e)}`);
  }
  const after = gateSafe(restored, "restored");
  if (!after.ok) {
    rec.gateReasons = [...new Set(after.reasons.map((r) => r.code))];
    return fail("gate", "The restored formula was rejected by the gate.");
  }

  // Write to the scratch sheet, read back, compare.
  let actual: CellValue[][];
  try {
    const placement = placementFor(reply, task);
    actual = await deps.writeAndRead(qualifyFormula(restored, task.sheet), placement, task);
  } catch (e) {
    return fail("mismatch", `Couldn't write or read the scratch sheet: ${message(e)}`);
  }
  let result: CompareResult;
  try {
    result = (deps.compare ?? defaultCompare)(actual, task.check);
  } catch (e) {
    return fail("mismatch", `Comparison failed: ${message(e)}`);
  }
  if (!result.pass) return fail("mismatch", result.detail);
  rec.pass = true;
  rec.stage = "ok";
  return rec;
}

/**
 * Runs every task under every condition, `runs` times each. Runs are interleaved (run 1 of every
 * task and condition, then run 2, ...) so a slow or failing provider period hits all conditions
 * alike. Stops early, returning what it has, when deps.signal aborts.
 */
export async function runBench(
  tasks: readonly BenchTask[],
  conditions: readonly Condition[] = CONDITIONS,
  runs: number = DEFAULT_RUNS,
  deps?: RunnerDeps,
): Promise<RunRecord[]> {
  if (!deps) throw new Error("runBench needs its deps.");
  const records: RunRecord[] = [];
  for (let run = 1; run <= runs; run++) {
    for (const task of tasks) {
      for (const condition of conditions) {
        if (deps.signal?.aborted) return records;
        const rec = await runTask(task, condition, run, deps);
        records.push(rec);
        deps.onRecord?.(rec);
      }
    }
  }
  return records;
}

// ---------------------------------------------------------------------------------------------
// Output

/** One JSON object per line, as written to bench/results/{date}-{model}.jsonl. */
export function toJsonl(records: readonly RunRecord[]): string {
  return records.map((r) => JSON.stringify(r)).join("\n") + (records.length ? "\n" : "");
}

/** bench/results file name: date plus model ID with "/" replaced by "_". */
export function resultsFileName(date: string, model: string): string {
  return `${date}-${model.replace(/\//g, "_")}.jsonl`;
}

export interface Rate {
  pass: number;
  runs: number;
  /** pass / runs, or null without runs. */
  rate: number | null;
}

export interface Medians {
  bytes: number | null;
  promptTokens: number | null;
  completionTokens: number | null;
  latencyMs: number | null;
}

export interface TaskSummary {
  task: string;
  category: Category;
  rates: Partial<Record<Condition, number | null>>;
  /** Null when the task can't be compared (no raw runs, or no substituted or structure-only runs). */
  matched: boolean | null;
}

export interface Summary {
  model?: string;
  tasks: number;
  runs: number;
  conditions: Condition[];
  /** Accuracy by category (plus "all") and condition. */
  accuracy: Record<string, Partial<Record<Condition, Rate>>>;
  medians: Partial<Record<Condition, Medians>>;
  stages: Partial<Record<Condition, Record<Stage, number>>>;
  gateRejections: Partial<Record<Condition, { runs: number; codes: Partial<Record<GateCode, number>> }>>;
  perTask: TaskSummary[];
  /** Tasks whose best production-condition pass rate is at least their raw pass rate. */
  matched: number;
  /** Tasks with both raw and production-condition runs. */
  comparable: number;
  /** Matched tasks where raw and both production conditions all scored 0. */
  matchedAtZero: number;
  /** Leak count from bench/leakcheck.ts over substituted and structure-only requests, when given. */
  leaks?: number;
  headline: string;
  headlineDefinition: string;
}

export const HEADLINE_DEFINITION =
  "A task matched raw-data accuracy when its best pass rate under substituted or structure-only " +
  "(whichever is higher) was at least its pass rate under raw. Tasks without raw runs or without " +
  "substituted and structure-only runs don't count as matched, and neither do tasks where every " +
  "condition scored 0%. \"Zero raw values sent\" is the " +
  "leak check's result over every substituted and structure-only request.";

function median(values: readonly number[]): number | null {
  if (values.length === 0) return null;
  const s = [...values].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 ? s[mid]! : (s[mid - 1]! + s[mid]!) / 2;
}

function rateOf(rs: readonly RunRecord[]): Rate {
  const pass = rs.filter((r) => r.pass).length;
  return { pass, runs: rs.length, rate: rs.length ? pass / rs.length : null };
}

/**
 * Accuracy by category and condition, medians, failure stages, gate rejections and the headline.
 * Pass `leaks` (the leak check's count over substituted and structure-only requests) to state the
 * headline's "zero raw values sent" part; without it the headline says the leak check wasn't run.
 */
export function summarize(records: readonly RunRecord[], opts: { leaks?: number; model?: string } = {}): Summary {
  const conditions = CONDITIONS.filter((c) => records.some((r) => r.condition === c));
  const taskIds = [...new Set(records.map((r) => r.task))];
  const categoryOf = new Map(records.map((r) => [r.task, r.category] as const));

  const accuracy: Summary["accuracy"] = {};
  const categories = [...CATEGORIES.filter((c) => records.some((r) => r.category === c)), "all"];
  for (const cat of categories) {
    const row: Partial<Record<Condition, Rate>> = {};
    for (const cond of conditions) {
      row[cond] = rateOf(records.filter((r) => r.condition === cond && (cat === "all" || r.category === cat)));
    }
    accuracy[cat] = row;
  }

  const medians: Summary["medians"] = {};
  const stages: Summary["stages"] = {};
  const gateRejections: Summary["gateRejections"] = {};
  for (const cond of conditions) {
    const rs = records.filter((r) => r.condition === cond);
    const pick = (k: keyof Medians) => median(rs.map((r) => r[k]).filter((v): v is number => typeof v === "number"));
    medians[cond] = { bytes: pick("bytes"), promptTokens: pick("promptTokens"), completionTokens: pick("completionTokens"), latencyMs: pick("latencyMs") };
    const counts = Object.fromEntries(STAGES.map((s) => [s, 0])) as Record<Stage, number>;
    for (const r of rs) counts[r.stage]++;
    stages[cond] = counts;
    const gated = rs.filter((r) => r.stage === "gate");
    const codes: Partial<Record<GateCode, number>> = {};
    for (const r of gated) for (const c of r.gateReasons) codes[c] = (codes[c] ?? 0) + 1;
    gateRejections[cond] = { runs: gated.length, codes };
  }

  const perTask: TaskSummary[] = taskIds.map((id) => {
    const rates: Partial<Record<Condition, number | null>> = {};
    for (const cond of conditions) {
      const rs = records.filter((r) => r.task === id && r.condition === cond);
      if (rs.length) rates[cond] = rateOf(rs).rate;
    }
    const raw = rates.raw;
    const best = Math.max(...(["substituted", "structure_only"] as const).map((c) => rates[c] ?? -1));
    const matched = raw === undefined || raw === null || best < 0 ? null : best >= raw;
    return { task: id, category: categoryOf.get(id)!, rates, matched };
  });
  const allZero = (t: (typeof perTask)[number]) =>
    t.rates.raw === 0 && (t.rates.substituted ?? 0) === 0 && (t.rates.structure_only ?? 0) === 0;
  // A tie where every condition failed is not "matching raw-data accuracy" in the headline.
  const matched = perTask.filter((t) => t.matched === true && !allZero(t)).length;
  const comparable = perTask.filter((t) => t.matched !== null).length;
  const matchedAtZero = perTask.filter(
    (t) => t.matched === true && allZero(t),
  ).length;

  const n = taskIds.length;
  let headline: string;
  if (opts.leaks === 0) headline = `${matched} of ${n} tasks matched raw-data accuracy with zero raw values sent.`;
  else if (opts.leaks === undefined) headline = `${matched} of ${n} tasks matched raw-data accuracy (leak check not run yet).`;
  else headline = `${matched} of ${n} tasks matched raw-data accuracy, but the leak check found ${opts.leaks} raw values sent.`;

  const summary: Summary = {
    tasks: n,
    runs: records.reduce((m, r) => Math.max(m, r.run), 0),
    conditions,
    accuracy,
    medians,
    stages,
    gateRejections,
    perTask,
    matched,
    comparable,
    matchedAtZero,
    headline,
    headlineDefinition: HEADLINE_DEFINITION,
  };
  const model = opts.model ?? records.find((r) => r.model)?.model;
  if (model) summary.model = model;
  if (opts.leaks !== undefined) summary.leaks = opts.leaks;
  return summary;
}

function pct(r: Rate | undefined): string {
  if (!r || r.rate === null) return "–";
  return `${Math.round(r.rate * 100)}% (${r.pass}/${r.runs})`;
}

function fmt(n: number | null | undefined): string {
  return n === null || n === undefined ? "–" : String(Math.round(n));
}

/** The summary as Markdown tables, for the README and bench/results. Failures included. */
export function toMarkdown(s: Summary): string {
  const conds = s.conditions;
  const lines: string[] = [];
  lines.push(`**${s.headline}**`, "");
  if (s.model) lines.push(`Model: \`${s.model}\`. ${s.tasks} tasks, ${s.runs} runs per condition.`, "");
  lines.push(`| Category | ${conds.join(" | ")} |`, `| --- |${conds.map(() => " --- |").join("")}`);
  for (const [cat, row] of Object.entries(s.accuracy)) {
    lines.push(`| ${cat === "all" ? "**all**" : cat} | ${conds.map((c) => pct(row[c])).join(" | ")} |`);
  }
  lines.push("", `| Median per run | ${conds.join(" | ")} |`, `| --- |${conds.map(() => " --- |").join("")}`);
  for (const k of ["bytes", "promptTokens", "completionTokens", "latencyMs"] as const) {
    lines.push(`| ${k} | ${conds.map((c) => fmt(s.medians[c]?.[k])).join(" | ")} |`);
  }
  lines.push("", `| Failure stage | ${conds.join(" | ")} |`, `| --- |${conds.map(() => " --- |").join("")}`);
  for (const st of STAGES) {
    lines.push(`| ${st} | ${conds.map((c) => String(s.stages[c]?.[st] ?? 0)).join(" | ")} |`);
  }
  lines.push("", "Gate rejections:");
  for (const c of conds) {
    const g = s.gateRejections[c];
    const codes = g ? Object.entries(g.codes).map(([k, v]) => `${k} ${v}`).join(", ") : "";
    lines.push(`- ${c}: ${g?.runs ?? 0}${codes ? ` (${codes})` : ""}`);
  }
  lines.push("", `Leak check: ${s.leaks === undefined ? "not run" : `${s.leaks} leaks in substituted and structure-only requests`}.`);
  lines.push(`Matched: ${s.matched} of ${s.tasks} (${s.comparable} comparable; ${s.matchedAtZero} of the matches are ties at 0%).`);
  lines.push("", s.headlineDefinition, "");
  return lines.join("\n");
}
