// Regression battery (fix pass, A5): ordinary analyst questions on the four bench workbooks, with the
// columns Nymform suggests keeping private.
//
// - Structure only must never block: nothing from a cell is sent, so a block there is a false one.
// - Nor may an ordinary question be rewritten on the way out (a word read as part of a private
//   value and mapped to its stand-in), in either mode: that is worse than a block.
// - With sample rows (substituted) a block can be right or wrong, so that pass measures: blocks per
//   workbook, pinned in the snapshot. Jorge decides what to do with a high rate (above 10% in any
//   workbook); this round doesn't change matching rules to lower it.
//
// `npm run battery` prints the table.
import { describe, expect, it } from "vitest";
import { PROMPTS, generateData, toRangeData, type BenchWorkbook } from "../bench/generate";
import type { AuditLocation, ProdMode, RangeData } from "../src/core/types";
import { Flow, type ColumnView, type HostAdapter, type Prepared } from "../src/taskpane/flow";

const MODEL = "openai/gpt-6-luna";
const REPORT = process.env.npm_lifecycle_event === "battery" || process.env.NYMFORM_BATTERY_REPORT === "1";
/** Above this block rate with sample rows, the round stops for Jorge's decision. */
export const SUBSTITUTED_LIMIT = 0.1;

const DATA = generateData();

function hostFor(data: RangeData): HostAdapter {
  return {
    readSelection: async () => data,
    listSheets: async () => [data.sheet],
    listTables: async () => [],
    readRange: async () => {
      throw new Error("no such range");
    },
    insertFormula: async () => ({ address: "", excelError: false }),
    isRangeEmpty: async () => true,
    isTooLarge: () => false,
  };
}

async function flowFor(key: string, data: RangeData): Promise<{ flow: Flow; columns: readonly ColumnView[] }> {
  const flow = new Flow({ adapter: hostFor(data), send: async () => ({ ok: false, message: "no network in tests" }) as never, model: MODEL, host: "openrouter.ai", build: "test" });
  const r = await flow.refresh();
  if (!r.ok) throw new Error(`${key}: ${r.message}`);
  return { flow, columns: flow.getState().ranges[0]!.columns };
}

/** True for a column of times of day ("h:mm AM/PM"): it gets no date phrases. */
function isTimeColumn(data: RangeData, c: number): boolean {
  const format = data.numberFormat.slice(1).map((r) => r[c]).find((f) => typeof f === "string" && f !== "General") ?? "";
  return /h/iu.test(format) && !/[dy]/iu.test(format);
}

/** Numbers of a column, sorted. */
function numbersOf(data: RangeData, c: number): number[] {
  return data.values
    .slice(1)
    .map((r) => r[c])
    .filter((v): v is number => typeof v === "number" && Number.isFinite(v))
    .sort((a, b) => a - b);
}

/**
 * The given prompts, then questions an analyst might ask, from the headers, types and the values of
 * the columns sent as they are: totals, averages and top 5 per category, the column's own values
 * ("Total Amount for West"), thresholds at the column's median, date phrases, TRUE/FALSE flags.
 * They are built from headers, sent values and fixed words, and avoid private values; one that
 * still names one shows up as rewritten.
 */
export function questionsFor(prompts: readonly string[], columns: readonly ColumnView[], data: RangeData): string[] {
  const idx = (c: ColumnView) => columns.indexOf(c);
  const nums = columns.filter((c) => c.type === "number");
  const dates = columns.filter((c) => c.type === "date" && !isTimeColumn(data, idx(c))).map((c) => c.header);
  const bools = columns.filter((c) => c.type === "boolean").map((c) => c.header);
  // Categories: text columns Nymform sends as they are, with few distinct values (Region, Department).
  const catCols = columns.filter((c) => c.type === "text" && !c.policy.private && c.stats.distinct <= 30);
  // A sent column can hold a private value too (a "Referred by" column of customer names).
  const privateTexts = new Set(
    columns.flatMap((c) => (c.policy.private ? data.text.slice(1).map((r) => (r[idx(c)] ?? "").trim().toLowerCase()) : [])),
  );
  const cats = catCols.map((c) => c.header);
  const out = [...prompts];
  for (const col of nums) {
    const n = col.header;
    // A threshold near the median, rounded to two significant digits. In a private column it is
    // never a value a cell holds or shows: the median itself can be a private salary.
    const sorted = numbersOf(data, idx(col));
    const held = new Set(sorted.map(String));
    for (const r of data.text.slice(1)) held.add((r[idx(col)] ?? "").trim());
    const median = sorted[Math.floor(sorted.length / 2)] ?? 1000;
    const step = median === 0 ? 1 : 10 ** (Math.floor(Math.log10(Math.abs(median))) - 1);
    const shown = (x: number) => (x === 0 ? 0 : x).toLocaleString("en-US", { maximumFractionDigits: 6 });
    let threshold = Math.round(median / step) * step;
    while (col.policy.private && (held.has(String(threshold)) || held.has(shown(threshold)))) threshold += step;
    const at = shown(threshold);
    out.push(`Rows where ${n} is over ${at}`, `Rows where ${n} is under ${at}`, `Rows where ${n} is over 1,000`, `Rank each row by ${n}, largest first`);
    for (const c of cats) {
      out.push(
        `Total ${n} per ${c}`,
        `Average ${n} per ${c}, one row per ${c.toLowerCase()}`,
        `Median ${n} per ${c}`,
        `Max ${n} per ${c}`,
        `Top five ${c} values by total ${n}`,
      );
    }
    for (const d of dates) {
      out.push(
        `Total ${n} in March 2025 by ${d}`,
        `Total ${n} between January 1, 2025 and March 31, 2025 by ${d}`,
        `Total ${n} in Q1 2025 by ${d}`,
        `${n} per month of ${d}`,
      );
    }
  }
  for (const col of catCols) {
    const c = col.header;
    out.push(`Count of rows per ${c}`, `How many distinct ${c} values are there?`);
    // Values of the column, which go out as they are anyway.
    const values = [...new Set(data.text.slice(1).map((r) => (r[idx(col)] ?? "").trim()))]
      .filter((v) => v.length >= 4 && !privateTexts.has(v.toLowerCase()))
      .slice(0, 4);
    for (const v of values) {
      out.push(`Count of rows where ${c} is ${v}`);
      for (const n of nums.slice(0, 2)) out.push(`Total ${n.header} for ${v}`);
    }
  }
  for (const d of dates) out.push(`Rows in the last 30 days of ${d}`, `The earliest and latest ${d}`, `Count of rows per year of ${d}`, `Rows with ${d} in 2026`);
  for (const b of bools) out.push(`Count rows where ${b} is TRUE`, `Share of rows where ${b} is FALSE`);
  return [...new Set(out)];
}

interface Outcome {
  question: string;
  /**
   * What blocked it and where the first match of each hit is: "A exact in the question", "A word in
   * column C of sample row 1", "B structural". Empty when it passed.
   */
  hits: string[];
  /** The question as it would go out, when that isn't the question as typed. */
  sent: string | null;
}

export interface WorkbookResult {
  key: string;
  questions: number;
  privateColumns: string[];
  structureOnly: { blocked: Outcome[]; rewritten: Outcome[] };
  substituted: { blocked: Outcome[]; rewritten: Outcome[] };
}

type Described = { columns?: unknown; sheet?: unknown } | null | undefined;

/** A path as key names and indexes ("rows[2][1]"). The audit's paths hold no text, only these. */
function pathText(path: readonly (string | number)[]): string {
  return path.map((p, i) => (typeof p === "number" ? `[${p}]` : i === 0 ? p : `.${p}`)).join("");
}

/**
 * Where a path inside one described range points. A sample row has a cell for each column that isn't
 * left out, in order, so rows[r][c] is the c-th column whose treatment isn't "exclude", not
 * columns[c].
 */
function placeInRange(block: Described, path: readonly (string | number)[]): string {
  const columns: unknown[] = Array.isArray(block?.columns) ? block.columns : [];
  const sent = columns.filter((c) => (c as { treatment?: unknown } | null)?.treatment !== "exclude");
  const letter = (list: unknown[], i: unknown) => {
    const l = typeof i === "number" ? (list[i] as { letter?: unknown } | undefined)?.letter : undefined;
    return typeof l === "string" ? l : `#${String(i)}`;
  };
  const [key, i, part] = path;
  if (key === "rows" && typeof i === "number") return `in column ${letter(sent, part)} of sample row ${i + 1}`;
  if (key === "columns" && part === "header") return `in the header of column ${letter(columns, i)}`;
  if (key === "columns" && part === "note") return `in the note on column ${letter(columns, i)}`;
  if (key === "sheet") return "in the sheet name";
  if (key === "table") return "in the table name";
  return `at ${pathText(path)}`;
}

/**
 * Where a location is, in words, from its field path and the messages as built: column letters, row
 * numbers and key names only, never the matched text, so the report can be printed.
 */
export function placeOf(loc: AuditLocation | undefined, messages: readonly { role: string; content: string }[]): string {
  const field = loc?.field;
  if (!field) return "(place unknown)";
  const turn = `message ${field.message} (${field.role})`;
  // Only the current question is a description the battery reads; it clears history between questions.
  if (field.role !== "user" || field.message !== messages.length - 1) {
    return field.path.length > 0 ? `in ${turn} at ${pathText(field.path)}` : `in ${turn}`;
  }
  let d: { context_ranges?: unknown } & Described;
  try {
    d = JSON.parse(messages[field.message]!.content) as typeof d;
  } catch {
    return `in ${turn}`;
  }
  const [first, k, ...rest] = field.path;
  if (first === "task") return "in the question";
  if (first === "allowed_ranges") return "in the allowed ranges";
  if (first === "context_ranges" && typeof k === "number") {
    const block = Array.isArray(d.context_ranges) ? (d.context_ranges[k] as Described) : undefined;
    return `${placeInRange(block, rest)} of lookup range ${k + 1}`;
  }
  return placeInRange(d, field.path);
}

/** What blocked a prepared request, each text hit with where it matched (see Outcome.hits). */
export function blockedBy(prepared: Prepared): string[] {
  const { outcome, built } = prepared;
  if (outcome.ok) return [];
  const where = (h: { column: string; variant: string }) =>
    placeOf(outcome.locations.find((l) => l.column === h.column && l.variant === h.variant), built.messages);
  return [
    ...outcome.result.hits.map((h) => `${h.column} ${h.variant} ${where(h)}`),
    ...(outcome.result.structural ?? []).map((s) => `${s.column} structural`),
  ];
}

function benchPrompts(key: string): string[] {
  return Object.entries(PROMPTS)
    .filter(([id]) => id.startsWith(`${key}-`))
    .map(([, p]) => p);
}

/**
 * Runs the battery on one selection (header row included), with the columns Nymform suggests. A
 * question counts as rewritten when it would go out different from how it was typed, or with a part
 * mapped or waiting for a pick: an ordinary question should go out as typed.
 */
export async function runBattery(key: string, data: RangeData, prompts: readonly string[], alsoPrivate: readonly string[] = []): Promise<WorkbookResult> {
  const { flow } = await flowFor(key, data);
  for (const letter of alsoPrivate) flow.setPolicy(0, letter, { private: true });
  const columns = flow.getState().ranges[0]!.columns;
  const questions = questionsFor(prompts, columns, data);
  const run = (mode: ProdMode) => {
    const blocked: Outcome[] = [];
    const rewritten: Outcome[] = [];
    for (const q of questions) {
      const r = flow.prepare(q, mode);
      if (!r.ok) throw new Error(`${key} "${q}": ${r.message}`);
      const { outcome, sentQuestion, partial } = r.prepared;
      const hits = blockedBy(r.prepared);
      const changed = sentQuestion !== q || partial.auto.length > 0 || partial.pending.length > 0;
      if (!outcome.ok) blocked.push({ question: q, hits, sent: sentQuestion === q ? null : sentQuestion });
      if (changed) rewritten.push({ question: q, hits, sent: sentQuestion });
      flow.clearHistory();
    }
    return { blocked, rewritten };
  };
  return {
    key,
    questions: questions.length,
    privateColumns: columns.filter((c) => c.policy.private).map((c) => `${c.letter} ${c.header} (${c.type})`),
    structureOnly: run("structure_only"),
    substituted: run("substituted"),
  };
}

const pct = (n: number, d: number) => `${((100 * n) / Math.max(1, d)).toFixed(1)}%`;

export function report(results: WorkbookResult[], title = "Regression battery: default suggestions"): string {
  const lines = [title, ""];
  lines.push(
    "| Workbook | Questions | Structure only: blocked | rewritten | Sample rows: blocked | rewritten | Sample-rows block rate |",
    "|---|---|---|---|---|---|---|",
  );
  for (const r of results) {
    const so = r.structureOnly;
    const su = r.substituted;
    lines.push(`| ${r.key} | ${r.questions} | ${so.blocked.length} | ${so.rewritten.length} | ${su.blocked.length} | ${su.rewritten.length} | ${pct(su.blocked.length, r.questions)} |`);
  }
  lines.push(
    "",
    "Sample rows are the first rows of the selection whatever the question, so a block that comes from a row",
    "blocks every question on that workbook: the rate is close to 0% or 100%. What matters is what matched.",
  );
  for (const r of results) {
    lines.push("", `${r.key}: private by default: ${r.privateColumns.join(", ") || "none"}`);
    for (const [label, part] of [["structure only", r.structureOnly], ["sample rows", r.substituted]] as const) {
      for (const b of part.blocked) lines.push(`- ${label}, blocked: ${JSON.stringify(b.question)}: ${b.hits.join(", ")}`);
      for (const w of part.rewritten) lines.push(`- ${label}, rewritten: ${JSON.stringify(w.question)} goes out as ${JSON.stringify(w.sent)}`);
    }
    const rate = r.substituted.blocked.length / Math.max(1, r.questions);
    if (rate > SUBSTITUTED_LIMIT) lines.push(`ABOVE LIMIT: ${r.key} blocks ${pct(r.substituted.blocked.length, r.questions)} with sample rows (limit ${SUBSTITUTED_LIMIT * 100}%). Stop for Jorge's decision; don't change matching rules in this round.`);
  }
  return lines.join("\n");
}

/**
 * A sheet with private dates, which no bench workbook has under the default suggestions: Date of
 * birth (suggested private) and Last contact (marked private here), with contacts all through
 * 2025 and 2026 except the two days the questions name. Date phrases in questions ("in March 2025",
 * "Q1 2025", "in 2026") then meet private dates of the same months and years, and must not match.
 */
function signups(): RangeData {
  const first = ["Maria", "Kenji", "Priya", "Omar", "Lena", "Tomas", "Ines", "Hugo", "Ana", "Felix"];
  const last = ["Lopez", "Watanabe", "Natarajan", "Haddad", "Fischer", "Novak", "Duarte", "Laurent", "Ruiz", "Bianchi"];
  const states = ["CA", "NY", "TX", "WA", "IL"];
  const header = ["Customer", "Date of birth", "Last contact", "State", "Orders"];
  const rows: (string | number)[][] = [];
  for (let i = 0; i < 60; i++) {
    const dob = 18264 + ((i * 3571) % 18000); // 1950 to 1999
    let signup = 45658 + ((i * 11) % 700); // 2025 and 2026
    if (signup === 45658 || signup === 45747) signup++; // not January 1 or March 31, 2025
    rows.push([`${first[i % 10]} ${last[(i * 7) % 10]}`, dob, signup, states[i % 5]!, 1 + ((i * 5) % 23)]);
  }
  const shown = (serial: number) => {
    const d = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
    return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`;
  };
  return {
    sheet: "Signups",
    address: "A1:E61",
    rowIndex: 0,
    columnIndex: 0,
    values: [header, ...rows],
    text: [header, ...rows.map((r) => [String(r[0]), shown(r[1] as number), shown(r[2] as number), String(r[3]), String(r[4])])],
    formulas: [header, ...rows],
    valueTypes: [header.map(() => "String"), ...rows.map((r) => r.map((v) => (typeof v === "number" ? "Double" : "String")))],
    numberFormat: [header.map(() => "General"), ...rows.map(() => ["General", "m/d/yyyy", "m/d/yyyy", "General", "General"])],
    table: null,
  };
}

/**
 * A sheet where a column sent as it is holds private values ("Referred by", names from Customer)
 * and the column before it is left out: every request with sample rows blocks on a row cell of C,
 * which is the second cell of each sent row.
 */
function referrals(): RangeData {
  const names = ["Maria Lopez", "Kenji Watanabe", "Priya Natarajan", "Omar Haddad", "Lena Fischer", "Tomas Novak"];
  const header = ["Customer", "Notes", "Referred by", "State"];
  const rows = names.map((n, i) => [n, `call back ${i + 1}`, names[(i + 1) % names.length]!, ["CA", "NY", "TX"][i % 3]!]);
  return {
    sheet: "Referrals",
    address: "A1:D7",
    rowIndex: 0,
    columnIndex: 0,
    values: [header, ...rows],
    text: [header, ...rows],
    formulas: [header, ...rows],
    valueTypes: [header.map(() => "String"), ...rows.map((r) => r.map(() => "String"))],
    numberFormat: [header.map(() => "General"), ...rows.map((r) => r.map(() => "General"))],
    table: null,
  };
}

describe("battery report: where a block matched", () => {
  const loc = (field: AuditLocation["field"]): AuditLocation => ({ column: "A", variant: "exact", start: 0, end: 1, origin: "value", field });
  const description = {
    columns: [
      { letter: "A", treatment: "stand_in" },
      { letter: "B", treatment: "exclude" },
      { letter: "C", treatment: "as_is" },
    ],
    context_ranges: [{ sheet: "Regions", columns: [{ letter: "A", treatment: "exclude" }, { letter: "B", treatment: "as_is" }] }],
    task: "",
  };
  const messages = [
    { role: "system", content: "" },
    { role: "user", content: JSON.stringify(description) },
  ];
  const user = (path: (string | number)[]) => loc({ message: 1, role: "user", path, within: "string" });

  it("reads a sample-row cell through the columns that are sent, not through columns[]", () => {
    expect(placeOf(user(["rows", 0, 1]), messages)).toBe("in column C of sample row 1");
    expect(placeOf(user(["rows", 4, 0]), messages)).toBe("in column A of sample row 5");
    expect(placeOf(user(["context_ranges", 0, "rows", 2, 0]), messages)).toBe("in column B of sample row 3 of lookup range 1");
  });

  it("names headers, notes, the question and other fields by letter or key only", () => {
    expect(placeOf(user(["columns", 1, "header"]), messages)).toBe("in the header of column B");
    expect(placeOf(user(["columns", 2, "note"]), messages)).toBe("in the note on column C");
    expect(placeOf(user(["context_ranges", 0, "columns", 0, "header"]), messages)).toBe("in the header of column A of lookup range 1");
    expect(placeOf(user(["task"]), messages)).toBe("in the question");
    expect(placeOf(user(["sheet"]), messages)).toBe("in the sheet name");
    expect(placeOf(user(["range", "address"]), messages)).toBe("at range.address");
    expect(placeOf(loc({ message: 0, role: "system", path: [], within: "text" }), messages)).toBe("in message 0 (system)");
    expect(placeOf(loc(null), messages)).toBe("(place unknown)");
    expect(placeOf(undefined, messages)).toBe("(place unknown)");
  });

  it("a real block from a sample row names the sent column, with an excluded column before it", async () => {
    const data = referrals();
    const customers = data.text.slice(1).map((row) => row[0]!);
    const { flow } = await flowFor("referrals", data);
    flow.setPolicy(0, "A", { private: true, treatment: "stand_in" });
    flow.setPolicy(0, "B", { private: false, treatment: "exclude" });
    flow.setPolicy(0, "C", { private: false, treatment: "as_is" });
    flow.setPolicy(0, "D", { private: false, treatment: "as_is" });
    const r = flow.prepare("Count of rows per State", "substituted");
    if (!r.ok) throw new Error(r.message);
    const hits = blockedBy(r.prepared);
    expect(hits.length).toBeGreaterThan(0);
    for (const hit of hits) {
      // Column C (the second cell of a sent row), never B (columns[1]).
      const m = /^A (exact|word) in column C of sample row (\d+)$/u.exec(hit);
      expect(m, hit).not.toBeNull();
      // And that cell does hold a customer's name.
      expect(customers).toContain(data.text[Number(m![2])]![2]);
    }
    // Where, never what: no name from the sheet is in the report line.
    for (const name of customers) expect(hits.join(" ")).not.toContain(name.split(" ")[0]);
    const passed = flow.prepare("Count of rows per State", "structure_only");
    expect(passed.ok && blockedBy(passed.prepared)).toEqual([]);
  });
});

describe("regression battery (default suggestions)", async () => {
  const results = await Promise.all([
    ...DATA.workbooks.map((wb: BenchWorkbook) => runBattery(wb.key, toRangeData(wb), benchPrompts(wb.key))),
    runBattery("signups (private dates)", signups(), ["Customers last contacted in 2026", "How many customers were last contacted in March 2025?"], ["C"]),
  ]);

  it.runIf(REPORT)("prints the report", () => {
    process.stdout.write(`\n${report(results)}\n\n`);
  });

  for (const r of results) {
    it(`${r.key}: every ordinary question passes with structure only, as typed`, () => {
      expect(r.questions).toBeGreaterThanOrEqual(20);
      expect(r.structureOnly.blocked, report([r])).toEqual([]);
      expect(r.structureOnly.rewritten, report([r])).toEqual([]);
    });
    it(`${r.key}: with sample rows, no ordinary question is rewritten (blocks are counted below)`, () => {
      expect(r.substituted.rewritten, report([r])).toEqual([]);
    });
  }

  it("the private-date sheet has private dates in the months and years its questions name", () => {
    const r = results.at(-1)!;
    expect(r.privateColumns).toEqual(expect.arrayContaining(["B Date of birth (date)", "C Last contact (date)"]));
    expect(r.questions).toBeGreaterThanOrEqual(20);
  });

  it("pins the counts per workbook, so a change in what blocks shows in review", () => {
    expect(
      Object.fromEntries(
        results.map((r) => [r.key, { questions: r.questions, sampleRowsBlocked: r.substituted.blocked.length, privateByDefault: r.privateColumns }]),
      ),
    ).toMatchSnapshot();
  });
});
