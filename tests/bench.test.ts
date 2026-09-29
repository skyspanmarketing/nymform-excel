// Benchmark tooling (spec §11): raw sender guard (T-P3), generator, task files, compare, leak check,
// runner. The workbooks and task files under bench/ are committed; these tests also confirm they
// are what the generator produces today.
import { mkdtempSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import { afterEach, describe, expect, it, vi } from "vitest";
import { CATEGORIES, compare, validateTask, valuesMatch, type BenchTask, type Check } from "../bench/compare";
import {
  BENCH_MARKER as GEN_MARKER,
  BENCH_SHEET,
  CREATOR,
  displayText,
  formatJson,
  generateData,
  partsOf,
  privateValuesFile,
  privateValuesOf,
  serialOf,
  toRangeData,
  type BenchWorkbook,
} from "../bench/generate";
import { extractEntries, isBinRendering, leakCheck, runCli } from "../bench/leakcheck";
import {
  BENCH_MARKER,
  NOT_A_BENCH_WORKBOOK,
  RAW_MODE,
  buildRawBody,
  sendRaw,
  type RawSendInput,
} from "../bench/rawSender";
import {
  CONDITIONS,
  HEADLINE_DEFINITION,
  qualifyFormula,
  resultsFileName,
  runBench,
  runTask,
  summarize,
  toJsonl,
  toMarkdown,
  type Condition,
  type Prepared,
  type RunRecord,
  type RunnerDeps,
} from "../bench/runner";
import { audit, buildAuditContext } from "../src/core/auditor";
import { RequestLog } from "../src/core/log";
import { buildStructureOnly, buildSubstituted } from "../src/core/payload";
import { SYSTEM_PROMPT } from "../src/core/prompt";
import { buildSheetContext, inferSchema } from "../src/core/schema";
import { defaultTreatment } from "../src/core/suggest";
import { createStandInMap, substituteUserText } from "../src/core/transform";
import type { CellValue, ColumnPolicy, ModelReply, ProdMode, RangeSpec } from "../src/core/types";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const BENCH_DIR = join(ROOT, "bench");
const DATA = generateData();
const wb = (key: string): BenchWorkbook => DATA.workbooks.find((w) => w.key === key)!;

afterEach(() => {
  vi.unstubAllGlobals();
});

// ---------------------------------------------------------------------------------------------

describe("raw sender (bench only)", () => {
  const base = (marker: string | null): RawSendInput => ({
    marker,
    data: toRangeData(wb("orders")),
    question: "Total Amount per Region, one row per region",
    model: "openai/gpt-6-luna",
    key: "sk-test",
    endpoint: "https://openrouter.ai/api/v1",
  });

  it("T-P3: sendRaw on a workbook without the bench marker throws and never calls fetch", async () => {
    const globalFetch = vi.fn();
    vi.stubGlobal("fetch", globalFetch);
    for (const marker of [null, "", "NYMFORM_BENCH_V2", "nymform_bench_v1", " NYMFORM_BENCH_V1"]) {
      const fetchImpl = vi.fn();
      await expect(sendRaw({ ...base(marker), fetchImpl: fetchImpl as unknown as typeof fetch })).rejects.toThrow(
        NOT_A_BENCH_WORKBOOK,
      );
      await expect(sendRaw(base(marker))).rejects.toThrow(
        "Not a bench workbook: the raw baseline runs only on workbooks with the bench marker.",
      );
      expect(fetchImpl).not.toHaveBeenCalled();
    }
    expect(() => buildRawBody(base(null))).toThrow(NOT_A_BENCH_WORKBOOK);
    expect(globalFetch).not.toHaveBeenCalled();
  });

  it("uses the same marker string as the generator and carries the raw_bench mode literal", () => {
    expect(BENCH_MARKER).toBe(GEN_MARKER);
    expect(RAW_MODE).toBe("raw_bench");
    expect(readFileSync(join(BENCH_DIR, "rawSender.ts"), "utf8")).toContain('"raw_bench"');
  });

  it("with the marker, sends every value in production request shape to the endpoint", async () => {
    const reply = JSON.stringify({ choices: [{ message: { content: "{}" } }], usage: { prompt_tokens: 900, completion_tokens: 40 } });
    const fetchImpl = vi.fn(async () => new Response(reply, { status: 200, headers: { "X-Provider-Name": "Azure" } }));
    let t = 1000;
    const res = await sendRaw({ ...base(BENCH_MARKER), fetchImpl: fetchImpl as unknown as typeof fetch, now: () => (t += 250) });
    expect(fetchImpl).toHaveBeenCalledTimes(1);
    const [url, init] = fetchImpl.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe("https://openrouter.ai/api/v1/chat/completions");
    expect(init.method).toBe("POST");
    expect((init.headers as Record<string, string>).Authorization).toBe("Bearer sk-test");
    expect(init.body).toBe(res.body);
    expect(res).toMatchObject({ raw: reply, status: 200, latencyMs: 250, mode: "raw_bench", tokens: { prompt: 900, completion: 40 }, providerName: "Azure" });
    expect(res.bytes).toBe(new TextEncoder().encode(res.body).length);

    const body = JSON.parse(res.body) as { model: string; messages: { role: string; content: string }[]; provider?: unknown };
    expect(Object.keys(body)[0]).toBe("model");
    expect(Object.keys(body)[1]).toBe("messages");
    expect(Object.keys(body).at(-1)).toBe("provider");
    expect(body.provider).toEqual({ zdr: true });
    expect(body.messages[0]).toEqual({ role: "system", content: SYSTEM_PROMPT });
    const user = JSON.parse(body.messages[1]!.content) as { mode: string; rows: CellValue[][]; rows_sent: number; rows_total: number; task: string };
    expect(user.mode).toBe("raw_bench");
    expect(user.rows_sent).toBe(500);
    expect(user.rows_total).toBe(500);
    expect(user.rows).toHaveLength(500);
    expect(user.task).toBe("Total Amount per Region, one row per region");
    // Every private value (display text) of the workbook is in the raw body; that is the baseline.
    const orders = wb("orders");
    for (const c of orders.canaries) expect(res.body).toContain(c);
    expect(user.rows[0]).toEqual([
      orders.rows[0]![0],
      orders.rows[0]![1],
      orders.rows[0]![2],
      displayText(orders.rows[0]![3]!, "m/d/yyyy"),
      orders.rows[0]![4],
      orders.rows[0]![5],
    ]);
  });

  it("leaves out the provider field for endpoints other than OpenRouter, and honors a row cap", () => {
    const built = buildRawBody({ ...base(BENCH_MARKER), endpoint: "https://api.example.com/v1", rowCap: 10 });
    const body = JSON.parse(built.body) as Record<string, unknown>;
    expect(body.provider).toBeUndefined();
    expect(built.rowsSent).toBe(10);
    expect(built.rowsTotal).toBe(500);
  });

  it("maps a network failure to a plain message without the key", async () => {
    const fetchImpl = vi.fn(async () => {
      throw new TypeError("network down");
    });
    const err = await sendRaw({ ...base(BENCH_MARKER), fetchImpl: fetchImpl as unknown as typeof fetch }).catch((e: Error) => e);
    expect(err).toBeInstanceOf(Error);
    expect((err as Error).message).toBe("Couldn't reach openrouter.ai.");
    expect((err as Error).message).not.toContain("sk-test");
  });
});

// ---------------------------------------------------------------------------------------------

describe("generator", () => {
  it("returns identical data on every call with the same seed", () => {
    const again = generateData();
    expect(again).toEqual(DATA);
    expect(JSON.stringify(again)).toBe(JSON.stringify(DATA));
    expect(generateData(7).workbooks[0]!.rows).not.toEqual(DATA.workbooks[0]!.rows);
  });

  it("builds the four workbooks from spec §11 with those columns and row counts", () => {
    const shape = DATA.workbooks.map((w) => ({ file: w.file, sheet: w.sheet, range: w.range, rows: w.rows.length, headers: w.columns.map((c) => c.header), private: w.privateLetters }));
    expect(shape).toEqual([
      { file: "orders.xlsx", sheet: "Orders", range: "A1:F501", rows: 500, headers: ["Customer", "Email", "Region", "Order date", "Product", "Amount"], private: ["A", "B"] },
      { file: "roster.xlsx", sheet: "Roster", range: "A1:F121", rows: 120, headers: ["Name", "Employee ID", "Department", "Rank", "Start date", "Salary"], private: ["A", "B", "F"] },
      { file: "sections.xlsx", sheet: "Sections", range: "A1:H201", rows: 200, headers: ["Course", "Section", "Instructor", "Days", "Start time", "Units", "Enrolled", "Cap"], private: ["C"] },
      { file: "appointments.xlsx", sheet: "Appointments", range: "A1:F301", rows: 300, headers: ["Client name", "Phone", "Location", "Appointment date", "Duration (min)", "No-show"], private: ["A", "B"] },
    ]);
  });

  it("uses only synthetic contact data: example.com / example.org emails and 555-01xx phones", () => {
    for (const r of wb("orders").rows) expect(r[1]).toMatch(/^[a-z0-9.]+@example\.(?:com|org)$/);
    for (const r of wb("appointments").rows) expect(r[1]).toMatch(/^\(\d{3}\) 555-01\d{2}$/);
    for (const r of wb("appointments").rows) expect(typeof r[5]).toBe("boolean");
  });

  it("plants 10 distinct canaries of 6+ characters once each in private columns", () => {
    const all = DATA.workbooks.flatMap((w) => w.canaries);
    expect(new Set(all).size).toBe(40);
    for (const w of DATA.workbooks) {
      expect(w.canaries).toHaveLength(10);
      const privateCols = w.privateLetters.map((l) => l.charCodeAt(0) - 65);
      for (const c of w.canaries) {
        expect(c.length).toBeGreaterThanOrEqual(6);
        const cells = w.rows.flatMap((r) => r.map((v, i) => ({ v, i }))).filter((x) => x.v === c);
        expect(cells).toHaveLength(1);
        expect(privateCols).toContain(cells[0]!.i);
      }
    }
  });

  it("formats display text as Excel shows it", () => {
    expect(displayText(serialOf(2025, 3, 14), "m/d/yyyy")).toBe("3/14/2025");
    expect(displayText(570 / 1440, "h:mm AM/PM")).toBe("9:30 AM");
    expect(displayText(720 / 1440, "h:mm AM/PM")).toBe("12:00 PM");
    expect(displayText(1234.5, "#,##0.00")).toBe("1,234.50");
    expect(displayText(84500, "#,##0")).toBe("84,500");
    expect(displayText(true, "General")).toBe("TRUE");
    expect(serialOf(2025, 1, 1)).toBe(45658);
    expect(partsOf(45658)).toEqual({ y: 2025, m: 1, d: 1, weekday: 3 });
  });

  it("lists every display text and raw value of the private cells", () => {
    const roster = wb("roster");
    const values = privateValuesOf(roster);
    const salary = roster.rows[0]![5] as number;
    expect(values).toContain(String(salary));
    expect(values).toContain(displayText(salary, "#,##0"));
    expect(values).toContain(roster.rows[0]![0]);
    expect(values).toContain(roster.rows[0]![1]);
    expect(values).not.toContain(roster.rows[0]![2]); // Department is not private
  });
});

// ---------------------------------------------------------------------------------------------

function readTaskFiles(): BenchTask[] {
  const dir = join(BENCH_DIR, "tasks");
  return readdirSync(dir)
    .filter((f) => f.endsWith(".json"))
    .sort()
    .map((f) => JSON.parse(readFileSync(join(dir, f), "utf8")) as BenchTask);
}

describe("task files", () => {
  const files = readTaskFiles();

  it("are 20 valid tasks in the spec §11 shape, spread across the seven categories", () => {
    expect(files).toHaveLength(20);
    for (const t of files) expect(validateTask(t), t.id).toEqual([]);
    const counts = Object.fromEntries(CATEGORIES.map((c) => [c, files.filter((t) => t.category === c).length]));
    expect(counts).toEqual({ conditional_totals: 3, counting: 3, lookup: 3, dates: 3, text: 3, ranking: 3, row_anomalies: 2 });
    const names = readdirSync(join(BENCH_DIR, "tasks")).filter((f) => f.endsWith(".json"));
    for (const t of files) expect(names).toContain(`${t.id}.json`);
  });

  it("match what the generator computes from the seeded data (committed files are current)", () => {
    const byId = new Map(DATA.tasks.map((t) => [t.id, t]));
    for (const t of files) expect(t, t.id).toEqual(byId.get(t.id));
    for (const t of DATA.tasks) {
      expect(readFileSync(join(BENCH_DIR, "tasks", `${t.id}.json`), "utf8")).toBe(formatJson(t) + "\n");
    }
    const pv = JSON.parse(readFileSync(join(BENCH_DIR, "workbooks", "private-values.json"), "utf8")) as unknown;
    expect(pv).toEqual(privateValuesFile(DATA));
  });

  it("point at their workbook's sheet, range and private columns", () => {
    for (const t of files) {
      const w = DATA.workbooks.find((x) => x.file === t.workbook)!;
      expect(t.sheet).toBe(w.sheet);
      expect(t.range).toBe(w.range);
      expect(t.private).toEqual(w.privateLetters);
      if (t.check.type === "column") expect(t.check.expected).toHaveLength(w.rows.length);
    }
  });

  it("have prompts with no private value or canary in them", () => {
    for (const t of files) {
      const w = DATA.workbooks.find((x) => x.file === t.workbook)!;
      const prompt = t.prompt.toLowerCase();
      for (const v of [...privateValuesOf(w), ...w.canaries].filter((v) => v.length >= 4)) {
        expect(prompt.includes(v.toLowerCase()), `${t.id} contains a private value`).toBe(false);
      }
    }
  });

  it("have expected values that agree with an independent recount of the data", () => {
    const task = (id: string) => DATA.tasks.find((t) => t.id === id)!;
    const orders = wb("orders").rows;
    const regionTotals = new Map<string, number>();
    for (const r of orders) regionTotals.set(r[2] as string, (regionTotals.get(r[2] as string) ?? 0) + (r[5] as number));
    const t3 = task("orders-03").check as Extract<Check, { type: "table" }>;
    expect(t3.expected.map((r) => r[0]).sort()).toEqual([...regionTotals.keys()].sort());
    for (const [region, total] of t3.expected) expect(total).toBeCloseTo(regionTotals.get(region as string)!, 2);

    const maxAmount = Math.max(...orders.map((r) => r[5] as number));
    expect(task("orders-01").check.expected).toBe(orders.find((r) => r[5] === maxAmount)![4]);

    const sections = wb("sections").rows;
    const maxEnrolled = Math.max(...sections.map((r) => r[6] as number));
    expect(task("sections-02").check.expected).toBe(sections.find((r) => r[6] === maxEnrolled)![2]);

    const weekend = wb("appointments").rows.filter((r) => [0, 6].includes(new Date(((r[3] as number) - 25569) * 86400000).getUTCDay()));
    expect(task("appointments-03").check.expected).toBe(weekend.length);

    const flags = (task("orders-07").check as Extract<Check, { type: "column" }>).expected;
    expect(flags.filter((f) => f === true)).toHaveLength(6);
    const doubles = (task("appointments-04").check as Extract<Check, { type: "column" }>).expected;
    expect(doubles.filter((f) => f === true).length).toBeGreaterThanOrEqual(8);
  });

  it("pass compare() when Excel returns exactly the expected values", () => {
    for (const t of files) {
      const c = t.check;
      const actual: CellValue[][] =
        c.type === "cell" ? [[c.expected]] : c.type === "column" ? c.expected.map((v) => [v]) : [...c.expected].reverse().map((r) => [...r]);
      expect(compare(actual, c), t.id).toEqual({ pass: true, detail: expect.any(String) });
    }
  });
});

// ---------------------------------------------------------------------------------------------

describe("committed workbooks", () => {
  it("hold the generated data, formats, metadata and the hidden bench sheet", async () => {
    for (const w of DATA.workbooks) {
      const book = new ExcelJS.Workbook();
      await book.xlsx.readFile(join(BENCH_DIR, "workbooks", w.file));
      expect(book.creator).toBe(CREATOR);
      expect(book.worksheets.map((s) => s.name)).toEqual([w.sheet, BENCH_SHEET]);
      const data = book.getWorksheet(w.sheet)!;
      const marker = book.getWorksheet(BENCH_SHEET)!;
      expect(data.state).toBe("visible");
      expect(marker.state).toBe("hidden");
      expect(marker.getCell("A1").value).toBe("NYMFORM_BENCH_V1");
      expect(marker.getRow(2).values).toEqual([undefined, "private", ...w.privateLetters]);
      expect(marker.getRow(3).values).toEqual([undefined, "canaries", ...w.canaries]);
      expect(data.getRow(1).values).toEqual([undefined, ...w.columns.map((c) => c.header)]);
      expect(data.rowCount).toBe(w.rows.length + 1);
      for (const r of [0, 1, Math.floor(w.rows.length / 2), w.rows.length - 1]) {
        w.columns.forEach((col, c) => {
          const cell = data.getRow(r + 2).getCell(c + 1);
          const expected = w.rows[r]![c]!;
          const got = cell.value instanceof Date ? cell.value.getTime() / 86400000 + 25569 : cell.value;
          const where = `${w.file} row ${r + 2} col ${c + 1}`;
          if (typeof expected === "number") expect(got, where).toBeCloseTo(expected, 6);
          else expect(got, where).toBe(expected);
          if (col.numFmt !== "General") expect(cell.numFmt).toBe(col.numFmt);
        });
      }
    }
  });
});

// ---------------------------------------------------------------------------------------------

describe("compare()", () => {
  it("matches numbers within tolerance, strings trimmed and case-insensitive, booleans exactly", () => {
    expect(compare([[100.004]], { type: "cell", expected: 100, tolerance: 0.01 }).pass).toBe(true);
    expect(compare([[100.02]], { type: "cell", expected: 100, tolerance: 0.01 }).pass).toBe(false);
    expect(compare([[7]], { type: "cell", expected: 7 }).pass).toBe(true);
    expect(compare([["7"]], { type: "cell", expected: 7 }).pass).toBe(true);
    expect(compare([["1,234"]], { type: "cell", expected: 1234 }).pass).toBe(false);
    expect(compare([["  standing DESK "]], { type: "cell", expected: "Standing Desk" }).pass).toBe(true);
    expect(compare([["Standing Desks"]], { type: "cell", expected: "Standing Desk" }).pass).toBe(false);
    expect(compare([[true]], { type: "cell", expected: true }).pass).toBe(true);
    expect(compare([["TRUE"]], { type: "cell", expected: true }).pass).toBe(true);
    expect(compare([[1]], { type: "cell", expected: true }).pass).toBe(false);
    expect(valuesMatch("", null)).toBe(true);
    expect(valuesMatch(null, "")).toBe(true);
  });

  it("wants one value for a cell check", () => {
    const r = compare([[5, 6]], { type: "cell", expected: 5 });
    expect(r.pass).toBe(false);
    expect(r.detail).toContain("1 row x 2 columns");
    expect(compare([[5], [null], [""]], { type: "cell", expected: 5 }).pass).toBe(true);
    expect(compare([], { type: "cell", expected: 5 }).detail).toBe("Nothing was returned.");
    expect(compare([["#N/A"]], { type: "cell", expected: 5 }).detail).toContain("Excel error");
  });

  it("compares a column one value per data row, in order", () => {
    const check: Check = { type: "column", expected: ["a.com", "b.org", "a.com"] };
    expect(compare([["A.COM"], ["b.org"], ["a.com "]], check).pass).toBe(true);
    const wrongOrder = compare([["b.org"], ["a.com"], ["a.com"]], check);
    expect(wrongOrder.pass).toBe(false);
    expect(wrongOrder.detail).toBe("2 of 3 values don't match; the first is data row 1.");
    expect(compare([["a.com"], ["b.org"]], check).detail).toBe("Expected 3 values (one per data row), got 2.");
    expect(compare([["a.com", 1], ["b.org", 1], ["a.com", 1]], check).pass).toBe(false);
    expect(compare([["Domain"], ["a.com"], ["b.org"], ["a.com"]], check).pass).toBe(true);
    expect(compare([["x"], ["#VALUE!"], ["a.com"]], check).detail).toContain("1 of them show an Excel error");
  });

  it("compares a table order-insensitive, ignoring a header row and a total row", () => {
    const check: Check = { type: "table", expected: [["East", 10.5], ["West", 20]], tolerance: 0.01 };
    expect(compare([["west", 20.001], ["East", 10.5]], check).pass).toBe(true);
    expect(compare([["Region", "Amount"], ["West", 20], ["East", 10.5], ["Total", 30.5]], check).pass).toBe(true);
    expect(compare([["West", 20], ["East", 10.5], ["North", 1]], check)).toEqual({ pass: false, detail: "2 of 2 expected rows found, 1 extra row." });
    expect(compare([["West", 20], ["East", 11]], check)).toEqual({ pass: false, detail: "1 of 2 expected rows found, 1 extra row." });
    expect(compare([["West"], ["East"]], check).detail).toBe("Expected 2 columns, got 1.");
    // A duplicate expected row needs two matching actual rows.
    const dup: Check = { type: "table", expected: [["A", 1], ["A", 1]] };
    expect(compare([["A", 1]], dup).pass).toBe(false);
    expect(compare([["A", 1], ["a", 1]], dup).pass).toBe(true);
  });

  it("never puts values in its details", () => {
    const secret = "Quorbel Vantrisk";
    const results = [
      compare([["Zanthe Pellowin"]], { type: "cell", expected: secret }),
      compare([["Zanthe Pellowin"]], { type: "column", expected: [secret] }),
      compare([["Zanthe Pellowin", 3]], { type: "table", expected: [[secret, 3]] }),
    ];
    for (const r of results) {
      expect(r.pass).toBe(false);
      expect(r.detail).not.toMatch(/Quorbel|Vantrisk|Zanthe|Pellowin/);
    }
  });

  it("validateTask rejects malformed tasks", () => {
    const good = DATA.tasks[0]!;
    expect(validateTask(good)).toEqual([]);
    expect(validateTask({ ...good, category: "misc" })).not.toEqual([]);
    expect(validateTask({ ...good, range: "Orders!A1:F501" })).not.toEqual([]);
    expect(validateTask({ ...good, check: { type: "table", expected: [[1], [1, 2]] } })).not.toEqual([]);
    expect(validateTask({ ...good, extra: 1 })).toEqual(['Unknown key "extra".']);
  });
});

// ---------------------------------------------------------------------------------------------

describe("leakCheck()", () => {
  const values = ["Maria Lopez", 'Ana "AJ" Ruiz', "(831) 555-0199", "Al", "E10234", "84,500", "84500"];
  const canaries = ["Quorbel Vantrisk"];
  const userTurn = (task: string, rows: unknown[][] = []) =>
    JSON.stringify({ nymform: "0.1", mode: "substituted", sheet: "Orders", rows, task });
  const bodyOf = (content: string) =>
    JSON.stringify({ model: "openai/gpt-6-luna", messages: [{ role: "system", content: SYSTEM_PROMPT }, { role: "user", content }] });

  it("finds a planted private value, its escaped forms, its digits, and a canary", () => {
    const entries = [
      { mode: "substituted", body: bodyOf(userTurn("total for MARIA LOPEZ")) },
      { mode: "substituted", body: bodyOf(userTurn("x", [['Ana "AJ" Ruiz', 1]])) },
      { mode: "structure_only", body: bodyOf(userTurn("call 8315550199")) },
      { mode: "structure_only", body: bodyOf(userTurn("hi quorbel vantrisk")) },
      { mode: "raw_bench", body: bodyOf(userTurn("x", [["Maria Lopez", "E10234"]])) },
    ];
    const report = leakCheck(entries, values, canaries);
    expect(report.leaks).toEqual([
      { entry: 0, mode: "substituted", kind: "value", index: 0, variant: "exact" },
      { entry: 1, mode: "substituted", kind: "value", index: 1, variant: "json2" },
      { entry: 2, mode: "structure_only", kind: "value", index: 2, variant: "digits" },
      { entry: 3, mode: "structure_only", kind: "canary", index: 0, variant: "canary" },
      { entry: 4, mode: "raw_bench", kind: "value", index: 0, variant: "exact" },
      { entry: 4, mode: "raw_bench", kind: "value", index: 4, variant: "exact" },
    ]);
    expect(report.byMode).toEqual({
      substituted: { requests: 2, leakingRequests: 2, leaks: 2 },
      structure_only: { requests: 2, leakingRequests: 2, leaks: 2 },
      raw_bench: { requests: 1, leakingRequests: 1, leaks: 2 },
    });
    expect(report.protectedLeaks).toBe(4);
    // The single-escaped form, as it would appear outside a nested JSON string.
    const once = leakCheck([{ mode: "substituted", body: JSON.stringify({ note: 'Ana "AJ" Ruiz' }) }], values, []);
    expect(once.leaks.map((l) => l.variant)).toEqual(["json"]);
  });

  it("reports zero on bodies that hold only stand-ins and renderings", () => {
    const rows = [
      ["PERSON_001", "EMAIL_001", "East", "3/14/2025", "Webcam", 120.5],
      ["NYMFORM_PERSON_002", "EMAIL_002", "West", "3/15/2025", "Keyboard", 80],
    ];
    const entries = [
      { mode: "substituted", body: bodyOf(userTurn("total for PERSON_001", rows)) },
      { mode: "structure_only", body: bodyOf(userTurn("Total Amount per Region")) },
      { mode: "substituted", body: bodyOf(JSON.stringify({ rows: [["80000-89999", "ID_001"]] })) },
    ];
    const report = leakCheck(entries, values, canaries);
    expect(report.leaks).toEqual([]);
    expect(report.protectedLeaks).toBe(0);
    expect(report.byMode.substituted).toEqual({ requests: 2, leakingRequests: 0, leaks: 0 });
  });

  it("skips values under 4 characters (the auditor's structural check covers those)", () => {
    const report = leakCheck([{ mode: "substituted", body: bodyOf(userTurn("Al and al")) }], values, []);
    expect(report.leaks).toEqual([]);
  });

  it("stays clean for the generator's own stand-in-only bodies across all bench workbooks", () => {
    const pv = privateValuesFile(DATA);
    for (const [key, entry] of Object.entries(pv)) {
      const w = wb(key);
      const rows = w.rows.slice(0, 50).map((r) => r.map((v, c) => (w.privateLetters.includes(String.fromCharCode(65 + c)) ? `TEXT_${String(c).padStart(3, "0")}` : displayText(v, w.columns[c]!.numFmt))));
      const report = leakCheck([{ mode: "substituted", body: bodyOf(userTurn("Total per group", rows)) }], entry.values, entry.canaries);
      expect(report.leaks, key).toEqual([]);
      // ...and the same rows with one real private cell put back are caught.
      const pc = w.privateLetters[0]!.charCodeAt(0) - 65;
      rows[3]![pc] = displayText(w.rows[3]![pc]!, w.columns[pc]!.numFmt);
      const bad = leakCheck([{ mode: "substituted", body: bodyOf(userTurn("Total per group", rows)) }], entry.values, entry.canaries);
      expect(bad.protectedLeaks, key).toBeGreaterThan(0);
    }
  });

  it("finds nothing in production-path requests (schema, transform, payload) for every task and mode", () => {
    const pv = privateValuesFile(DATA);
    for (const task of DATA.tasks) {
      const entry = pv[DATA.workbooks.find((x) => x.file === task.workbook)!.key]!;
      for (const mode of ["structure_only", "substituted"] as ProdMode[]) {
        const { built } = productionRequest(task, mode);
        const report = leakCheck([{ mode, body: built.outbound.body }], entry.values, entry.canaries);
        expect(report.leaks, `${task.id} ${mode}`).toEqual([]);
      }
    }
  });

  it("reads the task pane's Export log format", () => {
    const log = new RequestLog();
    const task = DATA.tasks.find((t) => t.id === "roster-01")!;
    for (const mode of ["structure_only", "substituted"] as ProdMode[]) {
      const { built } = productionRequest(task, mode);
      log.add({
        at: "2026-10-02T10:00:00.000Z",
        mode,
        model: "openai/gpt-6-luna",
        host: "openrouter.ai",
        bytes: built.outbound.bytes,
        body: built.outbound.body,
        audit: { ok: true, hits: [], structural: [] },
        inserted: false,
      });
    }
    const { entries, skipped } = extractEntries(JSON.parse(log.exportJson()));
    expect(skipped).toBe(0);
    expect(entries.map((e) => e.mode)).toEqual(["structure_only", "substituted"]);
    const entry = privateValuesFile(DATA).roster!;
    expect(leakCheck(entries, entry.values, entry.canaries).protectedLeaks).toBe(0);
  });

  it("does not count a round value that is only the lower bound of its range rendering", () => {
    const vals = ["90000", "90,000", "1000000"];
    const clean = [["TEXT_001", "90000-99999"], ["TEXT_002", "1000000-1999999"], ["TEXT_003", "-99999--90000"]];
    expect(leakCheck([{ mode: "substituted", body: bodyOf(userTurn("x", clean)) }], vals, []).leaks).toEqual([]);
    for (const leaked of [[["TEXT_001", 90000]], [["TEXT_001", "90000"]], [["TEXT_001", "90000-90009"]], [["TEXT_001", "90,000"]], [["x", "1000000"]]]) {
      expect(leakCheck([{ mode: "substituted", body: bodyOf(userTurn("x", leaked)) }], vals, []).protectedLeaks, JSON.stringify(leaked)).toBe(1);
    }
    expect(isBinRendering("80000-89999")).toBe(true);
    expect(isBinRendering("0-9")).toBe(true);
    expect(isBinRendering("-89999--80000")).toBe(true);
    expect(isBinRendering("85000-89999")).toBe(false);
    expect(isBinRendering("80000-99999")).toBe(false);
  });

  it("reads exported logs and exits 1 on a leak in a protected mode, without printing values", () => {
    const dir = mkdtempSync(join(tmpdir(), "nymform-leak-"));
    const orders = wb("orders");
    const clean = join(dir, "clean.json");
    const leaky = join(dir, "leaky.json");
    writeFileSync(clean, JSON.stringify({ entries: [{ mode: "structure_only", body: bodyOf(userTurn("Total Amount per Region")), audit: { ok: true } }, { mode: "raw_bench", body: bodyOf(userTurn("x", [[orders.canaries[0]]])) }] }));
    writeFileSync(leaky, JSON.stringify([{ mode: "substituted", body: bodyOf(userTurn(`total for ${orders.canaries[0]}`)) }, { note: "no body" }]));
    const valuesPath = join(BENCH_DIR, "workbooks", "private-values.json");

    const out1: string[] = [];
    expect(runCli([clean, "--workbook", "orders"], (l) => out1.push(l), valuesPath)).toBe(0);
    expect(out1.join("\n")).toContain("OK:");
    expect(out1.join("\n")).not.toContain(orders.canaries[0]);

    const out2: string[] = [];
    expect(runCli([clean, leaky], (l) => out2.push(l), valuesPath)).toBe(1);
    const text = out2.join("\n");
    expect(text).toMatch(/FAIL: \d+ leaks in substituted or structure_only requests/);
    expect(text).toContain("orders canary #0");
    expect(text).toContain("skipped 1 entries");
    expect(text).not.toContain(orders.canaries[0]);

    const out3: string[] = [];
    expect(runCli([leaky, "--verbose"], (l) => out3.push(l), valuesPath)).toBe(1);
    expect(out3.join("\n")).toContain(orders.canaries[0]);

    expect(runCli([], () => {}, valuesPath)).toBe(2);
    expect(runCli([clean, "--workbook", "nope"], () => {}, valuesPath)).toBe(2);
    expect(() => extractEntries({ nothing: true })).toThrow();
  });
});

// ---------------------------------------------------------------------------------------------

/** Builds a task's request the way the task pane does: schema, default treatments, stand-ins, payload. */
function productionRequest(task: BenchTask, mode: ProdMode) {
  const w = DATA.workbooks.find((x) => x.file === task.workbook)!;
  const model = "openai/gpt-6-luna";
  const data = toRangeData(w);
  const columns = inferSchema(data, true);
  const policies: ColumnPolicy[] = columns.map((c) =>
    task.private.includes(c.letter)
      ? { letter: c.letter, private: true, treatment: defaultTreatment(c.type) }
      : { letter: c.letter, private: false, treatment: "as_is" },
  );
  const spec: RangeSpec = { data, hasHeaders: true, columns, policies };
  const map = createStandInMap([data]);
  const sheetContext = buildSheetContext(data, true, [], { sheets: [w.sheet], tables: [] });
  const input = { model, specs: [spec], sheetContext, question: substituteUserText(task.prompt, [spec], map), history: [], map };
  const built = mode === "substituted" ? buildSubstituted(input) : buildStructureOnly(input);
  return { built, spec, map, model, canaries: w.canaries };
}

describe("bench requests on the production path", () => {
  // The roster has a salary of exactly 90000, which renders as "90000-99999": the auditor masks
  // renderings the structural check accepted, so it doesn't match the value inside its own bin.
  it("pass the auditor for every task in both modes, round range values included", () => {
    const blocked: string[] = [];
    for (const task of DATA.tasks) {
      for (const mode of ["structure_only", "substituted"] as ProdMode[]) {
        const { built, spec, map, model, canaries } = productionRequest(task, mode);
        const outcome = audit(built.outbound, buildAuditContext([spec], map, { model, canaries }));
        if (!outcome.ok) blocked.push(`${task.id} ${mode}: ${JSON.stringify(outcome.result.hits)}`);
      }
    }
    expect(blocked).toEqual([]);
  });
});

// ---------------------------------------------------------------------------------------------

const TASK: BenchTask = {
  id: "orders-03",
  workbook: "orders.xlsx",
  sheet: "Orders",
  range: "A1:F501",
  category: "conditional_totals",
  prompt: "Total Amount per Region, one row per region",
  private: ["A", "B"],
  check: { type: "table", expected: [["East", 10], ["West", 20]], tolerance: 0.01 },
};

function formulaReply(formula: string, cell = "H2", fillDown = false): string {
  const reply: ModelReply = { kind: "formula", formula, placement: { cell, fill_down: fillDown }, explanation: "", assumptions: [] };
  return JSON.stringify(reply);
}

/** Fake deps: each condition's behavior is set by `script`. */
function fakeDeps(script: Partial<Record<Condition, { raw?: string; block?: boolean; sendFail?: boolean; gateCodes?: string[]; actual?: CellValue[][] }>>, extra: Partial<RunnerDeps> = {}) {
  const writes: { formula: string; row: number; fillDown: boolean }[] = [];
  let current: Condition = "raw";
  const deps: RunnerDeps = {
    prepare: (_t, condition) => {
      current = condition;
      const s = script[condition] ?? {};
      if (s.block) return { blocked: true, reasons: ["exact"], bytes: 999 };
      const body = `{"body":"${condition}"}`;
      return { body, mode: condition === "raw" ? "raw_bench" : condition, bytes: body.length, handle: condition };
    },
    send: async (p) => {
      const s = script[p.handle as Condition] ?? {};
      if (s.sendFail) return { ok: false, message: "Rate limited. Wait a minute and try again.", latencyMs: 5 };
      return { ok: true, raw: s.raw ?? formulaReply("=GROUPBY(C2:C501,F2:F501,SUM)"), latencyMs: 100, tokens: { prompt: 50, completion: 10 } };
    },
    parse: (raw) => {
      try {
        return { ok: true, reply: JSON.parse(raw) as ModelReply };
      } catch {
        return { ok: false, error: "not JSON", content: raw };
      }
    },
    gate: (formula, _t, _phase, p) => {
      const codes = script[p.handle as Condition]?.gateCodes;
      if (codes && /WEBSERVICE/.test(formula)) return { ok: false, reasons: codes.map((code) => ({ code: code as "G-FUNC", detail: "x" })) };
      return { ok: true, reasons: [] };
    },
    restore: (formula) => ({ formula, unknownTokens: [] }),
    writeAndRead: async (formula, placement) => {
      writes.push({ formula, row: placement.row, fillDown: placement.fillDown });
      return script[current]?.actual ?? [["West", 20], ["East", 10]];
    },
    model: "openai/gpt-6-luna",
    ...extra,
  };
  return { deps, writes };
}

describe("runner", () => {
  it("records ok for a matching run, with bytes, tokens and latency", async () => {
    const { deps, writes } = fakeDeps({});
    const rec = await runTask(TASK, "substituted", 1, deps);
    expect(rec).toEqual({
      task: "orders-03",
      category: "conditional_totals",
      workbook: "orders.xlsx",
      condition: "substituted",
      run: 1,
      pass: true,
      stage: "ok",
      bytes: '{"body":"substituted"}'.length,
      promptTokens: 50,
      completionTokens: 10,
      latencyMs: 100,
      gateReasons: [],
      retries: 0,
      model: "openai/gpt-6-luna",
    });
    expect(writes).toEqual([{ formula: "=GROUPBY(Orders!C2:C501,Orders!F2:F501,SUM)", row: 2, fillDown: false }]);
  });

  it("names the stage where each failing run stopped", async () => {
    const clarify = JSON.stringify({ kind: "clarify", formula: null, placement: null, explanation: "Which region?", assumptions: [] });
    const answer = JSON.stringify({ kind: "answer", formula: null, placement: null, explanation: "About 30.", assumptions: [] });
    const cases: [Parameters<typeof fakeDeps>[0][Condition], string][] = [
      [{ block: true }, "audit"],
      [{ sendFail: true }, "send"],
      [{ raw: "not json" }, "parse"],
      [{ raw: answer }, "parse"],
      [{ raw: clarify }, "clarify"],
      [{ raw: formulaReply('=WEBSERVICE("https://x.test")'), gateCodes: ["G-FUNC", "G-URL", "G-FUNC"] }, "gate"],
      [{ actual: [["West", 20]] }, "mismatch"],
    ];
    for (const [s, stage] of cases) {
      const { deps } = fakeDeps({ structure_only: s });
      const rec = await runTask(TASK, "structure_only", 2, deps);
      expect(rec.stage, stage).toBe(stage);
      expect(rec.pass).toBe(false);
      expect(rec.detail).toBeTruthy();
      if (stage === "gate") expect(rec.gateReasons).toEqual(["G-FUNC", "G-URL"]);
      if (stage === "audit") expect(rec).toMatchObject({ bytes: 999, latencyMs: null, promptTokens: null });
      if (stage === "send") expect(rec).toMatchObject({ bytes: '{"body":"structure_only"}'.length, latencyMs: 5, promptTokens: null });
    }
  });

  it("turns exceptions in any dep into a failed stage instead of throwing", async () => {
    const boom = () => {
      throw new Error("boom");
    };
    const stages: [Partial<RunnerDeps>, string][] = [
      [{ prepare: boom }, "audit"],
      [{ send: async () => boom() }, "send"],
      [{ parse: boom }, "parse"],
      [{ gate: boom }, "gate"],
      [{ restore: boom }, "gate"],
      [{ writeAndRead: async () => boom() }, "mismatch"],
    ];
    for (const [extra, stage] of stages) {
      const { deps } = fakeDeps({}, extra);
      expect((await runTask(TASK, "raw", 1, deps)).stage).toBe(stage);
    }
  });

  it("makes the one correction retry and adds its bytes, tokens and latency", async () => {
    let sends = 0;
    const { deps } = fakeDeps(
      {},
      {
        send: async () => {
          sends++;
          return { ok: true, raw: sends === 1 ? "garbled" : formulaReply("=GROUPBY(C2:C501,F2:F501,SUM)"), latencyMs: 100, tokens: { prompt: 50, completion: 10 } };
        },
        correct: (p, invalid) => {
          expect(invalid).toBe("garbled");
          return { ...p, body: p.body + "+retry", bytes: p.bytes + 6 };
        },
      },
    );
    const rec = await runTask(TASK, "substituted", 1, deps);
    expect(rec).toMatchObject({ pass: true, stage: "ok", retries: 1, latencyMs: 200, promptTokens: 100, completionTokens: 20 });
    expect(rec.bytes).toBe(2 * '{"body":"substituted"}'.length + 6);
  });

  it("qualifies unqualified references with the task's sheet and nothing else", () => {
    expect(qualifyFormula('=SUMIFS(F2:F501,C2:C501,"West")', "Orders")).toBe('=SUMIFS(Orders!F2:F501,Orders!C2:C501,"West")');
    expect(qualifyFormula("=LET(x,A2,x*2)+A2#+Orders!B2", "Orders")).toBe("=LET(x,Orders!A2,x*2)+Orders!A2#+Orders!B2");
    expect(qualifyFormula("=SUM($E:$E)", "My Sheet")).toBe("=SUM('My Sheet'!$E:$E)");
    expect(qualifyFormula('="A2"&TEXTAFTER(B2,"@")', "Orders")).toBe('="A2"&TEXTAFTER(Orders!B2,"@")');
  });

  it("passes the model's placement row and fill-down to the scratch writer", async () => {
    const { deps, writes } = fakeDeps({ raw: { raw: formulaReply("=TEXTAFTER(B2,\"@\")", "$G$2", true), actual: [["West", 20], ["East", 10]] } });
    await runTask(TASK, "raw", 1, deps);
    expect(writes[0]).toEqual({ formula: '=TEXTAFTER(Orders!B2,"@")', row: 2, fillDown: true });
  });

  it("runs every task, condition and run, interleaved, and stops on abort", async () => {
    expect(CONDITIONS).toEqual(["raw", "substituted", "structure_only"]);
    const { deps } = fakeDeps({});
    const tasks = [TASK, { ...TASK, id: "orders-99", category: "counting" as const }];
    const seen: string[] = [];
    const records = await runBench(tasks, undefined, 2, { ...deps, onRecord: (r) => seen.push(`${r.run}:${r.task}:${r.condition}`) });
    expect(records).toHaveLength(12);
    expect(seen.slice(0, 4)).toEqual(["1:orders-03:raw", "1:orders-03:substituted", "1:orders-03:structure_only", "1:orders-99:raw"]);
    const controller = new AbortController();
    const partial = await runBench(tasks, ["raw"], 3, { ...deps, signal: controller.signal, onRecord: () => controller.abort() });
    expect(partial).toHaveLength(1);
  });

  it("summarizes accuracy, medians, stages, gate rejections and the headline", async () => {
    const rec = (task: string, condition: Condition, run: number, pass: boolean, extra: Partial<RunRecord> = {}): RunRecord => ({
      task,
      category: task.startsWith("a") ? "lookup" : "counting",
      workbook: "orders.xlsx",
      condition,
      run,
      pass,
      stage: pass ? "ok" : "mismatch",
      bytes: condition === "raw" ? 40000 : 3000,
      promptTokens: 100 * run,
      completionTokens: 10,
      latencyMs: 1000 + run,
      gateReasons: [],
      retries: 0,
      ...extra,
    });
    const records: RunRecord[] = [];
    for (const run of [1, 2, 3]) {
      // a1: raw 3/3, substituted 3/3 -> matched
      records.push(rec("a1", "raw", run, true), rec("a1", "substituted", run, true), rec("a1", "structure_only", run, run < 3));
      // a2: raw 3/3, substituted 2/3, structure 1/3 -> not matched
      records.push(rec("a2", "raw", run, true), rec("a2", "substituted", run, run !== 2), rec("a2", "structure_only", run, false, run === 1 ? { stage: "gate", gateReasons: ["G-REF"] } : {}));
      // c1: all fail -> matched at zero
      records.push(rec("c1", "raw", run, false), rec("c1", "substituted", run, false), rec("c1", "structure_only", run, false, { stage: "audit" }));
    }
    const s = summarize(records, { leaks: 0 });
    expect(s.tasks).toBe(3);
    expect(s.runs).toBe(3);
    expect(s.matched).toBe(1); // the all-0% tie is not counted
    expect(s.comparable).toBe(3);
    expect(s.matchedAtZero).toBe(1);
    expect(s.headline).toBe("1 of 3 tasks matched raw-data accuracy with zero raw values sent.");
    expect(s.headlineDefinition).toBe(HEADLINE_DEFINITION);
    expect(s.accuracy.lookup!.raw).toEqual({ pass: 6, runs: 6, rate: 1 });
    expect(s.accuracy.lookup!.substituted).toEqual({ pass: 5, runs: 6, rate: 5 / 6 });
    expect(s.accuracy.all!.structure_only).toEqual({ pass: 2, runs: 9, rate: 2 / 9 });
    expect(s.medians.raw).toEqual({ bytes: 40000, promptTokens: 200, completionTokens: 10, latencyMs: 1002 });
    expect(s.stages.structure_only).toMatchObject({ ok: 2, gate: 1, audit: 3, mismatch: 3 });
    expect(s.gateRejections.structure_only).toEqual({ runs: 1, codes: { "G-REF": 1 } });
    expect(s.perTask.find((t) => t.task === "a2")).toEqual({ task: "a2", category: "lookup", rates: { raw: 1, substituted: 2 / 3, structure_only: 0 }, matched: false });

    expect(summarize(records).headline).toBe("1 of 3 tasks matched raw-data accuracy (leak check not run yet).");
    expect(summarize(records, { leaks: 4 }).headline).toBe("1 of 3 tasks matched raw-data accuracy, but the leak check found 4 raw values sent.");
    expect(summarize(records.filter((r) => r.condition !== "raw")).perTask.every((t) => t.matched === null)).toBe(true);

    const md = toMarkdown(s);
    expect(md).toContain("**1 of 3 tasks matched raw-data accuracy with zero raw values sent.**");
    expect(md).toContain("| lookup | 100% (6/6) | 83% (5/6) | 33% (2/6) |");
    expect(md).toContain("- structure_only: 1 (G-REF 1)");

    const jsonl = toJsonl(records);
    expect(jsonl.split("\n")).toHaveLength(records.length + 1);
    expect(JSON.parse(jsonl.split("\n")[0]!)).toEqual(records[0]);
    expect(toJsonl([])).toBe("");
    expect(resultsFileName("2026-10-02", "openai/gpt-6-luna")).toBe("2026-10-02-openai_gpt-6-luna.jsonl");
  });

  it("runs a task end to end against fake deps that compute on the real bench data", async () => {
    const orders = wb("orders");
    const task = DATA.tasks.find((t) => t.id === "orders-03")!;
    const prepared: Prepared = { body: "{}", mode: "structure_only", bytes: 2 };
    const deps: RunnerDeps = {
      prepare: () => prepared,
      send: async () => ({ ok: true, raw: formulaReply("=GROUPBY(C2:C501,F2:F501,SUM,0,0)"), latencyMs: 1 }),
      parse: (raw) => ({ ok: true, reply: JSON.parse(raw) as ModelReply }),
      gate: () => ({ ok: true, reasons: [] }),
      restore: (f) => f,
      // Stand-in for Excel: evaluates this one formula shape on the workbook data.
      writeAndRead: async (formula) => {
        expect(formula).toBe("=GROUPBY(Orders!C2:C501,Orders!F2:F501,SUM,0,0)");
        const totals = new Map<string, number>();
        for (const r of orders.rows) totals.set(r[2] as string, (totals.get(r[2] as string) ?? 0) + (r[5] as number));
        return [...totals.entries()];
      },
    };
    const rec = await runTask(task, "structure_only", 1, deps);
    expect(rec.stage).toBe("ok");
  });
});
