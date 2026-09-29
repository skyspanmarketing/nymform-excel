// payload.ts (spec §7.6): the one builder of request bodies. Checks the exact body and user-turn
// format from the brief, and that the real auditor accepts what payload.ts builds.
import { afterEach, describe, expect, it, vi } from "vitest";
import { parseLocalRange } from "../src/core/a1";
import { audit, buildAuditContext } from "../src/core/auditor";
import {
  buildCorrection,
  buildStructureOnly,
  buildSubstituted,
  isReasoningModel,
  requestParams,
  type PayloadInput,
} from "../src/core/payload";
import { CORRECTION_PROMPT, SYSTEM_PROMPT } from "../src/core/prompt";
import { buildSheetContext, inferSchema } from "../src/core/schema";
import { createStandInMap, substituteUserText, type StandInMap } from "../src/core/transform";
import type { CellValue, ColumnPolicy, HistoryMessage, RangeData, RangeSpec } from "../src/core/types";

const MODEL = "openai/gpt-6-luna";

// ---------------------------------------------------------------------------------------------
// Fixtures

function serial(y: number, m: number, d: number): number {
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86_400_000;
}

interface DataOpts {
  /** Number format per column offset (data rows only). */
  formats?: Record<number, string>;
  /** Display text per "row,col" when it differs from String(value). */
  texts?: Record<string, string>;
  table?: string;
}

function rangeData(sheet: string, address: string, rows: CellValue[][], opts: DataOpts = {}): RangeData {
  const rect = parseLocalRange(address);
  if (!rect) throw new Error(`bad address ${address}`);
  const data: RangeData = {
    sheet,
    address,
    rowIndex: rect.r1,
    columnIndex: rect.c1,
    values: rows.map((r) => [...r]),
    text: rows.map((r, ri) => r.map((v, ci) => opts.texts?.[`${ri},${ci}`] ?? (v === null ? "" : String(v)))),
    formulas: rows.map((r) => [...r]),
    valueTypes: rows.map((r) =>
      r.map((v) => (v === null ? "Empty" : typeof v === "number" ? "Double" : typeof v === "boolean" ? "Boolean" : "String")),
    ),
    numberFormat: rows.map((r, ri) => r.map((_, ci) => (ri > 0 ? (opts.formats?.[ci] ?? "General") : "General"))),
  };
  if (opts.table) data.table = { name: opts.table, columns: rows[0]?.map(String) ?? [] };
  return data;
}

function spec(data: RangeData, policies: ColumnPolicy[], hasHeaders = true): RangeSpec {
  return { data, hasHeaders, columns: inferSchema(data, hasHeaders), policies };
}

// Words of 4+ characters from these notes must not appear in headers or the question: the
// auditor's word variant would (rightly) block them.
const NOTE_1 = "Called twice about the late delivery last month";
const NOTE_2 = "Asked for a refund and wants a callback soon";

function ordersData(): RangeData {
  return rangeData(
    "Orders",
    "A1:F6",
    [
      ["Customer", "Email", "Region", "Order date", "Notes", "Amount"],
      ["Maria Lopez", "maria.lopez@example.com", "East", serial(2025, 12, 1), NOTE_1, 83500],
      ['Ana "AJ" Ruiz', "ana.ruiz@example.com", "West", serial(2026, 1, 15), NOTE_2, 41250],
      ["Kenji Watanabe", "kenji@example.com", "East", serial(2026, 2, 3), null, 5],
      ["Maria Lopez", "maria.lopez@example.com", "North", serial(2026, 3, 9), "Wants the invoice by post today", -83500],
      ["Priya Natarajan", null, "South", serial(2026, 3, 30), "Prefers morning deliveries on weekdays", 0],
    ],
    {
      formats: { 3: "yyyy-mm-dd" },
      texts: { "1,3": "2025-12-01", "2,3": "2026-01-15", "3,3": "2026-02-03", "4,3": "2026-03-09", "5,3": "2026-03-30" },
    },
  );
}

const ORDERS_POLICIES: ColumnPolicy[] = [
  { letter: "A", private: true, treatment: "stand_in" },
  // A private column set to as_is is sent as stand_in (brief, decision 3).
  { letter: "B", private: true, treatment: "as_is" },
  { letter: "C", private: false, treatment: "as_is" },
  { letter: "D", private: true, treatment: "month" },
  { letter: "E", private: true, treatment: "exclude" },
  { letter: "F", private: true, treatment: "range", note: "  USD  " },
];

function regionsData(): RangeData {
  return rangeData("Regions", "A1:B5", [
    ["Region", "Manager"],
    ["East", "Dana Whitfield"],
    ["West", "Omar Haddad"],
    ["North", "Lucia Ferreira"],
    ["South", "Tobias Brandt"],
  ]);
}

const REGIONS_POLICIES: ColumnPolicy[] = [
  { letter: "A", private: false, treatment: "as_is" },
  { letter: "B", private: true, treatment: "stand_in" },
];

interface Setup {
  specs: RangeSpec[];
  map: StandInMap;
  input: PayloadInput;
}

function setup(overrides: Partial<PayloadInput> = {}, withContext = false): Setup {
  const orders = spec(ordersData(), ORDERS_POLICIES);
  const specs = withContext ? [orders, spec(regionsData(), REGIONS_POLICIES)] : [orders];
  const map = createStandInMap(specs.map((s) => s.data));
  const sheetContext = buildSheetContext(
    orders.data,
    true,
    specs.slice(1).map((s) => s.data),
    { sheets: ["Orders", "Regions", "Archive"], tables: [] },
  );
  const input: PayloadInput = {
    model: MODEL,
    specs,
    sheetContext,
    question: "Total Amount per Region in a new column",
    history: [],
    map,
    ...overrides,
  };
  return { specs, map, input };
}

function auditFor(s: Setup, body: { outbound: Parameters<typeof audit>[0] }) {
  return audit(body.outbound, buildAuditContext(s.specs, s.map, { model: MODEL }));
}

function lastUserTurn(body: string): Record<string, unknown> {
  const parsed = JSON.parse(body) as { messages: { role: string; content: string }[] };
  const last = parsed.messages[parsed.messages.length - 1]!;
  expect(last.role).toBe("user");
  return JSON.parse(last.content) as Record<string, unknown>;
}

const PRIVATE_VALUES = [
  "Maria Lopez",
  "Ana \"AJ\" Ruiz",
  "Kenji Watanabe",
  "Priya Natarajan",
  "maria.lopez@example.com",
  "ana.ruiz@example.com",
  "kenji@example.com",
  "83500",
  "41250",
  "2025-12-01",
  "2026-01-15",
  NOTE_1,
  NOTE_2,
  "Dana Whitfield",
  "Omar Haddad",
];

// ---------------------------------------------------------------------------------------------

describe("requestParams", () => {
  it("reasoning models get max_tokens 4000, low reasoning effort, JSON output and no temperature", () => {
    const p = requestParams(MODEL);
    expect(JSON.stringify(p)).toBe(
      '{"max_tokens":4000,"reasoning":{"effort":"low","exclude":true},"response_format":{"type":"json_object"}}',
    );
    expect(p).not.toHaveProperty("temperature");
    for (const id of ["openai/gpt-5-mini", "openai/o1", "openai/o3-mini", "openai/o4-mini", "openai/gpt-6-luna"]) {
      expect(isReasoningModel(id)).toBe(true);
    }
  });

  it("other models get temperature 0, max_tokens 800 and JSON output", () => {
    for (const id of ["anthropic/claude-sonnet-4.5", "openai/gpt-4.1", "google/gemini-2.5-pro"]) {
      expect(isReasoningModel(id)).toBe(false);
      expect(JSON.stringify(requestParams(id))).toBe(
        '{"temperature":0,"max_tokens":800,"response_format":{"type":"json_object"}}',
      );
    }
  });

  it("returns a fresh object each time", () => {
    const a = requestParams(MODEL);
    (a.reasoning as { effort: string }).effort = "high";
    expect(requestParams(MODEL)).toEqual({
      max_tokens: 4000,
      reasoning: { effort: "low", exclude: true },
      response_format: { type: "json_object" },
    });
  });
});

describe("buildStructureOnly", () => {
  it("builds the body with keys in order, no whitespace, and provider zdr for openrouter.ai (T-P2)", () => {
    const s = setup();
    const built = buildStructureOnly(s.input);
    const body = built.outbound.body;
    const parsed = JSON.parse(body) as Record<string, unknown>;

    expect(JSON.stringify(parsed)).toBe(body);
    expect(Object.keys(parsed)).toEqual(["model", "messages", "max_tokens", "reasoning", "response_format", "provider"]);
    expect(parsed.model).toBe(MODEL);
    expect(parsed.provider).toEqual({ zdr: true });
    expect(body).toContain('"provider":{"zdr":true}');
    expect(body.endsWith('"provider":{"zdr":true}}')).toBe(true);
    expect(body).not.toContain("temperature");

    const messages = parsed.messages as { role: string; content: string }[];
    expect(messages).toEqual([
      { role: "system", content: SYSTEM_PROMPT },
      { role: "user", content: built.userContent },
    ]);
    expect(built.messages).toEqual(messages);
  });

  it("sets mode, bytes (UTF-8 length) and an ISO createdAt on the Outbound", () => {
    const s = setup({ question: "Überblick: Summe je Region – café ☕" });
    const before = Date.now();
    const built = buildStructureOnly(s.input);
    const { outbound } = built;
    expect(outbound.mode).toBe("structure_only");
    expect(outbound.bytes).toBe(new TextEncoder().encode(outbound.body).byteLength);
    expect(outbound.bytes).toBeGreaterThan(outbound.body.length);
    expect(new Date(outbound.createdAt).toISOString()).toBe(outbound.createdAt);
    expect(Date.parse(outbound.createdAt)).toBeGreaterThanOrEqual(before - 1000);
    expect(built.rowsSent).toBeUndefined();
    expect(built.rowsTotal).toBeUndefined();
  });

  it("writes the user turn in the brief's format, keys in order", () => {
    const s = setup();
    const built = buildStructureOnly(s.input);
    const turn = JSON.parse(built.userContent) as Record<string, unknown>;
    expect(JSON.stringify(turn)).toBe(built.userContent);
    expect(Object.keys(turn)).toEqual([
      "nymform",
      "mode",
      "sheet",
      "table",
      "range",
      "columns",
      "context_ranges",
      "allowed_ranges",
      "task",
    ]);
    expect(turn.nymform).toBe("0.1");
    expect(turn.mode).toBe("structure_only");
    expect(turn.sheet).toBe("Orders");
    expect(turn.table).toBeNull();
    expect(turn.range).toEqual({ address: "A1:F6", header_row: 1, first_data_row: 2, last_data_row: 6 });
    expect(Object.keys(turn.range as object)).toEqual(["address", "header_row", "first_data_row", "last_data_row"]);
    expect(turn.context_ranges).toEqual([]);
    expect(turn.allowed_ranges).toEqual(["Orders!A1:F6"]);
    expect(turn.task).toBe("Total Amount per Region in a new column");
  });

  it("describes columns with effective treatments, notes and text-only stats", () => {
    const built = buildStructureOnly(setup().input);
    const cols = (JSON.parse(built.userContent) as { columns: Record<string, unknown>[] }).columns;
    expect(cols.map((c) => Object.keys(c))).toEqual(
      cols.map(() => ["letter", "header", "type", "private", "treatment", "note", "stats"]),
    );
    expect(cols.map((c) => c.letter)).toEqual(["A", "B", "C", "D", "E", "F"]);
    expect(cols.map((c) => c.header)).toEqual(["Customer", "Email", "Region", "Order date", "Notes", "Amount"]);
    expect(cols.map((c) => c.type)).toEqual(["text", "text", "text", "date", "text", "number"]);
    expect(cols.map((c) => c.private)).toEqual([true, true, false, true, true, true]);
    expect(cols.map((c) => c.treatment)).toEqual(["stand_in", "stand_in", "as_is", "month", "exclude", "range"]);
    expect(cols.map((c) => c.note)).toEqual([null, null, null, null, null, "USD"]);

    const statsA = cols[0]!.stats as Record<string, unknown>;
    expect(Object.keys(statsA)).toEqual(["blank", "distinct", "avg_words", "max_len"]);
    // Five names, one of three words: (2 + 3 + 2 + 2 + 2) / 5.
    expect(statsA).toEqual({ blank: 0, distinct: 4, avg_words: 2.2, max_len: 15 });
    expect(cols[1]!.stats).toMatchObject({ blank: 1, distinct: 3 });
    expect(Object.keys(cols[3]!.stats as object)).toEqual(["blank", "distinct"]);
    expect(Object.keys(cols[5]!.stats as object)).toEqual(["blank", "distinct"]);
  });

  it("never includes rows, rows_sent or rows_total, and no cell values", () => {
    const s = setup({}, true);
    const built = buildStructureOnly(s.input);
    const turn = JSON.parse(built.userContent) as Record<string, unknown>;
    for (const block of [turn, ...(turn.context_ranges as Record<string, unknown>[])]) {
      expect(block).not.toHaveProperty("rows");
      expect(block).not.toHaveProperty("rows_sent");
      expect(block).not.toHaveProperty("rows_total");
    }
    expect(built.outbound.body).not.toMatch(/"rows(_sent|_total)?"/);
    for (const v of PRIVATE_VALUES) expect(built.outbound.body).not.toContain(JSON.stringify(v).slice(1, -1));
    for (const v of ["East", "West", "North", "South"]) expect(built.userContent).not.toContain(v);
  });

  it("passes the real auditor", () => {
    const s = setup({}, true);
    const outcome = auditFor(s, buildStructureOnly(s.input));
    expect(outcome.ok).toBe(true);
  });

  it("sends the alias instead of a revealing header, and the original never appears", () => {
    const data = rangeData("Clinic", "A1:B3", [
      ["Patient", "Diagnosis"],
      ["Maria Lopez", "Asthma"],
      ["Kenji Watanabe", "Diabetes"],
    ]);
    const s0 = spec(data, [
      { letter: "A", private: true, treatment: "stand_in" },
      { letter: "B", private: true, treatment: "stand_in", alias: "Field B" },
    ]);
    const map = createStandInMap([data]);
    const input: PayloadInput = {
      model: MODEL,
      specs: [s0],
      sheetContext: buildSheetContext(data, true, [], { sheets: ["Clinic"], tables: [] }),
      question: "Count rows per Field B",
      history: [],
      map,
    };
    for (const built of [buildStructureOnly(input), buildSubstituted(input)]) {
      const cols = (JSON.parse(built.userContent) as { columns: { header: string }[] }).columns;
      expect(cols.map((c) => c.header)).toEqual(["Patient", "Field B"]);
      expect(built.outbound.body).not.toMatch(/diagnosis/i);
      expect(audit(built.outbound, buildAuditContext([s0], map, { model: MODEL })).ok).toBe(true);
    }
  });

  it("sends header_row null and 'Column {letter}' headers when the range has no header row", () => {
    const data = rangeData("Data", "C5:D7", [
      ["alpha", 1],
      ["beta", 2],
      ["gamma", 3],
    ]);
    const s0 = spec(
      data,
      [
        { letter: "C", private: false, treatment: "as_is" },
        { letter: "D", private: false, treatment: "as_is" },
      ],
      false,
    );
    const built = buildStructureOnly({
      model: MODEL,
      specs: [s0],
      sheetContext: buildSheetContext(data, false, [], { sheets: ["Data"], tables: [] }),
      question: "Sum column D",
      history: [],
      map: createStandInMap([data]),
    });
    const turn = JSON.parse(built.userContent) as { range: unknown; columns: { letter: string; header: string }[] };
    expect(turn.range).toEqual({ address: "C5:D7", header_row: null, first_data_row: 5, last_data_row: 7 });
    expect(turn.columns.map((c) => [c.letter, c.header])).toEqual([
      ["C", "Column C"],
      ["D", "Column D"],
    ]);
  });

  it("names the table when the selection is inside one", () => {
    const data = rangeData("Sales", "B2:C4", [["Rep", "Total"], ["x", 1], ["y", 2]], { table: "SalesTable" });
    const s0 = spec(data, [
      { letter: "B", private: true, treatment: "stand_in" },
      { letter: "C", private: false, treatment: "as_is" },
    ]);
    const built = buildStructureOnly({
      model: MODEL,
      specs: [s0],
      sheetContext: buildSheetContext(data, true, [], { sheets: ["Sales"], tables: [] }),
      question: "Total per rep",
      history: [],
      map: createStandInMap([data]),
    });
    const turn = JSON.parse(built.userContent) as { table: unknown; range: unknown };
    expect(turn.table).toBe("SalesTable");
    expect(turn.range).toEqual({ address: "B2:C4", header_row: 2, first_data_row: 3, last_data_row: 4 });
  });

  it("describes context ranges like the selection and lists them in allowed_ranges", () => {
    const s = setup({}, true);
    const turn = JSON.parse(buildStructureOnly(s.input).userContent) as Record<string, unknown>;
    const ctx = turn.context_ranges as Record<string, unknown>[];
    expect(ctx).toHaveLength(1);
    expect(Object.keys(ctx[0]!)).toEqual(["sheet", "table", "range", "columns"]);
    expect(ctx[0]).toMatchObject({
      sheet: "Regions",
      table: null,
      range: { address: "A1:B5", header_row: 1, first_data_row: 2, last_data_row: 5 },
    });
    const cols = ctx[0]!.columns as Record<string, unknown>[];
    expect(cols.map((c) => [c.letter, c.header, c.private, c.treatment])).toEqual([
      ["A", "Region", false, "as_is"],
      ["B", "Manager", true, "stand_in"],
    ]);
    expect(turn.allowed_ranges).toEqual(["Orders!A1:F6", "Regions!A1:B5"]);
    // A sheet the user didn't add is never sent.
    expect(JSON.stringify(turn)).not.toContain("Archive");
  });

  it("puts history between the system prompt and the new turn, at most 6 messages", () => {
    const history: HistoryMessage[] = Array.from({ length: 8 }, (_, i) => ({
      role: i % 2 === 0 ? "user" : "assistant",
      content: i % 2 === 0 ? JSON.stringify({ nymform: "0.1", n: i }) : `{"kind":"clarify","n":${i}}`,
    }));
    const s = setup({ history });
    const built = buildStructureOnly(s.input);
    expect(built.messages).toHaveLength(8);
    expect(built.messages[0]).toEqual({ role: "system", content: SYSTEM_PROMPT });
    expect(built.messages.slice(1, 7)).toEqual(history.slice(2));
    expect(built.messages[7]).toEqual({ role: "user", content: built.userContent });
  });
});

describe("buildSubstituted", () => {
  it("adds rows, rows_sent and rows_total after columns, with mode substituted", () => {
    const s = setup();
    const built = buildSubstituted(s.input);
    expect(built.outbound.mode).toBe("substituted");
    const turn = JSON.parse(built.userContent) as Record<string, unknown>;
    expect(Object.keys(turn)).toEqual([
      "nymform",
      "mode",
      "sheet",
      "table",
      "range",
      "columns",
      "rows",
      "rows_sent",
      "rows_total",
      "context_ranges",
      "allowed_ranges",
      "task",
    ]);
    expect(turn.mode).toBe("substituted");
    expect(turn.rows_sent).toBe(5);
    expect(turn.rows_total).toBe(5);
    expect(built.rowsSent).toBe(5);
    expect(built.rowsTotal).toBe(5);
  });

  it("sends stand-ins, ranges and months for private columns, drops excluded ones, keeps public cells", () => {
    const s = setup();
    const built = buildSubstituted(s.input);
    const turn = JSON.parse(built.userContent) as { rows: CellValue[][]; columns: { letter: string; treatment: string }[] };
    // Excluded column E stays in columns but not in rows: A, B, C, D, F.
    expect(turn.columns.find((c) => c.letter === "E")?.treatment).toBe("exclude");
    expect(turn.rows.every((r) => r.length === 5)).toBe(true);

    const original = [
      ["Maria Lopez", "maria.lopez@example.com"],
      ['Ana "AJ" Ruiz', "ana.ruiz@example.com"],
      ["Kenji Watanabe", "kenji@example.com"],
      ["Maria Lopez", "maria.lopez@example.com"],
      ["Priya Natarajan", null],
    ];
    turn.rows.forEach((row, i) => {
      const [name, email] = original[i]!;
      expect(s.map.isToken(row[0] as string)).toBe(true);
      expect(s.map.valueOf(row[0] as string)).toBe(name);
      if (email === null) expect(row[1]).toBeNull();
      else expect(s.map.valueOf(row[1] as string)).toBe(email);
    });
    // Same value, same token.
    expect(turn.rows[3]![0]).toBe(turn.rows[0]![0]);
    expect(turn.rows.map((r) => r[2])).toEqual(["East", "West", "East", "North", "South"]);
    expect(turn.rows.map((r) => r[3])).toEqual(["2025-12", "2026-01", "2026-02", "2026-03", "2026-03"]);
    expect(turn.rows.map((r) => r[4])).toEqual(["80000-89999", "40000-49999", "0-9", "-89999--80000", "0"]);

    for (const v of PRIVATE_VALUES) expect(built.outbound.body).not.toContain(JSON.stringify(v).slice(1, -1));
  });

  it("applies the row cap and reports rows_sent and rows_total", () => {
    const s = setup({ rowCap: 2 });
    const built = buildSubstituted(s.input);
    const turn = JSON.parse(built.userContent) as { rows: unknown[]; rows_sent: number; rows_total: number };
    expect(turn.rows).toHaveLength(2);
    expect(turn.rows_sent).toBe(2);
    expect(turn.rows_total).toBe(5);
    expect(built.rowsSent).toBe(2);
    expect(built.rowsTotal).toBe(5);
  });

  it("gives context ranges their own rows, after their columns", () => {
    const s = setup({}, true);
    const turn = JSON.parse(buildSubstituted(s.input).userContent) as Record<string, unknown>;
    const ctx = (turn.context_ranges as Record<string, unknown>[])[0]!;
    expect(Object.keys(ctx)).toEqual(["sheet", "table", "range", "columns", "rows", "rows_sent", "rows_total"]);
    const rows = ctx.rows as CellValue[][];
    expect(rows.map((r) => r[0])).toEqual(["East", "West", "North", "South"]);
    expect(rows.map((r) => s.map.valueOf(r[1] as string))).toEqual([
      "Dana Whitfield",
      "Omar Haddad",
      "Lucia Ferreira",
      "Tobias Brandt",
    ]);
    expect(ctx.rows_sent).toBe(4);
    expect(ctx.rows_total).toBe(4);
  });

  it("passes the real auditor, with the question's private value swapped for its stand-in", () => {
    const s = setup({}, true);
    s.input.question = substituteUserText("Total Amount for Maria Lopez and Dana Whitfield", s.specs, s.map);
    const built = buildSubstituted(s.input);
    const turn = JSON.parse(built.userContent) as { task: string };
    expect(turn.task).not.toContain("Maria");
    expect(turn.task).toContain(s.map.peek("Maria Lopez")!);
    const outcome = auditFor(s, built);
    expect(outcome.ok).toBe(true);
    if (outcome.ok) expect(outcome.outbound.body).toBe(built.outbound.body);
  });

  it("is blocked by the auditor if a private value gets into the body (sanity check of the pairing)", () => {
    const s = setup();
    const built = buildSubstituted({ ...s.input, question: "Total for Kenji Watanabe" });
    expect(auditFor(s, built).ok).toBe(false);
  });
});

describe("buildCorrection", () => {
  it("is [system, ...history, user turn, bad reply, correction prompt] with the same mode", () => {
    const s = setup();
    const first = buildSubstituted(s.input);
    const history: HistoryMessage[] = [];
    const retry = buildCorrection({
      model: MODEL,
      mode: "substituted",
      history,
      userContent: first.userContent,
      invalidReply: "Sure! The formula is =SUM(F2:F6)",
    });
    expect(retry.messages).toEqual([
      { role: "system", content: SYSTEM_PROMPT },
      { role: "user", content: first.userContent },
      { role: "assistant", content: "Sure! The formula is =SUM(F2:F6)" },
      { role: "user", content: CORRECTION_PROMPT },
    ]);
    expect(retry.userContent).toBe(first.userContent);
    expect(retry.outbound.mode).toBe("substituted");
    const parsed = JSON.parse(retry.outbound.body) as Record<string, unknown>;
    expect(JSON.stringify(parsed)).toBe(retry.outbound.body);
    expect(Object.keys(parsed)).toEqual(["model", "messages", "max_tokens", "reasoning", "response_format", "provider"]);
    expect(parsed.messages).toEqual(retry.messages);
    expect(retry.outbound.bytes).toBe(new TextEncoder().encode(retry.outbound.body).byteLength);
    expect(auditFor(s, retry).ok).toBe(true);
  });

  it("keeps history (at most 6 messages) before the turn being corrected", () => {
    const s = setup();
    const earlier = buildStructureOnly({ ...s.input, question: "Earlier question" });
    const history: HistoryMessage[] = [
      { role: "user", content: earlier.userContent },
      { role: "assistant", content: '{"kind":"clarify","formula":null,"placement":null,"explanation":"Which region?","assumptions":[]}' },
    ];
    const first = buildStructureOnly({ ...s.input, history });
    const retry = buildCorrection({
      model: MODEL,
      mode: "structure_only",
      history,
      userContent: first.userContent,
      invalidReply: "not json",
    });
    expect(retry.messages.map((m) => m.role)).toEqual(["system", "user", "assistant", "user", "assistant", "user"]);
    expect(retry.messages.slice(1, 3)).toEqual(history);
    expect(lastUserTurn(first.outbound.body)).toEqual(JSON.parse(first.userContent));
    expect(auditFor(s, retry).ok).toBe(true);
  });

  it("leaves out an empty bad reply", () => {
    const s = setup();
    const first = buildStructureOnly(s.input);
    const retry = buildCorrection({ model: MODEL, mode: "structure_only", history: [], userContent: first.userContent, invalidReply: "  " });
    expect(retry.messages.map((m) => m.role)).toEqual(["system", "user", "user"]);
    expect(retry.messages[2]!.content).toBe(CORRECTION_PROMPT);
  });

  it("uses a non-reasoning model's parameters", () => {
    const retry = buildCorrection({
      model: "anthropic/claude-sonnet-4.5",
      mode: "structure_only",
      history: [],
      userContent: '{"nymform":"0.1"}',
      invalidReply: "x",
    });
    const parsed = JSON.parse(retry.outbound.body) as Record<string, unknown>;
    expect(Object.keys(parsed)).toEqual(["model", "messages", "temperature", "max_tokens", "response_format", "provider"]);
    expect(parsed.temperature).toBe(0);
    expect(parsed.max_tokens).toBe(800);
  });
});

describe("endpoint other than openrouter.ai", () => {
  afterEach(() => {
    vi.doUnmock("../src/config");
    vi.resetModules();
  });

  it("leaves out the provider field (T-P2, negative case)", async () => {
    vi.resetModules();
    vi.doMock("../src/config", () => ({
      BENCH: false,
      ENDPOINT: "https://models.example.test/v1",
      ENDPOINT_HOST: "models.example.test",
      DEFAULT_MODEL: MODEL,
      VERSION: "0.1.0-alpha",
      COMMIT: "test",
      IS_OPENROUTER: false,
    }));
    const payload = await import("../src/core/payload");
    const s = setup();
    const built = payload.buildStructureOnly(s.input);
    const parsed = JSON.parse(built.outbound.body) as Record<string, unknown>;
    expect(Object.keys(parsed)).toEqual(["model", "messages", "max_tokens", "reasoning", "response_format"]);
    expect(built.outbound.body).not.toContain("provider");
  });
});
