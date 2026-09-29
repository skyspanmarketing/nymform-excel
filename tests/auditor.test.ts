// Auditor tests (spec §7.7, §10 T-A1..T-A10). Request bodies are built by hand in the exact
// format payload.ts produces, so most tests don't depend on payload.ts; a few end-to-end ones
// build the request with it.
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import ts from "typescript";
import { describe, expect, it } from "vitest";
import * as auditorModule from "../src/core/auditor";
import {
  AHO_MIN_TEXT,
  audit,
  buildAuditContext,
  chooseMatcher,
  findAll,
  INDEXOF_LIMIT,
  isAudited,
  type AuditContext,
  type AuditOutcome,
  type AuditRange,
  isRangeRendering,
  type ScanStrategy,
} from "../src/core/auditor";
import { parseLocalRange } from "../src/core/a1";
import { buildStructureOnly, buildSubstituted } from "../src/core/payload";
import { CORRECTION_PROMPT, SYSTEM_PROMPT } from "../src/core/prompt";
import { buildSheetContext, inferSchema } from "../src/core/schema";
import {
  createStandInMap,
  MONTH_RE,
  renderMonth,
  renderRange,
  substituteUserText,
  transformRows,
} from "../src/core/transform";
import type { AuditLocation, AuditResult, CellValue, ColType, HistoryMessage, Mode, RangeSpec, Treatment } from "../src/core/types";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const MODEL = "openai/gpt-6-luna";
const CREATED = "2026-09-24T12:00:00.000Z";

// ---------------------------------------------------------------------------------------------
// Fixtures

interface Col {
  letter: string;
  header: string;
  type: ColType;
  private: boolean;
  treatment: Treatment;
  values?: string[];
  aliasOriginal?: string;
  /** Header as sent; defaults to `header`. */
  sentHeader?: string;
  /** Note as sent; defaults to null. */
  note?: string;
  sentPrivate?: boolean;
  sentTreatment?: Treatment;
}

const ORDERS: Col[] = [
  {
    letter: "A",
    header: "Customer",
    type: "text",
    private: true,
    treatment: "stand_in",
    values: ["Maria Lopez", 'Ana "AJ" Ruiz', "Al", "Kenji Watanabe", "Bo"],
  },
  { letter: "B", header: "Region", type: "text", private: false, treatment: "as_is" },
  {
    letter: "C",
    header: "Phone",
    type: "text",
    private: true,
    treatment: "stand_in",
    values: ["(831) 555-0199", "(415) 555-0123"],
  },
  {
    letter: "D",
    header: "Amount",
    type: "number",
    private: true,
    treatment: "range",
    values: ["83500", "$83,500", "41250", "$41,250"],
  },
  {
    letter: "E",
    header: "Notes",
    type: "text",
    private: true,
    treatment: "exclude",
    values: ["reply about the formula"],
  },
  { letter: "F", header: "Signup", type: "date", private: true, treatment: "month", values: ["46096", "3/15/2026"] },
];

// Rows in layout A, B, C, D, F (E is excluded).
const ROWS: unknown[][] = [
  ["PERSON_001", "West", "TEXT_001", "80000-89999", "2026-03"],
  ["PERSON_002", "East", "TEXT_002", "40000-49999", "2026-01"],
  ["PERSON_003", "West", null, null, null],
  ["PERSON_004", "North", "TEXT_001", "0", "2025-12"],
  [null, "South", "TEXT_002", "-89999--80000", "2026-02"],
];

// Same columns, but the notes hold no word that appears in model replies.
const QUIET: Col[] = ORDERS.map((c) => (c.letter === "E" ? { ...c, values: ["called back twice"] } : c));

const TOKENS = new Set(["PERSON_001", "PERSON_002", "PERSON_003", "PERSON_004", "PERSON_005", "TEXT_001", "TEXT_002"]);

/** Data rows of a range with a header row, from its address ("A1:F6" has 5). */
function dataRowsOf(address: string): number {
  const rect = parseLocalRange(address)!;
  return rect.r2 - rect.r1;
}

function columnsJson(cols: readonly Col[], dataRows = 5): unknown[] {
  const distinct = Math.min(5, dataRows);
  return cols.map((c) => ({
    letter: c.letter,
    header: c.sentHeader ?? c.header,
    type: c.type,
    private: c.sentPrivate ?? c.private,
    treatment: c.sentTreatment ?? c.treatment,
    note: c.note ?? null,
    stats: c.type === "text" ? { blank: 0, distinct, avg_words: 2, max_len: 14 } : { blank: 0, distinct },
  }));
}

/** The range object payload.ts writes for an address with a header row. */
function rangeJson(address: string): Record<string, unknown> {
  const rect = parseLocalRange(address)!;
  return { address, header_row: rect.r1 + 1, first_data_row: rect.r1 + 2, last_data_row: rect.r2 + 1 };
}

interface DescOpts {
  mode?: "structure_only" | "substituted";
  cols?: readonly Col[];
  rows?: unknown[][];
  task?: string;
  sheet?: string;
  address?: string;
  contextRanges?: unknown[];
  extra?: Record<string, unknown>;
}

/** A nymform user turn, JSON.stringify'd with keys in the documented order. */
function desc(o: DescOpts = {}): string {
  const mode = o.mode ?? "structure_only";
  const sheet = o.sheet ?? "Orders";
  const address = o.address ?? "A1:F6";
  const dataRows = dataRowsOf(address);
  const d: Record<string, unknown> = {
    nymform: "0.1",
    mode,
    sheet,
    table: null,
    range: rangeJson(address),
    columns: columnsJson(o.cols ?? ORDERS, dataRows),
  };
  if (mode === "substituted" || o.rows) {
    const rows = o.rows ?? ROWS;
    d.rows = rows;
    d.rows_sent = rows.length;
    d.rows_total = dataRows;
  }
  d.context_ranges = o.contextRanges ?? [];
  d.allowed_ranges = [`${sheet}!${address}`];
  d.task = o.task ?? "Total Amount per Region in a new column";
  Object.assign(d, o.extra);
  return JSON.stringify(d);
}

type Msg = { role: string; content: string };

interface BodyOpts {
  user?: string;
  history?: Msg[];
  after?: Msg[];
  model?: string;
  system?: string;
  params?: Record<string, unknown>;
  provider?: unknown;
  extraTop?: Record<string, unknown>;
}

const REASONING_PARAMS = {
  max_tokens: 4000,
  reasoning: { effort: "low", exclude: true },
  response_format: { type: "json_object" },
};

/** `JSON.stringify({ model, messages, ...requestParams(model), provider })`, as payload.ts builds it. */
function makeBody(o: BodyOpts = {}): string {
  return JSON.stringify({
    model: o.model ?? MODEL,
    messages: [
      { role: "system", content: o.system ?? SYSTEM_PROMPT },
      ...(o.history ?? []),
      { role: "user", content: o.user ?? desc() },
      ...(o.after ?? []),
    ],
    ...(o.params ?? REASONING_PARAMS),
    provider: o.provider ?? { zdr: true },
    ...o.extraTop,
  });
}

function rangeOf(cols: readonly Col[], sheet = "Orders", address = "A1:F6"): AuditRange {
  return {
    sheet,
    address,
    hasHeaders: true,
    dataRows: dataRowsOf(address),
    columns: cols.map((c) => ({
      letter: c.letter,
      private: c.private,
      treatment: c.treatment,
      values: c.private ? [...(c.values ?? [])] : [],
      ...(c.aliasOriginal !== undefined ? { aliasOriginal: c.aliasOriginal } : {}),
    })),
  };
}

interface CtxOpts {
  cols?: readonly Col[];
  extraRanges?: AuditRange[];
  canaries?: string[];
  tokens?: Set<string>;
  isToken?: (s: string) => boolean;
  model?: string;
}

function makeCtx(o: CtxOpts = {}): AuditContext {
  const tokens = o.tokens ?? TOKENS;
  return {
    ranges: [rangeOf(o.cols ?? ORDERS), ...(o.extraRanges ?? [])],
    isToken: o.isToken ?? ((s: string) => tokens.has(s)),
    canaries: o.canaries ?? [],
    model: o.model ?? MODEL,
  };
}

function run(body: string, mode: Mode = "structure_only", ctx: AuditContext = makeCtx(), strategy?: ScanStrategy): AuditOutcome {
  const outbound = { body, bytes: Buffer.byteLength(body, "utf8"), mode, createdAt: CREATED };
  return strategy ? audit(outbound, ctx, { strategy }) : audit(outbound, ctx);
}

function blocked(r: AuditOutcome): AuditResult {
  if (r.ok) throw new Error("expected the audit to block, but it passed");
  expect(r.result.ok).toBe(false);
  return r.result;
}

function passed(r: AuditOutcome) {
  if (!r.ok) throw new Error(`expected a pass, got ${JSON.stringify(r.result)}`);
  return r.outbound;
}

// ---------------------------------------------------------------------------------------------
// Spec table T-A1..T-A10

describe("auditor: spec cases", () => {
  it("T-A1: structure-only body over a sheet with canaries passes", () => {
    const canaries = ["Quorbel Vantrisk", "Zentrofax Mirelle"];
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, values: [...(c.values ?? []), ...canaries] } : c));
    const body = makeBody({ user: desc({ cols }) });
    const out = passed(run(body, "structure_only", makeCtx({ cols, canaries })));
    expect(out.body).toBe(body);
    expect(out.audit).toEqual({ ok: true, hits: [], structural: [] });
  });

  it("T-A1: substituted body with stand-ins, ranges and months passes", () => {
    const body = makeBody({ user: desc({ mode: "substituted" }) });
    passed(run(body, "substituted"));
  });

  it("T-A1: a canary that reaches the body is caught", () => {
    const canaries = ["Quorbel Vantrisk"];
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, values: [...(c.values ?? []), ...canaries] } : c));
    const rows = ROWS.map((r) => [...r]);
    rows[1]![1] = "Quorbel Vantrisk"; // leaked into the non-private Region column
    const body = makeBody({ user: desc({ cols, mode: "substituted", rows }) });
    const result = blocked(run(body, "substituted", makeCtx({ cols, canaries })));
    expect(result.hits).toContainEqual({ column: "A", variant: "canary" });
  });

  it("T-A2: a private name injected into a body is blocked; the hit names the column, not the value", () => {
    const rows = ROWS.map((r) => [...r]);
    rows[2]![1] = "Maria Lopez"; // simulated bug: a private name in the non-private Region column
    const body = makeBody({ user: desc({ mode: "substituted", rows }) });
    const result = blocked(run(body, "substituted"));
    expect(result.hits).toContainEqual({ column: "A", variant: "exact" });
    for (const hit of result.hits) expect(Object.keys(hit).sort()).toEqual(["column", "variant"]);
    const serialized = JSON.stringify(result).toLowerCase();
    expect(serialized).not.toContain("maria");
    expect(serialized).not.toContain("lopez");
  });

  it("T-A2: the same injection in the question is blocked", () => {
    const body = makeBody({ user: desc({ task: "Total for Maria Lopez by month" }) });
    const result = blocked(run(body));
    expect(result.hits).toContainEqual({ column: "A", variant: "exact" });
    expect(JSON.stringify(result).toLowerCase()).not.toContain("maria");
  });

  it("T-A3: phone (831) 555-0199 appearing as 8315550199 is blocked", () => {
    const body = makeBody({ user: desc({ task: "Find the order for 8315550199" }) });
    const result = blocked(run(body));
    expect(result.hits).toEqual([{ column: "C", variant: "digits" }]);
    expect(result.error).toBe("Text from private column C is in the request. Nothing was sent.");
    expect(JSON.stringify(result)).not.toContain("8315550199");
  });

  it('T-A4: Ana "AJ" Ruiz JSON-escaped in a history message is blocked', () => {
    const history: Msg[] = [
      { role: "user", content: desc() },
      { role: "assistant", content: 'Here is the formula for Ana "AJ" Ruiz.' },
    ];
    const body = makeBody({ history });
    expect(body).toContain('Ana \\"AJ\\" Ruiz');
    const result = blocked(run(body));
    expect(result.hits).toContainEqual({ column: "A", variant: "json" });
    expect(JSON.stringify(result)).not.toContain("Ruiz");
  });

  it('T-A4: Ana "AJ" Ruiz in a nymform user turn (escaped twice) is blocked', () => {
    const body = makeBody({ user: desc({ task: 'Orders for Ana "AJ" Ruiz' }) });
    expect(body).toContain('Ana \\\\\\"AJ\\\\\\" Ruiz');
    const result = blocked(run(body));
    expect(result.hits).toContainEqual({ column: "A", variant: "json" });
  });

  it("T-A5: the first name alone from Maria Lopez is blocked (word variant)", () => {
    const body = makeBody({ user: desc({ task: "Total for Maria this year" }) });
    const result = blocked(run(body));
    expect(result.hits).toEqual([{ column: "A", variant: "word" }]);
  });

  it("T-A5: word variants need word boundaries", () => {
    // "Mariam" contains "maria" but not as a whole word.
    passed(run(makeBody({ user: desc({ task: "Total for Mariam this year" }) })));
  });

  it("T-A6: a private value Al in the question is NOT caught by the text scan (documented limit)", () => {
    // substituteUserText replaces it before the payload; the text scan skips values under 4 characters.
    const body = makeBody({ user: desc({ task: "Total for Al" }) });
    passed(run(body));
  });

  it("T-A7: a private value inside a history assistant turn is blocked", () => {
    const history: Msg[] = [
      { role: "user", content: desc() },
      { role: "assistant", content: '{"kind":"answer","formula":null,"placement":null,"explanation":"Kenji Watanabe has the most orders.","assumptions":[]}' },
    ];
    const result = blocked(run(makeBody({ history })));
    expect(result.hits).toContainEqual({ column: "A", variant: "exact" });
  });

  it("T-A7: a private value inside a history user turn is blocked", () => {
    const history: Msg[] = [
      { role: "user", content: desc({ task: "Orders for Kenji Watanabe" }) },
      { role: "assistant", content: '{"kind":"clarify","formula":null,"placement":null,"explanation":"Which month?","assumptions":[]}' },
    ];
    const result = blocked(run(makeBody({ history })));
    expect(result.hits).toContainEqual({ column: "A", variant: "exact" });
  });

  it("T-A8: the auditor throwing blocks the send (isToken throws)", () => {
    const ctx = makeCtx({
      isToken: () => {
        throw new Error("Maria Lopez");
      },
    });
    const r = run(makeBody({ user: desc({ mode: "substituted" }) }), "substituted", ctx);
    const result = blocked(r);
    expect(result.error).toBeTruthy();
    expect(result.error).not.toContain("Maria");
    expect("outbound" in r).toBe(false);
  });

  it("T-A8: a context that throws while being read blocks the send", () => {
    const ctx = makeCtx();
    Object.defineProperty(ctx, "ranges", {
      get() {
        throw new Error("boom");
      },
    });
    expect(blocked(run(makeBody(), "structure_only", ctx)).error).toBeTruthy();
  });

  it("T-A8: bad inputs never throw", () => {
    const bad: unknown[] = [null, undefined, 42, "x", {}, { body: 5 }, { body: makeBody(), mode: "structure_only" }];
    for (const outbound of bad) {
      let r: AuditOutcome | undefined;
      expect(() => {
        r = audit(outbound as never, makeCtx());
      }).not.toThrow();
      expect(r?.ok).toBe(false);
    }
    expect(audit({ body: makeBody(), bytes: 1, mode: "structure_only", createdAt: CREATED }, null as never).ok).toBe(false);
  });

  it("T-A9: a raw Al in a private column's row cell is blocked by the structural check", () => {
    const rows = ROWS.map((r) => [...r]);
    rows[0]![0] = "Al";
    const result = blocked(run(makeBody({ user: desc({ mode: "substituted", rows }) }), "substituted"));
    expect(result.structural).toEqual([{ column: "A" }]);
    expect(result.hits).toEqual([]);
    expect(result.error).toBe("A private cell in column A wasn't replaced with a stand-in. Nothing was sent.");
    expect(JSON.stringify(result)).not.toContain('"Al"');
  });

  it("T-A10: code outside the auditor can't construct an AuditedOutbound (type-check)", () => {
    const file = resolve(ROOT, "tests/auditedOutbound.typecheck.ts");
    const source = readFileSync(file, "utf8");
    const lines = source.split("\n");
    const directiveLines = lines.flatMap((l, i) => (/^\s*\/\/ @ts-expect-error/.test(l) ? [i] : []));
    expect(directiveLines.length).toBeGreaterThanOrEqual(3);

    const parsed = ts.getParsedCommandLineOfConfigFile(resolve(ROOT, "tsconfig.json"), {}, {
      ...ts.sys,
      onUnRecoverableConfigFileDiagnostic: () => undefined,
    });
    if (!parsed) throw new Error("tsconfig.json could not be read");
    const options: ts.CompilerOptions = { ...parsed.options, noEmit: true };

    const diagnosticsFor = (text: string): readonly ts.Diagnostic[] => {
      const host = ts.createCompilerHost(options);
      const original = host.getSourceFile.bind(host);
      host.getSourceFile = (name, languageVersion, onError, shouldCreate) =>
        resolve(name) === file
          ? ts.createSourceFile(name, text, languageVersion, true)
          : original(name, languageVersion, onError, shouldCreate);
      const program = ts.createProgram({ rootNames: [file], options, host });
      const sf = program.getSourceFile(file);
      if (!sf) throw new Error("typecheck file not loaded");
      return [...program.getSyntacticDiagnostics(sf), ...program.getSemanticDiagnostics(sf)];
    };
    const format = (d: ts.Diagnostic) => ts.flattenDiagnosticMessageText(d.messageText, "\n");

    // As written, every @ts-expect-error is used and nothing else fails.
    expect(diagnosticsFor(source).map(format)).toEqual([]);

    // Without the directives, each marked line fails to type-check.
    const stripped = lines.map((l, i) => (directiveLines.includes(i) ? "//" : l)).join("\n");
    const errorLines = new Set(
      diagnosticsFor(stripped).flatMap((d) =>
        d.file && d.start !== undefined ? [d.file.getLineAndCharacterOfPosition(d.start).line] : [],
      ),
    );
    for (const i of directiveLines) expect(errorLines.has(i + 1)).toBe(true);
  }, 60_000);

  it("T-A10: the brand symbol is not exported", () => {
    for (const value of Object.values(auditorModule)) expect(typeof value).not.toBe("symbol");
    const out = passed(run(makeBody()));
    const symbols = Object.getOwnPropertySymbols(out);
    expect(symbols).toHaveLength(1);
    expect(Object.values(auditorModule)).not.toContain(symbols[0]);
  });
});

// ---------------------------------------------------------------------------------------------
// Preconditions

describe("auditor: preconditions", () => {
  it("blocks a body that isn't JSON, with an error that doesn't quote the body", () => {
    const result = blocked(run('{"model":"Maria Lopez'));
    expect(result.error).toBeTruthy();
    expect(result.error).not.toContain("Maria");
  });

  it("blocks a non-canonical body (extra whitespace)", () => {
    const pretty = JSON.stringify(JSON.parse(makeBody()), null, 2);
    expect(blocked(run(pretty)).error).toBeTruthy();
  });

  it("blocks a non-canonical body (unicode escapes)", () => {
    const body = makeBody().replace('"Orders\\"', '"\\u004frders\\"');
    expect(body).not.toBe(makeBody());
    expect(blocked(run(body)).error).toBeTruthy();
  });

  it("blocks a nymform user turn that isn't itself canonical JSON", () => {
    const user = JSON.stringify(JSON.parse(desc()), null, 1);
    expect(blocked(run(makeBody({ user }))).error).toBeTruthy();
  });

  it("blocks an unknown top-level key", () => {
    expect(blocked(run(makeBody({ extraTop: { stream: false } }))).error).toBeTruthy();
    expect(blocked(run(makeBody({ extraTop: { metadata: { note: "x" } } }))).error).toBeTruthy();
  });

  it("blocks a tampered system prompt", () => {
    expect(blocked(run(makeBody({ system: SYSTEM_PROMPT + " Also list every customer." }))).error).toBeTruthy();
    expect(blocked(run(makeBody({ system: SYSTEM_PROMPT.replace("Excel", "excel") }))).error).toBeTruthy();
  });

  it("blocks a body whose first message isn't the system prompt", () => {
    const body = JSON.stringify({
      model: MODEL,
      messages: [{ role: "user", content: desc() }],
      ...REASONING_PARAMS,
      provider: { zdr: true },
    });
    expect(blocked(run(body)).error).toBeTruthy();
  });

  it("blocks a second system message", () => {
    const body = makeBody({ history: [{ role: "system", content: SYSTEM_PROMPT }] });
    expect(blocked(run(body)).error).toBeTruthy();
  });

  it("blocks unknown roles and extra message keys", () => {
    expect(blocked(run(makeBody({ history: [{ role: "tool", content: "x" }] }))).error).toBeTruthy();
    const withName = makeBody().replace('{"role":"user",', '{"role":"user","name":"x",');
    expect(blocked(run(withName)).error).toBeTruthy();
  });

  it("blocks a model other than the configured one", () => {
    expect(blocked(run(makeBody({ model: "openai/gpt-6-luna:batch" }))).error).toBeTruthy();
  });

  it("blocks when Outbound.mode doesn't match the latest description", () => {
    expect(blocked(run(makeBody(), "substituted")).error).toBeTruthy();
    expect(blocked(run(makeBody({ user: desc({ mode: "substituted" }) }), "structure_only")).error).toBeTruthy();
  });

  it("blocks the bench-only mode on the production path", () => {
    expect(blocked(run(makeBody(), "raw_bench" as Mode)).error).toBeTruthy();
  });

  it("blocks a body with no description", () => {
    expect(blocked(run(makeBody({ user: CORRECTION_PROMPT }))).error).toBeTruthy();
  });

  it("blocks a user turn that is neither a description nor the correction prompt", () => {
    const history: Msg[] = [
      { role: "user", content: "hello" },
      { role: "assistant", content: "hi" },
    ];
    expect(blocked(run(makeBody({ history }))).error).toBeTruthy();
  });

  it("blocks unexpected request parameters under the masked keys", () => {
    expect(blocked(run(makeBody({ provider: { zdr: true, order: ["Maria Lopez"] } }))).error).toBeTruthy();
    expect(blocked(run(makeBody({ params: { ...REASONING_PARAMS, response_format: { type: "json_object", schema: {} } } }))).error).toBeTruthy();
    expect(blocked(run(makeBody({ params: { ...REASONING_PARAMS, reasoning: { effort: "Maria" } } }))).error).toBeTruthy();
    expect(blocked(run(makeBody({ params: { temperature: "0", max_tokens: 800 } }))).error).toBeTruthy();
  });

  it("passes the non-reasoning parameter set and a body without provider", () => {
    const params = { temperature: 0, max_tokens: 800, response_format: { type: "json_object" } };
    passed(run(makeBody({ params })));
    const noProvider = JSON.stringify({
      model: MODEL,
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user", content: desc() },
      ],
      ...params,
    });
    passed(run(noProvider));
  });

  it("passes the correction retry: [system, user, assistant(invalid), user(CORRECTION_PROMPT)]", () => {
    const body = makeBody({
      after: [
        { role: "assistant", content: "Sure! Here it is: =SUM(D2:D6)" },
        { role: "user", content: CORRECTION_PROMPT },
      ],
    });
    passed(run(body));
  });

  it("passes a clean history exchange", () => {
    const history: Msg[] = [
      { role: "user", content: desc({ task: "Count orders per Region" }) },
      { role: "assistant", content: '{"kind":"formula","formula":"=COUNTIFS(B2:B6,B2)","placement":{"cell":"G2","fill_down":true},"explanation":"Counts orders.","assumptions":[]}' },
    ];
    passed(run(makeBody({ history }), "structure_only", makeCtx({ cols: QUIET })));
  });

  it("a private note containing the word 'formula' doesn't block a follow-up because a reply uses that key", () => {
    const history: Msg[] = [
      { role: "user", content: desc({ task: "Count orders per Region" }) },
      { role: "assistant", content: '{"kind":"formula","formula":"=COUNTIFS(B2:B6,B2)","placement":null,"explanation":"Counts orders.","assumptions":[]}' },
    ];
    passed(run(makeBody({ history })));
  });

  it("the same note word in a reply's explanation still blocks", () => {
    const history: Msg[] = [
      { role: "user", content: desc({ task: "Count orders per Region" }) },
      { role: "assistant", content: '{"kind":"formula","formula":"=COUNTIFS(B2:B6,B2)","placement":null,"explanation":"This formula counts orders.","assumptions":[]}' },
    ];
    expect(blocked(run(makeBody({ history }))).hits).toEqual([{ column: "E", variant: "word" }]);
  });

  it("blocks an unknown key in a description", () => {
    expect(blocked(run(makeBody({ user: desc({ extra: { samples: [["Al"]] } }) }))).error).toBeTruthy();
  });

  it("blocks a description of a range that isn't in the context", () => {
    expect(blocked(run(makeBody({ user: desc({ address: "A1:F7" }) }))).error).toBeTruthy();
    expect(blocked(run(makeBody({ user: desc({ sheet: "Other" }) }))).error).toBeTruthy();
  });
});

// ---------------------------------------------------------------------------------------------
// Structural check

describe("auditor: structural check", () => {
  it("blocks a structure-only body that contains rows", () => {
    const body = makeBody({ user: desc({ mode: "structure_only", rows: ROWS }) });
    expect(blocked(run(body, "structure_only")).error).toBeTruthy();
  });

  it("blocks a structure-only request whose history holds rows", () => {
    const history: Msg[] = [
      { role: "user", content: desc({ mode: "substituted" }) },
      { role: "assistant", content: '{"kind":"clarify","formula":null,"placement":null,"explanation":"Which?","assumptions":[]}' },
    ];
    expect(blocked(run(makeBody({ history }), "structure_only")).error).toBeTruthy();
  });

  it("blocks a row with the wrong length", () => {
    const rows = ROWS.map((r) => [...r]);
    rows[3] = [...rows[3]!, "extra"];
    expect(blocked(run(makeBody({ user: desc({ mode: "substituted", rows }) }), "substituted")).error).toBeTruthy();
    const short = ROWS.map((r) => r.slice(0, 4));
    expect(blocked(run(makeBody({ user: desc({ mode: "substituted", rows: short }) }), "substituted")).error).toBeTruthy();
  });

  it("blocks rows that include the excluded column", () => {
    const rows = ROWS.map((r) => [...r.slice(0, 4), "TEXT_001", r[4]]);
    expect(blocked(run(makeBody({ user: desc({ mode: "substituted", rows }) }), "substituted")).error).toBeTruthy();
  });

  it("blocks rows that aren't arrays of plain values", () => {
    const rows: unknown[][] = ROWS.map((r) => [...r]);
    rows[0]![1] = { v: "West" };
    expect(blocked(run(makeBody({ user: desc({ mode: "substituted", rows }) }), "substituted")).ok).toBe(false);
  });

  it("blocks columns whose private flags disagree with the context", () => {
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, sentPrivate: false } : c));
    const result = blocked(run(makeBody({ user: desc({ cols }) })));
    expect(result.error).toBeTruthy();
    expect(result.error).not.toContain("Maria");
  });

  it("blocks columns whose treatments disagree with the context", () => {
    const cols = ORDERS.map((c) => (c.letter === "D" ? { ...c, sentTreatment: "as_is" as Treatment } : c));
    expect(blocked(run(makeBody({ user: desc({ cols }) }))).error).toBeTruthy();
  });

  it("blocks columns whose letters disagree with the context", () => {
    const renamed = ORDERS.map((c) => (c.letter === "B" ? { ...c, letter: "Z" } : c));
    expect(blocked(run(makeBody({ user: desc({ cols: renamed }) }))).error).toBeTruthy();
    const missing = ORDERS.filter((c) => c.letter !== "E");
    expect(blocked(run(makeBody({ user: desc({ cols: missing }) }))).error).toBeTruthy();
  });

  it("uses the context's layout, not the payload's, for rows", () => {
    // The payload claims A is excluded; the context says stand_in. The column check blocks it.
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, sentTreatment: "exclude" as Treatment } : c));
    const rows = ROWS.map((r) => r.slice(1));
    expect(blocked(run(makeBody({ user: desc({ mode: "substituted", cols, rows }) }), "substituted")).ok).toBe(false);
  });

  it("treats a private column with treatment as_is as stand_in", () => {
    const ctxCols = ORDERS.map((c) => (c.letter === "A" ? { ...c, treatment: "as_is" as Treatment } : c));
    const body = makeBody({ user: desc({ mode: "substituted" }) }); // sends A as stand_in
    passed(run(body, "substituted", makeCtx({ cols: ctxCols })));
  });

  it("blocks a token the session map didn't allocate", () => {
    const rows = ROWS.map((r) => [...r]);
    rows[0]![0] = "PERSON_999";
    const result = blocked(run(makeBody({ user: desc({ mode: "substituted", rows }) }), "substituted"));
    expect(result.structural).toEqual([{ column: "A" }]);
    expect(result.hits).toEqual([]);
  });

  it("blocks raw numbers and booleans in private columns", () => {
    const rows = ROWS.map((r) => [...r]);
    rows[0]![3] = 83500;
    rows[1]![0] = true;
    const result = blocked(run(makeBody({ user: desc({ mode: "substituted", rows }) }), "substituted"));
    // 83500 also appears as a number, which the text scan finds.
    expect(result.structural).toEqual([{ column: "A" }, { column: "D" }]);
    expect(result.hits).toEqual([{ column: "D", variant: "exact" }]);
  });

  it("accepts range renderings and tokens in a range column, and rejects look-alikes", () => {
    for (const ok of ["0", "0-9", "10-19", "80000-89999", "-89999--80000", "-9--0", "1000000-1999999", "TEXT_001", null]) {
      const rows = ROWS.map((r) => [...r]);
      rows[0]![3] = ok;
      passed(run(makeBody({ user: desc({ mode: "substituted", rows }) }), "substituted"));
    }
    for (const bad of ["555-0199", "1990-2000", "80000-89998", "83500", "01-19", "2026-03", "Al", ""]) {
      const rows = ROWS.map((r) => [...r]);
      rows[0]![3] = bad;
      const result = blocked(run(makeBody({ user: desc({ mode: "substituted", rows }) }), "substituted"));
      expect(result.structural).toContainEqual({ column: "D" });
    }
  });

  it("accepts month renderings in a month column and rejects others", () => {
    const rows = ROWS.map((r) => [...r]);
    rows[0]![4] = "2026-13";
    rows[1]![4] = "15/03/26";
    const result = blocked(run(makeBody({ user: desc({ mode: "substituted", rows }) }), "substituted"));
    expect(result.structural).toEqual([{ column: "F" }]);
    const monthInStandIn = ROWS.map((r) => [...r]);
    monthInStandIn[0]![0] = "2026-03";
    expect(blocked(run(makeBody({ user: desc({ mode: "substituted", rows: monthInStandIn }) }), "substituted")).structural).toEqual([
      { column: "A" },
    ]);
  });

  it("checks context ranges against their own columns", () => {
    const regionCols: Col[] = [
      { letter: "A", header: "Region", type: "text", private: false, treatment: "as_is" },
      { letter: "B", header: "Manager", type: "text", private: true, treatment: "stand_in", values: ["Jo Park", "Lee"] },
    ];
    const regions = rangeOf(regionCols, "Regions", "A1:B5");
    const ctx = makeCtx({ extraRanges: [regions] });
    const contextRange = (rows: unknown[][]) => ({
      sheet: "Regions",
      table: null,
      range: { address: "A1:B5", header_row: 1, first_data_row: 2, last_data_row: 5 },
      columns: columnsJson(regionCols, 4),
      rows,
      rows_sent: rows.length,
      rows_total: 4,
    });
    const good = [["West", "PERSON_005"], ["East", null]];
    passed(run(makeBody({ user: desc({ mode: "substituted", contextRanges: [contextRange(good)] }) }), "substituted", ctx));

    const bad = [["West", "Lee"]];
    const result = blocked(
      run(makeBody({ user: desc({ mode: "substituted", contextRanges: [contextRange(bad)] }) }), "substituted", ctx),
    );
    expect(result.structural).toEqual([{ column: "Regions!B" }]);

    // A context range the user didn't add is refused.
    const unknown = { ...contextRange(good), sheet: "Payroll" };
    expect(blocked(run(makeBody({ user: desc({ mode: "substituted", contextRanges: [unknown] }) }), "substituted", ctx)).error).toBeTruthy();
  });

  it("reports structural hits and text hits together, sorted and deduped", () => {
    const rows = ROWS.map((r) => [...r]);
    rows[0]![0] = "Al";
    rows[1]![0] = "Bo";
    const body = makeBody({ user: desc({ mode: "substituted", rows, task: "Totals for Kenji Watanabe and Kenji Watanabe" }) });
    const result = blocked(run(body, "substituted"));
    expect(result.structural).toEqual([{ column: "A" }]);
    expect(result.hits).toEqual([
      { column: "A", variant: "exact" },
      { column: "A", variant: "word" },
    ]);
  });
});

// ---------------------------------------------------------------------------------------------
// Text scan and the constant mask

describe("auditor: text scan and constant mask", () => {
  it("does not block a private notes value made of words from the system prompt", () => {
    // "reply", "about" and "formula" all occur in SYSTEM_PROMPT, which is masked.
    for (const w of ["reply", "about", "formula"]) expect(SYSTEM_PROMPT.toLowerCase()).toContain(w);
    passed(run(makeBody()));
  });

  it("blocks when the same word appears in the task", () => {
    const result = blocked(run(makeBody({ user: desc({ task: "Write a formula for the totals" }) })));
    expect(result.hits).toEqual([{ column: "E", variant: "word" }]);
  });

  it("does not block a private boolean-ish value 'true' because of JSON literals", () => {
    const cols: Col[] = [
      ...ORDERS,
      { letter: "G", header: "Active", type: "boolean", private: true, treatment: "stand_in", values: ["true", "TRUE", "false", "FALSE"] },
      { letter: "H", header: "Paid", type: "boolean", private: false, treatment: "as_is" },
    ];
    const rows = ROWS.map((r, i) => [...r, i % 2 === 0 ? "TEXT_001" : null, i % 2 === 0]);
    const ctx = makeCtx({ cols });
    const body = makeBody({ user: desc({ cols, mode: "substituted", rows, address: "A1:H6" }) });
    const ctxH = { ...ctx, ranges: [{ ...ctx.ranges[0]!, address: "A1:H6" }] };
    expect(body).toContain(":true");
    expect(body).toContain(",true]");
    passed(run(body, "substituted", ctxH));
    const asked = makeBody({ user: desc({ cols, mode: "substituted", rows, address: "A1:H6", task: "Is it true that B2 is West?" }) });
    expect(blocked(run(asked, "substituted", ctxH)).hits).toEqual([{ column: "G", variant: "exact" }]);
  });

  it("does not block a private client named Luna because of the model ID, but does in the task", () => {
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, values: ["Luna", "Maria Lopez"] } : c));
    const ctx = makeCtx({ cols });
    passed(run(makeBody({ user: desc({ cols }) }), "structure_only", ctx));
    const result = blocked(run(makeBody({ user: desc({ cols, task: "Orders for Luna" }) }), "structure_only", ctx));
    expect(result.hits).toEqual([{ column: "A", variant: "exact" }]);
  });

  it("finds a word right after an escaped newline in a nymform user turn", () => {
    const body = makeBody({ user: desc({ task: "names:\nMaria and others" }) });
    expect(body).toContain("\\\\nMaria");
    expect(blocked(run(body)).hits).toEqual([{ column: "A", variant: "word" }]);
  });

  it("finds a word right after an escaped newline or tab in an assistant turn", () => {
    for (const sep of ["\n", "\t", "\r", '"', "\\"]) {
      const history: Msg[] = [
        { role: "user", content: desc() },
        { role: "assistant", content: `names:${sep}Lopez${sep}and others` },
      ];
      expect(blocked(run(makeBody({ history }))).hits).toEqual([{ column: "A", variant: "word" }]);
    }
    // A control character that isn't whitespace can't be seen: it is skipped, not a word break
    // (eighth review).
    const history: Msg[] = [
      { role: "user", content: desc() },
      { role: "assistant", content: "names:\u0001Lo\u0001pez\u0001 and others" },
    ];
    expect(blocked(run(makeBody({ history }))).hits).toEqual([{ column: "A", variant: "word" }]);
  });

  it("finds normalized variants (case, spacing)", () => {
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, values: ["  Maria   Lopez "] } : c));
    const result = blocked(run(makeBody({ user: desc({ cols, task: "Totals for MARIA LOPEZ" }) }), "structure_only", makeCtx({ cols })));
    expect(result.hits).toContainEqual({ column: "A", variant: "normalized" });
  });

  it("finds a private value in a header, note or sheet name", () => {
    const cols = ORDERS.map((c) => (c.letter === "B" ? { ...c, sentHeader: "Kenji Watanabe" } : c));
    expect(blocked(run(makeBody({ user: desc({ cols }) }))).hits).toContainEqual({ column: "A", variant: "exact" });
  });

  it("blocks the original header of an aliased column", () => {
    const cols: Col[] = [
      ...ORDERS.slice(0, 5),
      { letter: "F", header: "Diagnosis", sentHeader: "Field F", type: "text", private: true, treatment: "exclude", values: [], aliasOriginal: "Diagnosis" },
    ];
    const ctx = makeCtx({ cols });
    passed(run(makeBody({ user: desc({ cols }) }), "structure_only", ctx));
    const leaked = cols.map((c) => (c.letter === "F" ? { ...c, sentHeader: "Diagnosis" } : c));
    expect(blocked(run(makeBody({ user: desc({ cols: leaked }) }), "structure_only", ctx)).hits).toEqual([
      { column: "F", variant: "exact" },
    ]);
  });

  it("masks enum values only when they are allowed members", () => {
    // A private value "number" isn't found in "type":"number", but is found in the task.
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, values: ["number", "stand_in"] } : c));
    const ctx = makeCtx({ cols });
    passed(run(makeBody({ user: desc({ cols }) }), "structure_only", ctx));
    expect(blocked(run(makeBody({ user: desc({ cols, task: "the number of orders" }) }), "structure_only", ctx)).hits).toEqual([
      { column: "A", variant: "exact" },
    ]);
  });

  it("scans numbers: a private 4+ digit value that appears as a number blocks", () => {
    // Only checked row numbers and counts are masked; any other number is scanned.
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, values: ["12345"] } : c));
    const rows = ROWS.map((r) => [...r]);
    rows[0]![1] = 12345;
    const body = makeBody({ user: desc({ mode: "substituted", cols, rows }) });
    expect(body).toContain(",12345,");
    expect(blocked(run(body, "substituted", makeCtx({ cols }))).hits).toEqual([{ column: "A", variant: "exact" }]);
    // A count too large for the selection isn't masked: the request blocks before the scan.
    const d = JSON.parse(desc({ cols })) as { columns: { stats: { distinct: number } }[] };
    d.columns[1]!.stats.distinct = 12345;
    const result = blocked(run(makeBody({ user: JSON.stringify(d) }), "structure_only", makeCtx({ cols })));
    expect(result.hits).toEqual([]);
    expect(result.error).toBe("A description in the request has row numbers or counts that don't match your selection.");
  });

  it("a private amount equal to its bin's lower bound doesn't block: a checked rendering is masked", () => {
    // 80000 renders as "80000-89999", which contains the private value itself. The structural check
    // accepted that cell as a rendering, so it is replacement text Nymform wrote, not the value.
    const cols = ORDERS.map((c) => (c.letter === "D" ? { ...c, values: ["80000", "$80,000"] } : c));
    const rows = ROWS.map((r) => [...r]);
    rows[0]![3] = "80000-89999";
    passed(run(makeBody({ user: desc({ mode: "substituted", rows }) }), "substituted", makeCtx({ cols })));
    // The same value anywhere else is still found.
    const inTask = makeBody({ user: desc({ mode: "substituted", rows, task: "rows over 80000" }) });
    expect(blocked(run(inTask, "substituted", makeCtx({ cols }))).hits).toEqual([{ column: "D", variant: "exact" }]);
    // And a rendering in a non-private column is scanned like any other text.
    const open = ORDERS.map((c) => (c.letter === "D" ? { ...c, private: false, treatment: "as_is" as const, values: [] } : c));
    const leakCols = open.map((c) => (c.letter === "A" ? { ...c, values: [...(c.values ?? []), "80000"] } : c));
    expect(blocked(run(makeBody({ user: desc({ mode: "substituted", cols: leakCols, rows }) }), "substituted", makeCtx({ cols: leakCols }))).hits).toContainEqual({ column: "A", variant: "exact" });
  });

  it("stand-in tokens in the task are masked, so a private word inside a token doesn't block", () => {
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, values: [...(c.values ?? []), "Person"] } : c));
    passed(run(makeBody({ user: desc({ task: "Total for PERSON_001" }) }), "structure_only", makeCtx({ cols })));
    // A string that only looks like a token, but the session didn't allocate, is scanned.
    const notToken = makeBody({ user: desc({ task: "Total for PERSON_999 person" }) });
    expect(blocked(run(notToken, "structure_only", makeCtx({ cols }))).hits.length).toBeGreaterThan(0);
  });

  it("skips variants shorter than 4 characters and canaries shorter than 6", () => {
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, values: ["Bo", "Al", "Wu"] } : c));
    passed(run(makeBody({ user: desc({ cols, task: "Bo and Al Wu" }) }), "structure_only", makeCtx({ cols, canaries: ["Regio"] })));
  });

  it("matches canaries case-insensitively", () => {
    const body = makeBody({ user: desc({ task: "see QUORBEL VANTRISK" }) });
    expect(blocked(run(body, "structure_only", makeCtx({ canaries: ["Quorbel Vantrisk"] }))).hits).toEqual([
      { column: "*", variant: "canary" },
    ]);
  });

  it("uses a length-preserving lowercase so positions line up", () => {
    // U+0130 lowercases to two UTF-16 units; the scan must still find values after it.
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, values: ["İstanbul Kaya"] } : c));
    const ctx = makeCtx({ cols });
    expect(blocked(run(makeBody({ user: desc({ cols, task: "İİİ Kaya" }) }), "structure_only", ctx)).hits).toEqual([
      { column: "A", variant: "word" },
    ]);
    expect(blocked(run(makeBody({ user: desc({ cols, task: "from İstanbul" }) }), "structure_only", ctx)).hits).toEqual([
      { column: "A", variant: "word" },
    ]);
  });

  it("finds words with non-ASCII letters on correct boundaries", () => {
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, values: ["José Núñez"] } : c));
    const ctx = makeCtx({ cols });
    expect(blocked(run(makeBody({ user: desc({ cols, task: "for Núñez" }) }), "structure_only", ctx)).hits).toEqual([
      { column: "A", variant: "word" },
    ]);
    passed(run(makeBody({ user: desc({ cols, task: "for ÁNúñezÉ" }) }), "structure_only", ctx));
  });
});

// ---------------------------------------------------------------------------------------------
// The audited result

describe("auditor: AuditedOutbound", () => {
  it("is frozen and carries the identical body string", () => {
    const body = makeBody({ user: desc({ mode: "substituted" }) });
    const outbound = { body, bytes: Buffer.byteLength(body, "utf8"), mode: "substituted" as const, createdAt: CREATED };
    const r = audit(outbound, makeCtx());
    const out = passed(r);
    expect(out.body).toBe(body);
    expect(out.bytes).toBe(outbound.bytes);
    expect(out.mode).toBe("substituted");
    expect(out.createdAt).toBe(CREATED);
    expect(out.audit).toEqual({ ok: true, hits: [], structural: [] });
    expect(Object.isFrozen(out)).toBe(true);
    expect(Object.isFrozen(out.audit)).toBe(true);
    expect(Object.isFrozen(out.audit.hits)).toBe(true);
    expect(out).not.toBe(outbound);
    expect(() => {
      (out as { body: string }).body = "{}";
    }).toThrow();
  });

  it("returns the body it checked, even if the input's body changes when read again", () => {
    const clean = makeBody();
    const dirty = makeBody({ user: desc({ task: "Orders for Maria Lopez" }) });
    let reads = 0;
    const outbound = {
      get body() {
        reads += 1;
        return reads === 1 ? clean : dirty;
      },
      bytes: 1,
      mode: "structure_only" as const,
      createdAt: CREATED,
    };
    const out = passed(audit(outbound, makeCtx()));
    expect(out.body).toBe(clean);
  });
});

// ---------------------------------------------------------------------------------------------
// Matchers: indexOf and Aho-Corasick must agree

function mulberry32(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function pick<T>(rand: () => number, items: readonly T[]): T {
  return items[Math.floor(rand() * items.length)]!;
}

function randomString(rand: () => number, alphabet: readonly string[], min: number, max: number): string {
  const n = min + Math.floor(rand() * (max - min + 1));
  let s = "";
  for (let i = 0; i < n; i++) s += pick(rand, alphabet);
  return s;
}

describe("auditor: matchers", () => {
  it("indexOf and Aho-Corasick report identical occurrences on a random corpus", () => {
    const rand = mulberry32(42);
    const alphabet = ["a", "b", "c", "n", "\\", '"', " ", "é", "İ", "😀"];
    for (let t = 0; t < 400; t++) {
      const hay = randomString(rand, alphabet, 0, 300);
      const count = 1 + Math.floor(rand() * 40);
      const patterns: string[] = [];
      for (let i = 0; i < count; i++) patterns.push(randomString(rand, alphabet, 1, 6));
      if (hay.length > 10 && rand() < 0.5) patterns.push(hay.slice(3, 9));
      const collect = (strategy: ScanStrategy) => {
        const out: string[] = [];
        findAll(hay, patterns, strategy, (p, s) => {
          out.push(`${p}:${s}`);
        });
        return out.sort();
      };
      const viaIndexOf = collect("indexOf");
      expect(collect("ahoCorasick")).toEqual(viaIndexOf);
      // Brute force agrees too.
      const brute: string[] = [];
      patterns.forEach((p, pi) => {
        for (let s = 0; s + p.length <= hay.length; s++) if (hay.startsWith(p, s)) brute.push(`${pi}:${s}`);
      });
      expect(viaIndexOf).toEqual(brute.sort());
    }
  });

  it("audit gives identical results with either matcher on random bodies", () => {
    const rand = mulberry32(7);
    const vocab = ["Maria", "Lopez", "Kenji", "Watanabe", "Ruiz", "formula", "reply", "true", "Luna", "West", "8315550199", "abcd", "abcde", "Ab", "région", "Østby"];
    const seps = [" ", "\n", "\t", '"', "\\", "-", "_", ".", ""];
    const phrase = (min: number, max: number) => {
      const n = min + Math.floor(rand() * (max - min + 1));
      const parts: string[] = [];
      for (let i = 0; i < n; i++) parts.push(pick(rand, vocab) + pick(rand, seps));
      return parts.join("");
    };
    let passes = 0;
    let blocks = 0;
    for (let t = 0; t < 200; t++) {
      const cols = ORDERS.map((c) =>
        c.private ? { ...c, values: Array.from({ length: 1 + Math.floor(rand() * 4) }, () => phrase(1, 3)) } : c,
      );
      const history: Msg[] =
        rand() < 0.5 ? [{ role: "user", content: desc({ cols }) }, { role: "assistant", content: phrase(0, 6) }] : [];
      const body = makeBody({ history, user: desc({ cols, task: phrase(0, 5) }) });
      const ctx = makeCtx({ cols, canaries: rand() < 0.3 ? [phrase(2, 2)] : [] });
      const a = run(body, "structure_only", ctx, "indexOf");
      const b = run(body, "structure_only", ctx, "ahoCorasick");
      const c = run(body, "structure_only", ctx, "auto");
      const view = (r: AuditOutcome) => (r.ok ? "ok" : r.result);
      expect(view(b)).toEqual(view(a));
      expect(view(c)).toEqual(view(a));
      if (a.ok) passes++;
      else blocks++;
    }
    expect(passes).toBeGreaterThan(10);
    expect(blocks).toBeGreaterThan(10);
  });

  it(`switches to Aho-Corasick above ${INDEXOF_LIMIT} variants and still finds hits`, () => {
    expect(INDEXOF_LIMIT).toBe(5000);
    const values = Array.from({ length: 3000 }, (_, i) => `Givenname${i} Familyname${i}`);
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, values } : c));
    const ctx = makeCtx({ cols });
    passed(run(makeBody({ user: desc({ cols }) }), "structure_only", ctx, "auto"));
    const body = makeBody({ user: desc({ cols, task: "orders for Familyname2718" }) });
    const auto = blocked(run(body, "structure_only", ctx, "auto"));
    expect(auto.hits).toEqual([{ column: "A", variant: "word" }]);
    expect(run(body, "structure_only", ctx, "indexOf")).toEqual(run(body, "structure_only", ctx, "auto"));
  });
});

// ---------------------------------------------------------------------------------------------
// buildAuditContext

function specFixture(): RangeSpec {
  const text = [
    ["Customer", "Region", "Amount", "Diagnosis", "Active"],
    ["Maria Lopez", "West", "$83,500", "Flu", "TRUE"],
    ["Kenji Watanabe", "East", "$41,250", "", "FALSE"],
    ["", "West", "", "Cold", ""],
  ];
  const values = [
    ["Customer", "Region", "Amount", "Diagnosis", "Active"],
    ["Maria Lopez", "West", 83500, "Flu", true],
    ["Kenji Watanabe", "East", 41250, "", false],
    [null, "West", null, "Cold", null],
  ];
  return {
    data: {
      sheet: "Orders",
      address: "A1:E4",
      rowIndex: 0,
      columnIndex: 0,
      values,
      text,
      formulas: values,
      valueTypes: values.map((r) => r.map((v) => (v === null || v === "" ? "Empty" : typeof v === "number" ? "Double" : typeof v === "boolean" ? "Boolean" : "String"))),
      numberFormat: values.map((r) => r.map(() => "General")),
      table: null,
    },
    hasHeaders: true,
    columns: [
      { letter: "A", header: "Customer", type: "text", stats: { blank: 1, distinct: 2, avgWords: 2, maxLen: 14 } },
      { letter: "B", header: "Region", type: "text", stats: { blank: 0, distinct: 2, avgWords: 1, maxLen: 4 } },
      { letter: "C", header: "Amount", type: "number", stats: { blank: 1, distinct: 2 } },
      { letter: "D", header: "Diagnosis", type: "text", stats: { blank: 1, distinct: 2, avgWords: 1, maxLen: 4 } },
      { letter: "E", header: "Active", type: "boolean", stats: { blank: 1, distinct: 2 } },
    ],
    policies: [
      { letter: "A", private: true, treatment: "stand_in" },
      { letter: "B", private: false, treatment: "as_is" },
      { letter: "C", private: true, treatment: "range" },
      { letter: "D", private: true, treatment: "exclude", alias: "Field D" },
      { letter: "E", private: true, treatment: "as_is" },
    ],
  };
}

describe("auditor: buildAuditContext", () => {
  it("builds one range per spec with private values, aliases and effective treatments", () => {
    const spec = specFixture();
    const map = { isToken: (s: string) => s === "PERSON_001" };
    const ctx = buildAuditContext([spec], map, { model: MODEL, canaries: ["Quorbel Vantrisk"] });
    expect(ctx.model).toBe(MODEL);
    expect(ctx.canaries).toEqual(["Quorbel Vantrisk"]);
    expect(ctx.isToken("PERSON_001")).toBe(true);
    expect(ctx.isToken("PERSON_002")).toBe(false);
    expect(ctx.ranges).toHaveLength(1);
    const range = ctx.ranges[0]!;
    expect(range.sheet).toBe("Orders");
    expect(range.address).toBe("A1:E4");
    const [a, b, c, d, e] = range.columns;
    expect(a!.letter).toBe("A");
    expect(a!.private).toBe(true);
    expect(a!.values).toEqual(expect.arrayContaining(["Maria Lopez", "Kenji Watanabe"]));
    expect(a!.values).not.toContain("Customer");
    expect(a!.aliasOriginal).toBeUndefined();
    expect(b!.private).toBe(false);
    expect(b!.values).toEqual([]);
    expect(c!.treatment).toBe("range");
    expect(c!.values).toEqual(expect.arrayContaining(["$83,500", "83500", "$41,250", "41250"]));
    expect(d!.aliasOriginal).toBe("Diagnosis");
    expect(d!.values).toEqual(expect.arrayContaining(["Flu", "Cold"]));
    expect(e!.treatment).toBe("stand_in");
    expect(e!.values).toEqual(expect.arrayContaining(["TRUE", "FALSE"]));
    for (const col of range.columns) expect(col.values).not.toContain("");
  });

  it("covers context ranges and works with audit end to end", () => {
    const spec = specFixture();
    const ctx = buildAuditContext([spec, { ...spec, data: { ...spec.data, sheet: "Copy" } }], { isToken: () => false }, { model: MODEL });
    expect(ctx.canaries).toEqual([]);
    expect(ctx.ranges.map((r) => r.sheet)).toEqual(["Orders", "Copy"]);
    const cols: Col[] = [
      { letter: "A", header: "Customer", type: "text", private: true, treatment: "stand_in" },
      { letter: "B", header: "Region", type: "text", private: false, treatment: "as_is" },
      { letter: "C", header: "Amount", type: "number", private: true, treatment: "range" },
      { letter: "D", header: "Diagnosis", sentHeader: "Field D", type: "text", private: true, treatment: "exclude" },
      { letter: "E", header: "Active", type: "boolean", private: true, treatment: "stand_in" },
    ];
    const clean = makeBody({ user: desc({ cols, address: "A1:E4" }) });
    passed(run(clean, "structure_only", ctx));
    const leaky = makeBody({ user: desc({ cols, address: "A1:E4", task: "patients with Diagnosis Flu" }) });
    // Both ranges hold the aliased Diagnosis header; the context range's hit names its sheet.
    expect(blocked(run(leaky, "structure_only", ctx)).hits).toEqual([
      { column: "D", variant: "exact" },
      { column: "Copy!D", variant: "exact" },
    ]);
  });
});

// ---------------------------------------------------------------------------------------------
// The structural check accepts exactly what transform.ts renders

describe("auditor: agrees with transform.ts renderings", () => {
  it("accepts every renderRange output as a range rendering", () => {
    const rand = mulberry32(99);
    const xs = [0, 0.5, -0.5, 5, -5, 9.99, 10, -10, 15, 99, 100, 83500, -83500, 1e6, 123456789, -1e15];
    for (let i = 0; i < 500; i++) xs.push((rand() - 0.5) * 10 ** Math.floor(rand() * 12));
    for (const x of xs) {
      const r = renderRange(x);
      expect(isRangeRendering(r), `${x} -> ${r}`).toBe(true);
    }
  });

  it("rejects strings that only look like ranges", () => {
    for (const s of ["555-0199", "1990-2000", "01-19", "10-18", "-9-9", "0-10", "80000-89999-1", "", "-0"]) {
      expect(isRangeRendering(s), s).toBe(false);
    }
  });

  it("renderMonth output is accepted in a month column", () => {
    for (const [value, text] of [
      [46096, "3/15/2026"],
      [45658, "1/1/2025"],
      ["2026-03-15", "2026-03-15"],
    ] as const) {
      const m = renderMonth(value, text);
      expect(m).not.toBeNull();
      expect(MONTH_RE.test(m!)).toBe(true);
    }
  });
});

describe("auditor: end to end with transform.ts", () => {
  it("rows from transformRows and a question from substituteUserText pass the audit", () => {
    const spec = specFixture();
    const map = createStandInMap([spec.data]);
    const { rows } = transformRows(spec, map);
    const task = substituteUserText("Total for Maria Lopez and Kenji Watanabe", [spec], map);
    expect(task).not.toContain("Maria");
    const ctx = buildAuditContext([spec], map, { model: MODEL });
    const cols: Col[] = [
      { letter: "A", header: "Customer", type: "text", private: true, treatment: "stand_in" },
      { letter: "B", header: "Region", type: "text", private: false, treatment: "as_is" },
      { letter: "C", header: "Amount", type: "number", private: true, treatment: "range" },
      { letter: "D", header: "Diagnosis", sentHeader: "Field D", type: "text", private: true, treatment: "exclude" },
      { letter: "E", header: "Active", type: "boolean", private: true, treatment: "stand_in" },
    ];
    const body = makeBody({ user: desc({ mode: "substituted", cols, address: "A1:E4", rows, task }) });
    passed(run(body, "substituted", ctx));

    // A cell put back raw, however short, is caught.
    const leaked = rows.map((r) => [...r]);
    leaked[0]![3] = true;
    const result = blocked(run(makeBody({ user: desc({ mode: "substituted", cols, address: "A1:E4", rows: leaked, task }) }), "substituted", ctx));
    expect(result.structural).toContainEqual({ column: "E" });
  });
});

// ---------------------------------------------------------------------------------------------
// Red-team findings: other spellings of a private value, a header row that is data, a forged pass

/** A structure-only body whose task is `task`, audited with ORDERS columns holding `values`. */
function scanTask(task: string, values: Partial<Record<string, string[]>>): AuditOutcome {
  const cols = ORDERS.map((c) => (values[c.letter] ? { ...c, values: values[c.letter] } : c));
  return run(makeBody({ user: desc({ cols, task }) }), "structure_only", makeCtx({ cols }));
}

describe("auditor: other spellings of a private value", () => {
  it("folds Greek final sigma: a name stored in capitals is found when typed in lowercase", () => {
    const stored = "Γιώργος Παππάς".toUpperCase();
    expect(stored).toBe("ΓΙΏΡΓΟΣ ΠΑΠΠΆΣ");
    expect(blocked(scanTask("Τι μισθό παίρνει ο Γιώργος Παππάς;", { A: [stored] })).hits).toContainEqual({ column: "A", variant: "exact" });
    expect(blocked(scanTask("πόσα πουλάει ο Παππάς;", { A: [stored] })).hits).toEqual([{ column: "A", variant: "word" }]);
    // And the other way round: stored with final sigma, typed in capitals.
    expect(blocked(scanTask("ΠΑΠΠΆΣ", { A: ["Γιώργος Παππάς"] })).hits).toEqual([{ column: "A", variant: "word" }]);
  });

  it("finds a value stored decomposed (NFD) when typed composed (NFC), and the reverse", () => {
    const nfc = "José Muñoz".normalize("NFC");
    const nfd = "José Muñoz".normalize("NFD");
    expect(nfc).not.toBe(nfd);
    expect(blocked(scanTask(`total for ${nfc}`, { A: [nfd] })).hits).toContainEqual({ column: "A", variant: "exact" });
    expect(blocked(scanTask(`total for ${nfd}`, { A: [nfc] })).hits).toContainEqual({ column: "A", variant: "exact" });
    // Word variants come from both forms, with combining marks kept in their word.
    expect(blocked(scanTask("total for Muñoz".normalize("NFC"), { A: [nfd] })).hits).toEqual([{ column: "A", variant: "word" }]);
    expect(blocked(scanTask("total for Muñoz".normalize("NFD"), { A: [nfc] })).hits).toEqual([{ column: "A", variant: "word" }]);
  });

  it("folds typographic apostrophes, whichever one is stored or typed", () => {
    for (const [stored, typed] of [
      ["O’Sullivan", "O'Sullivan"],
      ["D'Alessandro", "D’Alessandro"],
      ["N‘Diaye", "Nʼdiaye"],
      ["D′Angelo", "Dʹangelo"],
    ] as const) {
      expect(blocked(scanTask(`total for ${typed}`, { A: [stored] })).hits, typed).toEqual([{ column: "A", variant: "exact" }]);
    }
  });

  it("finds a Japanese name written without the space it is stored with", () => {
    const A = ["山田 太郎", "佐藤　花子"];
    expect(blocked(scanTask("山田太郎さんの給与は?", { A })).hits).toEqual([{ column: "A", variant: "normalized" }]);
    expect(blocked(scanTask("佐藤花子の部署", { A })).hits).toEqual([{ column: "A", variant: "normalized" }]);
    expect(blocked(scanTask("Total for 山田太郎san", { A })).hits).toEqual([{ column: "A", variant: "normalized" }]);
  });

  it("ends a word at a Chinese, Japanese or Korean character", () => {
    // "Lopezさん": the surname from Maria Lopez, followed by an honorific.
    expect(blocked(scanTask("Lopezさんの売上合計は?", {})).hits).toEqual([{ column: "A", variant: "word" }]);
    expect(blocked(scanTask("アレクサンダーさんの売上", { A: ["アレクサンダー スミス"] })).hits).toEqual([{ column: "A", variant: "word" }]);
  });

  it("finds a Latin name written without its space only as a whole word", () => {
    expect(blocked(scanTask("orders for MaLi", { A: ["Ma Li"] })).hits).toEqual([{ column: "A", variant: "normalized" }]);
    // "Ma Li" without its space is "mali", which "normalize" contains.
    passed(scanTask("normalize the totals", { A: ["Ma Li"] }));
  });

  it("finds a phone number typed with other punctuation (digits projection)", () => {
    const C = ["831-555-0199", "4155550123"];
    for (const task of [
      "orders from (831) 555-0199",
      "orders from 831.555.0199",
      "call +1 831 555 0199",
      "call 415-555-0123",
      "call (415) 555 0123",
      "call 415 –\n555 0123",
    ]) {
      const result = blocked(scanTask(task, { C }));
      // The last group ("0199") is also a free-text word hit; the whole number is a digits hit.
      expect(result.hits, task).toContainEqual({ column: "C", variant: "digits" });
      expect(JSON.stringify(result)).not.toMatch(/555/);
    }
  });

  it("joins digits only across number punctuation, so separate numbers don't run together", () => {
    const C = ["1234567"];
    expect(blocked(scanTask("rows 123 4567", { C })).hits).toEqual([{ column: "C", variant: "digits" }]);
    passed(scanTask("rows 123, 4567", { C }));
    passed(scanTask("rows 123 and 4567", { C }));
  });

  it("finds an amount written with thousands separators or without its sign", () => {
    const D = ["83500", "-41250.5"];
    for (const task of ["earning 83,500", "earning 83.500", "earning 83 500", "owed 41,250.5", "owed 41.250,5", "owed 41250.5 back"]) {
      expect(blocked(scanTask(task, { D })).hits, task).toEqual([{ column: "D", variant: "normalized" }]);
    }
    passed(scanTask("earning 835 or 8,350", { D }));
  });
});

describe("auditor: a header row that is really data", () => {
  // The user selected A2:D3, without the header row, and left "has a header row" on.
  function headerless(): RangeSpec {
    const values: CellValue[][] = [
      ["Maria Lopez", "maria.lopez@example.com", 8315550199, "East"],
      ["Ana Ruiz", "ana.ruiz@example.com", 4155550123, "West"],
    ];
    const text = [
      ["Maria Lopez", "maria.lopez@example.com", "(831) 555-0199", "East"],
      ["Ana Ruiz", "ana.ruiz@example.com", "(415) 555-0123", "West"],
    ];
    return {
      data: {
        sheet: "Orders",
        address: "A2:D3",
        rowIndex: 1,
        columnIndex: 0,
        values,
        text,
        formulas: values,
        valueTypes: values.map((r) => r.map((v) => (typeof v === "number" ? "Double" : "String"))),
        numberFormat: values.map((r) => r.map(() => "General")),
      },
      hasHeaders: true,
      columns: [
        { letter: "A", header: "Maria Lopez", type: "text", stats: { blank: 0, distinct: 1, avgWords: 2, maxLen: 8 } },
        { letter: "B", header: "maria.lopez@example.com", type: "text", stats: { blank: 0, distinct: 1, avgWords: 1, maxLen: 20 } },
        { letter: "C", header: "(831) 555-0199", type: "number", stats: { blank: 0, distinct: 1 } },
        { letter: "D", header: "East", type: "text", stats: { blank: 0, distinct: 1, avgWords: 1, maxLen: 4 } },
      ],
      policies: [
        { letter: "A", private: true, treatment: "stand_in" },
        { letter: "B", private: true, treatment: "stand_in" },
        { letter: "C", private: true, treatment: "stand_in" },
        { letter: "D", private: false, treatment: "as_is" },
      ],
    };
  }

  it("scans a private column's header cell when it looks like data", () => {
    const ctx = buildAuditContext([headerless()], { isToken: () => false }, { model: MODEL });
    const [a, b, c, d] = ctx.ranges[0]!.columns;
    expect(b!.values).toEqual(expect.arrayContaining(["maria.lopez@example.com", "ana.ruiz@example.com"]));
    expect(c!.values).toEqual(expect.arrayContaining(["(831) 555-0199", "8315550199"]));
    // Plain text isn't data-shaped; the pane decides about the header row itself.
    expect(a!.values).not.toContain("Maria Lopez");
    expect(d!.values).toEqual([]);
  });

  it("blocks a request that sends such a header cell, in either mode", () => {
    const spec = headerless();
    const ctx = buildAuditContext([spec], { isToken: () => false }, { model: MODEL });
    const cols: Col[] = spec.columns.map((c, i) => ({
      letter: c.letter,
      header: c.header,
      type: c.type,
      private: spec.policies[i]!.private,
      treatment: spec.policies[i]!.treatment,
    }));
    for (const mode of ["structure_only", "substituted"] as const) {
      const rows = mode === "substituted" ? [[null, null, null, "West"]] : undefined;
      const body = makeBody({ user: desc({ mode, cols, rows, address: "A2:D3", task: "Count clients per region" }) });
      const result = blocked(run(body, mode, ctx));
      expect(result.hits, mode).toContainEqual({ column: "B", variant: "exact" });
      expect(result.hits, mode).toContainEqual({ column: "C", variant: "exact" });
      expect(result.structural).toEqual([]);
      expect(result.error).toBe("Text from private columns B, C is in the request. Nothing was sent.");
    }
  });
});

describe("auditor: isAudited", () => {
  it("can't be made to pass by patching WeakSet.prototype.has or Reflect.apply after load", () => {
    const genuine = passed(run(makeBody()));
    const copy = { ...genuine };
    const has = WeakSet.prototype.has;
    const apply = Reflect.apply;
    let forged: boolean;
    let kept: boolean;
    try {
      WeakSet.prototype.has = () => true;
      Object.defineProperty(Reflect, "apply", { value: () => true, configurable: true, writable: true });
      forged = isAudited(copy);
      kept = isAudited(genuine);
    } finally {
      WeakSet.prototype.has = has;
      Object.defineProperty(Reflect, "apply", { value: apply, configurable: true, writable: true });
    }
    expect(forged).toBe(false);
    expect(kept).toBe(true);
  });
});

// ---------------------------------------------------------------------------------------------
// Second review: false blocks from numbers, displayed numbers, more spellings, the matcher rule

/** A RangeSpec from rows, with display texts and number formats where they differ. */
function specOf(
  sheet: string,
  address: string,
  rows: CellValue[][],
  privateCols: Partial<Record<string, Treatment>>,
  opts: { text?: Record<string, string>; formats?: Record<string, string>; alias?: Record<string, string> } = {},
): RangeSpec {
  const rect = parseLocalRange(address)!;
  const width = rows[0]!.length;
  const letters = Array.from({ length: width }, (_, i) => String.fromCharCode(65 + rect.c1 + i));
  const data = {
    sheet,
    address,
    rowIndex: rect.r1,
    columnIndex: rect.c1,
    values: rows,
    text: rows.map((r, ri) => r.map((v, ci) => opts.text?.[`${ri},${ci}`] ?? (v === null ? "" : String(v)))),
    formulas: rows,
    valueTypes: rows.map((r) => r.map((v) => (v === null ? "Empty" : typeof v === "number" ? "Double" : typeof v === "boolean" ? "Boolean" : "String"))),
    numberFormat: rows.map((r, ri) => r.map((_, ci) => opts.formats?.[`${ri},${ci}`] ?? "General")),
  };
  return {
    data,
    hasHeaders: true,
    columns: inferSchema(data, true),
    policies: letters.map((letter) => ({
      letter,
      private: privateCols[letter] !== undefined,
      treatment: privateCols[letter] ?? "as_is",
      ...(opts.alias?.[letter] ? { alias: opts.alias[letter] } : {}),
    })),
  };
}

/** The request payload.ts builds for `spec` and `question` (substituted first), audited as flow.ts does. */
function auditSpec(spec: RangeSpec, question: string, mode: "structure_only" | "substituted", substitute = true): AuditOutcome {
  const map = createStandInMap([spec.data]);
  const sheetContext = buildSheetContext(spec.data, spec.hasHeaders, [], { sheets: [spec.data.sheet], tables: [] });
  const task = substitute ? substituteUserText(question, [spec], map) : question;
  const input = { model: MODEL, specs: [spec], sheetContext, question: task, history: [], map, rowCap: 50 };
  const built = mode === "substituted" ? buildSubstituted(input) : buildStructureOnly(input);
  return audit(built.outbound, buildAuditContext([spec], map, { model: MODEL }));
}

describe("auditor: numbers in other formats (second review)", () => {
  // Qty and Price are sent as they are; Amount is private.
  const cols: Col[] = [
    { letter: "A", header: "Customer", type: "text", private: true, treatment: "stand_in", values: ["Maria Lopez"] },
    { letter: "B", header: "Qty", type: "number", private: false, treatment: "as_is" },
    { letter: "C", header: "Price", type: "number", private: false, treatment: "as_is" },
    { letter: "D", header: "Amount", type: "number", private: true, treatment: "range", values: ["5390", "93496"] },
  ];
  const ctx = { ...makeCtx({ cols }), ranges: [rangeOf(cols, "Orders", "A1:D3")] };
  const scan = (rows: unknown[][], task: string) =>
    run(makeBody({ user: desc({ mode: "substituted", cols, rows, address: "A1:D3", task }) }), "substituted", ctx);

  it("doesn't read a grouped form across JSON values or inside another number", () => {
    // 5 and 390 are two cells: the row is written [..., 5,390, ...]. -93.496157 is another number.
    passed(scan([["PERSON_001", 5, 390, "5000-5999"], ["PERSON_002", 1, -93.496157, "90000-99999"]], "Total per Customer"));
    passed(scan([["PERSON_001", 5, 390, null]], "rows 15,390 or 5,390,000 or 193.496"));
  });

  it("still finds a grouped form in the question and in text cells", () => {
    const rows = [["PERSON_001", 5, 390, null]];
    for (const task of ["earning 5,390", "earning 5,390.", "earning 5.390?", "(5 390)", "-93.496"]) {
      expect(blocked(scan(rows, task)).hits, task).toEqual([{ column: "D", variant: "normalized" }]);
    }
    const textCell = [["PERSON_001", "5,390", 390, null]];
    expect(blocked(scan(textCell, "Total per Customer")).hits).toEqual([{ column: "D", variant: "normalized" }]);
  });

  it("makes number forms from amounts, ZIP codes and IDs, not from dates or booleans (third review)", () => {
    const spec = specOf(
      "Orders",
      "A1:F3",
      [
        ["Amount", "ZIP", "Order date", "Employee ID", "Raise", "Active"],
        [83500.4, 93496, 46096, 1000123, 0.0525, true],
        ["$91,250", 2134, 46100, 1000124, 0.031, false],
      ],
      { A: "range", B: "stand_in", C: "month", D: "stand_in", E: "range", F: "stand_in" },
      { text: { "1,0": "$83,500", "1,4": "5.25%", "2,4": "3.10%" }, formats: { "1,2": "m/d/yyyy", "2,2": "m/d/yyyy" } },
    );
    const [amount, zip, date, id, raise, active] = buildAuditContext([spec], { isToken: () => false }, { model: MODEL }).ranges[0]!.columns;
    expect(amount!.numbers).toEqual(expect.arrayContaining(["83500", "91250"]));
    expect(zip!.numbers).toEqual(expect.arrayContaining(["93496", "2134"]));
    expect(id!.numbers).toEqual(expect.arrayContaining(["1000123", "1000124"]));
    expect(date!.numbers).toEqual([]);
    expect(active!.numbers).toEqual([]);
    // "3.1" is shorter than 4 characters.
    expect(raise!.numbers).toEqual(expect.arrayContaining(["5.25", "3.10"]));
    expect(raise!.numbers).not.toContain("3.1");
  });

  it("finds the number the user sees, not only the raw value", () => {
    const spec = specOf(
      "Staff",
      "A1:D4",
      [
        ["Employee", "Salary", "Raise", "Paid"],
        ["Maria Lopez", 83500.4, 0.0525, "$91,250"],
        ["Ana Ruiz", 1234.2000000000003, 0.031, "$7,000"],
        ["Kenji Watanabe", 835000, 0.1, "$12,50,000"],
      ],
      { A: "stand_in", B: "range", C: "range", D: "stand_in" },
      { text: { "1,1": "$83,500", "2,1": "1,234.20", "3,1": "₹8,35,000", "1,2": "5.25%", "2,2": "3.10%", "3,2": "10.00%" } },
    );
    for (const question of [
      "who earns 83,500",
      "who earns 83500",
      "amount 1234.2",
      "amount 1.234,20",
      "salary 8,35,000",
      "who got 5.25",
      "who was paid 91,250",
      "who was paid 1250000",
    ]) {
      // Sent without substitution, the auditor blocks it; with it, the value is a stand-in.
      expect(auditSpec(spec, question, "structure_only", false).ok, question).toBe(false);
      expect(auditSpec(spec, question, "structure_only").ok, question).toBe(true);
    }
    passed(auditSpec(spec, "Average Salary per Raise band", "substituted"));
  });

  it("finds words of fewer than 7 digits only in free text, such as the last part of a phone number (third review)", () => {
    const C = ["(831) 555-0199", "415 555 0123"];
    // The question is free text: the last group of a phone number blocks there.
    expect(blocked(scanTask("order 0199 shipped in 2026, 0123 left", { C })).hits).toEqual([{ column: "C", variant: "word" }]);
    expect(blocked(scanTask("call 831 555 0199", { C })).hits).toEqual([
      { column: "C", variant: "digits" },
      { column: "C", variant: "word" },
    ]);
    expect(blocked(scanTask("call (415) 555 0123", { C })).hits).toContainEqual({ column: "C", variant: "digits" });
    // Not inside another number, and never in row cells or JSON numbers (see the third-review block).
    passed(scanTask("order 10199 or 01990 or 0199x", { C }));
    // Words of letters, and words of 7+ digits, still count anywhere.
    expect(blocked(scanTask("ref 1234567", { A: ["Order 1234567 A"] })).hits).toContainEqual({ column: "A", variant: "word" });
    expect(blocked(scanTask("for Watanabe", {})).hits).toEqual([{ column: "A", variant: "word" }]);
  });
});

/** The parts of a parsed description the count tests change. */
interface WireDesc {
  range: Record<string, unknown>;
  rows_sent?: number;
  rows_total?: number;
  columns: { stats: Record<string, number> }[];
  context_ranges?: unknown[];
}

describe("auditor: row numbers and counts (second review)", () => {
  // Header row 1000, data rows 1001 to 3000.
  const big = "A1000:F3000";
  const bigCtx = (cols: readonly Col[], extra: Partial<AuditRange> = {}) => ({
    ...makeCtx({ cols }),
    ranges: [{ ...rangeOf(cols, "Orders", big), ...extra }],
  });
  const countsError = "A description in the request has row numbers or counts that don't match your selection.";

  it("masks the row numbers and counts it has checked: a count equal to a private number is a coincidence", () => {
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, values: [...(c.values ?? []), "1500", "1999", "2000", "1001"] } : c));
    const d = JSON.parse(desc({ cols, address: big, mode: "substituted" })) as { columns: { stats: Record<string, number> }[] };
    d.columns[1]!.stats.distinct = 1500;
    d.columns[1]!.stats.max_len = 1999;
    // first_data_row is 1001 and rows_total 2000.
    const body = makeBody({ user: JSON.stringify(d) });
    expect(body).toContain('\\"first_data_row\\":1001');
    expect(body).toContain('\\"rows_total\\":2000');
    passed(run(body, "substituted", bigCtx(cols)));
    // The same numbers anywhere else are scanned.
    const inTask = makeBody({ user: desc({ cols, address: big, task: "the 1500 rows" }) });
    expect(blocked(run(inTask, "structure_only", bigCtx(cols))).hits).toEqual([{ column: "A", variant: "exact" }]);
  });

  it("blocks row numbers and counts that don't match the selection, with a plain error", () => {
    const edits: [string, (d: WireDesc) => void][] = [
      ["header row", (d) => (d.range.header_row = 1001)],
      ["first data row", (d) => (d.range.first_data_row = 1000)],
      ["last data row", (d) => (d.range.last_data_row = 3001)],
      ["no header row", (d) => (d.range.header_row = null)],
      ["rows total", (d) => (d.rows_total = 1999)],
      ["rows sent", (d) => (d.rows_sent = 4)],
      ["rows sent over total", (d) => (d.rows_sent = 2001)],
      ["blank", (d) => (d.columns[0]!.stats.blank = 2001)],
      ["distinct", (d) => (d.columns[0]!.stats.distinct = 12345)],
      ["negative", (d) => (d.columns[0]!.stats.blank = -1)],
      ["fraction", (d) => (d.columns[0]!.stats.distinct = 2.5)],
      ["max_len", (d) => (d.columns[0]!.stats.max_len = 32768)],
      ["avg_words decimals", (d) => (d.columns[0]!.stats.avg_words = 2.25)],
      ["avg_words size", (d) => (d.columns[0]!.stats.avg_words = 40000)],
      ["avg_words sign", (d) => (d.columns[0]!.stats.avg_words = -1)],
    ];
    for (const [label, edit] of edits) {
      const d = JSON.parse(desc({ address: big, mode: "substituted" })) as WireDesc;
      edit(d);
      const result = blocked(run(makeBody({ user: JSON.stringify(d) }), "substituted", bigCtx(ORDERS)));
      expect(result.error, label).toBe(countsError);
      expect(result.hits, label).toEqual([]);
    }
    // The same edits in a context range are checked too.
    const d = JSON.parse(desc({ address: big })) as WireDesc;
    const regions = { ...rangeOf(ORDERS, "Regions", "A1:F6") };
    d.context_ranges = [{ sheet: "Regions", table: null, range: { ...rangeJson("A1:F6"), last_data_row: 7 }, columns: columnsJson(ORDERS) }];
    const ctx = { ...bigCtx(ORDERS), ranges: [bigCtx(ORDERS).ranges[0]!, regions] };
    expect(blocked(run(makeBody({ user: JSON.stringify(d) }), "structure_only", ctx)).error).toBe(countsError);
  });

  it("checks a range without a header row, and blocks when the context doesn't know the row count", () => {
    const d = JSON.parse(desc({ address: big })) as WireDesc;
    d.range = { address: big, header_row: null, first_data_row: 1000, last_data_row: 3000 };
    passed(run(makeBody({ user: JSON.stringify(d) }), "structure_only", bigCtx(ORDERS, { hasHeaders: false, dataRows: 2001 })));
    const unknown = { ...bigCtx(ORDERS), ranges: [{ ...rangeOf(ORDERS, "Orders", big), dataRows: undefined as unknown as number }] };
    const result = blocked(run(makeBody({ user: desc({ address: big }) }), "structure_only", unknown));
    expect(result.error).toBe("The check couldn't work out the rows of your selection, so nothing was sent. Select your data, or one cell in it, and choose Refresh from selection.");
    // A context that disagrees with its own address blocks too.
    expect(blocked(run(makeBody({ user: desc({ address: big }) }), "structure_only", bigCtx(ORDERS, { dataRows: 1999 }))).error).toBe(countsError);
  });

  it("buildAuditContext records the header row and data row count", () => {
    const spec = specFixture();
    const range = buildAuditContext([spec, { ...spec, hasHeaders: false }], { isToken: () => false }, { model: MODEL }).ranges;
    expect(range[0]).toMatchObject({ hasHeaders: true, dataRows: 3 });
    expect(range[1]).toMatchObject({ hasHeaders: false, dataRows: 4 });
  });
});

describe("auditor: more spellings (second review)", () => {
  it("folds Turkish İ and ı, in any case and either form, with positions kept", () => {
    const A = ["İbrahim Kaya", "Ayşe Yıldız"];
    for (const task of ["for ibrahim", "for IBRAHIM", "for İBRAHİM", "for Ibrahim", "for yildiz", "for YILDIZ"]) {
      expect(blocked(scanTask(task, { A })).hits, task).toEqual([{ column: "A", variant: "word" }]);
    }
    // Decomposed (NFD) İ is I and a combining dot; either way it is "ibrahim".
    expect(blocked(scanTask("for İbrahim", { A })).hits).toEqual([{ column: "A", variant: "word" }]);
    expect(blocked(scanTask("İ İ i̇brahim kaya", { A })).hits).toContainEqual({ column: "A", variant: "exact" });
    // After a dropped dot, masks and word boundaries still line up.
    passed(scanTask("İİİ PERSON_001 ibrahimx", { A: ["Person Kaya", "Ibrahim Test"] }));
    expect(blocked(scanTask("İİİ PERSON_001 ibrahim", { A: ["Person Kaya", "Ibrahim Test"] })).hits).toEqual([
      { column: "A", variant: "word" },
    ]);
  });

  it("folds apostrophe look-alikes, full-width digits and hyphens", () => {
    for (const typed of ["O´Neill", "O`Neill", "O＇Neill", "O’Neill"]) {
      expect(blocked(scanTask(`total for ${typed}`, { A: ["O'Neill"] })).hits, typed).toEqual([{ column: "A", variant: "exact" }]);
    }
    const C = ["090-1234-5678"];
    expect(blocked(scanTask("０９０−１２３４−５６７８の注文", { C })).hits).toContainEqual({ column: "C", variant: "exact" });
    expect(blocked(scanTask("０９０（１２３４）５６７８", { C })).hits).toContainEqual({ column: "C", variant: "digits" });
  });

  it("finds amounts in lakh grouping", () => {
    const D = ["835000", "1250000"];
    for (const task of ["salary 8,35,000", "salary 12,50,000", "salary 835,000"]) {
      expect(blocked(scanTask(task, { D })).hits, task).toEqual([{ column: "D", variant: "normalized" }]);
    }
  });
});

describe("auditor: a year header (second review)", () => {
  const budget = (sheet: string) =>
    specOf(
      sheet,
      "A1:C4",
      [
        ["Account", 2023, 2024],
        ["Rent", 10000, 11000],
        ["Travel", 11375, 12450],
        ["Legal", 12750, 13900],
      ],
      { B: "range", C: "range" },
      { alias: { C: "Next year" } },
    );

  it("doesn't scan a bare number header, as a value or under its alias", () => {
    const ctx = buildAuditContext([budget("Budget 2024")], { isToken: () => false }, { model: MODEL });
    const [, b, c] = ctx.ranges[0]!.columns;
    expect(b!.values).not.toContain("2023");
    expect(c!.values).not.toContain("2024");
    expect(c!.aliasOriginal).toBeUndefined();
    for (const mode of ["structure_only", "substituted"] as const) passed(auditSpec(budget("Budget 2024"), "Total per Account", mode));
  });

  it("still scans a header that identifies someone, and an ID over a column of IDs", () => {
    const spec = specOf(
      "Orders",
      "A1:C3",
      [
        ["maria@example.com", "EMP-00123", "FY2024"],
        ["ana@example.com", "EMP-00124", 1200],
        ["kenji@example.com", "EMP-00125", 1300],
      ],
      { A: "stand_in", B: "stand_in", C: "range" },
    );
    const [a, b, c] = buildAuditContext([spec], { isToken: () => false }, { model: MODEL }).ranges[0]!.columns;
    expect(a!.values).toContain("maria@example.com");
    expect(b!.values).toContain("EMP-00123");
    expect(c!.values).not.toContain("FY2024");
  });
});

describe("auditor: matcher choice (second review)", () => {
  it(`uses Aho-Corasick only above ${INDEXOF_LIMIT} variants in a text longer than 16 KB`, () => {
    expect(AHO_MIN_TEXT).toBe(16 * 1024);
    expect(chooseMatcher("auto", 20000, 10_000)).toBe("indexOf");
    expect(chooseMatcher("auto", 20000, 16 * 1024)).toBe("indexOf");
    expect(chooseMatcher("auto", 20000, 16 * 1024 + 1)).toBe("ahoCorasick");
    expect(chooseMatcher("auto", 5000, 100_000)).toBe("indexOf");
    expect(chooseMatcher("auto", 5001, 100_000)).toBe("ahoCorasick");
    expect(chooseMatcher("indexOf", 50_000, 100_000)).toBe("indexOf");
    expect(chooseMatcher("ahoCorasick", 1, 1)).toBe("ahoCorasick");
  });

  it("finds the same hits either way on a long body with many variants", () => {
    const values = Array.from({ length: 3000 }, (_, i) => `Givenname${i} Familyname${i}`);
    const cols = ORDERS.map((c) => (c.letter === "A" ? { ...c, values } : c));
    const ctx = makeCtx({ cols });
    const task = `${"Total per Region. ".repeat(1200)}orders for Familyname2718`;
    const body = makeBody({ user: desc({ cols, task }) });
    expect(body.length).toBeGreaterThan(AHO_MIN_TEXT);
    const auto = run(body, "structure_only", ctx, "auto");
    expect(blocked(auto).hits).toEqual([{ column: "A", variant: "word" }]);
    expect(run(body, "structure_only", ctx, "indexOf")).toEqual(auto);
    expect(run(body, "structure_only", ctx, "ahoCorasick")).toEqual(auto);
  });
});

// ---------------------------------------------------------------------------------------------
// Third review: regressions and gaps

/** A description with a note on one column (desc() writes every note as null). */
function withNote(user: string, letter: string, note: string): string {
  const d = JSON.parse(user) as { columns: { letter: string; note: string | null }[] };
  d.columns.find((c) => c.letter === letter)!.note = note;
  return JSON.stringify(d);
}

describe("auditor: amounts stored as text (third review)", () => {
  // A CSV opened in Excel: the salaries came in as text.
  const staff = () =>
    specOf(
      "Staff",
      "A1:C4",
      [
        ["Employee", "Dept", "Salary"],
        ["Maria Lopez", "Sales", "83500"],
        ["Ana Ruiz", "Ops", "125000"],
        ["Kenji Watanabe", "IT", "1234.50"],
      ],
      { A: "stand_in", C: "stand_in" },
    );

  it("gives a private column of digit text number forms, whatever its token kind", () => {
    const [, , salary] = buildAuditContext([staff()], { isToken: () => false }, { model: MODEL }).ranges[0]!.columns;
    expect(salary!.numbers).toEqual(expect.arrayContaining(["83500", "125000", "1234.50", "1234.5"]));
  });

  it("blocks the amount typed in another format, and the question goes through substituted", () => {
    for (const question of ["who earns 83,500?", "who earns $125,000?", "salary 1,25,000", "Gehalt 125.000", "83 500", "invoice 1,234.50", "1.234,50"]) {
      expect(auditSpec(staff(), question, "structure_only", false).ok, question).toBe(false);
      expect(auditSpec(staff(), question, "structure_only").ok, question).toBe(true);
      expect(auditSpec(staff(), question, "substituted").ok, question).toBe(true);
    }
  });

  it("number forms of ZIP codes and IDs count only as whole numbers inside text", () => {
    const cols: Col[] = [
      { letter: "A", header: "Customer", type: "text", private: true, treatment: "stand_in", values: ["Maria Lopez"] },
      { letter: "B", header: "Qty", type: "number", private: false, treatment: "as_is" },
      { letter: "C", header: "Code", type: "text", private: false, treatment: "as_is" },
      { letter: "D", header: "ZIP", type: "number", private: true, treatment: "stand_in", values: ["93496", "2134"] },
    ];
    const ctx = { ...makeCtx({ cols }), ranges: [rangeOf(cols, "Orders", "A1:D3")] };
    const scan = (rows: unknown[][], task: string) =>
      run(makeBody({ user: desc({ mode: "substituted", cols, rows, address: "A1:D3", task }) }), "substituted", ctx);
    // The grouped forms ("93.496", "2,134") don't count inside JSON numbers or other numbers.
    passed(scan([["PERSON_001", 93.496, "93.4961", "TEXT_001"], ["PERSON_002", 2.134, "2,1345", "TEXT_002"]], "Total per Customer"));
    expect(blocked(scan([["PERSON_001", 1, "x", "TEXT_001"]], "orders from 93,496")).hits).toEqual([{ column: "D", variant: "normalized" }]);
  });
});

describe("auditor: short digit groups in free text (third review)", () => {
  const C = ["(831) 555-0199", "4111 1111 1111 1234", "123-45-6789"];
  const cols = ORDERS.map((c) => (c.letter === "C" ? { ...c, values: C } : c));
  const ctx = makeCtx({ cols });
  const WORD = [{ column: "C", variant: "word" }];
  const reply = (fields: Record<string, unknown>) =>
    JSON.stringify({ kind: "answer", formula: null, placement: null, explanation: "Done.", assumptions: [], ...fields });
  const withHistory = (task: string, content: string) =>
    makeBody({ history: [{ role: "user", content: desc({ cols, task }) }, { role: "assistant", content }], user: desc({ cols }) });

  it("blocks the last digits of a phone, card or SSN in the question", () => {
    for (const task of [
      "orders from 555-0199",
      "orders from 555 0199",
      "orders from 555.0199",
      "the customer ending 0199",
      "card ending 1234",
      "paid with **** 1234",
      "SSN ending 6789",
      "xxx-xx-6789",
    ]) {
      expect(blocked(scanTask(task, { C })).hits, task).toEqual(WORD);
    }
  });

  it("blocks them in notes, headers, sheet names, earlier questions and earlier replies' explanations and assumptions", () => {
    expect(blocked(run(makeBody({ user: withNote(desc({ cols }), "B", "call 555-0199 first") }), "structure_only", ctx)).hits).toEqual(WORD);
    const header = cols.map((c) => (c.letter === "B" ? { ...c, sentHeader: "Card 1234" } : c));
    expect(blocked(run(makeBody({ user: desc({ cols: header }) }), "structure_only", ctx)).hits).toEqual(WORD);
    const sheetCtx = { ...ctx, ranges: [rangeOf(cols, "Calls 0199", "A1:F6")] };
    expect(blocked(run(makeBody({ user: desc({ cols, sheet: "Calls 0199" }) }), "structure_only", sheetCtx)).hits).toEqual(WORD);
    expect(blocked(run(withHistory("orders from 555-0199", reply({})), "structure_only", ctx)).hits).toEqual(WORD);
    expect(blocked(run(withHistory("Count orders", reply({ explanation: "The card ending 1234 has 3 orders." })), "structure_only", ctx)).hits).toEqual(WORD);
    expect(blocked(run(withHistory("Count orders", reply({ assumptions: ["SSN ending 6789 is one person"] })), "structure_only", ctx)).hits).toEqual(WORD);
    // A reply that isn't JSON is free text throughout, and so is any text a reply has besides its
    // kind, formula and cell.
    expect(blocked(run(withHistory("Count orders", "The one ending 0199."), "structure_only", ctx)).hits).toEqual(WORD);
    expect(blocked(run(withHistory("Count orders", `Sure. ${reply({})} Ends 0199.`), "structure_only", ctx)).hits).toEqual(WORD);
    expect(blocked(run(withHistory("Count orders", reply({ note: "ending 0199" })), "structure_only", ctx)).hits).toEqual(WORD);
  });

  it("doesn't find them in row cells, JSON numbers or formulas, or inside other numbers", () => {
    const rows = [
      ["PERSON_001", "Store 0199", "TEXT_001", "80000-89999", "2026-03"],
      ["PERSON_002", 1234, "TEXT_002", "40000-49999", "2026-01"],
      ["PERSON_003", "Unit 6789-B", null, null, null],
    ];
    passed(run(makeBody({ user: desc({ mode: "substituted", cols, rows }) }), "substituted", ctx));
    const formula = reply({ kind: "formula", formula: '=COUNTIF(B2:B6,"*0199")+1234', placement: { cell: "G2", fill_down: false } });
    passed(run(withHistory("Count orders", formula), "structure_only", ctx));
    passed(scanTask("orders 10199, 01990, 12345 and 0199x", { C }));
  });

  it("leaves out groups that read as a year, and the groups of numbers and dates", () => {
    passed(scanTask("orders placed in 2026", { C: ["(831) 555-2026"] }));
    passed(scanTask("order 12345 or 2024", { C: ["$12,345.50", "12345.5", "3/15/2024"] }));
  });
});

describe("auditor: header cells (third review)", () => {
  it("scans an ID-shaped header cell without a word in it, but not a number such as a year", () => {
    const spec = specOf(
      "Staff",
      "A1:C4",
      [
        ["Maria Lopez", "B-201", 2024],
        ["Ana Ruiz", "Remote", 2350],
        ["Kenji Watanabe", "Home office", 2780],
        ["Omar Haddad", "Remote", 1850],
      ],
      { A: "stand_in", B: "stand_in", C: "range" },
    );
    const [, b, c] = buildAuditContext([spec], { isToken: () => false }, { model: MODEL }).ranges[0]!.columns;
    expect(b!.values).toContain("B-201");
    expect(c!.values).not.toContain("2024");
  });
});

describe("auditor: numbers with zero decimals (third review)", () => {
  const cols = ORDERS.map((c) => (c.letter === "D" ? { ...c, values: ["125000", "93496"] } : c));
  const ctx = makeCtx({ cols });
  const withNoteOn = (note: string) => run(makeBody({ user: withNote(desc({ cols }), "B", note) }), "structure_only", ctx);

  it("finds a number followed by a decimal mark and only zeros, in a note or the question", () => {
    for (const note of ["Kenji's 125,000.00 includes relocation", "Jahresgehalt, z. B. 125.000,00 €", "125000.00", "93,496.000", "1,25,000.00"]) {
      expect(blocked(withNoteOn(note)).hits.map((h) => h.column), note).toContain("D");
    }
    expect(blocked(scanTask("who earns 125,000.00?", { D: ["125000"] })).hits).toEqual([{ column: "D", variant: "normalized" }]);
  });

  it("doesn't find it followed by other digits or by another group", () => {
    // In free text, other decimals after a grouped form no longer hide it (eighth review).
    for (const note of ["93.496157", "125.000.000", "93,496,000", "1125,000"]) passed(withNoteOn(note));
  });
});

describe("auditor: other spellings (third review)", () => {
  it("finds a number stored with its trunk 0 typed in international form, and the reverse", () => {
    const C = ["020 7946 0958", "07700 900123"];
    for (const task of ["call +44 20 7946 0958", "call +442079460958", "call 0044 20 7946 0958", "call +44 7700 900123"]) {
      expect(blocked(scanTask(task, { C })).hits, task).toContainEqual({ column: "C", variant: "digits" });
    }
    expect(blocked(scanTask("call (831) 555-0199", { C: ["+18315550199"] })).hits).toContainEqual({ column: "C", variant: "digits" });
  });

  it("finds a value typed without its accents, and ß as ss", () => {
    const A = ["José Muñoz", "François Müller", "ΝΊΚΟΣ ΠΑΠΑΔΌΠΟΥΛΟΣ", "Łukasz Wójcik", "Gerhard Weiß"];
    const cases: [string, "exact" | "word"][] = [
      ["total for Jose Munoz", "exact"],
      ["total for munoz", "word"],
      ["Francois Muller", "exact"],
      ["ο Νικος Παπαδοπουλος", "exact"],
      ["Lukasz Wojcik", "exact"],
      ["Gerhard Weiss", "exact"],
      ["WEISS", "word"],
    ];
    for (const [task, variant] of cases) expect(blocked(scanTask(task, { A })).hits, task).toContainEqual({ column: "A", variant });
    // A value stored without accents is found typed with them, decomposed or not.
    for (const task of ["total for José", "total for José".normalize("NFD"), "total for JOSÉ"]) {
      expect(blocked(scanTask(task, { A: ["Jose Munoz"] })).hits, task).toEqual([{ column: "A", variant: "word" }]);
    }
    // A match takes its accents, so it still needs a word boundary after them.
    passed(scanTask("total for Joséx".normalize("NFD"), { A: ["Jose Munoz"] }));
  });

  it("folds full-width Latin letters", () => {
    expect(blocked(scanTask("total for Ｍａｒｉａ Ｌｏｐｅｚ", {})).hits).toContainEqual({ column: "A", variant: "exact" });
    expect(blocked(scanTask("total for Ｗａｔａｎａｂｅ", {})).hits).toEqual([{ column: "A", variant: "word" }]);
  });

  it("treats _ as a word boundary and finds a name joined with _", () => {
    for (const task of ["rows for maria_lopez", "use Bonus_Maria_Lopez", "KENJI_WATANABE"]) {
      expect(blocked(scanTask(task, {})).hits, task).toContainEqual({ column: "A", variant: "normalized" });
    }
    expect(blocked(scanTask("use Bonus_Watanabe_2026", {})).hits).toEqual([{ column: "A", variant: "word" }]);
    passed(scanTask("use mariax_lopezz", {}));
  });

  it("folds Arabic-Indic, Persian and Devanagari digits and the Arabic thousands separator", () => {
    const C = ["0501234567", "9876543210"];
    for (const task of ["call ٠٥٠١٢٣٤٥٦٧", "call ۰۵۰۱۲۳۴۵۶۷", "call ९८७६५४३२१०", "call ٠٥٠-١٢٣-٤٥٦٧"]) {
      expect(blocked(scanTask(task, { C })).hits.map((h) => h.column), task).toContain("C");
    }
    for (const task of ["salary ٨٣٥٠٠", "salary ٨٣٬٥٠٠", "salary ۸۳٬۵۰۰"]) {
      expect(blocked(scanTask(task, { D: ["83500"] })).hits.map((h) => h.column), task).toContain("D");
    }
  });
});

describe("auditor: dates in other formats (third review)", () => {
  const hr = () =>
    specOf(
      "HR",
      "A1:C3",
      [
        ["Employee", "Date of birth", "Dept"],
        ["Maria Lopez", 31121, "Sales"],
        ["Ana Ruiz", "4/2/1990", "Ops"],
      ],
      { A: "stand_in", B: "month" },
      { text: { "1,1": "3/15/1985" }, formats: { "1,1": "m/d/yyyy" } },
    );

  it("records a private column's dates as ISO dates, every reading (fifth review: no list of written forms)", () => {
    const [a, b] = buildAuditContext([hr()], { isToken: () => false }, { model: MODEL }).ranges[0]!.columns;
    expect(b!.dates).toEqual(["1985-03-15", "1990-04-02", "1990-02-04"]);
    expect(a!.dates).toEqual([]);
  });

  it("blocks a private date written in another format, and the question goes through substituted", () => {
    for (const question of [
      "born 1985-03-15",
      "born 15/03/1985",
      "born 15.03.1985",
      "born 15.3.1985",
      "born March 15, 1985",
      "born 15 March 1985",
      "born Mar 15 1985",
      "born 2 April 1990",
      "born 1990-02-04",
    ]) {
      expect(auditSpec(hr(), question, "structure_only", false).ok, question).toBe(false);
      expect(auditSpec(hr(), question, "structure_only").ok, question).toBe(true);
    }
  });

  it("matches only whole dates: a year or a month alone passes", () => {
    for (const question of ["born in 1985", "born in March 1985", "hired 1985-03-16", "rows 11985-03-15"]) {
      passed(auditSpec(hr(), question, "structure_only", false));
    }
  });
});

// ---------------------------------------------------------------------------------------------
// Fourth review

describe("auditor: date phrases in free text (fourth review)", () => {
  const serial = (y: number, m: number, d: number) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  const hr = () =>
    specOf(
      "HR",
      "A1:C3",
      [
        ["Employee", "Date of birth", "Dept"],
        ["Maria Lopez", serial(1985, 3, 15), "Sales"],
        ["Ines Duarte", serial(1988, 7, 4), "Ops"],
      ],
      { A: "stand_in", B: "month" },
      { text: { "1,1": "3/15/1985", "2,1": "7/4/1988" }, formats: { "1,1": "m/d/yyyy", "2,1": "m/d/yyyy" } },
    );
  const TYPED = [
    "born 3/15/85",
    "born 03/15/85",
    "born 15/03/85",
    "born 15.3.85",
    "born 15-Mar-1985",
    "born 15-Mar-85",
    "born Mar-15-1985",
    "born 1985/03/15",
    "born 1985/3/15",
    "born 1985.03.15",
    "born March 15th, 1985",
    "born 15th March 1985",
    "born Mar. 15, 1985",
    "born 15 Mar. 1985",
    "born 19850315",
    "born on the 4th of July 1988",
    "born July 4th, 1988",
  ];

  it("blocks a private date written in any date form in the question, naming the column as normalized", () => {
    for (const question of TYPED) {
      expect(blocked(auditSpec(hr(), question, "structure_only", false)).hits, question).toEqual([{ column: "B", variant: "normalized" }]);
    }
  });

  it("passes the same questions once substituteUserText has replaced the date, in either mode", () => {
    for (const question of TYPED) {
      for (const mode of ["structure_only", "substituted"] as const) passed(auditSpec(hr(), question, mode));
    }
  });

  it("finds them in notes, headers, sheet names and earlier replies, and across an escaped line break", () => {
    const cols = ORDERS.map((c) => (c.letter === "F" ? { ...c, values: ["3/15/1985"] } : c));
    const ctx = makeCtx({ cols });
    const F = [{ column: "F", variant: "normalized" }];
    expect(blocked(run(makeBody({ user: withNote(desc({ cols }), "B", "the one born 15-Mar-85") }), "structure_only", ctx)).hits).toEqual(F);
    const header = cols.map((c) => (c.letter === "B" ? { ...c, sentHeader: "Region (born March 15th, 1985)" } : c));
    expect(blocked(run(makeBody({ user: desc({ cols: header }) }), "structure_only", ctx)).hits).toEqual(F);
    const sheetCtx = { ...ctx, ranges: [rangeOf(cols, "Born 1985.03.15", "A1:F6")] };
    expect(blocked(run(makeBody({ user: desc({ cols, sheet: "Born 1985.03.15" }) }), "structure_only", sheetCtx)).hits).toEqual(F);
    const reply = JSON.stringify({ kind: "answer", formula: null, placement: null, explanation: "Born 3/15/85.", assumptions: [] });
    const history = [{ role: "user", content: desc({ cols, task: "Who is oldest?" }) }, { role: "assistant", content: reply }];
    expect(blocked(run(makeBody({ history, user: desc({ cols }) }), "structure_only", ctx)).hits).toEqual(F);
    expect(blocked(scanTask("born March\n15, 1985", { F: ["3/15/1985"] })).hits).toEqual(F);
  });

  it("reads them in row cells too (fifth review), and only whole dates that are private", () => {
    const cols = ORDERS.map((c) => (c.letter === "F" ? { ...c, values: ["3/15/1985"] } : c));
    const withCell = (cell: string) => ROWS.map((r, i) => (i === 0 ? [r[0], cell, ...r.slice(2)] : r));
    // A row cell of a column that isn't private: the same date there blocks, as the same name would.
    for (const cell of ["3/15/85", "15.03.85", "1985/3/15", "15-Mar-85", "Mar. 15, 1985", "SO-1985-03-15-001"]) {
      const body = makeBody({ user: desc({ mode: "substituted", cols, rows: withCell(cell) }) });
      expect(blocked(run(body, "substituted", makeCtx({ cols }))).hits, cell).toEqual([{ column: "F", variant: "normalized" }]);
    }
    // The loosest forms (8 digits, numbers separated by spaces) are as often an ID or a phone
    // number: in row cells they don't count.
    for (const cell of ["19850315", "15 03 85", "06 15 03 85 12", "06.45.15.03.85"]) {
      passed(run(makeBody({ user: desc({ mode: "substituted", cols, rows: withCell(cell) }) }), "substituted", makeCtx({ cols })));
    }
    for (const task of ["born 3/16/85", "born in March 1985", "born in 1985", "id 119850315", "born 4th of July 1988"]) {
      passed(scanTask(task, { F: ["3/15/1985"] }));
    }
    for (const task of ["born 19850315", "born 15 03 85"]) {
      expect(blocked(scanTask(task, { F: ["3/15/1985"] })).hits, task).toEqual([{ column: "F", variant: "normalized" }]);
    }
  });

  it("reads the dates of a context that lists other forms than the ISO one", () => {
    const ctx = makeCtx();
    ctx.ranges[0]!.columns[5]!.dates = ["March 15, 1985"];
    expect(blocked(run(makeBody({ user: desc({ task: "born 3/15/85" }) }), "structure_only", ctx)).hits).toEqual([{ column: "F", variant: "normalized" }]);
    passed(run(makeBody({ user: desc({ task: "born 3/16/85" }) }), "structure_only", ctx));
  });

  it("gives text dates with 2-digit years their dates", () => {
    const spec = specOf(
      "HR",
      "A1:B3",
      [
        ["Employee", "Date of birth"],
        ["Maria Lopez", "03/15/85"],
        ["Ana Ruiz", "04/02/90"],
      ],
      { A: "stand_in", B: "stand_in" },
    );
    for (const question of ["born March 15, 1985", "born 1985-03-15", "born 15 March 1985", "born 3/15/85"]) {
      expect(blocked(auditSpec(spec, question, "structure_only", false)).hits, question).toContainEqual({ column: "B", variant: "normalized" });
      passed(auditSpec(spec, question, "structure_only"));
    }
  });
});

describe("auditor: round hundreds in free text (fifth review: the fourth-round skip is reverted)", () => {
  it("blocks the last 4 digits of a phone, SSN or card number even when they are a round number of hundreds", () => {
    const C = ["(831) 555-1500", "(415) 555-5000", "123-45-6700", "4111 1111 1111 3400", "(646) 555-0100"];
    for (const task of [
      "How many orders came from 555-1500?",
      "Who has the phone ending 1500?",
      "Which orders are over 5000?",
      "Which employee's SSN ends in 6700?",
      "Look up xxx-xx-6700",
      "card ending 3400",
      "the phone ending 0100",
    ]) {
      expect(blocked(scanTask(task, { C })).hits, task).toEqual([{ column: "C", variant: "word" }]);
    }
    // Years are still skipped.
    passed(scanTask("sales in 2024", { C: ["(212) 555-2024"] }));
  });
});

describe("auditor: identifiers in a header row that is really data (fourth review)", () => {
  it("scans a private column's IBAN, NI number, postcode or address header cell", () => {
    for (const cell of ["DE89 3704 0044 0532 0130 00", "QQ 12 34 56 C", "SW1A 1AA", "12 Baker Street"]) {
      const spec = specOf(
        "Customers",
        "A2:C3",
        [
          ["Maria Lopez", cell, "London"],
          ["Ana Ruiz", "GB29 NWBK 6016 1331 9268 19", "Leeds"],
        ],
        { A: "stand_in", B: "stand_in" },
      );
      const b = buildAuditContext([spec], { isToken: () => false }, { model: MODEL }).ranges[0]!.columns[1]!;
      expect(b.values, cell).toContain(cell);
      expect(blocked(auditSpec(spec, "How many customers per city?", "structure_only")).hits, cell).toContainEqual({ column: "B", variant: "exact" });
    }
  });
});

// ---------------------------------------------------------------------------------------------
// Fifth review

/** Excel's serial for a date in the default 1900 date system. */
const serial1900 = (y: number, m: number, d: number) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);

describe("auditor: private dates read, not listed in every written form (fifth review)", () => {
  const p2 = (n: number) => String(n).padStart(2, "0");

  it("keeps at most the four readings of a 2-digit-year date per cell, and still finds each date in any form", () => {
    const rows: CellValue[][] = [["Employee", "Date of birth"]];
    for (let i = 0; i < 2000; i++) {
      const d = new Date(Date.UTC(1950, 0, 1 + i * 9));
      rows.push([`Person ${i}`, `${p2(d.getUTCMonth() + 1)}/${p2(d.getUTCDate())}/${String(d.getUTCFullYear()).slice(2)}`]);
    }
    const spec = specOf("HR", `A1:B${rows.length}`, rows, { A: "stand_in", B: "stand_in" });
    const b = buildAuditContext([spec], { isToken: () => false }, { model: MODEL }).ranges[0]!.columns[1]!;
    expect(b.dates!.length).toBeLessThanOrEqual(4 * 2000);
    expect(b.dates!.every((d) => /^\d{4}-\d{2}-\d{2}$/.test(d))).toBe(true);
    // Row 8 is 1/1/1950 + 63 days: March 5, 1950.
    expect(rows[8]![1]).toBe("03/05/50");
    for (const question of ["born 1950-03-05", "born March 5th, 1950", "born 5.3.1950", "born 5-Mar-50", "born 3/5/50"]) {
      expect(blocked(auditSpec(spec, question, "structure_only", false)).hits, question).toEqual([{ column: "B", variant: "normalized" }]);
      passed(auditSpec(spec, question, "structure_only"));
    }
  });

  it("finds a private date written in another language, or in Chinese, Japanese or Korean form, in the question", () => {
    const spec = specOf(
      "HR",
      "A1:B2",
      [
        ["Employee", "Date of birth"],
        ["Maria Lopez", serial1900(1985, 3, 15)],
      ],
      { A: "stand_in", B: "month" },
      { text: { "1,1": "3/15/1985" }, formats: { "1,1": "m/d/yyyy" } },
    );
    for (const question of ["born 15. März 1985", "born 15 mars 1985", "born 15 de marzo de 1985", "born 1985年3月15日", "born 1985년 3월 15일"]) {
      expect(blocked(auditSpec(spec, question, "structure_only", false)).hits, question).toEqual([{ column: "B", variant: "normalized" }]);
      passed(auditSpec(spec, question, "substituted"));
    }
  });
});

describe("auditor: private dates stored with a time zone or in another language (fifth review)", () => {
  const hr = (stored: string) =>
    specOf(
      "HR",
      "A1:B2",
      [
        ["Employee", "Date of birth"],
        ["Maria Lopez", stored],
      ],
      { A: "stand_in", B: "stand_in" },
    );

  it("reads ISO text with a time and an offset or zone, so the date typed in other forms is found", () => {
    const cases: [string, string[]][] = [
      ["1985-03-15T00:00:00-05:00", ["Who was born on 1985-03-15?", "Who was born on March 15, 1985?", "born 3/15/1985"]],
      ["1990-04-02 00:00:00+00", ["Who was born 4/2/1990?", "born April 2, 1990"]],
      ["1982-02-18-08:00", ["Who was born 1982-02-18?", "born Feb 18, 1982"]],
      ["2001-01-09 00:00:00 UTC", ["born 2001-01-09", "born 9 Jan 2001"]],
      ["1988-07-04T00:00:00.000+01:00", ["Who was born July 4, 1988?", "born 7/4/1988"]],
    ];
    for (const [stored, questions] of cases) {
      for (const question of questions) {
        expect(blocked(auditSpec(hr(stored), question, "structure_only", false)).hits, `${stored}: ${question}`).toContainEqual({
          column: "B",
          variant: "normalized",
        });
        passed(auditSpec(hr(stored), question, "structure_only"));
      }
    }
  });

  it("reads dates written in another language, so the date typed with numbers is found", () => {
    for (const stored of ["15. März 1985", "15 mars 1985", "15 de marzo de 1985", "1985年3月15日", "1985년 3월 15일"]) {
      for (const question of ["born 1985-03-15", "born 15.03.1985", "born 3/15/1985", "born March 15, 1985"]) {
        expect(blocked(auditSpec(hr(stored), question, "structure_only", false)).hits, `${stored}: ${question}`).toContainEqual({
          column: "B",
          variant: "normalized",
        });
        passed(auditSpec(hr(stored), question, "substituted"));
      }
    }
  });
});

describe("auditor: dates joined without spaces (fifth review)", () => {
  it("finds each date of a range typed without spaces, and dates with a stray space or comma", () => {
    for (const task of ["born 3/15/85-4/2/90", "born 3/15/85–4/2/90", "born 3/15/85,4/2/90", "born 15.03.85-02.04.90", "born 1985/03/15-1990/04/02"]) {
      // The first date of the pair, and the second.
      expect(blocked(scanTask(task, { F: ["3/15/1985"] })).hits, task).toEqual([{ column: "F", variant: "normalized" }]);
      expect(blocked(scanTask(task, { F: ["4/2/1990"] })).hits, task).toEqual([{ column: "F", variant: "normalized" }]);
    }
    expect(blocked(scanTask("born 15-Mar-85/02-Apr-90", { F: ["4/2/1990"] })).hits).toEqual([{ column: "F", variant: "normalized" }]);
    for (const task of ["born 15.3. 1985", "born 3/15, 1985"]) {
      expect(blocked(scanTask(task, { F: ["3/15/1985"] })).hits, task).toEqual([{ column: "F", variant: "normalized" }]);
    }
  });
});

describe("auditor: dates in a workbook that uses the 1904 date system (fifth review)", () => {
  it("reads the date a date cell shows, not only the date its serial names in the 1900 system", () => {
    // In the 1904 date system, March 15, 1985 is serial 29659; read in the 1900 system that is 1981-03-14.
    const spec = specOf(
      "HR",
      "A1:B2",
      [
        ["Employee", "Date of birth"],
        ["Maria Lopez", serial1900(1985, 3, 15) - 1462],
      ],
      { A: "stand_in", B: "month" },
      { text: { "1,1": "3/15/1985" }, formats: { "1,1": "m/d/yyyy" } },
    );
    for (const question of ["born 1985-03-15", "born March 15, 1985", "born 15.03.1985"]) {
      expect(blocked(auditSpec(spec, question, "structure_only", false)).hits, question).toContainEqual({ column: "B", variant: "normalized" });
      passed(auditSpec(spec, question, "structure_only"));
    }
  });
});

describe("auditor: digits between letters (fifth review)", () => {
  const C = ["(831) 555-0199 x204", "+1 212 555 0147 ext 3", "GB29 NWBK 6016 1331 9268 19"];

  it("finds a phone number without its extension, and an IBAN's account number, typed as digits", () => {
    for (const task of ["Orders for 8315550199?", "Orders for +18315550199?", "Orders for 2125550147?", "account 60161331926819", "call 831.555.0199"]) {
      expect(blocked(scanTask(task, { C })).hits, task).toContainEqual({ column: "C", variant: "digits" });
    }
  });

  it("replaces them in the question", () => {
    const spec = specOf(
      "Contacts",
      "A1:B3",
      [
        ["Name", "Phone"],
        ["Maria Lopez", C[0]!],
        ["Ana Ruiz", C[1]!],
      ],
      { A: "stand_in", B: "stand_in" },
    );
    for (const question of ["Orders for 8315550199?", "Orders for 2125550147?"]) passed(auditSpec(spec, question, "substituted"));
  });
});

describe("auditor: plain numbers inside other numbers (fifth review)", () => {
  // B and C are sent as they are; D (ZIP) and E (commission) are private; F holds a 7-digit ID.
  const cols: Col[] = [
    { letter: "A", header: "Customer", type: "text", private: true, treatment: "stand_in", values: ["Maria Lopez"] },
    { letter: "B", header: "Longitude", type: "number", private: false, treatment: "as_is" },
    { letter: "C", header: "Note", type: "text", private: false, treatment: "as_is" },
    { letter: "D", header: "ZIP", type: "number", private: true, treatment: "range", values: ["93496"] },
    { letter: "E", header: "Commission", type: "number", private: true, treatment: "range", values: ["2345", "5.25"] },
    { letter: "F", header: "Employee ID", type: "number", private: true, treatment: "range", values: ["1234567"] },
  ];
  const ctx = { ...makeCtx({ cols }), ranges: [rangeOf(cols, "Data", "A1:F3")] };
  const scan = (b: unknown, c: unknown, task = "Count per Note") =>
    run(makeBody({ user: desc({ mode: "substituted", cols, sheet: "Data", address: "A1:F3", task, rows: [["PERSON_001", b, c, "90000-99999", "1000-1999", "1000000-1999999"]] }) }), "substituted", ctx);

  it("doesn't find a private number of fewer than 7 digits with a digit right next to it in a number cell", () => {
    // Nor in the decimals of a number with one decimal point (seventh review): see below.
    for (const b of [-120.934961, 934961, 123456, 12345.25, 15.255]) passed(scan(b, "ok"));
    // In a text cell it does (eleventh review): "Transfer to 48392001" holds an account number.
    for (const c of ["934961", "-120.934961", "total 123456"]) expect(blocked(scan(1, c)).hits.length, c).toBeGreaterThan(0);
    // In free text it does (tenth review).
    expect(blocked(scan(1, "ok", "rows with 934961 or 123456")).hits).toEqual([
      { column: "D", variant: "exact" },
      { column: "E", variant: "exact" },
    ]);
  });

  it("still finds it on its own, next to letters or a dash, or with zero decimals", () => {
    for (const c of ["ZIP93496", "93496-1234", "93496.00", "(93496)"]) {
      expect(blocked(scan(1, c)).hits, c).toContainEqual({ column: "D", variant: "exact" });
    }
    expect(blocked(scan(93496, "ok")).hits).toContainEqual({ column: "D", variant: "exact" });
    expect(blocked(scan(1, "ok", "orders from 93496")).hits).toContainEqual({ column: "D", variant: "exact" });
    expect(blocked(scan(2345, "ok")).hits).toContainEqual({ column: "E", variant: "exact" });
  });

  it("finds it next to a comma or a decimal mark, in the question, a note or a JSON number (sixth review)", () => {
    for (const c of ["93496,94103", "=IF(B2=93496,1,0)", "{93496,94103}", "93496.5", "ZIP 93496,5", "1,93496"]) {
      expect(blocked(scan(1, c)).hits, c).toContainEqual({ column: "D", variant: "exact" });
    }
    for (const task of ["ZIPs 93496,94103?", "Why does =IF(B2=93496,1,0) give 0?", "=SUMPRODUCT((B2:B5={93496,94103})*C2:C5)"]) {
      expect(blocked(scan(1, "ok", task)).hits, task).toContainEqual({ column: "D", variant: "exact" });
    }
    const note = withNote(desc({ mode: "substituted", cols, sheet: "Data", address: "A1:F3", rows: [["PERSON_001", 1, "ok", "90000-99999", "1000-1999", "1000000-1999999"]] }), "C", "Only ZIPs 93496,94103 matter");
    expect(blocked(run(makeBody({ user: note }), "substituted", ctx)).hits).toContainEqual({ column: "D", variant: "exact" });
    // A JSON number in a row cell of a column sent as it is.
    for (const b of [93496.5, 2345.75]) expect(blocked(scan(b, "ok")).hits.length, String(b)).toBeGreaterThan(0);
  });

  it("finds it after zeros that follow no digit, and takes zero decimals after its own as the same number (sixth review)", () => {
    for (const c of ["093496", "0093496", "#093496", "ZIP 093496-1234"]) expect(blocked(scan(1, c)).hits, c).toContainEqual({ column: "D", variant: "exact" });
    for (const c of ["5.250", "rate 5.2500"]) expect(blocked(scan(1, c)).hits, c).toContainEqual({ column: "E", variant: "exact" });
    expect(blocked(scan(1, "5,250")).hits).toContainEqual({ column: "E", variant: "normalized" });
    // Zeros after a digit, or other decimals after its own, make another number when it is stored
    // as a number; in a text cell they are found anyway (eleventh review).
    for (const b of [1093496, 10093496, 5.251, 5.2501, 15.25]) passed(scan(b, "ok"));
    for (const c of ["1093496", "10093496", "5.251", "5.2501", "15.25"]) expect(blocked(scan(1, c)).hits.length, c).toBeGreaterThan(0);
  });

  it("keeps finding a number of 7+ digits anywhere, a phone number typed with a country code for one", () => {
    expect(blocked(scan(91234567.5, "ok")).hits).toContainEqual({ column: "F", variant: "exact" });
    expect(blocked(scan(1, "+11234567")).hits).toContainEqual({ column: "F", variant: "exact" });
  });

  it("doesn't find a date's serial inside another number", () => {
    const withSerial = ORDERS.map((c) => (c.letter === "F" ? { ...c, values: ["31121", "3/15/1985"] } : c));
    const rows = ROWS.map((r, i) => (i === 0 ? [r[0], 1311215, ...r.slice(2)] : r));
    passed(run(makeBody({ user: desc({ mode: "substituted", cols: withSerial, rows }) }), "substituted", makeCtx({ cols: withSerial })));
    const exact = ROWS.map((r, i) => (i === 0 ? [r[0], 31121, ...r.slice(2)] : r));
    expect(blocked(run(makeBody({ user: desc({ mode: "substituted", cols: withSerial, rows: exact }) }), "substituted", makeCtx({ cols: withSerial }))).hits).toContainEqual({
      column: "F",
      variant: "exact",
    });
  });
});

describe("auditor: range addresses the auditor derives itself (fifth review)", () => {
  const cols = ORDERS.map((c) => (c.letter === "D" ? { ...c, values: ["1000", "$1,000", "20000"] } : c));
  const ctxFor = (sheet: string, address: string) => ({ ...makeCtx({ cols }), ranges: [rangeOf(cols, sheet, address)] });

  it("doesn't match a private number inside the range's address or allowed range", () => {
    passed(run(makeBody({ user: desc({ cols, address: "A1:F1000" }) }), "structure_only", ctxFor("Orders", "A1:F1000")));
    passed(run(makeBody({ user: desc({ cols, address: "A1:F20000" }) }), "structure_only", ctxFor("Orders", "A1:F20000")));
    // A sheet name with a space is quoted in an allowed range, as payload.ts writes it.
    const quoted = { allowed_ranges: ["'My Sheet'!A1:F20000"] };
    passed(run(makeBody({ user: desc({ cols, address: "A1:F20000", sheet: "My Sheet", extra: quoted }) }), "structure_only", ctxFor("My Sheet", "A1:F20000")));
  });

  it("still scans the sheet name, an allowed range the context doesn't have, and the question", () => {
    const D = [{ column: "D", variant: "exact" }];
    expect(blocked(run(makeBody({ user: desc({ cols, address: "A1:F6", sheet: "Budget 1000" }) }), "structure_only", ctxFor("Budget 1000", "A1:F6"))).hits).toEqual(D);
    const extra = { allowed_ranges: ["Orders!A1:F1000", "Orders!B1000"] };
    expect(blocked(run(makeBody({ user: desc({ cols, address: "A1:F1000", extra }) }), "structure_only", ctxFor("Orders", "A1:F1000"))).hits).toEqual(D);
    expect(
      blocked(run(makeBody({ user: desc({ cols, address: "A1:F1000", task: "amounts over 1000" }) }), "structure_only", ctxFor("Orders", "A1:F1000"))).hits,
    ).toEqual(D);
  });

  it("lets a DOB column of 20,000 dates through when one of them is serial 20000", () => {
    const rows: CellValue[][] = [["Date of birth"]];
    for (let i = 0; i < 19999; i++) rows.push([18000 + ((i * 7) % 20000)]);
    rows[500]![0] = 20000;
    const text: Record<string, string> = {};
    const formats: Record<string, string> = {};
    rows.forEach((r, i) => {
      if (i === 0) return;
      const d = new Date(Date.UTC(1899, 11, 30) + (r[0] as number) * 86400000);
      text[`${i},0`] = `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`;
      formats[`${i},0`] = "m/d/yyyy";
    });
    const spec = specOf("HR", "A1:A20000", rows, { A: "month" }, { text, formats });
    passed(auditSpec(spec, "How many were born in each decade?", "structure_only"));
  });
});

describe("auditor: more spellings (fifth review)", () => {
  it("finds a name typed without its apostrophe or hyphen", () => {
    const A = ["Mary O'Brien", "Anne-Marie", "Siobhan O’Neill", "Jean-Luc Picard"];
    const cases: [string, "word" | "normalized"][] = [
      ["rows for OBrien", "word"],
      ["rows for Obrien", "word"],
      ["rows for ONeill", "word"],
      ["rows for AnneMarie", "normalized"],
      ["rows for Anne Marie", "normalized"],
      ["rows for JeanLuc", "word"],
    ];
    for (const [task, variant] of cases) expect(blocked(scanTask(task, { A })).hits, task).toEqual([{ column: "A", variant }]);
    // Only where a capital letter follows: "e-mail" and "don't" don't give "email" and "dont".
    for (const task of ["Email the list", "dont count them"]) passed(scanTask(task, { A: ["Sent the e-mail", "don't know"] }));
  });

  it("finds a name typed with a ligature written out, or with ä ö ü as ae oe ue", () => {
    expect(blocked(scanTask("rows for Griffiths", { A: ["Griﬃths"] })).hits).toContainEqual({ column: "A", variant: "exact" });
    expect(blocked(scanTask("rows for Hans Mueller", { A: ["Hans Müller"] })).hits).toContainEqual({ column: "A", variant: "exact" });
    expect(blocked(scanTask("rows for Mueller", { A: ["Hans Müller"] })).hits).toEqual([{ column: "A", variant: "word" }]);
    expect(blocked(scanTask("rows for Schoenfeld", { A: ["Schönfeld"] })).hits).toContainEqual({ column: "A", variant: "exact" });
  });

  it("finds an email address's local part when it has a digit or . _ -", () => {
    const A = ["sunflower77@gmail.com", "bluejay_fan@yahoo.com", "kw.tokyo.1988@example.jp", "maria@example.com"];
    for (const task of ["rows for sunflower77", "rows for sunflower77@gmail", "rows for bluejay_fan", "rows for kw.tokyo.1988"]) {
      expect(blocked(scanTask(task, { A })).hits, task).toEqual([{ column: "A", variant: "word" }]);
    }
    // A plain one is as often an ordinary word or a first name.
    passed(scanTask("rows for maria", { A }));
    passed(scanTask("rows for sunflower770", { A }));
  });

  it("finds a private number written exactly in thousands or millions, or grouped with _", () => {
    const D = ["83500", "1200000"];
    for (const task of ["salary 83.5k", "salary 83.5 K", "salary 83,5k", "salary $83.5K", "salary 83_500", "budget 1.2M", "budget 1200k", "budget 1_200_000"]) {
      // "1_200_000" is also its digits with "_" between them (eighth review).
      expect(blocked(scanTask(task, { D })).hits, task).toContainEqual({ column: "D", variant: "normalized" });
    }
    for (const task of ["distance 83.5km", "salary 83.51k", "salary 183.5k", "budget 1.25M"]) passed(scanTask(task, { D }));
  });
});

describe("auditor: a number shown as ##### (fifth review)", () => {
  it("finds the last 4 digits of a 7+ digit number whose column is too narrow to show it", () => {
    const spec = specOf(
      "Contacts",
      "A1:C3",
      [
        ["Name", "Phone", "Dept"],
        ["Maria Lopez", 8315550199, "Sales"],
        ["Ana Ruiz", 8315552024, "Ops"],
      ],
      { A: "stand_in", B: "range" },
      { text: { "1,1": "##########", "2,1": "##########" }, formats: { "1,1": "(###) ###-####", "2,1": "(###) ###-####" } },
    );
    expect(buildAuditContext([spec], { isToken: () => false }, { model: MODEL }).ranges[0]!.columns[1]!.groups).toEqual(["0199", "2024"]);
    for (const question of ["calls from 555-0199", "the phone ending 0199"]) {
      expect(blocked(auditSpec(spec, question, "structure_only")).hits, question).toEqual([{ column: "B", variant: "word" }]);
    }
    // A last group that reads as a year is skipped, as for any phone number.
    passed(auditSpec(spec, "calls in 2024", "structure_only"));
  });
});

// ---------------------------------------------------------------------------------------------
// Sixth review

describe("auditor: plain numbers typed in lists, formulas and with zeros before them (sixth review)", () => {
  // ZIP codes, salaries and employee numbers stored as numbers, as Excel keeps them (2139 is a
  // ZIP whose leading zero Excel dropped), and a text column of the last 4 digits of cards.
  const staff = () =>
    specOf(
      "Customers",
      "A1:F4",
      [
        ["Customer", "ZIP", "Card last 4", "Salary", "Employee no", "Rate"],
        ["Maria Lopez", 93496, "0199", 83500, 4521, 12.75],
        ["Ana Ruiz", 94103, "4242", 91250, 1187, 18.5],
        ["Kenji Watanabe", 2139, "1881", 69000, 30422, 22.25],
      ],
      { A: "stand_in", B: "range", C: "stand_in", D: "range", E: "stand_in", F: "range" },
    );
  const cases: [string, string][] = [
    ["How many orders came from ZIPs 93496,94103?", "B"],
    ["Why does =IF(B2=93496,1,0) give 0 for her?", "B"],
    ["=SUMPRODUCT((B2:B5={93496,94103})*G2:G5) is wrong", "B"],
    ["Total orders for ZIP 93496,10001 and 60614", "B"],
    ["ZIP 93496.5?", "B"],
    ["How many orders came from 02139?", "B"],
    ["Salaries 83500,91250 - who?", "D"],
    ["Who has card 4242,1881?", "C"],
    ["Who called from 831.555.0199?", "C"],
    ["Hours for employee 004521?", "E"],
    ["Hours for employee #04521?", "E"],
    ["Who earns 12.750 per hour?", "F"],
  ];

  it("blocks the private number when only a comma, a decimal mark or zeros that follow no digit are next to it", () => {
    for (const [question, column] of cases) {
      expect(blocked(auditSpec(staff(), question, "structure_only", false)).hits, question).toContainEqual({ column, variant: "exact" });
    }
  });

  it("replaces it in the question, so the request passes without it", () => {
    for (const [question] of cases) {
      const out = passed(auditSpec(staff(), question, "structure_only"));
      for (const value of ["93496", "2139", "83500", "4242", "0199", "4521", "12.75"]) expect(out.body, question).not.toContain(value);
    }
  });

  it("blocks a number with a digit right next to the private one in the question (tenth review)", () => {
    const cases: [string, string][] = [
      ["id 934961", "B"],
      ["longitude -120.934961", "B"],
      ["code 102139", "B"],
      ["ref 45210", "E"],
      ["ref 1004521", "E"],
      ["rate 12.751", "F"],
      ["total 835001", "D"],
    ];
    for (const [question, column] of cases) {
      expect(blocked(auditSpec(staff(), question, "structure_only", false)).hits, question).toContainEqual({ column, variant: "exact" });
      // Not replaced: the substitution only takes a number no digit continues.
      expect(blocked(auditSpec(staff(), question, "structure_only")).hits, question).toContainEqual({ column, variant: "exact" });
    }
    // Only a cell's own value: "12,75", the rate 12.75 written with a decimal comma, still needs no
    // digit next to it.
    passed(auditSpec(staff(), "About 12,751 orders", "structure_only", false));
    expect(blocked(auditSpec(staff(), "About 12,75 per hour", "structure_only", false)).hits).toContainEqual({ column: "F", variant: "normalized" });
  });

  it("blocks a list of private numbers in a note, which is never replaced", () => {
    const s = staff();
    const noted = { ...s, policies: s.policies.map((p) => (p.letter === "A" ? { ...p, note: "Only ZIPs 93496,94103 matter" } : p)) };
    expect(blocked(auditSpec(noted, "Total per ZIP", "structure_only")).hits).toContainEqual({ column: "B", variant: "exact" });
  });
});

describe("auditor: dates written together, with _ or with other spaces next to the month name (sixth review)", () => {
  const dob = (y: number, m: number, d: number, shown: string) =>
    specOf(
      "HR",
      "A1:B2",
      [
        ["Employee", "Date of birth"],
        ["Maria Lopez", serial1900(y, m, d)],
      ],
      { A: "stand_in", B: "month" },
      { text: { "1,1": shown }, formats: { "1,1": "m/d/yyyy" } },
    );
  const glued = [
    "15MAR1985",
    "15Mar1985",
    "15March1985",
    "15thMarch1985",
    "15MAR85",
    "March15, 1985",
    "Mar15 1985",
    "15_March_1985",
    "March_15,_1985",
    "15_Mar_1985",
    "15 Mar 1985",
    "15 Mar 1985",
    "15﻿Mar﻿1985",
    "15​Mar​1985",
    "March the 15th, 1985",
    "1985MAR15",
  ];

  it("finds a private date typed that way in the question, and the question has it replaced", () => {
    for (const typed of glued) {
      const question = `Who was born ${typed}?`;
      expect(blocked(auditSpec(dob(1985, 3, 15, "3/15/1985"), question, "structure_only", false)).hits, question).toEqual([{ column: "B", variant: "normalized" }]);
      passed(auditSpec(dob(1985, 3, 15, "3/15/1985"), question, "structure_only"));
    }
  });

  it("reads a slash look-alike as a slash in a numeric date", () => {
    for (const typed of ["4∕2∕1990", "4⁄2⁄1990", "4／2／1990"]) {
      const question = `Who was born ${typed}?`;
      // The full-width one also folds to the display text itself.
      expect(blocked(auditSpec(dob(1990, 4, 2, "4/2/1990"), question, "structure_only", false)).hits, question).toContainEqual({ column: "B", variant: "normalized" });
      passed(auditSpec(dob(1990, 4, 2, "4/2/1990"), question, "structure_only"));
    }
  });

  it("finds it in a note, which is never replaced", () => {
    const s = dob(1985, 3, 15, "3/15/1985");
    for (const note of ["Born 15MAR1985", "Born 15_March_1985", "Born March15, 1985"]) {
      const noted = { ...s, policies: s.policies.map((p) => (p.letter === "A" ? { ...p, note } : p)) };
      expect(blocked(auditSpec(noted, "Count per month", "structure_only")).hits, note).toEqual([{ column: "B", variant: "normalized" }]);
    }
  });

  it("still needs a separator between numbers alone, and finds no date where there is none", () => {
    // "1985_03_15" is read since the seventh review.
    for (const question of ["born 031585", "ref 15031985", "March 151985", "Q1mar2024 plan", "15MAR1986"]) {
      passed(auditSpec(dob(1985, 3, 15, "3/15/1985"), question, "structure_only", false));
    }
  });

  it("reads a private date stored as text written together, with _, or as 8 digits", () => {
    const hr = (stored: string) =>
      specOf(
        "HR",
        "A1:B2",
        [
          ["Employee", "Date of birth"],
          ["Maria Lopez", stored],
        ],
        { A: "stand_in", B: "stand_in" },
      );
    const cases: [string, string[]][] = [
      ["15MAR1985", ["born 1985-03-15", "born 3/15/1985", "born March 15, 1985", "born 15 Mar 1985"]],
      ["15_Mar_1985", ["born 1985-03-15", "born 3/15/1985"]],
      ["19910919", ["born 9/19/1991", "born 1991-09-19", "born September 19, 1991"]],
    ];
    for (const [stored, questions] of cases) {
      for (const question of questions) {
        expect(blocked(auditSpec(hr(stored), question, "structure_only", false)).hits, `${stored}: ${question}`).toContainEqual({
          column: "B",
          variant: "normalized",
        });
        passed(auditSpec(hr(stored), question, "substituted"));
      }
    }
  });
});

// ---------------------------------------------------------------------------------------------
// Seventh review

describe("auditor: two dates next to each other in the question (seventh review)", () => {
  // March 15 and September 2, 1985 are private: "1985 - September 2" reads as the second.
  const hr = () =>
    specOf(
      "HR",
      "A1:B3",
      [
        ["Employee", "Date of birth"],
        ["Maria Lopez", serial1900(1985, 3, 15)],
        ["Ana Ruiz", serial1900(1985, 9, 2)],
      ],
      { A: "stand_in", B: "month" },
      { text: { "1,1": "3/15/1985", "2,1": "9/2/1985" }, formats: { "1,1": "m/d/yyyy", "2,1": "m/d/yyyy" } },
    );

  it("sends neither date: both are replaced, with nothing of either left", () => {
    for (const question of [
      "How many people were born March 15, 1985 - September 2, 1985?",
      "How many people were born 15 March 1985 - September 2 1985?",
      "Count DOBs 15 Mar 1985 \u2013 Sep 2 1985",
      "Born 15MAR1985, MAR_15_1985?",
      "Born 15 Mar 1985, March 15 1985?",
    ]) {
      const out = passed(auditSpec(hr(), question, "structure_only"));
      const task = (JSON.parse((JSON.parse(out.body) as { messages: { content: string }[] }).messages[1]!.content) as { task: string }).task;
      expect(task, question).toMatch(/^[^\d]*TEXT_\d{3}[^\d]*(?:TEXT_\d{3}[^\d]*)?$/u);
      for (const piece of ["March", "Mar ", "MAR", "15", "Sep"]) expect(task, question).not.toContain(piece);
    }
  });
});

describe("auditor: plain numbers in the decimals of another number (seventh review)", () => {
  // A customer list: the ZIP is private, latitude and longitude are sent as they are.
  const cols: Col[] = [
    { letter: "A", header: "Customer", type: "text", private: true, treatment: "stand_in", values: ["Maria Lopez"] },
    { letter: "B", header: "Latitude", type: "number", private: false, treatment: "as_is" },
    { letter: "C", header: "Longitude", type: "number", private: false, treatment: "as_is" },
    { letter: "D", header: "ZIP", type: "number", private: true, treatment: "range", values: ["81939", "60112", "93496"] },
    { letter: "E", header: "Phone", type: "text", private: true, treatment: "stand_in", values: ["0199"] },
  ];
  const ctx = { ...makeCtx({ cols }), ranges: [rangeOf(cols, "Data", "A1:E3")] };
  const scan = (lat: unknown, lon: unknown, task = "Count per ZIP") =>
    run(makeBody({ user: desc({ mode: "substituted", cols, sheet: "Data", address: "A1:E3", task, rows: [["PERSON_001", lat, lon, "90000-99999", "TEXT_001"]] }) }), "substituted", ctx);

  it("doesn't find a private ZIP in the decimals of a latitude or longitude", () => {
    passed(scan(43.060112, -116.081939));
    passed(scan(1.93496, -0.093496));
    // Only in a number stored as a number: typed in the question they count (eighth review).
    for (const task of ["near -116.081939", "lat 43.060112?", "x 1.93496"]) expect(blocked(scan(1, 2, task)).hits, task).toContainEqual({ column: "D", variant: "exact" });
  });

  it("still finds it in the last group of a dotted number, a list, after zeros or a comma, and on its own", () => {
    const D = { column: "D", variant: "exact" };
    for (const task of ["ZIPs 10001,93496", "ZIP 060112", "ZIP 93496.0", "value 1,93496", "ZIP 81939"]) expect(blocked(scan(1, 2, task)).hits, task).toContainEqual(D);
    for (const task of ["call 831.555.0199", "call 1.831.555.0199"]) expect(blocked(scan(1, 2, task)).hits, task).toContainEqual({ column: "E", variant: "exact" });
    expect(blocked(scan(81939, 2)).hits).toContainEqual(D);
  });

  it("lets a sheet of customers with private ZIPs and unmarked coordinates through", () => {
    // Each ZIP is also the decimals of another customer's latitude and longitude ("43.010939",
    // "-116.010939"), as happens by chance in a sheet of thousands.
    const zips = Array.from({ length: 60 }, (_, i) => 10939 + i * 1013);
    const rows: CellValue[][] = [["Customer", "ZIP", "Latitude", "Longitude", "Visits"]];
    zips.forEach((zip, i) => {
      const lat = 43 + zips[(i + 13) % 60]! / 1e6;
      const lon = -(116 + (i % 7)) - zips[(i + 7) % 60]! / 1e6;
      rows.push([`Person ${i}`, zip, Math.round(lat * 1e6) / 1e6, Math.round(lon * 1e6) / 1e6, i % 40]);
    });
    const spec = specOf("Contacts", "A1:E61", rows, { A: "stand_in", B: "range" });
    passed(auditSpec(spec, "Count of visits per ZIP", "substituted"));
  });
});

describe("auditor: dates of 8-digit cells count only in free text (seventh review)", () => {
  // Lot numbers are 8 digits that read as dates, next to an unmarked production date column.
  const lots = () => {
    const rows: CellValue[][] = [["Lot", "Produced", "Units"]];
    const text: Record<string, string> = {};
    const formats: Record<string, string> = {};
    for (let i = 0; i < 30; i++) {
      const d = new Date(Date.UTC(2024, 0, 2 + i));
      const [y, m, day] = [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()];
      rows.push([`${y}${String(m).padStart(2, "0")}${String(day).padStart(2, "0")}`, serial1900(y, m, day), 100 + i]);
      text[`${i + 1},1`] = `${m}/${day}/${y}`;
      formats[`${i + 1},1`] = "m/d/yyyy";
    }
    return specOf("Lots", "A1:C31", rows, { A: "stand_in" }, { text, formats });
  };

  it("lets the unmarked dates in the sample rows through", () => {
    passed(auditSpec(lots(), "Total units per month", "substituted"));
    passed(auditSpec(lots(), "Units produced on 1/15/2024?", "substituted"));
  });

  it("still finds such a date typed in the question or a note", () => {
    expect(blocked(auditSpec(lots(), "Units produced on 1/15/2024?", "structure_only", false)).hits).toEqual([{ column: "A", variant: "normalized" }]);
    const s = lots();
    const noted = { ...s, policies: s.policies.map((p) => (p.letter === "C" ? { ...p, note: "Lot of Jan 15, 2024 was recalled" } : p)) };
    expect(blocked(auditSpec(noted, "Total units per month", "substituted")).hits).toEqual([{ column: "A", variant: "normalized" }]);
  });

  it("finds a date of birth stored as 8 digits, text or number, typed in the question", () => {
    for (const stored of ["19910919", 19910919, "09191991", "19091991"] as CellValue[]) {
      const spec = specOf("HR", "A1:B2", [["Employee", "Date of birth"], ["Maria Lopez", stored]], { A: "stand_in", B: typeof stored === "number" ? "range" : "stand_in" });
      for (const question of ["Who was born 9/19/1991?", "Who was born Sept 19, 1991?", "Who was born 1991-09-19?"]) {
        expect(blocked(auditSpec(spec, question, "structure_only", false)).hits, `${stored}: ${question}`).toContainEqual({ column: "B", variant: "normalized" });
        const out = passed(auditSpec(spec, question, "structure_only"));
        expect(out.body).not.toContain("1991");
      }
    }
  });

  it("reads a SAS date-time stored as text", () => {
    const spec = specOf("HR", "A1:B2", [["Employee", "Date of birth"], ["Maria Lopez", "15MAR1985:00:00:00"]], { A: "stand_in", B: "stand_in" });
    expect(blocked(auditSpec(spec, "Who was born 3/15/1985?", "structure_only", false)).hits).toContainEqual({ column: "B", variant: "normalized" });
    passed(auditSpec(spec, "Who was born 3/15/1985?", "structure_only"));
  });
});

describe("auditor: characters that can't be seen (seventh review)", () => {
  const people = () =>
    specOf(
      "HR",
      "A1:C3",
      [
        ["Employee", "Date of birth", "Notes"],
        ["Maria Lopez", serial1900(1985, 3, 15), "ok"],
        ["Ana Ruiz", serial1900(1990, 4, 2), "ok"],
      ],
      { A: "stand_in", B: "month" },
      { text: { "1,1": "3/15/1985", "2,1": "4/2/1990" }, formats: { "1,1": "m/d/yyyy", "2,1": "m/d/yyyy" } },
    );
  const withCell = (cell: string) => {
    const s = people();
    s.data.values[2]![2] = cell;
    s.data.text[2]![2] = cell;
    return s;
  };
  const noted = (note: string) => {
    const s = people();
    return { ...s, policies: s.policies.map((p) => (p.letter === "C" ? { ...p, note } : p)) };
  };
  const marks = ["\u200e", "\u200f", "\u061c", "\u00ad", "\u202a", "\u202e", "\u2066", "\u2069"];

  it("finds a date or a name with them inside, in the question, a note and an unmarked cell", () => {
    const B = { column: "B", variant: "normalized" };
    for (const m of marks) {
      const date = `${m}3/${m}15/${m}1985`;
      expect(blocked(auditSpec(people(), `Who was born ${date}?`, "structure_only", false)).hits, JSON.stringify(m)).toContainEqual(B);
      expect(blocked(auditSpec(noted(`Born ${date}`), "Count per month", "structure_only")).hits, JSON.stringify(m)).toContainEqual(B);
      expect(blocked(auditSpec(withCell(`born ${date}`), "Count per month", "substituted")).hits, JSON.stringify(m)).toContainEqual(B);
      const name = `Ma${m}ria Lo${m}pez`;
      expect(blocked(auditSpec(people(), `Who is ${name}?`, "structure_only", false)).hits, JSON.stringify(m)).toContainEqual({ column: "A", variant: "exact" });
      expect(blocked(auditSpec(noted(`See ${name}`), "Count per month", "structure_only")).hits, JSON.stringify(m)).toContainEqual({ column: "A", variant: "exact" });
      expect(blocked(auditSpec(withCell(`see ${name}`), "Count per month", "substituted")).hits, JSON.stringify(m)).toContainEqual({ column: "A", variant: "exact" });
      // In the question they are replaced.
      passed(auditSpec(people(), `Who is ${name}, born ${date}?`, "structure_only"));
    }
  });

  it("finds a date or a name with form feeds in a note", () => {
    expect(blocked(auditSpec(noted("Born 15\fMar\f1985"), "Count per month", "structure_only")).hits).toEqual([{ column: "B", variant: "normalized" }]);
    // In a date a form feed counts as a space: "3/ 15/ 1985" is the date.
    expect(blocked(auditSpec(noted("Born 3/\f15/\f1985"), "Count per month", "structure_only")).hits).toContainEqual({ column: "B", variant: "normalized" });
    // Matching a value drops it, as a character that can't be seen, and it still separates two
    // words (ninth review): "Ma\fria Lopez" is the name, "Maria\fLopez" holds both of its words.
    expect(blocked(auditSpec(noted("See Ma\fria Lopez"), "Count per month", "structure_only")).hits).toContainEqual({ column: "A", variant: "exact" });
    expect(blocked(auditSpec(noted("See Maria\fLopez"), "Count per month", "structure_only")).hits).toContainEqual({ column: "A", variant: "word" });
    expect(blocked(auditSpec(withCell("born 15\fMar\f1985"), "Count per month", "substituted")).hits).toEqual([{ column: "B", variant: "normalized" }]);
  });

  it("keeps a backslash typed before an f, as in a path", () => {
    const spec = specOf("Files", "A1:B2", [["Owner", "Folder"], ["Maria Lopez", "C:\\finance\\2024"]], { A: "stand_in", B: "stand_in" });
    const s = { ...spec, policies: spec.policies.map((p) => (p.letter === "A" ? { ...p, note: "see C:\\finance\\2024" } : p)) };
    expect(blocked(auditSpec(s, "Count per owner", "structure_only")).hits).toContainEqual({ column: "B", variant: "json" });
    passed(auditSpec(spec, "Count per owner", "structure_only"));
  });
});

describe("auditor: numeric dates with _ or \\ (seventh review)", () => {
  const dob = () =>
    specOf(
      "HR",
      "A1:C3",
      [
        ["Employee", "Date of birth", "Notes"],
        ["Maria Lopez", serial1900(1985, 3, 15), "ok"],
        ["Ana Ruiz", serial1900(1985, 11, 15), "ok"],
      ],
      { A: "stand_in", B: "month" },
      { text: { "1,1": "3/15/1985", "2,1": "11/15/1985" }, formats: { "1,1": "m/d/yyyy", "2,1": "m/d/yyyy" } },
    );
  const B = [{ column: "B", variant: "normalized" }];

  it("finds such a date in the question, a note and an unmarked cell", () => {
    // A backslash typed before a letter is not a JSON escape: "15\Nov\1985" is read too.
    for (const date of ["1985_03_15", "15_03_1985", "3\\15\\1985", "15\\Mar\\1985", "15\\Nov\\1985"]) {
      expect(blocked(auditSpec(dob(), `Who was born ${date}?`, "structure_only", false)).hits, date).toEqual(B);
      passed(auditSpec(dob(), `Who was born ${date}?`, "structure_only"));
      const s = dob();
      const noted = { ...s, policies: s.policies.map((p) => (p.letter === "C" ? { ...p, note: `Born ${date}` } : p)) };
      expect(blocked(auditSpec(noted, "Count per month", "structure_only")).hits, date).toEqual(B);
      const cell = dob();
      cell.data.values[1]![2] = `born ${date}`;
      cell.data.text[1]![2] = `born ${date}`;
      expect(blocked(auditSpec(cell, "Count per month", "substituted")).hits, date).toEqual(B);
    }
  });

  it("still reads a backslash typed before n as a space in a note", () => {
    const s = dob();
    const noted = { ...s, policies: s.policies.map((p) => (p.letter === "C" ? { ...p, note: "Born 3\\n15\\n1985" } : p)) };
    expect(blocked(auditSpec(noted, "Count per month", "structure_only")).hits).toContainEqual(B[0]);
  });
});

describe("auditor: codes typed without their separators (seventh review)", () => {
  const fleet = () =>
    specOf(
      "Fleet",
      "A1:C3",
      [
        ["Driver", "Plate", "Depot"],
        ["Maria Lopez", "AB-1234-CD", "North"],
        ["Ana Ruiz", "XY 987 Z", "South"],
      ],
      { A: "stand_in", B: "stand_in" },
    );

  it("finds a private code typed without its - / . or spaces, as a whole word", () => {
    for (const question of ["Who drives AB1234CD?", "Who drives ab1234cd?", "Who drives XY987Z?"]) {
      expect(blocked(auditSpec(fleet(), question, "structure_only", false)).hits, question).toEqual([{ column: "B", variant: "normalized" }]);
      const out = passed(auditSpec(fleet(), question, "structure_only"));
      expect(out.body.toLowerCase(), question).not.toContain(question.slice(11, -1).toLowerCase());
    }
    // A letter touching a code of 6+ letters and digits doesn't hide it (eighth review).
    expect(blocked(auditSpec(fleet(), "Who drives XAB1234CD?", "structure_only", false)).hits).toEqual([{ column: "B", variant: "normalized" }]);
  });
});

// ---------------------------------------------------------------------------------------------
// Eighth review

/** A spec with a note on one column. */
function withColumnNote(spec: RangeSpec, letter: string, note: string): RangeSpec {
  return { ...spec, policies: spec.policies.map((p) => (p.letter === letter ? { ...p, note } : p)) };
}

/** Audits a structure-only request whose history has one earlier question and this reply. */
function auditWithReply(spec: RangeSpec, reply: string, question = "And per group?"): AuditOutcome {
  const map = createStandInMap([spec.data]);
  const sheetContext = buildSheetContext(spec.data, spec.hasHeaders, [], { sheets: [spec.data.sheet], tables: [] });
  const base = { model: MODEL, specs: [spec], sheetContext, map, rowCap: 50 };
  const first = buildStructureOnly({ ...base, question: "How many rows?", history: [] });
  const history: HistoryMessage[] = [
    { role: "user", content: first.userContent },
    { role: "assistant", content: reply },
  ];
  const built = buildStructureOnly({ ...base, question, history });
  return audit(built.outbound, buildAuditContext([spec], map, { model: MODEL }));
}

describe("auditor: a plain number after digits and a period (eighth review)", () => {
  // Customers: the ZIP is private, the latitude is sent as it is, stored as a number.
  const customers = (cell = "ok", lat: CellValue = 43.1) =>
    specOf(
      "Customers",
      "A1:D3",
      [
        ["Customer", "ZIP", "Latitude", "Notes"],
        ["Maria Lopez", 93496, lat, cell],
        ["Ana Ruiz", 94103, 43.2, "ok"],
      ],
      { A: "stand_in", B: "range" },
    );
  const B = { column: "B", variant: "exact" };

  it("finds it typed after digits and a period in the question, a note or a text cell", () => {
    for (const q of ["ZIP 10001.93496?", "Orders per ZIP for:\n1.93496\n2.10001", "ZIPs #1.93496", "ZIP 0.93496"]) {
      expect(blocked(auditSpec(customers(), q, "structure_only", false)).hits, q).toContainEqual(B);
    }
    expect(blocked(auditSpec(withColumnNote(customers(), "D", "Key ZIPs: 1.93496 2.10001"), "Orders per region", "structure_only")).hits).toContainEqual(B);
    expect(blocked(auditSpec(customers("see 1.93496"), "Orders per region", "substituted")).hits).toContainEqual(B);
    expect(blocked(auditWithReply(customers(), "Most orders come from ZIP 1.93496.")).hits).toContainEqual(B);
  });

  it("sends the question with each such number replaced", () => {
    for (const q of ["ZIPs 93496.94103?", "Orders per ZIP for:\n1.93496\n2.94103", "ZIPs 10001.93496.94103"]) {
      const out = passed(auditSpec(customers(), q, "structure_only"));
      expect(out.body, q).not.toMatch(/93496|94103/u);
    }
  });

  it("still lets it through in the decimals of a number stored as a number", () => {
    passed(auditSpec(customers("ok", 43.093496), "Orders per region", "substituted"));
    passed(auditSpec(customers("ok", -116.094103), "Orders per region", "substituted"));
  });
});

describe("auditor: characters that can't be seen, all of them (eighth review)", () => {
  const people = (cell = "ok") =>
    specOf(
      "HR",
      "A1:E3",
      [
        ["Employee", "Date of birth", "ZIP", "Plate", "Notes"],
        ["Maria Lopez", serial1900(1985, 3, 15), 93496, "AB-1234-CD", cell],
        ["Ana Ruiz", serial1900(1990, 4, 2), 94103, "XY-5678-ZT", "ok"],
      ],
      { A: "stand_in", B: "month", C: "range", D: "stand_in" },
      { text: { "1,1": "3/15/1985", "2,1": "4/2/1990" }, formats: { "1,1": "m/d/yyyy", "2,1": "m/d/yyyy" } },
    );
  const hidden = ["\u200b", "\u200c", "\u200d", "\u2060", "\ufeff", "\u180e", "\u2061", "\u206f", "\ufe0f", "\u{e0020}", "\u3164", "\uffa0", "\u034f", "\b", "\u007f", "\u0001", "\u009f"];
  const texts = (m: string): [string, string][] => [
    [`93${m}496`, "C"],
    [`AB${m}1234${m}CD`, "D"],
    [`3${m}/15/19${m}85`, "B"],
    [`Ma${m}ria`, "A"],
  ];

  it("finds a ZIP, a code, a date or a name with them inside, in a note and an unmarked cell", () => {
    for (const m of hidden) {
      for (const [text, col] of texts(m)) {
        const k = `${JSON.stringify(m)} ${JSON.stringify(text)}`;
        expect(blocked(auditSpec(withColumnNote(people(), "E", `See ${text}`), "Count per month", "structure_only")).hits.map((h) => h.column), k).toContain(col);
        expect(blocked(auditSpec(people(`see ${text}`), "Count per month", "substituted")).hits.map((h) => h.column), k).toContain(col);
      }
    }
  });

  it("finds them written as JSON escapes in an earlier reply", () => {
    // A model reply that escapes what JSON doesn't have to: "\u200b" and "\b" inside a string.
    for (const esc of ["\\u200b", "\\b", "\\u0001", "\\udb40\\udc20"]) {
      const reply = `{"kind":"answer","formula":null,"placement":null,"explanation":"ZIP 93${esc}496 is Ma${esc}ria's","assumptions":[]}`;
      const hits = blocked(auditWithReply(people(), reply)).hits.map((h) => h.column);
      expect(hits, esc).toContain("C");
      expect(hits, esc).toContain("A");
    }
  });

  it("reads a whitespace control as a space", () => {
    // Not a vertical tab or a form feed (ninth review): see the test below.
    for (const m of ["\t", "\n", "\u0085"]) {
      // "93 496", the ZIP grouped with a space.
      expect(blocked(auditSpec(withColumnNote(people(), "E", `See 93${m}496`), "Count per month", "structure_only")).hits, JSON.stringify(m)).toContainEqual({ column: "C", variant: "normalized" });
      expect(blocked(auditSpec(people(`see 93${m}496`), "Count per month", "substituted")).hits, JSON.stringify(m)).toContainEqual({ column: "C", variant: "normalized" });
    }
  });

  it("sends the question with them replaced", () => {
    for (const m of hidden) {
      for (const [text] of texts(m)) passed(auditSpec(people(), `Who has ${text === `Ma${m}ria` ? `${text} Lopez` : text}?`, "structure_only"));
    }
  });
});

describe("auditor: dates with two separators or look-alikes of - and \\ (eighth review)", () => {
  const people = (cell = "ok") =>
    specOf(
      "HR",
      "A1:D3",
      [
        ["Employee", "Date of birth", "Phone", "Notes"],
        ["Maria Lopez", serial1900(1985, 3, 15), "(831) 555-0199", cell],
        ["Ana Ruiz", serial1900(1990, 4, 2), "(415) 555-0142", "ok"],
      ],
      { A: "stand_in", B: "month", C: "stand_in" },
      { text: { "1,1": "3/15/1985", "2,1": "4/2/1990" }, formats: { "1,1": "m/d/yyyy", "2,1": "m/d/yyyy" } },
    );
  const B = { column: "B", variant: "normalized" };

  it("finds such a date in the question, a note and an unmarked cell", () => {
    for (const date of ["3_15/1985", "3_15-1985", "1985-03/15", "3\ufe6815\ufe681985", "3\u29f515\u29f51985", "3\ufe6315\ufe631985", "3\ufe5815\ufe581985", "15\ufe63Mar\ufe631985"]) {
      expect(blocked(auditSpec(people(), `Who was born ${date}?`, "structure_only", false)).hits, date).toContainEqual(B);
      passed(auditSpec(people(), `Who was born ${date}?`, "structure_only"));
      expect(blocked(auditSpec(withColumnNote(people(), "D", `Born ${date}`), "Count per month", "structure_only")).hits, date).toContainEqual(B);
      expect(blocked(auditSpec(people(`born ${date}`), "Count per month", "substituted")).hits, date).toContainEqual(B);
    }
  });

  it("finds a phone number typed with a dash look-alike or _ in an unmarked cell", () => {
    for (const phone of ["831\ufe63555\ufe630199", "831\ufe58555\ufe580199", "831_555_0199"]) {
      expect(blocked(auditSpec(people(`call ${phone}`), "Count per month", "substituted")).hits, phone).toContainEqual({ column: "C", variant: "digits" });
    }
  });
});

describe("auditor: codes with some or other separators (eighth review)", () => {
  const fleet = (d = "ok", e = "ok") =>
    specOf(
      "Fleet",
      "A1:E3",
      [
        ["Driver", "Plate", "Part", "Notes", "More"],
        ["Maria Lopez", "AB-1234-CD", "X7-42-B9", d, e],
        ["Ana Ruiz", "XY-5678-ZT", "Q9-17-C4", "ok", "ok"],
      ],
      { A: "stand_in", B: "stand_in", C: "stand_in" },
    );
  const forms: [string, string][] = [
    ["AB-1234CD", "B"],
    ["AB 1234CD", "B"],
    ["AB\u20131234CD", "B"],
    ["AB.1234.CD", "B"],
    ["AB/1234/CD", "B"],
    ["ab 12 34 cd", "B"],
    ["X7_42_B9", "C"],
    ["xAB1234CD", "B"],
    ["AB1234CDs", "B"],
  ];
  /** A letter touches these, so the question keeps them (only a whole code is replaced) and is blocked. */
  const touching = new Set(["xAB1234CD", "AB1234CDs"]);

  it("finds it in an unmarked cell, a note and the question", () => {
    for (const [code, col] of forms) {
      const hit = { column: col, variant: "normalized" };
      expect(blocked(auditSpec(fleet(`see ${code}`), "Count per driver", "substituted")).hits, code).toContainEqual(hit);
      expect(blocked(auditSpec(withColumnNote(fleet(), "D", `See ${code}`), "Count per driver", "structure_only")).hits, code).toContainEqual(hit);
      expect(blocked(auditSpec(fleet(), `Who has ${code}?`, "structure_only", false)).hits, code).toContainEqual(hit);
      if (touching.has(code)) {
        expect(blocked(auditSpec(fleet(), `Who has ${code}?`, "structure_only")).hits, code).toContainEqual(hit);
        continue;
      }
      const out = passed(auditSpec(fleet(), `Who has ${code}?`, "structure_only"));
      expect(out.body.toLowerCase(), code).not.toContain("1234");
    }
  });

  it("doesn't join letters and digits across JSON values or other characters", () => {
    passed(auditSpec(fleet("AB", "1234CD"), "Count per driver", "substituted"));
    passed(auditSpec(fleet("AB, 1234CD"), "Count per driver", "substituted"));
    passed(auditSpec(withColumnNote(fleet(), "D", "AB; 12-34CD"), "Count per driver", "structure_only"));
  });
});

describe("auditor: 8-digit values are dates only in a column of dates (eighth review)", () => {
  // Invoice numbers 20240001, 20240002 ...: only now and then a date ("20240101").
  const invoices = (isoDates = false) => {
    const rows: CellValue[][] = [["Invoice no", "Customer", "Invoice date", "Amount"]];
    const text: Record<string, string> = {};
    const formats: Record<string, string> = {};
    for (let i = 0; i < 200; i++) {
      const d = new Date(Date.UTC(2024, 0, 1 + Math.floor(i / 2)));
      const [y, m, day] = [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()];
      const iso = `${y}-${String(m).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      rows.push([20240001 + i, i % 2 === 0 ? "Maria Lopez" : "Ana Ruiz", isoDates ? iso : serial1900(y, m, day), 100 + i]);
      if (!isoDates) {
        text[`${i + 1},2`] = `${m}/${day}/${y}`;
        formats[`${i + 1},2`] = "m/d/yyyy";
      }
    }
    return specOf("Invoices", "A1:D201", rows, { A: "stand_in", B: "stand_in" }, { text, formats });
  };

  it("lets dates through that only an invoice number reads as", () => {
    passed(auditSpec(withColumnNote(invoices(), "C", "Invoice dates, from 1/1/2024 on"), "Total per month", "structure_only"));
    // Their digits ("112024", "12024") are no invoice's.
    for (const q of ["Total after 1/1/2024", "Amount on January 1, 2024"]) {
      passed(auditSpec(invoices(), q, "structure_only", false));
      // Nor is the date replaced with an invoice's stand-in.
      const out = passed(auditSpec(invoices(), q, "structure_only"));
      expect(out.body, q).not.toMatch(/ID_\d{3}/u);
    }
    // Sample rows with the invoice dates shown 1/1/2024.
    passed(auditSpec(invoices(), "Total per month", "substituted"));
    // Written 2024-01-01 in the unmarked column, a date holds an invoice number's digits: blocked,
    // as every 8-digit value's digits count wherever they are (tenth review).
    expect(blocked(auditSpec(invoices(true), "Total per month", "substituted")).hits).toEqual([{ column: "A", variant: "digits" }]);
  });

  it("still finds the invoice number itself", () => {
    for (const q of ["Amount of 20240101?", "Amount of 2024 0101?"]) {
      expect(blocked(auditSpec(invoices(), q, "structure_only", false)).hits.map((h) => h.column), q).toContain("A");
    }
  });

  it("still reads a column of dates of birth stored as 8 digits", () => {
    for (const stored of ["19850315", 19850315] as CellValue[]) {
      const spec = specOf("HR", "A1:C3", [["Employee", "Date of birth", "Notes"], ["Maria Lopez", stored, "ok"], ["Ana Ruiz", typeof stored === "number" ? 19900402 : "19900402", "ok"]], { A: "stand_in", B: "stand_in" });
      expect(blocked(auditSpec(withColumnNote(spec, "C", "Born 3/15/1985"), "Count", "structure_only")).hits, String(stored)).toEqual([{ column: "B", variant: "normalized" }]);
      expect(blocked(auditSpec(spec, "Born 1985-03-15?", "structure_only", false)).hits, String(stored)).toContainEqual({ column: "B", variant: "normalized" });
    }
  });
});

describe("auditor: grouped numbers in free text (eighth review)", () => {
  const cols = ORDERS.map((c) => (c.letter === "D" ? { ...c, values: ["93496", "94103", "5390"] } : c));
  const ctx = makeCtx({ cols });
  const withNoteOn = (note: string) => run(makeBody({ user: withNote(desc({ cols }), "B", note) }), "structure_only", ctx);

  it("finds a grouped form in free text with a decimal part after it", () => {
    for (const note of ["ZIPs 93 496,94 103", "93,496.5", "5,390.25", "93 496 5", "5.390,25"]) {
      expect(blocked(withNoteOn(note)).hits, note).toEqual([{ column: "D", variant: "normalized" }]);
    }
    expect(blocked(withNoteOn("ZIPs 94 103,93 496")).hits).toEqual([{ column: "D", variant: "normalized" }]);
  });

  it("doesn't find one a digit, or the same separator and 3 digits, continue", () => {
    for (const note of ["5,390,000", "1 093 496", "15,390", "5.390.000"]) passed(withNoteOn(note));
  });

  it("keeps the rule outside free text", () => {
    // In a sample-row cell a grouped form followed by other decimals is still another number.
    const spec = specOf("Customers", "A1:C3", [["Customer", "ZIP", "Notes"], ["Maria Lopez", 93496, "93 496,94 103"], ["Ana Ruiz", 94103, "ok"]], { A: "stand_in", B: "range" });
    passed(auditSpec(spec, "Orders per region", "substituted"));
  });
});

// ---------------------------------------------------------------------------------------------
// Ninth review

describe("auditor: a dropped character still separates words and numbers (ninth review)", () => {
  // Customers: full names, ZIP numbers and salaries ("$83,500") private; Notes unmarked.
  const customers = (cell = "ok") =>
    specOf(
      "Customers",
      "A1:D4",
      [
        ["Customer", "ZIP", "Salary", "Notes"],
        ["Kenji Watanabe", 93496, 83500, cell],
        ["Maria Lopez", 94103, 91200, "ok"],
        ["Ana Ruiz", 10001, 72000, "ok"],
      ],
      { A: "stand_in", B: "stand_in", C: "stand_in" },
      { text: { "1,2": "$83,500", "2,2": "$91,200", "3,2": "$72,000" }, formats: { "1,2": "$#,##0", "2,2": "$#,##0", "3,2": "$#,##0" } },
    );
  const hidden = ["\u200b", "\u200c", "\u2060", "\ufeff", "\u180e", "\u3164", "\uffa0", "\ufe0f", "\u{e0020}", "\u034f", "\u0001", "\b", "\u007f", "\u009f", "\u00ad", "\u200e", "\u202a", "\u2066", "\f", "\v"];
  const forms = (x: string): [string, string][] => [
    [`Mr${x}Watanabe`, "A"],
    [`Kenji${x}san`, "A"],
    [`93496${x}94103`, "B"],
    [`10001${x}2`, "B"],
    [`83,500${x}91,200`, "C"],
  ];

  it("finds the value next to it in a note, an unmarked cell and an earlier reply", () => {
    for (const x of hidden) {
      for (const [text, col] of forms(x)) {
        const k = `${JSON.stringify(x)} ${JSON.stringify(text)}`;
        expect(blocked(auditSpec(withColumnNote(customers(), "D", `Key: ${text}`), "Count per region", "structure_only")).hits.map((h) => h.column), k).toContain(col);
        expect(blocked(auditSpec(customers(`key ${text}`), "Count per region", "substituted")).hits.map((h) => h.column), k).toContain(col);
        const reply = JSON.stringify({ kind: "answer", formula: null, placement: null, explanation: `The top one is ${text}.`, assumptions: [] });
        expect(blocked(auditWithReply(customers(), reply)).hits.map((h) => h.column), k).toContain(col);
      }
    }
  });

  it("finds it in the question, and sends a whole value there replaced", () => {
    for (const x of hidden) {
      for (const [text, col] of forms(x)) {
        expect(blocked(auditSpec(customers(), `Orders for ${text}?`, "structure_only", false)).hits.map((h) => h.column), `${JSON.stringify(x)} ${text}`).toContain(col);
      }
      // ZIPs and salaries are whole values: replaced. A surname alone is a part of one: blocked.
      const out = passed(auditSpec(customers(), `Orders for ZIPs 93496${x}94103 and 10001${x}2?`, "structure_only"));
      expect(out.body, JSON.stringify(x)).not.toMatch(/93496|94103|10001/u);
      blocked(auditSpec(customers(), `Orders for Mr${x}Watanabe?`, "structure_only"));
    }
  });

  it("still finds a value with one inside it, a form feed and a vertical tab too", () => {
    for (const x of hidden) {
      for (const text of [`Wata${x}nabe`, `934${x}96`]) {
        const k = `${JSON.stringify(x)} ${JSON.stringify(text)}`;
        expect(blocked(auditSpec(withColumnNote(customers(), "D", `Key: ${text}`), "Count per region", "structure_only")).hits.length, k).toBeGreaterThan(0);
        expect(blocked(auditSpec(customers(`key ${text}`), "Count per region", "substituted")).hits.length, k).toBeGreaterThan(0);
      }
      passed(auditSpec(customers(), `Orders for ZIP 934${x}96?`, "structure_only"));
    }
  });
});

describe("auditor: a dropped character ends a date (ninth review)", () => {
  const people = () =>
    specOf(
      "HR",
      "A1:C3",
      [
        ["Employee", "Date of birth", "Notes"],
        ["Maria Lopez", serial1900(1985, 3, 15), "ok"],
        ["Ana Ruiz", serial1900(1990, 4, 2), "ok"],
      ],
      { A: "stand_in", B: "stand_in" },
      { text: { "1,1": "3/15/1985", "2,1": "4/2/1990" }, formats: { "1,1": "m/d/yyyy", "2,1": "m/d/yyyy" } },
    );
  const B = { column: "B", variant: "normalized" };

  it("finds the date in a note and the question, and sends the question with it replaced", () => {
    for (const x of ["\u200b", "\u2060", "\ufeff", "\u00ad", "\u0001", "\b"]) {
      for (const text of [`3/15/1985${x}5`, `12${x}3/15/1985`, `15${x}03${x}1985`]) {
        const k = `${JSON.stringify(x)} ${JSON.stringify(text)}`;
        expect(blocked(auditSpec(withColumnNote(people(), "C", `Born ${text}`), "Count", "structure_only")).hits, k).toContainEqual(B);
        expect(blocked(auditSpec(people(), `Who was born ${text}?`, "structure_only", false)).hits, k).toContainEqual(B);
        passed(auditSpec(people(), `Who was born ${text}?`, "structure_only"));
      }
    }
  });
});

describe("auditor: space-grouped numbers in a list (ninth review)", () => {
  const staff = () =>
    specOf(
      "Staff",
      "A1:C4",
      [
        ["Employee", "Salary", "Notes"],
        ["Maria Lopez", 83500, "ok"],
        ["Ana Ruiz", 390000, "ok"],
        ["Kenji Watanabe", 91200, "ok"],
      ],
      { A: "stand_in", B: "stand_in" },
      { text: { "1,1": "$83,500", "2,1": "$390,000", "3,1": "$91,200" }, formats: { "1,1": "$#,##0", "2,1": "$#,##0", "3,1": "$#,##0" } },
    );
  const B = { column: "B", variant: "normalized" };

  it("finds one whatever number is next to it, a digit right next to it aside", () => {
    for (const note of ["Bands: 40 000 83 500 120 000", "Which exist?\n40 000\n83 500\n", "Top 10 83 500", "Paid 1 83\u00a0500", "Total 5 390 000", "40\u202f000 91\u2009200"]) {
      expect(blocked(auditSpec(withColumnNote(staff(), "C", note), "Count", "structure_only")).hits, note).toContainEqual(B);
      expect(blocked(auditSpec(staff(), note, "structure_only", false)).hits, note).toContainEqual(B);
      passed(auditSpec(staff(), note, "structure_only"));
    }
    for (const note of ["Total 183 500", "Total 83 5001"]) passed(auditSpec(withColumnNote(staff(), "C", note), "Count", "structure_only"));
  });

  it("keeps the rule for , . and ' grouping", () => {
    for (const note of ["Total 5,390,000", "Total 5.390.000", "Total 5'390'000"]) passed(auditSpec(withColumnNote(staff(), "C", note), "Count", "structure_only"));
  });
});

describe("auditor: 8-digit columns with placeholders (ninth review)", () => {
  const born = (placeholder: CellValue, k: number) => {
    const rows: CellValue[][] = [["Employee", "DOB", "Notes"]];
    for (let i = 0; i < 40; i++) rows.push([`Person ${i}`, i >= 40 - k ? placeholder : `${1950 + i}0${1 + (i % 9)}${10 + (i % 18)}`, "ok"]);
    return specOf("Staff", "A1:C41", rows, { A: "stand_in", B: "stand_in" });
  };
  const B = { column: "B", variant: "normalized" };

  it("reads a column of dates of birth with a few placeholders", () => {
    for (const ph of ["00000000", "99999999", "19000000"]) {
      for (const k of [3, 8]) {
        const s = born(ph, k);
        expect(blocked(auditSpec(withColumnNote(s, "C", "Oldest born 1953-04-13"), "Count", "structure_only")).hits, `${ph} ${k}`).toContainEqual(B);
        expect(blocked(auditSpec(s, "Who was born 4/13/1953?", "structure_only", false)).hits, `${ph} ${k}`).toContainEqual(B);
        const out = passed(auditSpec(s, "Who was born 1953-04-13?", "structure_only"));
        expect(out.body, `${ph} ${k}`).not.toContain("1953-04-13");
      }
    }
  });
});

describe("auditor: codes (ninth review)", () => {
  it("finds an aliased header typed without its separators", () => {
    const s = specOf("Budget", "A1:C3", [["Employee", "Falcon X9-2024", "Dept"], ["Maria Lopez", 100, "Ops"], ["Ana Ruiz", 101, "Ops"]], { A: "stand_in" }, { alias: { B: "Project" } });
    expect(blocked(auditSpec(s, "Budget for FalconX92024?", "structure_only", false)).hits.map((h) => h.column)).toContain("B");
    const out = passed(auditSpec(s, "Budget for FalconX92024?", "structure_only"));
    expect(out.body).toContain("Budget for Project?");
  });

  it("doesn't read a cell range as a code, nor a stand-in and the number after it", () => {
    const bins = specOf("Stock", "A1:B3", [["Bin", "Qty"], ["A10-B20", 5], ["C12-D22", 7]], { A: "stand_in" });
    for (const q of ["Sum of A10:B20", "What does =SUM(A10:B20) give?"]) {
      const out = passed(auditSpec(bins, q, "structure_only"));
      expect(out.body, q).toContain(q.replace(/"/gu, '\\"'));
    }
    const rows: CellValue[][] = [["Case", "Hours"], ...Array.from({ length: 20 }, (_, i): CellValue[] => [`ID-${String(i + 1).padStart(4, "0")}`, 3 + i])];
    const cases = specOf("Cases", "A1:B21", rows, { A: "stand_in" });
    for (const q of ["Hours for ID-0001 2024?", "Compare ID-0001 12 hours with ID-0002", "Hours for ID-0001 17?"]) {
      const out = passed(auditSpec(cases, q, "structure_only"));
      expect(out.body, q).not.toContain("ID-00");
    }
    // Typed by itself, the code is still found.
    expect(blocked(auditSpec(cases, "Hours for ID 0012?", "structure_only", false)).hits.map((h) => h.column)).toContain("A");
    // A colon that isn't a cell range's still separates the parts of a code, in an unmarked cell too.
    const fleet = (cell: string) => specOf("Fleet", "A1:C2", [["Driver", "Plate", "Notes"], ["Maria Lopez", "AB-1234-CD", cell]], { A: "stand_in", B: "stand_in" });
    for (const code of ["AB:1234:CD", "AB\uff1a1234\uff1aCD"]) {
      expect(blocked(auditSpec(fleet(`see ${code}`), "Count", "substituted")).hits, code).toContainEqual({ column: "B", variant: "normalized" });
    }
    passed(auditSpec(fleet("see A10:B20"), "Count", "substituted"));
  });

  it("doesn't replace a code inside other letters", () => {
    // "area 10023" holds the letters and digits of "A10023": the number is replaced, not "a 10023".
    const members = specOf("Members", "A1:B3", [["Member", "Fee"], ["A10023", 20], ["A10150", 25]], { A: "stand_in" });
    const out = passed(auditSpec(members, "How many members live in area 10023?", "structure_only"));
    expect(out.body).toMatch(/live in area [A-Z]+_\d{3}\?/u);
  });
});

// ---------------------------------------------------------------------------------------------
// Tenth review

describe("auditor: an 8-digit value typed as a date (tenth review)", () => {
  // Invoice numbers 20240001 ... 20240200: only now and then a date, so the column isn't read as dates.
  const invoices = () =>
    specOf(
      "Invoices",
      "A1:C201",
      [["Invoice no", "Amount", "Notes"], ...Array.from({ length: 200 }, (_, i): CellValue[] => [20240001 + i, 100 + i, "ok"])],
      { A: "stand_in" },
    );
  const typed = ["2024-01-15", "2024-01-15​", "2024-01-15‎", "2024/01/15", "2024.01.15"];

  it("finds the invoice number's digits typed as a date, a character that can't be seen after it or not", () => {
    for (const t of typed) {
      expect(blocked(auditSpec(invoices(), `Total for invoice ${t}?`, "structure_only", false)).hits, t).toEqual([{ column: "A", variant: "digits" }]);
    }
    expect(blocked(auditSpec(invoices(), "Total for invoice 2024-01-15​5", "structure_only", false)).hits).toEqual([{ column: "A", variant: "digits" }]);
    // A note is never replaced.
    expect(blocked(auditSpec(withColumnNote(invoices(), "C", "Invoice 2024-01-15 is disputed"), "Count", "structure_only")).hits).toEqual([
      { column: "A", variant: "digits" },
    ]);
  });

  it("has it replaced in the question", () => {
    for (const t of typed) {
      const out = passed(auditSpec(invoices(), `Total for invoice ${t}?`, "substituted"));
      expect(out.body, t).not.toContain("2024");
      expect(out.body, t).toMatch(/Total for invoice ID_\d{3}/u);
    }
  });

  it("still passes dates whose digits are no invoice's", () => {
    for (const q of ["Total after 1/1/2024", "Amount on January 15, 2024", "Monthly totals for 2024"]) {
      passed(auditSpec(invoices(), q, "structure_only", false));
    }
    passed(auditSpec(withColumnNote(invoices(), "C", "from 1/1/2024 on"), "Count", "structure_only"));
  });
});

describe("auditor: a plain number with a digit next to it in free text (tenth review)", () => {
  // The ZIP and a 6-digit account number are private; Region and Extra are sent as they are.
  const customers = (extra: CellValue = "ok", header = "Extra", sheet = "Customers") =>
    specOf(
      sheet,
      "A1:E3",
      [
        ["Customer", "ZIP", "Account", "Region", header],
        ["Maria Lopez", 93496, 483920, "West", extra],
        ["Ana Ruiz", 94103, 571046, "East", "ok"],
      ],
      { A: "stand_in", B: "range", C: "range" },
    );
  const cases: [string, string][] = [
    ["How many orders came from 934961234?", "B"],
    ["Orders from 93496 1234 or 934961234", "B"],
    ["Payments for reference 4839201?", "C"],
    ["Payments for 48392001?", "C"],
    ["Payments for REF4839201?", "C"],
  ];

  it("finds it in the question, whether or not the question is replaced first", () => {
    for (const [question, column] of cases) {
      expect(blocked(auditSpec(customers(), question, "structure_only", false)).hits, question).toContainEqual({ column, variant: "exact" });
      expect(blocked(auditSpec(customers(), question, "substituted")).hits, question).toContainEqual({ column, variant: "exact" });
    }
  });

  it("finds it in a note, a header, a sheet name and an earlier reply", () => {
    expect(blocked(auditSpec(withColumnNote(customers(), "D", "Mostly 934961234 and 941031234"), "Count", "structure_only")).hits).toEqual([
      { column: "B", variant: "exact" },
    ]);
    expect(blocked(auditSpec(customers("ok", "Ref 4839201"), "Count", "structure_only")).hits).toContainEqual({ column: "C", variant: "exact" });
    expect(blocked(auditSpec(customers("ok", "Extra", "ZIP 934961234"), "Count", "structure_only")).hits).toContainEqual({ column: "B", variant: "exact" });
    const reply = JSON.stringify({ kind: "answer", formula: null, placement: null, explanation: "Most orders came from 934961234.", assumptions: [] });
    expect(blocked(auditWithReply(customers(), reply)).hits).toContainEqual({ column: "B", variant: "exact" });
  });

  it("passes it in a number cell and finds it in a text cell (eleventh review)", () => {
    for (const extra of [934961234, 4839201, 1.93496]) passed(auditSpec(customers(extra), "Count", "substituted"));
    for (const extra of ["key 934961234", "4839201", "Transfer to 48392001"]) {
      expect(blocked(auditSpec(customers(extra), "Count", "substituted")).hits.length, extra).toBeGreaterThan(0);
    }
  });

  it("still passes numbers in free text that hold no private one", () => {
    for (const q of ["How many orders came from 9349?", "Payments over 48392?", "Total for 2024", "Rows 2 to 1000"]) {
      passed(auditSpec(customers(), q, "structure_only", false));
    }
  });
});

// ---------------------------------------------------------------------------------------------
// Where a hit was found (AuditOutcome.locations): offsets and a value-free field path, for the
// pane's Log only. They add nothing to the result and change no verdict.

describe("auditor: where each hit was found", () => {
  function locationsOf(r: AuditOutcome): readonly AuditLocation[] {
    if (r.ok) throw new Error("expected the audit to block, but it passed");
    return r.locations;
  }
  /** The first location of a hit. */
  function locate(r: AuditOutcome, column: string, variant: AuditLocation["variant"]): AuditLocation {
    const loc = locationsOf(r).find((l) => l.column === column && l.variant === variant);
    if (!loc) throw new Error(`no location for ${column} ${variant}`);
    return loc;
  }
  const slice = (body: string, l: AuditLocation) => body.slice(l.start, l.end);
  const inTask = { message: 1, role: "user", path: ["task"], within: "string" };

  it("gives one location per hit, in the order of the hits, with no text in it", () => {
    const body = makeBody({ user: desc({ task: "Totals for Kenji Watanabe on 3/15/2026, call 8315550199" }) });
    const r = run(body);
    const locations = locationsOf(r);
    expect(locations.map(({ column, variant }) => ({ column, variant }))).toEqual(blocked(r).hits);
    for (const l of locations) expect(Object.keys(l).sort()).toEqual(["column", "end", "field", "origin", "start", "variant"]);
    expect(JSON.stringify(locations)).not.toMatch(/kenji|watanabe|8315550199|2026/i);
  });

  it("a whole value in the question: the slice is the value, in the task", () => {
    const body = makeBody({ user: desc({ task: "Total for Maria Lopez by month" }) });
    const exact = locate(run(body), "A", "exact");
    expect(slice(body, exact)).toBe("Maria Lopez");
    expect(exact).toMatchObject({ origin: "value", field: inTask });
    expect(slice(body, locate(run(body), "A", "word"))).toBe("Maria");
  });

  it("a word of a multi-word value in the header of another column", () => {
    const cols = ORDERS.map((c) => (c.letter === "B" ? { ...c, sentHeader: "Watanabe region" } : c));
    const body = makeBody({ user: desc({ cols }) });
    const r = run(body);
    expect(blocked(r).hits).toEqual([{ column: "A", variant: "word" }]);
    const word = locate(r, "A", "word");
    expect(slice(body, word)).toBe("Watanabe");
    expect(word).toMatchObject({ origin: "value", field: { message: 1, role: "user", path: ["columns", 1, "header"] } });
  });

  it("a date phrase, however it is written", () => {
    for (const phrase of ["March 15, 2026", "15.03.2026", "2026-03-15"]) {
      const body = makeBody({ user: desc({ task: `Signups on ${phrase} by region` }) });
      const date = locate(run(body), "F", "normalized");
      expect(slice(body, date), phrase).toBe(phrase);
      expect(date.field).toEqual(inTask);
    }
  });

  it("a digit group, and digits typed with other punctuation", () => {
    const group = makeBody({ user: desc({ task: "The customer whose phone ends in 0199" }) });
    expect(slice(group, locate(run(group), "C", "word"))).toBe("0199");
    const dotted = makeBody({ user: desc({ task: "The customer on 831.555.0199" }) });
    const digits = locate(run(dotted), "C", "digits");
    expect(slice(dotted, digits)).toBe("831.555.0199");
    expect(digits.field).toEqual(inTask);
  });

  it("a value escaped twice in the question: the slice is its escaped form", () => {
    const body = makeBody({ user: desc({ task: 'Orders for Ana "AJ" Ruiz' }) });
    const json = locate(run(body), "A", "json");
    const twice = JSON.stringify(JSON.stringify('Ana "AJ" Ruiz').slice(1, -1)).slice(1, -1);
    expect(slice(body, json)).toBe(twice);
    expect(json.field).toEqual(inTask);
  });

  it("the original header of a renamed column: origin alias, and the column isn't called private", () => {
    const diagnosis: Col = { letter: "F", header: "Diagnosis", sentHeader: "Field F", type: "text", private: true, treatment: "exclude", values: [], aliasOriginal: "Diagnosis" };
    const cols: Col[] = [...ORDERS.slice(0, 4), { ...ORDERS[4]!, note: "sorted by diagnosis" }, diagnosis];
    const body = makeBody({ user: desc({ cols }) });
    const r = run(body, "structure_only", makeCtx({ cols }));
    expect(blocked(r).hits).toEqual([{ column: "F", variant: "exact" }]);
    expect(blocked(r).error).toBe("Text from column F's original header is in the request. Nothing was sent.");
    const alias = locate(r, "F", "exact");
    expect(slice(body, alias)).toBe("diagnosis");
    expect(alias).toMatchObject({ origin: "alias", field: { message: 1, role: "user", path: ["columns", 4, "note"] } });

    // A column that isn't private, renamed: its original header in the question.
    const region: Col[] = ORDERS.map((c) => (c.letter === "B" ? { ...c, sentHeader: "Area", aliasOriginal: "Region" } : c));
    const rb = run(makeBody({ user: desc({ cols: region, task: "Total per region" }) }), "structure_only", makeCtx({ cols: region }));
    expect(blocked(rb).error).toBe("Text from column B's original header is in the request. Nothing was sent.");
    expect(locate(rb, "B", "exact")).toMatchObject({ origin: "alias", field: inTask });

    // With a private value too, each is named for what it is.
    const both = run(makeBody({ user: desc({ cols: region, task: "Total per region for Kenji Watanabe" }) }), "structure_only", makeCtx({ cols: region }));
    expect(blocked(both).error).toBe(
      "Text from private column A is in the request. Text from column B's original header is in the request. Nothing was sent.",
    );
    expect(locate(both, "A", "exact").origin).toBe("value");
  });

  it("names the field: a sample row cell, the sheet name, a lookup range's header, earlier turns", () => {
    const rows = ROWS.map((r) => [...r]);
    rows[2]![1] = "Maria Lopez";
    const inRow = makeBody({ user: desc({ mode: "substituted", rows }) });
    const cell = locate(run(inRow, "substituted"), "A", "exact");
    expect(slice(inRow, cell)).toBe("Maria Lopez");
    expect(cell.field).toEqual({ message: 1, role: "user", path: ["rows", 2, 1], within: "string" });

    const sheetCtx = { ...makeCtx(), ranges: [rangeOf(ORDERS, "Watanabe Orders")] };
    const inSheet = run(makeBody({ user: desc({ sheet: "Watanabe Orders" }) }), "structure_only", sheetCtx);
    expect(locate(inSheet, "A", "word").field).toEqual({ message: 1, role: "user", path: ["sheet"], within: "string" });

    const regionCols: Col[] = [
      { letter: "A", header: "Region", type: "text", private: false, treatment: "as_is" },
      { letter: "B", header: "Manager", sentHeader: "Kenji Watanabe", type: "text", private: false, treatment: "as_is" },
    ];
    const lookup = {
      sheet: "Regions",
      table: null,
      range: { address: "A1:B5", header_row: 1, first_data_row: 2, last_data_row: 5 },
      columns: columnsJson(regionCols, 4),
    };
    const ctx = makeCtx({ extraRanges: [rangeOf(regionCols, "Regions", "A1:B5")] });
    const inLookup = run(makeBody({ user: desc({ contextRanges: [lookup] }) }), "structure_only", ctx);
    expect(locate(inLookup, "A", "exact").field).toEqual({
      message: 1,
      role: "user",
      path: ["context_ranges", 0, "columns", 1, "header"],
      within: "string",
    });

    const reply = '{"kind":"answer","formula":null,"placement":null,"explanation":"Kenji Watanabe has the most orders.","assumptions":[]}';
    const inReply = run(makeBody({ history: [{ role: "user", content: desc() }, { role: "assistant", content: reply }] }));
    expect(locate(inReply, "A", "exact").field).toEqual({ message: 2, role: "assistant", path: ["explanation"], within: "string" });
    const plain = run(makeBody({ history: [{ role: "user", content: desc() }, { role: "assistant", content: "Here it is for Kenji Watanabe." }] }));
    expect(locate(plain, "A", "exact").field).toEqual({ message: 2, role: "assistant", path: [], within: "text" });
    const earlier = run(
      makeBody({
        history: [
          { role: "user", content: desc({ task: "Orders for Kenji Watanabe" }) },
          { role: "assistant", content: '{"kind":"clarify","formula":null,"placement":null,"explanation":"Which month?","assumptions":[]}' },
        ],
      }),
    );
    expect(locate(earlier, "A", "exact").field).toEqual(inTask);
  });

  it("says what a match is in: a string of the turn's JSON under any key, a number of it, or plain text", () => {
    // A number cell of a sample row: in the description's JSON, but not in a string.
    const amounts: Col[] = ORDERS.map((c) => (c.letter === "B" ? { ...c, type: "number" } : c));
    const rows = ROWS.map((r) => [...r]);
    rows[1]![1] = 83500;
    const inNumber = makeBody({ user: desc({ mode: "substituted", cols: amounts, rows }) });
    const number = locationsOf(run(inNumber, "substituted", makeCtx({ cols: amounts }))).find((l) => l.column === "D");
    expect(number && slice(inNumber, number)).toBe("83500");
    expect(number?.field).toEqual({ message: 1, role: "user", path: ["rows", 1, 1], within: "json" });

    // A reply key the path stops at (it could hold typed text) is still a string of the reply.
    const reply = '{"kind":"answer","formula":null,"placement":null,"explanation":"x","assumptions":[],"extra":"Ask Kenji Watanabe"}';
    const inExtra = run(makeBody({ history: [{ role: "user", content: desc() }, { role: "assistant", content: reply }] }));
    expect(locate(inExtra, "A", "exact").field).toEqual({ message: 2, role: "assistant", path: [], within: "string" });

    // Text around a reply's JSON is not in it.
    const around = run(makeBody({ history: [{ role: "user", content: desc() }, { role: "assistant", content: `Kenji Watanabe: ${reply}` }] }));
    expect(locate(around, "A", "exact").field).toEqual({ message: 2, role: "assistant", path: [], within: "text" });
  });

  it("a renamed private column whose value and original header both match: one hit, both matches, called private", () => {
    const nickname: Col = {
      letter: "E",
      header: "Nickname",
      sentHeader: "Handle",
      type: "text",
      private: true,
      treatment: "exclude",
      values: ["Lulubelle"],
      aliasOriginal: "Nickname",
    };
    const cols: Col[] = [...ORDERS.slice(0, 2), { ...ORDERS[2]!, note: "nickname Lulubelle" }, ORDERS[3]!, nickname, ORDERS[5]!];
    const body = makeBody({ user: desc({ cols }) });
    const ctx = makeCtx({ cols });
    for (const strategy of ["indexOf", "ahoCorasick"] as const) {
      const r = run(body, "structure_only", ctx, strategy);
      // The hit and the sentence are what they were: the column's own value is in the request.
      expect(blocked(r).hits, strategy).toEqual([{ column: "E", variant: "exact" }]);
      expect(blocked(r).error, strategy).toBe("Text from private column E is in the request. Nothing was sent.");
      const locations = locationsOf(r);
      expect(locations.map((l) => [l.column, l.variant]), strategy).toEqual([
        ["E", "exact"],
        ["E", "exact"],
      ]);
      expect(new Set(locations.map((l) => `${l.origin}:${slice(body, l)}`)), strategy).toEqual(new Set(["alias:nickname", "value:Lulubelle"]));
      for (const l of locations) expect(l.field, strategy).toEqual({ message: 1, role: "user", path: ["columns", 2, "note"], within: "string" });
    }
    // Aho-Corasick reports in text order, so the header leads; indexOf finds the values first.
    expect(locationsOf(run(body, "structure_only", ctx, "ahoCorasick")).map((l) => l.origin)).toEqual(["alias", "value"]);
    expect(locationsOf(run(body, "structure_only", ctx, "indexOf")).map((l) => l.origin)).toEqual(["value", "alias"]);
  });

  it("the Aho-Corasick matcher gives the same locations when each hit occurs once", () => {
    const body = makeBody({ user: desc({ task: "Totals for Kenji Watanabe on 3/15/2026, call 8315550199" }) });
    expect(locationsOf(run(body, "structure_only", makeCtx(), "ahoCorasick"))).toEqual(locationsOf(run(body, "structure_only", makeCtx(), "indexOf")));
  });

  it("a request refused before the scan, an exception, or an unreplaced cell alone has no locations", () => {
    expect(locationsOf(run(makeBody({ model: "other/model" })))).toEqual([]);
    const throwing = makeCtx({
      isToken: () => {
        throw new Error("Maria Lopez");
      },
    });
    expect(locationsOf(run(makeBody({ user: desc({ mode: "substituted" }) }), "substituted", throwing))).toEqual([]);
    const rows = ROWS.map((r) => [...r]);
    rows[0]![0] = "Al";
    const structural = run(makeBody({ user: desc({ mode: "substituted", rows }) }), "substituted");
    expect(blocked(structural).structural).toEqual([{ column: "A" }]);
    expect(locationsOf(structural)).toEqual([]);
  });
});
