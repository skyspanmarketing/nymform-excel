// metrics.ts (spec §7.14, T-M2): strict schema for local evaluation records. Invalid records are
// dropped without throwing; the export holds EvalRecord fields only.
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { EvalRecordSchema, MetricsStore, MODEL_ID_RE, validateEvalRecord } from "../src/core/metrics";
import type { EvalRecord } from "../src/core/types";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
// Built by concatenation, as in invariants.test.ts.
const RAW_MARKER = "raw" + "_bench";

const FIELDS = [
  "build",
  "model",
  "mode",
  "rows",
  "columns",
  "privateColumns",
  "treatments",
  "bytes",
  "tokens",
  "latencyMs",
  "audit",
  "auditReasons",
  "gate",
  "gateCodes",
  "replyKind",
  "retries",
  "inserted",
  "copied",
  "excelError",
  "rating",
];

function record(extra: Partial<EvalRecord> = {}): EvalRecord {
  return {
    build: "0.1.0-alpha (test)",
    model: "openai/gpt-6-luna",
    mode: "substituted",
    rows: 500,
    columns: 6,
    privateColumns: 3,
    treatments: { as_is: 3, stand_in: 2, range: 1, month: 0, exclude: 0 },
    bytes: 4096,
    tokens: { prompt: 1200, completion: 80 },
    latencyMs: 2345,
    audit: "pass",
    auditReasons: [],
    gate: "pass",
    gateCodes: [],
    replyKind: "formula",
    retries: 0,
    inserted: true,
    copied: false,
    excelError: false,
    rating: "worked",
    ...extra,
  };
}

function rejected(value: unknown): void {
  const store = new MetricsStore();
  expect(store.add(value as EvalRecord)).toBeNull();
  expect(store.records()).toEqual([]);
  expect(validateEvalRecord(value)).toBeNull();
}

describe("EvalRecord schema", () => {
  it("accepts a full record and a minimal one", () => {
    const store = new MetricsStore();
    expect(store.add(record())).toEqual(expect.any(Number));
    const minimal: EvalRecord = {
      build: "0.1.0-alpha (abc1234)",
      model: "openai/gpt-6-luna",
      mode: "structure_only",
      rows: 0,
      columns: 1,
      privateColumns: 0,
      treatments: { as_is: 1, stand_in: 0, range: 0, month: 0, exclude: 0 },
      bytes: 10,
      audit: "block",
      auditReasons: ["exact", "structural", "error"],
      gateCodes: [],
      retries: 0,
      inserted: false,
      copied: false,
    };
    expect(store.add(minimal)).toEqual(expect.any(Number));
    expect(store.records()).toEqual([record(), minimal]);
  });

  it("T-M2: rejects a record with an unknown field", () => {
    rejected({ ...record(), question: "Total for Maria Lopez" });
    rejected({ ...record(), sheet: "Orders" });
    rejected({ ...record(), treatments: { ...record().treatments, secret: 1 } });
    rejected({ ...record(), tokens: { prompt: 1, completion: 2, text: "PERSON_001" } });
  });

  it("T-M2: rejects free text in an enum field", () => {
    rejected(record({ mode: "Maria Lopez" as EvalRecord["mode"] }));
    rejected(record({ audit: "blocked: column A had Maria Lopez" as EvalRecord["audit"] }));
    rejected(record({ auditReasons: ["A"] as unknown as EvalRecord["auditReasons"] }));
    rejected(record({ auditReasons: ["exact", "Maria Lopez"] as EvalRecord["auditReasons"] }));
    rejected(record({ gate: "ok" as EvalRecord["gate"] }));
    rejected(record({ gateCodes: ["G-FUNC: Uses WEBSERVICE"] as unknown as EvalRecord["gateCodes"] }));
    rejected(record({ replyKind: "=SUM(A1:A9)" as EvalRecord["replyKind"] }));
    rejected(record({ rating: "great" as EvalRecord["rating"] }));
  });

  it("rejects the raw bench mode", () => {
    rejected(record({ mode: RAW_MARKER as EvalRecord["mode"] }));
  });

  it("model must have the shape of an OpenRouter model ID", () => {
    for (const ok of ["openai/gpt-6-luna", "anthropic/claude-sonnet-4.5", "meta-llama/llama-3.1-8b-instruct:free", "x~y/z"]) {
      expect(MODEL_ID_RE.test(ok)).toBe(true);
      expect(validateEvalRecord(record({ model: ok }))).not.toBeNull();
    }
    for (const bad of ["sk-or-v1-0123456789abcdef", "", "openai", "openai/gpt 6", "Maria Lopez", "a/b/c", "openai/gpt-6\n"]) {
      rejected(record({ model: bad }));
    }
  });

  it("rejects negative, fractional or non-numeric counts and wrong types", () => {
    rejected(record({ rows: -1 }));
    rejected(record({ columns: 1.5 }));
    rejected(record({ bytes: Number.NaN }));
    rejected(record({ latencyMs: Number.POSITIVE_INFINITY }));
    rejected(record({ retries: "1" as unknown as number }));
    rejected(record({ inserted: "yes" as unknown as boolean }));
    rejected(record({ tokens: { prompt: 1, completion: -2 } }));
    rejected({ ...record(), treatments: { as_is: 1, stand_in: 0, range: 0, month: 0 } });
    rejected(record({ build: "x".repeat(101) }));
  });

  it("never throws, whatever it is given", () => {
    const circular: Record<string, unknown> = { ...record() };
    circular.self = circular;
    const throwing = Object.defineProperty({ ...record() }, "rating", {
      enumerable: true,
      get() {
        throw new Error("boom");
      },
    });
    for (const v of [null, undefined, "record", 42, [], circular, throwing]) {
      expect(() => rejected(v)).not.toThrow();
    }
  });
});

describe("MetricsStore", () => {
  it("stores a copy, so later changes to the caller's object don't leak in", () => {
    const store = new MetricsStore();
    const r = record();
    store.add(r);
    (r as unknown as Record<string, unknown>).question = "Total for Maria Lopez";
    r.rows = 1;
    expect(store.records()[0]).toEqual(record());
  });

  it("update merges a patch and re-validates", () => {
    const store = new MetricsStore();
    const id = store.add(record({ rating: undefined, inserted: false }))!;
    store.update(id, { inserted: true, rating: "partly", excelError: true });
    expect(store.get(id)).toMatchObject({ inserted: true, rating: "partly", excelError: true });
  });

  it("update drops a record that no longer validates, without throwing", () => {
    const store = new MetricsStore();
    const keep = store.add(record())!;
    const drop = store.add(record({ mode: "structure_only" }))!;
    expect(() => store.update(drop, { question: "Maria Lopez" } as Partial<EvalRecord>)).not.toThrow();
    expect(store.get(drop)).toBeUndefined();
    expect(store.get(keep)).toEqual(record());
    expect(() => store.update(12345, { inserted: true })).not.toThrow();
    expect(store.records()).toHaveLength(1);
  });

  it("addRetry raises retries and adds the retry's bytes, tokens and latency", () => {
    const store = new MetricsStore();
    const id = store.add(record({ bytes: 1000, tokens: { prompt: 100, completion: 10 }, latencyMs: 500, retries: 0 }))!;
    store.addRetry(id, { bytes: 1200, tokens: { prompt: 150, completion: 20 }, latencyMs: 700 });
    expect(store.get(id)).toMatchObject({
      retries: 1,
      bytes: 2200,
      tokens: { prompt: 250, completion: 30 },
      latencyMs: 1200,
    });
    const noTokens = store.add(record({ tokens: undefined, latencyMs: undefined, bytes: 10 }))!;
    store.addRetry(noTokens, { bytes: 5 });
    expect(store.get(noTokens)).toMatchObject({ retries: 1, bytes: 15 });
    expect(store.get(noTokens)?.tokens).toBeUndefined();
  });

  it("exportJson is a pretty JSON array holding only EvalRecord fields", () => {
    const store = new MetricsStore();
    store.add(record());
    store.add(record({ mode: "structure_only", audit: "block", auditReasons: ["word", "canary"], gate: undefined }));
    store.add({ ...record(), headers: ["Customer"] } as EvalRecord);
    const json = store.exportJson();
    const parsed = JSON.parse(json) as Record<string, unknown>[];
    expect(Array.isArray(parsed)).toBe(true);
    expect(parsed).toHaveLength(2);
    for (const r of parsed) {
      for (const k of Object.keys(r)) expect(FIELDS).toContain(k);
      expect(Object.keys(r.treatments as object).sort()).toEqual(["as_is", "exclude", "month", "range", "stand_in"]);
      if (r.tokens) expect(Object.keys(r.tokens as object).sort()).toEqual(["completion", "prompt"]);
      expect(EvalRecordSchema.safeParse(r).success).toBe(true);
    }
    expect(json).not.toContain("Customer");
    expect(json).toBe(JSON.stringify(store.records(), null, 2));
    expect(new MetricsStore().exportJson()).toBe("[]");
  });

  it("the schema's fields are exactly the EvalRecord fields", () => {
    expect(Object.keys(EvalRecordSchema.shape).sort()).toEqual([...FIELDS].sort());
  });
});

describe("metrics.ts source", () => {
  it("does not contain the raw bench marker (T-P4)", () => {
    const src = readFileSync(resolve(ROOT, "src/core/metrics.ts"), "utf8");
    expect(src.includes(RAW_MARKER)).toBe(false);
  });
});
