// Local evaluation metadata (spec §7.14). Describes how Nymform performed, never what the data was.
// Kept in memory next to the log; nothing is transmitted. Export happens only when the user clicks.
import { z } from "zod";
import type { EvalRecord } from "./types";

/** The shape of an OpenRouter model ID, so a key pasted into the model field can't reach a report. */
export const MODEL_ID_RE = /^[\w.~-]+\/[\w.:~-]+$/;

const count = z.number().int().nonnegative();
const tokens = z.strictObject({ prompt: count, completion: count });

/** Every string field except `build` and `model` is an enum. Unknown fields are rejected. */
export const EvalRecordSchema = z.strictObject({
  build: z.string().max(100),
  model: z.string().max(200).regex(MODEL_ID_RE),
  mode: z.enum(["structure_only", "substituted"]),
  rows: count,
  columns: count,
  privateColumns: count,
  treatments: z.strictObject({
    as_is: count,
    stand_in: count,
    range: count,
    month: count,
    exclude: count,
  }),
  bytes: count,
  tokens: tokens.optional(),
  latencyMs: z.number().nonnegative().optional(),
  audit: z.enum(["pass", "block"]),
  auditReasons: z.array(z.enum(["exact", "normalized", "digits", "json", "word", "canary", "structural", "error"])),
  gate: z.enum(["pass", "block"]).optional(),
  gateCodes: z.array(z.enum(["G-START", "G-LEN", "G-FUNC", "G-NAME", "G-EXT", "G-REF", "G-URL", "G-PARSE"])),
  replyKind: z.enum(["formula", "answer", "clarify", "invalid"]).optional(),
  retries: count,
  inserted: z.boolean(),
  copied: z.boolean(),
  excelError: z.boolean().optional(),
  rating: z.enum(["worked", "partly", "wrong"]).optional(),
});

// The schema and the EvalRecord type must describe the same closed list of fields.
type SchemaRecord = z.infer<typeof EvalRecordSchema>;
const _sameFields: [SchemaRecord extends EvalRecord ? true : false, EvalRecord extends SchemaRecord ? true : false] = [
  true,
  true,
];
void _sameFields;

/** The validated record, or null. Never throws. */
export function validateEvalRecord(record: unknown): EvalRecord | null {
  try {
    const parsed = EvalRecordSchema.safeParse(record);
    return parsed.success ? parsed.data : null;
  } catch {
    return null;
  }
}

export class MetricsStore {
  readonly #records = new Map<number, EvalRecord>();
  #nextId = 1;

  /** Validates and stores a record; returns its id, or null when the record is dropped. */
  add(record: EvalRecord): number | null {
    const valid = validateEvalRecord(record);
    if (!valid) return null;
    const id = this.#nextId++;
    this.#records.set(id, valid);
    return id;
  }

  /** Merges a patch, re-validates, and drops the record if it no longer validates. */
  update(id: number, patch: Partial<EvalRecord>): void {
    const current = this.#records.get(id);
    if (!current) return;
    let merged: unknown;
    try {
      merged = { ...current, ...patch };
    } catch {
      merged = null;
    }
    const valid = validateEvalRecord(merged);
    if (valid) this.#records.set(id, valid);
    else this.#records.delete(id);
  }

  /**
   * The §7.9 correction retry: the same record, with `retries` raised by one and the retry's
   * bytes, tokens and latency added.
   */
  addRetry(id: number, retry: { bytes: number; tokens?: { prompt: number; completion: number }; latencyMs?: number }): void {
    const current = this.#records.get(id);
    if (!current) return;
    const patch: Partial<EvalRecord> = { retries: current.retries + 1, bytes: current.bytes + retry.bytes };
    if (retry.tokens) {
      patch.tokens = {
        prompt: (current.tokens?.prompt ?? 0) + retry.tokens.prompt,
        completion: (current.tokens?.completion ?? 0) + retry.tokens.completion,
      };
    }
    if (retry.latencyMs !== undefined) patch.latencyMs = (current.latencyMs ?? 0) + retry.latencyMs;
    this.update(id, patch);
  }

  get(id: number): EvalRecord | undefined {
    return this.#records.get(id);
  }

  /** In the order they were added. */
  records(): readonly EvalRecord[] {
    return [...this.#records.values()];
  }

  clear(): void {
    this.#records.clear();
  }

  /** JSON for the Export evaluation report button: validated records only, nothing else. */
  exportJson(): string {
    const out: EvalRecord[] = [];
    for (const r of this.#records.values()) {
      const valid = validateEvalRecord(r);
      if (valid) out.push(valid);
    }
    return JSON.stringify(out, null, 2);
  }
}
