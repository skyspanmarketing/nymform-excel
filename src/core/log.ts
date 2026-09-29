// In-memory request log (spec §7.12). Holds what was sent (stand-ins only), the audit result, the raw
// reply and the gate result. Never the key or the stand-in map. Gone when the task pane closes.
// Beside an entry it may keep pane-only notes (LogLocal) that never go into the entry or an export.
import type { AuditOrigin, AuditVariant, LogEntry } from "./types";

export const LOG_CAPACITY = 200;

/**
 * What a blocked request matched, for the Log screen only: the matched text and the text around it
 * come from the request that was blocked, so they hold a private value. Kept beside the entry, never
 * in it: entries(), get() and exportJson() leave it out, and so do the evaluation report and history.
 */
export interface LocalMatch {
  column: string;
  variant: AuditVariant;
  origin: AuditOrigin;
  /** The matched text, unescaped for reading. */
  text: string;
  /** Up to 32 characters before and after it. */
  before: string;
  after: string;
  /** Where in the request it was found, in words ("your question", "the header of column J"). */
  where: string;
  /** The column whose header (as sent) holds the match, when it is in a header. */
  inHeader?: string;
}

export interface LogLocal {
  matches?: readonly LocalMatch[];
}

/** The fields a log entry may carry. Anything else a caller passes in is dropped. */
const LOG_KEYS = [
  "at",
  "mode",
  "model",
  "host",
  "bytes",
  "body",
  "audit",
  "replyRaw",
  "gate",
  "inserted",
  "latencyMs",
  "tokens",
  "providerName",
  "error",
] as const satisfies readonly (keyof LogEntry)[];

function pick(source: Partial<LogEntry>): Partial<LogEntry> {
  const out: Record<string, unknown> = {};
  for (const k of LOG_KEYS) {
    if (Object.prototype.hasOwnProperty.call(source, k)) out[k] = source[k];
  }
  return out as Partial<LogEntry>;
}

/** A frozen copy of the pane-only notes, so a caller can't change them afterwards. */
function copyLocal(local: LogLocal | undefined): LogLocal | undefined {
  if (!local) return undefined;
  return Object.freeze(local.matches ? { matches: Object.freeze(local.matches.map((m) => Object.freeze({ ...m }))) } : {});
}

export class RequestLog {
  readonly #items: { id: number; entry: LogEntry; local: LogLocal | undefined }[] = [];
  #nextId = 1;

  /**
   * Adds an entry, dropping the oldest past LOG_CAPACITY. Returns the entry's id. `local` is kept
   * beside the entry for the pane (see items()) and dropped with it.
   */
  add(entry: LogEntry, local?: LogLocal): number {
    const id = this.#nextId++;
    this.#items.push({ id, entry: pick(entry) as LogEntry, local: copyLocal(local) });
    while (this.#items.length > LOG_CAPACITY) this.#items.shift();
    return id;
  }

  /** Merges a patch into an entry. Does nothing when the entry has already been dropped. */
  update(id: number, patch: Partial<LogEntry>): void {
    const item = this.#items.find((i) => i.id === id);
    if (item) item.entry = { ...item.entry, ...pick(patch) };
  }

  /** Oldest first. */
  entries(): readonly LogEntry[] {
    return this.#items.map((i) => i.entry);
  }

  /** Oldest first, each entry with its pane-only notes. For the Log screen; never exported. */
  items(): readonly { entry: LogEntry; local: LogLocal | undefined }[] {
    return this.#items.map((i) => ({ entry: i.entry, local: i.local }));
  }

  get(id: number): LogEntry | undefined {
    return this.#items.find((i) => i.id === id)?.entry;
  }

  get size(): number {
    return this.#items.length;
  }

  clear(): void {
    this.#items.length = 0;
  }

  /** JSON for the Export log button: sent bodies, audit, raw reply, gate. Never the key or the map. */
  exportJson(): string {
    return JSON.stringify(this.entries(), null, 2);
  }
}
