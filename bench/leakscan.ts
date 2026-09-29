// The leak scan behind bench/leakcheck.ts, with no Node imports so the bench panel can run it
// in the task pane (bench builds only). Deliberately independent of src/core/auditor.ts.

/** Values shorter than this are not scanned as substrings (they would match constantly). */
export const MIN_VALUE_LENGTH = 4;
/** Values with at least this many digits are also scanned as their digits alone. */
export const MIN_DIGITS = 7;
/** Modes whose requests must never contain a private value. */
export const PROTECTED_MODES: readonly string[] = ["substituted", "structure_only"];

export type LeakVariant = "exact" | "json" | "json2" | "digits" | "canary";

export interface Leak {
  /** Index of the request in the entries passed in. */
  entry: number;
  mode: string;
  kind: "value" | "canary";
  /** Index into the privateValues (or canaries) list. */
  index: number;
  variant: LeakVariant;
}

export interface ModeCounts {
  requests: number;
  /** Requests with at least one leak. */
  leakingRequests: number;
  /** Distinct (request, value) leaks. */
  leaks: number;
}

export interface LeakReport {
  byMode: Record<string, ModeCounts>;
  leaks: Leak[];
  /** Leaks in substituted and structure_only requests. */
  protectedLeaks: number;
}

interface Pattern {
  kind: "value" | "canary";
  index: number;
  /** The value is a plain number, so a match that is only the bound of a range rendering doesn't count. */
  numeric: boolean;
  variants: { variant: LeakVariant; text: string }[];
}

const NUMERIC_RE = /^-?\d{1,3}(?:,?\d{3})*(?:\.\d+)?$/;

function isBin(lo: string, hi: string): boolean {
  if (lo === "0") return hi === "9";
  const m = /^([1-9])(0+)$/.exec(lo);
  return m !== null && hi === m[1]! + "9".repeat(m[2]!.length);
}

/**
 * A `range` treatment rendering such as "80000-89999", "0-9" or "-89999--80000". A round private
 * value (80000) is the lower bound of its own bin, so its text appears inside the rendering even
 * though the rendering says no more about it than about 84500.
 */
export function isBinRendering(s: string): boolean {
  let m = /^(\d+)-(\d+)$/.exec(s);
  if (m) return isBin(m[1]!, m[2]!);
  m = /^-(\d+)--?(\d+)$/.exec(s);
  if (m) return isBin(m[2]!, m[1]!);
  return false;
}

/** True when the match at [start, end) is one bound of a range rendering and nothing more. */
function insideBinRendering(body: string, start: number, end: number): boolean {
  let a = start;
  let b = end;
  while (a > 0 && /[0-9-]/.test(body[a - 1]!)) a--;
  while (b < body.length && /[0-9-]/.test(body[b]!)) b++;
  return isBinRendering(body.slice(a, b));
}

function occursIn(body: string, text: string, numeric: boolean): boolean {
  if (!numeric) return body.includes(text);
  for (let i = body.indexOf(text); i !== -1; i = body.indexOf(text, i + 1)) {
    if (!insideBinRendering(body, i, i + text.length)) return true;
  }
  return false;
}

function jsonEscaped(v: string): string {
  return JSON.stringify(v).slice(1, -1);
}

function patternsFor(privateValues: readonly string[], canaries: readonly string[]): Pattern[] {
  const out: Pattern[] = [];
  privateValues.forEach((value, index) => {
    const variants: Pattern["variants"] = [];
    const add = (variant: LeakVariant, text: string) => {
      const t = text.toLowerCase();
      if (t.length >= MIN_VALUE_LENGTH && !variants.some((v) => v.text === t)) variants.push({ variant, text: t });
    };
    if (value.length >= MIN_VALUE_LENGTH) {
      add("exact", value);
      add("json", jsonEscaped(value));
      add("json2", jsonEscaped(jsonEscaped(value)));
    }
    const digits = value.replace(/\D/g, "");
    if (digits.length >= MIN_DIGITS) add("digits", digits);
    if (variants.length > 0) out.push({ kind: "value", index, numeric: NUMERIC_RE.test(value), variants });
  });
  canaries.forEach((canary, index) => {
    if (canary.length > 0) {
      out.push({ kind: "canary", index, numeric: false, variants: [{ variant: "canary", text: canary.toLowerCase() }] });
    }
  });
  return out;
}

/**
 * Scans each request body, case-insensitively, for every private value and canary. The one
 * exception: a plain-number value found only as the bound of a range rendering ("80000" in
 * "80000-89999") is not a leak.
 */
export function leakCheck(
  entries: readonly { mode: string; body: string }[],
  privateValues: readonly string[],
  canaries: readonly string[],
): LeakReport {
  const patterns = patternsFor(privateValues, canaries);
  const byMode: Record<string, ModeCounts> = {};
  const leaks: Leak[] = [];
  entries.forEach((e, entry) => {
    const counts = (byMode[e.mode] ??= { requests: 0, leakingRequests: 0, leaks: 0 });
    counts.requests++;
    const body = e.body.toLowerCase();
    let found = 0;
    for (const p of patterns) {
      const hit = p.variants.find((v) => occursIn(body, v.text, p.numeric));
      if (!hit) continue;
      found++;
      leaks.push({ entry, mode: e.mode, kind: p.kind, index: p.index, variant: hit.variant });
    }
    counts.leaks += found;
    if (found > 0) counts.leakingRequests++;
  });
  const protectedLeaks = leaks.filter((l) => PROTECTED_MODES.includes(l.mode)).length;
  return { byMode, leaks, protectedLeaks };
}

// ---------------------------------------------------------------------------------------------
// Log files

export interface PrivateValuesFile {
  [workbook: string]: { sheet: string; range: string; privateLetters: string[]; values: string[]; canaries: string[] };
}

/**
 * Pulls { mode, body } requests out of an exported log. Accepts a list of entries or an object
 * holding one (under entries, log, requests or records). Entries without a string mode and body are
 * counted as skipped.
 */
export function extractEntries(json: unknown): { entries: { mode: string; body: string }[]; skipped: number } {
  let list: unknown = json;
  if (list !== null && typeof list === "object" && !Array.isArray(list)) {
    const o = list as Record<string, unknown>;
    list = o.entries ?? o.log ?? o.requests ?? o.records;
  }
  if (!Array.isArray(list)) throw new Error("No list of log entries found in the file.");
  const entries: { mode: string; body: string }[] = [];
  let skipped = 0;
  for (const item of list) {
    const o = item as Record<string, unknown> | null;
    if (o && typeof o === "object" && typeof o.mode === "string" && typeof o.body === "string") {
      entries.push({ mode: o.mode, body: o.body });
    } else {
      skipped++;
    }
  }
  return { entries, skipped };
}
