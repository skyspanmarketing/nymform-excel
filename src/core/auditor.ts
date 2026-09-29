// Auditor (spec §7.7). Both checks run on the exact `outbound.body` string that provider.ts sends.
// The brand symbol is not exported, so only this file can create an AuditedOutbound (invariant 3).
import { columnIndex, parseLocalRange, quoteSheet } from "./a1";
import { CORRECTION_PROMPT, SYSTEM_PROMPT } from "./prompt";
import { columnOffset, headerCellIsBareValue, headerCellIsIdentifier, rowCount } from "./schema";
import {
  MONTH_RE,
  RANGE_RE,
  TOKEN_RE,
  aliasRevealsHeader,
  codeLetters,
  codeWithoutSeparators,
  dateKey,
  droppedLength,
  emailLocalPart,
  findDates,
  foldUnit,
  isAccentMark,
  isCodeSeparatorAt,
  isDroppedCodePoint,
  isHiddenNumberText,
  isIgnorableCodePoint,
  joinedNames,
  letterSeparatedParts,
  looseDatesOf,
  privateDates,
  privateLooseDates,
  privateNumbers,
  privateValues,
  shortNumbers,
  shownNumbers,
  textDates,
  withoutCountryCode,
  withoutMarks,
  withoutSharpS,
  withoutUmlauts,
} from "./transform";
import type { DateMatch } from "./transform";
import type {
  AuditField,
  AuditLocation,
  AuditOrigin,
  AuditResult,
  AuditVariant,
  ColType,
  Outbound,
  ProdMode,
  RangeSpec,
  Treatment,
} from "./types";

const audited: unique symbol = Symbol("nymform.audited");
/** Every AuditedOutbound this file has issued. A spread copy carries the brand symbol but isn't in here. */
const issued = new WeakSet<object>();
// Taken at load, so patching WeakSet.prototype.has or Reflect.apply later can't make isAudited pass.
const weakSetHas = WeakSet.prototype.has;
const apply = Reflect.apply;

export type AuditedOutbound = Outbound & {
  readonly [audited]: true;
  readonly audit: AuditResult & { ok: true };
};

export interface AuditColumn {
  letter: string;
  private: boolean;
  treatment: Treatment;
  /** Every non-empty display text and raw value in this column across the whole range (private columns only). */
  values: string[];
  /**
   * The numbers a user reads in this column's cells, as plain unsigned decimals ("83500", "5.25"),
   * matched plainly and with grouping. Left out, they are read from `values`.
   */
  numbers?: string[];
  /**
   * This column's dates as ISO dates ("1985-03-15"): every date a cell can mean (see privateDates
   * in transform.ts). Every date phrase in the body ("3/15/85", "4th of July 1988"; see findDates)
   * is compared with them. Other date text here is read with textDates. Left out, they are read
   * from date text in `values`.
   */
  dates?: string[];
  /**
   * The dates this column's 8-digit cells can mean, as ISO dates (see privateLooseDates in
   * transform.ts): "19850315" and "03151985" may be dates of birth, or IDs and lot numbers, so a
   * column gives them only when at least half of its 8-digit values are dates. They count only in
   * free text (see BodyMask.free). Left out, they are read from 8-digit `values` the same way (see
   * looseDatesOf). Either way the digits of an 8-digit value count wherever they are typed, as a
   * date too ("2024-01-15" for invoice number 20240115).
   */
  looseDates?: string[];
  /**
   * Groups of 4 to 6 digits a cell doesn't show, matched like the last group of a phone number
   * (only in free text): the last 4 digits of a 7+ digit number whose display is "#####" because
   * the column is too narrow. Left out, there are none.
   */
  groups?: string[];
  /** Original header text when the column is sent under an alias. */
  aliasOriginal?: string;
}

export interface AuditRange {
  sheet: string;
  /** Local address, e.g. "A1:F501". */
  address: string;
  /** True when the first row of the range is its header row. */
  hasHeaders: boolean;
  /** Number of data rows: the rows below the header row, or every row without one. */
  dataRows: number;
  columns: AuditColumn[];
}

export interface AuditContext {
  /** [0] is the selection; the rest are context ranges. */
  ranges: AuditRange[];
  /** True when a string is a token from the session's stand-in map. */
  isToken: (s: string) => boolean;
  /** Planted canary strings (bench workbooks); matched exactly, length 6+. */
  canaries: string[];
  /** The configured model ID, the only value allowed in the top-level "model" field. */
  model: string;
}

/**
 * A block also says where each hit's first match is in the body (`locations`, in the order of
 * `result.hits`; empty when the request was refused before the text scan). A hit whose text was
 * found both in a column's values and in its original header has two, the first match of each,
 * the one recorded first leading; any other hit has one. Like the result it holds no text, only
 * offsets and paths.
 */
export type AuditOutcome =
  | { ok: true; outbound: AuditedOutbound }
  | { ok: false; result: AuditResult; locations: readonly AuditLocation[] };

/**
 * True only for the exact object audit() returned. provider.ts calls this before sending, so a copy
 * such as `{ ...audited, body: other }` (which still type-checks) is refused at runtime.
 */
export function isAudited(o: unknown): o is AuditedOutbound {
  return typeof o === "object" && o !== null && apply(weakSetHas, issued, [o]) === true;
}

/** Plain indexOf scanning up to this many distinct variants; Aho-Corasick above it, for long texts. */
export const INDEXOF_LIMIT = 5000;
/**
 * Aho-Corasick only pays for building its trie on a long text: below 16 KB, indexOf is faster
 * even with 20,000 variants (measured on 2,000- and 20,000-row sheets).
 */
export const AHO_MIN_TEXT = 16 * 1024;

export type ScanStrategy = "auto" | "indexOf" | "ahoCorasick";

export interface AuditOptions {
  /** Which matcher the text scan uses. "auto" (default): see chooseMatcher. */
  strategy?: ScanStrategy;
}

const MIN_VARIANT = 4;
const MIN_CANARY = 6;
const MIN_DIGITS = 7;

const TOP_KEYS = new Set(["model", "messages", "temperature", "max_tokens", "response_format", "reasoning", "provider"]);
const PARAM_KEYS = new Set(["temperature", "max_tokens", "response_format", "reasoning", "provider"]);
const ROLES = new Set(["system", "user", "assistant"]);
const PROD_MODES = new Set<string>(["structure_only", "substituted"]);
const COL_TYPES = new Set<string>(["text", "number", "date", "boolean", "mixed", "empty"] satisfies ColType[]);
const TREATMENTS = new Set<string>(["as_is", "stand_in", "range", "month", "exclude"] satisfies Treatment[]);

const DESCRIPTION_KEYS = new Set([
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
const CONTEXT_RANGE_KEYS = new Set(["sheet", "table", "range", "columns", "rows", "rows_sent", "rows_total"]);
const RANGE_KEYS = new Set(["address", "header_row", "first_data_row", "last_data_row"]);
const COLUMN_KEYS = new Set(["letter", "header", "type", "private", "treatment", "note", "stats"]);
const STATS_KEYS = new Set(["blank", "distinct", "avg_words", "max_len"]);
const ROW_KEYS = ["rows", "rows_sent", "rows_total"];

const REASONING_KEYS = new Set(["effort", "exclude", "max_tokens", "enabled"]);
const EFFORTS = new Set(["minimal", "low", "medium", "high"]);
const PROVIDER_KEYS = new Set(["zdr", "data_collection"]);

/** Key names the body and a nymform user turn use. Other keys are scanned like any other text. */
const KNOWN_KEYS = new Set([
  ...TOP_KEYS,
  "role",
  "content",
  "type",
  ...REASONING_KEYS,
  ...PROVIDER_KEYS,
  ...DESCRIPTION_KEYS,
  ...RANGE_KEYS,
  ...COLUMN_KEYS,
  ...STATS_KEYS,
]);

/** String values that are masked inside a nymform user turn when they are allowed members. */
const ENUM_VALUES = new Map<string, (v: string) => boolean>([
  ["nymform", (v) => v === "0.1"],
  ["mode", (v) => PROD_MODES.has(v)],
  ["type", (v) => COL_TYPES.has(v)],
  ["treatment", (v) => TREATMENTS.has(v)],
]);

const VARIANT_ORDER: AuditVariant[] = ["exact", "normalized", "digits", "json", "word", "canary"];

/** Keys of a model reply (spec §8), masked in assistant turns that parse as JSON. */
const REPLY_KEYS = new Set(["kind", "formula", "placement", "cell", "fill_down", "explanation", "assumptions"]);
const REPLY_KINDS = new Set(["formula", "answer", "clarify"]);

type Hit = AuditResult["hits"][number];

// ---------------------------------------------------------------------------------------------
// Public API

/** Checks the exact body. Only a pass on both checks returns an AuditedOutbound. Never throws. */
export function audit(outbound: Outbound, ctx: AuditContext, options: AuditOptions = {}): AuditOutcome {
  try {
    return runAudit(outbound, ctx, options);
  } catch {
    // The message stays fixed: an exception's text could quote the body or a cell value.
    return blocked("The check stopped on an unexpected problem, so nothing was sent.");
  }
}

export function buildAuditContext(
  specs: readonly RangeSpec[],
  map: { isToken(s: string): boolean },
  opts: { model: string; canaries?: string[] },
): AuditContext {
  const ranges = specs.map((spec): AuditRange => {
    const found = privateValues(spec);
    const numbers = privateNumbers(spec);
    const dates = privateDates(spec);
    const looseDates = privateLooseDates(spec);
    const columns = spec.policies.map((policy, i): AuditColumn => {
      const info = spec.columns.find((c) => c.letter === policy.letter) ?? spec.columns[i];
      const offset = columnOffset(spec.data, policy.letter, i);
      const values = policy.private
        ? unique([...(found.get(policy.letter) ?? []), ...columnValues(spec, policy.letter, i), ...headerData(spec, offset)])
        : [];
      const col: AuditColumn = {
        letter: policy.letter,
        private: policy.private,
        treatment: effectiveTreatment(policy.private, policy.treatment),
        values,
      };
      if (policy.private) {
        col.numbers = numbers.get(policy.letter) ?? [];
        col.dates = dates.get(policy.letter) ?? [];
        col.looseDates = looseDates.get(policy.letter) ?? [];
        col.groups = hiddenGroups(spec, policy.letter, i);
      }
      // A header that is a bare number or date, such as a year, is not scanned for under its alias,
      // as headerData leaves it out: "2024" would block every request on a sheet named "Budget 2024".
      // Nor is one the alias still shows ("Order Date" or "Order date (month)" for "Order date"):
      // it goes out in the alias, where the scan would find it in every request.
      const bare = spec.hasHeaders && headerCellIsBareValue(spec.data, offset);
      if (policy.alias && info && !aliasRevealsHeader(policy.alias, info.header) && !bare) col.aliasOriginal = info.header;
      return col;
    });
    const dataRows = Math.max(0, rowCount(spec.data) - (spec.hasHeaders ? 1 : 0));
    return { sheet: spec.data.sheet, address: spec.data.address, hasHeaders: spec.hasHeaders, dataRows, columns };
  });
  return {
    ranges,
    isToken: (s: string) => map.isToken(s),
    canaries: [...(opts.canaries ?? [])],
    model: opts.model,
  };
}

/** The matcher "auto" picks: Aho-Corasick only for more than INDEXOF_LIMIT patterns in a text longer than AHO_MIN_TEXT. */
export function chooseMatcher(strategy: ScanStrategy, patterns: number, textLength: number): "indexOf" | "ahoCorasick" {
  if (strategy !== "auto") return strategy;
  return patterns > INDEXOF_LIMIT && textLength > AHO_MIN_TEXT ? "ahoCorasick" : "indexOf";
}

/**
 * Calls `onMatch(patternIndex, start)` for every occurrence of every pattern in `haystack`,
 * overlapping ones included. Both matchers report the same set of occurrences; only the order
 * differs. With indexOf, returning true from `onMatch` skips the rest of that pattern.
 */
export function findAll(
  haystack: string,
  patterns: readonly string[],
  strategy: ScanStrategy,
  onMatch: (pattern: number, start: number) => boolean | void,
): void {
  if (chooseMatcher(strategy, patterns.length, haystack.length) === "ahoCorasick") ahoCorasick(haystack, patterns, onMatch);
  else indexOfScan(haystack, patterns, onMatch);
}

// ---------------------------------------------------------------------------------------------
// The audit

interface CtxColumn {
  letter: string;
  label: string;
  order: number;
  private: boolean;
  treatment: string;
  values: string[];
  /** Null: read from `values`. */
  numbers: string[] | null;
  /** Null: read from `values`. */
  dates: string[] | null;
  /** Null: read from `values`. */
  looseDates: string[] | null;
  groups: string[];
  aliasOriginal?: string;
}

interface CtxRange {
  sheet: string;
  address: string;
  /** Null when the context doesn't say; a description of the range then blocks. */
  hasHeaders: boolean | null;
  dataRows: number | null;
  columns: CtxColumn[];
}

interface Snapshot {
  ranges: CtxRange[];
  isToken: (s: string) => boolean;
  canaries: string[];
  model: string;
}

class Blocked extends Error {}

function blocked(
  error: string,
  hits: Hit[] = [],
  structural: { column: string }[] = [],
  locations: readonly AuditLocation[] = [],
): { ok: false; result: AuditResult; locations: readonly AuditLocation[] } {
  return { ok: false, result: { ok: false, hits, structural, error }, locations };
}

function runAudit(outbound: Outbound, ctx: AuditContext, options: AuditOptions): AuditOutcome {
  if (!outbound || typeof outbound !== "object") return blocked("There is no request to check.");
  // Read each field once, so the string checked is the string returned.
  const body: unknown = outbound.body;
  const mode: unknown = outbound.mode;
  const bytes: unknown = outbound.bytes;
  const createdAt: unknown = outbound.createdAt;
  if (typeof body !== "string" || typeof bytes !== "number" || typeof createdAt !== "string") {
    return blocked("The request is incomplete.");
  }
  if (mode !== "structure_only" && mode !== "substituted") {
    return blocked("Only structure-only and substituted requests can be sent.");
  }
  const snap = snapshot(ctx);

  let parsed: unknown;
  try {
    parsed = JSON.parse(body);
  } catch {
    return blocked("The request isn't valid JSON.");
  }
  if (JSON.stringify(parsed) !== body) return blocked("The request isn't in the exact form Nymform builds.");

  try {
    const descriptions = checkEnvelope(parsed, snap, mode);
    // The first match of each hit is kept with it, as offsets into the body, for the pane. So is the
    // first match of the other origin, when the hit's text was found both in the column's values and
    // in its original header: otherwise a header seen first would hide the value, and the pane would
    // blame only the header. The hit is the same either way.
    const hits = new Map<string, { hit: Hit; order: number; kind: number; first: Match; also?: Match }>();
    const addHit: AddHit = (col, variant, start, end) => {
      const key = `${col.order}|${variant}`;
      const known = hits.get(key);
      if (known === undefined) {
        hits.set(key, { hit: { column: col.label, variant }, order: col.order, kind: VARIANT_ORDER.indexOf(variant), first: { origin: col.origin, start, end } });
      } else if (known.also === undefined && known.first.origin !== col.origin) {
        known.also = { origin: col.origin, start, end };
      }
    };
    const structural = new Map<number, string>();
    const addStructural = (col: { label: string; order: number }) => {
      if (!structural.has(col.order)) structural.set(col.order, col.label);
    };
    for (const d of descriptions) checkDescription(d.value, snap, mode, addStructural);
    textScan(body, parsed as Record<string, unknown>, descriptions, snap, options.strategy ?? "auto", addHit);

    if (hits.size > 0 || structural.size > 0) {
      const found = [...hits.values()].sort((a, b) => a.order - b.order || a.kind - b.kind);
      const sorted = found.map((h) => h.hit);
      const cells = [...structural.entries()].sort((a, b) => a[0] - b[0]).map(([, column]) => ({ column }));
      const fieldOf = fieldFinder(body);
      const locations = found.flatMap((h) =>
        (h.also ? [h.first, h.also] : [h.first]).map(
          (m): AuditLocation => ({ ...h.hit, start: m.start, end: m.end, origin: m.origin, field: fieldOf(m.start) }),
        ),
      );
      return blocked(describeHits(sorted, cells, locations), sorted, cells, locations);
    }
  } catch (e) {
    if (e instanceof Blocked) return blocked(e.message);
    throw e;
  }

  const hits: Hit[] = [];
  Object.freeze(hits);
  const structural: { column: string }[] = [];
  Object.freeze(structural);
  const result: AuditResult & { ok: true } = { ok: true, hits, structural };
  Object.freeze(result);
  const out: AuditedOutbound = { body, bytes, mode: mode satisfies ProdMode, createdAt, [audited]: true, audit: result };
  Object.freeze(out);
  issued.add(out);
  return { ok: true, outbound: out };
}

/**
 * A plain sentence naming columns by letter only, never values. A column whose matches all came
 * from its original header (renamed under Header sent) isn't called private: it may not be.
 */
function describeHits(hits: Hit[], structural: { column: string }[], locations: readonly AuditLocation[]): string {
  const cols = (keep: (h: Hit) => boolean) => [...new Set(hits.filter(keep).map((h) => h.column))];
  const unreplaced = [...new Set(structural.map((s) => s.column))];
  const fromValues = new Set(locations.filter((l) => l.origin === "value" && l.variant !== "canary").map((l) => l.column));
  const found = cols((h) => h.variant !== "canary" && fromValues.has(h.column));
  const headers = cols((h) => h.variant !== "canary" && !fromValues.has(h.column));
  const parts: string[] = [];
  if (unreplaced.length > 0) {
    parts.push(`A private cell in ${plural(unreplaced)} wasn't replaced with a stand-in.`);
  }
  if (found.length > 0) parts.push(`Text from private ${plural(found)} is in the request.`);
  if (headers.length === 1) parts.push(`Text from column ${headers[0]}'s original header is in the request.`);
  else if (headers.length > 1) parts.push(`Text from the original headers of ${plural(headers)} is in the request.`);
  if (hits.some((h) => h.variant === "canary")) parts.push("A bench canary is in the request.");
  parts.push("Nothing was sent.");
  return parts.join(" ");
}

function plural(labels: string[]): string {
  return `${labels.length === 1 ? "column" : "columns"} ${labels.join(", ")}`;
}

function effectiveTreatment(isPrivate: boolean, treatment: Treatment): Treatment {
  // A private column is never sent as is (brief, decision 3).
  return isPrivate && treatment === "as_is" ? "stand_in" : treatment;
}

/** Copies what the audit needs out of the context, so it can't change halfway through. */
function snapshot(ctx: AuditContext): Snapshot {
  if (!ctx || typeof ctx !== "object") throw new Error("no context");
  const rangesIn: unknown = ctx.ranges;
  if (!Array.isArray(rangesIn) || rangesIn.length === 0) throw new Error("no ranges");
  let order = 0;
  const ranges = rangesIn.map((r: AuditRange, ri): CtxRange => {
    if (!r || typeof r.sheet !== "string" || typeof r.address !== "string" || !Array.isArray(r.columns)) {
      throw new Error("bad range");
    }
    const columns = r.columns.map((col: AuditColumn): CtxColumn => {
      if (!col || typeof col.letter !== "string") throw new Error("bad column");
      // Anything but an explicit false counts as private.
      const isPrivate = col.private !== false;
      const values = Array.isArray(col.values) ? col.values.flatMap(asText) : [];
      const numbers: unknown = col.numbers;
      const dates: unknown = col.dates;
      const looseDates: unknown = col.looseDates;
      const groups: unknown = col.groups;
      const c: CtxColumn = {
        letter: col.letter,
        label: ri === 0 ? col.letter : `${quoteSheet(r.sheet)}!${col.letter}`,
        order: order++,
        private: isPrivate,
        treatment: effectiveTreatment(isPrivate, col.treatment),
        values,
        numbers: Array.isArray(numbers) ? numbers.filter((n): n is string => typeof n === "string") : null,
        dates: Array.isArray(dates) ? dates.filter((n): n is string => typeof n === "string") : null,
        looseDates: Array.isArray(looseDates) ? looseDates.filter((n): n is string => typeof n === "string") : null,
        groups: Array.isArray(groups) ? groups.filter((n): n is string => typeof n === "string") : [],
      };
      if (typeof col.aliasOriginal === "string") c.aliasOriginal = col.aliasOriginal;
      return c;
    });
    const hasHeaders: unknown = r.hasHeaders;
    const dataRows: unknown = r.dataRows;
    return {
      sheet: r.sheet,
      address: r.address,
      hasHeaders: typeof hasHeaders === "boolean" ? hasHeaders : null,
      dataRows: typeof dataRows === "number" && Number.isInteger(dataRows) && dataRows >= 0 ? dataRows : null,
      columns,
    };
  });
  const isTokenFn: unknown = ctx.isToken;
  if (typeof isTokenFn !== "function") throw new Error("no isToken");
  const canariesIn: unknown = ctx.canaries;
  const canaries = Array.isArray(canariesIn) ? canariesIn.filter((c): c is string => typeof c === "string") : [];
  if (typeof ctx.model !== "string") throw new Error("no model");
  return {
    ranges,
    isToken: (s: string) => (isTokenFn as (s: string) => unknown).call(ctx, s) === true,
    canaries,
    model: ctx.model,
  };
}

function asText(v: unknown): string[] {
  if (typeof v === "string") return [v];
  if (typeof v === "number" || typeof v === "boolean") return [String(v)];
  return [];
}

// ---------------------------------------------------------------------------------------------
// Preconditions

interface Description {
  /** Index in `messages`. */
  message: number;
  value: Record<string, unknown>;
}

function isObject(v: unknown): v is Record<string, unknown> {
  return v !== null && typeof v === "object" && !Array.isArray(v);
}

function has(o: Record<string, unknown>, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(o, key);
}

function onlyKeys(o: Record<string, unknown>, allowed: Set<string>): boolean {
  return Object.keys(o).every((k) => allowed.has(k));
}

function isCount(v: unknown): boolean {
  return typeof v === "number" && Number.isInteger(v) && v > 0;
}

/** Checks the request's outer shape and returns the nymform descriptions in message order. */
function checkEnvelope(parsed: unknown, snap: Snapshot, mode: ProdMode): Description[] {
  if (!isObject(parsed)) throw new Blocked("The request isn't in the form Nymform builds.");
  if (!onlyKeys(parsed, TOP_KEYS)) throw new Blocked("The request has a field Nymform doesn't send.");
  if (parsed.model !== snap.model) throw new Blocked("The request names a model other than the one in Setup.");
  checkParams(parsed);

  const messages = parsed.messages;
  if (!Array.isArray(messages) || messages.length < 2) throw new Blocked("The request has no question in it.");
  const descriptions: Description[] = [];
  messages.forEach((m: unknown, i) => {
    if (!isObject(m) || Object.keys(m).length !== 2 || !has(m, "role") || !has(m, "content")) {
      throw new Blocked("A message in the request isn't in the form Nymform builds.");
    }
    const { role, content } = m;
    if (typeof role !== "string" || !ROLES.has(role) || typeof content !== "string") {
      throw new Blocked("A message in the request isn't in the form Nymform builds.");
    }
    if (i === 0) {
      if (role !== "system" || content !== SYSTEM_PROMPT) throw new Blocked("The system message isn't Nymform's system prompt.");
      return;
    }
    if (role === "system") throw new Blocked("The request has an extra system message.");
    if (role !== "user" || content === CORRECTION_PROMPT) return;
    const value = parseDescription(content);
    if (!value) throw new Blocked("A user message isn't a Nymform description.");
    descriptions.push({ message: i, value });
  });

  const latest = descriptions[descriptions.length - 1];
  if (!latest) throw new Blocked("The request has no description of your data.");
  if (latest.value.mode !== mode) throw new Blocked("The request's mode doesn't match its description.");
  return descriptions;
}

/** A user turn that is a canonical JSON object with a "nymform" key, or null. */
function parseDescription(content: string): Record<string, unknown> | null {
  let value: unknown;
  try {
    value = JSON.parse(content);
  } catch {
    return null;
  }
  if (!isObject(value) || !has(value, "nymform")) return null;
  if (JSON.stringify(value) !== content) return null;
  return value;
}

/** Request parameters are masked in the text scan, so they must be exactly the known shapes. */
function checkParams(top: Record<string, unknown>): void {
  const bad = () => new Blocked("The request has a parameter Nymform doesn't send.");
  if (has(top, "temperature") && !(typeof top.temperature === "number" && Number.isFinite(top.temperature))) throw bad();
  if (has(top, "max_tokens") && !isCount(top.max_tokens)) throw bad();
  if (has(top, "response_format")) {
    const rf = top.response_format;
    if (!isObject(rf) || Object.keys(rf).length !== 1 || (rf.type !== "json_object" && rf.type !== "text")) throw bad();
  }
  if (has(top, "reasoning")) {
    const r = top.reasoning;
    if (!isObject(r) || !onlyKeys(r, REASONING_KEYS)) throw bad();
    if (has(r, "effort") && !(typeof r.effort === "string" && EFFORTS.has(r.effort))) throw bad();
    if (has(r, "exclude") && typeof r.exclude !== "boolean") throw bad();
    if (has(r, "enabled") && typeof r.enabled !== "boolean") throw bad();
    if (has(r, "max_tokens") && !isCount(r.max_tokens)) throw bad();
  }
  if (has(top, "provider")) {
    const p = top.provider;
    if (!isObject(p) || !onlyKeys(p, PROVIDER_KEYS)) throw bad();
    if (has(p, "zdr") && typeof p.zdr !== "boolean") throw bad();
    if (has(p, "data_collection") && p.data_collection !== "deny" && p.data_collection !== "allow") throw bad();
  }
}

// ---------------------------------------------------------------------------------------------
// Structural check

/**
 * Records a hit and where it was found: [start, end) of the body. It is called once per hit and
 * origin (see Owner.key), so a hit can be reported twice; the first call makes the hit.
 */
type AddHit = (col: { label: string; order: number; origin: AuditOrigin }, variant: AuditVariant, start: number, end: number) => void;
/** One match recorded for a hit, as offsets into the body, and where its text came from. */
interface Match {
  origin: AuditOrigin;
  start: number;
  end: number;
}
type AddStructural = (col: { label: string; order: number }) => void;

function checkDescription(d: Record<string, unknown>, snap: Snapshot, mode: ProdMode, addStructural: AddStructural): void {
  const malformed = () => new Blocked("A description in the request isn't in the form Nymform builds.");
  if (!onlyKeys(d, DESCRIPTION_KEYS)) throw malformed();
  if (typeof d.nymform !== "string" || typeof d.mode !== "string" || !PROD_MODES.has(d.mode)) throw malformed();
  const structureOnly = mode === "structure_only" || d.mode === "structure_only";
  if (has(d, "task") && typeof d.task !== "string") throw malformed();
  if (has(d, "allowed_ranges") && !(Array.isArray(d.allowed_ranges) && d.allowed_ranges.every((a) => typeof a === "string"))) {
    throw malformed();
  }
  checkBlock(d, snap, structureOnly, addStructural);
  if (has(d, "context_ranges")) {
    if (!Array.isArray(d.context_ranges)) throw malformed();
    for (const cr of d.context_ranges) {
      if (!isObject(cr) || !onlyKeys(cr, CONTEXT_RANGE_KEYS)) throw malformed();
      checkBlock(cr, snap, structureOnly, addStructural);
    }
  }
}

/** Checks one described range (the selection or a context range) against the context. */
function checkBlock(b: Record<string, unknown>, snap: Snapshot, structureOnly: boolean, addStructural: AddStructural): void {
  const malformed = () => new Blocked("A description in the request isn't in the form Nymform builds.");
  const range = b.range;
  if (typeof b.sheet !== "string" || !isObject(range) || !onlyKeys(range, RANGE_KEYS) || typeof range.address !== "string") {
    throw malformed();
  }
  if (has(b, "table") && b.table !== null && typeof b.table !== "string") throw malformed();
  if (has(range, "header_row") && range.header_row !== null && typeof range.header_row !== "number") throw malformed();
  for (const k of ["first_data_row", "last_data_row"]) {
    if (has(range, k) && typeof range[k] !== "number") throw malformed();
  }
  for (const k of ["rows_sent", "rows_total"]) {
    if (has(b, k) && typeof b[k] !== "number") throw malformed();
  }
  if (structureOnly && ROW_KEYS.some((k) => has(b, k))) {
    throw new Blocked("A structure-only request can't contain rows.");
  }

  const target = snap.ranges.find((r) => r.sheet === b.sheet && r.address === range.address);
  if (!target) throw new Blocked("The request describes a range that isn't in your selection or its context ranges.");
  const dataRows = checkRowNumbers(b, range, target);

  // Columns must match the context's letters, private flags and treatments, in order.
  const columns = b.columns;
  if (!Array.isArray(columns)) throw malformed();
  if (columns.length !== target.columns.length) {
    throw new Blocked("The request describes a different set of columns from the ones you set.");
  }
  columns.forEach((col: unknown, i) => {
    const expected = target.columns[i]!;
    if (!isObject(col) || !onlyKeys(col, COLUMN_KEYS)) throw malformed();
    if (col.letter !== expected.letter) {
      throw new Blocked("The request describes a different set of columns from the ones you set.");
    }
    if (col.private !== expected.private || col.treatment !== expected.treatment) {
      throw new Blocked(`Column ${expected.label} is described differently from how you set it.`);
    }
    if (has(col, "header") && typeof col.header !== "string") throw malformed();
    if (has(col, "type") && typeof col.type !== "string") throw malformed();
    if (has(col, "note") && col.note !== null && typeof col.note !== "string") throw malformed();
    if (has(col, "stats")) {
      const stats = col.stats;
      if (!isObject(stats) || !onlyKeys(stats, STATS_KEYS)) throw malformed();
      if (!Object.values(stats).every((v) => typeof v === "number")) throw malformed();
      checkStats(stats, dataRows);
    }
  });

  if (!has(b, "rows")) return;
  const rows = b.rows;
  if (!Array.isArray(rows)) throw malformed();
  if (has(b, "rows_sent") && b.rows_sent !== rows.length) throw countsDiffer();
  // The layout comes from the context: its non-excluded columns, in order.
  const layout = target.columns.filter((c) => c.treatment !== "exclude");
  for (const row of rows) {
    if (!Array.isArray(row) || row.length !== layout.length) {
      throw new Blocked("A sample row in the request has the wrong number of cells.");
    }
    row.forEach((cell: unknown, i) => {
      const col = layout[i]!;
      if (cell !== null && typeof cell !== "string" && typeof cell !== "number" && typeof cell !== "boolean") {
        if (col.private) addStructural(col);
        else throw malformed();
        return;
      }
      if (col.private && !isReplaced(cell, col.treatment, snap)) addStructural(col);
    });
  }
}

// Row numbers and counts are numbers Nymform computes, so the text scan masks them (a count that
// equals a private number is a coincidence, not the value). In exchange they must be exactly what
// the range implies, or counts no larger than its data rows.

const countsDiffer = () => new Blocked("A description in the request has row numbers or counts that don't match your selection.");

/** Excel's cell limit: 32,767 characters, so no more words either. */
const MAX_CELL_CHARS = 32767;

/** Checks header_row, first_data_row, last_data_row, rows_sent and rows_total. Returns the data rows. */
function checkRowNumbers(b: Record<string, unknown>, range: Record<string, unknown>, target: CtxRange): number {
  const rect = parseLocalRange(target.address);
  if (!rect || target.hasHeaders === null || target.dataRows === null) {
    throw new Blocked("The check couldn't work out the rows of your selection, so nothing was sent. Select your data, or one cell in it, and choose Refresh from selection.");
  }
  const top = rect.r1 + 1;
  const first = target.hasHeaders ? top + 1 : top;
  const last = rect.r2 + 1;
  if (last - first + 1 !== target.dataRows) throw countsDiffer();
  const expected: [string, number | null][] = [
    ["header_row", target.hasHeaders ? top : null],
    ["first_data_row", first],
    ["last_data_row", last],
  ];
  for (const [key, value] of expected) if (has(range, key) && range[key] !== value) throw countsDiffer();
  if (has(b, "rows_total") && b.rows_total !== target.dataRows) throw countsDiffer();
  if (has(b, "rows_sent") && !isBoundedInt(b.rows_sent, target.dataRows)) throw countsDiffer();
  return target.dataRows;
}

/** blank and distinct count data rows; max_len and avg_words are bounded by Excel's cell size. */
function checkStats(stats: Record<string, unknown>, dataRows: number): void {
  for (const [key, v] of Object.entries(stats)) {
    const ok =
      key === "avg_words"
        ? typeof v === "number" && v <= MAX_CELL_CHARS && /^\d+(?:\.\d)?$/.test(String(v))
        : isBoundedInt(v, key === "max_len" ? MAX_CELL_CHARS : dataRows);
    if (!ok) throw countsDiffer();
  }
}

function isBoundedInt(v: unknown, max: number): boolean {
  return typeof v === "number" && Number.isInteger(v) && v >= 0 && v <= max;
}

function isReplaced(cell: string | number | boolean | null, treatment: string, snap: Snapshot): boolean {
  if (cell === null) return true;
  if (typeof cell !== "string") return false;
  if (snap.isToken(cell)) return true;
  if (treatment === "range" && isRangeRendering(cell)) return true;
  if (treatment === "month" && MONTH_RE.test(cell)) return true;
  return false;
}

/**
 * A `range` rendering: "0", or a bin "lo-hi" where lo = d * 10^k (k >= 1) and hi = lo + 10^k - 1,
 * such as "0-9", "10-19", "80000-89999", or its mirror for negatives, "-89999--80000".
 */
export function isRangeRendering(s: string): boolean {
  if (s === "0") return true;
  if (!RANGE_RE.test(s)) return false;
  const m = /^(-?)(\d+)-(-?)(\d+)$/.exec(s);
  if (!m) return false;
  const [sign1, a, sign2, b] = [m[1] ?? "", m[2] ?? "", m[3] ?? "", m[4] ?? ""];
  if (sign1 === "" && sign2 === "") return isBin(a, b);
  if (sign1 === "-" && sign2 === "-") return isBin(b, a);
  if (sign1 === "-" && sign2 === "" && b === "0") return isBin(b, a);
  return false;
}

function isBin(lo: string, hi: string): boolean {
  if (lo === "0") return hi === "9";
  const m = /^([1-9])(0+)$/.exec(lo);
  return m !== null && hi === m[1]! + "9".repeat(m[2]!.length);
}

// ---------------------------------------------------------------------------------------------
// Text scan

/**
 * Where a variant counts:
 * - "any": anywhere.
 * - "word": only on word boundaries.
 * - "free": on word boundaries and only inside free text (see BodyMask.free).
 * - "plain": a private value that is a plain number of fewer than 7 digits before any decimals,
 *   such as a ZIP code, an amount or a date's serial. In text (the question, notes, headers, text
 *   cells) it counts anywhere, digits next to it or not: "934961234" (a ZIP+4 typed without its
 *   hyphen) holds 93496, and "Transfer to 48392001" holds 483920. Inside a JSON number of a
 *   description (a cell or stat stored as a number) it counts only where no digit is right next
 *   to it and not as the decimals of a number with one decimal point (see noDigitNextTo): the
 *   number 934961 doesn't hold 93496, nor does a longitude -120.93496.
 * - "plainText": a number the user reads written the same way ("83500" for "$83,500", "5,25" for
 *   5.25), only inside string content and only where no digit is right next to it, in free text
 *   too: "5,25" is not in "5,251".
 * - "number": a number the user reads in other forms ("83,500", "83.5k"), only with no digit or
 *   decimal mark continuing it on either side, and only inside string content. In free text a
 *   grouped form ends only where a digit, or its own separator and 3 digits, don't continue it
 *   (see groupedNumberEnds).
 */
type MatchRule = "any" | "word" | "free" | "number" | "plain" | "plainText";

interface Owner {
  label: string;
  order: number;
  kind: AuditVariant;
  rule: MatchRule;
  /**
   * What textScan records once: the hit this owner reports (column and kind) and its origin. A hit
   * found in a column's original header is still looked for in its values, and the other way round,
   * so the pane can show both (see AddHit); the hit itself is the first match either way.
   */
  key: string;
  /** Whether the text came from the column's values or its original header (see AuditOrigin). */
  origin: AuditOrigin;
}

function textScan(
  body: string,
  parsed: Record<string, unknown>,
  descriptions: Description[],
  snap: Snapshot,
  strategy: ScanStrategy,
  addHit: AddHit,
): void {
  const { patterns, owners, digitPatterns, digitOwners, codePatterns, codeOwners, dateOwners } = buildVariants(snap);
  if (patterns.length === 0 && dateOwners.size === 0 && digitPatterns.length === 0 && codePatterns.length === 0) return;

  const { mask, content, free, nested, number } = buildMask(body, parsed, descriptions, snap);
  const n = body.length;
  // unmasked[i] = number of unmasked characters before position i; inString, inFree and inNumber likewise.
  const unmasked = new Int32Array(n + 1);
  const inString = new Int32Array(n + 1);
  const inFree = new Int32Array(n + 1);
  const inNumber = new Int32Array(n + 1);
  for (let i = 0; i < n; i++) {
    unmasked[i + 1] = unmasked[i]! + (mask[i] === 1 ? 0 : 1);
    inString[i + 1] = inString[i]! + content[i]!;
    inFree[i + 1] = inFree[i]! + free[i]!;
    inNumber[i + 1] = inNumber[i]! + number[i]!;
  }
  const wordChar = wordCharMap(body);
  const lower = lowerUnits(body);
  // Matching runs on `folded`: accents, characters that can't be seen, form feeds and vertical
  // tabs dropped, and other escaped whitespace read as a space (see foldBody). `back` maps its
  // positions to the body; every rule below looks at `folded` around a match, and the mask, string
  // content and free text at the body.
  const { folded, at: back, gap } = foldBody(body, lower, nested);
  const fn = folded.length;
  const fw = back === null ? wordChar : wordCharsAt(wordChar, back);
  const bodyAt = (k: number) => (back === null ? k : k < fn ? back[k]! : n);
  const digitAt = (k: number) => k >= 0 && k < fn && fw[k] === 1 && isAsciiDigit(folded.charCodeAt(k));
  // A decimal mark: "." anywhere, "," only inside string content (between JSON values it separates them).
  const markAt = (k: number) => k >= 0 && (folded[k] === "." || (folded[k] === "," && content[bodyAt(k)] === 1));
  // A dropped character still separates what is on either side of it: the rules below look past a
  // match's ends with these, so "Mr\u200bWatanabe" holds the word "Watanabe" and "93496\u200b94103"
  // the number 93496, as they would with a space there. -1 stands for "nothing next to it".
  // The unit before folded[k], when nothing was dropped between them.
  const left = (k: number) => (k <= 0 || (gap !== null && gap[k] === 1) ? -1 : k - 1);
  // folded[k], when nothing was dropped between it and the unit before it.
  const right = (k: number) => (k <= 0 || (gap !== null && gap[k] === 1) ? -1 : k);
  const near: Near = { digitAt, markAt, left, right };
  const done = new Set<string>();

  findAll(folded, patterns, strategy, (p, pos) => {
    const list = owners[p]!;
    const pattern = patterns[p]!;
    const next = pos + pattern.length;
    const start = bodyAt(pos);
    // A match takes the accents that follow its last letter ("José" for "jose").
    const end = bodyAt(next);
    if (unmasked[end]! - unmasked[start]! === 0) return false;
    let boundary: boolean | undefined;
    let plain: boolean | undefined;
    let bare: boolean | undefined;
    const inText = () => inString[end]! - inString[start]! === end - start;
    const isFree = () => inFree[end]! - inFree[start]! === end - start;
    for (const o of list) {
      if (done.has(o.key)) continue;
      if (o.rule === "word" || o.rule === "free") {
        if (o.rule === "free" && !isFree()) continue;
        // A side needs a boundary only where the match's own end is a word character. A dropped
        // character is one (see left and right).
        boundary ??= (fw[pos] === 0 || fw[left(pos)] !== 1) && (fw[next - 1] === 0 || fw[right(next)] !== 1);
        if (!boundary) continue;
      } else if (o.rule === "plain" || o.rule === "plainText") {
        // A private value in text (the question, notes, headers, a text cell): a digit next to it
        // doesn't matter. Inside a JSON number of a description (a row cell or a stat stored as a
        // number) it does, and so do the decimals of a number with one decimal point, such as a
        // longitude. A number read only from how a cell is shown ("plainText") needs no digit next
        // to it anywhere.
        if (o.rule === "plainText" || inNumber[end]! - inNumber[start]! === end - start) {
          plain ??= noDigitNextTo(pos, next, pattern, folded, near, () => inNumber[end]! - inNumber[start]! === end - start);
          if (!plain) continue;
        }
        if (o.rule === "plainText" && !inText()) continue;
      } else if (o.rule === "number") {
        if (!inText()) continue;
        const separator = isFree() ? groupSeparatorOf(pattern) : null;
        if (separator !== null) {
          bare ??= groupedNumberEnds(pos, next, separator, folded, near);
        } else {
          // "93.496" is not in "-93.496157", and "5,390" is not the two JSON numbers in "5,390,0".
          // A decimal mark counts only with a digit behind it, so "earning 5,390." is still found.
          // A form ending in a letter ("83.5k") needs a non-letter after it.
          bare ??=
            !digitAt(left(pos)) &&
            !(markAt(left(pos)) && digitAt(left(left(pos)))) &&
            numberEndsAt(next, pattern, folded, near) &&
            !(LETTER_END_RE.test(pattern) && fw[right(next)] === 1);
        }
        if (!bare) continue;
      }
      done.add(o.key);
      addHit(o, o.kind, start, end);
    }
    return list.every((o) => done.has(o.key));
  });

  // Date phrases anywhere in the body that mean a private date, however they are written
  // ("3/15/85", "4th of July 1988", "15. März 1985"; see findDates in transform.ts): in the
  // question, notes, headers, sheet names, history and row cells. JSON's escapes of whitespace
  // ("\n", "\\n", a form feed's "\f") are read as spaces, and escapes of characters that can't be
  // seen ("\b", "\u200b") are skipped (see escapesForDates). A match counts where it touches
  // unmasked text, like a word; the loosest forms (see DateMatch.weak: 8 digits in a row, numbers
  // separated only by spaces, a run of 2-digit numbers such as a phone number) only in free text,
  // and so do the dates of a column's 8-digit cells (owners with the "free" rule).
  let phrases: DateMatch[] | null = null;
  const bodyDates = (): DateMatch[] => {
    if (phrases !== null) return phrases;
    // A backslash typed before n, r, t, f or b is read both ways: kept, for "15\Nov\1985", and as
    // a space, for "3\n15\n1985" typed with backslashes.
    const exact = escapesForDates(body, lower, nested, false);
    const every = escapesForDates(body, lower, nested, true);
    phrases = every === exact ? findDates(exact) : [...findDates(exact), ...findDates(every)];
    return phrases;
  };
  if (dateOwners.size > 0) {
    for (const { start, end, weak, dates } of bodyDates()) {
      if (unmasked[end]! - unmasked[start]! === 0) continue;
      const isFree = inFree[end]! - inFree[start]! === end - start;
      if (weak && !isFree) continue;
      for (const date of dates) {
        for (const o of dateOwners.get(dateKey(date)) ?? []) {
          if (done.has(o.key) || (o.rule === "free" && !isFree)) continue;
          done.add(o.key);
          // escapesForDates keeps the body's positions, so the phrase is at the same place in it.
          addHit(o, o.kind, start, end);
        }
      }
    }
  }

  if (digitPatterns.length > 0) {
    // Digits typed with other punctuation, "(831) 555-0199" for 831.555.0199, or as a date,
    // "2024-01-15" for invoice number 20240115. `at` maps each projected digit back to the body, so
    // the mask applies as above.
    const { digits, at } = digitsProjection(body);
    findAll(digits, digitPatterns, strategy, (p, start) => {
      const list = digitOwners[p]!;
      const first = at[start]!;
      const last = at[start + digitPatterns[p]!.length - 1]!;
      if (unmasked[last + 1]! - unmasked[first]! === 0) return false;
      for (const o of list) {
        if (done.has(o.key)) continue;
        done.add(o.key);
        addHit(o, o.kind, first, last + 1);
      }
      return list.every((o) => done.has(o.key));
    });
  }

  if (codePatterns.length > 0) {
    // A code of 6+ letters and digits with only some of its separators, or other ones, or with a
    // letter or digit touching it ("AB 1234CD", "AB.1234.CD", "xAB1234CD" for "AB-1234-CD"; see
    // codeLetters in transform.ts). `at` maps each projected unit to `folded`, and so to the body.
    // A match counts only where all of it is unmasked text: a stand-in and the number after it
    // ("ID_001 2024") are not a private "ID-0012".
    const { letters, at } = lettersProjection(folded, fw);
    findAll(letters, codePatterns, strategy, (p, start) => {
      const list = codeOwners[p]!;
      const first = bodyAt(at[start]!);
      const last = bodyAt(at[start + codePatterns[p]!.length - 1]!);
      if (unmasked[last + 1]! - unmasked[first]! !== last + 1 - first) return false;
      for (const o of list) {
        if (done.has(o.key)) continue;
        done.add(o.key);
        addHit(o, o.kind, first, last + 1);
      }
      return list.every((o) => done.has(o.key));
    });
  }
}

/** The body's word-character map (see wordCharMap) at each unit of a folded text: `back` maps the text to the body. */
function wordCharsAt(wordChar: Uint8Array, back: Int32Array): Uint8Array {
  const out = new Uint8Array(back.length);
  for (let k = 0; k < back.length; k++) out[k] = wordChar[back[k]!]!;
  return out;
}

/** The code points of the JSON escapes "\n", "\r", "\t", "\f" and "\b". */
const ESCAPED: Readonly<Record<string, number>> = { n: 0x0a, r: 0x0d, t: 0x09, f: 0x0c, b: 0x08 };

/**
 * The character a JSON escape in the body stands for (`cp`), with its letter at `j` after the
 * backslashes, when it is a whitespace control ("\n", "\r", "\t", "\f", "\u000b", "\u0085") or a
 * character that can't be seen ("\b", "\u200b", "\u0007", or a tag character's two "\u" escapes,
 * each after `width` backslashes); null for anything else. `end` is just past the escape.
 */
function escapeAt(body: string, j: number, width: number): { cp: number; end: number } | null {
  const c = body[j];
  if (c === undefined) return null;
  const simple = ESCAPED[c];
  if (simple !== undefined) return { cp: simple, end: j + 1 };
  if (c !== "u") return null;
  let cp = hexAt(body, j + 1);
  if (cp < 0) return null;
  let end = j + 5;
  if (cp >= 0xd800 && cp <= 0xdbff) {
    // A character outside the Basic Multilingual Plane is two escapes, as in "\udb40\udc20".
    let k = end;
    while (k < end + width && body.charCodeAt(k) === 0x5c) k++;
    const low = k === end + width && body[k] === "u" ? hexAt(body, k + 1) : -1;
    if (low < 0xdc00 || low > 0xdfff) return null;
    cp = (cp - 0xd800) * 0x400 + (low - 0xdc00) + 0x10000;
    end = k + 5;
  }
  return isWhitespaceControl(cp) || isIgnorableCodePoint(cp) ? { cp, end } : null;
}

/** The value of the 4 hex digits at s[i], or -1. */
function hexAt(s: string, i: number): number {
  const hex = s.slice(i, i + 4);
  return /^[0-9a-fA-F]{4}$/.test(hex) ? parseInt(hex, 16) : -1;
}

/** Tab, line feed, vertical tab, form feed, carriage return and NEL: the whitespace controls. */
function isWhitespaceControl(cp: number): boolean {
  return (cp >= 0x09 && cp <= 0x0d) || cp === 0x85;
}

/** Finds what foldBody changes: an accent, a character that can't be seen, or a backslash. */
// eslint-disable-next-line no-control-regex
const FOLD_RE = /[\u0300-\u036f\\\p{Default_Ignorable_Code_Point}\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f]/u;

/**
 * The lowercased body as the value scan reads it. Dropped: combining accents (U+0300–036F), so
 * "José" typed decomposed is "jose" and "İ" decomposed is "i"; the characters matching drops (see
 * isDroppedCodePoint: those that can't be seen, form feed and vertical tab), so "Ma\u200bria" is
 * "maria" and "934\f96" is "93496"; and the JSON escapes of those ("\b", "\u0007", "\f",
 * "\u000b"). Each escape of another whitespace control ("\n", "\t"), single or double (see
 * isEscapeRun), becomes one space, so "93\n496" reads "93 496" and "Maria\nLopez" reads "maria
 * lopez". Backslashes typed before a letter stay (see isEscapeRun). `at` maps each unit of
 * `folded` back to the body; null when nothing changed. `gap[k]` is 1 where characters matching
 * drops were taken out right before the k-th unit (at k = folded.length: after the last one), so
 * the rules around a match still see a boundary there; null when none were.
 */
function foldBody(body: string, lower: string, nested: Uint8Array): { folded: string; at: Int32Array | null; gap: Uint8Array | null } {
  if (!FOLD_RE.test(body)) return { folded: lower, at: null, gap: null };
  const n = body.length;
  const codes = new Uint16Array(n);
  const at = new Int32Array(n);
  const gap = new Uint8Array(n + 1);
  let gaps = false;
  let k = 0;
  for (let i = 0; i < n; ) {
    const c = body.charCodeAt(i);
    if (c === 0x5c) {
      let j = i + 1;
      while (body.charCodeAt(j) === 0x5c) j++;
      const run = j - i;
      const width = run % 2 === 1 ? 1 : 2;
      const escape = isEscapeRun(run, nested[i] === 1) ? escapeAt(body, j, width) : null;
      // Backslashes before the escape were typed.
      const typed = escape === null ? run : run - width;
      for (let m = 0; m < typed; m++) {
        codes[k] = 0x5c;
        at[k++] = i + m;
      }
      if (escape === null) {
        i = j;
        continue;
      }
      if (isDroppedCodePoint(escape.cp)) {
        gap[k] = 1;
        gaps = true;
      } else {
        codes[k] = 0x20;
        at[k++] = i + typed;
      }
      i = escape.end;
      continue;
    }
    // Before accents: the combining grapheme joiner U+034F is one of the characters that can't be seen.
    const skip = droppedLength(body, i);
    if (skip > 0) {
      gap[k] = 1;
      gaps = true;
      i += skip;
      continue;
    }
    const l = lower.charCodeAt(i);
    if (isAccentMark(l)) {
      i++;
      continue;
    }
    codes[k] = l;
    at[k++] = i;
    i++;
  }
  let folded = "";
  for (let i = 0; i < k; i += 8192) folded += String.fromCharCode(...codes.subarray(i, Math.min(k, i + 8192)));
  return { folded, at: at.subarray(0, k), gap: gaps ? gap.subarray(0, k + 1) : null };
}

/** A run of U+200B, which findDates skips, standing for an escape of a character that can't be seen. */
const HIDDEN = "\u200b";

/**
 * The body (lowercased) for reading dates, positions unchanged: each JSON escape of whitespace,
 * single or double (`\n`, `\\n`, `\t`, `\r`, `\f`, `\u000b`; see isEscapeRun), becomes as many
 * spaces, so a date split by a line break is still read; each escape of a character that can't be
 * seen (`\b`, `\u200b`) becomes as many U+200B, which findDates skips. Any other `\u` escape
 * becomes spaces, so its hex digits don't join a date. Backslashes typed before one of those
 * letters stay, so "15\Nov\1985" is read, unless `every`, which reads them as escapes too.
 */
function escapesForDates(body: string, lower: string, nested: Uint8Array, every: boolean): string {
  if (!body.includes("\\")) return lower;
  const parts: string[] = [];
  let last = 0;
  for (let i = body.indexOf("\\"); i !== -1; i = body.indexOf("\\", i)) {
    let j = i + 1;
    while (body.charCodeAt(j) === 0x5c) j++;
    const run = j - i;
    const letter = lower[j];
    const hex = letter === "u" && hexAt(body, j + 1) >= 0;
    if (!hex && !(letter === "n" || letter === "r" || letter === "t" || letter === "f" || letter === "b")) {
      i = j;
      continue;
    }
    if (!hex && !every && !isEscapeRun(run, nested[i] === 1)) {
      i = j + 1;
      continue;
    }
    const escape = escapeAt(body, j, run % 2 === 1 ? 1 : 2);
    const end = escape?.end ?? (hex ? j + 5 : j + 1);
    parts.push(lower.slice(last, i), (escape !== null && !isWhitespaceControl(escape.cp) ? HIDDEN : " ").repeat(end - i));
    last = end;
    i = end;
  }
  if (last === 0) return lower;
  parts.push(lower.slice(last));
  return parts.join("");
}

/**
 * True when a run of `run` backslashes before a letter in the body escapes that letter (a line
 * break's "\n", a form feed's "\f"): an odd run (text written into JSON once), or a run of 2, 6,
 * 10 ... inside a string written into JSON twice (`nested`: a description's or an earlier reply's
 * strings). Otherwise the backslashes were typed before the letter, as in a path "C:\files".
 */
function isEscapeRun(run: number, nested: boolean): boolean {
  return run % 2 === 1 || (nested && run % 4 === 2);
}

/** A variant that ends in a letter, such as "83.5k". */
const LETTER_END_RE = /\p{L}$/u;

/**
 * What the number rules look at around a match in the folded body: a digit or a decimal mark at a
 * position, and the positions next to one, which are -1 ("nothing next to it") where a dropped
 * character stands between them (see textScan).
 */
interface Near {
  digitAt: (k: number) => boolean;
  markAt: (k: number) => boolean;
  /** The unit before folded[k], or -1. */
  left: (k: number) => number;
  /** folded[k] itself, or -1 when something was dropped right before it. */
  right: (k: number) => number;
}

/**
 * True when no digit is right next to a plain number matched at [start, end) of the folded body
 * (the "plain" rule outside free text, and the "plainText" rule): 93496 is in "93496,94103",
 * "=IF(B2=93496,1,0)", "93496.5" and "1.93496", but not in "934961". A comma or decimal point next
 * to it doesn't matter, except that inside a JSON number of a description (`inNumber`: a row cell
 * or a stat stored as a number) the decimals of a number with one decimal point are not a match:
 * -116.093496 (a longitude) doesn't hold 93496, but "831.555.0199" typed anywhere holds 0199.
 * Zeros before it don't count when no digit comes before them, so "02139", "004521" and "#04521"
 * hold 2139 and 4521, but "102139" doesn't hold 2139. Zeros after a number with decimals don't
 * count either: "12.750" is 12.75. A dropped character is not "next to" anything:
 * "93496\u200b94103" holds 93496.
 */
function noDigitNextTo(start: number, end: number, pattern: string, text: string, near: Near, inNumber: () => boolean): boolean {
  const { digitAt, left, right } = near;
  let before = left(start);
  while (digitAt(before) && text[before] === "0") before = left(before);
  if (digitAt(before)) return false;
  if (before >= 0 && text[before] === "." && digitAt(left(before)) && inNumber()) {
    let k = left(before);
    while (digitAt(k)) k = left(k);
    if (k < 0 || text[k] !== ".") return false;
  }
  let after = right(end);
  if (pattern.includes(".") || pattern.includes(",")) while (digitAt(after) && text[after] === "0") after = right(after + 1);
  return !digitAt(after);
}

/**
 * True when a number form matched just before `end` isn't continued by more of a number. Zero
 * decimals still end it: "125,000.00", "125.000,00" and "83500.0" hold 125000 and 83500, and a
 * form with decimals may have zeros added ("1,234.50" for 1234.5). Other decimals ("93.496157")
 * and a group of three after the form's own separator ("5,390,000") continue it.
 */
function numberEndsAt(end: number, pattern: string, text: string, near: Near): boolean {
  const { digitAt, markAt, right } = near;
  let k = right(end);
  if (/[.,]/.test(pattern)) while (digitAt(k) && text[k] === "0") k = right(k + 1);
  if (k === end && markAt(k) && digitAt(right(k + 1))) {
    const mark = text[k]!;
    let z = right(k + 1);
    let zeros = true;
    let digits = 0;
    while (digitAt(z)) {
      if (text[z] !== "0") zeros = false;
      digits++;
      z = right(z + 1);
    }
    if (!zeros || (pattern.includes(mark) && digits === 3)) return false;
    k = z;
  }
  return !digitAt(k) && !(markAt(k) && digitAt(right(k + 1)));
}

/** A number grouped with one separator throughout ("93 496", "5,390.25", "8,35,000"): the separator. Null for any other form. */
function groupSeparatorOf(pattern: string): string | null {
  const m = /^\d{1,3}([,.' _])\d{2,3}(?:\1\d{2,3})*(?:[.,]\d+)?$/u.exec(pattern);
  return m ? m[1]! : null;
}

/**
 * In free text, true when a grouped number form matched at [start, end) of the folded body isn't
 * continued by more of it. Grouped with a space of any kind (folded to " "), only a digit right
 * next to it continues it: a list typed one number after another, "40 000 83 500" or "40 000\n83
 * 500", holds 83 500, and "5 390 000" holds 390 000. Grouped with "," "." "'" or "_", a digit or its
 * own separator and 3 digits on either side continue it: "5,390,000" doesn't hold 5,390, but
 * "5,390.25" and "93 496,94 103" hold 5,390 and 93 496. A decimal part after it doesn't matter.
 */
function groupedNumberEnds(start: number, end: number, separator: string, text: string, near: Near): boolean {
  const { digitAt, left, right } = near;
  const before = left(start);
  const after = right(end);
  if (digitAt(before) || digitAt(after)) return false;
  if (separator === " ") return true;
  if (before >= 0 && text[before] === separator && digitAt(left(before))) return false;
  if (after < 0 || text[after] !== separator) return true;
  const a = right(after + 1);
  const b = right(a + 1);
  return !(digitAt(a) && digitAt(b) && digitAt(right(b + 1)));
}

/**
 * Characters a typed number may have between its digits: spaces, ( ) . + - _ / and dashes, also
 * full-width. Characters that can't be seen are skipped (see isIgnorableCodePoint).
 */
const DIGIT_SEPARATOR_RE = /[\s\u0085().+\-_/\u2010-\u2015\u2212\ufe58\ufe63\uff08\uff09\uff0b\uff0d\uff0e\uff0f]/u;

/**
 * The body's digits (full-width, Arabic-Indic, Persian and Devanagari ones as 0-9), where only
 * separators, characters that can't be seen, and the escapes of either (`\n`, `\\n`, a form feed's
 * `\f`, `\b`, `\u200b`) between two digits are dropped. Anything else, such as the commas, quotes
 * and keys between JSON values, ends a run with "|", so separate numbers never run together.
 */
function digitsProjection(body: string): { digits: string; at: Int32Array } {
  const out: string[] = [];
  const at: number[] = [];
  let gap = false;
  for (let i = 0; i < body.length; ) {
    const c = body.charCodeAt(i);
    const f = c < 0x80 ? c : foldUnit(c);
    const digit = f >= 0x30 && f <= 0x39 ? f : 0;
    if (digit !== 0) {
      if (gap && out.length > 0 && out[out.length - 1] !== "|") {
        out.push("|");
        at.push(i);
      }
      gap = false;
      out.push(String.fromCharCode(digit));
      at.push(i);
      i++;
      continue;
    }
    if (c === 0x5c) {
      let j = i;
      while (body[j] === "\\") j++;
      const escape = escapeAt(body, j, (j - i) % 2 === 1 ? 1 : 2);
      if (escape !== null) {
        i = escape.end;
        continue;
      }
      gap = true;
      // The hex digits of a \u escape are not digits of the text.
      i = body[j] === "u" && hexAt(body, j + 1) >= 0 ? j + 5 : j;
      continue;
    }
    const skip = droppedLength(body, i);
    if (skip > 0) {
      i += skip;
      continue;
    }
    if (!DIGIT_SEPARATOR_RE.test(body[i]!)) gap = true;
    i++;
  }
  return { digits: out.join(""), at: Int32Array.from(at) };
}

/**
 * The folded body's letters and digits (its word characters, `fw`), where separators between them
 * (see isCodeSeparatorAt: spaces, - _ . / \ : · and look-alikes, but not the colon of a cell range
 * such as "A10:B20") are dropped. Anything else ends a run with "|". `at` maps each projected unit
 * back to `folded`.
 */
function lettersProjection(folded: string, fw: Uint8Array): { letters: string; at: Int32Array } {
  const out: string[] = [];
  const at: number[] = [];
  let gap = false;
  for (let k = 0; k < folded.length; k++) {
    if (fw[k] === 1) {
      if (gap && out.length > 0 && out[out.length - 1] !== "|") {
        out.push("|");
        at.push(k);
      }
      gap = false;
      out.push(folded[k]!);
      at.push(k);
    } else if (!isCodeSeparatorAt(folded, k)) {
      gap = true;
    }
  }
  return { letters: out.join(""), at: Int32Array.from(at) };
}

interface Variants {
  patterns: string[];
  owners: Owner[][];
  /** Digits-only forms of 7+ digits, scanned in the body's digits projection. */
  digitPatterns: string[];
  digitOwners: Owner[][];
  /** Codes' letters and digits (see codeLetters), scanned in the body's letters projection. */
  codePatterns: string[];
  codeOwners: Owner[][];
  /** Private dates by dateKey ("1985-03-15"), compared with every date phrase in the body. */
  dateOwners: Map<string, Owner[]>;
}

/** A private number of this many digits or more that starts with 0 is also scanned for without it. */
const MIN_TRUNK_DIGITS = 9;

/** A group of 4 digits that reads as a year: questions name years all the time. */
function isYearLike(digits: string): boolean {
  return digits.length === 4 && Number(digits) >= 1900 && Number(digits) <= 2100;
}

const HAS_DIGIT_RE = /\p{Nd}/u;
const FOUR_DIGITS_RE = /\p{Nd}{4}/u;

/** A plain number ("93496", "-120.934961", "5.25"), with the digits before any decimals. */
const PLAIN_NUMBER_RE = /^[-+]?(\d+)(?:\.\d+)?$/u;
/** A number grouped with one kind of separator ("2,345", "83.500", "83 500"), and optional decimals. */
const GROUPED_NUMBER_RE = /^[-+]?(\d{1,3}([,.' \u00a0\u202f])\d{3}(?:\2\d{3})*)(?:[.,]\d+)?$/u;

/**
 * The match rule for a value's whole forms (exact, normalized, JSON-escaped): "plain" for a plain
 * number of fewer than 7 digits before any decimals (a ZIP code, an amount, a date's serial), so
 * it isn't found inside other numbers; "number" for such a number shown grouped ("2,345"), which
 * outside string content could only be two JSON numbers side by side ("512,34567"); "any" for
 * everything else, and for numbers of 7+ digits (a phone number typed with a country code).
 */
function wholeValueRule(v: string): MatchRule {
  const plain = PLAIN_NUMBER_RE.exec(v);
  if (plain) return plain[1]!.length < MIN_DIGITS ? "plain" : "any";
  const grouped = GROUPED_NUMBER_RE.exec(v);
  if (grouped) return grouped[1]!.replace(/\D/gu, "").length < MIN_DIGITS ? "number" : "any";
  return "any";
}

const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}$/u;

/**
 * The dates (as dateKey) a column's date list names: an ISO date ("1985-03-15", as
 * buildAuditContext lists them) as it is, and other date text ("March 15, 1985") read with
 * textDates, every reading.
 */
function dateKeysOf(dates: readonly string[]): Set<string> {
  const keys = new Set<string>();
  for (const d of dates) {
    if (ISO_DATE_RE.test(d)) keys.add(d);
    else for (const date of textDates(d)) keys.add(dateKey(date));
  }
  return keys;
}

/**
 * The dates (as dateKey) of a column's 8-digit cells: its `looseDates` (ISO dates as they are,
 * 8-digit text read with looseDatesOf), or, when the context doesn't list them, its values read
 * with looseDatesOf. None for a column that doesn't hold dates stored as 8 digits.
 */
function looseKeysOf(col: CtxColumn): Set<string> {
  const keys = new Set<string>();
  const other: string[] = [];
  for (const d of col.looseDates ?? col.values) {
    if (col.looseDates !== null && ISO_DATE_RE.test(d)) keys.add(d);
    else other.push(d);
  }
  for (const date of looseDatesOf(other)) keys.add(dateKey(date));
  return keys;
}

/** The column a variant stands for, and whether it came from the column's values or its original header. */
interface Source {
  label: string;
  order: number;
  origin: AuditOrigin;
}

/** Distinct lowercased variants and, for each, the columns and kinds it stands for. */
function buildVariants(snap: Snapshot): Variants {
  const index = new Map<string, number>();
  const patterns: string[] = [];
  const owners: Owner[][] = [];
  const digitIndex = new Map<string, number>();
  const digitPatterns: string[] = [];
  const digitOwners: Owner[][] = [];
  const codeIndex = new Map<string, number>();
  const codePatterns: string[] = [];
  const codeOwners: Owner[][] = [];
  const dateOwners = new Map<string, Owner[]>();

  const add = (text: string, col: Source, kind: AuditVariant, rule: MatchRule, min: number) => {
    if (text.length < min) return;
    const low = foldText(text);
    if (low.length < min) return;
    let p = index.get(low);
    if (p === undefined) {
      p = patterns.length;
      index.set(low, p);
      patterns.push(low);
      owners.push([]);
    }
    const list = owners[p]!;
    const canary = kind === "canary";
    for (const o of list) {
      if (o.order !== col.order) continue;
      // One owner per column and match rule; the first kind added wins (exact before json, and so
      // on). Canaries are reported as their own kind.
      if (o.rule === rule && (o.kind === "canary") === canary) return;
      // A number form adds nothing where the same text already counts anywhere or as a plain
      // number (a "number" form nothing where it counts as "plainText" either: that rule is
      // wider), and a free-text form nothing where it counts anywhere or on any word boundary.
      if (o.kind === "canary" || (rule !== "number" && rule !== "plainText" && rule !== "free")) continue;
      if (o.rule === "any") return;
      if (rule === "free" && o.rule === "word") return;
      if ((rule === "number" || rule === "plainText") && o.rule === "plain") return;
      if (rule === "number" && o.rule === "plainText") return;
    }
    list.push({ label: col.label, order: col.order, kind, rule, key: `${col.order}|${kind}|${col.origin}`, origin: col.origin });
  };

  const addDigits = (digits: string, col: Source) => {
    let p = digitIndex.get(digits);
    if (p === undefined) {
      p = digitPatterns.length;
      digitIndex.set(digits, p);
      digitPatterns.push(digits);
      digitOwners.push([]);
    }
    const list = digitOwners[p]!;
    if (list.some((o) => o.order === col.order)) return;
    list.push({ label: col.label, order: col.order, kind: "digits", rule: "any", key: `${col.order}|digits|${col.origin}`, origin: col.origin });
  };

  // A code of 6+ letters and digits (see codeLetters), found with some or other separators, or
  // with a letter or digit touching it.
  const addCode = (letters: string, col: Source) => {
    let p = codeIndex.get(letters);
    if (p === undefined) {
      p = codePatterns.length;
      codeIndex.set(letters, p);
      codePatterns.push(letters);
      codeOwners.push([]);
    }
    const list = codeOwners[p]!;
    if (list.some((o) => o.order === col.order)) return;
    list.push({ label: col.label, order: col.order, kind: "normalized", rule: "any", key: `${col.order}|normalized|${col.origin}`, origin: col.origin });
  };

  // add() folds every variant the same way the body is folded (case, accents, look-alike
  // apostrophes, digits, hyphens and spaces, full-width letters, Turkish i, final sigma), so
  // "O’Sullivan" and "O'Sullivan" are one variant. `whole` is the rule for the value's whole
  // forms (see wholeValueRule).
  const addForm = (v: string, col: Source, whole: MatchRule) => {
    const trimmed = v.trim();
    // A form equal to one already added adds nothing (the first kind wins), so it is skipped.
    add(v, col, "exact", whole, MIN_VARIANT);
    const normalized = trimmed.replace(/\s+/g, " ");
    if (normalized !== v) add(normalized, col, "normalized", whole, MIN_VARIANT);
    const once = JSON.stringify(v).slice(1, -1);
    if (once !== v) {
      add(once, col, "json", whole, MIN_VARIANT);
      add(JSON.stringify(once).slice(1, -1), col, "json", whole, MIN_VARIANT);
    }
    const parts = trimmed.split(/\s+/);
    if (parts.length > 1) {
      // Combining marks stay in their word, so the NFD form of "Muñoz" is one word. "_" splits
      // words, as it does in the body ("maria_lopez" holds "maria").
      for (const w of trimmed.split(/[^\p{L}\p{M}\p{N}]+/u)) {
        // Digit groups are added by addDigitGroups.
        if (/^\p{Nd}+$/u.test(w)) continue;
        add(w, col, "word", "word", MIN_VARIANT);
      }
      // Written without its spaces, as Japanese and Chinese names usually are ("山田太郎" for
      // "山田 太郎"), or with "_" for them ("maria_lopez"). Word boundaries apply only at Latin
      // letters and digits.
      add(parts.join(""), col, "normalized", "word", MIN_VARIANT);
      add(parts.join("_"), col, "normalized", "word", MIN_VARIANT);
    }
  };

  // A name typed without its apostrophes or hyphens, on word boundaries (see joinedNames): each
  // word of it that has one ("OBrien" of "Mary O'Brien", "LopezRuiz" of "Maria Lopez-Ruiz"), and
  // a one-word value with its hyphens as spaces too ("Anne Marie" for "Anne-Marie"). The other
  // words of a longer value are words of the value already.
  const addJoined = (v: string, col: Source) => {
    const parts = v.trim().split(/\s+/);
    for (const part of parts) {
      for (const joined of joinedNames(part)) add(joined, col, parts.length === 1 ? "normalized" : "word", "word", MIN_VARIANT);
    }
  };

  // A number of 7+ digits (the digits of a whole value, or of a part between letters), found
  // with any separators: as a digits-only variant, in the body's digits projection, without a
  // trunk 0, and without a country code.
  const addDigitForms = (text: string, col: Source) => {
    if (!HAS_DIGIT_RE.test(text)) return;
    const digits = foldText(text).replace(/\D/g, "");
    if (digits.length < MIN_DIGITS) return;
    add(digits, col, "digits", "any", MIN_VARIANT);
    addDigits(digits, col);
    // A number stored with its trunk 0, typed in international form: "+44 20 7946 0958" for
    // "020 7946 0958".
    if (digits.length >= MIN_TRUNK_DIGITS && digits.startsWith("0")) addDigits(digits.slice(1), col);
    // And one stored in international form, typed as dialled at home: "(831) 555-0199" for
    // "+18315550199".
    for (const national of withoutCountryCode(text)) addDigits(national, col);
  };

  // The digit groups of a value that isn't a number or a date, such as the parts of a phone
  // number, SSN or card number. A group of 7+ digits counts anywhere on word boundaries. A group
  // of 4 to 6 digits ("0199" of "(831) 555-0199", "6789" of "123-45-6789") counts only in free
  // text a person typed or Nymform copied from labels (the question, notes, headers, sheet and
  // table names, earlier turns), never in sample-row cells or JSON numbers, where counts and
  // other numbers matched it so often that requests blocked for no reason. A group that reads as
  // a year (1900 to 2100) is left out: questions name years all the time.
  const addDigitGroups = (v: string, col: Source) => {
    const t = v.trim();
    if (!FOUR_DIGITS_RE.test(t)) return;
    const groups = foldText(t)
      .split(/[^\p{L}\p{M}\p{N}]+/u)
      .filter((w) => /^\d+$/u.test(w) && (w.length >= MIN_DIGITS || (w.length >= MIN_VARIANT && !isYearLike(w))));
    if (groups.length === 0) return;
    if (PLAIN_NUMBER_RE.test(t) || /^[-+]?\d+,\d+$/u.test(t) || numbersIn(t).length > 0 || textDates(t).length > 0) return;
    for (const w of groups) add(w, col, "word", w.length >= MIN_DIGITS ? "word" : "free", MIN_VARIANT);
  };

  for (const range of snap.ranges) {
    for (const col of range.columns) {
      const values = col.private ? [...col.values] : [];
      const own = values.length;
      if (col.aliasOriginal !== undefined) values.push(col.aliasOriginal);
      // The dates of 8-digit cells, for a column that holds dates stored that way.
      const looseKeys = col.private ? looseKeysOf(col) : new Set<string>();
      // The column's own values, then its original header: the first owner of a variant wins,
      // so a header that is also one of the values reports as a value.
      const fromValues: Source = { label: col.label, order: col.order, origin: "value" };
      const fromAlias: Source = { label: col.label, order: col.order, origin: "alias" };
      for (const [index, v] of values.entries()) {
        if (v.trim() === "") continue;
        const src = index < own ? fromValues : fromAlias;
        // The same text with accents composed (NFC) and decomposed (NFD), without its marks
        // ("Jose Munoz"; the scan also folds accents on letters), with ß as ss, with ä ö ü as
        // ae oe ue ("Mueller"), and with compatibility characters written out (NFKC: the
        // ligature in "Griﬃths").
        const plain = withoutMarks(v);
        // eslint-disable-next-line no-control-regex
        const ascii = /^[\x00-\x7f]*$/.test(v);
        const forms = ascii
          ? [v]
          : unique([
              v,
              v.normalize("NFC"),
              v.normalize("NFD"),
              plain,
              withoutSharpS(v),
              withoutSharpS(plain),
              v.normalize("NFKC"),
              withoutUmlauts(v),
              withoutSharpS(withoutUmlauts(v)),
            ]);
        const whole = wholeValueRule(v.trim());
        for (const form of forms) {
          addForm(form, src, whole);
          addJoined(form, src);
        }
        addDigitGroups(v, src);
        addDigitForms(v, src);
        // The digits between letters: a phone number without its extension ("(831) 555-0199"
        // of "(831) 555-0199 x204"), an account number without the rest of its IBAN.
        if (HAS_DIGIT_RE.test(v)) for (const part of letterSeparatedParts(v)) addDigitForms(part, src);
        // An email address's local part, when it is distinctive ("sunflower77").
        const local = v.includes("@") ? emailLocalPart(v) : null;
        if (local !== null) add(local, src, "word", "word", MIN_VARIANT);
        // A code of 6+ letters and digits with some or other separators: "AB 1234CD", "AB.1234.CD"
        // or "xAB1234CD" for "AB-1234-CD" (see codeLetters). A shorter one, and the original header
        // of an aliased column ("FalconX92024" for "Falcon X9-2024"), typed without its separators,
        // on word boundaries: "A1030" for "A-1030".
        const letters = index < own ? codeLetters(foldText(v)) : null;
        if (letters !== null) addCode(letters, src);
        else {
          const code = codeWithoutSeparators(v);
          if (code !== null) add(code, src, "normalized", "word", MIN_VARIANT);
        }
      }
      if (!col.private) continue;
      for (const group of col.groups) if (!isYearLike(group)) add(group, fromValues, "word", "free", MIN_VARIANT);
      // Numbers as the user reads them, plain and grouped ("83,500" for "$83,500"). Written
      // plainly with fewer than 7 digits before any decimals, they follow the "plain" rule.
      const numbers = col.numbers ?? col.values.flatMap(numbersIn);
      for (const plain of unique(numbers)) {
        const short = (plain.split(".")[0] ?? "").length < MIN_DIGITS;
        const bare = new Set(short ? plainForms(plain) : []);
        for (const form of numberForms(plain)) add(form, fromValues, "normalized", bare.has(form) ? "plainText" : "number", MIN_VARIANT);
      }
      // Dates: any date phrase in the body that means one of them ("3/15/85", "15-Mar-1985"); the
      // dates of 8-digit cells only in free text. A date the column holds both ways counts anywhere.
      const addDate = (key: string, rule: MatchRule) => {
        const list = dateOwners.get(key);
        const owner: Owner = { label: col.label, order: col.order, kind: "normalized", rule, key: `${col.order}|normalized|value`, origin: "value" };
        if (list === undefined) dateOwners.set(key, [owner]);
        else if (!list.some((o) => o.order === col.order)) list.push(owner);
      };
      for (const key of dateKeysOf(col.dates ?? col.values.flatMap((v) => textDates(v).map(dateKey)))) addDate(key, "word");
      for (const key of looseKeys) addDate(key, "free");
    }
  }

  const allColumns = snap.ranges.flatMap((r) => r.columns);
  const lastOrder = allColumns.length;
  for (const canary of snap.canaries) {
    if (canary.length < MIN_CANARY) continue;
    const low = foldText(canary);
    const home = allColumns.find((c) => c.private && c.values.some((v) => foldText(v).includes(low)));
    add(canary, { label: home?.label ?? "*", order: home?.order ?? lastOrder, origin: "value" }, "canary", "any", MIN_CANARY);
  }
  return { patterns, owners, digitPatterns, digitOwners, codePatterns, codeOwners, dateOwners };
}

/** The numbers in a context value when the context doesn't list them: a plain number or an amount ("$83,500"). */
function numbersIn(v: string): string[] {
  const t = v.trim();
  return /^-?\d+(?:\.\d+)?$/.test(t) ? shownNumbers(Number(t), t, true) : shownNumbers(v, v, false);
}

/** A number (plain, unsigned) written plainly, with "." or "," as its decimal mark: "83500", "5.25", "5,25". */
function plainForms(plain: string): string[] {
  return unique([plain, plain.replace(".", ",")]);
}

/**
 * A number (plain, unsigned) as people write it: plainly (see plainForms), with "," "." " " "'"
 * or "_" grouping ("83,500", "83.500", "83 500", "83'500", "83_500"), in lakh from 100,000
 * ("8,35,000"), and a whole number exactly in thousands or millions ("83.5k", "83,5 k", "1.2m";
 * see shortNumbers). Zero decimals after it are allowed when matching (see numberEndsAt and
 * noDigitNextTo).
 */
function numberForms(plain: string): string[] {
  const [int = "", frac] = plain.split(".");
  const tail = (point: string) => (frac === undefined ? "" : point + frac);
  const forms = plainForms(plain);
  if (int.length >= 4) {
    const group = (sep: string) => int.replace(/\B(?=(?:\d{3})+$)/g, sep);
    forms.push(group(",") + tail("."), group(".") + tail(","), group(" ") + tail(","), group(" ") + tail("."));
    forms.push(group("'") + tail("."), group("_") + tail("."));
  }
  if (int.length >= 6) {
    forms.push(`${int.slice(0, -3).replace(/\B(?=(?:\d{2})+$)/g, ",")},${int.slice(-3)}${tail(".")}`);
  }
  if (frac === undefined) {
    for (const [number, unit] of shortNumbers(int)) {
      for (const n of unique([number, number.replace(".", ",")])) forms.push(`${n}${unit}`, `${n} ${unit}`);
    }
  }
  return unique(forms);
}

// ---------------------------------------------------------------------------------------------
// Constant mask: fixed, code-owned text in the body. A match counts only if it touches unmasked text.

interface BodyMask {
  /** 1 where the body is fixed text Nymform writes. */
  mask: Uint8Array;
  /** 1 inside the content of a JSON string, at the innermost level (a description's own strings). */
  content: Uint8Array;
  /**
   * 1 inside free text a person typed or Nymform copied from labels: a description's task, sheet
   * and table names, headers (or aliases), notes and allowed ranges, and an earlier reply's
   * explanation, assumptions and any other text but its kind, formula and cell (the whole reply
   * when it isn't JSON). Not row cells, numbers or formulas.
   */
  free: Uint8Array;
  /**
   * 1 inside the strings of a description or of an earlier reply's JSON: text written into JSON
   * twice, where a form feed is `\\f` in the body (see isEscapeRun).
   */
  nested: Uint8Array;
  /** 1 inside a JSON number of a description: a row cell or a stat stored as a number. */
  number: Uint8Array;
}

/** Keys whose strings are free text in a description; items of an array take the array's key. */
const FREE_DESCRIPTION_KEYS = new Set(["task", "sheet", "table", "header", "note", "allowed_ranges"]);
/** Keys whose strings are not free text in a model reply: everything else is (fail closed). */
const FIXED_REPLY_KEYS = new Set(["kind", "formula", "placement", "cell", "fill_down"]);

/**
 * Marks the content of every string whose key passes `isFree` in a parsed JSON text, mapped to
 * the body as in markStrings. Items of an array take the array's key.
 */
function markFree(
  n: JNode,
  key: string | null,
  isFree: (key: string | null) => boolean,
  offset: number,
  starts: Int32Array,
  free: Uint8Array,
): void {
  switch (n.k) {
    case "str":
      if (isFree(key)) free.fill(1, starts[offset + n.s + 1]!, starts[offset + n.e - 1]!);
      return;
    case "arr":
      for (const item of n.items) markFree(item, key, isFree, offset, starts, free);
      return;
    case "obj":
      for (const { key: k, value } of n.entries) markFree(value, k.value, isFree, offset, starts, free);
      return;
    default:
      return;
  }
}

function buildMask(body: string, parsed: Record<string, unknown>, descriptions: Description[], snap: Snapshot): BodyMask {
  const tree = new SpanParser(body).parse();
  const mask = new Uint8Array(body.length).fill(1);
  const content = new Uint8Array(body.length);
  const free = new Uint8Array(body.length);
  const nested = new Uint8Array(body.length);
  const number = new Uint8Array(body.length);
  markStrings(tree, 0, null, content);
  const describedMessages = new Set(descriptions.map((d) => d.message));
  const messages = parsed.messages as { role: string; content: string }[];

  const unmaskString = (node: JStr) => mask.fill(0, node.s + 1, node.e - 1);
  const unmaskAll = (node: JNode): void => {
    switch (node.k) {
      case "str":
        unmaskString(node);
        return;
      case "num":
        mask.fill(0, node.s, node.e);
        return;
      case "arr":
        node.items.forEach(unmaskAll);
        return;
      case "obj":
        for (const { key, value } of node.entries) {
          if (!KNOWN_KEYS.has(key.value)) unmaskString(key);
          unmaskAll(value);
        }
        return;
      default:
        return;
    }
  };

  if (tree.k !== "obj") throw new Error("body is not an object");
  for (const { key, value } of tree.entries) {
    if (PARAM_KEYS.has(key.value)) continue; // validated in checkParams
    if (key.value === "model") {
      if (!(value.k === "str" && value.value === snap.model)) unmaskAll(value);
      continue;
    }
    if (key.value !== "messages" || value.k !== "arr") {
      if (!KNOWN_KEYS.has(key.value)) unmaskString(key);
      unmaskAll(value);
      continue;
    }
    value.items.forEach((item, i) => {
      if (item.k !== "obj") return unmaskAll(item);
      for (const entry of item.entries) {
        const v = entry.value;
        if (!KNOWN_KEYS.has(entry.key.value)) unmaskString(entry.key);
        if (entry.key.value === "role" && v.k === "str" && ROLES.has(v.value)) continue;
        if (entry.key.value !== "content" || v.k !== "str") {
          unmaskAll(v);
          continue;
        }
        if (v.value !== messages[i]?.content) throw new Error("span parser disagrees with JSON.parse");
        if (v.value === SYSTEM_PROMPT || v.value === CORRECTION_PROMPT) continue;
        unmaskString(v);
        if (describedMessages.has(i)) remaskDescription(body, v, { mask, content, free, nested, number }, snap);
        else if (messages[i]?.role === "assistant") remaskReply(body, v, mask, content, free, nested);
      }
    });
  }
  // Stand-in tokens are text Nymform generates (a kind and a number), so they can't carry a
  // private value; a private "Person" must not match inside PERSON_014 in the question.
  const tokens = new RegExp(TOKEN_RE.source, "g");
  for (let m = tokens.exec(body); m !== null; m = tokens.exec(body)) {
    if (snap.isToken(m[0])) mask.fill(1, m.index, m.index + m[0].length);
  }
  return { mask, content, free, nested, number };
}

/**
 * Masks the fixed parts of a nymform user turn, mapped from decoded positions back to the body.
 * Private row cells the structural check accepts as a stand-in or a range or month rendering are
 * masked too: they are the replacement Nymform wrote, so a round private value such as 90000 must
 * not match inside its own bin "90000-99999".
 */
function remaskDescription(body: string, node: JStr, marks: BodyMask, snap: Snapshot): void {
  const { mask, content, free, nested, number } = marks;
  const { text, starts } = decodeWithMap(body, node.s, node.e);
  if (text !== node.value) throw new Error("decode mismatch");
  const inner = new SpanParser(text).parse();
  const innerMask = new Uint8Array(text.length).fill(1);
  const unmaskString = (s: JStr) => innerMask.fill(0, s.s + 1, s.e - 1);
  const visit = (n: JNode, key: string | null): void => {
    switch (n.k) {
      case "str": {
        const allowed = key !== null ? ENUM_VALUES.get(key) : undefined;
        if (!allowed?.(n.value)) unmaskString(n);
        return;
      }
      case "num":
        innerMask.fill(0, n.s, n.e);
        // A number has no escapes, so it is in one piece in the body.
        number.fill(1, starts[n.s]!, starts[n.e]!);
        return;
      case "arr":
        for (const item of n.items) visit(item, null);
        return;
      case "obj":
        for (const { key: k, value } of n.entries) {
          if (!KNOWN_KEYS.has(k.value)) unmaskString(k);
          visit(value, k.value);
        }
        return;
      default:
        return;
    }
  };
  visit(inner, null);
  if (inner.k === "obj") {
    const blocks = [inner];
    const context = inner.entries.find((e) => e.key.value === "context_ranges")?.value;
    if (context?.k === "arr") for (const cr of context.items) if (cr.k === "obj") blocks.push(cr);
    for (const block of blocks) {
      maskReplacedCells(block, snap, innerMask);
      maskCounts(block, innerMask);
      maskAddress(block, text, snap, innerMask);
    }
    maskAllowedRanges(inner, text, snap, innerMask);
  }
  for (let i = 0; i < text.length; i++) {
    if (innerMask[i] === 1) mask.fill(1, starts[i]!, starts[i + 1]!);
  }
  // Only string content counts for number forms: a grouped "5,390" must not be read across the
  // "5,390" of two JSON numbers.
  content.fill(0, node.s + 1, node.e - 1);
  markStrings(inner, 0, starts, content);
  markStrings(inner, 0, starts, nested);
  markFree(inner, null, (key) => key !== null && FREE_DESCRIPTION_KEYS.has(key), 0, starts, free);
}

/**
 * Masks the row numbers and counts of one described range (header_row, first_data_row,
 * last_data_row, rows_sent, rows_total and every stats value). checkRowNumbers and checkStats
 * have checked each of them.
 */
function maskCounts(block: Extract<JNode, { k: "obj" }>, innerMask: Uint8Array): void {
  const get = (o: JNode | undefined, key: string) => (o?.k === "obj" ? o.entries.find((e) => e.key.value === key)?.value : undefined);
  const maskNumber = (n: JNode | undefined) => {
    if (n?.k === "num") innerMask.fill(1, n.s, n.e);
  };
  const range = get(block, "range");
  for (const key of ["header_row", "first_data_row", "last_data_row"]) maskNumber(get(range, key));
  maskNumber(get(block, "rows_sent"));
  maskNumber(get(block, "rows_total"));
  const columns = get(block, "columns");
  if (columns?.k !== "arr") return;
  for (const col of columns.items) {
    const stats = get(col, "stats");
    if (stats?.k === "obj") for (const { value } of stats.entries) maskNumber(value);
  }
}

// Range addresses are text the auditor can derive itself: the address of each range in the
// context. The cell reference in them ("A1:J1000") is masked when it is exactly such an address,
// so a private 1000 or a date serial 20000 doesn't match the last row of the range. The sheet
// name before a "!" is never masked; it is scanned like any other text.

/** A range address as it is written in the text: a cell reference with nothing JSON escapes. */
function isCellReference(address: string): boolean {
  return /^[$A-Za-z0-9:]+$/u.test(address) && parseLocalRange(address) !== null;
}

/** Masks a described range's "address" when the context has a range on that sheet with that address. */
function maskAddress(block: Extract<JNode, { k: "obj" }>, text: string, snap: Snapshot, innerMask: Uint8Array): void {
  const get = (o: Extract<JNode, { k: "obj" }>, key: string) => o.entries.find((e) => e.key.value === key)?.value;
  const sheet = get(block, "sheet");
  const range = get(block, "range");
  const address = range?.k === "obj" ? get(range, "address") : undefined;
  if (sheet?.k !== "str" || address?.k !== "str" || !isCellReference(address.value)) return;
  if (!snap.ranges.some((r) => r.sheet === sheet.value && r.address === address.value)) return;
  if (text.slice(address.s + 1, address.e - 1) !== address.value) return;
  innerMask.fill(1, address.s + 1, address.e - 1);
}

/** Masks the cell reference at the end of each allowed range that names a context range in full ("HR!A1:A20000"). */
function maskAllowedRanges(description: Extract<JNode, { k: "obj" }>, text: string, snap: Snapshot, innerMask: Uint8Array): void {
  const allowed = description.entries.find((e) => e.key.value === "allowed_ranges")?.value;
  if (allowed?.k !== "arr") return;
  const known = new Map(snap.ranges.map((r) => [`${quoteSheet(r.sheet)}!${r.address}`, r.address]));
  for (const item of allowed.items) {
    const address = item.k === "str" ? known.get(item.value) : undefined;
    if (item.k !== "str" || address === undefined || !isCellReference(address)) continue;
    const start = item.e - 1 - address.length;
    if (text.slice(start, item.e - 1) !== address || text[start - 1] !== "!") continue;
    innerMask.fill(1, start, item.e - 1);
  }
}

/**
 * Marks the content of every string in a parsed JSON text (not its quotes), mapped to the body:
 * text unit i came from body [starts[i], starts[i + 1]); without `starts` the text is the body.
 */
function markStrings(n: JNode, offset: number, starts: Int32Array | null, content: Uint8Array): void {
  switch (n.k) {
    case "str": {
      const at = (i: number) => (starts ? starts[offset + i]! : offset + i);
      content.fill(1, at(n.s + 1), at(n.e - 1));
      return;
    }
    case "arr":
      for (const item of n.items) markStrings(item, offset, starts, content);
      return;
    case "obj":
      for (const { key, value } of n.entries) {
        markStrings(key, offset, starts, content);
        markStrings(value, offset, starts, content);
      }
      return;
    default:
      return;
  }
}

/** Masks the private row cells of one described range that the structural check accepts. */
function maskReplacedCells(block: Extract<JNode, { k: "obj" }>, snap: Snapshot, innerMask: Uint8Array): void {
  const get = (o: Extract<JNode, { k: "obj" }>, key: string) => o.entries.find((e) => e.key.value === key)?.value;
  const sheet = get(block, "sheet");
  const range = get(block, "range");
  const rows = get(block, "rows");
  if (sheet?.k !== "str" || range?.k !== "obj" || rows?.k !== "arr") return;
  const address = get(range, "address");
  if (address?.k !== "str") return;
  const target = snap.ranges.find((r) => r.sheet === sheet.value && r.address === address.value);
  if (!target) return;
  const layout = target.columns.filter((c) => c.treatment !== "exclude");
  for (const row of rows.items) {
    if (row.k !== "arr" || row.items.length !== layout.length) continue;
    row.items.forEach((cell, i) => {
      const col = layout[i]!;
      if (col.private && cell.k === "str" && isReplaced(cell.value, col.treatment, snap)) innerMask.fill(1, cell.s, cell.e);
    });
  }
}

/**
 * Masks the fixed parts of an earlier model reply (spec §8's JSON: its key names and "kind"
 * values). The explanation, assumptions, formula and any text around the JSON stay scanned.
 */
function remaskReply(body: string, node: JStr, mask: Uint8Array, content: Uint8Array, free: Uint8Array, nested: Uint8Array): void {
  const { text, starts } = decodeWithMap(body, node.s, node.e);
  // Text that isn't the reply's JSON is free text; within the JSON, every string but its kind,
  // formula and cell (the explanation and assumptions, and any key a reply shouldn't have).
  const allFree = (): void => void free.fill(1, node.s + 1, node.e - 1);
  let inner: JNode;
  let offset = 0;
  let length = text.length;
  try {
    inner = new SpanParser(text).parse();
  } catch {
    const a = text.indexOf("{");
    const b = text.lastIndexOf("}");
    if (a < 0 || b <= a) return allFree();
    try {
      inner = new SpanParser(text.slice(a, b + 1)).parse();
    } catch {
      return allFree();
    }
    offset = a;
    length = b + 1 - a;
  }
  if (inner.k !== "obj") return allFree();
  free.fill(1, node.s + 1, starts[offset]!);
  free.fill(1, starts[offset + length]!, node.e - 1);
  markFree(inner, null, (key) => key === null || !FIXED_REPLY_KEYS.has(key), offset, starts, free);
  const innerMask = new Uint8Array(text.length);
  innerMask.fill(1, offset, offset + length);
  const unmask = (s: number, e: number) => innerMask.fill(0, offset + s, offset + e);
  const visit = (n: JNode, key: string | null): void => {
    switch (n.k) {
      case "str":
        if (!(key === "kind" && REPLY_KINDS.has(n.value))) unmask(n.s + 1, n.e - 1);
        return;
      case "num":
        unmask(n.s, n.e);
        return;
      case "arr":
        for (const item of n.items) visit(item, null);
        return;
      case "obj":
        for (const { key: k, value } of n.entries) {
          if (!REPLY_KEYS.has(k.value)) unmask(k.s + 1, k.e - 1);
          visit(value, k.value);
        }
        return;
      default:
        return;
    }
  };
  visit(inner, null);
  for (let i = 0; i < text.length; i++) {
    if (innerMask[i] === 1) mask.fill(1, starts[i]!, starts[i + 1]!);
  }
  // Inside the reply's JSON, only its strings are content; text around it stays content.
  content.fill(0, starts[offset]!, starts[offset + length]!);
  markStrings(inner, offset, starts, content);
  markStrings(inner, offset, starts, nested);
}

// ---------------------------------------------------------------------------------------------
// Where a match is, for the pane (AuditLocation.field). Runs only after a block, on a body the
// envelope check has accepted, so every message is { role, content } and every description has
// only the keys Nymform writes.

/** Key names a path may hold: the ones Nymform writes and a reply's own. Any other key ends the path, so a path never carries typed text. */
const PATH_KEYS = new Set([...KNOWN_KEYS, ...REPLY_KEYS]);

/**
 * Returns a function giving the field at a body offset: the message, its role, and inside a
 * description or a reply that is JSON (as remaskReply reads it), the keys and indexes down to the
 * value there and whether the offset is in one of its strings (AuditField.within). Each message is
 * decoded the first time an offset falls in it. Null when the offset isn't inside a message.
 */
function fieldFinder(body: string): (at: number) => AuditField | null {
  let messages: JNode[] = [];
  try {
    const tree = new SpanParser(body).parse();
    const list = tree.k === "obj" ? tree.entries.find((e) => e.key.value === "messages")?.value : undefined;
    if (list?.k === "arr") messages = list.items;
  } catch {
    // No paths; the hits still stand.
  }
  const decoded = new Map<number, { starts: Int32Array; inner: JNode; offset: number } | null>();
  const decode = (i: number, role: string, content: JStr) => {
    if (decoded.has(i)) return decoded.get(i) ?? null;
    let out: { starts: Int32Array; inner: JNode; offset: number } | null = null;
    try {
      const { text, starts } = decodeWithMap(body, content.s, content.e);
      try {
        out = { starts, inner: new SpanParser(text).parse(), offset: 0 };
      } catch {
        // A reply with text around its JSON (see remaskReply).
        const a = text.indexOf("{");
        const b = text.lastIndexOf("}");
        if (role === "assistant" && a >= 0 && b > a) out = { starts, inner: new SpanParser(text.slice(a, b + 1)).parse(), offset: a };
      }
    } catch {
      out = null;
    }
    decoded.set(i, out);
    return out;
  };
  const fieldAt = (at: number): AuditField | null => {
    const i = messages.findIndex((m) => m.s <= at && at < m.e);
    const m = messages[i];
    if (m?.k !== "obj") return null;
    const role = m.entries.find((e) => e.key.value === "role")?.value;
    const content = m.entries.find((e) => e.key.value === "content")?.value;
    if (role?.k !== "str" || !ROLES.has(role.value) || content?.k !== "str") return null;
    const field: AuditField = { message: i, role: role.value as AuditField["role"], path: [], within: "text" };
    if (at <= content.s || at >= content.e - 1) return field;
    const json = decode(i, role.value, content);
    if (json === null) return field;
    // The decoded unit the offset falls in: the last one starting at or before it.
    let lo = 0;
    let hi = json.starts.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (json.starts[mid]! <= at) lo = mid;
      else hi = mid - 1;
    }
    pathAt(json.inner, lo - json.offset, field.path);
    field.within = withinAt(json.inner, lo - json.offset);
    return field;
  };
  return (at) => {
    try {
      return fieldAt(at);
    } catch {
      // A path is only a pointer for the pane; it never changes the hits or the message.
      return null;
    }
  };
}

/**
 * What position `k` of a message's JSON is in (see AuditField.within): a string, key or value, at
 * any depth and under any key, known or not; another part of the JSON; or text around it.
 */
function withinAt(root: JNode, k: number): AuditField["within"] {
  if (k < root.s || k >= root.e) return "text";
  let n = root;
  for (;;) {
    let next: JNode | undefined;
    if (n.k === "arr") next = n.items.find((item) => item.s <= k && k < item.e);
    else if (n.k === "obj") {
      const entry = n.entries.find((e) => e.key.s <= k && k < e.value.e);
      if (entry !== undefined) next = k < entry.key.e ? entry.key : k >= entry.value.s ? entry.value : undefined;
    }
    if (next === undefined) break;
    n = next;
  }
  // Inside the quotes: s is the opening quote and e - 1 the closing one.
  return n.k === "str" && n.s < k && k < n.e - 1 ? "string" : "json";
}

/** Appends to `path` the keys and indexes from `n` down to the value holding position `k`. */
function pathAt(n: JNode, k: number, path: (string | number)[]): void {
  if (n.k === "arr") {
    const i = n.items.findIndex((item) => item.s <= k && k < item.e);
    if (i < 0) return;
    path.push(i);
    pathAt(n.items[i]!, k, path);
  } else if (n.k === "obj") {
    const entry = n.entries.find((e) => e.key.s <= k && k < e.value.e);
    if (!entry || !PATH_KEYS.has(entry.key.value)) return;
    path.push(entry.key.value);
    if (k >= entry.value.s) pathAt(entry.value, k, path);
  }
}

// ---------------------------------------------------------------------------------------------
// A JSON parser that keeps the source span of every value and key.

interface JStr {
  k: "str";
  /** Opening quote. */
  s: number;
  /** Just past the closing quote. */
  e: number;
  value: string;
}

type JNode =
  | JStr
  | { k: "num" | "lit"; s: number; e: number }
  | { k: "arr"; s: number; e: number; items: JNode[] }
  | { k: "obj"; s: number; e: number; entries: { key: JStr; value: JNode }[] };

class SpanParser {
  private i = 0;
  constructor(private readonly t: string) {}

  parse(): JNode {
    this.ws();
    const v = this.value();
    this.ws();
    if (this.i !== this.t.length) throw new Error("trailing text");
    return v;
  }

  private ws(): void {
    const t = this.t;
    while (this.i < t.length) {
      const c = t.charCodeAt(this.i);
      if (c !== 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) break;
      this.i++;
    }
  }

  private value(): JNode {
    const c = this.t[this.i];
    if (c === "{") return this.object();
    if (c === "[") return this.array();
    if (c === '"') return this.string();
    if (c === "t" || c === "f" || c === "n") return this.literal();
    if (c === "-" || (c !== undefined && c >= "0" && c <= "9")) return this.number();
    throw new Error("unexpected character");
  }

  private object(): JNode {
    const s = this.i++;
    const entries: { key: JStr; value: JNode }[] = [];
    this.ws();
    if (this.t[this.i] === "}") {
      this.i++;
      return { k: "obj", s, e: this.i, entries };
    }
    for (;;) {
      this.ws();
      if (this.t[this.i] !== '"') throw new Error("expected key");
      const key = this.string();
      this.ws();
      if (this.t[this.i++] !== ":") throw new Error("expected colon");
      this.ws();
      entries.push({ key, value: this.value() });
      this.ws();
      const c = this.t[this.i++];
      if (c === "}") return { k: "obj", s, e: this.i, entries };
      if (c !== ",") throw new Error("expected comma");
    }
  }

  private array(): JNode {
    const s = this.i++;
    const items: JNode[] = [];
    this.ws();
    if (this.t[this.i] === "]") {
      this.i++;
      return { k: "arr", s, e: this.i, items };
    }
    for (;;) {
      this.ws();
      items.push(this.value());
      this.ws();
      const c = this.t[this.i++];
      if (c === "]") return { k: "arr", s, e: this.i, items };
      if (c !== ",") throw new Error("expected comma");
    }
  }

  private string(): JStr {
    const s = this.i;
    const end = scanString(this.t, s);
    this.i = end;
    return { k: "str", s, e: end, value: decodeWithMap(this.t, s, end).text };
  }

  private literal(): JNode {
    const s = this.i;
    for (const word of ["true", "false", "null"]) {
      if (this.t.startsWith(word, s)) {
        this.i += word.length;
        return { k: "lit", s, e: this.i };
      }
    }
    throw new Error("bad literal");
  }

  private number(): JNode {
    const s = this.i;
    const m = /-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/y;
    m.lastIndex = s;
    if (!m.exec(this.t)) throw new Error("bad number");
    this.i = m.lastIndex;
    return { k: "num", s, e: this.i };
  }
}

/** Index just past the closing quote of the string starting at `s`. */
function scanString(t: string, s: number): number {
  let i = s + 1;
  for (;;) {
    const c = t.charCodeAt(i);
    if (Number.isNaN(c)) throw new Error("unterminated string");
    if (c === 0x22) return i + 1;
    if (c === 0x5c) i += t[i + 1] === "u" ? 6 : 2;
    else if (c < 0x20) throw new Error("control character in string");
    else i++;
  }
}

const SIMPLE_ESCAPES: Record<string, string> = { '"': '"', "\\": "\\", "/": "/", b: "\b", f: "\f", n: "\n", r: "\r", t: "\t" };

/**
 * Decodes the JSON string token at [s, e) and maps each decoded UTF-16 unit to its source:
 * unit i came from [starts[i], starts[i + 1]). starts[text.length] is the closing quote.
 */
function decodeWithMap(t: string, s: number, e: number): { text: string; starts: Int32Array } {
  const last = e - 1;
  if (t[s] !== '"' || t[last] !== '"') throw new Error("not a string token");
  const parts: string[] = [];
  const starts: number[] = [];
  let i = s + 1;
  while (i < last) {
    starts.push(i);
    const ch = t[i]!;
    if (ch !== "\\") {
      parts.push(ch);
      i++;
      continue;
    }
    const esc = t[i + 1];
    if (esc === "u") {
      const hex = t.slice(i + 2, i + 6);
      if (!/^[0-9a-fA-F]{4}$/.test(hex)) throw new Error("bad unicode escape");
      parts.push(String.fromCharCode(parseInt(hex, 16)));
      i += 6;
    } else if (esc !== undefined && Object.prototype.hasOwnProperty.call(SIMPLE_ESCAPES, esc)) {
      parts.push(SIMPLE_ESCAPES[esc]!);
      i += 2;
    } else {
      throw new Error("bad escape");
    }
  }
  if (i !== last) throw new Error("string token overrun");
  starts.push(last);
  return { text: parts.join(""), starts: Int32Array.from(starts) };
}

// ---------------------------------------------------------------------------------------------
// Text helpers

const lowerCache = new Map<number, number>();

/** A unit lowercased, or the unit itself when its lowercase has another length (U+0130 İ). */
function lowerOne(c: number): number {
  const lowered = String.fromCharCode(c).toLowerCase();
  return lowered.length === 1 ? lowered.charCodeAt(0) : c;
}

/**
 * Lowercases and folds one UTF-16 unit at a time, so positions line up with the body. Folds are
 * transform.ts's foldUnit, shared with substituteUserText: a letter with accents to its base letter
 * ("é" to "e", "Ά" to "α", "İ" to "i"), Greek final sigma to sigma, Turkish ı to i, look-alike
 * apostrophes to "'", the minus sign to "-", full-width ASCII to ASCII ("Ｍａｒｉａ" to "maria"),
 * Arabic-Indic, Persian and Devanagari digits to 0-9, Arabic decimal and thousands separators to
 * "." and ",", no-break and thin spaces to " ", and ł ø đ ħ ŧ to l o d h t. Body and variants go
 * through the same function.
 */
function lowerUnits(s: string): string {
  // eslint-disable-next-line no-control-regex
  if (/^[\x00-\x7f]*$/.test(s)) return s.toLowerCase().replace(/`/g, "'");
  const codes = new Uint16Array(s.length);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x80) {
      codes[i] = c >= 0x41 && c <= 0x5a ? c + 0x20 : c === 0x60 ? 0x27 : c;
      continue;
    }
    let l = lowerCache.get(c);
    if (l === undefined) {
      // Fold, lowercase, fold again: "İ" is "I" with a dot, then "i"; "Ｍ" is "M", then "m".
      l = foldUnit(lowerOne(foldUnit(lowerOne(c))));
      lowerCache.set(c, l);
    }
    codes[i] = l;
  }
  let out = "";
  for (let i = 0; i < codes.length; i += 8192) {
    out += String.fromCharCode(...codes.subarray(i, i + 8192));
  }
  return out;
}

/** What foldText drops: combining accents and the characters matching drops (see isDroppedCodePoint). */
// eslint-disable-next-line no-control-regex
const FOLD_DROP_RE = /[\u0300-\u036f\p{Default_Ignorable_Code_Point}\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f]/gu;

/**
 * lowerUnits, then combining accents (U+0300–036F) and the characters matching drops (see
 * isDroppedCodePoint: those that can't be seen, form feed and vertical tab) taken out, and tab and
 * line breaks read as a space, as foldBody reads the body. For variants, where positions don't matter.
 */
function foldText(s: string): string {
  if (/^[\x20-\x5f\x61-\x7e]*$/.test(s)) return s.toLowerCase();
  return lowerUnits(s).replace(FOLD_DROP_RE, "").replace(/[\t\n\r]/g, " ");
}

function isAsciiDigit(c: number): boolean {
  return c >= 0x30 && c <= 0x39;
}

const wordCache = new Map<number, boolean>();
/** Letters (with their combining marks) and digits. "_" is not one: "maria_lopez" holds the word "maria". */
const WORD_RE = /^[\p{L}\p{M}\p{N}]$/u;
/** Scripts written without spaces between words; for boundaries they count as non-word. */
const CJK_RE = /^[\p{scx=Han}\p{scx=Hiragana}\p{scx=Katakana}\p{scx=Hangul}]$/u;

function isWordCodePoint(cp: number): boolean {
  if (cp < 0x80) {
    return (cp >= 0x30 && cp <= 0x39) || (cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a);
  }
  let w = wordCache.get(cp);
  if (w === undefined) {
    const ch = String.fromCodePoint(cp);
    w = WORD_RE.test(ch) && !CJK_RE.test(ch);
    wordCache.set(cp, w);
  }
  return w;
}

const ESCAPE_LETTERS = new Set(["n", "r", "t", "b", "f", '"', "/"]);

/**
 * 1 where the body has a word character ([\p{L}\p{M}\p{N}], except Han, Hiragana, Katakana and
 * Hangul, which are written without spaces: "Lopezさん" ends a word at "z"), 0 elsewhere. JSON escape
 * sequences, single or double (`\n`, `\\n`, `\"`, `é`, `\\u00e9`), count as non-word, so
 * "\nMaria" still starts a word at "M".
 */
function wordCharMap(body: string): Uint8Array {
  const map = new Uint8Array(body.length);
  for (let i = 0; i < body.length; ) {
    const cp = body.codePointAt(i)!;
    const size = cp > 0xffff ? 2 : 1;
    if (isWordCodePoint(cp)) map.fill(1, i, i + size);
    i += size;
  }
  for (let i = 0; i < body.length; ) {
    if (body[i] !== "\\") {
      i++;
      continue;
    }
    let j = i;
    while (body[j] === "\\") j++;
    const c = body[j];
    if (c !== undefined && ESCAPE_LETTERS.has(c)) {
      map.fill(0, i, j + 1);
      i = j + 1;
    } else if (c === "u" && /^[0-9a-fA-F]{4}$/.test(body.slice(j + 1, j + 5))) {
      map.fill(0, i, j + 5);
      i = j + 5;
    } else {
      i = j;
    }
  }
  return map;
}

function unique(values: string[]): string[] {
  return [...new Set(values)];
}

/** Display text and raw value of every non-empty data cell in one column of a spec. */
function columnValues(spec: RangeSpec, letter: string, fallbackIndex: number): string[] {
  const data = spec.data;
  const fromLetter = columnIndex(letter) - data.columnIndex;
  const c = columnIndex(letter) >= 0 && fromLetter >= 0 ? fromLetter : fallbackIndex;
  const rows = Math.max(data.values.length, data.text.length);
  const out: string[] = [];
  for (let r = spec.hasHeaders ? 1 : 0; r < rows; r++) {
    const text = data.text[r]?.[c];
    if (typeof text === "string" && text.trim() !== "") out.push(text);
    const value = data.values[r]?.[c];
    if (typeof value === "string" && value.trim() !== "") out.push(value);
    else if (typeof value === "number" && Number.isFinite(value)) out.push(String(value));
    else if (typeof value === "boolean") out.push(String(value));
  }
  return out;
}

/**
 * The last 4 digits of every 7+ digit whole number in one column whose display is only "#"
 * characters (the column is too narrow to show it): the display would have given its digit groups,
 * such as the "0199" of a phone number shown as (831) 555-0199.
 */
function hiddenGroups(spec: RangeSpec, letter: string, fallbackIndex: number): string[] {
  const data = spec.data;
  const fromLetter = columnIndex(letter) - data.columnIndex;
  const c = columnIndex(letter) >= 0 && fromLetter >= 0 ? fromLetter : fallbackIndex;
  const rows = Math.max(data.values.length, data.text.length);
  const out = new Set<string>();
  for (let r = spec.hasHeaders ? 1 : 0; r < rows; r++) {
    const value = data.values[r]?.[c];
    if (typeof value !== "number" || !Number.isSafeInteger(value) || !isHiddenNumberText(data.text[r]?.[c])) continue;
    const digits = String(Math.abs(value));
    if (digits.length >= MIN_DIGITS) out.add(digits.slice(-4));
  }
  return [...out];
}

/**
 * The header cell's display text and raw value when it identifies someone or something (an email
 * address, a phone number, an SSN, or an ID over a column of IDs; see schema.ts): a selection made
 * without its header row sends its first data row as headers, so those cells are scanned like the
 * column's values. Bare numbers and dates are left out: a year header would block every request
 * that mentions the year.
 */
function headerData(spec: RangeSpec, offset: number): string[] {
  if (!spec.hasHeaders || !headerCellIsIdentifier(spec.data, offset)) return [];
  const text = spec.data.text[0]?.[offset] ?? "";
  const value = spec.data.values[0]?.[offset] ?? null;
  return [text, value === null ? "" : String(value)].filter((s) => s.trim() !== "");
}

// ---------------------------------------------------------------------------------------------
// Matchers

function indexOfScan(hay: string, patterns: readonly string[], onMatch: (p: number, s: number) => boolean | void): void {
  for (let p = 0; p < patterns.length; p++) {
    const pat = patterns[p]!;
    if (pat.length === 0) continue;
    for (let i = hay.indexOf(pat); i !== -1; i = hay.indexOf(pat, i + 1)) {
      if (onMatch(p, i) === true) break;
    }
  }
}

/** Aho-Corasick over UTF-16 code units. No dependency: a trie, failure links and output links. */
function ahoCorasick(hay: string, patterns: readonly string[], onMatch: (p: number, s: number) => boolean | void): void {
  const RADIX = 0x10000;
  const edges = new Map<number, number>();
  const firstChild: number[] = [-1];
  const nextSibling: number[] = [-1];
  const charOf: number[] = [0];
  const out: (number[] | null)[] = [null];

  for (let p = 0; p < patterns.length; p++) {
    const pat = patterns[p]!;
    if (pat.length === 0) continue;
    let node = 0;
    for (let i = 0; i < pat.length; i++) {
      const c = pat.charCodeAt(i);
      let next = edges.get(node * RADIX + c);
      if (next === undefined) {
        next = firstChild.length;
        edges.set(node * RADIX + c, next);
        firstChild.push(-1);
        nextSibling.push(firstChild[node]!);
        firstChild[node] = next;
        charOf.push(c);
        out.push(null);
      }
      node = next;
    }
    (out[node] ??= []).push(p);
  }

  const count = firstChild.length;
  const fail = new Int32Array(count);
  // Nearest proper suffix node that ends a pattern, or -1.
  const dict = new Int32Array(count).fill(-1);
  const queue = new Int32Array(count);
  let head = 0;
  let tail = 0;
  for (let v = firstChild[0]!; v !== -1; v = nextSibling[v]!) queue[tail++] = v;
  while (head < tail) {
    const u = queue[head++]!;
    for (let v = firstChild[u]!; v !== -1; v = nextSibling[v]!) {
      const c = charOf[v]!;
      let f = fail[u]!;
      let target = edges.get(f * RADIX + c);
      while (target === undefined && f !== 0) {
        f = fail[f]!;
        target = edges.get(f * RADIX + c);
      }
      const fv = target ?? 0;
      fail[v] = fv;
      dict[v] = out[fv] ? fv : dict[fv]!;
      queue[tail++] = v;
    }
  }

  let node = 0;
  for (let i = 0; i < hay.length; i++) {
    const c = hay.charCodeAt(i);
    let next = edges.get(node * RADIX + c);
    while (next === undefined && node !== 0) {
      node = fail[node]!;
      next = edges.get(node * RADIX + c);
    }
    node = next ?? 0;
    for (let t = out[node] ? node : dict[node]!; t !== -1; t = dict[t]!) {
      for (const p of out[t]!) onMatch(p, i - patterns[p]!.length + 1);
    }
  }
}
