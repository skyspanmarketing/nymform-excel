// Stand-ins and row transforms (spec §7.4). The stand-in map lives in memory only (invariant 6).
import { cellKind, columnOffset, isDateFormat, isEmptyCell, rowCount } from "./schema";
import { headerWords, isEmailText, isIdShaped, isPhoneText, isSsnText } from "./suggest";
import type { CellValue, ColumnInfo, ColumnPolicy, RangeData, RangeSpec, TokenKind, Treatment } from "./types";

/** Matches any stand-in token, with or without the collision prefix. Global: reset lastIndex. */
export const TOKEN_RE = /\b(?:NYMFORM_)?(?:PERSON|EMAIL|ID|TEXT)_\d{3,}\b/g;
/** Collision guard pattern from the spec, run over the selection's display text at session start. */
export const COLLISION_RE = /\b(PERSON|EMAIL|ID|TEXT)_\d+\b/;
/** A `range` treatment rendering, e.g. "80000-89999", "0-9", "-89999--80000", or "0". */
export const RANGE_RE = /^(?:0|-?\d+--?\d+)$/;
/** A `month` treatment rendering, e.g. "2026-03". */
export const MONTH_RE = /^\d{4}-(?:0[1-9]|1[0-2])$/;

export const DEFAULT_ROW_CAP = 50;
export const MAX_ROW_CAP = 200;

const TOKEN_KINDS: readonly TokenKind[] = ["PERSON", "EMAIL", "ID", "TEXT"];

/** Session-scoped, memory-only map between private values and tokens (invariant 6). */
export class StandInMap {
  readonly prefix: "" | "NYMFORM_";
  readonly #byValue = new Map<string, string>();
  readonly #byToken = new Map<string, string>();
  readonly #counters = new Map<TokenKind, number>();
  readonly #reserved: ReadonlySet<string>;

  /** `reserved` holds token-shaped strings already in the workbook; they are never allocated. */
  constructor(prefix: "" | "NYMFORM_" = "", reserved: Iterable<string> = []) {
    this.prefix = prefix;
    this.#reserved = new Set(reserved);
  }

  /** Token for `value`, allocated on first use and stable afterwards. Format PERSON_001 (>= 3 digits). */
  tokenFor(value: string, kind: TokenKind): string {
    const known = this.#byValue.get(value);
    if (known !== undefined) return known;
    if (!TOKEN_KINDS.includes(kind)) kind = "TEXT";
    let n = this.#counters.get(kind) ?? 0;
    let token: string;
    do {
      n++;
      token = `${this.prefix}${kind}_${String(n).padStart(3, "0")}`;
    } while (this.#reserved.has(token) || this.#byToken.has(token));
    this.#counters.set(kind, n);
    this.#byValue.set(value, token);
    this.#byToken.set(token, value);
    return token;
  }

  /** The token already allocated for `value`, without allocating one. */
  peek(value: string): string | undefined {
    return this.#byValue.get(value);
  }

  /** The original value for a token, or undefined when the token is unknown. */
  valueOf(token: string): string | undefined {
    return this.#byToken.get(token);
  }

  /** True when `s` is exactly a token this map allocated. */
  isToken(s: string): boolean {
    return this.#byToken.has(s);
  }

  /** True when `s` is a token-shaped string that is already in the ranges (text from the sheet). */
  isReserved(s: string): boolean {
    return this.#reserved.has(s);
  }

  get size(): number {
    return this.#byToken.size;
  }

  /** The map is never serialized, logged or exported. */
  toJSON(): string {
    return "[stand-in map: not exported]";
  }
}

/** Every display text, string value and string formula in the ranges. */
function* rangeStrings(ranges: readonly RangeData[]): Generator<string> {
  for (const r of ranges) {
    for (const row of r.text) for (const t of row) if (typeof t === "string" && t !== "") yield t;
    for (const row of r.values) for (const v of row) if (typeof v === "string" && v !== "") yield v;
    for (const row of r.formulas) for (const f of row) if (typeof f === "string" && f !== "") yield f;
  }
}

/** True when any display text in the ranges matches COLLISION_RE. */
export function hasTokenCollision(ranges: readonly RangeData[]): boolean {
  for (const s of rangeStrings(ranges)) if (COLLISION_RE.test(s)) return true;
  return false;
}

/** Token-shaped strings that already appear in the ranges, with or without the prefix. */
function tokensInRanges(ranges: readonly RangeData[]): Set<string> {
  const found = new Set<string>();
  const re = new RegExp(String.raw`(?:NYMFORM_)?(?:PERSON|EMAIL|ID|TEXT)_\d+`, "g");
  for (const s of rangeStrings(ranges)) for (const m of s.matchAll(re)) found.add(m[0]);
  return found;
}

/** New map for a new session; uses the NYMFORM_ prefix when hasTokenCollision is true. */
export function createStandInMap(ranges: readonly RangeData[]): StandInMap {
  return new StandInMap(hasTokenCollision(ranges) ? "NYMFORM_" : "", tokensInRanges(ranges));
}

// ---------------------------------------------------------------------------------------------
// Token kinds

const NAME_WORDS = new Set(["name", "names", "surname", "firstname", "lastname", "fullname", "forename", "fname", "lname"]);
const PERSON_WORDS = new Set([
  "customer", "client", "patient", "person", "contact", "instructor", "teacher", "professor", "employee",
  "student", "member", "owner", "manager", "applicant", "recipient", "sender", "author", "guardian",
  "parent", "spouse", "driver", "buyer", "tenant", "doctor", "physician", "nurse", "advisor", "rep",
]);
/** "name" next to one of these is not a person's name ("Product name"). */
const NOT_PERSON = new Set([
  "product", "item", "company", "business", "brand", "store", "shop", "course", "project", "file",
  "sheet", "table", "team", "department", "dept", "event", "place", "location", "city", "street",
  "school", "organization", "organisation", "org", "account", "domain", "host", "server", "app",
  "model", "category", "vendor", "supplier", "program", "plan", "campaign", "region", "country", "state",
]);
const SOLO_NAME = new Set(["first", "last", "middle"]);
const EMAIL_WORDS = new Set(["email", "emails", "mail"]);
const STRONG_ID_WORDS = new Set(["id", "ids", "ssn", "emplid", "uid", "uuid", "guid", "studentid", "employeeid"]);
const WEAK_ID_WORDS = new Set([
  "phone", "mobile", "cell", "tel", "telephone", "fax", "phonenumber", "cellphone", "number", "num", "no",
  "code", "zip", "zipcode", "postal", "postcode", "social",
]);

const VALUE_SHARE = 0.5;
const KIND_SAMPLE = 200;

/**
 * PERSON for name-like headers, EMAIL, ID, else TEXT. `texts` (the column's display texts) lets
 * values decide when the header doesn't: mostly email addresses gives EMAIL; in columns that aren't
 * plain numbers, mostly ID-shaped, phone-shaped or 9-digit values gives ID.
 */
export function tokenKindFor(column: ColumnInfo, texts?: readonly string[]): TokenKind {
  const words = headerWords(column.header);
  const has = (set: ReadonlySet<string>) => words.some((w) => set.has(w));
  const sample = (texts ?? []).filter((t) => t.trim() !== "").slice(0, KIND_SAMPLE);
  const share = (test: (t: string) => boolean) =>
    sample.length === 0 ? 0 : sample.filter(test).length / sample.length;

  if (has(EMAIL_WORDS) || share(isEmailText) > VALUE_SHARE) return "EMAIL";
  if (has(STRONG_ID_WORDS)) return "ID";
  if (has(NAME_WORDS) && !has(NOT_PERSON)) return "PERSON";
  if (has(WEAK_ID_WORDS)) return "ID";
  if (has(PERSON_WORDS) || (words.length === 1 && SOLO_NAME.has(words[0]!))) return "PERSON";
  // Plain numbers (amounts, counts) and dates are not IDs just because they are unique.
  const plain = column.type === "number" || column.type === "date";
  if (!plain && share((t) => isIdShaped(t) || isPhoneText(t) || isSsnText(t)) > VALUE_SHARE) return "ID";
  return "TEXT";
}

// ---------------------------------------------------------------------------------------------
// Renderings

/** `range` treatment: 83500 -> "80000-89999"; 0 -> "0"; negatives mirror. */
export function renderRange(x: number): string {
  if (typeof x !== "number" || !Number.isFinite(x)) throw new RangeError("renderRange needs a finite number");
  if (x === 0) return "0";
  const a = Math.abs(x);
  let lo: string;
  let hi: string;
  if (a < 10) {
    lo = "0";
    hi = "9";
  } else {
    let e = Math.floor(Math.log10(a));
    if (10 ** (e + 1) <= a) e++;
    if (10 ** e > a) e--;
    const lead = Math.min(9, Math.max(1, Math.floor(a / 10 ** e)));
    // Built as strings so large magnitudes never print in exponent form.
    lo = `${lead}${"0".repeat(e)}`;
    hi = `${lead}${"9".repeat(e)}`;
  }
  return x > 0 ? `${lo}-${hi}` : `-${hi}-${lo === "0" ? "0" : `-${lo}`}`;
}

const MONTHS: Record<string, number> = {
  jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3, apr: 4, april: 4, may: 5, jun: 6, june: 6,
  jul: 7, july: 7, aug: 8, august: 8, sep: 9, sept: 9, september: 9, oct: 10, october: 10, nov: 11,
  november: 11, dec: 12, december: 12,
};

const MAX_SERIAL = 2958465; // 9999-12-31
const EPOCH = Date.UTC(1899, 11, 30);
const DAY_MS = 86400000;

function ym(year: number, month: number): string | null {
  if (!Number.isInteger(year) || !Number.isInteger(month) || year < 1000 || year > 9999 || month < 1 || month > 12) {
    return null;
  }
  return `${year}-${String(month).padStart(2, "0")}`;
}

const DAYS_IN_MONTH = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

/** True when `day` is a day of that month (1 to 12) and year, leap years included. */
function validDay(year: number, month: number, day: number): boolean {
  if (!Number.isInteger(day) || day < 1) return false;
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const last = month === 2 && leap ? 29 : DAYS_IN_MONTH[month - 1];
  return last !== undefined && day <= last;
}

/** Month of an Excel date serial (1900 date system), or null. */
export function serialMonth(serial: number): string | null {
  if (!Number.isFinite(serial) || serial < 0 || serial >= MAX_SERIAL + 1) return null;
  const s = Math.floor(serial);
  if (s === 0) return "1900-01"; // Excel shows serial 0 as 1900-01-00
  if (s === 60) return "1900-02"; // Excel's 1900-02-29
  // The 1899-12-30 epoch is right from serial 61 on; before Excel's leap-day slot it is one day early.
  const d = new Date(EPOCH + (s < 60 ? s + 1 : s) * DAY_MS);
  return ym(d.getUTCFullYear(), d.getUTCMonth() + 1);
}

/** Month of a date written as text, or null. Ambiguous numeric forms (3/4/2026) give null. */
export function textMonth(text: string): string | null {
  const t = text.trim();
  let m = /^(\d{4})[-/.](\d{1,2})(?:[-/.](\d{1,2}))?(?:[T\s].*)?$/u.exec(t);
  if (m) {
    const [y, mo] = [Number(m[1]), Number(m[2])];
    if (m[3] !== undefined && !validDay(y, mo, Number(m[3]))) return null;
    return ym(y, mo);
  }
  m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})(?:[T\s].*)?$/u.exec(t);
  if (m) {
    const [a, b, y] = [Number(m[1]), Number(m[2]), Number(m[3])];
    if (a > 12 && b <= 12 && validDay(y, b, a)) return ym(y, b);
    if (b > 12 && a <= 12 && validDay(y, a, b)) return ym(y, a);
    if (a === b && a <= 12 && validDay(y, a, a)) return ym(y, a);
    return null;
  }
  // "March 4, 2026", "Mar 2026", "Mar-2026"
  m = /^([A-Za-z]{3,9})\.?[\s\-/]+(?:(\d{1,2})(?:st|nd|rd|th)?,?[\s\-/]+)?(\d{4})(?:[T\s,].*)?$/u.exec(t);
  if (m) {
    const mo = MONTHS[m[1]!.toLowerCase()];
    const y = Number(m[3]);
    if (mo === undefined || (m[2] !== undefined && !validDay(y, mo, Number(m[2])))) return null;
    return ym(y, mo);
  }
  // "4 March 2026", "4-Mar-2026"
  m = /^(\d{1,2})(?:st|nd|rd|th)?[\s\-/]+([A-Za-z]{3,9})\.?,?[\s\-/]+(\d{4})(?:[T\s,].*)?$/u.exec(t);
  if (m) {
    const mo = MONTHS[m[2]!.toLowerCase()];
    const y = Number(m[3]);
    if (mo === undefined || !validDay(y, mo, Number(m[1]))) return null;
    return ym(y, mo);
  }
  // Leading weekday: "Wednesday, March 4, 2026"
  m = /^(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?,?\s+(.+)$/iu.exec(t);
  if (m) return textMonth(m[1]!);
  return null;
}

/** `month` treatment: an Excel date serial (or ISO/date text) -> "YYYY-MM"; null when not a date. */
export function renderMonth(value: CellValue, text: string): string | null {
  if (typeof value === "number") return serialMonth(value);
  if (typeof value === "string" && value.trim() !== "") return textMonth(value);
  if (value === null && typeof text === "string" && text.trim() !== "") return textMonth(text);
  return null;
}

// ---------------------------------------------------------------------------------------------
// Rows

export interface TransformedRows {
  /** Data rows in column order, excluded columns omitted. */
  rows: CellValue[][];
  rowsSent: number;
  rowsTotal: number;
}

/** A private column is never sent as is (brief, decision 3). */
export function effectiveTreatment(policy: Pick<ColumnPolicy, "private" | "treatment">): Treatment {
  return policy.private && policy.treatment === "as_is" ? "stand_in" : policy.treatment;
}

/** Clamps a requested row cap to 1..MAX_ROW_CAP; anything unusable gives the default. */
export function clampRowCap(rowCap?: number): number {
  if (rowCap === undefined || !Number.isFinite(rowCap)) return DEFAULT_ROW_CAP;
  return Math.min(MAX_ROW_CAP, Math.max(1, Math.floor(rowCap)));
}

interface ColumnPlan {
  offset: number;
  info: ColumnInfo;
  policy: ColumnPolicy;
  treatment: Treatment;
}

function plans(spec: RangeSpec): ColumnPlan[] {
  return spec.columns.map((info, i) => {
    // A column without a policy is treated as private: fail closed.
    const policy: ColumnPolicy = spec.policies.find((p) => p.letter === info.letter) ?? {
      letter: info.letter,
      private: true,
      treatment: "stand_in",
    };
    return { offset: columnOffset(spec.data, info.letter, i), info, policy, treatment: effectiveTreatment(policy) };
  });
}

function firstDataRow(spec: RangeSpec): number {
  return spec.hasHeaders ? 1 : 0;
}

/** Display texts of a column's data cells (non-empty), for token kinds. */
function columnTexts(spec: RangeSpec, offset: number, limit = KIND_SAMPLE): string[] {
  const out: string[] = [];
  const d = spec.data;
  for (let r = firstDataRow(spec); r < rowCount(d) && out.length < limit; r++) {
    const key = cellKey(d.values[r]?.[offset], d.text[r]?.[offset]);
    if (key !== null) out.push(key);
  }
  return out;
}

/** The map key for a cell: its display text, or its raw value when the display is blank. */
function cellKey(value: CellValue | undefined, text: string | undefined): string | null {
  if (typeof text === "string" && text !== "") return text;
  if (value === null || value === undefined || value === "") return null;
  return String(value);
}

/** A non-empty cell of a private column: where it is, its value, and the stand-in it goes out as. */
export interface PrivateCellRef {
  /** The column's letter on the sheet. */
  letter: string;
  header: string;
  /** Row index into the range's data (0 is its first row). */
  row: number;
  /** The value as the map knows it: the display text, or the raw value when the display is blank. */
  value: string;
  /** The column's token kind. */
  kind: () => TokenKind;
  /** The value's stand-in, allocated on first call and the same one substitution uses. */
  token: () => string;
}

/** Every non-empty data cell of the private columns of `spec`, column by column. */
export function privateCellsOf(spec: RangeSpec, map: StandInMap): PrivateCellRef[] {
  const out: PrivateCellRef[] = [];
  for (const p of plans(spec)) {
    if (!p.policy.private) continue;
    let kind: TokenKind | undefined;
    const kindOf = () => (kind ??= tokenKindFor(p.info, columnTexts(spec, p.offset)));
    for (let r = firstDataRow(spec); r < rowCount(spec.data); r++) {
      const value = cellKey(spec.data.values[r]?.[p.offset] ?? null, spec.data.text[r]?.[p.offset]);
      if (value === null) continue;
      out.push({ letter: p.info.letter, header: p.info.header, row: r, value, kind: kindOf, token: () => map.tokenFor(value, kindOf()) });
    }
  }
  return out;
}

/** Substituted-mode rows (first `rowCap` data rows). Private cells become tokens / ranges / months. */
export function transformRows(spec: RangeSpec, map: StandInMap, rowCap?: number): TransformedRows {
  const d = spec.data;
  const first = firstDataRow(spec);
  const rowsTotal = Math.max(0, rowCount(d) - first);
  const rowsSent = Math.min(clampRowCap(rowCap), rowsTotal);
  const sent = plans(spec).filter((p) => p.treatment !== "exclude");
  const kinds = new Map<number, TokenKind>();
  const kindOf = (p: ColumnPlan) => {
    let k = kinds.get(p.offset);
    if (k === undefined) {
      k = tokenKindFor(p.info, columnTexts(spec, p.offset));
      kinds.set(p.offset, k);
    }
    return k;
  };

  const rows: CellValue[][] = [];
  for (let r = first; r < first + rowsSent; r++) {
    const row: CellValue[] = [];
    for (const p of sent) {
      const value = d.values[r]?.[p.offset];
      const text = d.text[r]?.[p.offset];
      const vt = d.valueTypes[r]?.[p.offset];
      const fmt = d.numberFormat[r]?.[p.offset];
      if (isEmptyCell(value, text, vt)) {
        row.push(null);
        continue;
      }
      if (!p.policy.private) {
        row.push(plainCell(value, text, cellKind(value, text, vt, fmt), p.info.type));
        continue;
      }
      row.push(privateCell(p, value, text, fmt, map, kindOf));
    }
    rows.push(row);
  }
  return { rows, rowsSent, rowsTotal };
}

function plainCell(value: CellValue | undefined, text: string | undefined, kind: string, colType: string): CellValue {
  const shown = text ?? (value === null || value === undefined ? "" : String(value));
  if (kind === "date" || colType === "date") return shown;
  if ((kind === "number" || kind === "boolean") && (typeof value === "number" || typeof value === "boolean")) {
    return value;
  }
  return shown;
}

function privateCell(
  p: ColumnPlan,
  value: CellValue | undefined,
  text: string | undefined,
  numberFormat: string | undefined,
  map: StandInMap,
  kindOf: (p: ColumnPlan) => TokenKind,
): CellValue {
  if (p.treatment === "range" && typeof value === "number" && Number.isFinite(value)) {
    return renderRange(value);
  }
  if (p.treatment === "month") {
    let month: string | null = null;
    // A number is a date serial only when its cell is formatted as a date (or the column is one).
    if (typeof value === "number") {
      if (isDateFormat(numberFormat ?? "") || p.info.type === "date") month = serialMonth(value);
    } else if (typeof value === "string") {
      month = textMonth(value);
    }
    if (month !== null) return month;
  }
  const key = cellKey(value ?? null, text);
  return key === null ? null : map.tokenFor(key, kindOf(p));
}

/**
 * Every non-empty private value per column letter, display text and raw value, across the whole
 * range (not just rows sent). Used by the auditor and by substituteUserText.
 */
export function privateValues(spec: RangeSpec): Map<string, string[]> {
  const out = new Map<string, string[]>();
  const d = spec.data;
  for (const p of plans(spec)) {
    if (!p.policy.private) continue;
    const seen = new Set<string>();
    for (let r = firstDataRow(spec); r < rowCount(d); r++) {
      const text = d.text[r]?.[p.offset];
      const value = d.values[r]?.[p.offset];
      if (typeof text === "string" && text.trim() !== "") seen.add(text);
      if (value !== null && value !== undefined && String(value).trim() !== "") seen.add(String(value));
    }
    out.set(p.info.letter, [...(out.get(p.info.letter) ?? []), ...seen]);
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Numbers as the user sees them

/** Shortest plain or grouped number form that is matched (spec §7.7: 4 characters). */
const MIN_NUMBER = 4;

/** A number read from display text, as a plain unsigned decimal ("83500", "1234.20"). */
interface ReadNumber {
  plain: string;
  decimals: number;
  percent: boolean;
}

// An integer part with thousands separators (one separator throughout: "83,500", "1.234.567",
// "83 500") or Indian lakh grouping ("8,35,000"), or plain digits; then an optional decimal part.
const SHOWN_NUMBER_RE =
  /^(?:(\d{1,3})([,.'\u00a0\u202f ])(\d{3}(?:\2\d{3})*)|(\d{1,2}(?:,\d{2})+,\d{3})|(\d+))(?:([.,])(\d+))?$/u;

/**
 * The number in a display string, or [] when there is none: currency symbols, a percent sign,
 * sign, parentheses and grouping are taken out ("$83,500" -> 83500, "5.25%" -> 5.25). "1,234"
 * and "1.234" can be read two ways; a text cell's comma (`textCell`) is read as grouping. Plain
 * digits count in text cells too: an amount imported from a CSV is often text ("83500").
 */
function readNumbers(text: string, textCell: boolean): ReadNumber[] {
  let t = foldUnits(text).trim();
  if (!/\d/u.test(t)) return [];
  const percent = /%$/u.test(t);
  t = t.replace(/%$/u, "").replace(/\p{Sc}/gu, "").trim();
  // Currency codes, units and scale letters around the number: "83 500,00 kr", "CHF 91,250",
  // "Rs. 69,000", "83.5K".
  // Only whole runs of up to 4 letters ("kr", "CHF", "EUR", "Rs."), so "Order 1234567 A" isn't read as a number.
  t = t.replace(/^[\p{L}.]{1,4}(?![\p{L}.])\s*/u, "").replace(/\s*(?<![\p{L}.])[\p{L}.]{1,4}$/u, "").trim();
  if (t.startsWith("(") && t.endsWith(")")) t = t.slice(1, -1).trim();
  t = t.replace(/^[-+]\s*/u, "");
  const m = SHOWN_NUMBER_RE.exec(t);
  if (!m) return [];
  const [, head, sep, tail, lakh, digits, , frac] = m;
  const whole = (head !== undefined ? head + tail! : (lakh ?? digits!)).replace(/\D/g, "").replace(/^0+(?=\d)/u, "");
  const read = (int: string, f: string | undefined): ReadNumber => ({
    plain: f === undefined ? int : `${int}.${f}`,
    decimals: f?.length ?? 0,
    percent,
  });
  const out = [read(whole, frac)];
  // One group and no decimals: "1.234" may be 1.234 as well as 1234.
  if (head !== undefined && tail!.length === 3 && frac === undefined && (sep === "." || (sep === "," && !textCell))) {
    out.push(read(head.replace(/^0+(?=\d)/u, ""), tail));
  }
  return out;
}

/**
 * The numbers a user reads in a cell, as plain unsigned decimals of 4+ characters: the display
 * text without currency symbols, percent signs and grouping ("$83,500" -> "83500", "5.25%" ->
 * "5.25"), and the raw value rounded to the displayed decimals (83500.4 shown as "$83,500" ->
 * "83500"). Trailing zeros are also dropped ("1,234.20" -> "1234.20" and "1234.2"). `isNumber`
 * is false for a text cell, whose text then has to read as a number or an amount ("83500",
 * "$83,500", "12.5%", "1,250").
 */
export function shownNumbers(value: CellValue | undefined, text: string | undefined, isNumber: boolean): string[] {
  const raw = isNumber && typeof value === "number" && Number.isFinite(value) ? Math.abs(value) : null;
  let shown = typeof text === "string" && text.trim() !== "" ? text : typeof value === "string" ? value : "";
  if (shown === "" && raw !== null) shown = String(raw);
  const out = new Set<string>();
  for (const r of readNumbers(shown, !isNumber)) {
    if (raw === null) {
      out.add(r.plain);
      continue;
    }
    // The display has to be the raw value rounded (a percentage shows the value times 100).
    const x = r.percent ? raw * 100 : raw;
    if (Math.abs(Number(r.plain) - x) > 0.5 * 10 ** -r.decimals + 1e-9 * Math.max(1, x)) continue;
    out.add(r.plain);
    const rounded = x.toFixed(r.decimals);
    if (/^\d+(?:\.\d+)?$/u.test(rounded)) out.add(rounded);
  }
  // The stored value itself, whatever the display format shows: a format with letters ("kr",
  // "CHF", a "K" scale) or fewer decimals must not hide the number the cell holds.
  if (raw !== null) {
    // Float noise (1234.2000000000003) is read to two decimals.
    const stored = String(raw);
    out.add(/^\d+(?:\.\d{1,4})?$/u.test(stored) ? stored : raw.toFixed(2));
    if (!Number.isInteger(raw)) out.add(raw.toFixed(0));
  }
  for (const p of [...out]) if (p.includes(".")) out.add(p.replace(/\.?0+$/u, ""));
  return [...out].filter((p) => /^\d+(?:\.\d+)?$/u.test(p) && p.length >= MIN_NUMBER);
}

/** shownNumbers for one data cell: number-formatted and amount text cells only, not dates or booleans. */
function cellNumbers(d: RangeData, r: number, offset: number): string[] {
  const value = d.values[r]?.[offset];
  const text = d.text[r]?.[offset];
  const kind = cellKind(value, text, d.valueTypes[r]?.[offset], d.numberFormat[r]?.[offset]);
  if (kind === "number") return shownNumbers(value, text, true);
  if (kind === "text") return shownNumbers(value, text, false);
  return [];
}

/**
 * The numbers a user reads in each private column (see shownNumbers), per column letter, whatever
 * the column holds: amounts, numbers stored as text, and also ZIP codes and IDs. The auditor
 * matches them only inside text and with no digit continuing them, so an ID's number forms don't
 * turn up inside other numbers. Dates and booleans give none.
 */
export function privateNumbers(spec: RangeSpec): Map<string, string[]> {
  const out = new Map<string, string[]>();
  const d = spec.data;
  for (const p of plans(spec)) {
    if (!p.policy.private) continue;
    const found = new Set<string>();
    for (let r = firstDataRow(spec); r < rowCount(d); r++) for (const n of cellNumbers(d, r, p.offset)) found.add(n);
    out.set(p.info.letter, [...(out.get(p.info.letter) ?? []), ...found]);
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Dates as the user writes them

/** A calendar date. */
export interface Ymd {
  y: number;
  m: number;
  d: number;
}

function ymd(y: number, m: number, d: number): Ymd | null {
  if (!Number.isInteger(y) || y < 1000 || y > 9999 || !Number.isInteger(m) || m < 1 || m > 12 || !validDay(y, m, d)) return null;
  return { y, m, d };
}

/** The date of an Excel date serial (1900 date system), or null (serial 0, Excel's 1900-02-29, times). */
export function serialDate(serial: number): Ymd | null {
  if (!Number.isFinite(serial) || serial < 1 || serial >= MAX_SERIAL + 1) return null;
  const s = Math.floor(serial);
  if (s === 60) return null;
  const t = new Date(EPOCH + (s < 60 ? s + 1 : s) * DAY_MS);
  return ymd(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate());
}

/** A date written as "1985-03-15": the key private dates are compared by. */
export function dateKey(date: Ymd): string {
  return `${date.y}-${String(date.m).padStart(2, "0")}-${String(date.d).padStart(2, "0")}`;
}

/** A date phrase found in a text: where it is, and every date it can mean. */
export interface DateMatch {
  start: number;
  end: number;
  dates: Ymd[];
  /**
   * True for the loosest forms, which are as often something else: 8 digits in a row ("19850315",
   * also an ID), numbers separated only by spaces ("15 03 85", also part of a phone number such as
   * "06 15 03 85 12"), and three numbers of 1 or 2 digits inside a longer run with the same
   * separator ("23.12.80" of "06.45.23.12.80"). The auditor counts these only in free text.
   */
  weak: boolean;
}

/**
 * Month names and abbreviations as foldUnit folds them (accents off) and lowercased, in English,
 * German, French, Spanish, Italian, Portuguese and Dutch: "März" is "marz", "août" is "aout".
 */
const MONTH_WORDS: ReadonlyMap<string, number> = new Map(
  [
    ["jan", "january", "januar", "janner", "janvier", "janv", "enero", "ene", "gennaio", "gen", "janeiro", "januari"],
    ["feb", "february", "februar", "fevrier", "fevr", "fev", "febrero", "febbraio", "fevereiro", "februari"],
    ["mar", "march", "marz", "maerz", "mrz", "mars", "marzo", "marco", "maart", "mrt"],
    ["apr", "april", "avril", "avr", "abril", "abr", "aprile"],
    ["may", "mai", "mayo", "maggio", "mag", "maio", "mei"],
    ["jun", "june", "juni", "juin", "junio", "giugno", "giu", "junho"],
    ["jul", "july", "juli", "juillet", "juil", "julio", "luglio", "lug", "julho"],
    ["aug", "august", "aout", "agosto", "ago", "augustus"],
    ["sep", "sept", "september", "septembre", "septiembre", "setiembre", "settembre", "set", "setembro"],
    ["oct", "october", "oktober", "okt", "octobre", "octubre", "ottobre", "ott", "outubro", "out"],
    ["nov", "november", "novembre", "noviembre", "novembro"],
    ["dec", "december", "dezember", "dez", "decembre", "diciembre", "dic", "dicembre", "dezembro"],
  ].flatMap((words, i) => words.map((w): [string, number] => [w, i + 1])),
);
/** The longest month word, so a longer run of letters is not looked up. */
const MONTH_WORD_MAX = 10;

const NUMBER_CHAR_RE = /\p{N}/u;
const LETTER_CHAR_RE = /\p{L}/u;

function isDigitCode(c: number): boolean {
  return c >= 0x30 && c <= 0x39;
}

/** A digit of any script at t[i] (false past either end). */
function numberAt(t: string, i: number): boolean {
  const c = t.charCodeAt(i);
  if (c < 0x80) return isDigitCode(c);
  return !Number.isNaN(c) && NUMBER_CHAR_RE.test(t[i]!);
}

/** A letter of any script at t[i] (false past either end). */
function letterAt(t: string, i: number): boolean {
  const c = t.charCodeAt(i);
  if (c < 0x80) return (c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a);
  return !Number.isNaN(c) && LETTER_CHAR_RE.test(t[i]!);
}

/**
 * A space of any kind between two parts of a date: tab, line breaks, vertical tab and form feed,
 * NEL (U+0085), every Unicode space (\p{Zs}), and the line and paragraph separators (U+2028,
 * U+2029). The characters that can't be seen (see isIgnorableCodePoint) are taken out before.
 */
function isSpaceCode(c: number): boolean {
  if (c < 0x80) return c === 0x20 || (c >= 0x09 && c <= 0x0d);
  return (
    c === 0x85 ||
    c === 0xa0 ||
    c === 0x1680 ||
    (c >= 0x2000 && c <= 0x200a) ||
    c === 0x2028 ||
    c === 0x2029 ||
    c === 0x202f ||
    c === 0x205f ||
    c === 0x3000
  );
}

const DEFAULT_IGNORABLE_RE = /^\p{Default_Ignorable_Code_Point}$/u;
const ignorableCache = new Map<number, boolean>();

/**
 * A character that can't be seen, which matching skips as if it weren't there: every default
 * ignorable code point (zero-width space, non-joiner and joiner, word joiner, BOM, direction marks,
 * embeddings and isolates, soft hyphen, U+180E, U+2061 to U+206F, variation selectors, tag
 * characters, Hangul fillers), and every control character but the whitespace ones (backspace,
 * DEL, the C1 controls but NEL). Tab, line breaks, vertical tab, form feed and NEL count as a space.
 */
export function isIgnorableCodePoint(cp: number): boolean {
  if (cp < 0xa0) return cp < 0x20 ? cp < 0x09 || cp > 0x0d : cp >= 0x7f && cp !== 0x85;
  if (cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff)) return false;
  let v = ignorableCache.get(cp);
  if (v === undefined) {
    v = DEFAULT_IGNORABLE_RE.test(String.fromCodePoint(cp));
    ignorableCache.set(cp, v);
  }
  return v;
}

/**
 * The number of UTF-16 units of the character at s[i] when it can't be seen (see
 * isIgnorableCodePoint): 1, or 2 for a tag character; 0 for any other character or past the end.
 */
export function ignorableLength(s: string, i: number): number {
  const c = s.charCodeAt(i);
  if (c < 0x7f) return c < 0x20 && (c < 0x09 || c > 0x0d) ? 1 : 0;
  if (Number.isNaN(c)) return 0;
  if (c >= 0xd800 && c <= 0xdbff) {
    const d = s.charCodeAt(i + 1);
    return d >= 0xdc00 && d <= 0xdfff && isIgnorableCodePoint((c - 0xd800) * 0x400 + (d - 0xdc00) + 0x10000) ? 2 : 0;
  }
  return isIgnorableCodePoint(c) ? 1 : 0;
}

/** Finds a character that can't be seen (see isIgnorableCodePoint), quickly. */
// eslint-disable-next-line no-control-regex
export const IGNORABLE_RE = /[\p{Default_Ignorable_Code_Point}\x00-\x08\x0e-\x1f\x7f-\x84\x86-\x9f]/u;
// eslint-disable-next-line no-control-regex
const IGNORABLE_ALL_RE = /[\p{Default_Ignorable_Code_Point}\x00-\x08\x0e-\x1f\x7f-\x84\x86-\x9f]/gu;

/**
 * `s` without the characters that can't be seen, and where each kept unit was: `at[k]` is the
 * index in `s` of the k-th unit of `text`, and `at[text.length]` is `s.length`. `at` is null when
 * nothing was taken out (the usual case).
 */
function stripIgnorable(s: string): { text: string; at: number[] | null } {
  if (!IGNORABLE_RE.test(s)) return { text: s, at: null };
  let text = "";
  const at: number[] = [];
  for (let i = 0; i < s.length; ) {
    const skip = ignorableLength(s, i);
    if (skip > 0) {
      i += skip;
      continue;
    }
    text += s[i]!;
    at.push(i);
    i++;
  }
  at.push(s.length);
  return { text, at };
}

/** A value without the characters that can't be seen (see isIgnorableCodePoint). */
export function withoutIgnorable(s: string): string {
  return IGNORABLE_RE.test(s) ? s.replace(IGNORABLE_ALL_RE, "") : s;
}

/**
 * A character that matching values drops: one that can't be seen (see isIgnorableCodePoint), or a
 * form feed or vertical tab, so "Wata\fnabe" and "934\f96" hold Watanabe and 93496. Reading dates,
 * a form feed and a vertical tab count as a space.
 */
export function isDroppedCodePoint(cp: number): boolean {
  return cp === 0x0b || cp === 0x0c || isIgnorableCodePoint(cp);
}

/** ignorableLength, and 1 for a form feed or vertical tab (see isDroppedCodePoint). */
export function droppedLength(s: string, i: number): number {
  const c = s.charCodeAt(i);
  return c === 0x0b || c === 0x0c ? 1 : ignorableLength(s, i);
}

// eslint-disable-next-line no-control-regex
const DROPPED_RE = /[\p{Default_Ignorable_Code_Point}\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f]/u;

/**
 * `s` without the characters matching drops (see isDroppedCodePoint), where each kept unit was
 * (`at`, as in stripIgnorable), and where something was dropped: `gap[k]` is 1 when one or more
 * characters were dropped right before the k-th unit of `text` (at k = text.length: after the
 * last one). A dropped character still counts as a boundary between the units on either side of
 * it: "Mr\u200bWatanabe" holds the word "Watanabe", and "93496\u200b94103" the number 93496, while
 * "Wata\u200bnabe" also holds "Watanabe". `at` and `gap` are null when nothing was dropped.
 */
export function stripDropped(s: string): { text: string; at: number[] | null; gap: Uint8Array | null } {
  if (!DROPPED_RE.test(s)) return { text: s, at: null, gap: null };
  const parts: string[] = [];
  const at: number[] = [];
  const marks: number[] = [];
  let dropped = false;
  for (let i = 0; i < s.length; ) {
    const skip = droppedLength(s, i);
    if (skip > 0) {
      dropped = true;
      i += skip;
      continue;
    }
    if (dropped) marks.push(at.length);
    dropped = false;
    parts.push(s[i]!);
    at.push(i);
    i++;
  }
  if (dropped) marks.push(at.length);
  at.push(s.length);
  const gap = new Uint8Array(at.length);
  for (const k of marks) gap[k] = 1;
  return { text: parts.join(""), at, gap };
}

/**
 * The mark withGapMarks puts where characters were dropped: U+0000, which is no letter, digit,
 * space or separator, and which no value holds (matching drops it from values too).
 */
const GAP_MARK = "\u0000";

/**
 * A text from stripDropped with GAP_MARK wherever `gap` says characters were dropped, so a
 * pattern's word and number boundaries hold there; `at[k]` is the index in `text` of the k-th unit
 * (for a mark, of the unit after it), and `at[marked.length]` is `text.length`.
 */
function withGapMarks(text: string, gap: Uint8Array): { text: string; at: number[] } {
  const parts: string[] = [];
  const at: number[] = [];
  for (let k = 0; k <= text.length; k++) {
    if (gap[k] === 1) {
      parts.push(GAP_MARK);
      at.push(k);
    }
    if (k < text.length) {
      parts.push(text[k]!);
      at.push(k);
    }
  }
  at.push(text.length);
  return { text: parts.join(""), at };
}

/** The code point that ends right before s[k] (a surrogate pair read whole), or -1 at the start. */
function codePointBefore(s: string, k: number): number {
  if (k <= 0) return -1;
  const low = s.charCodeAt(k - 1);
  if (low >= 0xdc00 && low <= 0xdfff && k >= 2) {
    const high = s.charCodeAt(k - 2);
    if (high >= 0xd800 && high <= 0xdbff) return (high - 0xd800) * 0x400 + (low - 0xdc00) + 0x10000;
  }
  return low;
}

/** "-" or a character that looks like it: U+2010 to U+2015, the minus sign, U+FE58, U+FE63 and U+FF0D. */
export function isDashCode(c: number): boolean {
  return c === 0x2d || (c >= 0x2010 && c <= 0x2015) || c === 0x2212 || c === 0xfe58 || c === 0xfe63 || c === 0xff0d;
}

/** "\" or a character that looks like it: U+FE68, U+29F5 and U+FF3C. */
export function isBackslashCode(c: number): boolean {
  return c === 0x5c || c === 0xfe68 || c === 0x29f5 || c === 0xff3c;
}

/** "/" or a character that looks like it: the division slash U+2215, the fraction slash U+2044, full-width U+FF0F. */
function isSlashCode(c: number): boolean {
  return c === 0x2f || c === 0x2215 || c === 0x2044 || c === 0xff0f;
}

/** A character that may stand next to a month name in a date: "_", ".", ",", "/", "\", a dash, or any space. */
function isMonthSeparatorCode(c: number): boolean {
  return c === 0x5f || c === 0x2e || c === 0x2c || isBackslashCode(c) || isSlashCode(c) || isDashCode(c) || isSpaceCode(c);
}

/** Index just past the run of isMonthSeparatorCode characters starting at i (i itself when there are none). */
function monthSeparatorEnd(t: string, i: number): number {
  while (isMonthSeparatorCode(t.charCodeAt(i))) i++;
  return i;
}

/** Index just past the ASCII digits starting at i. */
function digitsEnd(t: string, i: number): number {
  while (isDigitCode(t.charCodeAt(i))) i++;
  return i;
}

function spacesEnd(t: string, i: number): number {
  while (isSpaceCode(t.charCodeAt(i))) i++;
  return i;
}

/** Kinds of separator between the parts of a date. */
const SEP_SPACE = 1;
const SEP_SLASH = 2;
const SEP_DOT = 3;
const SEP_DASH = 4;
const SEP_COMMA = 5;
const SEP_UNDERSCORE = 6;

/**
 * A separator at t[i]: one of / \ . , _ or a dash with optional spaces around it, or spaces
 * alone. A run of "\" counts as one: the body of a request doubles them. Null when there is none.
 */
function separatorAt(t: string, i: number): { end: number; kind: number } | null {
  const j = spacesEnd(t, i);
  const c = t.charCodeAt(j);
  const kind = separatorKind(c);
  if (kind !== 0) {
    let k = j + 1;
    if (c === 0x5c) while (t.charCodeAt(k) === 0x5c) k++;
    return { end: spacesEnd(t, k), kind };
  }
  return j > i ? { end: j, kind: SEP_SPACE } : null;
}

/** True when a number and a separator (see separatorAt) come right before t[i]. */
function numberJoinedBefore(t: string, i: number): boolean {
  let k = i - 1;
  while (isSpaceCode(t.charCodeAt(k))) k--;
  if (separatorKind(t.charCodeAt(k)) === 0) return false;
  k--;
  while (isBackslashCode(t.charCodeAt(k)) || isSpaceCode(t.charCodeAt(k))) k--;
  return numberAt(t, k);
}

/** True when a separator (see separatorAt) and a number come right after t[e - 1]. */
function numberJoinedAfter(t: string, e: number): boolean {
  const s = separatorAt(t, e);
  return s !== null && s.kind !== SEP_SPACE && numberAt(t, s.end);
}

/** A separator that joins the parts of a numeric date on its own: /, "\", ".", "_" or a dash. */
function isJoiningKind(kind: number): boolean {
  return kind === SEP_SLASH || kind === SEP_DOT || kind === SEP_DASH || kind === SEP_UNDERSCORE;
}

/** The kind of a one-character separator ("/", "\" or a look-alike of either, . , _ or a dash), or 0. */
function separatorKind(c: number): number {
  if (isSlashCode(c) || isBackslashCode(c)) return SEP_SLASH;
  return c === 0x2e ? SEP_DOT : c === 0x2c ? SEP_COMMA : c === 0x5f ? SEP_UNDERSCORE : isDashCode(c) ? SEP_DASH : 0;
}

/** A date ends where no digit continues it, nor a decimal point and a digit ("3/15/1985.5" is no date). */
function dateEndsAt(t: string, i: number): boolean {
  return !numberAt(t, i) && !(t.charCodeAt(i) === 0x2e && numberAt(t, i + 1));
}

const ORDINALS = new Set(["st", "nd", "rd", "th", "er"]);

/** An ordinal's letters at t[i] ("th" of "15th", "er" of French "1er"), whatever follows them. */
function isOrdinalAt(t: string, i: number): boolean {
  return ORDINALS.has(t.slice(i, i + 2).toLowerCase());
}

/** After a day's digits: past an ordinal with no letter after it ("15th", "1st", "1er"), or where it was. */
function ordinalEnd(t: string, i: number): number {
  return isOrdinalAt(t, i) && !letterAt(t, i + 2) ? i + 2 : i;
}

/**
 * Past a joining word and the separators after it ("4th of July", "15 de marzo de 1985", "March
 * the 15th", "15_of_March"), or where it was.
 */
function connectorEnd(t: string, i: number): number {
  for (const word of ["of", "de", "del", "the"]) {
    if (t.slice(i, i + word.length).toLowerCase() === word && isMonthSeparatorCode(t.charCodeAt(i + word.length))) {
      return monthSeparatorEnd(t, i + word.length);
    }
  }
  return i;
}

/**
 * The month name after a day's digits: right after them or after separators of any kind, an
 * ordinal or "of" ("15MAR", "15_March", "15th March", "15thMarch", "15th of March", "15. März").
 * Null when there is none.
 */
function monthAfterDay(t: string, dayEnd: number): { month: number; end: number } | null {
  if (isOrdinalAt(t, dayEnd)) {
    const month = monthAt(t, connectorEnd(t, monthSeparatorEnd(t, dayEnd + 2)));
    if (month) return month;
  }
  return monthAt(t, connectorEnd(t, monthSeparatorEnd(t, dayEnd)));
}

/** The month named by the word starting at t[i] (accents folded or skipped: "März", "août"), or null. */
function monthAt(t: string, i: number): { month: number; end: number } | null {
  let word = "";
  let j = i;
  for (; j < t.length && word.length <= MONTH_WORD_MAX; j++) {
    const code = t.charCodeAt(j);
    if (isAccentMark(code)) continue;
    const c = code < 0x80 ? code : foldUnit(code);
    if (c >= 0x61 && c <= 0x7a) word += String.fromCharCode(c);
    else if (c >= 0x41 && c <= 0x5a) word += String.fromCharCode(c + 0x20);
    else break;
  }
  if (word.length < 3 || letterAt(t, j)) return null;
  const month = MONTH_WORDS.get(word);
  return month === undefined ? null : { month, end: j };
}

/** A year at t[i]: 4 digits, or 2 ("85", "'85") read both as 19xx and 20xx. */
function yearAt(t: string, i: number): { years: number[]; end: number } | null {
  const apostrophe = t.charCodeAt(i) === 0x27;
  const s = apostrophe ? i + 1 : i;
  const e = digitsEnd(t, s);
  const length = e - s;
  if (!(length === 2 || (length === 4 && !apostrophe)) || !dateEndsAt(t, e)) return null;
  const n = Number(t.slice(s, e));
  return { years: length === 4 ? [n] : [1900 + n, 2000 + n], end: e };
}

/** A year of 4 digits, or of 2 read both as 19xx and 20xx. */
function yearsOf(year: string): number[] {
  return year.length === 4 ? [Number(year)] : [1900 + Number(year), 2000 + Number(year)];
}

/** Every reading of three numbers: y-m-d when the first is a year, m-d-y and d-m-y when the last is. */
function numericReadings(a: string, b: string, c: string): (Ymd | null)[] {
  const [x, y, z] = [Number(a), Number(b), Number(c)];
  if (a.length === 4) return c.length <= 2 ? [ymd(x, y, z)] : [];
  const out: (Ymd | null)[] = [];
  if (c.length === 4 || c.length === 2) for (const year of yearsOf(c)) out.push(ymd(year, x, y), ymd(year, y, x));
  if (a.length === 2 && c.length <= 2) for (const year of yearsOf(a)) out.push(ymd(year, y, z));
  return out;
}

const YEAR_MARKS = new Set([0x5e74, 0xb144]); // 年 년
const MONTH_MARKS = new Set([0x6708, 0xc6d4]); // 月 월
const DAY_MARKS = new Set([0x65e5, 0xc77c]); // 日 일

/** "1985年3月15日", "1985년 3월 15일": a year, month and day, each followed by its sign. */
function markedDate(t: string, i: number, yearEnd: number, push: Push): void {
  let j = spacesEnd(t, yearEnd);
  if (!YEAR_MARKS.has(t.charCodeAt(j))) return;
  const parts: number[] = [Number(t.slice(i, yearEnd))];
  for (const marks of [MONTH_MARKS, DAY_MARKS]) {
    const s = spacesEnd(t, j + 1);
    const e = digitsEnd(t, s);
    if (e - s < 1 || e - s > 2) return;
    parts.push(Number(t.slice(s, e)));
    j = spacesEnd(t, e);
    if (!marks.has(t.charCodeAt(j))) return;
  }
  push(i, j + 1, [ymd(parts[0]!, parts[1]!, parts[2]!)], false);
}

type Push = (start: number, end: number, dates: (Ymd | null)[], weak: boolean) => void;

/** The date phrases starting with the number at t[i] (no digit before it). */
function datesFromNumber(t: string, i: number, compact: boolean, push: Push): void {
  const aEnd = digitsEnd(t, i);
  const aLength = aEnd - i;
  if (aLength === 8) {
    // "19850315", not inside a longer number ("1.19850315", "19850315,5").
    const before = t.charCodeAt(i - 1);
    const after = t.charCodeAt(aEnd);
    if (!compact || ((before === 0x2e || before === 0x2c) && numberAt(t, i - 2))) return;
    if (numberAt(t, aEnd) || ((after === 0x2e || after === 0x2c) && numberAt(t, aEnd + 1))) return;
    push(i, aEnd, [ymd(Number(t.slice(i, i + 4)), Number(t.slice(i + 4, i + 6)), Number(t.slice(i + 6, aEnd)))], true);
    return;
  }
  if (aLength !== 1 && aLength !== 2 && aLength !== 4) return;
  const a = t.slice(i, aEnd);
  if (aLength === 4) markedDate(t, i, aEnd, push);

  // Three numbers: "3/15/85", "1985-03-15", "15.3. 1985", "15 03 1985", and "3/15, 1985".
  const s1 = separatorAt(t, aEnd);
  if (s1 && s1.kind !== SEP_COMMA) {
    const bEnd = digitsEnd(t, s1.end);
    const bLength = bEnd - s1.end;
    const s2 = bLength === 1 || bLength === 2 ? separatorAt(t, bEnd) : null;
    if (s2) {
      const cEnd = digitsEnd(t, s2.end);
      const cLength = cEnd - s2.end;
      const yearAfterComma = s1.kind === SEP_SLASH && (s2.kind === SEP_COMMA || s2.kind === SEP_SPACE) && cLength === 4 && aLength <= 2;
      // Two different ones of / \ . - _ may separate the parts when the year has 4 digits and the
      // numbers aren't part of a longer run of them: "3_15/1985", "3_15-1985", "1985-03/15", but
      // not the "03/15-1990" of "1985/03/15-1990/04/02".
      const mixed =
        isJoiningKind(s1.kind) && isJoiningKind(s2.kind) && (aLength === 4 || cLength === 4) && !numberJoinedBefore(t, i) && !numberJoinedAfter(t, cEnd);
      if ((s2.kind === s1.kind || yearAfterComma || mixed) && cLength >= 1 && cLength !== 3 && cLength <= 4 && dateEndsAt(t, cEnd)) {
        // Numbers separated by spaces alone are weak, and so are three numbers of 1 or 2 digits
        // inside a longer run of numbers with the same separator ("23.12.80" of "06.45.23.12.80",
        // a phone number).
        const chained =
          aLength <= 2 &&
          cLength <= 2 &&
          ((separatorKind(t.charCodeAt(i - 1)) === s1.kind && numberAt(t, i - 2)) ||
            (separatorKind(t.charCodeAt(cEnd)) === s2.kind && numberAt(t, cEnd + 1)));
        push(i, cEnd, numericReadings(a, t.slice(s1.end, bEnd), t.slice(s2.end, cEnd)), s1.kind === SEP_SPACE || chained);
      }
    }
  }

  // Next to a month name, the separator may be any mix of "_", ".", ",", "/", dashes and spaces
  // of any kind, or nothing at all (see monthSeparatorEnd).
  if (aLength <= 2) {
    // A day, then a month name: "15 Mar 1985", "15-Mar-85", "15. März 1985", "4th of July 1988",
    // "15MAR1985", "15_March_1985".
    const month = monthAfterDay(t, aEnd);
    const year = month ? yearAt(t, connectorEnd(t, monthSeparatorEnd(t, month.end))) : null;
    if (month && year) push(i, year.end, year.years.map((yr) => ymd(yr, month.month, Number(a))), false);
  } else {
    // A year, then a month name and a day: "1985-Mar-15", "1985 March 15", "1985MAR15".
    const month = monthAt(t, monthSeparatorEnd(t, aEnd));
    const d = month ? monthSeparatorEnd(t, month.end) : -1;
    const dEnd = d >= 0 ? digitsEnd(t, d) : -1;
    if (month && dEnd - d >= 1 && dEnd - d <= 2 && dateEndsAt(t, dEnd)) {
      push(i, ordinalEnd(t, dEnd), [ymd(Number(a), month.month, Number(t.slice(d, dEnd)))], false);
    }
  }
}

/**
 * The date phrase starting with the month name at t[i]: "March 15, 1985", "Mar-15-85", "July 4th
 * of 1988", "March the 15th, 1985", "March15, 1985", "March_15,_1985". Between the month name and
 * the day any separators or none; between the day and the year an ordinal, separators or "of".
 */
function datesFromMonth(t: string, i: number, push: Push): void {
  const month = monthAt(t, i);
  if (!month) return;
  const d = connectorEnd(t, monthSeparatorEnd(t, month.end));
  const dEnd = digitsEnd(t, d);
  if (dEnd - d < 1 || dEnd - d > 2) return;
  const y = connectorEnd(t, monthSeparatorEnd(t, ordinalEnd(t, dEnd)));
  // "March 151985" is not read: the day and the year need something between them.
  const year = y > dEnd ? yearAt(t, y) : null;
  if (year) push(i, year.end, year.years.map((yr) => ymd(yr, month.month, Number(t.slice(d, dEnd)))), false);
}

/**
 * The date phrases in a text, each with every date it can mean. A phrase is tried at every number
 * with no digit before it and at every word with no letter before it, so dates joined without
 * spaces ("3/15/85-4/2/90", "15-Mar-85/02-Apr-90") are each found. The forms: numbers with "/"
 * (or a look-alike: "3∕2∕1990"), "-", "." or spaces in y-m-d, m-d-y or d-m-y order ("1985-03-15",
 * "3/15/85" as 1985 and 2085, "4/2/1990" as April 2 and February 4, "15.3. 1985", "1985. 3. 15.",
 * "3/15, 1985"); a day and a month name in either order, in English, German, French, Spanish,
 * Italian, Portuguese or Dutch, with any separators next to the month name or none ("15-Mar-1985",
 * "July 4th, 1988", "4th of July 1988", "March the 15th, 1985", "15. März 1985", "15 de marzo de
 * 1985", "15MAR1985", "March15, 1985", "15_March_1985"); "1985年3月15日" and "1985년 3월 15일"; and,
 * with `compact`, 8 digits in year-month-day order ("19850315"). Numbers may also be separated by
 * "_" or "\" ("1985_03_15", "3\15\1985"); alone, they need a separator between them. Only whole
 * dates: a year, or a month and a year, is not one. Phrases may overlap. Numbers may mix those
 * separators when the year has 4 digits ("3_15/1985"). Dash look-alikes count as "-" and backslash
 * look-alikes as "\". Characters that can't be seen (see isIgnorableCodePoint) are skipped:
 * "3/15/1985" copied from Windows with a direction mark before each part is read, and the phrase's
 * start and end are where it is in `text`. They are also read as a space, so they still end a date.
 */
export function findDates(text: string, compact = true): DateMatch[] {
  if (!/\d/u.test(text)) return [];
  const { text: shown, at } = stripIgnorable(text);
  const found = scanDates(shown, compact);
  if (at === null) return found;
  for (const f of found) {
    f.start = at[f.start]!;
    f.end = at[f.end - 1]! + 1;
  }
  // Such a character still separates what is on either side of it, as a space would: "3/15/1985"
  // with one and a digit after it is still the date, and "15\u200b03\u200b1985" reads as "15 03 1985".
  const seen = new Set(found.map((f) => `${f.start}:${f.end}`));
  for (const f of scanDates(text.replace(IGNORABLE_ALL_RE, (m) => " ".repeat(m.length)), compact)) {
    if (!seen.has(`${f.start}:${f.end}`)) found.push(f);
  }
  return found;
}

/** findDates on a text with no character that can't be seen. */
function scanDates(text: string, compact: boolean): DateMatch[] {
  const out: DateMatch[] = [];
  if (!/\d/u.test(text)) return out;
  const push: Push = (start, end, dates, weak) => {
    const kept = keep(dates);
    if (kept.length > 0) out.push({ start, end, dates: kept, weak });
  };
  const n = text.length;
  for (let i = 0; i < n; i++) {
    const c = text.charCodeAt(i);
    if (isDigitCode(c)) {
      if (!numberAt(text, i - 1)) datesFromNumber(text, i, compact, push);
      i = digitsEnd(text, i) - 1;
    } else if ((c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a)) {
      if (!letterAt(text, i - 1)) datesFromMonth(text, i, push);
      while (letterAt(text, i + 1) && text.charCodeAt(i + 1) < 0x80) i++;
    }
  }
  return out;
}

const WEEKDAY_RE = /^(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?,?\s+/iu;
/**
 * A time after a date ("10:30 AM", "T00:00:00.000", and ":00:00:00" as SAS writes it:
 * "15MAR1985:00:00:00"), with an optional zone ("Z", " UTC", "+00", "-05:00").
 */
const TIME_AT_END_RE = /(?:,?\s+|T|:)\d{1,2}:\d{2}(?::\d{2}(?:[.,]\d+)?)?(?:\s?[ap]\.?m\.?)?(?:\s?(?:Z|UTC|GMT)|\s?[+-]\d{2}(?::?\d{2})?)?$/iu;
/** A zone right after a date, with no time: "1982-02-18-08:00", "1985-03-15Z". */
const ZONE_AT_END_RE = /(?<=\d)(?:Z|\s?(?:UTC|GMT)|[+-]\d{2}:\d{2})$/iu;
/**
 * Characters a date written as text can have: letters, digits, spaces of any kind, "_", "\" and
 * , . / : ' (and look-alike slashes and backslashes) and dashes. Anything else (@, $, parentheses)
 * means it isn't one.
 */
const DATE_TEXT_RE = /^[\p{L}\p{N}\s\p{Zs}_,./\\\u2215\u2044\ufe68\u29f5:'\-‐-―\u2212\ufe58\ufe63]+$/u;
/** ASCII digits alone: never a date without a separator (8 of them are read by eightDigitDates). */
const DIGITS_ONLY_RE = /^\d+$/u;

/**
 * The dates a text can mean when the whole text is one date (see findDates): "1985-03-15",
 * "3/15/1985", "15.03.1985", "March 15, 1985", "15-Mar-85", "15MAR1985", "15_Mar_1985",
 * "1985_03_15", "15. März 1985", "1985年3月15日", with an optional weekday before and time or
 * zone after ("1985-03-15T00:00:00-05:00", "2001-01-09 00:00:00 UTC", "1982-02-18-08:00", SAS's
 * "15MAR1985:00:00:00"). "4/2/1990" gives both readings, and "03/15/85" both 1985 and 2085.
 * Characters that can't be seen are skipped. 8 digits alone are not read here: they are as often
 * an ID, so their dates count only in free text (see eightDigitDates).
 */
export function textDates(text: string): Ymd[] {
  if (DIGITS_ONLY_RE.test(text)) return [];
  let t = withoutIgnorable(foldUnits(text)).trim();
  if (t.length < 6 || t.length > 64 || !/\d/u.test(t) || DIGITS_ONLY_RE.test(t)) return [];
  t = t.replace(WEEKDAY_RE, "").replace(TIME_AT_END_RE, "").replace(ZONE_AT_END_RE, "");
  if (!DATE_TEXT_RE.test(t)) return [];
  // A final period after the day is allowed ("1985. 3. 15.").
  const last = t.endsWith(".") ? t.length - 1 : t.length;
  const whole = findDates(t, false).filter((f) => f.start === 0 && (f.end === t.length || f.end === last));
  return keep(whole.flatMap((f) => f.dates));
}

/**
 * Every date 8 digits can mean with a year from 1900 to 2100: year-month-day ("19850315"),
 * month-day-year ("03151985") and day-month-year ("15031985"). Empty for other text. A column
 * of such values may hold dates of birth imported from a CSV, or IDs and lot numbers; they are
 * read as dates only for a column where at least half of them are dates, placeholders such as
 * "00000000" left out (see columnLooseDates), and those dates count only in free text (the
 * question, notes, headers, names, allowed ranges, earlier replies).
 */
export function eightDigitDates(text: string): Ymd[] {
  const t = eightDigitText(text);
  return t === null ? [] : eightDigitReadings(t);
}

/** The 8 ASCII digits a text is, spaces and characters that can't be seen around them taken off; null for other text. */
function eightDigitText(text: string): string | null {
  if (isEightDigits(text)) return text;
  if (text.length < 8) return null;
  // Most texts are not 8 digits: give up at the first character that is none of a digit, a space
  // or a character that can't be seen.
  for (let i = 0; i < text.length; ) {
    const c = text.charCodeAt(i);
    const skip = isDigitCode(c) || isSpaceCode(c) ? 1 : ignorableLength(text, i);
    if (skip === 0 && !numberAt(text, i)) return null;
    i += Math.max(1, skip);
  }
  const t = withoutIgnorable(foldUnits(text)).trim();
  return isEightDigits(t) ? t : null;
}

/** The dates of 8 ASCII digits (see eightDigitDates). */
function eightDigitReadings(t: string): Ymd[] {
  const a = pairAt(t, 0);
  const b = pairAt(t, 2);
  const c = pairAt(t, 4);
  const d = pairAt(t, 6);
  const first = a * 100 + b;
  const last = c * 100 + d;
  const ymdDate = first >= 1900 && first <= 2100 ? ymd(first, c, d) : null;
  if (last < 1900 || last > 2100) return ymdDate ? [ymdDate] : [];
  return keep([ymdDate, ymd(last, a, b), ymd(last, b, a)]);
}

/** The 2-digit number at t[k] (ASCII digits). */
function pairAt(t: string, k: number): number {
  return (t.charCodeAt(k) - 0x30) * 10 + t.charCodeAt(k + 1) - 0x30;
}

/** Exactly 8 ASCII digits. */
function isEightDigits(t: string): boolean {
  if (t.length !== 8) return false;
  for (let i = 0; i < 8; i++) if (!isDigitCode(t.charCodeAt(i))) return false;
  return true;
}

function keep(dates: (Ymd | null)[]): Ymd[] {
  const out: Ymd[] = [];
  for (const d of dates) if (d && !out.some((o) => o.y === d.y && o.m === d.m && o.d === d.d)) out.push(d);
  return out;
}

/** Display text that is only "#" characters: Excel's sign that the column is too narrow for the number. */
export function isHiddenNumberText(text: string | undefined): boolean {
  return typeof text === "string" && /^#+$/u.test(text.trim());
}

/**
 * The dates of one data cell. A number formatted as a date (or in a date column) is a serial,
 * read in Excel's default 1900 date system; its display text is also read, because in a workbook
 * that uses the 1904 date system the serial names a date four years off. Text is read with textDates.
 */
function cellDates(d: RangeData, r: number, offset: number, colType: string): Ymd[] {
  const value = d.values[r]?.[offset];
  const text = d.text[r]?.[offset];
  if (typeof value === "number") {
    const fmt = d.numberFormat[r]?.[offset] ?? "";
    if (!isDateFormat(fmt) && colType !== "date") return [];
    const date = serialDate(value);
    const shown = typeof text === "string" && text.trim() !== "" && !isHiddenNumberText(text) ? textDates(text) : [];
    // When the display names the serial's date, that date alone: a display such as "4/2/1990"
    // also reads as February 4, which the cell doesn't hold.
    if (date && (shown.length === 0 || shown.some((s) => s.y === date.y && s.m === date.m && s.d === date.d))) return [date];
    return keep([date, ...shown]);
  }
  const out: Ymd[] = [];
  for (const t of [text, typeof value === "string" && value !== text ? value : undefined]) {
    if (typeof t === "string" && t.trim() !== "") out.push(...textDates(t));
  }
  return keep(out);
}

/**
 * The 8 ASCII digits a data cell holds (see eightDigitDates): text such as "19850315" or
 * "03151985", or a whole number of 8 digits without a date format, as Excel stores "19850315" from
 * a CSV. The display text and a raw value that differs from it each give theirs; null for a cell
 * that doesn't hold 8 digits.
 */
function cellEightDigits(d: RangeData, r: number, offset: number): string[] | null {
  const value = d.values[r]?.[offset];
  if (typeof value === "number") {
    if (!Number.isInteger(value) || value < 10000000 || value > 99999999) return null;
    if (isDateFormat(d.numberFormat[r]?.[offset] ?? "")) return null;
    return [String(value)];
  }
  const text = d.text[r]?.[offset];
  const shown = typeof text === "string" ? eightDigitText(text) : null;
  const raw = typeof value === "string" && value !== text ? eightDigitText(value) : null;
  if (shown === null && raw === null) return null;
  return [...(shown === null ? [] : [shown]), ...(raw === null ? [] : [raw])];
}

/**
 * 8 digits that stand for "no date" rather than for a date or an ID: one digit repeated
 * ("00000000", "99999999"), or a year from 1900 to 2100 with month or day 00 ("19000000",
 * "19850300", "00001985"). They don't count when deciding whether a column's 8-digit values are
 * dates (see columnLooseDates).
 */
function isDatePlaceholder(t: string): boolean {
  if (/^(\d)\1{7}$/u.test(t)) return true;
  const first = pairAt(t, 0) * 100 + pairAt(t, 2);
  const last = pairAt(t, 4) * 100 + pairAt(t, 6);
  if (first >= 1900 && first <= 2100 && (pairAt(t, 4) === 0 || pairAt(t, 6) === 0)) return true;
  return last >= 1900 && last <= 2100 && (pairAt(t, 0) === 0 || pairAt(t, 2) === 0);
}

/**
 * The share of a column's 8-digit values, placeholders left out (see isDatePlaceholder), that must
 * be dates for the column to be read as dates stored as 8 digits. Invoice numbers such as 20240001,
 * 20240002 ... are dates only now and then; a column of dates of birth may have a few "00000000"
 * or "19000000" for unknown ones.
 */
const LOOSE_DATE_SHARE = 0.5;

/** Counts 8-digit values for LOOSE_DATE_SHARE: `eight` values that aren't placeholders, `valid` of them dates. */
interface LooseCount {
  eight: number;
  valid: number;
}

/** Adds one value's 8 digits (one text or more) to `count`, and returns their dates. */
function countLoose(texts: readonly string[], count: LooseCount): Ymd[] {
  const dates = keep(texts.flatMap(eightDigitReadings));
  if (dates.length > 0) {
    count.eight++;
    count.valid++;
  } else if (!texts.every(isDatePlaceholder)) {
    count.eight++;
  }
  return dates;
}

/** True when `count` says the values are dates stored as 8 digits (see LOOSE_DATE_SHARE). */
function areLooseDates(count: LooseCount): boolean {
  return count.eight > 0 && count.valid >= LOOSE_DATE_SHARE * count.eight;
}

/**
 * The dates of one column's 8-digit cells (see cellEightDigits), by row, when the column holds
 * dates stored as 8 digits: at least half of its non-empty 8-digit values that aren't placeholders
 * are a date. Null for any other column: its 8-digit values are not read as dates. Their digits
 * are still replaced wherever they are typed, as a date too ("2024-01-15" for invoice number
 * 20240115).
 */
function columnLooseDates(d: RangeData, first: number, offset: number): Map<number, Ymd[]> | null {
  const count: LooseCount = { eight: 0, valid: 0 };
  const byRow = new Map<number, Ymd[]>();
  for (let r = first; r < rowCount(d); r++) {
    const texts = cellEightDigits(d, r, offset);
    if (texts === null) continue;
    const dates = countLoose(texts, count);
    if (dates.length > 0) byRow.set(r, dates);
  }
  return areLooseDates(count) ? byRow : null;
}

/**
 * The dates of a list of texts that are 8 digits (see eightDigitDates), every reading, when at
 * least half of them that aren't placeholders are a date (see columnLooseDates); otherwise none.
 */
export function looseDatesOf(texts: readonly string[]): Ymd[] {
  const count: LooseCount = { eight: 0, valid: 0 };
  const out = new Map<string, Ymd>();
  for (const text of texts) {
    const t = eightDigitText(text);
    if (t === null) continue;
    for (const date of countLoose([t], count)) out.set(dateKey(date), date);
  }
  return areLooseDates(count) ? [...out.values()] : [];
}

/**
 * The dates of every private column's 8-digit cells (see columnLooseDates), per column letter, as
 * ISO dates, every reading; none for a column that doesn't hold dates stored as 8 digits. The
 * auditor counts them only in free text.
 */
export function privateLooseDates(spec: RangeSpec): Map<string, string[]> {
  const out = new Map<string, string[]>();
  const d = spec.data;
  for (const p of plans(spec)) {
    if (!p.policy.private) continue;
    const found = new Set<string>();
    for (const dates of columnLooseDates(d, firstDataRow(spec), p.offset)?.values() ?? []) for (const date of dates) found.add(dateKey(date));
    out.set(p.info.letter, [...(out.get(p.info.letter) ?? []), ...found]);
  }
  return out;
}

/**
 * The dates of every private column, per column letter, as ISO dates ("1985-03-15"; see dateKey):
 * every date a cell can mean, so a 2-digit year gives both 19xx and 20xx and "4/2/1990" both
 * April 2 and February 4.
 */
export function privateDates(spec: RangeSpec): Map<string, string[]> {
  const out = new Map<string, string[]>();
  const d = spec.data;
  for (const p of plans(spec)) {
    if (!p.policy.private) continue;
    const found = new Set<string>();
    for (let r = firstDataRow(spec); r < rowCount(d); r++) for (const date of cellDates(d, r, p.offset, p.info.type)) found.add(dateKey(date));
    out.set(p.info.letter, [...(out.get(p.info.letter) ?? []), ...found]);
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// The user's question

interface Candidate {
  /** Folded like the text it is matched in (see foldUnits and foldLower). */
  needle: string;
  replacement: () => string;
  /** Matched in the lowercased question: headers, and values of 4+ characters. */
  caseless: boolean;
  /** A number matched in another format; `needle` is then the number written plainly ("83500"). */
  pattern?: () => RegExp;
  /** Where it is in the text it is matched in, found some other way (a code; see codeLetters). */
  spans?: () => { index: number; length: number }[];
}

/**
 * The letters and digits of a folded text with the separators between them taken out (see
 * isCodeSeparatorAt), and "|" wherever anything else splits them; `at[k]` is the index in `text`
 * of the k-th unit.
 */
function codeProjection(text: string): { letters: string; at: number[] } {
  const out: string[] = [];
  const at: number[] = [];
  let gap = false;
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    const size = cp > 0xffff ? 2 : 1;
    if (isCodeLetterCodePoint(cp)) {
      if (gap && out.length > 0) {
        out.push("|");
        at.push(i);
      }
      gap = false;
      for (let k = 0; k < size; k++) {
        out.push(text[i + k]!);
        at.push(i + k);
      }
    } else if (!isCodeSeparatorAt(text, i)) {
      gap = true;
    }
    i += size;
  }
  return { letters: out.join(""), at };
}

/** True when a run of letters and digits (see codeProjection) mixes both and has `min` or more. */
function hasCodeRun(letters: string, min: number): boolean {
  for (const run of letters.split("|")) if (run.length >= min && /\p{L}/u.test(run) && /\p{N}/u.test(run)) return true;
  return false;
}

/** Letters and digits. "_" is not one: "maria_lopez" holds the word "maria". */
const WORD_CHAR = /[\p{L}\p{N}]/u;
/** Scripts written without spaces between words: no whole-word guard next to them. */
const CJK = String.raw`\p{scx=Han}\p{scx=Hiragana}\p{scx=Katakana}\p{scx=Hangul}`;
const CJK_CHAR = new RegExp(`[${CJK}]`, "u");
const LATIN_WORD = String.raw`(?:(?![${CJK}])[\p{L}\p{N}])`;
const CASE_INSENSITIVE_FROM = 4;
const MIN_DIGITS = 7;
/** A number of this many digits or more that starts with 0 is also matched without it (a trunk prefix). */
const MIN_TRUNK_DIGITS = 9;
/** Characters a typed number may have between its digits (the auditor drops the same ones). */
const DIGIT_SEPARATORS = String.raw`[\s().+\-_/\u2010-\u2015\u2212\ufe58\ufe63\uff08\uff09\uff0b\uff0d\uff0e\uff0f]*`;
/** Where a typed value may have spaces the stored one hasn't, or "_" for them ("maria_lopez"). */
const SPACES = String.raw`[\s\u3000_]*`;

/**
 * One-for-one folds (the auditor applies the same ones): look-alike apostrophes (’ ‘ ʼ ′ ʹ ´ ` ＇)
 * to "'", dashes and the minus sign (U+2010 to U+2015, U+2212, U+FE58, U+FE63) to "-", Arabic
 * decimal and thousands separators to "." and ",", every other space (no-break, thin, ideographic,
 * the line and paragraph separators, NEL) to " ", Turkish ı to i, final sigma to sigma, and
 * letters with a stroke (ł ø đ ħ ŧ) to their base letter. Full-width ASCII, and Arabic-Indic,
 * Persian and Devanagari digits, are folded by range in foldUnit.
 */
const UNIT_FOLD = new Map<number, number>([
  [0x2019, 0x27], [0x2018, 0x27], [0x02bc, 0x27], [0x2032, 0x27], [0x02b9, 0x27], [0x00b4, 0x27], [0x0060, 0x27],
  [0x2010, 0x2d], [0x2011, 0x2d], [0x2012, 0x2d], [0x2013, 0x2d], [0x2014, 0x2d], [0x2015, 0x2d], [0x2212, 0x2d],
  [0xfe58, 0x2d], [0xfe63, 0x2d],
  [0x066b, 0x2e], [0x066c, 0x2c],
  [0x0085, 0x20], [0x00a0, 0x20], [0x1680, 0x20], [0x2000, 0x20], [0x2001, 0x20], [0x2002, 0x20], [0x2003, 0x20],
  [0x2004, 0x20], [0x2005, 0x20], [0x2006, 0x20], [0x2007, 0x20], [0x2008, 0x20], [0x2009, 0x20], [0x200a, 0x20],
  [0x2028, 0x20], [0x2029, 0x20], [0x202f, 0x20], [0x205f, 0x20], [0x3000, 0x20],
  [0x0131, 0x69], [0x03c2, 0x03c3],
  [0x0142, 0x6c], [0x0141, 0x4c], [0x00f8, 0x6f], [0x00d8, 0x4f], [0x0111, 0x64], [0x0110, 0x44],
  [0x0127, 0x68], [0x0126, 0x48], [0x0167, 0x74], [0x0166, 0x54],
]);

const unitFoldCache = new Map<number, number>();

/**
 * One UTF-16 unit folded for matching, case kept: UNIT_FOLD; full-width ASCII (U+FF01–FF5E) to
 * ASCII ("Ｍａｒｉａ" to "Maria"); Arabic-Indic, Persian and Devanagari digits to 0-9; and a letter
 * with accents to its base letter, when its decomposition is one letter and accents from U+0300 to
 * U+036F ("é" to "e", "Ά" to "Α", "İ" to "I", "й" to "и"). The auditor and substituteUserText both
 * use it, so a value and the text it is looked for in fold alike.
 */
export function foldUnit(c: number): number {
  if (c < 0x60) return c;
  let f = unitFoldCache.get(c);
  if (f !== undefined) return f;
  f = UNIT_FOLD.get(c);
  if (f === undefined) {
    if (c >= 0xff01 && c <= 0xff5e) f = c - 0xfee0;
    else if (c >= 0x0660 && c <= 0x0669) f = c - 0x0660 + 0x30;
    else if (c >= 0x06f0 && c <= 0x06f9) f = c - 0x06f0 + 0x30;
    else if (c >= 0x0966 && c <= 0x096f) f = c - 0x0966 + 0x30;
    else f = accentBase(c);
  }
  unitFoldCache.set(c, f);
  return f;
}

/** The base letter of a letter with accents (U+0300–036F), or the unit itself. */
function accentBase(c: number): number {
  if (c < 0xc0 || (c >= 0xd800 && c <= 0xdfff)) return c;
  const d = String.fromCharCode(c).normalize("NFD");
  if (d.length < 2) return c;
  for (let i = 1; i < d.length; i++) if (!isAccentMark(d.charCodeAt(i))) return c;
  return d.charCodeAt(0);
}

/** Combining accents (U+0300–036F): dropped when matching, as the auditor does. */
export function isAccentMark(c: number): boolean {
  return c >= 0x0300 && c <= 0x036f;
}

// A unit foldUnit changes: anything from U+0060 on that isn't plain ASCII, and the backtick.
// eslint-disable-next-line no-control-regex
const PLAIN_RE = /^[\x00-\x5f\x61-\x7f]*$/u;

/** foldUnit on every unit of `s`. Positions don't change. */
function foldUnits(s: string): string {
  if (PLAIN_RE.test(s)) return s;
  let out = "";
  let last = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    const f = foldUnit(c);
    if (f === c) continue;
    out += s.slice(last, i) + String.fromCharCode(f);
    last = i + 1;
  }
  return last === 0 ? s : out + s.slice(last);
}

/**
 * For a number stored in international form ("+18315550199", "+44 20 7946 0958", "0044 ..."), its
 * digits without a country code of 1 to 3 digits, 7 digits or more each: one of them is the
 * number as dialled at home ("8315550199", "2079460958"). Empty for other values.
 */
export function withoutCountryCode(v: string): string[] {
  const t = foldUnits(v).trim();
  const m = /^(?:\+|00)\s*\(?\d/u.exec(t);
  if (!m) return [];
  const digits = t.replace(/\D/g, "").replace(/^00/u, "");
  if (digits.length < 10 || digits.length > 15) return [];
  const out: string[] = [];
  for (let k = 1; k <= 3; k++) if (digits.length - k >= MIN_DIGITS) out.push(digits.slice(k));
  return out;
}

/** A value with every combining mark taken off (NFD, then \p{Mn} dropped): "José Muñoz" as "Jose Munoz". */
export function withoutMarks(v: string): string {
  // eslint-disable-next-line no-control-regex
  if (/^[\x00-\x7f]*$/.test(v)) return v;
  return v.normalize("NFD").replace(/\p{Mn}/gu, "").normalize("NFC");
}

/**
 * True when a "Header sent" alias still shows the column's original header: the same text in
 * another case, spacing, width or accents ("Order Date" or "Order date " for "Order date"), or
 * with the header's words in it whole ("Order date (month)"). The header then goes out anyway, in
 * the alias, so it isn't hidden: the question isn't rewritten for it, and the check doesn't look for
 * it, as it would find it in the alias Nymform sends.
 *
 * It folds what the check's own scan folds, so an alias it calls showing is one the scan would
 * find the header in: case, width, look-alike letters, digits and punctuation ("Lodz" shows
 * "Łódź", "Q1–Q2" shows "Q1-Q2"), U+0300–036F accents and characters that can't be seen. It also
 * reads text as NFC and lowercases astral letters. A Thai tone mark, an Indic vowel sign or a
 * dakuten makes another word ("ピザ番号" doesn't show "ビザ番号"), and only whole words count,
 * with a mark part of its word: "Tracer" doesn't show "Race", nor "बीमारी" "बीमार".
 */
export function aliasRevealsHeader(alias: string, header: string): boolean {
  // The same text typed again shows it, even when the fold leaves nothing to compare.
  if (alias.trim() === header.trim()) return true;
  const a = headerKey(alias);
  const h = headerKey(header);
  if (a === "" || h === "") return false;
  for (let i = a.indexOf(h); i !== -1; i = a.indexOf(h, i + 1)) {
    if (!nameCharBefore(a, i) && !nameCharAt(a, i + h.length)) return true;
  }
  return false;
}

/**
 * A header or alias as aliasRevealsHeader compares them: case, width and look-alikes folded, U+0300–036F
 * accents and characters that can't be seen dropped (foldUnits, foldLower, as the question), spacing
 * collapsed. No other marks are dropped.
 */
function headerKey(s: string): string {
  return foldLower(withoutIgnorable(foldUnits(s.normalize("NFC"))))
    .text.replace(/\s+/gu, " ")
    .trim();
}

/** Letters with their marks, and digits: what continues a word, for a column's name (the auditor's word rule). */
const NAME_CHAR = /[\p{L}\p{M}\p{N}]/u;

/** True when the code point at `i` continues a word (see NAME_CHAR). */
export function nameCharAt(s: string, i: number): boolean {
  const cp = s.codePointAt(i);
  return cp !== undefined && NAME_CHAR.test(String.fromCodePoint(cp));
}

/** True when the code point ending just before `i` continues a word (see NAME_CHAR). */
export function nameCharBefore(s: string, i: number): boolean {
  if (i === 0) return false;
  const low = s.charCodeAt(i - 1);
  const high = i >= 2 ? s.charCodeAt(i - 2) : 0;
  const pair = low >= 0xdc00 && low <= 0xdfff && high >= 0xd800 && high <= 0xdbff;
  return nameCharAt(s, pair ? i - 2 : i - 1);
}

/** A value with ß (and ẞ) written as ss, as it is typed without a German keyboard. */
export function withoutSharpS(v: string): string {
  return /[\u00df\u1e9e]/u.test(v) ? v.replace(/\u00df/gu, "ss").replace(/\u1e9e/gu, "SS") : v;
}

const UMLAUTS: Readonly<Record<string, string>> = { "ä": "ae", "ö": "oe", "ü": "ue", "Ä": "Ae", "Ö": "Oe", "Ü": "Ue" };

/** A value with ä, ö and ü written ae, oe and ue, as German is typed without umlauts ("Mueller" for "Müller"). */
export function withoutUmlauts(v: string): string {
  const t = v.normalize("NFC");
  return /[äöüÄÖÜ]/u.test(t) ? t.replace(/[äöüÄÖÜ]/gu, (c) => UMLAUTS[c] ?? c) : v;
}

/** An apostrophe (or a look-alike one) or a hyphen. */
const JOINER_RE = /['\-\u2010-\u2015\u2018\u2019\u02bc\u2032\u02b9\u00b4\u0060\uff07\u2212]/u;
// An apostrophe or hyphen between a letter and a capital letter: "O'Brien", "D'Angelo", "Anne-Marie".
const NAME_JOINER_RE = /(?<=\p{L})['\-\u2010-\u2015](?=\p{Lu})/gu;
const NAME_HYPHEN_RE = /(?<=\p{L})[-\u2010-\u2015](?=\p{Lu})/gu;

/**
 * A name typed without its apostrophes and hyphens ("OBrien" for "O'Brien", "AnneMarie" for
 * "Anne-Marie"), and a one-word name with its hyphens as spaces ("Anne Marie"; the words of a
 * longer name are found on their own). Only where a capital letter follows, so "e-mail" and
 * "don't" don't give common words. Empty when there is no such mark.
 */
export function joinedNames(v: string): string[] {
  if (!JOINER_RE.test(v)) return [];
  const t = foldUnits(v).trim();
  const joined = t.replace(NAME_JOINER_RE, "");
  if (joined === t) return [];
  const spaced = /\s/u.test(t) ? t : t.replace(NAME_HYPHEN_RE, " ");
  return spaced === t ? [joined] : [joined, spaced];
}

/**
 * The part of an email address before the @, when it has 4+ characters and a digit or one of
 * . _ - ("sunflower77", "bluejay_fan", "kw.tokyo.1988"): typed alone, it still names the person.
 * Null otherwise; a plain local part ("maria") is as often an ordinary word or first name.
 */
export function emailLocalPart(v: string): string | null {
  const local = /^([^\s@]+)@[^\s@]+\.[^\s@]+$/u.exec(v.trim())?.[1];
  return local !== undefined && local.length >= 4 && /[\d._-]/u.test(local) ? local : null;
}

/**
 * The parts of a value between runs of letters that hold digits: "(831) 555-0199 x204" gives
 * "(831) 555-0199 " and "204", "GB29 NWBK 6016 1331 9268 19" gives "29 " and " 6016 1331 9268 19".
 * Empty for a value without letters. A part of 7+ digits is matched like a whole number, so a phone
 * number is found without its extension and an account number without the rest of its IBAN.
 */
export function letterSeparatedParts(v: string): string[] {
  if (!/\p{L}/u.test(v)) return [];
  return v.split(/\p{L}+/u).filter((part) => /\d/u.test(part));
}

/** A code of letters and digits with "-", "/", "." or spaces between its parts: "AB-1234-CD", "X7 42.B". */
const SEPARATED_CODE_RE = /^[\p{L}\p{N}]+(?:[-/.\s]+[\p{L}\p{N}]+)+$/u;
const CODE_SEPARATOR_RE = /[-/.\s]/u;

/**
 * A code of letters and digits written with "-", "/", "." or spaces between its parts, without
 * them: "AB1234CD" for "AB-1234-CD", as people type codes both ways. Characters that can't be seen
 * are skipped. Null when the value has any other character, has no letter or no digit, or would
 * be shorter than 4 characters.
 */
export function codeWithoutSeparators(v: string): string | null {
  const t = withoutIgnorable(v).trim();
  if (t.length < 5 || !CODE_SEPARATOR_RE.test(t) || !/\p{N}/u.test(t) || !/\p{L}/u.test(t) || !SEPARATED_CODE_RE.test(t)) return null;
  const joined = t.replace(/[-/.\s]+/gu, "");
  return joined.length >= 4 ? joined : null;
}

/**
 * A private value that mixes letters and digits and has this many of them or more is also found
 * with only some of its separators, or other ones (see codeLetters).
 */
export const MIN_CODE_LETTERS = 6;

const CODE_LETTER_RE = /^[\p{L}\p{M}\p{N}]$/u;
const CODE_CJK_RE = /^[\p{scx=Han}\p{scx=Hiragana}\p{scx=Katakana}\p{scx=Hangul}]$/u;
const codeLetterCache = new Map<number, boolean>();

/**
 * A letter, a mark or a digit, as the auditor's word characters are: not Chinese, Japanese or
 * Korean characters, which are written without spaces.
 */
export function isCodeLetterCodePoint(cp: number): boolean {
  if (cp < 0x80) return (cp >= 0x30 && cp <= 0x39) || (cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a);
  let v = codeLetterCache.get(cp);
  if (v === undefined) {
    const ch = String.fromCodePoint(cp);
    v = CODE_LETTER_RE.test(ch) && !CODE_CJK_RE.test(ch);
    codeLetterCache.set(cp, v);
  }
  return v;
}

/**
 * What may stand between the letters and digits of a code typed with other separators: any space,
 * "-", "_", ".", "/", "\" and "·", and their look-alikes. A colon only where it isn't a cell range's
 * (see isCodeSeparatorAt).
 */
export function isCodeSeparatorCode(c: number): boolean {
  return isSpaceCode(c) || isDashCode(c) || isSlashCode(c) || isBackslashCode(c) || c === 0x5f || c === 0x2e || c === 0xb7 || c === 0xff3f || c === 0xff0e;
}

/**
 * True when text[i], in a text folded for matching (the full-width colon is ":" there), separates
 * the parts of a code (see isCodeSeparatorCode). A colon does, "AB:1234:CD", except between the two
 * cells of a range: "A10:B20" is a range, not the code "A10-B20".
 */
export function isCodeSeparatorAt(text: string, i: number): boolean {
  const c = text.charCodeAt(i);
  return c === 0x3a ? !isRangeColon(text, i) : isCodeSeparatorCode(c);
}

/** An ASCII letter or digit at t[i] (false past either end). */
function isAsciiAlnumAt(t: string, i: number): boolean {
  const c = t.charCodeAt(i);
  return (c >= 0x30 && c <= 0x39) || (c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a);
}

/** The colon at t[i] of a cell range "A10:B20": 1 to 3 letters and 1 to 7 digits on each side, as whole words. */
function isRangeColon(t: string, i: number): boolean {
  const run = (k: number, step: number, test: (c: number) => boolean, max: number): number => {
    let n = 0;
    while (n < max && test(t.charCodeAt(k + step * n))) n++;
    return n;
  };
  const letter = (c: number) => (c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a);
  const before = run(i - 1, -1, isDigitCode, 7);
  const cols = before > 0 ? run(i - 1 - before, -1, letter, 3) : 0;
  if (cols === 0 || isAsciiAlnumAt(t, i - 1 - before - cols)) return false;
  const cols2 = run(i + 1, 1, letter, 3);
  const after = cols2 > 0 ? run(i + 1 + cols2, 1, isDigitCode, 7) : 0;
  return after > 0 && !isAsciiAlnumAt(t, i + 1 + cols2 + after);
}

/**
 * The letters and digits of a value already folded for matching (lowercased, accents off), when it
 * mixes letters and digits and has 6 or more of them: "ab1234cd" for "ab-1234-cd", "x742b9" for
 * "x7_42.b9". The auditor looks for them in the body's letters and digits with the separators
 * between them taken out (see isCodeSeparatorCode), so "AB 1234CD", "AB–1234CD", "AB.1234.CD" and
 * "xAB1234CD" all hold "AB-1234-CD"; substituteUserText does the same in the question. Null for
 * any other value.
 */
export function codeLetters(folded: string): string | null {
  let out = "";
  let letter = false;
  let digit = false;
  for (let i = 0; i < folded.length; ) {
    const cp = folded.codePointAt(i)!;
    const size = cp > 0xffff ? 2 : 1;
    if (isCodeLetterCodePoint(cp)) {
      if (cp >= 0x30 && cp <= 0x39) digit = true;
      else if (cp >= 0x80 && /^\p{N}$/u.test(String.fromCodePoint(cp))) digit = true;
      else letter = true;
      out += folded.slice(i, i + size);
    }
    i += size;
  }
  return letter && digit && out.length >= MIN_CODE_LETTERS ? out : null;
}

/**
 * A whole number of thousands or millions written short, exactly, as [number, unit]: 83500 gives
 * ["83.5", "k"], 1200000 gives ["1200", "k"] and ["1.2", "m"]. At most 2 decimals, so nothing is
 * rounded: 83512 gives none.
 */
export function shortNumbers(plain: string): [string, "k" | "m"][] {
  if (!/^\d+$/u.test(plain)) return [];
  const out: [string, "k" | "m"][] = [];
  for (const [unit, zeros] of [["k", 3], ["m", 6]] as const) {
    if (plain.length <= zeros) continue;
    const int = plain.slice(0, -zeros).replace(/^0+(?=\d)/u, "");
    const frac = plain.slice(-zeros).replace(/0+$/u, "");
    if (frac.length <= 2) out.push([frac === "" ? int : `${int}.${frac}`, unit]);
  }
  return out;
}

/**
 * Lowercases one character at a time (after foldUnits), keeping characters whose lowercase changes
 * length, and drops combining accents (U+0300–036F), as the auditor does. `at[k]` is the index in
 * `s` of the k-th unit of `text`, and `at[text.length]` is `s.length`; `at` is null when every
 * position is kept.
 */
function foldLower(s: string): { text: string; at: number[] | null } {
  // eslint-disable-next-line no-control-regex
  if (/^[\x00-\x7f]*$/.test(s)) return { text: s.toLowerCase(), at: null };
  let text = "";
  const at: number[] = [];
  const fold = (u: string) => (u.length === 1 ? String.fromCharCode(foldUnit(u.charCodeAt(0))) : u);
  for (let i = 0; i < s.length; ) {
    const ch = String.fromCodePoint(s.codePointAt(i)!);
    if (ch.length === 1 && isAccentMark(ch.charCodeAt(0))) {
      i++;
      continue;
    }
    // Fold, lowercase, fold again, as the auditor does: "İ" is "I" with a dot, then "i".
    const f = fold(ch);
    let l = f.toLowerCase();
    l = l.length === f.length ? fold(l) : f;
    text += l;
    for (let k = 0; k < l.length; k++) at.push(i + k);
    i += ch.length;
  }
  at.push(s.length);
  return { text, at };
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\/]/gu, "\\$&");
}

/**
 * The parts of a value between which the question may have spaces: its own whitespace, and
 * between two Chinese, Japanese or Korean characters ("山田太郎" typed "山田 太郎" or "山 田 太 郎").
 */
function pieces(needle: string): string[] {
  if (!CJK_CHAR.test(needle)) return needle.split(/\s+/u).filter((part) => part !== "");
  const out: string[] = [];
  for (const part of needle.split(/\s+/u)) {
    let cur = "";
    let prevCjk = false;
    for (const ch of part) {
      const cjk = CJK_CHAR.test(ch);
      if (cur !== "" && cjk && prevCjk) {
        out.push(cur);
        cur = "";
      }
      cur += ch;
      prevCjk = cjk;
    }
    if (cur !== "") out.push(cur);
  }
  return out;
}

/**
 * The first characters of a value's first piece (see pieces): up to 4, stopping at a space or a
 * Chinese, Japanese or Korean character, where a typed value may have a space the stored one
 * hasn't. Every match of the value holds it. Empty (found in any text) when the value starts there.
 */
function firstPiecePrefix(needle: string): string {
  let k = 0;
  while (k < 4 && k < needle.length) {
    const c = needle.charCodeAt(k);
    if (c <= 0x20 || c === 0xa0 || c >= 0x1100) break;
    k++;
  }
  return needle.slice(0, k);
}

/**
 * A whole-word pattern for a folded value. Its pieces match with any run of whitespace between
 * them, or none ("山田太郎" for "山田 太郎").
 */
function needlePattern(needle: string): RegExp {
  const chars = [...needle];
  const firstCh = chars[0]!;
  const lastCh = chars[chars.length - 1]!;
  const guarded = (ch: string) => WORD_CHAR.test(ch) && !CJK_CHAR.test(ch);
  // Whole word: no letter, digit or underscore right before or after, except in Chinese, Japanese
  // and Korean, where words aren't spaced ("山田太郎さん"). Numbers also don't continue through a
  // decimal point or thousands separator ("5" is not found inside "5.5" or "1,500").
  const before = guarded(firstCh) ? String.raw`(?<!${LATIN_WORD})(?<!\p{N}[.,])` : "";
  const after = guarded(lastCh) ? String.raw`(?!${LATIN_WORD})(?![.,]\p{N})` : "";
  return new RegExp(`${before}${pieces(needle).map(escapeRegExp).join(SPACES)}${after}`, "gu");
}

/**
 * A number of 7+ digits with any separators between its digits: "(831) 555-0199", "831.555.0199".
 * With `countryCode`, a country code may come first ("+44 20 7946 0958", "0044 20..."): for a
 * number stored with its trunk 0 ("020 7946 0958"), `digits` is then the number without it.
 */
function digitsPattern(digits: string, countryCode = false): RegExp {
  const prefix = countryCode ? String.raw`(?:(?:\+|00)\d{1,3}${DIGIT_SEPARATORS})?` : "";
  return new RegExp(String.raw`(?<!\d)${prefix}[(\uff08]?${[...digits].join(DIGIT_SEPARATORS)}(?!\d)`, "gu");
}

/** A space of any kind as a thousands separator (no-break, narrow no-break and thin spaces are folded to " "). */
const SPACE_GROUPING = String.raw`\s`;

/**
 * Thousands separators a typed number may have, each with the decimal marks that can follow a
 * number grouped with it: never the separator itself, so "5,390,000" is not 5,390 with zero
 * decimals.
 */
const GROUPINGS: readonly (readonly [separator: string, decimalMark: string])[] = [
  [",", String.raw`\.`],
  [String.raw`\.`, ","],
  [SPACE_GROUPING, "[.,]"],
  ["'", "[.,]"],
  ["_", "[.,]"],
];

/** A plain number of 4+ characters and fewer than 7 digits before any decimals: "93496", "0199", "12.75". */
const SHORT_PLAIN_RE = /^\d{1,6}(?:\.\d+)?$/u;

/**
 * A plain number of fewer than 7 digits before any decimals ("93496", "0199", "12.75"), as a
 * pattern source. It is found wherever no digit is right next to it: "93496,94103",
 * "=IF(B2=93496,1,0)", "93496.5", "ZIP93496", "1.93496" and "#1.4528" hold 93496 and 4528, but
 * "934961" doesn't. Zeros before it are taken with it when no digit comes before them ("02139" for
 * 2139, "#04521" for 4521), and so are zero decimals after it ("93496.00", "12.750" for 12.75). "."
 * or "," may be its decimal mark. Only such a clean occurrence is replaced. The auditor's "plain"
 * rule is the same outside free text; in free text, the question included, it also finds the
 * number with a digit next to it ("934961234", a ZIP+4 typed without its hyphen), so such a
 * question is blocked rather than a part of a longer number replaced.
 */
function plainNumberSource(plain: string): string {
  const [whole = "", frac] = plain.split(".");
  const tail = frac === undefined ? "(?:[.,]0+)?" : `[.,]${frac}0*`;
  return String.raw`(?<!\p{N})0*${whole}${tail}(?!\p{N})`;
}

/**
 * A number the user reads in a cell, written plainly ("83500", "5.25"), or with one kind of
 * thousands separator throughout ("83,500", "83.500", "83 500", "83'500", "83_500", and "8,35,000"
 * in lakh from 100,000), "." or "," as the decimal mark (a grouped number's decimal mark differs
 * from its separator), optional zero decimals after a whole number ("83,500.00"), or exactly in
 * thousands or millions ("83.5k", "83,5 K", "1.2M"; see shortNumbers). Written plainly with fewer
 * than 7 digits before any decimals, only a digit right next to it stops a match (see
 * plainNumberSource). Grouped with "," "." "'" or "_", only a digit, or the same separator and 3
 * digits, continuing it on either side: "5,390,000" doesn't hold 5,390, but "93 496,94 103" holds
 * 93 496 and "5,390.25" holds 5,390 (the question is free text, as in the auditor). Grouped with a
 * space of any kind, only a digit right next to it: a list typed one number after another, such as
 * "40 000 83 500" or "5 390 000", holds 83 500 and 390 000. Otherwise no digit or decimal mark may
 * continue it on either side, and no letter may follow k or M.
 */
function numberPattern(plain: string): RegExp {
  const [whole = "", frac] = plain.split(".");
  const tail = (mark: string) => (frac === undefined ? `(?:${mark}0+)?` : `${mark}${frac}`);
  const alternatives = whole.length < MIN_DIGITS ? [plainNumberSource(plain)] : [];
  const forms = whole.length < MIN_DIGITS ? [] : [`${whole}${tail("[.,]")}`];
  if (whole.length >= 4) {
    for (const [separator, mark] of GROUPINGS) {
      const grouped = [whole.replace(/\B(?=(?:\d{3})+$)/g, separator)];
      if (whole.length >= 6) grouped.push(`${whole.slice(0, -3).replace(/\B(?=(?:\d{2})+$)/g, separator)}${separator}${whole.slice(-3)}`);
      // A space doesn't carry a number on: "40 000 83 500" is two numbers as often as one.
      const around = separator === SPACE_GROUPING ? ["", ""] : [String.raw`(?<!\p{N}${separator})`, String.raw`(?!${separator}\p{N}{3})`];
      alternatives.push(String.raw`(?<!\p{N})${around[0]}(?:${grouped.join("|")})${tail(mark)}(?!\p{N})${around[1]}`);
    }
  }
  if (frac === undefined) {
    for (const [number, unit] of shortNumbers(whole)) {
      forms.push(String.raw`${number.replace(".", "[.,]")}\s?[${unit}${unit.toUpperCase()}](?!\p{L})`);
    }
  }
  if (forms.length > 0) alternatives.push(String.raw`(?<!\p{N})(?<!\p{N}[.,])(?:${forms.join("|")})(?!\p{N})(?![.,]\p{N})`);
  return new RegExp(alternatives.join("|"), "gu");
}

/**
 * Replaces private values (and original headers of aliased columns) in the user's question.
 * The question comes back NFC-normalized. Values match with either form of accents or none
 * ("Jose Munoz" for "José Muñoz"), ß as ss, ä ö ü as ae oe ue ("Mueller"), ligatures written out
 * ("Griffiths" for "Griﬃths"), a name without its apostrophe or hyphen ("OBrien", "AnneMarie"),
 * full-width letters, look-alike apostrophes, digits and hyphens, Turkish dotted and dotless i,
 * "_" for a space ("maria_lopez"), any date phrase that means a private date ("1985-03-15",
 * "3/15/85", "March 15th, 1985" or "15MAR1985" for 3/15/1985; see findDates), numbers in other
 * formats ("(831) 555-0199" for 831.555.0199, "+44 20 7946 0958" for 020 7946 0958, "8315550199"
 * for "(831) 555-0199 x204", "83,500", "83_500" or "83.5k" for 83500 or "83500" stored as text,
 * "5.25" for 5.25%), and a plain number of fewer than 7 digits wherever no digit is right next to
 * it, zeros before it included ("93496,94103", "=IF(B2=93496,1,0)", "02139" for 2139; see
 * plainNumberSource). Parts of a value, such as the last four digits of a phone number or an
 * email address's local part, are not replaced; the auditor blocks them.
 */
export function substituteUserText(text: string, specs: readonly RangeSpec[], map: StandInMap): string {
  if (typeof text !== "string" || text === "") return text;
  text = text.normalize("NFC");

  // Matching runs on folded copies of the question. `folded` keeps every position; `plain` is
  // `folded` without the characters matching drops (see isDroppedCodePoint), and `plainAt` maps its
  // positions back; `lower.at` maps positions in `lower` to `plain`. Other whitespace controls (a
  // tab, a line break) are spaces. The text around each match is kept as typed.
  const folded = foldUnits(text);
  const { text: plain, at: plainAt, gap } = stripDropped(folded);
  const lower = foldLower(plain);
  const plainDigits = plain.replace(/\D/g, "");
  const toText = (k: number) => (plainAt ? plainAt[k]! : k);
  // The end of a match in `plain` as a position in `text`: just past its last unit, so a dropped
  // character right after a match stays outside it.
  const endInText = (k: number) => (plainAt ? (k > 0 ? plainAt[k - 1]! + 1 : 0) : k);
  // Where characters were dropped, `plain` again with a mark in their place (see withGapMarks), so
  // a value is also found where a dropped character is the boundary next to it: "Mr\u200bWatanabe"
  // or "93496\u200b94103". `marked.at` maps its positions to `plain`.
  const marked = gap === null ? null : withGapMarks(plain, gap);
  const markedLower = marked === null ? null : foldLower(marked.text);
  // A word boundary right before plain[k], as a code's replacement needs at both of its ends: the
  // start or end of the question, a dropped character, or no letter or digit on one side.
  const boundaryAt = (k: number): boolean =>
    k <= 0 || k >= plain.length || gap?.[k] === 1 || !isCodeLetterCodePoint(codePointBefore(plain, k)) || !isCodeLetterCodePoint(plain.codePointAt(k)!);
  // The question's letters and digits, where a code typed with some or other separators is looked
  // for (see codeLetters); only when it has a run of them that mixes letters and digits.
  let code: { letters: string; at: number[] } | null = codeProjection(lower.text);
  if (!hasCodeRun(code.letters, 4)) code = null;
  const mayHoldCode = code !== null;
  // Every date phrase in the question, once needed.
  let phrases: DateMatch[] | null = null;
  const datePhrases = () => (phrases ??= findDates(folded));
  // A value is kept as a candidate only when the question holds the start of it (see
  // firstPiecePrefix), or all of its digits for a number: the others can't match, and a sheet
  // has thousands of them.
  const mayHold = (needle: string, caseless: boolean) => (caseless ? lower.text : plain).includes(firstPiecePrefix(needle));

  const candidates = new Map<string, Candidate>();
  const addDigits = (digits: string, replacement: () => string, countryCode: boolean): void => {
    const key = `${countryCode ? "t" : "d"}\u0000${digits}`;
    if (!candidates.has(key) && plainDigits.includes(digits)) {
      candidates.set(key, { needle: digits, replacement, caseless: false, pattern: () => digitsPattern(digits, countryCode) });
    }
  };
  // A number of 7+ digits typed with any separators ("2024-01-15" for invoice number 20240115),
  // without its trunk 0 or country code too.
  const addDigitForms = (s: string, replacement: () => string): void => {
    const digits = s.replace(/\D/g, "");
    if (digits.length >= MIN_DIGITS) addDigits(digits, replacement, false);
    if (digits.length >= MIN_TRUNK_DIGITS && digits.startsWith("0")) addDigits(digits.slice(1), replacement, true);
    for (const national of withoutCountryCode(s)) addDigits(national, replacement, false);
  };
  // Values of 4+ characters match in any case; shorter ones exactly, so "An" doesn't rewrite "an".
  // `digitForms` false skips a value's digits typed with other separators (the value they come
  // from has them already).
  const addNeedle = (form: string, replacement: () => string, anyCase: boolean, digitForms = true): void => {
    // Plain ASCII (the usual case) needs no normalizing or folding but lowercasing.
    const ascii = PLAIN_RE.test(form);
    const value = withoutIgnorable(ascii ? form : foldUnits(form.normalize("NFC"))).trim();
    if (value === "") return;
    if (!anyCase && value.length >= MIN_NUMBER && SHORT_PLAIN_RE.test(value)) {
      // A plain number such as a ZIP code: found with no digit right next to it (see plainNumberSource).
      const key = `p\u0000${value}`;
      if (mayHold(value, false) && !candidates.has(key)) {
        candidates.set(key, { needle: value, replacement, caseless: false, pattern: () => new RegExp(plainNumberSource(value), "gu") });
      }
      return;
    }
    const caseless = anyCase || value.length >= CASE_INSENSITIVE_FROM;
    const needle = caseless ? (ascii ? value.toLowerCase() : foldLower(value).text) : value;
    const key = `${caseless ? "i" : "c"}\u0000${needle}`;
    if (mayHold(needle, caseless) && !candidates.has(key)) candidates.set(key, { needle, replacement, caseless });
    if (anyCase || !digitForms || !/\d/.test(value)) return;
    addDigitForms(value, replacement);
    // The digits between letters: a phone number without its extension ("(831) 555-0199 x204").
    for (const part of letterSeparatedParts(value)) addDigitForms(part, replacement);
  };
  // A value as typed, without its accents (foldUnits folds accented letters; this takes off other
  // marks), with ß as ss, ä ö ü as ae oe ue, compatibility characters such as the ligature "ﬃ"
  // written out (NFKC), and a name without its apostrophes or hyphens ("OBrien" for "O'Brien").
  // needlePattern lets "_" stand for a space.
  const addValue = (v: string, replacement: () => string): void => {
    const plain = withoutMarks(v);
    // eslint-disable-next-line no-control-regex
    const forms = /^[\x00-\x7f]*$/.test(v)
      ? [v]
      : [v, plain, withoutSharpS(v), withoutSharpS(plain), v.normalize("NFKC"), withoutUmlauts(v), withoutSharpS(withoutUmlauts(v))];
    for (const f of new Set(forms)) {
      addNeedle(f, replacement, false);
      for (const joined of joinedNames(f)) addNeedle(joined, replacement, false);
    }
    if (!mayHoldCode) return;
    // A code of 6+ letters and digits typed with some of its separators, or other ones, or with a
    // letter or digit touching it: "AB 1234CD", "AB–1234CD", "AB.1234.CD", "xAB1234CD" for
    // "AB-1234-CD" (see codeLetters), as the auditor finds it.
    const ascii = PLAIN_RE.test(v);
    const letters = codeLetters(ascii ? v.toLowerCase() : foldLower(withoutIgnorable(foldUnits(v.normalize("NFC")))).text);
    if (letters !== null) {
      const key = `k\u0000${letters}`;
      if (code !== null && !candidates.has(key) && code.letters.includes(letters)) {
        const { letters: hay, at } = code;
        // Only a whole code is replaced, one that starts and ends on a word boundary: "area 10023"
        // holds the letters and digits of "A10023", but replaced they would give "areID_001".
        const spans = () => {
          const out: { index: number; length: number }[] = [];
          for (let i = hay.indexOf(letters); i !== -1; i = hay.indexOf(letters, i + 1)) {
            const start = at[i]!;
            const end = at[i + letters.length - 1]! + 1;
            if (boundaryAt(lower.at ? lower.at[start]! : start) && boundaryAt(lower.at ? lower.at[end]! : end)) out.push({ index: start, length: end - start });
          }
          return out;
        };
        candidates.set(key, { needle: letters, replacement, caseless: true, spans });
      }
      return;
    }
    // A shorter code typed without its separators, as a whole word: "A1030" for "A-1030".
    const joined = codeWithoutSeparators(v);
    if (joined !== null) addNeedle(joined, replacement, false, false);
  };

  const headers: [string, string][] = [];
  const numbersTried = new Set<string>();
  /** Private dates by dateKey: the stand-in of the first cell holding each. */
  const dateReplacements = new Map<string, () => string>();
  for (const spec of specs) {
    const d = spec.data;
    for (const p of plans(spec)) {
      if (p.policy.alias && !aliasRevealsHeader(p.policy.alias, p.info.header)) headers.push([p.info.header, p.policy.alias]);
      if (!p.policy.private) continue;
      let kind: TokenKind | undefined;
      const kindOf = () => (kind ??= tokenKindFor(p.info, columnTexts(spec, p.offset)));
      // The dates of 8-digit cells, for a column that holds dates stored that way.
      const loose = columnLooseDates(d, firstDataRow(spec), p.offset);
      for (let r = firstDataRow(spec); r < rowCount(d); r++) {
        const value = d.values[r]?.[p.offset];
        const shown = d.text[r]?.[p.offset];
        const key = cellKey(value ?? null, shown);
        if (key === null) continue;
        const replacement = () => map.tokenFor(key, kindOf());
        const raw = value === null || value === undefined ? undefined : String(value);
        // "#####" only says the column is too narrow for the number.
        if (typeof shown === "string" && !isHiddenNumberText(shown)) addValue(shown, replacement);
        if (raw !== undefined && raw !== shown) addValue(raw, replacement);
        // A date, as a whole date only: any date phrase in the question that means it (see
        // findDates). The dates 8 digits can mean count too, in a column of them: the question is
        // free text.
        for (const date of [...cellDates(d, r, p.offset, p.info.type), ...(loose?.get(r) ?? [])]) {
          const k = dateKey(date);
          if (!dateReplacements.has(k)) dateReplacements.set(k, replacement);
        }
        // The number the user sees, in other formats, whatever the column holds.
        for (const number of cellNumbers(d, r, p.offset)) {
          const key = `g\u0000${number}`;
          if (numbersTried.has(key)) continue;
          numbersTried.add(key);
          // The digits every typed form holds: all of them, but only "835" of 83500 ("83.5k").
          const short = shortNumbers(number).at(-1);
          const digits = (short ? short[0] : number).replace(/\D/g, "");
          if (plainDigits.includes(digits)) {
            candidates.set(key, { needle: number, replacement, caseless: false, pattern: () => numberPattern(number) });
          }
        }
      }
    }
  }
  // Headers match in any case: they are sent only under the alias. A value with the same text wins.
  // A header that is a code is also found typed without its separators, as a whole word:
  // "FalconX92024" for "Falcon X9-2024".
  for (const [header, alias] of headers) {
    const code = codeWithoutSeparators(header);
    for (const form of code === null ? [header] : [header, code]) {
      const needle = foldLower(withoutIgnorable(foldUnits(form.normalize("NFC"))).trim()).text;
      const key = `i\u0000${needle}`;
      if (needle === "" || candidates.has(key) || !mayHold(needle, true)) continue;
      candidates.set(key, { needle, replacement: () => alias, caseless: true });
    }
  }

  // Tokens already in the question stay as they are.
  const claimed: { start: number; end: number; replacement: (() => string) | null }[] = [];
  for (const m of text.matchAll(new RegExp(TOKEN_RE.source, "g"))) {
    if (map.isToken(m[0])) claimed.push({ start: m.index, end: m.index + m[0].length, replacement: null });
  }
  const overlaps = (s: number, e: number) => claimed.some((c) => s < c.end && c.start < e);

  const ordered = [...candidates.values()].sort((a, b) => b.needle.length - a.needle.length);
  // Date phrases that mean a private date ("4th of July 1988", "3/15/85"). Among themselves the
  // one that starts first wins, the longer one when two start together: in "March 15, 1985 -
  // September 2, 1985" both dates are replaced, not the "1985 - September 2" that joins them. They
  // are then claimed in the same longest-first order as the values, so a longer private text
  // holding a date still wins.
  type DateSpan = { start: number; end: number; replacement: () => string };
  const found: DateSpan[] = [];
  if (dateReplacements.size > 0) {
    for (const f of datePhrases()) {
      const k = f.dates.map(dateKey).find((key) => dateReplacements.has(key));
      if (k !== undefined) found.push({ start: f.start, end: f.end, replacement: dateReplacements.get(k)! });
    }
    found.sort((a, b) => a.start - b.start || b.end - a.end);
  }
  const dates: DateSpan[] = [];
  let reach = 0;
  for (const x of found) {
    if (x.start < reach) continue;
    dates.push(x);
    reach = x.end;
  }
  dates.sort((a, b) => b.end - b.start - (a.end - a.start) || a.start - b.start);
  let nextDate = 0;
  const claimDates = (minLength: number): void => {
    for (; nextDate < dates.length && dates[nextDate]!.end - dates[nextDate]!.start >= minLength; nextDate++) {
      const x = dates[nextDate]!;
      if (!overlaps(x.start, x.end)) claimed.push(x);
    }
  };
  // Where a match in `lower`, `marked` or its lowercase copy is in `plain`.
  const inPlain = (at: number[] | null, k: number) => (at ? at[k]! : k);
  for (const c of ordered) {
    claimDates(c.needle.length);
    const hay = c.caseless ? lower.text : plain;
    // Cheap filter before building a pattern: every piece of the value must appear somewhere.
    if (!c.pattern && !c.spans && !pieces(c.needle).every((part) => hay.includes(part))) continue;
    // Spans in `plain`: in `hay`, and in `marked` (or its lowercase copy) when characters were dropped.
    const spans: { start: number; end: number }[] = [];
    const back = c.caseless ? lower.at : null;
    if (c.spans) {
      for (const m of c.spans()) spans.push({ start: inPlain(back, m.index), end: inPlain(back, m.index + m.length) });
    } else {
      const re = c.pattern ? c.pattern() : needlePattern(c.needle);
      for (const m of hay.matchAll(re)) spans.push({ start: inPlain(back, m.index), end: inPlain(back, m.index + m[0].length) });
      if (marked !== null && markedLower !== null) {
        const hay2 = c.caseless ? markedLower.text : marked.text;
        const back2 = c.caseless ? markedLower.at : null;
        for (const m of hay2.matchAll(re)) {
          spans.push({ start: marked.at[inPlain(back2, m.index)]!, end: marked.at[inPlain(back2, m.index + m[0].length)]! });
        }
      }
    }
    for (const m of spans) {
      if (m.end <= m.start) continue;
      const start = toText(m.start);
      const end = endInText(m.end);
      if (overlaps(start, end)) continue;
      claimed.push({ start, end, replacement: c.replacement });
    }
  }
  claimDates(0);

  // No letter or digit of a private date phrase may stay outside the replaced text, as in "March
  // 15, TEXT_001, 1985". Where one would, the phrase and every replaced part it overlaps become one
  // stand-in. Only separators may be left between two stand-ins ("TEXT_001 - TEXT_002").
  const replacedAt = (i: number) => claimed.some((c) => c.start <= i && i < c.end);
  for (const x of found) {
    let whole = true;
    for (let i = x.start; i < x.end && whole; i++) if ((numberAt(folded, i) || letterAt(folded, i)) && !replacedAt(i)) whole = false;
    if (whole) continue;
    let start = x.start;
    let end = x.end;
    for (let grew = true; grew; ) {
      grew = false;
      for (let k = claimed.length - 1; k >= 0; k--) {
        const c = claimed[k]!;
        if (c.start >= end || start >= c.end) continue;
        start = Math.min(start, c.start);
        end = Math.max(end, c.end);
        claimed.splice(k, 1);
        grew = true;
      }
    }
    claimed.push({ start, end, replacement: x.replacement });
  }

  // Tokens are allocated in reading order.
  claimed.sort((a, b) => a.start - b.start);
  let out = "";
  let last = 0;
  for (const c of claimed) {
    out += text.slice(last, c.start) + (c.replacement ? c.replacement() : text.slice(c.start, c.end));
    last = c.end;
  }
  return out + text.slice(last);
}
