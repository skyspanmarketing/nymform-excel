// Local defaults for the "Keep private" checkbox (spec §7.3). Suggestions only: the user's checkbox
// is the decision. Nothing here leaves the task pane.
import { columnIndex } from "./a1";
import { cellKind, columnCount, columnOffset, rowCount, wordCount, type CellKind } from "./schema";
import type { ColType, ColumnInfo, RangeData, Treatment } from "./types";

export interface Suggestion {
  letter: string;
  private: boolean;
  treatment: Treatment;
  /** Suggested neutral header for headers that are revealing on their own. */
  alias?: string;
  /** Plain reasons shown in the UI, e.g. "Header mentions email". */
  reasons: string[];
}

/** Header words that suggest a column is private (spec §7.3), matched as whole words. */
export const HEADER_KEYWORDS: readonly string[] = [
  "name", "first", "last", "surname", "email", "e-mail", "phone", "mobile", "cell", "ssn", "social",
  "dob", "birth", "birthday", "address", "street", "city", "zip", "postal", "id", "student", "employee",
  "emplid", "salary", "wage", "pay", "income", "gpa", "grade", "score", "diagnosis", "medical",
  "health", "gender", "sex", "race", "ethnicity", "religion", "disability", "notes", "comment", "reason",
  // Beyond the spec's list: headers naming the person a row is about ("Customer", "Client name").
  "customer", "client", "patient", "instructor", "teacher", "contact", "member", "applicant", "person",
  "tenant", "donor", "guardian", "parent", "recipient",
];

/** Headers that reveal something on their own, so the column gets a neutral alias. */
export const REVEALING_KEYWORDS: readonly string[] = [
  "diagnosis", "medical", "health", "disability", "religion", "race", "ethnicity", "gender", "sex",
];

/** Run-together forms of the keywords above, common in exported data ("FirstName" is split anyway). */
const COMPOUND_KEYWORDS: Record<string, string> = {
  firstname: "first",
  lastname: "last",
  fullname: "name",
  username: "name",
  fname: "first",
  lname: "last",
  emailaddress: "email",
  phonenumber: "phone",
  cellphone: "phone",
  telephone: "phone",
  zipcode: "zip",
  postcode: "postal",
  postalcode: "postal",
  dateofbirth: "dob",
  birthdate: "birth",
  studentid: "student",
  employeeid: "employee",
};

const KEYWORD_SET = new Set(HEADER_KEYWORDS.filter((k) => k !== "e-mail"));
const REVEALING_SET = new Set(REVEALING_KEYWORDS);

/** Up to this many non-empty cells per column are checked for value patterns. */
export const SAMPLE_CELLS = 200;

export const EMAIL_RE = /[^\s@<>()"',;:]+@[^\s@<>()"',;:]+\.[A-Za-z]{2,}/u;
const EMAIL_EXACT_RE = /^[^\s@<>()"',;:]+@[^\s@<>()"',;:]+\.[A-Za-z]{2,}$/u;
const PHONE_SHAPE_RE = /^[+(]?[\d\s().\-/+]+(?:\s*(?:x|ext\.?)\s*\d+)?$/iu;
const SSN_RE = /^(?:\d{3}-\d{2}-\d{4}|\d{3} \d{2} \d{4}|\d{9})$/u;
const ID_SHAPE_RE = /^[\p{L}\p{N}][\p{L}\p{N}\-_./#]*$/u;

/**
 * Lowercase words of a header: split on anything that isn't a letter or digit, on camelCase and on
 * letter/digit changes. "E-mail" and "e mail" also yield "email".
 */
export function headerWords(header: string): string[] {
  const spaced = header
    .replace(/(\p{Ll})(\p{Lu})/gu, "$1 $2")
    .replace(/(\p{Lu})(\p{Lu}\p{Ll})/gu, "$1 $2")
    .replace(/(\p{L})(\p{N})/gu, "$1 $2")
    .replace(/(\p{N})(\p{L})/gu, "$1 $2");
  const words = spaced
    .toLowerCase()
    .split(/[^\p{L}\p{N}]+/u)
    .filter((w) => w !== "");
  const out = [...words];
  for (let i = 0; i + 1 < words.length; i++) {
    if (words[i] === "e" && words[i + 1] === "mail") out.push("email");
  }
  return out;
}

/** Keywords from HEADER_KEYWORDS found in a header, in the order they appear. */
export function headerKeywords(header: string): string[] {
  const found: string[] = [];
  for (const w of headerWords(header)) {
    const k = KEYWORD_SET.has(w) ? w : COMPOUND_KEYWORDS[w];
    if (k !== undefined && !found.includes(k)) found.push(k);
  }
  return found;
}

export function isEmailText(text: string): boolean {
  return EMAIL_EXACT_RE.test(text.trim());
}

export function containsEmail(text: string): boolean {
  return EMAIL_RE.test(text);
}

/** 10 or more digits in something shaped like a phone number (digits, spaces, dashes, parentheses). */
export function isPhoneText(text: string): boolean {
  const t = text.trim();
  if (!PHONE_SHAPE_RE.test(t)) return false;
  const digits = t.replace(/\D/g, "").length;
  return digits >= 10 && digits <= 15;
}

export function isSsnText(text: string): boolean {
  return SSN_RE.test(text.trim());
}

/** Shape of one ID-like value: 5 to 15 characters, no spaces, at least one digit. */
export function isIdShaped(text: string): boolean {
  const t = text.trim();
  return t.length >= 5 && t.length <= 15 && /\p{N}/u.test(t) && ID_SHAPE_RE.test(t);
}

/** A cell of one column, as the value-pattern checks see it. */
export interface SampleCell {
  text: string;
  value: RangeData["values"][number][number] | undefined;
  kind: CellKind;
}

/** The first `limit` non-empty data cells of the column at `offset`. */
export function sampleColumn(data: RangeData, offset: number, hasHeaders: boolean, limit = SAMPLE_CELLS): SampleCell[] {
  const out: SampleCell[] = [];
  const rows = rowCount(data);
  for (let r = hasHeaders ? 1 : 0; r < rows && out.length < limit; r++) {
    const value = data.values[r]?.[offset];
    const text = data.text[r]?.[offset];
    const kind = cellKind(value, text, data.valueTypes[r]?.[offset], data.numberFormat[r]?.[offset]);
    if (kind === "empty") continue;
    out.push({ text: text !== undefined && text !== "" ? text : String(value ?? ""), value, kind });
  }
  return out;
}

const MAJORITY = 0.5;
const ID_SHARE = 0.8;
const ID_DISTINCT = 0.9;
const FREE_TEXT_WORDS = 6;
const UNSURE_DISTINCT = 0.5;

interface ValueFindings {
  email: boolean;
  phone: boolean;
  ssn: boolean;
  id: boolean;
  freeText: boolean;
}

/** Value patterns from spec §7.3 over up to 200 non-empty cells. */
export function valuePatterns(cells: readonly SampleCell[], avgWords?: number): ValueFindings {
  const n = cells.length;
  const none = { email: false, phone: false, ssn: false, id: false, freeText: false };
  if (n === 0) return none;
  let email = 0;
  let phone = 0;
  let ssn = 0;
  let idShaped = 0;
  let words = 0;
  const distinct = new Set<string>();
  for (const cell of cells) {
    const t = cell.text;
    if (containsEmail(t)) email++;
    if (isPhoneText(t)) phone++;
    if (isSsnText(t)) ssn++;
    if (idCandidate(cell) && isIdShaped(t)) idShaped++;
    words += wordCount(t);
    distinct.add(t);
  }
  const avg = avgWords ?? words / n;
  return {
    // Any address at all: an email in a cell is personal data wherever it sits.
    email: email > 0,
    phone: phone / n >= MAJORITY,
    ssn: ssn / n >= MAJORITY,
    id: idShaped / n >= ID_SHARE && distinct.size / n > ID_DISTINCT,
    freeText: avg > FREE_TEXT_WORDS,
  };
}

const DATE_TEXT_RE = /^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}$/u;

/** Text that isn't a date, or a whole number: amounts with decimals and dates are not IDs. */
function idCandidate(cell: SampleCell): boolean {
  if (cell.kind === "text") return !DATE_TEXT_RE.test(cell.text.trim());
  return cell.kind === "number" && typeof cell.value === "number" && Number.isInteger(cell.value);
}

/** Default treatment for a private column of a given type (brief, decision 4). */
export function defaultTreatment(type: ColType): Treatment {
  if (type === "number") return "range";
  if (type === "date") return "month";
  return "stand_in";
}

// ---------------------------------------------------------------------------------------------
// Formula references (derived columns)

/** Sheet columns a formula may read, or "all" when it reads whole rows or a whole table. */
export type FormulaRefs = { all: boolean; columns: Set<number>; names: Set<string> };

const CELL = String.raw`\$?([A-Za-z]{1,3})\$?(\d{1,7})`;
const RANGE_A1_RE = new RegExp(String.raw`(?<![\p{L}\p{N}_.])${CELL}\s*:\s*${CELL}(?![\p{L}\p{N}_(\[])`, "gu");
const CELL_RE = new RegExp(String.raw`(?<![\p{L}\p{N}_.])${CELL}(?![\p{L}\p{N}_(\[])`, "gu");
const COLS_RE = /(?<![\p{L}\p{N}_.$])\$?([A-Za-z]{1,3})\s*:\s*\$?([A-Za-z]{1,3})(?![\p{L}\p{N}_([])/gu;
const ROWS_RE = /(?<![\p{L}\p{N}_.$])\$?(\d{1,7})\s*:\s*\$?(\d{1,7})(?![\p{L}\p{N}_.([])/u;
const WHOLE_TABLE_RE = /[\p{L}_\\][\p{L}\p{N}_.\\]*\[\s*(?:#(?:All|Data|This Row|Headers|Totals)\s*)?\]/iu;
const INDIRECT_RE = /(?<![\p{L}\p{N}_.])(?:_xlfn\.)?INDIRECT\s*\(/iu;
const BRACKET_RE = /\[([^[\]]*)\]/gu;

/**
 * References in a formula, read with simple patterns that err toward finding too many: A1 cells,
 * A1 ranges, whole columns, whole rows, and structured references by column name. Sheet names are
 * ignored, so a reference to column B on any sheet counts as column B. Text inside string literals
 * is not a reference, except through INDIRECT, which counts as reading everything.
 */
export function formulaRefs(formula: string): FormulaRefs {
  const refs: FormulaRefs = { all: false, columns: new Set(), names: new Set() };
  const f = formula.replace(/"(?:[^"]|"")*"?/gu, " ").replace(/'(?:[^']|'')*'?!/gu, "!");
  const addSpan = (a: string, b: string) => {
    const i = columnIndex(a);
    const j = columnIndex(b);
    if (i < 0 || j < 0) return;
    for (let k = Math.min(i, j); k <= Math.max(i, j); k++) refs.columns.add(k);
  };
  for (const m of f.matchAll(RANGE_A1_RE)) addSpan(m[1]!, m[3]!);
  for (const m of f.matchAll(CELL_RE)) addSpan(m[1]!, m[1]!);
  for (const m of f.matchAll(COLS_RE)) addSpan(m[1]!, m[2]!);
  if (ROWS_RE.test(f) || WHOLE_TABLE_RE.test(f) || INDIRECT_RE.test(f)) refs.all = true;
  for (const m of f.matchAll(BRACKET_RE)) {
    let inner = m[1]!.trim();
    if (inner.startsWith("@")) {
      inner = inner.slice(1).trim();
      if (inner === "") refs.all = true; // Table[@]: the whole row
    }
    if (inner === "" || inner.startsWith("#")) continue;
    // Table column names escape special characters with an apostrophe.
    refs.names.add(inner.replace(/'(.)/gu, "$1").toLowerCase());
  }
  return refs;
}

// ---------------------------------------------------------------------------------------------

export function suggestPolicies(data: RangeData, columns: readonly ColumnInfo[], hasHeaders = true): Suggestion[] {
  const rows = rowCount(data);
  const dataRows = Math.max(0, rows - (hasHeaders ? 1 : 0));

  const states = columns.map((col, i) => {
    const offset = columnOffset(data, col.letter, i);
    const keywords = hasHeaders ? headerKeywords(col.header) : [];
    return {
      col,
      offset,
      keywords,
      revealing: keywords.filter((k) => REVEALING_SET.has(k)),
      isPrivate: false,
      exclude: false,
      reasons: [] as string[],
    };
  });

  for (const s of states) {
    const { col } = s;
    if (s.keywords.length > 0) {
      s.isPrivate = true;
      s.reasons.push(`Header mentions ${s.keywords.join(", ")}`);
    }

    const cells = sampleColumn(data, s.offset, hasHeaders);
    const found = valuePatterns(cells, col.type === "text" ? col.stats.avgWords : undefined);
    const valueReasons: string[] = [];
    if (found.email) valueReasons.push("Values contain email addresses");
    if (found.phone) valueReasons.push("Values look like phone numbers");
    if (found.ssn) valueReasons.push("Values look like 9-digit ID numbers");
    if (found.id && !found.phone && !found.ssn) valueReasons.push("Values look like unique IDs");
    if (found.freeText && (col.type === "text" || col.type === "mixed")) {
      valueReasons.push("Values are long free text, so the column is left out of sample rows");
      s.exclude = true;
    }
    if (valueReasons.length > 0) {
      s.isPrivate = true;
      s.reasons.push(...valueReasons);
    }

    // Without a header row there is nothing to say what a column holds, so every column starts
    // private ("when unsure, default to private"); the user unticks what can go out as is.
    if (!s.isPrivate && !hasHeaders && col.type !== "empty") {
      s.isPrivate = true;
      s.reasons.push("There is no header row to say what this column holds, so it is kept private to be safe");
    }

    if (!s.isPrivate && (col.type === "text" || col.type === "mixed")) {
      const nonEmpty = dataRows - col.stats.blank;
      if (nonEmpty > 0 && col.stats.distinct / nonEmpty > UNSURE_DISTINCT) {
        s.isPrivate = true;
        s.reasons.push("Many different values, so it is kept private to be safe");
      }
    }
  }

  // Derived columns: formulas that read a private column make their column private too.
  const byName = new Map<string, number[]>();
  const addName = (name: string, i: number) => {
    const key = name.trim().toLowerCase();
    if (key === "") return;
    byName.set(key, [...(byName.get(key) ?? []), i]);
  };
  states.forEach((s, i) => {
    if (hasHeaders) addName(s.col.header, i);
    const tableName = data.table?.columns[s.offset];
    if (tableName !== undefined) addName(tableName, i);
  });
  const refsByColumn = states.map((s) => columnFormulaRefs(data, s.offset, hasHeaders));
  const selfIndex = new Map<number, number>(); // sheet column -> index in states
  states.forEach((s, i) => selfIndex.set(data.columnIndex + s.offset, i));

  let changed = true;
  while (changed) {
    changed = false;
    states.forEach((s, i) => {
      const refs = refsByColumn[i];
      if (s.isPrivate || !refs) return;
      const readsPrivate =
        (refs.all && states.some((o) => o.isPrivate)) ||
        [...refs.columns].some((c) => {
          const j = selfIndex.get(c);
          return j !== undefined && j !== i && states[j]!.isPrivate;
        }) ||
        [...refs.names].some((name) => (byName.get(name) ?? []).some((j) => j !== i && states[j]!.isPrivate));
      if (readsPrivate) {
        s.isPrivate = true;
        s.reasons.push("Formulas use a private column");
        changed = true;
      }
    });
  }

  return states.map((s): Suggestion => {
    const out: Suggestion = {
      letter: s.col.letter,
      private: s.isPrivate,
      treatment: !s.isPrivate ? "as_is" : s.exclude ? "exclude" : defaultTreatment(s.col.type),
      reasons: s.reasons,
    };
    if (s.revealing.length > 0) {
      out.alias = neutralAlias(s.col.letter);
      out.reasons.push("Header is revealing on its own, so a neutral name is suggested");
    }
    return out;
  });

}

/** Every reference made by the formulas in one column's data cells, or null when it holds none. */
function columnFormulaRefs(data: RangeData, offset: number, hasHeaders: boolean): FormulaRefs | null {
  if (offset < 0 || offset >= columnCount(data)) return null;
  let merged: FormulaRefs | null = null;
  for (let r = hasHeaders ? 1 : 0; r < rowCount(data); r++) {
    const f = data.formulas[r]?.[offset];
    if (typeof f !== "string" || !f.startsWith("=")) continue;
    const refs = formulaRefs(f);
    if (!merged) {
      merged = refs;
      continue;
    }
    merged.all ||= refs.all;
    for (const c of refs.columns) merged.columns.add(c);
    for (const n of refs.names) merged.names.add(n);
  }
  return merged;
}

/** "Field D" style alias for a column letter. */
export function neutralAlias(letter: string): string {
  return `Field ${letter}`;
}
