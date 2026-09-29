// Bench task shape and result comparison (spec §11). Pure: no Office.js, no network.
// Details describe positions and counts only, never cell values, so result files can be shared.
import type { CellValue } from "../src/core/types";

export type Category = "conditional_totals" | "counting" | "lookup" | "dates" | "text" | "ranking" | "row_anomalies";

export const CATEGORIES: readonly Category[] = [
  "conditional_totals",
  "counting",
  "lookup",
  "dates",
  "text",
  "ranking",
  "row_anomalies",
];

/**
 * cell: one value. column: one value per data row, in row order. table: rows compared
 * order-insensitive. `tolerance` applies to numbers (absolute difference).
 */
export type Check =
  | { type: "cell"; expected: CellValue; tolerance?: number }
  | { type: "column"; expected: CellValue[]; tolerance?: number }
  | { type: "table"; expected: CellValue[][]; tolerance?: number };

export interface BenchTask {
  id: string;
  workbook: string;
  sheet: string;
  /** Local address of the selection, e.g. "A1:F501". Header row is row 1. */
  range: string;
  category: Category;
  prompt: string;
  /** Private column letters. */
  private: string[];
  check: Check;
}

export interface CompareResult {
  pass: boolean;
  detail: string;
}

/** Used when a check gives no tolerance: exact up to floating-point noise. */
export const DEFAULT_TOLERANCE = 1e-9;

const TOTAL_LABELS = new Set(["total", "grand total"]);
const EXCEL_ERROR_RE = /^#(?:N\/A|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!)$/i;

function isBlank(v: CellValue | undefined): boolean {
  return v === null || v === undefined || v === "";
}

function isExcelError(v: CellValue | undefined): boolean {
  return typeof v === "string" && EXCEL_ERROR_RE.test(v.trim());
}

function plainNumber(s: string): number | null {
  const t = s.trim();
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(t)) return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}

/**
 * One value against one expected value.
 * Numbers: within tolerance (a string holding a plain number counts as that number).
 * Strings: trimmed, case-insensitive (a number read back is compared by its plain text form).
 * Booleans: equal, or the text TRUE/FALSE. Empty expected: null or "".
 */
export function valuesMatch(actual: CellValue | undefined, expected: CellValue, tolerance = DEFAULT_TOLERANCE): boolean {
  if (expected === null || expected === "") return isBlank(actual);
  if (actual === null || actual === undefined) return false;
  if (typeof expected === "number") {
    const n = typeof actual === "number" ? actual : typeof actual === "string" ? plainNumber(actual) : null;
    return n !== null && Math.abs(n - expected) <= tolerance + 1e-12;
  }
  if (typeof expected === "boolean") {
    if (typeof actual === "boolean") return actual === expected;
    return typeof actual === "string" && actual.trim().toUpperCase() === (expected ? "TRUE" : "FALSE");
  }
  const e = expected.trim().toLowerCase();
  if (typeof actual === "string") return actual.trim().toLowerCase() === e;
  if (typeof actual === "number") return String(actual) === e;
  return false;
}

/** Drops trailing empty rows and trailing empty columns (a spill or fill never needs them). */
export function trimActual(actual: readonly (readonly CellValue[])[]): CellValue[][] {
  const rows = actual.map((r) => [...r]);
  while (rows.length > 0 && rows[rows.length - 1]!.every(isBlank)) rows.pop();
  let width = 0;
  for (const r of rows) {
    for (let c = r.length - 1; c >= 0; c--) {
      if (!isBlank(r[c])) {
        width = Math.max(width, c + 1);
        break;
      }
    }
  }
  return rows.map((r) => {
    const out = r.slice(0, width);
    while (out.length < width) out.push(null);
    return out;
  });
}

function rowsMatch(actual: readonly CellValue[], expected: readonly CellValue[], tolerance: number): boolean {
  if (actual.length !== expected.length) return false;
  return expected.every((e, i) => valuesMatch(actual[i], e, tolerance));
}

/** A leading row of labels: every cell is non-empty text and it matches no expected row. */
function isHeaderRow(row: readonly CellValue[], expected: readonly (readonly CellValue[])[], tolerance: number): boolean {
  if (row.length === 0 || !row.every((v) => typeof v === "string" && v.trim() !== "")) return false;
  const expectsNonText = expected.some((r) => r.some((v) => typeof v === "number" || typeof v === "boolean"));
  return expectsNonText && !expected.some((r) => rowsMatch(row, r, tolerance));
}

function isTotalRow(row: readonly CellValue[], expected: readonly (readonly CellValue[])[]): boolean {
  const first = row[0];
  if (typeof first !== "string" || !TOTAL_LABELS.has(first.trim().toLowerCase())) return false;
  return !expected.some((r) => typeof r[0] === "string" && TOTAL_LABELS.has(r[0].trim().toLowerCase()));
}

function shape(rows: readonly (readonly CellValue[])[]): string {
  const width = rows[0]?.length ?? 0;
  return `${rows.length} row${rows.length === 1 ? "" : "s"} x ${width} column${width === 1 ? "" : "s"}`;
}

function describeCell(actual: CellValue | undefined, expected: CellValue): string {
  if (isBlank(actual)) return "the cell is empty";
  if (isExcelError(actual)) return "the cell shows an Excel error";
  const want = expected === null ? "empty" : typeof expected === "string" ? "text" : typeof expected;
  const got = typeof actual === "string" ? "text" : typeof actual;
  return want === got ? `the ${want} differs` : `expected ${want}, got ${got}`;
}

function compareCell(actual: CellValue[][], expected: CellValue, tolerance: number): CompareResult {
  if (actual.length === 0) return { pass: false, detail: "Nothing was returned." };
  if (actual.length !== 1 || actual[0]!.length !== 1) {
    return { pass: false, detail: `Expected one value, got ${shape(actual)}.` };
  }
  const v = actual[0]![0];
  if (valuesMatch(v, expected, tolerance)) return { pass: true, detail: "Value matches." };
  return { pass: false, detail: `Value doesn't match: ${describeCell(v, expected)}.` };
}

function compareColumn(actual: CellValue[][], expected: readonly CellValue[], tolerance: number): CompareResult {
  if (actual.length === 0) return { pass: false, detail: "Nothing was returned." };
  const width = actual[0]!.length;
  if (width !== 1) return { pass: false, detail: `Expected one column, got ${shape(actual)}.` };
  let values = actual.map((r) => r[0] ?? null);
  // One leading label row is allowed, e.g. a formula placed on the header row with a VSTACK label.
  if (values.length === expected.length + 1 && typeof values[0] === "string" && !valuesMatch(values[0], expected[0] ?? null, tolerance)) {
    values = values.slice(1);
  }
  // A column whose last expected values are empty may come back shorter after trimming.
  while (values.length < expected.length && isBlank(expected[values.length])) values.push(null);
  if (values.length !== expected.length) {
    return { pass: false, detail: `Expected ${expected.length} values (one per data row), got ${values.length}.` };
  }
  const bad: number[] = [];
  let errors = 0;
  expected.forEach((e, i) => {
    if (!valuesMatch(values[i], e, tolerance)) {
      bad.push(i);
      if (isExcelError(values[i])) errors++;
    }
  });
  if (bad.length === 0) return { pass: true, detail: `All ${expected.length} values match.` };
  const first = bad[0]! + 1;
  const errNote = errors > 0 ? ` ${errors} of them show an Excel error.` : "";
  return {
    pass: false,
    detail: `${bad.length} of ${expected.length} values don't match; the first is data row ${first}.${errNote}`,
  };
}

function compareTable(actual: CellValue[][], expected: readonly (readonly CellValue[])[], tolerance: number): CompareResult {
  if (actual.length === 0) return { pass: false, detail: "Nothing was returned." };
  let rows = actual;
  if (rows.length > 0 && isHeaderRow(rows[0]!, expected, tolerance)) rows = rows.slice(1);
  rows = rows.filter((r) => !isTotalRow(r, expected));
  const width = expected[0]?.length ?? 0;
  const actualWidth = rows[0]?.length ?? 0;
  if (actualWidth !== width) {
    return { pass: false, detail: `Expected ${width} column${width === 1 ? "" : "s"}, got ${actualWidth}.` };
  }
  const used = new Array<boolean>(rows.length).fill(false);
  let missing = 0;
  for (const e of expected) {
    const i = rows.findIndex((r, k) => !used[k] && rowsMatch(r, e, tolerance));
    if (i === -1) missing++;
    else used[i] = true;
  }
  const extra = used.filter((u) => !u).length;
  if (missing === 0 && extra === 0) return { pass: true, detail: `All ${expected.length} rows match.` };
  const parts = [`${expected.length - missing} of ${expected.length} expected rows found`];
  if (extra > 0) parts.push(`${extra} extra row${extra === 1 ? "" : "s"}`);
  return { pass: false, detail: `${parts.join(", ")}.` };
}

/**
 * Compares what was read back from the scratch sheet with a task's check.
 * `actual` is the written range (a fill-down column or a spill), top-left first.
 * Trailing empty rows and columns are ignored. In table checks, one leading header row and any
 * "Total" / "Grand Total" row are ignored unless the expected rows contain them.
 */
export function compare(actual: readonly (readonly CellValue[])[], check: Check): CompareResult {
  const tolerance = check.tolerance ?? DEFAULT_TOLERANCE;
  const trimmed = trimActual(actual);
  switch (check.type) {
    case "cell":
      return compareCell(trimmed, check.expected, tolerance);
    case "column":
      return compareColumn(trimmed, check.expected, tolerance);
    case "table":
      return compareTable(trimmed, check.expected, tolerance);
    default:
      return { pass: false, detail: "Unknown check type." };
  }
}

// ---------------------------------------------------------------------------------------------
// Task validation, for the task files and the bench panel.

const LETTER_RE = /^[A-Z]{1,3}$/;
const RANGE_RE = /^[A-Z]{1,3}[1-9]\d*:[A-Z]{1,3}[1-9]\d*$/;

function isCellValue(v: unknown): v is CellValue {
  return v === null || typeof v === "string" || typeof v === "boolean" || (typeof v === "number" && Number.isFinite(v));
}

/** Problems with a parsed task file; empty when it is valid. */
export function validateTask(t: unknown): string[] {
  const p: string[] = [];
  if (typeof t !== "object" || t === null || Array.isArray(t)) return ["Task is not an object."];
  const o = t as Record<string, unknown>;
  const allowed = new Set(["id", "workbook", "sheet", "range", "category", "prompt", "private", "check"]);
  for (const k of Object.keys(o)) if (!allowed.has(k)) p.push(`Unknown key "${k}".`);
  for (const k of ["id", "workbook", "sheet", "range", "prompt"]) {
    if (typeof o[k] !== "string" || (o[k] as string).trim() === "") p.push(`"${k}" must be a non-empty string.`);
  }
  if (typeof o.workbook === "string" && !o.workbook.endsWith(".xlsx")) p.push(`"workbook" must name an .xlsx file.`);
  if (typeof o.range === "string" && !RANGE_RE.test(o.range)) p.push(`"range" must be a local address like A1:F501.`);
  if (!CATEGORIES.includes(o.category as Category)) p.push(`"category" must be one of ${CATEGORIES.join(", ")}.`);
  if (!Array.isArray(o.private) || !o.private.every((x) => typeof x === "string" && LETTER_RE.test(x))) {
    p.push(`"private" must be a list of column letters.`);
  }
  const c = o.check as Record<string, unknown> | undefined;
  if (typeof c !== "object" || c === null) {
    p.push(`"check" must be an object.`);
    return p;
  }
  for (const k of Object.keys(c)) if (!["type", "expected", "tolerance"].includes(k)) p.push(`Unknown check key "${k}".`);
  if (c.tolerance !== undefined && !(typeof c.tolerance === "number" && c.tolerance >= 0)) {
    p.push(`"check.tolerance" must be a number of 0 or more.`);
  }
  if (c.type === "cell") {
    if (!isCellValue(c.expected)) p.push(`A cell check expects one value.`);
  } else if (c.type === "column") {
    if (!Array.isArray(c.expected) || c.expected.length === 0 || !c.expected.every(isCellValue)) {
      p.push(`A column check expects a non-empty list of values.`);
    }
  } else if (c.type === "table") {
    const e = c.expected;
    if (!Array.isArray(e) || e.length === 0 || !e.every((r) => Array.isArray(r) && r.length > 0 && r.every(isCellValue))) {
      p.push(`A table check expects a non-empty list of non-empty rows.`);
    } else if (!e.every((r) => (r as unknown[]).length === (e[0] as unknown[]).length)) {
      p.push(`Table rows must all have the same width.`);
    }
  } else {
    p.push(`"check.type" must be cell, column or table.`);
  }
  return p;
}
