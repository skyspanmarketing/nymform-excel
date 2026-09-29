// A1-style address helpers shared by schema, payload, formulaGate and the adapter.
// Pure functions, no Office.js.

export const MAX_COLUMNS = 16384; // XFD
export const MAX_ROWS = 1048576;

/** 0-based column index to letters: 0 -> "A", 25 -> "Z", 26 -> "AA". */
export function columnLetter(index: number): string {
  if (!Number.isInteger(index) || index < 0 || index >= MAX_COLUMNS) {
    throw new RangeError(`Column index out of range: ${index}`);
  }
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/** Letters to 0-based column index: "A" -> 0, "AA" -> 26. Returns -1 when not a valid column. */
export function columnIndex(letters: string): number {
  if (!/^[A-Za-z]{1,3}$/.test(letters)) return -1;
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1 < MAX_COLUMNS ? n - 1 : -1;
}

/** A rectangle in 0-based, inclusive coordinates. */
export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** Parses a cell like "B7" or "$B$7" into 0-based row and column, or null. */
export function parseCell(ref: string): { r: number; c: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(ref);
  if (!m) return null;
  const c = columnIndex(m[1]!);
  const r = Number(m[2]) - 1;
  if (c < 0 || r < 0 || r >= MAX_ROWS) return null;
  return { r, c };
}

/**
 * Parses a local range address ("A1", "A1:F501", "$A$1:$F$501") into a rectangle.
 * Returns null for anything else (whole columns, rows, names).
 */
export function parseLocalRange(address: string): Rect | null {
  const parts = address.split(":");
  if (parts.length === 1) {
    const a = parseCell(parts[0]!);
    return a ? { r1: a.r, c1: a.c, r2: a.r, c2: a.c } : null;
  }
  if (parts.length !== 2) return null;
  const a = parseCell(parts[0]!);
  const b = parseCell(parts[1]!);
  if (!a || !b) return null;
  return {
    r1: Math.min(a.r, b.r),
    c1: Math.min(a.c, b.c),
    r2: Math.max(a.r, b.r),
    c2: Math.max(a.c, b.c),
  };
}

/** Quotes a sheet name for use in a reference when needed: "My Sheet" -> "'My Sheet'". */
export function quoteSheet(sheet: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) && !parseCell(sheet)) return sheet;
  return `'${sheet.replace(/'/g, "''")}'`;
}

/** Splits "Sheet!A1:B2" or "'My Sheet'!A1" into sheet and local address. */
export function splitSheetAddress(full: string): { sheet: string | null; address: string } {
  const bang = full.lastIndexOf("!");
  if (bang < 0) return { sheet: null, address: full };
  let sheet = full.slice(0, bang);
  if (sheet.startsWith("'") && sheet.endsWith("'") && sheet.length >= 2) {
    sheet = sheet.slice(1, -1).replace(/''/g, "'");
  }
  return { sheet, address: full.slice(bang + 1) };
}

/** Formats a rectangle as a local address, e.g. "A1:F501" (or "A1" for one cell). */
export function formatRect(rect: Rect): string {
  const a = `${columnLetter(rect.c1)}${rect.r1 + 1}`;
  const b = `${columnLetter(rect.c2)}${rect.r2 + 1}`;
  return a === b ? a : `${a}:${b}`;
}

/** "Orders" + "A1:F501" -> "Orders!A1:F501" (quoted when needed). */
export function qualify(sheet: string, address: string): string {
  return `${quoteSheet(sheet)}!${address}`;
}
