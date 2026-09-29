// Converts an .xlsx into the workbook JSON the fake Excel host reads (tests/e2e/fakeOffice.js).
// Display text is an approximation of Excel's: General numbers, #,##0.00-style number formats,
// percent, dates and times by their number format (m/d/yyyy, h:mm AM/PM, ...), TRUE/FALSE.
//
//   npx tsx tests/e2e/workbookToFake.ts <input.xlsx> [output.json] [--select=Sheet!A1:F501]
//
// Without an output path the JSON goes to stdout.
import { writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";
import ExcelJS from "exceljs";

export type FakeType = "String" | "Double" | "Boolean" | "Error" | "Empty";

export interface FakeCell {
  value: string | number | boolean | null;
  text: string;
  type: FakeType;
  format: string;
  formula?: string;
}

export interface FakeSheet {
  name: string;
  visibility: "Visible" | "Hidden" | "VeryHidden";
  cells: Record<string, FakeCell>;
  /** The sheet's AutoFilter range, e.g. "A1:F501". */
  autoFilter?: string;
}

export interface FakeWorkbook {
  sheets: FakeSheet[];
  /** `address` includes the totals row when `showTotals` is set, as in Excel. */
  tables: { name: string; sheet: string; address: string; showTotals?: boolean }[];
  selection: { sheet: string; address: string };
  calculationMode: "Automatic" | "AutomaticExceptTables" | "Manual";
}

const MS_PER_DAY = 86_400_000;
const EPOCH_OFFSET_DAYS = 25_569; // serial of 1970-01-01 with Excel's 1899-12-30 epoch

/** A JS Date (exceljs reads date-formatted cells as UTC dates) back to an Excel serial. */
export function dateToSerial(d: Date): number {
  const ms = d.getTime() + EPOCH_OFFSET_DAYS * MS_PER_DAY;
  const days = Math.floor(ms / MS_PER_DAY);
  const rem = ms - days * MS_PER_DAY;
  if (rem === 0) return days;
  // Whole seconds divide by 86400, which gives the same double as minutes / 1440.
  return rem % 1000 === 0 ? days + rem / 1000 / 86_400 : days + rem / MS_PER_DAY;
}

// ---------------------------------------------------------------------------------------------
// Display text

function grouped(n: number, decimals: number, thousands: boolean): string {
  const fixed = Math.abs(n).toFixed(decimals);
  const [int = "0", frac] = fixed.split(".");
  const body = thousands ? int.replace(/\B(?=(\d{3})+(?!\d))/g, ",") : int;
  return body + (frac !== undefined ? `.${frac}` : "");
}

export function generalText(n: number): string {
  if (!Number.isFinite(n)) return "#NUM!";
  if (Number.isInteger(n) && Math.abs(n) < 1e11) return String(n);
  if (Math.abs(n) >= 1e11 || (n !== 0 && Math.abs(n) < 1e-9)) return n.toExponential(5).replace("e", "E");
  return String(parseFloat(n.toPrecision(10)));
}

/** Removes quoted literals, escapes and [...] codes, for format classification. */
function bareFormat(section: string): string {
  return section
    .replace(/"[^"]*"/g, "")
    .replace(/\\./g, "")
    .replace(/\[[^\]]*\]/g, "")
    .replace(/_.|\*./g, "");
}

function sections(format: string): string[] {
  const out: string[] = [];
  let cur = "";
  let quoted = false;
  for (let i = 0; i < format.length; i++) {
    const ch = format[i]!;
    if (ch === '"') quoted = !quoted;
    if (ch === "\\" && !quoted) {
      cur += ch + (format[i + 1] ?? "");
      i++;
      continue;
    }
    if (ch === ";" && !quoted) {
      out.push(cur);
      cur = "";
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out;
}

export function isDateFormat(format: string): boolean {
  const bare = bareFormat(sections(format)[0] ?? "");
  return /[dmyhs]/i.test(bare) && !/^general$/i.test(bare.trim());
}

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

function formatDate(serial: number, section: string): string {
  const days = Math.floor(serial);
  let secs = Math.round((serial - days) * 86_400);
  let dayNum = days;
  if (secs >= 86_400) {
    secs -= 86_400;
    dayNum += 1;
  }
  const dt = new Date((dayNum - EPOCH_OFFSET_DAYS) * MS_PER_DAY);
  const y = dt.getUTCFullYear();
  const mo = dt.getUTCMonth();
  const d = dt.getUTCDate();
  const wd = dt.getUTCDay();
  const h = Math.floor(secs / 3600);
  const mi = Math.floor((secs % 3600) / 60);
  const s = secs % 60;
  const ampm = /AM\/PM|A\/P/i.test(section);
  const pad = (n: number, w: number) => String(n).padStart(w, "0");

  // Tokenize first so "m" can be told apart as month or minute.
  type Tok = { kind: "lit" | "y" | "m" | "d" | "h" | "s" | "ampm" | "ap"; text: string };
  const toks: Tok[] = [];
  for (let i = 0; i < section.length; ) {
    const rest = section.slice(i);
    const ch = section[i]!;
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      const stop = end < 0 ? section.length : end;
      toks.push({ kind: "lit", text: section.slice(i + 1, stop) });
      i = stop + 1;
    } else if (ch === "\\") {
      toks.push({ kind: "lit", text: section[i + 1] ?? "" });
      i += 2;
    } else if (ch === "[") {
      const end = section.indexOf("]", i);
      i = end < 0 ? section.length : end + 1;
    } else if (/^AM\/PM/i.test(rest)) {
      toks.push({ kind: "ampm", text: rest.slice(0, 5) });
      i += 5;
    } else if (/^A\/P/i.test(rest)) {
      toks.push({ kind: "ap", text: rest.slice(0, 3) });
      i += 3;
    } else if (/[ymdhs]/i.test(ch)) {
      const m = new RegExp(`^${ch}+`, "i").exec(rest)!;
      toks.push({ kind: ch.toLowerCase() as Tok["kind"], text: m[0] });
      i += m[0].length;
    } else if (ch === "_" || ch === "*") {
      i += 2;
    } else {
      toks.push({ kind: "lit", text: ch });
      i += 1;
    }
  }
  const prevNonLit = (j: number) => toks.slice(0, j).reverse().find((t) => t.kind !== "lit");
  const nextNonLit = (j: number) => toks.slice(j + 1).find((t) => t.kind !== "lit");

  let out = "";
  toks.forEach((t, j) => {
    const n = t.text.length;
    switch (t.kind) {
      case "lit":
        out += t.text;
        break;
      case "y":
        out += n <= 2 ? pad(y % 100, 2) : String(y);
        break;
      case "m": {
        const minute = n <= 2 && (prevNonLit(j)?.kind === "h" || nextNonLit(j)?.kind === "s");
        if (minute) out += n === 2 ? pad(mi, 2) : String(mi);
        else if (n === 1) out += String(mo + 1);
        else if (n === 2) out += pad(mo + 1, 2);
        else if (n === 3) out += MONTHS[mo]!.slice(0, 3);
        else if (n === 5) out += MONTHS[mo]!.slice(0, 1);
        else out += MONTHS[mo]!;
        break;
      }
      case "d":
        if (n === 1) out += String(d);
        else if (n === 2) out += pad(d, 2);
        else if (n === 3) out += DAYS[wd]!.slice(0, 3);
        else out += DAYS[wd]!;
        break;
      case "h": {
        const hh = ampm ? (h % 12 === 0 ? 12 : h % 12) : h;
        out += n >= 2 ? pad(hh, 2) : String(hh);
        break;
      }
      case "s":
        out += n >= 2 ? pad(s, 2) : String(s);
        break;
      case "ampm": {
        const upper = t.text[0] === "A";
        const v = h < 12 ? "AM" : "PM";
        out += upper ? v : v.toLowerCase();
        break;
      }
      case "ap":
        out += h < 12 ? (t.text[0] === "A" ? "A" : "a") : t.text[0] === "A" ? "P" : "p";
        break;
    }
  });
  return out;
}

function formatNumber(n: number, section: string, negativeSection: boolean): string {
  // Literal prefix/suffix around the digit placeholders; "$#,##0.00" keeps its "$".
  let literal = "";
  let pattern = "";
  const parts: { lit: string; pat: string }[] = [];
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]!;
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      const stop = end < 0 ? section.length : end;
      literal += section.slice(i + 1, stop);
      i = stop;
    } else if (ch === "\\") {
      literal += section[i + 1] ?? "";
      i++;
    } else if (ch === "[") {
      const end = section.indexOf("]", i);
      i = end < 0 ? section.length : end;
    } else if (ch === "_" || ch === "*") {
      i++;
    } else if (/[0#?.,]/.test(ch) || (/[eE]/.test(ch) && pattern !== "") || (/[+-]/.test(ch) && /[eE]$/.test(pattern))) {
      if (literal) parts.push({ lit: literal, pat: "" });
      literal = "";
      pattern += ch;
    } else {
      if (pattern) parts.push({ lit: "", pat: pattern });
      pattern = "";
      literal += ch;
    }
  }
  if (pattern) parts.push({ lit: "", pat: pattern });
  if (literal) parts.push({ lit: literal, pat: "" });

  const pat = parts.map((p) => p.pat).join("");
  const percent = parts.some((p) => p.lit.includes("%"));
  let v = percent ? n * 100 : n;
  if (negativeSection) v = Math.abs(v);
  let digits: string;
  const sci = /E[+-]?/i.exec(pat);
  if (sci) {
    const mant = pat.slice(0, sci.index);
    const dec = (mant.split(".")[1] ?? "").replace(/[^0#?]/g, "").length;
    const [m = "0", e = "0"] = Math.abs(v).toExponential(dec).split("e");
    const exp = Number(e);
    digits = `${m}E${exp < 0 ? "-" : "+"}${String(Math.abs(exp)).padStart(2, "0")}`;
  } else {
    const [intPat = "", decPat = ""] = pat.split(".");
    // Trailing commas scale by thousands.
    let scale = 0;
    let ip = intPat;
    while (ip.endsWith(",")) {
      scale++;
      ip = ip.slice(0, -1);
    }
    v = v / 1000 ** scale;
    const decimals = decPat.replace(/[^0#?]/g, "").length;
    const minInt = (ip.match(/0/g) ?? []).length;
    digits = grouped(v, decimals, ip.includes(","));
    if (!pat.includes("0") && !pat.includes("?") && digits === "0") digits = "";
    if (minInt === 0 && digits.startsWith("0.")) digits = digits.slice(1);
    // Optional decimals ("0.##") drop trailing zeros.
    if (decimals > 0 && /#/.test(decPat) && !/0/.test(decPat)) digits = digits.replace(/\.?0+$/, "");
  }
  const sign = !negativeSection && v < 0 ? "-" : "";
  let out = "";
  let placed = false;
  for (const p of parts) {
    if (p.pat) {
      if (!placed) out += digits.replace(/^-/, "");
      placed = true;
    } else out += p.lit;
  }
  return sign + (placed ? out : out + digits);
}

/** The text Excel would show for `value` under `format` (an approximation). */
export function displayText(value: string | number | boolean | null, format = "General"): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "string") return value;
  const secs = sections(format);
  if (!format || /^general$/i.test(format.trim())) return generalText(value);
  if (isDateFormat(format)) {
    if (value < 0) return "#".repeat(10);
    return formatDate(value, secs[0] ?? format);
  }
  const neg = value < 0 && secs.length >= 2;
  const zero = value === 0 && secs.length >= 3;
  const section = zero ? secs[2]! : neg ? secs[1]! : secs[0]!;
  if (/^general$/i.test(bareFormat(section).trim())) return generalText(neg ? -value : value);
  if (bareFormat(section).trim() === "@") return generalText(value);
  return formatNumber(value, section, neg);
}

// ---------------------------------------------------------------------------------------------
// Conversion

type ExcelValue = ExcelJS.CellValue;

function plainOf(v: ExcelValue): { value: string | number | boolean | null; type: FakeType } {
  if (v === null || v === undefined) return { value: null, type: "Empty" };
  if (typeof v === "number") return { value: v, type: "Double" };
  if (typeof v === "string") return { value: v, type: v === "" ? "Empty" : "String" };
  if (typeof v === "boolean") return { value: v, type: "Boolean" };
  if (v instanceof Date) return { value: dateToSerial(v), type: "Double" };
  if (typeof v === "object") {
    if ("error" in v && typeof v.error === "string") return { value: v.error, type: "Error" };
    if ("richText" in v && Array.isArray(v.richText)) {
      return { value: v.richText.map((r) => r.text).join(""), type: "String" };
    }
    if ("text" in v && "hyperlink" in v) {
      const t = v.text;
      const text = typeof t === "string" ? t : Array.isArray((t as { richText?: unknown }).richText) ? (t as ExcelJS.CellRichTextValue).richText.map((r) => r.text).join("") : "";
      return { value: text, type: "String" };
    }
  }
  return { value: String(v), type: "String" };
}

function cellOf(cell: ExcelJS.Cell): FakeCell | null {
  const format = typeof cell.numFmt === "string" && cell.numFmt !== "" ? cell.numFmt : "General";
  const raw = cell.value;
  let formula: string | undefined;
  let plain: { value: string | number | boolean | null; type: FakeType };
  if (raw !== null && typeof raw === "object" && !(raw instanceof Date) && ("formula" in raw || "sharedFormula" in raw)) {
    formula = cell.formula ? `=${cell.formula}` : undefined;
    const result = (raw as { result?: ExcelValue }).result;
    plain = plainOf(result === undefined ? null : result);
    // No cached result: the fake host shows 0 for formulas it can't evaluate.
    if (plain.type === "Empty") plain = { value: 0, type: "Double" };
  } else {
    plain = plainOf(raw);
  }
  if (plain.type === "Empty" && !formula) {
    return format === "General" ? null : { value: null, text: "", type: "Empty", format };
  }
  const out: FakeCell = { value: plain.value, text: displayText(plain.value, format), type: plain.type, format };
  if (formula) out.formula = formula;
  return out;
}

const VISIBILITY: Record<string, FakeSheet["visibility"]> = {
  visible: "Visible",
  hidden: "Hidden",
  veryHidden: "VeryHidden",
};

export interface ConvertOptions {
  /** Initial selection, e.g. { sheet: "Orders", address: "A1:F501" }. Default: A1 of the first visible sheet. */
  selection?: { sheet: string; address: string };
}

/** ExcelJS reads a sheet's AutoFilter as its ref ("A1:F501"); a written one may be { from, to }. */
function autoFilterOf(filter: ExcelJS.AutoFilter | undefined): string | null {
  if (!filter) return null;
  if (typeof filter === "string") return filter;
  const cell = (a: string | { row: number; column: number }) =>
    typeof a === "string" ? a : `${columnName(a.column)}${a.row}`;
  return `${cell(filter.from)}:${cell(filter.to)}`;
}

/** 1-based column number to letters, as ExcelJS numbers columns: 1 -> "A". */
function columnName(n: number): string {
  let out = "";
  for (let x = n; x > 0; x = Math.floor((x - 1) / 26)) out = String.fromCharCode(65 + ((x - 1) % 26)) + out;
  return out;
}

export function workbookToFake(book: ExcelJS.Workbook, opts: ConvertOptions = {}): FakeWorkbook {
  const sheets: FakeSheet[] = [];
  const tables: FakeWorkbook["tables"] = [];
  for (const ws of book.worksheets) {
    const cells: Record<string, FakeCell> = {};
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        const c = cellOf(cell);
        if (c) cells[cell.address] = c;
      });
    });
    const sheet: FakeSheet = { name: ws.name, visibility: VISIBILITY[ws.state] ?? "Visible", cells };
    const filter = autoFilterOf(ws.autoFilter);
    if (filter) sheet.autoFilter = filter;
    sheets.push(sheet);
    const wsTables =
      (ws as unknown as { tables?: Record<string, { table?: { name?: string; tableRef?: string; ref?: string; totalsRow?: boolean } }> })
        .tables ?? {};
    for (const [key, t] of Object.entries(wsTables)) {
      const model = t.table ?? {};
      const address = model.tableRef ?? model.ref;
      if (!address) continue;
      const table: FakeWorkbook["tables"][number] = { name: model.name ?? key, sheet: ws.name, address };
      if (model.totalsRow) table.showTotals = true;
      tables.push(table);
    }
  }
  const first = sheets.find((s) => s.visibility === "Visible") ?? sheets[0];
  const selection = opts.selection ?? { sheet: first?.name ?? "Sheet1", address: "A1" };
  return { sheets, tables, selection, calculationMode: "Automatic" };
}

export async function xlsxToFake(input: string | Buffer | ArrayBuffer, opts: ConvertOptions = {}): Promise<FakeWorkbook> {
  const book = new ExcelJS.Workbook();
  if (typeof input === "string") await book.xlsx.readFile(input);
  else await book.xlsx.load(input as ArrayBuffer);
  return workbookToFake(book, opts);
}

/** JSON with one cell per line, so fixture diffs stay readable. */
export function stringifyFake(wb: FakeWorkbook): string {
  const lines: string[] = ["{", '  "sheets": ['];
  wb.sheets.forEach((s, i) => {
    const filter = s.autoFilter ? `"autoFilter": ${JSON.stringify(s.autoFilter)}, ` : "";
    lines.push(`    { "name": ${JSON.stringify(s.name)}, "visibility": ${JSON.stringify(s.visibility)}, ${filter}"cells": {`);
    const keys = Object.keys(s.cells);
    keys.forEach((k, j) => lines.push(`      ${JSON.stringify(k)}: ${JSON.stringify(s.cells[k])}${j < keys.length - 1 ? "," : ""}`));
    lines.push(`    } }${i < wb.sheets.length - 1 ? "," : ""}`);
  });
  lines.push("  ],");
  lines.push(`  "tables": ${JSON.stringify(wb.tables)},`);
  lines.push(`  "selection": ${JSON.stringify(wb.selection)},`);
  lines.push(`  "calculationMode": ${JSON.stringify(wb.calculationMode)}`);
  lines.push("}");
  return lines.join("\n") + "\n";
}

async function main(argv: string[]): Promise<void> {
  const flags = argv.filter((a) => a.startsWith("--"));
  const args = argv.filter((a) => !a.startsWith("--"));
  const [input, output] = args;
  if (!input) {
    process.stderr.write("Usage: npx tsx tests/e2e/workbookToFake.ts <input.xlsx> [output.json] [--select=Sheet!A1:F501]\n");
    process.exitCode = 2;
    return;
  }
  let selection: ConvertOptions["selection"];
  const sel = flags.find((f) => f.startsWith("--select="))?.slice("--select=".length);
  if (sel) {
    const bang = sel.lastIndexOf("!");
    if (bang < 0) throw new Error("--select needs Sheet!Address");
    let sheet = sel.slice(0, bang);
    if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
    selection = { sheet, address: sel.slice(bang + 1) };
  }
  const json = stringifyFake(await xlsxToFake(resolve(input), { selection }));
  if (output) {
    writeFileSync(resolve(output), json);
    return;
  }
  // Piping into `head` closes stdout early; that isn't an error.
  process.stdout.on("error", (e: NodeJS.ErrnoException) => {
    if (e.code === "EPIPE") process.exit(0);
    throw e;
  });
  process.stdout.write(json);
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  main(process.argv.slice(2)).catch((e: unknown) => {
    process.stderr.write(`${e instanceof Error ? e.message : String(e)}\n`);
    process.exitCode = 1;
  });
}
