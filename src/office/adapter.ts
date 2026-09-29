// The only file that touches Office.js (spec §7.1). Excel JavaScript API, requirement set
// ExcelApi 1.9; spill readback uses 1.12 when the host has it. Everything else in the add-in gets
// plain data from here, and every Office error leaves this file as a plain Error message.
import {
  columnLetter,
  formatRect,
  MAX_COLUMNS,
  MAX_ROWS,
  parseCell,
  parseLocalRange,
  splitSheetAddress,
  type Rect,
} from "../core/a1";
import type { CellValue, InsertCheck, RangeData, Resolution, ResolutionKind, TableInfo } from "../core/types";
import { BENCH } from "../config";

// Declared in core/types.ts, which can't import from here; RangeData carries it.
export type { Resolution, ResolutionKind };

export const CELL_LIMIT = 20000;
export const SCRATCH_SHEET = "_nymform_scratch";
export const BENCH_SHEET = "_nymform_bench";
/** Requirement set the add-in needs (getTables, autoFill). */
export const REQUIRED_API = "1.9";
/** Requirement set for spill readback (getSpillingToRangeOrNullObject). */
export const SPILL_API = "1.12";

export const TOO_LARGE_MESSAGE = "Select fewer cells (limit 20,000).";
export const NOT_IN_EXCEL_MESSAGE = "Nymform runs inside Excel. Open it from Excel's ribbon.";
/** readSelection found nothing it can read as data: no cells in use, or fewer than two rows. */
export const RESOLVE_MESSAGE = "Select your data, or one cell in it, and choose Refresh from selection.";
/** readRange was given a sheet Excel doesn't have. */
export const NO_SUCH_SHEET_MESSAGE = "There's no sheet with that name. Check the sheet name in the address.";

/** "The data around E4 has 28,000 cells. Select fewer cells (limit 20,000)." */
export function tooLargeAroundMessage(from: string, cellCount: number): string {
  return `The data around ${from} has ${cellCount.toLocaleString("en-US")} cells. ${TOO_LARGE_MESSAGE}`;
}

export class SelectionTooLargeError extends Error {
  readonly cellCount: number;
  constructor(cellCount = CELL_LIMIT + 1, message = TOO_LARGE_MESSAGE) {
    super(message);
    this.name = "SelectionTooLargeError";
    this.cellCount = cellCount;
  }
}

/** An Office error turned into a plain sentence. `code` is Office's error code, when there was one. */
export class ExcelHostError extends Error {
  readonly code: string | undefined;
  constructor(message: string, code?: string) {
    super(message);
    this.name = "ExcelHostError";
    this.code = code;
    // The UI shows messages only; keep Office internals out of what it could print.
    this.stack = `${this.name}: ${message}`;
  }
}

const CODE_MESSAGES: Record<string, string> = {
  ItemNotFound: "Excel couldn't find that sheet, table or range. Check the name and try again.",
  InvalidArgument: "Excel didn't accept that address or formula. Check it and try again.",
  InvalidReference: "Excel didn't accept that reference. Check it and try again.",
  InvalidSelection: "Select one rectangular range and try again.",
  InvalidOperation: "Excel couldn't do that right now. Try again.",
  InvalidOperationInCellEditMode: "Finish editing the cell in Excel (press Enter or Esc), then try again.",
  AccessDenied: "Excel didn't allow the change. The sheet or workbook may be protected.",
  ItemAlreadyExists: "A sheet with that name already exists.",
  RequestPayloadSizeLimitExceeded: "That range is too large for Excel to send. Select fewer cells.",
  ResponsePayloadSizeLimitExceeded: "That range is too large for Excel to send. Select fewer cells.",
  ActivityLimitReached: "Excel is busy. Wait a moment and try again.",
  RequestAborted: "Excel stopped the request. Try again.",
  ApiNotFound: "This version of Excel doesn't support that. Update Excel and try again.",
  Conflict: "Someone else is editing that part of the workbook. Try again in a moment.",
  GeneralException: "Excel couldn't finish that. Try again.",
};

/**
 * Converts anything thrown inside Excel.run into a plain Error with a short message. The original
 * message is not passed on, so nothing from the workbook can reach the UI through it; the Office
 * error code is kept.
 */
export function toPlainError(e: unknown): Error {
  if (e instanceof SelectionTooLargeError || e instanceof ExcelHostError) return e;
  const raw = typeof e === "object" && e !== null && "code" in e ? (e as { code: unknown }).code : undefined;
  const code = typeof raw === "string" && /^[A-Za-z]{1,64}$/.test(raw) ? raw : undefined;
  if (code !== undefined && CODE_MESSAGES[code]) return new ExcelHostError(CODE_MESSAGES[code]!, code);
  return new ExcelHostError(code ? `Excel reported a problem (${code}). Try again.` : "Excel reported a problem. Try again.", code);
}

function hasOffice(): boolean {
  return typeof Office !== "undefined" && typeof Office.onReady === "function";
}

function isSetSupported(version: string): boolean {
  try {
    return typeof Office !== "undefined" && Office.context?.requirements?.isSetSupported("ExcelApi", version) === true;
  } catch {
    return false;
  }
}

async function run<T>(batch: (ctx: Excel.RequestContext) => Promise<T>): Promise<T> {
  if (typeof Excel === "undefined" || typeof Excel.run !== "function") throw new ExcelHostError(NOT_IN_EXCEL_MESSAGE);
  try {
    return await Excel.run(batch);
  } catch (e) {
    throw toPlainError(e);
  }
}

/** Resolves once Office.js is ready. Rejects outside Excel, or when Excel is too old. */
export async function ready(timeoutMs = 15000): Promise<{ host: string; platform: string }> {
  if (!hasOffice()) throw new ExcelHostError(NOT_IN_EXCEL_MESSAGE);
  let timer: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new ExcelHostError(NOT_IN_EXCEL_MESSAGE)), timeoutMs);
  });
  let info: { host: Office.HostType | null; platform: Office.PlatformType | null };
  try {
    info = await Promise.race([Office.onReady(), timeout]);
  } catch (e) {
    throw toPlainError(e);
  } finally {
    clearTimeout(timer);
  }
  if (!info || info.host !== Office.HostType.Excel) throw new ExcelHostError(NOT_IN_EXCEL_MESSAGE);
  if (!isSetSupported(REQUIRED_API)) {
    throw new ExcelHostError("This version of Excel is too old for Nymform. Update Excel and try again.");
  }
  return { host: String(info.host), platform: String(info.platform ?? "") };
}

/** Office's theme, for the task pane's light or dark look. Null when the host doesn't say. */
export function officeTheme(): { isDarkTheme?: boolean; bodyBackgroundColor?: string } | null {
  try {
    const theme: unknown = typeof Office !== "undefined" ? Office.context?.officeTheme : undefined;
    if (theme === null || typeof theme !== "object") return null;
    const { isDarkTheme, bodyBackgroundColor } = theme as Record<string, unknown>;
    const out: { isDarkTheme?: boolean; bodyBackgroundColor?: string } = {};
    if (typeof isDarkTheme === "boolean") out.isDarkTheme = isDarkTheme;
    if (typeof bodyBackgroundColor === "string") out.bodyBackgroundColor = bodyBackgroundColor;
    return out;
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------------------------
// Reading

function cellValue(v: unknown, type: string | undefined, text: string | undefined): CellValue {
  if (type === "Empty") return null;
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return v;
  if (v === null || v === undefined) return null;
  return text ?? String(v);
}

function formulaValue(v: unknown): string | number | boolean | null {
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return v;
  if (v === null || v === undefined) return null;
  return String(v);
}

/** Loaded on a range to know where it is and how large, without any cell content. */
const GEOMETRY = "rowIndex,columnIndex,rowCount,columnCount";

type Geometry = Pick<Excel.Range, "rowIndex" | "columnIndex" | "rowCount" | "columnCount">;

function rectOf(r: Geometry): Rect {
  return { r1: r.rowIndex, c1: r.columnIndex, r2: r.rowIndex + r.rowCount - 1, c2: r.columnIndex + r.columnCount - 1 };
}

function intersect(a: Rect, b: Rect): Rect | null {
  const r = { r1: Math.max(a.r1, b.r1), c1: Math.max(a.c1, b.c1), r2: Math.min(a.r2, b.r2), c2: Math.min(a.c2, b.c2) };
  return r.r1 <= r.r2 && r.c1 <= r.c2 ? r : null;
}

function holdsCell(rect: Rect, r: number, c: number): boolean {
  return r >= rect.r1 && r <= rect.r2 && c >= rect.c1 && c <= rect.c2;
}

/** Counted from the sides: Office reports cellCount as -1 past 2^31-1 cells (a whole sheet). */
function cellsIn(rect: Rect): number {
  return (rect.r2 - rect.r1 + 1) * (rect.c2 - rect.c1 + 1);
}

/**
 * Reads one range into plain data. Checks the cell count before loading any cell content, a table's
 * column names included (they are its header cells). A caller that already knows the count, or the
 * table the range is, passes it and saves a round trip to Excel each.
 */
async function readRangeData(
  ctx: Excel.RequestContext,
  range: Excel.Range,
  known: { count?: number; table?: Excel.Table } = {},
): Promise<RangeData> {
  let count = known.count;
  if (count === undefined) {
    range.load("rowCount,columnCount");
    await ctx.sync();
    count = range.rowCount * range.columnCount;
  }
  if (count > CELL_LIMIT) throw new SelectionTooLargeError(count);

  range.load("address,values,text,formulas,valueTypes,numberFormat,rowIndex,columnIndex,columnCount");
  const sheet = range.worksheet;
  sheet.load("name");
  const tables = known.table ? null : range.getTables(false);
  tables?.load("items/name");
  // The table's name was loaded with its position; its column names come only now, after the count.
  known.table?.columns.load("items/name");
  await ctx.sync();

  const valueTypes = range.valueTypes.map((row) => row.map((t) => String(t)));
  const text = range.text.map((row) => row.map((t) => (t === null || t === undefined ? "" : String(t))));
  const values = range.values.map((row, r) => row.map((v, c) => cellValue(v, valueTypes[r]?.[c], text[r]?.[c])));
  const formulas = range.formulas.map((row) => row.map(formulaValue));
  const numberFormat = range.numberFormat.map((row) => row.map((f) => (typeof f === "string" ? f : "General")));

  const data: RangeData = {
    sheet: sheet.name,
    address: splitSheetAddress(range.address).address,
    rowIndex: range.rowIndex,
    columnIndex: range.columnIndex,
    values,
    text,
    formulas,
    valueTypes,
    numberFormat,
    // A table's range starts at its first column, so its column names line up as they are.
    table: known.table ? { name: known.table.name, columns: known.table.columns.items.map((c) => c.name) } : null,
  };

  if (tables && tables.items.length > 0) {
    const found = tables.items.map((t) => {
      const r = t.getRange();
      r.load(GEOMETRY);
      t.columns.load("items/name");
      return { t, r };
    });
    await ctx.sync();
    // Prefer the table holding the selection's top-left cell.
    const pick = found.find(({ r }) => holdsCell(rectOf(r), range.rowIndex, range.columnIndex)) ?? found[0]!;
    const names = pick.t.columns.items.map((c) => c.name);
    // Aligned with the selection's columns: "" where a selected column lies outside the table.
    const columns = Array.from(
      { length: range.columnCount },
      (_, i) => names[range.columnIndex + i - pick.r.columnIndex] ?? "",
    );
    data.table = { name: pick.t.name, columns };
  }
  return data;
}

function resolveError(): ExcelHostError {
  return new ExcelHostError(RESOLVE_MESSAGE);
}

/**
 * What a pointer stands for (spec §7.1), from its anchor (one cell's own cell; a row's first filled
 * cell): the Excel table holding the anchor (else the only table where the row holds something)
 * without the totals row; else the sheet's AutoFilter range when it holds the anchor, cut to its
 * filled cells; else the block of filled cells around the anchor (Excel's current region), cut to
 * the cells in use. Loads positions, sizes and table names, and for a row the value
 * types of that one row (empty or not, never the values) to find where it holds anything; a table's
 * column names are its header cells, loaded by readRangeData after the size check. Null when there
 * is no block.
 */
async function expandPointer(
  ctx: Excel.RequestContext,
  ws: Excel.Worksheet,
  pointer: Rect,
  trimmed: boolean,
): Promise<{ kind: ResolutionKind; rect: Rect; table?: Excel.Table } | null> {
  const range = ws.getRange(formatRect(pointer));
  // A row's first filled cell anchors the region. The trimmed row starts there when Excel's used
  // range of a range is the box around its filled cells; when it is the range cut to the sheet's
  // used range instead, the row can start on an empty cell. Types only, never values.
  const multi = pointer.c2 > pointer.c1;
  if (multi) range.load("valueTypes");
  const tables = range.getTables(false);
  tables.load("items/name,items/showTotals");
  const filter = ws.autoFilter.getRangeOrNullObject();
  filter.load(GEOMETRY);
  // Asked for in the same round trip, so a pointer outside any table and filter needs no further
  // one. The region is Excel's for the top-left cell; around an empty cell next to data it includes
  // that cell's empty row or column, which its own used range cuts off.
  let region = range.getSurroundingRegion().getUsedRangeOrNullObject(true);
  region.load(GEOMETRY);
  await ctx.sync();

  let anchor = { r: pointer.r1, c: pointer.c1 };
  const types = multi ? (range.valueTypes[0] ?? []).map(String) : [];
  if (multi) {
    const first = types.findIndex((t) => t !== "Empty");
    // A selected row with nothing in it stands for the data around it, as one empty cell does; a
    // whole row with nothing in the data's columns doesn't.
    if (first === -1 && trimmed) return null;
    if (first > 0) {
      // The row starts on an empty cell: grow the region around its first filled cell instead.
      anchor = { r: pointer.r1, c: pointer.c1 + first };
      region = ws.getRange(formatRect({ r1: anchor.r, c1: anchor.c, r2: anchor.r, c2: anchor.c })).getSurroundingRegion().getUsedRangeOrNullObject(true);
      region.load(GEOMETRY);
    }
  }

  const found = tables.items.map((t) => {
    const r = t.getRange();
    r.load(GEOMETRY);
    return { t, r };
  });
  // Excel reports the range the filter was set on, which can be whole columns (A:F) or run far past
  // the data; only its filled cells are data. Read from a range of its own, not from the null object.
  const filterUsed =
    !filter.isNullObject && holdsCell(rectOf(filter), anchor.r, anchor.c)
      ? ws.getRange(formatRect(rectOf(filter))).getUsedRangeOrNullObject(true)
      : null;
  filterUsed?.load(GEOMETRY);
  if (found.length > 0 || filterUsed || anchor.c !== pointer.c1) await ctx.sync();

  // A table counts when it holds the anchor, or when the row holds something inside the table's
  // columns: a table the row crosses where it is blank isn't the data, wherever it lies along the row.
  const filledIn = (r: Rect): boolean => {
    if (!multi || r.r1 > pointer.r1 || pointer.r1 > r.r2) return false;
    for (let c = Math.max(r.c1, pointer.c1); c <= Math.min(r.c2, pointer.c2); c++) if (types[c - pointer.c1] !== "Empty") return true;
    return false;
  };
  const touched = found.filter(({ r }) => holdsCell(rectOf(r), anchor.r, anchor.c) || filledIn(rectOf(r)));
  const pick = touched.find(({ r }) => holdsCell(rectOf(r), anchor.r, anchor.c)) ?? (touched.length === 1 ? touched[0] : undefined);
  if (pick) {
    const whole = rectOf(pick.r);
    // The totals row summarizes the rows above it; it isn't data.
    const rect = pick.t.showTotals && whole.r2 > whole.r1 ? { ...whole, r2: whole.r2 - 1 } : whole;
    return { kind: "table", rect, table: pick.t };
  }
  if (filterUsed && !filterUsed.isNullObject) {
    const rect = rectOf(filterUsed);
    // A filter over a header row alone says nothing about where the data ends.
    if (holdsCell(rect, anchor.r, anchor.c) && rect.r2 > rect.r1) return { kind: "filter", rect };
  }
  return region.isNullObject ? null : { kind: "region", rect: rectOf(region) };
}

/**
 * Reads the data the selection points at (spec §7.1). Whole rows or columns are trimmed to the
 * cells in use. Unless `exact` is set, a pointer (one cell, or one row) stands for the table, the
 * filtered range or the block of filled cells around it, and needs a header and a data row. Sizes
 * are checked before any cell content is loaded. Selecting is not a write: nothing here changes the
 * workbook or the selection (invariant 9).
 *
 * Throws SelectionTooLargeError above CELL_LIMIT cells ("The data around E4 has ..." when the range
 * grew from a pointer), and ExcelHostError(RESOLVE_MESSAGE) when there is nothing to read.
 */
export async function readSelection(opts: { exact?: boolean } = {}): Promise<RangeData> {
  return run(async (ctx) => {
    const selected = ctx.workbook.getSelectedRange();
    selected.load(`address,isEntireRow,isEntireColumn,${GEOMETRY}`);
    const ws = selected.worksheet;
    // The filled cells inside the selection, asked for in the same round trip; used only for whole
    // rows or columns.
    const used = selected.getUsedRangeOrNullObject(true);
    used.load(GEOMETRY);
    await ctx.sync();

    const from = splitSheetAddress(selected.address).address.replace(/\$/g, "");
    let rect: Rect | null = rectOf(selected);
    let kind: ResolutionKind = "selection";
    // Whole rows or columns of any size: 4:4 is 16,384 cells, nearly all of them empty.
    if (selected.isEntireRow || selected.isEntireColumn) {
      rect = used.isNullObject ? null : intersect(rect, rectOf(used));
      if (!rect) throw resolveError();
      kind = "trimmed";
    }
    const resolution: Resolution = { kind, from, exact: formatRect(rect) };

    let table: Excel.Table | undefined;
    if (!opts.exact && rect.r1 === rect.r2) {
      const found = await expandPointer(ctx, ws, rect, kind === "trimmed");
      if (!found) throw resolveError();
      rect = found.rect;
      resolution.kind = found.kind;
      if (found.table) {
        table = found.table;
        resolution.table = found.table.name;
      }
    }

    const count = cellsIn(rect);
    if (count > CELL_LIMIT) {
      const grown = resolution.kind === "table" || resolution.kind === "filter" || resolution.kind === "region";
      throw new SelectionTooLargeError(count, grown ? tooLargeAroundMessage(from, count) : TOO_LARGE_MESSAGE);
    }
    // One row is a header with nothing under it. Only an exact read may return it.
    if (!opts.exact && rect.r2 - rect.r1 + 1 < 2) throw resolveError();

    const data = await readRangeData(ctx, ws.getRange(formatRect(rect)), { count, table });
    // The flow, the gate and the placement all work from this address.
    if (!parseLocalRange(data.address)) throw resolveError();
    data.resolution = resolution;
    return data;
  });
}

/** Queues the sheet and range of each loaded table; the getter reads them after the next sync. */
function queueTableInfo(tables: Excel.TableCollection): () => TableInfo[] {
  const found = tables.items.map((t) => {
    const ws = t.worksheet;
    ws.load("name");
    const r = t.getRange();
    r.load("address");
    return { t, ws, r };
  });
  return () => found.map(({ t, ws, r }) => ({ name: t.name, sheet: ws.name, address: splitSheetAddress(r.address).address }));
}

/** Queues each loaded sheet's own names; the getter reads them, as "Sheet!Name", after the next sync. */
function queueSheetNames(sheets: Excel.WorksheetCollection): () => string[] {
  const perSheet = sheets.items.map((ws) => {
    const n = ws.names;
    n.load("items/name");
    return { ws, n };
  });
  return () => perSheet.flatMap(({ ws, n }) => n.items.map((item) => `${ws.name}!${item.name}`));
}

/** Every sheet name in tab order, hidden sheets included (the gate needs to know every real sheet). */
export async function listSheets(): Promise<string[]> {
  return run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    return sheets.items.map((s) => s.name);
  });
}

export async function listTables(): Promise<TableInfo[]> {
  return run(async (ctx) => {
    const tables = ctx.workbook.tables;
    tables.load("items/name");
    await ctx.sync();
    const info = queueTableInfo(tables);
    await ctx.sync();
    return info();
  });
}

/**
 * Workbook-scoped and sheet-scoped defined names. The gate uses them: a function passed by name, as
 * in GROUPBY(..., SUM), is refused when a defined name could shadow it.
 */
export async function listNames(): Promise<string[]> {
  return run(async (ctx) => {
    const names = ctx.workbook.names;
    names.load("items/name");
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const sheetNames = queueSheetNames(sheets);
    await ctx.sync();
    return [...names.items.map((n) => n.name), ...sheetNames()];
  });
}

/**
 * listSheets, listTables and listNames in one Excel.run of two round trips instead of three runs
 * of five. Fails as a whole; the flow then falls back to the three, each with its own default.
 */
export async function readWorkbookInfo(): Promise<{ sheets: string[]; tables: TableInfo[]; names: string[] }> {
  return run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    const tables = ctx.workbook.tables;
    tables.load("items/name");
    const names = ctx.workbook.names;
    names.load("items/name");
    await ctx.sync();
    const tableInfo = queueTableInfo(tables);
    const sheetNames = queueSheetNames(sheets);
    await ctx.sync();
    return {
      sheets: sheets.items.map((s) => s.name),
      tables: tableInfo(),
      names: [...names.items.map((n) => n.name), ...sheetNames()],
    };
  });
}

/**
 * True when every cell in the range is empty. Reads value types only (never values), so Insert can
 * ask before it replaces existing cells. Capped at CELL_LIMIT cells; larger ranges count as not empty.
 */
/** Selecting several separate cells at once (RangeAreas.select) needs this API set. */
const SELECT_AREAS_API = "1.18";
/** At most this many cells are selected at once, so the address list stays short. */
export const MAX_SELECT_CELLS = 150;
const LOCAL_ADDRESS_RE = /^\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?$/u;

/**
 * Selects `cells` (local addresses such as "D2" or "A1:F501") on `sheet` and shows that sheet, so
 * the user can see which cells a value is in. Selection only, not a workbook write (invariant 9).
 * Several cells are selected together where Excel supports it; otherwise the first one is.
 */
export async function selectCells(sheet: string, cells: readonly string[]): Promise<void> {
  const list = cells.filter((c) => LOCAL_ADDRESS_RE.test(c)).slice(0, MAX_SELECT_CELLS);
  if (list.length === 0) throw new ExcelHostError("There are no cells to show.");
  await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    ws.activate();
    if (list.length > 1 && isSetSupported(SELECT_AREAS_API)) ws.getRanges(list.join(", ")).select();
    else ws.getRange(list[0]).select();
    await ctx.sync();
  });
}

export async function isRangeEmpty(sheet: string, address: string): Promise<boolean> {
  return run(async (ctx) => {
    const range = ctx.workbook.worksheets.getItem(sheet).getRange(address);
    range.load("cellCount");
    await ctx.sync();
    if (range.cellCount > CELL_LIMIT) return false;
    range.load("valueTypes");
    await ctx.sync();
    return range.valueTypes.every((row) => row.every((t) => t === "Empty"));
  });
}

/** Reads a user-added context range, e.g. "Regions!A1:B5" or "'Price list'!A1:C9". */
export async function readRange(fullAddress: string): Promise<RangeData> {
  const { sheet, address } = splitSheetAddress(fullAddress.trim());
  if (!parseLocalRange(address.trim())) throw new ExcelHostError("Enter a range like Regions!A1:B5.");
  if (sheet !== null && sheet.trim() === "") throw new ExcelHostError("Enter a range like Regions!A1:B5.");
  try {
    return await run((ctx) => {
      const sheets = ctx.workbook.worksheets;
      const ws = sheet === null ? sheets.getActiveWorksheet() : sheets.getItem(sheet);
      return readRangeData(ctx, ws.getRange(address.trim()));
    });
  } catch (e) {
    if (e instanceof ExcelHostError && e.code === "ItemNotFound") {
      throw new ExcelHostError(NO_SUCH_SHEET_MESSAGE, e.code);
    }
    throw e;
  }
}

/**
 * Letter of the first column at or right of `fromColumn` (0-based) whose cells in sheet rows
 * `firstRow`..`lastRow` are all empty, for the default placement (spec §7.13). Reads value types
 * only. Looks at most `maxColumns` columns, then falls back to `fromColumn`.
 */
export async function firstEmptyColumn(
  sheet: string,
  fromColumn: number,
  firstRow: number,
  lastRow: number,
  maxColumns = 64,
): Promise<string> {
  const int = (v: number, fallback: number) => (Number.isFinite(v) ? Math.floor(v) : fallback);
  const from = Math.min(Math.max(0, int(fromColumn, 0)), MAX_COLUMNS - 1);
  const top = Math.min(Math.max(1, int(firstRow, 1)), MAX_ROWS);
  const bottom = Math.min(Math.max(top, int(lastRow, top)), top + CELL_LIMIT - 1, MAX_ROWS);
  const rows = bottom - top + 1;
  const perChunk = Math.max(1, Math.floor(CELL_LIMIT / rows));
  const stop = Math.min(MAX_COLUMNS, from + Math.max(1, int(maxColumns, 64)));
  return run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    for (let start = from; start < stop; start += perChunk) {
      const end = Math.min(start + perChunk, stop) - 1;
      const range = ws.getRange(formatRect({ r1: top - 1, c1: start, r2: bottom - 1, c2: end }));
      range.load("valueTypes");
      await ctx.sync();
      for (let c = 0; c <= end - start; c++) {
        if (range.valueTypes.every((row) => row[c] === "Empty")) return columnLetter(start + c);
      }
    }
    return columnLetter(from);
  });
}

// ---------------------------------------------------------------------------------------------
// Insert (invariant 9: the only production write)

/** Loads `valueTypes` of at most CELL_LIMIT cells of `rect` and says whether any is an error. */
async function anyError(ctx: Excel.RequestContext, ws: Excel.Worksheet, rect: Rect): Promise<boolean> {
  const width = rect.c2 - rect.c1 + 1;
  const maxRows = Math.max(1, Math.floor(CELL_LIMIT / width));
  const bounded = { ...rect, r2: Math.min(rect.r2, rect.r1 + maxRows - 1) };
  const check = ws.getRange(formatRect(bounded));
  check.load("valueTypes");
  await ctx.sync();
  return check.valueTypes.some((row) => row.some((t) => t === "Error"));
}

function spillSupported(): boolean {
  return isSetSupported(SPILL_API);
}

export interface InsertResult {
  /** The cells written: the filled span, or only the cell when `fillFailed`. */
  address: string;
  check: InsertCheck;
  /** Set for "ok" and "error" only (spec §7.14). */
  excelError?: boolean;
  spill?: string;
  /** The formula is in the cell, but filling it down failed. */
  fillFailed?: boolean;
}

/**
 * Writes `formula` into `cell` on `sheet` and, when `fillDown` is set and `lastRow` (a sheet row
 * number) is below the cell, fills it down to `lastRow`. Then reads back only the value types of
 * the written cells, or of the spill range for a spilling formula, to set `excelError`. Values are
 * never read. When the workbook calculates manually, `excelError` is left unset.
 *
 * The write has its own sync. If it fails, this throws and nothing was written. Once it succeeds
 * the cell holds the formula, so a later failure (fill down, reading back) is reported in the
 * result, never thrown: the caller must not tell the user the insert failed, or insert again.
 * Called only from the Insert button (invariant 9).
 */
export async function insertFormula(
  sheet: string,
  cell: string,
  formula: string,
  fillDown: boolean,
  lastRow: number,
): Promise<InsertResult> {
  const at = parseCell(cell.trim());
  if (!at) throw new ExcelHostError("Pick a single cell for the formula, like G2.");
  const row = at.r + 1;
  const col = columnLetter(at.c);
  const fill = fillDown && Number.isInteger(lastRow) && lastRow > row && lastRow <= MAX_ROWS;
  const written: Rect = { r1: at.r, c1: at.c, r2: fill ? lastRow - 1 : at.r, c2: at.c };
  const address = formatRect(written);
  const cellOnly = `${col}${row}`;

  return run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    const target = ws.getRange(cellOnly);
    target.formulas = [[formula]];
    await ctx.sync();

    // From here on the formula is in the workbook.
    if (fill) {
      try {
        target.autoFill(ws.getRange(address), "FillDefault");
        await ctx.sync();
      } catch {
        return { address: cellOnly, check: "unverified", fillFailed: true };
      }
    }
    try {
      const app = ctx.workbook.application;
      app.load("calculationMode");
      await ctx.sync();
      if (app.calculationMode === "Manual") return { address, check: "manual" };

      if (!fill && spillSupported()) {
        const spill = target.getSpillingToRangeOrNullObject();
        spill.load("rowIndex,columnIndex,rowCount,columnCount");
        await ctx.sync();
        if (!spill.isNullObject) {
          const rect: Rect = {
            r1: spill.rowIndex,
            c1: spill.columnIndex,
            r2: spill.rowIndex + spill.rowCount - 1,
            c2: spill.columnIndex + spill.columnCount - 1,
          };
          const excelError = await anyError(ctx, ws, rect);
          return { address, spill: formatRect(rect), check: excelError ? "error" : "ok", excelError };
        }
      }
      const excelError = await anyError(ctx, ws, written);
      return { address, check: excelError ? "error" : "ok", excelError };
    } catch {
      return { address, check: "unverified" };
    }
  });
}

// ---------------------------------------------------------------------------------------------
// Bench only. The bench panel is the only caller; release builds never reach these.

/** Selects `address` on `sheet`, so readSelection() returns it. Not a workbook write. */
export async function selectRange(sheet: string, address: string): Promise<void> {
  await run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    ws.activate();
    ws.getRange(address).select();
    await ctx.sync();
  });
}

/**
 * Clears the bench scratch sheet (creating it if missing), writes `formula` in column A at
 * `rows.row`, and fills it down to `rows.lastRow` when `fillDown` is set and the formula doesn't
 * spill. Returns the local address to read back: the spill range, the filled range or the cell.
 */
export async function writeScratch(
  formula: string,
  fillDown: boolean,
  rows: { row: number; lastRow: number },
): Promise<string> {
  // Invariant 9: outside bench builds, Insert is the only write.
  if (!BENCH) throw new ExcelHostError("The scratch sheet is only used in bench builds.");
  const row = rows.row;
  if (!Number.isInteger(row) || row < 1 || row > MAX_ROWS) throw new ExcelHostError("The scratch row is out of range.");
  return run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    let ws = sheets.getItemOrNullObject(SCRATCH_SHEET);
    await ctx.sync();
    if (ws.isNullObject) ws = sheets.add(SCRATCH_SHEET);
    ws.getRange().clear("All");
    const anchor = ws.getRange(`A${row}`);
    anchor.formulas = [[formula]];
    await ctx.sync();

    if (spillSupported()) {
      const spill = anchor.getSpillingToRangeOrNullObject();
      spill.load("address");
      await ctx.sync();
      if (!spill.isNullObject) return splitSheetAddress(spill.address).address;
    }
    const lastRow = rows.lastRow;
    if (fillDown && Number.isInteger(lastRow) && lastRow > row && lastRow <= MAX_ROWS) {
      const dest = `A${row}:A${lastRow}`;
      anchor.autoFill(ws.getRange(dest), "FillDefault");
      await ctx.sync();
      return dest;
    }
    return `A${row}`;
  });
}

/** Reads values from the bench scratch sheet. Empty cells are null. */
export async function readScratch(address: string): Promise<CellValue[][]> {
  const local = splitSheetAddress(address.trim()).address;
  if (!parseLocalRange(local)) throw new ExcelHostError("The scratch address is not a range.");
  return run(async (ctx) => {
    const range = ctx.workbook.worksheets.getItem(SCRATCH_SHEET).getRange(local);
    range.load("cellCount");
    await ctx.sync();
    if (range.cellCount > CELL_LIMIT) throw new SelectionTooLargeError(range.cellCount);
    range.load("values,valueTypes,text");
    await ctx.sync();
    return range.values.map((r, i) => r.map((v, j) => cellValue(v, range.valueTypes[i]?.[j], range.text[i]?.[j])));
  });
}

const BENCH_READ = "A1:AZ3";

/** Reads the bench marker sheet, if present: marker in A1, private letters, canaries. */
export async function readBenchInfo(): Promise<{ marker: string | null; privateLetters: string[]; canaries: string[] }> {
  return run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(BENCH_SHEET);
    await ctx.sync();
    if (ws.isNullObject) return { marker: null, privateLetters: [], canaries: [] };
    const range = ws.getRange(BENCH_READ);
    range.load("values");
    await ctx.sync();
    const rowOf = (i: number): unknown[] => range.values[i] ?? [];
    const a1 = rowOf(0)[0];
    const marker = a1 === null || a1 === undefined || a1 === "" ? null : String(a1);
    // Row 2 and row 3 may start with a label ("private", "canaries") in column A.
    const strings = (cells: unknown[], label: string): string[] =>
      cells
        .map((v) => (v === null || v === undefined ? "" : String(v).trim()))
        .filter((s, i) => s !== "" && !(i === 0 && s.toLowerCase() === label));
    const privateLetters = strings(rowOf(1), "private")
      .filter((s) => /^[A-Za-z]{1,3}$/.test(s))
      .map((s) => s.toUpperCase());
    const canaries = strings(rowOf(2), "canaries");
    return { marker, privateLetters, canaries };
  });
}
