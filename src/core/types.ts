// Core types (spec §6). Pure TypeScript: nothing in src/core imports Office.js.
// `AuditedOutbound` is declared in auditor.ts, because only the auditor may create one.

export type ColType = "text" | "number" | "date" | "boolean" | "mixed" | "empty";
export type Treatment = "as_is" | "stand_in" | "range" | "month" | "exclude";
export type Mode = "structure_only" | "substituted" | "raw_bench";
/** Modes the production path can send. `raw_bench` exists only in bench/rawSender.ts. */
export type ProdMode = Exclude<Mode, "raw_bench">;

/** A cell value as the adapter hands it over. Empty cells are `null`. */
export type CellValue = string | number | boolean | null;

/**
 * How the adapter turned Excel's selection into the range it read (spec §7.1): as selected, whole
 * rows or columns trimmed to the cells in use, or a pointer (one cell, one row) grown to the table,
 * the filtered range or the block of filled cells around it.
 */
export type ResolutionKind = "selection" | "trimmed" | "table" | "filter" | "region";

export interface Resolution {
  kind: ResolutionKind;
  /** The selection as Excel reported it, without the sheet: "E4", "4:4", "A:F". */
  from: string;
  /** The selected cells after trimming, as an address parseLocalRange reads ("E4", "A4:F4"). */
  exact: string;
  /** The table's name, for kind "table". */
  table?: string;
}

/**
 * Plain data read from one rectangular range by office/adapter.ts.
 * Arrays are row-major and all the same shape (rows x columns).
 */
export interface RangeData {
  sheet: string;
  /** Address without the sheet, e.g. "A1:F501". */
  address: string;
  /** 0-based sheet row of the top-left cell (row 1 is 0). */
  rowIndex: number;
  /** 0-based sheet column of the top-left cell (column A is 0). */
  columnIndex: number;
  values: CellValue[][];
  /** Display strings as Excel shows them. Empty cells are "". */
  text: string[][];
  /** Formulas ("=...") or constants, as Excel returns `formulas`. */
  formulas: (string | number | boolean | null)[][];
  /** Excel.RangeValueType names: "Empty", "String", "Double", "Integer", "Boolean", "Error", "Unknown", "RichValue". */
  valueTypes: string[][];
  numberFormat: string[][];
  /** The Excel table the range sits in, if any. */
  table?: { name: string; columns: string[] } | null;
  /** Set when the range was read from the selection. Shown in the pane; never sent. */
  resolution?: Resolution;
}

export interface ColumnInfo {
  letter: string;
  header: string;
  alias?: string;
  type: ColType;
  stats: { blank: number; distinct: number; avgWords?: number; maxLen?: number };
}

export interface ColumnPolicy {
  letter: string;
  private: boolean;
  treatment: Treatment;
  note?: string;
  /** Neutral header sent instead of the real one (spec §7.3, §7.6). */
  alias?: string;
}

/**
 * What Insert found once the formula was written (spec §7.1, §7.13). The cells changed in every
 * case: "ok" (no written or spilled cell shows an Excel error), "error" (one does, #SPILL!
 * included), "manual" (the workbook calculates manually, so nothing is calculated yet) or
 * "unverified" (reading the result back, or filling down, failed).
 */
export type InsertCheck = "ok" | "error" | "manual" | "unverified";

/** One described range: the selection, or a user-added context range (spec §7.6). */
export interface RangeSpec {
  data: RangeData;
  hasHeaders: boolean;
  columns: ColumnInfo[];
  /** Same order and letters as `columns`. */
  policies: ColumnPolicy[];
}

export interface TableInfo {
  name: string;
  sheet: string;
  /** Address without the sheet, e.g. "A1:F501". */
  address: string;
}

export interface SheetContext {
  sheet: string;
  table?: string;
  address: string;
  headerRow: number;
  firstDataRow: number;
  lastDataRow: number;
  sheets: string[];
  tableNames: string[];
  /** Selection plus user-added context ranges, e.g. "Orders!A1:F501". */
  allowedRanges: string[];
  /** Tables in the workbook with their ranges, so G-REF can check structured references. */
  tables: TableInfo[];
  /**
   * Workbook and sheet defined names. A bare function name such as SUM (as in GROUPBY(..., SUM)) is
   * only allowed when no defined name shadows it. Undefined means the names weren't read.
   */
  definedNames?: string[];
}

export interface Outbound {
  body: string;
  bytes: number;
  mode: Mode;
  createdAt: string;
}

/** Text-scan detection kinds (spec §7.7). */
export type AuditVariant = "exact" | "normalized" | "digits" | "json" | "word" | "canary";

export interface AuditResult {
  ok: boolean;
  /** Text scan: column letter and variant kind only, never the value itself. */
  hits: { column: string; variant: AuditVariant }[];
  /**
   * Structural check: private columns with a row cell that isn't a stand-in or a rendering.
   * A reason code for that check, kept apart from the text-scan kinds (Jorge's decision 7).
   */
  structural: { column: string }[];
  error?: string;
}

/**
 * Where a text-scan variant came from: a private column's own values, or the original header of a
 * column sent under an alias ("Header sent"), which is checked like a private value.
 */
export type AuditOrigin = "value" | "alias";

/**
 * Where a match sits in the request, as a path with no text in it: the message's index and role,
 * then keys and indexes inside that message's JSON (["task"], ["columns", 3, "header"],
 * ["rows", 12, 4], ["context_ranges", 0, "columns", 1, "header"], a reply's ["explanation"]).
 * The path is [] for a turn that isn't JSON.
 */
export interface AuditField {
  message: number;
  role: "system" | "user" | "assistant";
  path: (string | number)[];
  /**
   * What the match is in, so the pane can read the text around it: a string of the message's own
   * JSON (a description's or a reply's, under any key, even one the path stops at), which the body
   * holds escaped twice; another part of that JSON (a number in a sample row); or text outside it (a
   * turn that isn't JSON, or the text around a reply's JSON), escaped once.
   */
  within: "string" | "json" | "text";
}

/**
 * The first match recorded for one hit (column and variant) and origin, as offsets into the exact
 * body: body.slice(start, end). Value-free, like AuditResult; the pane reads the text from the body it
 * still holds, and neither the offsets nor that text go into the log entry or the report.
 */
export interface AuditLocation {
  column: string;
  variant: AuditVariant;
  start: number;
  end: number;
  origin: AuditOrigin;
  /** Null when the path can't be worked out. */
  field: AuditField | null;
}

export interface ModelReply {
  kind: "formula" | "answer" | "clarify";
  formula: string | null;
  placement: { cell: string; fill_down: boolean } | null;
  explanation: string;
  assumptions: string[];
}

export type GateCode = "G-START" | "G-LEN" | "G-FUNC" | "G-NAME" | "G-EXT" | "G-REF" | "G-URL" | "G-PARSE";

export interface GateResult {
  ok: boolean;
  reasons: { code: GateCode; detail: string }[];
}

export interface LogEntry {
  at: string;
  mode: Mode;
  model: string;
  host: string;
  bytes: number;
  /** Exactly what was sent (or would have been, if blocked); stand-ins only. */
  body: string;
  audit: AuditResult;
  replyRaw?: string;
  gate?: GateResult;
  inserted: boolean;
  latencyMs?: number;
  tokens?: { prompt: number; completion: number };
  /** Which upstream provider served the request, from OpenRouter's X-Provider-Name header. */
  providerName?: string;
  /** A plain error message when the send failed. */
  error?: string;
}

/** Non-content evaluation metadata (spec §7.14). A closed list: adding a field is a spec change. */
export interface EvalRecord {
  build: string;
  model: string;
  mode: ProdMode;
  rows: number;
  columns: number;
  privateColumns: number;
  treatments: Record<Treatment, number>;
  bytes: number;
  tokens?: { prompt: number; completion: number };
  latencyMs?: number;
  audit: "pass" | "block";
  auditReasons: (AuditVariant | "structural" | "error")[];
  gate?: "pass" | "block";
  gateCodes: GateCode[];
  replyKind?: ModelReply["kind"] | "invalid";
  retries: number;
  inserted: boolean;
  copied: boolean;
  excelError?: boolean;
  rating?: "worked" | "partly" | "wrong";
}

export type TokenKind = "PERSON" | "EMAIL" | "ID" | "TEXT";

/** A message kept for the model. Always in stand-in form (invariant 5). */
export interface HistoryMessage {
  role: "user" | "assistant";
  content: string;
}
