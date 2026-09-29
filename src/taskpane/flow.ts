// Orchestration for the task pane (spec §9). Plain TypeScript with no React, so it can be tested with
// a fake Excel host and a fake sender. The API key is passed in for each send and never kept here.
//
// Order of a question: substituteUserText -> payload -> audit -> (user clicks Send) -> provider ->
// reply -> gate -> restore -> gate again -> (user clicks Insert) -> adapter.insertFormula.
import { COMMIT, DEFAULT_MODEL, ENDPOINT_HOST, VERSION, buildLabel } from "../config";
import { MAX_COLUMNS, MAX_ROWS, columnIndex, columnLetter, formatRect, parseCell, parseLocalRange, qualify, splitSheetAddress } from "../core/a1";
import { audit, buildAuditContext, type AuditContext, type AuditOutcome, type AuditedOutbound } from "../core/auditor";
import { ALLOWLIST } from "../core/allowlist";
import { check, checkFill, tokenize } from "../core/formulaGate";
import { applyChoices, findPartialMatches, type PartialMatch } from "../core/partial";
import { RequestLog, type LocalMatch, type LogLocal } from "../core/log";
import { MODEL_ID_RE, MetricsStore } from "../core/metrics";
import { buildCorrection, buildStructureOnly, buildSubstituted, type BuiltPayload, type PayloadInput } from "../core/payload";
import { appendExchange } from "../core/prompt";
import { send as providerSend, type SendErr, type SendOk } from "../core/provider";
import { parseReply, type ParseResult } from "../core/reply";
import { restoreFormula, restoreText } from "../core/restore";
import { buildSheetContext, columnCount, firstRowLooksLikeData, headerCellLooksLikeData, inferSchema, rowCount } from "../core/schema";
import { defaultTreatment, suggestPolicies, type Suggestion } from "../core/suggest";
import {
  DEFAULT_ROW_CAP,
  TOKEN_RE,
  clampRowCap,
  createStandInMap,
  effectiveTreatment,
  privateCellsOf,
  substituteUserText,
  type StandInMap,
} from "../core/transform";
import type {
  AuditField,
  AuditLocation,
  AuditOrigin,
  AuditResult,
  AuditVariant,
  ColType,
  ColumnInfo,
  ColumnPolicy,
  EvalRecord,
  GateResult,
  HistoryMessage,
  InsertCheck,
  LogEntry,
  ModelReply,
  ProdMode,
  RangeData,
  RangeSpec,
  Resolution,
  ResolutionKind,
  SheetContext,
  TableInfo,
  Treatment,
} from "../core/types";
import { COPY } from "./copy";

// ---------------------------------------------------------------------------------------------
// Dependencies

/** What the flow needs from the Excel host. index.tsx wires this to src/office/adapter.ts. */
export interface HostAdapter {
  /**
   * The data the selection points at: one cell or one row stands for the table, filtered range or
   * block of filled cells around it, unless `exact` is set (spec §7.1).
   */
  readSelection(opts?: { exact?: boolean }): Promise<RangeData>;
  listSheets(): Promise<string[]>;
  listTables(): Promise<TableInfo[]>;
  /** Workbook and sheet defined names ("Sheet!Name" for sheet-scoped ones). */
  listNames?(): Promise<string[]>;
  /** Sheets, tables and names in one read. Optional: without it, or when it fails, the three above are used. */
  readWorkbookInfo?(): Promise<{ sheets: string[]; tables: TableInfo[]; names: string[] }>;
  readRange(fullAddress: string): Promise<RangeData>;
  /**
   * Called only from insert(), which only the Insert button calls (invariant 9). Throws only when
   * nothing was written; once the formula is in the cell, what happened next is in `check`.
   */
  insertFormula(
    sheet: string,
    cell: string,
    formula: string,
    fillDown: boolean,
    lastRow: number,
  ): Promise<{ address: string; check?: InsertCheck; excelError?: boolean; fillFailed?: boolean }>;
  /**
   * Letter of the first column at or right of `fromColumn` (0-based) that is empty in rows
   * firstRow..lastRow. Reads value types only. Optional: without it, the column right of the
   * selection is used.
   */
  firstEmptyColumn?(sheet: string, fromColumn: number, firstRow: number, lastRow: number): Promise<string>;
  /** True when every cell in the range is empty (value types only). Lets Insert ask before replacing. */
  isRangeEmpty?(sheet: string, address: string): Promise<boolean>;
  /**
   * Selects `cells` (local addresses) on `sheet` in Excel, so the user sees them. Changes the
   * selection only, never the workbook. Optional: without it, Show in sheet isn't offered.
   */
  selectCells?(sheet: string, cells: readonly string[]): Promise<void>;
  /** True when a read failed because the range has too many cells. */
  isTooLarge?(error: unknown): boolean;
  /** The host's own plain message for an error, when it has one (never a stack trace or a value). */
  plainMessage?(error: unknown): string | null;
}

export type SendFn = (outbound: AuditedOutbound, key: string, signal?: AbortSignal) => Promise<SendOk | SendErr>;

export interface FlowDeps {
  adapter: HostAdapter;
  /** Defaults to provider.send, the one sender. */
  send?: SendFn;
  model?: string;
  /** Endpoint host shown in the UI and written to the log. */
  host?: string;
  /** Version and build commit, as in the footer. */
  build?: string;
  /** Bench canaries (bench workbooks only). */
  canaries?: string[];
  now?: () => Date;
}

export type ActionResult = { ok: true; message?: string } | { ok: false; message: string };

// ---------------------------------------------------------------------------------------------
// State the UI reads

export interface ColumnView {
  letter: string;
  header: string;
  type: ColType;
  stats: ColumnInfo["stats"];
  policy: Readonly<ColumnPolicy>;
  /** "Suggested: header mentions email. Your choice decides." or null. */
  hint: string | null;
  suggestedPrivate: boolean;
}

export interface RangeView {
  /** "Orders!A1:F501" */
  label: string;
  sheet: string;
  address: string;
  table: string | null;
  hasHeaders: boolean;
  /** Shown under the header checkbox when row 1 looked like data and is treated as data. */
  notice: string | null;
  dataRows: number;
  /** "Orders, A1:F501 · 500 rows" */
  summary: string;
  /**
   * How the range came from the selection, e.g. "From the table Orders around E4."; null when it
   * was used as selected. `exact` is true when "Use exactly the selected cells" can undo the choice.
   */
  resolution: { kind: ResolutionKind; text: string; exact: boolean } | null;
  columns: readonly ColumnView[];
}

export interface FlowState {
  /** 0 before the first read. A new session has a fresh stand-in map and empty history. */
  sessionId: number;
  /** Changes whenever columns, ranges or history change; a preview is only sendable at its version. */
  version: number;
  /**
   * True once a check at Send or Insert found that the workbook no longer matches what the session
   * read (spec §7.1): nothing more is previewed, sent or inserted until Refresh from selection.
   */
  stale: boolean;
  ranges: readonly RangeView[];
  context: SheetContext | null;
  /** Messages kept for the model (stand-in form). */
  historyMessages: number;
  logSize: number;
  /** Changes on every update, for React. */
  revision: number;
}

export type Segment = { kind: "text"; text: string } | { kind: "token"; text: string };

export interface MessageView {
  role: "system" | "user" | "assistant";
  /** True for a Nymform description turn (the JSON user turn). */
  description: boolean;
  segments: Segment[];
}

/** A readable view of the exact body. Built by parsing the body, so it shows only what is sent. */
export interface RequestView {
  model: string;
  params: { key: string; value: string }[];
  system: string;
  earlier: MessageView[];
  current: MessageView | null;
  bytes: number;
  /** The literal string that is audited and sent. */
  exact: string;
}

export interface AuditStatus {
  ok: boolean;
  headline: string;
  detail?: string;
  action?: string;
  /** Where in the request the check found text ("Found in your question."), never the text itself. */
  found?: string;
}

export interface GateStatus {
  ok: boolean;
  headline: string;
  reasons: string[];
}

export interface Prepared {
  readonly id: number;
  readonly sessionId: number;
  readonly version: number;
  readonly model: string;
  readonly mode: ProdMode;
  readonly rowCap: number;
  /** The question as sent, after private values were swapped for stand-ins. */
  readonly sentQuestion: string;
  readonly built: BuiltPayload;
  readonly outcome: AuditOutcome;
  readonly status: AuditStatus;
  readonly view: RequestView;
  /** Parts of private values in the question and what became of them. Pane only; never logged. */
  readonly partial: PartialView;
  /**
   * Set when this preview is the one correction request for a reply that couldn't be read (spec
   * §7.9): the question again, the model's reply and a note asking for the format. Like any
   * request it is audited, shown and sent only when the user chooses Send; it never leads to another.
   */
  readonly correction?: { readonly metricsId: number | null; readonly userContent: string };
}

/** A part of a private value typed in the question ("Felix" of "Felix Bianchi"), and its stand-in. */
export interface PartialPick {
  key: string;
  fragment: string;
  values: readonly string[];
  tokens: readonly string[];
}

export interface PartialView {
  /** Parts the user needs to pick for: several candidates, or one that is free text. */
  pending: readonly PartialMatch[];
  /** Parts read as their only candidate without asking; shown so the user can undo it. */
  auto: readonly PartialPick[];
  /** Parts the user picked for. */
  chosen: readonly PartialPick[];
  /** Parts the user chose to keep as typed (the check then blocks them). */
  kept: readonly { key: string; fragment: string }[];
}

/** For each part (by PartialMatch.key), the stand-ins chosen for it; an empty list keeps it as typed. */
export type PartialChoices = ReadonlyMap<string, readonly string[]>;

export interface FlowResult {
  readonly id: number;
  readonly kind: ModelReply["kind"];
  readonly mode: ProdMode;
  /** Formula to show and insert: restored when the first gate pass succeeded, else the model's own. */
  readonly formula: string | null;
  /** The model's formula, stand-in form. */
  readonly modelFormula: string | null;
  readonly gate: GateResult | null;
  readonly gateStatus: GateStatus | null;
  /** Both gate passes succeeded and every stand-in in the formula's strings was put back. */
  readonly canInsert: boolean;
  /** Restored for display only. */
  readonly explanation: string;
  readonly assumptions: readonly string[];
  /** Stand-ins anywhere in the reply that this session didn't create. */
  readonly unknownTokens: readonly string[];
  /**
   * The ones inside the formula's string literals. The formula would compare against the stand-in
   * text rather than a value, so it can't be inserted or copied, whatever the gate says.
   */
  readonly unresolvedTokens: readonly string[];
  /**
   * Values put back into the formula that Excel may read differently from the sheet: a number or
   * date comes back as text, and *, ? and ~ are wildcards in criteria. Shown before Insert; pane only.
   */
  readonly cautions: readonly string[];
  /**
   * Whole-column references in the formula ("E:E"). Allowed for the chosen columns (spec §7.10,
   * G-REF), and shown so the user knows the formula reads rows outside the selection, in Excel.
   */
  readonly wholeColumns: readonly string[];
  readonly placement: { cell: string; fillDown: boolean } | null;
  readonly sheet: string;
  readonly lastRow: number;
  readonly retries: number;
}

export type SendOutcome =
  | { ok: true; result: FlowResult }
  /** With `correction`: the reply couldn't be read, and this is the one correction request to preview. */
  | { ok: false; message: string; detail?: string; correction?: Prepared };

export type InsertOutcome =
  | { ok: true; message: string; address: string; check: InsertCheck }
  | {
      ok: false;
      message: string;
      reasons?: string[];
      /** The target may have values: insert again with confirmedSpan set to `span` to replace them. */
      needsConfirm?: boolean;
      /** The cells the confirmation is for, e.g. "G2:G501". */
      span?: string;
    };

// ---------------------------------------------------------------------------------------------
// Internal state

interface RangeSlot {
  data: RangeData;
  hasHeaders: boolean;
  /** True once the user set the header checkbox; otherwise hasHeaders was detected from row 1. */
  headersChosen: boolean;
  columns: ColumnInfo[];
  policies: ColumnPolicy[];
  suggestions: Suggestion[];
  /** Columns whose alias is the automatic "Column {letter}" given because the header looks like data. */
  autoAliases: Set<string>;
}

/** Which parts of the workbook info were read, rather than filled with their fail-closed defaults. */
interface WorkbookRead {
  tables: boolean;
  names: boolean;
}

type WorkbookInfo = { sheets: string[]; tables: TableInfo[]; names?: string[]; read?: WorkbookRead };

interface Session {
  id: number;
  ranges: RangeSlot[];
  workbook: WorkbookInfo;
  context: SheetContext;
  /** Memory only; never logged, exported or handed to the UI (invariant 6). */
  map: StandInMap;
  /** Stand-in form only (invariant 5). */
  history: HistoryMessage[];
  /** The workbook changed since the ranges were read (see FlowState.stale). */
  stale: boolean;
}

interface ResultRecord {
  view: FlowResult;
  sessionId: number;
  logId: number;
  metricsId: number | null;
}

const PRIVATE_TREATMENTS: readonly Treatment[] = ["stand_in", "range", "month", "exclude"];
const PLAIN_TREATMENTS: readonly Treatment[] = ["as_is", "exclude"];

export function treatmentsFor(isPrivate: boolean): readonly Treatment[] {
  return isPrivate ? PRIVATE_TREATMENTS : PLAIN_TREATMENTS;
}

export class Flow {
  readonly #adapter: HostAdapter;
  readonly #send: SendFn;
  readonly #host: string;
  readonly #build: string;
  readonly #canaries: string[];
  readonly #now: () => Date;
  readonly #log = new RequestLog();
  readonly #metrics = new MetricsStore();
  readonly #results = new Map<number, ResultRecord>();
  /** Correction previews whose retry is already in their evaluation record. */
  readonly #retriesCounted = new Set<number>();
  readonly #listeners = new Set<() => void>();

  #model: string;
  #sending = false;
  #inserting = false;
  #session: Session | null = null;
  #sessionCounter = 0;
  #version = 0;
  #revision = 0;
  #counter = 0;
  #state: FlowState;

  constructor(deps: FlowDeps) {
    this.#adapter = deps.adapter;
    this.#send = deps.send ?? ((outbound, key, signal) => providerSend(outbound, key, signal));
    this.#model = (deps.model ?? DEFAULT_MODEL).trim();
    this.#host = deps.host ?? ENDPOINT_HOST;
    this.#build = deps.build ?? buildLabel(VERSION, COMMIT);
    this.#canaries = [...(deps.canaries ?? [])];
    this.#now = deps.now ?? (() => new Date());
    this.#state = this.#snapshot();
  }

  // ----- store interface (for useSyncExternalStore)

  subscribe = (listener: () => void): (() => void) => {
    this.#listeners.add(listener);
    return () => {
      this.#listeners.delete(listener);
    };
  };

  getState = (): FlowState => this.#state;

  get model(): string {
    return this.#model;
  }

  get host(): string {
    return this.#host;
  }

  setModel(model: string): void {
    this.#model = model.trim();
    this.#emit();
  }

  /** The history kept for the model: stand-in form only. A copy. */
  history(): HistoryMessage[] {
    return (this.#session?.history ?? []).map((m) => ({ ...m }));
  }

  logEntries(): readonly LogEntry[] {
    return this.#log.entries();
  }

  /**
   * The log for the Log screen: each entry with what a blocked request matched (LogLocal). Pane
   * only: the matches hold the private text, and never go into exportLog() or exportReport().
   */
  logItems(): readonly { entry: LogEntry; local: LogLocal | undefined }[] {
    return this.#log.items();
  }

  /** Export log: sent bodies, audit results, raw replies, gate results. Never the key or the map. */
  exportLog(): string {
    return this.#log.exportJson();
  }

  /** Export evaluation report: validated EvalRecords only (spec §7.14). */
  exportReport(): string {
    return this.#metrics.exportJson();
  }

  evalRecords(): readonly EvalRecord[] {
    return this.#metrics.records();
  }

  // ----- sessions and columns

  /**
   * Reads the selection and starts a new session (fresh stand-in map, empty history). The pane
   * reads the selection only here: until the next Refresh, clicks in Excel change nothing.
   */
  async refresh(): Promise<ActionResult> {
    let data: RangeData;
    try {
      data = await this.#adapter.readSelection();
    } catch (e) {
      return fail(this.#tooLarge(e) ? this.#tooLargeMessage(e, COPY.tooLarge) : (this.#hostMessage(e) ?? COPY.readFailed));
    }
    return this.#startWithMain(data);
  }

  /**
   * "Use exactly the selected cells": reads the cells the selection resolved from (the "E4" behind
   * "From the table Orders around E4."), on that sheet, rather than Excel's current selection, which
   * may have moved since. Starts a new session, as Refresh does.
   */
  async useExactSelection(): Promise<ActionResult> {
    const main = this.#session?.ranges[0];
    const exact = main?.data.resolution?.exact;
    const rect = typeof exact === "string" ? parseLocalRange(exact) : null;
    if (!main || !rect) return fail(COPY.noSelection);
    let data: RangeData;
    try {
      data = await this.#adapter.readRange(qualify(main.data.sheet, formatRect(rect)));
    } catch (e) {
      if (this.#tooLarge(e)) return fail(COPY.tooLarge);
      const m = this.#hostMessage(e);
      // The sheet was renamed or deleted since; the user typed no address to check.
      return fail(m === COPY.noSuchSheet ? COPY.exactSheetGone : (m ?? COPY.readFailed));
    }
    return this.#startWithMain(data);
  }

  /** A new session on `data` as the main range, keeping the header choice and lookup ranges. */
  async #startWithMain(data: RangeData): Promise<ActionResult> {
    if (!isUsable(data)) return fail(COPY.noSelection);
    const workbook = await this.#readWorkbook(data.sheet);
    const prev = this.#session;
    const main = makeSlot(data, chosenHeaders(prev?.ranges[0], data), prev?.ranges[0]);

    // Lookup ranges the user added stay, read again.
    const slots: RangeSlot[] = [main];
    const dropped: string[] = [];
    for (const old of prev?.ranges.slice(1) ?? []) {
      const label = labelOf(old.data);
      if (slots.some((s) => sameRange(s.data, old.data.sheet, old.data.address))) continue;
      try {
        const read = await this.#adapter.readRange(label);
        // Read again by its address, which carries no resolution: keep the line that says how the
        // range grew from the cell the user picked ("From the filtered range around B7").
        const fresh = old.data.resolution && sameRange(old.data, read.sheet, read.address) ? { ...read, resolution: old.data.resolution } : read;
        if (isUsable(fresh)) slots.push(makeSlot(fresh, chosenHeaders(old, fresh), old));
        else dropped.push(label);
      } catch {
        dropped.push(label);
      }
    }
    // Everything was read just now, so the session is no longer stale.
    this.#startSession(slots, workbook, true);
    if (dropped.length > 0) {
      return { ok: true, message: `Couldn't read ${dropped.join(", ")} again, so it was removed from the lookup ranges.` };
    }
    return { ok: true };
  }

  /** "This range has a header row", for the main range or a lookup range. Starts a new session. */
  setHasHeaders(rangeIndex: number, hasHeaders: boolean): void {
    const s = this.#session;
    const slot = s?.ranges[rangeIndex];
    if (!s || !slot) return;
    if (slot.hasHeaders === hasHeaders) {
      if (slot.headersChosen) return;
      slot.headersChosen = true;
      this.#emit();
      return;
    }
    const slots = s.ranges.map((r, i) => (i === rangeIndex ? makeSlot(r.data, hasHeaders) : r));
    this.#startSession(slots, s.workbook);
  }

  /** Changes one column's policy. Any policy change clears the history; the map stays. */
  setPolicy(rangeIndex: number, letter: string, patch: Partial<Omit<ColumnPolicy, "letter">>): void {
    const s = this.#session;
    const slot = s?.ranges[rangeIndex];
    if (!s || !slot) return;
    const i = slot.policies.findIndex((p) => p.letter === letter);
    const old = slot.policies[i];
    const col = slot.columns[i];
    if (!old || !col) return;

    const next: ColumnPolicy = { ...old, ...patch, letter: old.letter };
    if (patch.private !== undefined && patch.private !== old.private && patch.treatment === undefined) {
      next.treatment = old.treatment === "exclude" ? "exclude" : next.private ? defaultTreatment(col.type) : "as_is";
    }
    // An alias the user typed is theirs to keep.
    if (patch.alias !== undefined) slot.autoAliases.delete(col.letter);
    if (next.private && !old.private && patch.alias === undefined && next.alias === undefined && headerLooksLikeData(slot, i)) {
      next.alias = autoAlias(col.letter);
      slot.autoAliases.add(col.letter);
    }
    // Not private any more: the automatic alias goes, so the header is sent as it is.
    if (!next.private && old.private && slot.autoAliases.has(col.letter)) {
      delete next.alias;
      slot.autoAliases.delete(col.letter);
    }
    if (!treatmentsFor(next.private).includes(next.treatment)) {
      next.treatment = next.private ? defaultTreatment(col.type) : "as_is";
    }
    if (typeof next.alias === "string" && next.alias.trim() === "") delete next.alias;
    if (typeof next.note === "string" && next.note === "") delete next.note;
    if (samePolicy(old, next)) return;

    slot.policies = slot.policies.map((p, j) => (j === i ? next : p));
    s.history = [];
    this.#version++;
    this.#emit();
  }

  /**
   * Adds a lookup range by address, e.g. "Regions!A1:B5". Starts a new session. Without a choice
   * for hasHeaders, row 1 is a header row unless it looks like data.
   */
  async addContextRange(address: string, hasHeaders?: boolean): Promise<ActionResult> {
    const s = this.#session;
    if (!s) return fail(COPY.noSelection);
    const { sheet, address: local } = splitSheetAddress(address.trim());
    const rect = parseLocalRange(local.trim());
    if (!rect || (sheet !== null && sheet.trim() === "")) return fail(COPY.rangeFormat);
    const sheetName = sheet ?? s.context.sheet;
    const localAddress = formatRect(rect);
    if (s.ranges.some((r) => sameRange(r.data, sheetName, localAddress))) return fail(COPY.rangeDuplicate);
    let data: RangeData;
    try {
      data = await this.#adapter.readRange(qualify(sheetName, localAddress));
    } catch (e) {
      return fail(this.#tooLarge(e) ? COPY.tooLargeRange : (this.#hostMessage(e) ?? COPY.rangeReadFailed));
    }
    return this.#addContextData(data, hasHeaders);
  }

  /** Adds the cells currently selected in Excel as a lookup range. Starts a new session. */
  async addSelectionAsContext(hasHeaders?: boolean): Promise<ActionResult> {
    if (!this.#session) return fail(COPY.noSelection);
    let data: RangeData;
    try {
      data = await this.#adapter.readSelection();
    } catch (e) {
      if (this.#tooLarge(e)) return fail(this.#tooLargeMessage(e, COPY.tooLargeRange));
      const m = this.#hostMessage(e);
      // The adapter's advice is about the main range; here the user is adding a lookup.
      return fail(m === COPY.resolveSelection ? COPY.resolveLookup : (m ?? COPY.readFailed));
    }
    return this.#addContextData(data, hasHeaders);
  }

  removeContextRange(rangeIndex: number): void {
    const s = this.#session;
    if (!s || rangeIndex < 1 || rangeIndex >= s.ranges.length) return;
    this.#startSession(
      s.ranges.filter((_, i) => i !== rangeIndex),
      s.workbook,
    );
  }

  clearHistory(): void {
    const s = this.#session;
    if (!s || s.history.length === 0) return;
    s.history = [];
    this.#version++;
    this.#emit();
  }

  // ----- a question

  /**
   * Builds and audits the request. Nothing is sent. A blocked request is logged and recorded.
   * A part of a private value in the question ("Felix") is replaced by the stand-in `choices` gives
   * for it; without a choice, a part with one candidate naming a person, email or ID stands for it,
   * and any other part waits for the user (Prepared.partial.pending) and blocks.
   */
  prepare(
    question: string,
    mode: ProdMode,
    rowCap: number = DEFAULT_ROW_CAP,
    choices: PartialChoices = new Map(),
  ): { ok: true; prepared: Prepared } | { ok: false; message: string } {
    const s = this.#session;
    if (!s) return fail(COPY.noSelection);
    if (s.stale) return fail(COPY.dataChanged);
    // A header row alone gives the model nothing to compute on (one clicked cell used to end up here).
    const main = s.ranges[0];
    if (!main || dataRows(main) === 0) return fail(COPY.noDataRows);
    const text = typeof question === "string" ? question.trim() : "";
    if (text === "") return fail(COPY.noQuestion);
    const model = this.#model;
    if (!MODEL_ID_RE.test(model)) return fail(COPY.modelInvalid);
    if (mode !== "structure_only" && mode !== "substituted") return fail(COPY.unexpected);

    try {
      const cap = clampRowCap(rowCap);
      const specs = specsOf(s);
      const substituted = substituteUserText(text, specs, s.map);
      const matches = findPartialMatches(substituted, specs, s.map);
      const decided = new Map<string, readonly string[]>();
      const partial: { pending: PartialMatch[]; auto: PartialPick[]; chosen: PartialPick[]; kept: { key: string; fragment: string }[] } = {
        pending: [],
        auto: [],
        chosen: [],
        kept: [],
      };
      const pick = (m: PartialMatch, tokens: readonly string[]): PartialPick => ({
        key: m.key,
        fragment: m.fragment,
        tokens,
        values: tokens.map((t) => m.candidates.find((c) => c.token === t)?.value ?? ""),
      });
      for (const m of matches) {
        const chosen = choices.get(m.key);
        const valid = chosen?.filter((t) => m.candidates.some((c) => c.token === t)) ?? [];
        if (chosen && chosen.length === 0) {
          decided.set(m.key, []);
          partial.kept.push({ key: m.key, fragment: m.fragment });
        } else if (valid.length > 0) {
          decided.set(m.key, valid);
          partial.chosen.push(pick(m, valid));
        } else if (m.autoMap) {
          const only = [m.candidates[0]!.token];
          decided.set(m.key, only);
          partial.auto.push(pick(m, only));
        } else {
          partial.pending.push(m);
        }
      }
      const sentQuestion = applyChoices(substituted, matches, decided);
      const input: PayloadInput = {
        model,
        specs,
        sheetContext: s.context,
        question: sentQuestion,
        history: s.history,
        map: s.map,
        rowCap: cap,
      };
      const built = mode === "substituted" ? buildSubstituted(input) : buildStructureOnly(input);
      const outcome = audit(built.outbound, this.#auditContext(s, specs, model));
      // Read what the check matched from the body now; a blocked body isn't kept anywhere else.
      const found = outcome.ok ? [] : localMatches(built.outbound.body, outcome.locations);
      const status = outcome.ok ? auditStatus(outcome.outbound.audit) : auditStatus(outcome.result, outcome.locations);
      if (found.length > 0) status.found = COPY.auditFound(uniq(found.map((m) => m.where)));
      const prepared: Prepared = Object.freeze({
        id: ++this.#counter,
        sessionId: s.id,
        version: this.#version,
        model,
        mode,
        rowCap: cap,
        sentQuestion,
        built,
        outcome,
        status,
        view: formatRequest(built.outbound.body, built.outbound.bytes, (t) => s.map.isToken(t)),
        partial: Object.freeze(partial),
      });
      if (!outcome.ok) {
        this.#logBlocked(model, mode, built.outbound.bytes, outcome.result, status.headline, found);
        this.#metrics.add(this.#record(s, model, mode, built.outbound.bytes, outcome.result));
        this.#emit();
      }
      return { ok: true, prepared };
    } catch {
      return fail(COPY.unexpected);
    }
  }

  /** True when this preview can still be sent as is. */
  /** True when the host can select cells, so the pane can offer Show in sheet. */
  canShowCells(): boolean {
    return typeof this.#adapter.selectCells === "function";
  }

  /**
   * Selects a candidate's cells in Excel so the user can see which value it is. Selection only;
   * nothing is written. The cells come from a Prepared's partial matches.
   */
  async showCells(sheet: string, cells: readonly string[]): Promise<ActionResult> {
    if (!this.#adapter.selectCells) return fail(COPY.showCellsUnavailable);
    if (cells.length === 0) return fail(COPY.unexpected);
    try {
      await this.#adapter.selectCells(sheet, cells);
      return { ok: true };
    } catch (e) {
      return fail(this.#hostMessage(e) ?? COPY.showCellsFailed);
    }
  }

  /**
   * Puts the selection back after Show in sheet moved it: the cells the user had selected (E4, not
   * the block around it), so a later Refresh finds the same range the same way.
   */
  async restoreSelection(): Promise<void> {
    const r = this.#session?.ranges[0]?.data;
    if (!r || !this.#adapter.selectCells) return;
    try {
      await this.#adapter.selectCells(r.sheet, [r.resolution?.exact ?? r.address]);
    } catch {
      // The selection is a convenience; the questions don't depend on it.
    }
  }

  isCurrent(prepared: Prepared): boolean {
    return (
      this.#session !== null && !this.#session.stale && this.#session.id === prepared.sessionId && prepared.version === this.#version
    );
  }

  /**
   * Sends an audited request, parses the reply, checks the formula, restores it, checks it again.
   * Appends the sent user turn and the model's own reply text (stand-in form) to the history.
   * First reads the ranges and the workbook's tables and names again: if they changed since the
   * preview was built, nothing is sent (spec §7.1). A reply that can't be read comes back as the one
   * correction request (`correction`), audited, for the user to preview and send (spec §7.9).
   */
  async send(prepared: Prepared, key: string, signal?: AbortSignal): Promise<SendOutcome> {
    if (!prepared.outcome.ok) return fail(prepared.status.headline);
    const s = this.#session;
    if (s?.stale) return fail(COPY.dataChanged);
    if (!s || !this.isCurrent(prepared)) return fail(COPY.stalePreview);
    if (this.#sending) return fail(COPY.alreadySending);
    const { mode, model } = prepared;

    this.#sending = true;
    try {
      const fresh = await this.#recheck(s);
      if (fresh === "changed") return fail(COPY.dataChangedSend);
      if (fresh === "failed") return fail(COPY.recheckFailedSend);
      if (!this.isCurrent(prepared)) return fail(COPY.stalePreview);
      const versionAtSend = this.#version;

      const outbound = prepared.outcome.outbound;
      const res = await this.#sendLogged(outbound, model, key, signal);
      let metricsId: number | null;
      if (prepared.correction) {
        metricsId = prepared.correction.metricsId;
        // Sent again after a failure, the same correction isn't a second retry (spec §7.14: retries is 1).
        if (metricsId !== null && !this.#retriesCounted.has(prepared.id)) {
          this.#retriesCounted.add(prepared.id);
          const retry: { bytes: number; tokens?: { prompt: number; completion: number }; latencyMs?: number } = {
            bytes: outbound.bytes,
            latencyMs: res.res.latencyMs,
          };
          if (res.res.ok && res.res.tokens) retry.tokens = res.res.tokens;
          this.#metrics.addRetry(metricsId, retry);
        }
      } else {
        const record = this.#record(s, model, mode, outbound.bytes, outbound.audit);
        if (res.res.ok && res.res.tokens) record.tokens = res.res.tokens;
        record.latencyMs = res.res.latencyMs;
        metricsId = this.#metrics.add(record);
      }
      this.#emit();
      if (!res.res.ok) {
        if (prepared.correction) this.#updateMetrics(metricsId, { replyKind: "invalid" });
        return fail(res.res.message);
      }

      const parsed: ParseResult = parseReply(res.res.raw, mode);
      if (!parsed.ok) {
        this.#updateMetrics(metricsId, { replyKind: "invalid" });
        this.#emit();
        // One correction only, and only while the preview it corrects still stands.
        if (prepared.correction || this.#session !== s || this.#version !== versionAtSend) {
          return { ok: false, message: COPY.unreadable, detail: parsed.error };
        }
        return this.#correction(s, prepared, parsed.content ?? "", metricsId);
      }

      // History gets the exact user turn that was sent and the model's own text: stand-ins only.
      // If the columns changed meanwhile, the history was cleared and this exchange is left out.
      const userContent = prepared.correction?.userContent ?? prepared.built.userContent;
      if (this.#session === s && this.#version === versionAtSend) {
        s.history = appendExchange(s.history, userContent, parsed.content);
        this.#version++;
      }

      const fallback = parsed.reply.kind === "formula" ? await this.#defaultCell(s) : null;
      const view = this.#makeResult(s, parsed.reply, mode, prepared.correction ? 1 : 0, fallback);
      if (view.gate) this.#log.update(res.logId, { gate: view.gate });
      const patch: Partial<EvalRecord> = { replyKind: parsed.reply.kind };
      if (view.gate) {
        patch.gate = view.gate.ok ? "pass" : "block";
        patch.gateCodes = uniq(view.gate.reasons.map((r) => r.code));
      }
      this.#updateMetrics(metricsId, patch);
      this.#results.set(view.id, { view, sessionId: s.id, logId: res.logId, metricsId });
      this.#emit();
      return { ok: true, result: view };
    } catch {
      this.#emit();
      return fail(COPY.unexpected);
    } finally {
      this.#sending = false;
    }
  }

  /**
   * The one correction request for a reply that couldn't be read: built by payload.ts and audited
   * like any request, then handed back to be previewed. Nothing is sent here.
   */
  #correction(s: Session, prepared: Prepared, invalidReply: string, metricsId: number | null): SendOutcome {
    const { model, mode } = prepared;
    const built = buildCorrection({ model, mode, history: s.history, userContent: prepared.built.userContent, invalidReply });
    const outcome = audit(built.outbound, this.#auditContext(s, specsOf(s), model));
    if (!outcome.ok) {
      const status = auditStatus(outcome.result, outcome.locations);
      const matches = localMatches(built.outbound.body, outcome.locations);
      this.#logBlocked(model, mode, built.outbound.bytes, outcome.result, status.headline, matches);
      this.#updateMetrics(metricsId, { audit: "block", auditReasons: auditReasons(outcome.result) });
      this.#emit();
      return fail(COPY.retryBlocked);
    }
    const next: Prepared = Object.freeze({
      id: ++this.#counter,
      sessionId: s.id,
      version: this.#version,
      model,
      mode,
      rowCap: prepared.rowCap,
      sentQuestion: prepared.sentQuestion,
      built,
      outcome,
      status: auditStatus(outcome.outbound.audit),
      view: formatRequest(built.outbound.body, built.outbound.bytes, (t) => s.map.isToken(t)),
      partial: Object.freeze({ pending: [], auto: [], chosen: [], kept: [] }),
      correction: Object.freeze({ metricsId, userContent: prepared.built.userContent }),
    });
    return { ok: false, message: COPY.correctionReady, correction: next };
  }

  /**
   * Reads the session's ranges again at their addresses, and the workbook's tables and defined
   * names, and compares them with what the session was built from (spec §7.1). Local reads only:
   * nothing is sent or written. "changed" marks the session stale until Refresh from selection.
   */
  async #recheck(s: Session): Promise<"same" | "changed" | "failed"> {
    try {
      for (const slot of s.ranges) {
        let now: RangeData;
        try {
          now = await this.#adapter.readRange(labelOf(slot.data));
        } catch (e) {
          // The sheet was renamed or deleted: that is a change. Anything else, try again.
          if (this.#hostMessage(e) === COPY.noSuchSheet) return this.#markStale(s);
          return "failed";
        }
        if (!sameCells(slot.data, now)) return this.#markStale(s);
      }
      // Read strictly: a failed read here means "try again", not "the names changed".
      const info = this.#adapter.readWorkbookInfo
        ? await this.#adapter.readWorkbookInfo()
        : {
            tables: await this.#adapter.listTables(),
            names: this.#adapter.listNames ? await this.#adapter.listNames() : s.workbook.names,
          };
      if (!Array.isArray(info.tables) || (info.names !== undefined && !Array.isArray(info.names))) return "failed";
      // A part the session holds as its fail-closed default wasn't read then; there is nothing to compare.
      if (!sameWorkbook(s.workbook, info, s.workbook.read ?? { tables: true, names: true })) return this.#markStale(s);
      return "same";
    } catch {
      return "failed";
    }
  }

  #markStale(s: Session): "changed" {
    if (this.#session === s && !s.stale) {
      s.stale = true;
      this.#emit();
    }
    return "changed";
  }

  /**
   * Writes the checked formula into the sheet. Only the Insert button calls this (invariant 9).
   * `confirmedSpan` is the span the user agreed to replace; it counts only for that exact span.
   */
  async insert(resultId: number, cell: string, fillDown: boolean, opts: { confirmedSpan?: string } = {}): Promise<InsertOutcome> {
    // One insert at a time. A second call while one is in flight (a double click) would pass the
    // same checks and reach the write again; it is refused here, before any read. The flag is
    // released on every path: written, refused, asked to confirm, or failed.
    if (this.#inserting) return fail(COPY.alreadyInserting);
    this.#inserting = true;
    try {
      return await this.#insert(resultId, cell, fillDown, opts);
    } finally {
      this.#inserting = false;
    }
  }

  async #insert(resultId: number, cell: string, fillDown: boolean, opts: { confirmedSpan?: string }): Promise<InsertOutcome> {
    const rec = this.#results.get(resultId);
    if (!rec) return fail(COPY.unexpected);
    const r = rec.view;
    if (r.unresolvedTokens.length > 0) return fail(COPY.insertUnresolved);
    if (r.kind !== "formula" || !r.canInsert || r.formula === null) return fail(COPY.insertBlocked);
    const s = this.#session;
    if (!s || s.id !== rec.sessionId) return fail(COPY.staleResult);
    if (s.stale) return fail(COPY.dataChangedInsert);

    // Check again against the ranges as they are now; fail closed.
    const again = check(r.formula, s.context);
    if (!again.ok) return { ok: false, message: COPY.insertBlocked, reasons: uniq(again.reasons.map((x) => x.detail)) };

    const target = normalizeCell(cell);
    const pos = target ? parseCell(target) : null;
    if (!target || !pos) return fail(COPY.placementInvalid);
    if (insideRanges(s, pos)) return fail(COPY.placementInside);
    // The formula is written for the first data row; filled from any other row, every copy reads the wrong row.
    const first = s.context.firstDataRow;
    if (fillDown && pos.r + 1 !== first) return fail(COPY.fillStart(`${columnPart(target)}${first}`));
    const doFill = fillDown && pos.r + 1 < r.lastRow;
    const span = doFill ? `${target}:${columnPart(target)}${r.lastRow}` : target;
    if (doFill) {
      if (spanInsideRanges(s, pos.c, pos.r, r.lastRow - 1)) return fail(COPY.placementInside);
      // Excel moves every row number without a $ in each filled cell; each copy must stay in the ranges.
      const filled = checkFill(r.formula, s.context, r.lastRow - (pos.r + 1));
      if (!filled.ok) return { ok: false, message: COPY.fillBlocked, reasons: uniq(filled.reasons.map((x) => x.detail)) };
    }

    // The data the answer was built for must still be what's in the workbook (spec §7.1).
    const fresh = await this.#recheck(s);
    if (fresh === "changed") return fail(COPY.dataChangedInsert);
    if (fresh === "failed") return fail(COPY.recheckFailedInsert);

    // Office.js writes can't be undone with Ctrl+Z, so ask before replacing anything. Without the
    // check, or when it fails, ask rather than overwrite.
    if (opts.confirmedSpan !== span) {
      let empty: boolean | null;
      try {
        const res = this.#adapter.isRangeEmpty ? await this.#adapter.isRangeEmpty(r.sheet, span) : null;
        empty = typeof res === "boolean" ? res : null;
      } catch {
        empty = null;
      }
      if (empty !== true) {
        return { ok: false, needsConfirm: true, span, message: empty === false ? COPY.overwriteAsk(span) : COPY.overwriteUnknown(span) };
      }
    }

    // Refresh or a stale check may have happened during the reads above.
    if (this.#session !== s || s.stale) return fail(s.stale ? COPY.dataChangedInsert : COPY.staleResult);

    let out: { address: string; check?: InsertCheck; excelError?: boolean; fillFailed?: boolean };
    try {
      out = await this.#adapter.insertFormula(r.sheet, target, r.formula, doFill, r.lastRow);
    } catch (e) {
      // The adapter throws only when the write itself failed: nothing was written.
      return fail(this.#hostMessage(e) ?? COPY.insertFailed);
    }
    // Written. Whatever the check found, the cells changed: never report this as a failure, and
    // never write again on the user's behalf (Insert asks before replacing, as for any other cells).
    const found = insertCheckOf(out);
    this.#log.update(rec.logId, { inserted: true });
    const patch: Partial<EvalRecord> = { inserted: true };
    if (found === "ok" || found === "error") patch.excelError = found === "error";
    this.#updateMetrics(rec.metricsId, patch);
    this.#emit();
    const shown = localAddress(out?.address, r.sheet) ?? span;
    const message =
      found === "ok"
        ? COPY.inserted(shown)
        : found === "error"
          ? COPY.insertedError(shown)
          : found === "manual"
            ? COPY.insertedManual(shown)
            : out?.fillFailed === true
              ? COPY.insertedFillFailed(shown, span)
              : COPY.insertedUnverified(shown);
    return { ok: true, message, address: shown, check: found };
  }

  /** The formula for the Copy button, only when it passed both gate checks. */
  copyText(resultId: number): string | null {
    const rec = this.#results.get(resultId);
    if (!rec || !rec.view.canInsert || rec.view.formula === null) return null;
    return rec.view.formula;
  }

  markCopied(resultId: number): void {
    const rec = this.#results.get(resultId);
    if (!rec) return;
    this.#updateMetrics(rec.metricsId, { copied: true });
    this.#emit();
  }

  rate(resultId: number, rating: "worked" | "partly" | "wrong"): void {
    const rec = this.#results.get(resultId);
    if (!rec) return;
    this.#updateMetrics(rec.metricsId, { rating });
    this.#emit();
  }

  // ----- internals

  #tooLarge(e: unknown): boolean {
    try {
      return this.#adapter.isTooLarge?.(e) === true;
    } catch {
      return false;
    }
  }

  /**
   * The adapter's too-large message when it says more than the plain one, e.g. "The data around E4
   * has 28,000 cells. ..." for a range grown from one cell; otherwise `fallback`.
   */
  #tooLargeMessage(e: unknown, fallback: string): string {
    const m = this.#hostMessage(e);
    return m !== null && m !== COPY.tooLarge ? m : fallback;
  }

  #hostMessage(e: unknown): string | null {
    try {
      const m = this.#adapter.plainMessage?.(e);
      return typeof m === "string" && m.trim() !== "" ? m.trim() : null;
    } catch {
      return null;
    }
  }

  async #readWorkbook(fallbackSheet: string): Promise<{ sheets: string[]; tables: TableInfo[]; names: string[]; read: WorkbookRead }> {
    // One round trip to Excel when the host can; otherwise, or if it fails, each part with its own default.
    try {
      const info = this.#adapter.readWorkbookInfo ? await this.#adapter.readWorkbookInfo() : null;
      if (info && Array.isArray(info.sheets) && Array.isArray(info.tables) && Array.isArray(info.names)) {
        return { sheets: info.sheets, tables: info.tables, names: info.names, read: { tables: true, names: true } };
      }
    } catch {
      // Read them one at a time below.
    }
    let sheets: string[] = [fallbackSheet];
    let tables: TableInfo[];
    // Unknown names fail closed: every allowed function name counts as shadowed, so the gate refuses
    // functions passed by name (GROUPBY(..., SUM)) rather than guess.
    let names: string[] = [...ALLOWLIST];
    const read: WorkbookRead = { tables: false, names: false };
    try {
      if (this.#adapter.listNames) {
        names = await this.#adapter.listNames();
        read.names = true;
      }
    } catch {
      // Keep the fail-closed default.
    }
    try {
      sheets = await this.#adapter.listSheets();
    } catch {
      // The selection's own sheet is enough to go on.
    }
    try {
      tables = await this.#adapter.listTables();
      read.tables = true;
    } catch {
      tables = [];
    }
    return { sheets, tables, names, read };
  }

  async #addContextData(data: RangeData, hasHeaders: boolean | undefined): Promise<ActionResult> {
    const s = this.#session;
    if (!s) return fail(COPY.noSelection);
    if (!isUsable(data)) return fail(COPY.rangeReadFailed);
    if (s.ranges.some((r) => sameRange(r.data, data.sheet, data.address))) return fail(COPY.rangeDuplicate);
    const workbook = s.workbook.sheets.includes(data.sheet) ? s.workbook : await this.#readWorkbook(data.sheet);
    this.#startSession([...s.ranges, makeSlot(data, hasHeaders)], workbook);
    return { ok: true };
  }

  /**
   * `fresh` is true only when every range was just read from the workbook (Refresh, Use exactly the
   * selected cells). Other changes (headers, lookup ranges) keep a stale session stale.
   */
  #startSession(slots: RangeSlot[], workbook: WorkbookInfo, fresh = false): void {
    const stale = !fresh && (this.#session?.stale ?? false);
    const [main, ...contexts] = slots;
    if (!main) return;
    const context = buildSheetContext(
      main.data,
      main.hasHeaders,
      contexts.map((c) => c.data),
      workbook,
    );
    this.#session = {
      id: ++this.#sessionCounter,
      ranges: slots,
      workbook,
      context,
      // Fresh map for every session; the collision guard runs over every range in it.
      map: createStandInMap(slots.map((r) => r.data)),
      history: [],
      stale,
    };
    this.#version++;
    this.#emit();
  }

  #auditContext(s: Session, specs: RangeSpec[], model: string): AuditContext {
    return buildAuditContext(specs, s.map, { model, canaries: this.#canaries });
  }

  async #sendLogged(
    outbound: AuditedOutbound,
    model: string,
    key: string,
    signal?: AbortSignal,
  ): Promise<{ res: SendOk | SendErr; logId: number }> {
    const at = this.#now().toISOString();
    let res: SendOk | SendErr;
    try {
      res = await this.#send(outbound, key, signal);
    } catch {
      res = { ok: false, message: COPY.unexpected, latencyMs: 0 };
    }
    const entry: LogEntry = {
      at,
      mode: outbound.mode,
      model,
      host: this.#host,
      bytes: outbound.bytes,
      body: outbound.body,
      audit: { ok: true, hits: [], structural: [] },
      inserted: false,
      latencyMs: res.latencyMs,
    };
    if (res.ok) {
      entry.replyRaw = res.raw;
      if (res.tokens) entry.tokens = res.tokens;
      if (res.providerName) entry.providerName = res.providerName;
    } else {
      entry.error = res.message;
    }
    return { res, logId: this.#log.add(entry) };
  }

  /**
   * A blocked request is logged without its body: the body holds the private value that blocked it.
   * What it matched stays beside the entry for the Log screen only (LogLocal), never in it.
   */
  #logBlocked(model: string, mode: ProdMode, bytes: number, result: AuditResult, headline: string, matches: readonly LocalMatch[]): void {
    const entry: LogEntry = {
      at: this.#now().toISOString(),
      mode,
      model,
      host: this.#host,
      bytes,
      body: "",
      audit: {
        ok: false,
        hits: result.hits.map((h) => ({ column: h.column, variant: h.variant })),
        structural: (result.structural ?? []).map((c) => ({ column: c.column })),
      },
      inserted: false,
      error: headline,
    };
    if (result.error !== undefined) entry.audit.error = result.error;
    this.#log.add(entry, matches.length > 0 ? { matches } : undefined);
  }

  #record(s: Session, model: string, mode: ProdMode, bytes: number, result: AuditResult): EvalRecord {
    const main = s.ranges[0]!;
    const treatments: Record<Treatment, number> = { as_is: 0, stand_in: 0, range: 0, month: 0, exclude: 0 };
    for (const p of main.policies) treatments[effectiveTreatment(p)]++;
    return {
      build: this.#build,
      model,
      mode,
      rows: dataRows(main),
      columns: main.columns.length,
      privateColumns: main.policies.filter((p) => p.private).length,
      treatments,
      bytes,
      audit: result.ok ? "pass" : "block",
      auditReasons: auditReasons(result),
      gateCodes: [],
      retries: 0,
      inserted: false,
      copied: false,
    };
  }

  #updateMetrics(id: number | null, patch: Partial<EvalRecord>): void {
    if (id !== null) this.#metrics.update(id, patch);
  }

  /** First empty column right of the selection, at the first data row; the next column when unknown. */
  async #defaultCell(s: Session): Promise<string | null> {
    const main = s.ranges[0];
    if (!main) return null;
    const next = defaultPlacementCell(s.context, main.data);
    if (!this.#adapter.firstEmptyColumn || next === null) return next;
    const from = main.data.columnIndex + Math.max(1, columnCount(main.data));
    try {
      const letter = await this.#adapter.firstEmptyColumn(s.context.sheet, from, s.context.firstDataRow, s.context.lastDataRow);
      const cell = typeof letter === "string" ? normalizeCell(`${letter}${s.context.firstDataRow}`) : null;
      const pos = cell ? parseCell(cell) : null;
      if (cell && pos && pos.c >= from && !insideRanges(s, pos)) return cell;
    } catch {
      // Reading failed; the next column is a fine default.
    }
    return next;
  }

  #makeResult(s: Session, reply: ModelReply, mode: ProdMode, retries: number, fallbackCell: string | null): FlowResult {
    const unknown = new Set<string>();
    let unresolved: string[] = [];
    let cautions: string[] = [];
    let formula: string | null = null;
    let modelFormula: string | null = null;
    let gate: GateResult | null = null;
    let canInsert = false;

    if (reply.kind === "formula" && typeof reply.formula === "string") {
      modelFormula = reply.formula;
      const firstPass = check(modelFormula, s.context);
      if (firstPass.ok) {
        const restored = restoreFormula(modelFormula, s.map);
        for (const t of restored.unknownTokens) unknown.add(t);
        // Token-shaped text that is in the ranges (with the NYMFORM_ prefix in use) may be the sheet's
        // own value: that stays a warning. Anything else can't be put back, so it blocks.
        unresolved = restored.unknownTokens.filter((t) => !s.map.isReserved(t));
        const secondPass = check(restored.formula, s.context);
        formula = restored.formula;
        gate = combineGates(firstPass, secondPass);
        // A safe formula isn't necessarily a correct one: a stand-in left in a string would be
        // compared as text against the real values, so it can't be inserted.
        canInsert = firstPass.ok && secondPass.ok && unresolved.length === 0;
        cautions = restoreCautions(s, restored.restoredTokens);
      } else {
        // Rejected before restoring: the formula stays in stand-in form.
        formula = modelFormula;
        gate = firstPass;
      }
    }

    const explanation = restoreText(reply.explanation, s.map);
    for (const t of explanation.unknownTokens) unknown.add(t);
    const assumptions = reply.assumptions.map((a) => {
      const r = restoreText(a, s.map);
      for (const t of r.unknownTokens) unknown.add(t);
      return r.text;
    });

    return Object.freeze({
      id: ++this.#counter,
      kind: reply.kind,
      mode,
      formula,
      modelFormula,
      gate,
      gateStatus: gate ? gateStatus(gate) : null,
      canInsert,
      explanation: explanation.text,
      assumptions: Object.freeze(assumptions),
      unknownTokens: Object.freeze([...unknown]),
      unresolvedTokens: Object.freeze([...unresolved]),
      cautions: Object.freeze(cautions),
      wholeColumns: Object.freeze(formula !== null && canInsert ? wholeColumns(formula) : []),
      placement: reply.kind === "formula" ? choosePlacement(s, reply.placement, fallbackCell) : null,
      sheet: s.context.sheet,
      lastRow: s.context.lastDataRow,
      retries,
    });
  }

  #snapshot(): FlowState {
    const s = this.#session;
    return {
      sessionId: s?.id ?? 0,
      version: this.#version,
      stale: s?.stale ?? false,
      ranges: s ? s.ranges.map(rangeView) : [],
      context: s?.context ?? null,
      historyMessages: s?.history.length ?? 0,
      logSize: this.#log.size,
      revision: ++this.#revision,
    };
  }

  #emit(): void {
    this.#state = this.#snapshot();
    for (const l of [...this.#listeners]) {
      try {
        l();
      } catch {
        // A listener's failure must not stop the flow.
      }
    }
  }
}

// ---------------------------------------------------------------------------------------------
// Helpers (exported where the UI or tests use them)

function fail(message: string): { ok: false; message: string } {
  return { ok: false, message };
}

/** A range the flow can work with. The gate, placement and Insert need an address like "A1:F501", never "4:4". */
function isUsable(data: RangeData | null | undefined): data is RangeData {
  return (
    !!data &&
    typeof data.sheet === "string" &&
    typeof data.address === "string" &&
    parseLocalRange(data.address) !== null &&
    rowCount(data) > 0 &&
    columnCount(data) > 0
  );
}

function labelOf(data: RangeData): string {
  return qualify(data.sheet, data.address);
}

function sameRange(data: RangeData, sheet: string, address: string): boolean {
  const a = parseLocalRange(data.address);
  const b = parseLocalRange(address);
  const sameRect = a && b ? formatRect(a) === formatRect(b) : data.address.toUpperCase() === address.toUpperCase();
  return data.sheet.toLowerCase() === sheet.toLowerCase() && sameRect;
}

function dataRows(slot: RangeSlot): number {
  return Math.max(0, rowCount(slot.data) - (slot.hasHeaders ? 1 : 0));
}

function specsOf(s: Session): RangeSpec[] {
  return s.ranges.map((r) => ({ data: r.data, hasHeaders: r.hasHeaders, columns: r.columns, policies: r.policies }));
}

/**
 * Cautions for values put back into a formula that Excel may read differently than the sheet holds
 * them (SECURITY.md, "Restored values are text"): a stand-in for a number or date comes back as its
 * shown text, and *, ? and ~ act as wildcards in criteria. The formula is still the user's call.
 */
function restoreCautions(s: Session, tokens: readonly string[]): string[] {
  if (tokens.length === 0) return [];
  let notText: Map<string, string> | null = null;
  const out: string[] = [];
  for (const token of tokens) {
    const value = s.map.valueOf(token);
    if (value === undefined) continue;
    notText ??= notTextPrivateValues(s);
    const letter = notText.get(value);
    if (letter !== undefined) out.push(COPY.restoredNumber(token, letter));
    if (/[*?~]/u.test(value)) out.push(COPY.restoredWildcard(token));
  }
  return uniq(out);
}

/**
 * True when `b`, read again, holds what `a` held: the same sheet and address, and cell for cell the
 * same formulas, number formats, and stored values and types of constants. A formula's result isn't
 * compared, since volatile functions (TODAY, RAND) change it on every read, nor is display text,
 * which a narrower column turns into ####.
 */
export function sameCells(a: RangeData, b: RangeData): boolean {
  if (!sameRange(a, b.sheet, b.address)) return false;
  const rows = rowCount(a);
  const cols = columnCount(a);
  if (rowCount(b) !== rows || columnCount(b) !== cols) return false;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const f = a.formulas[r]?.[c];
      if (f !== b.formulas[r]?.[c]) return false;
      if (a.numberFormat[r]?.[c] !== b.numberFormat[r]?.[c]) return false;
      if (typeof f === "string" && f.startsWith("=")) continue;
      if (a.values[r]?.[c] !== b.values[r]?.[c] || a.valueTypes[r]?.[c] !== b.valueTypes[r]?.[c]) return false;
    }
  }
  return true;
}

/**
 * True when the workbook's tables (name, sheet, address) and defined names are the same. They are
 * what the formula gate checks names and structured references against. Sheets aren't compared: a
 * new sheet changes nothing, and a renamed or deleted one fails the read of its range instead.
 */
export function sameWorkbook(
  a: { tables: readonly TableInfo[]; names?: readonly string[] },
  b: { tables: readonly TableInfo[]; names?: readonly string[] },
  compare: { tables: boolean; names: boolean } = { tables: true, names: true },
): boolean {
  const tables = (w: { tables: readonly TableInfo[] }) =>
    w.tables.map((t) => `${t.name.toUpperCase()}\u0000${t.sheet.toUpperCase()}\u0000${t.address.toUpperCase()}`).sort().join("\u0001");
  const names = (w: { names?: readonly string[] }) =>
    [...(w.names ?? [])]
      .filter((n) => !isExcelInternalName(n))
      .map((n) => n.toUpperCase())
      .sort()
      .join("\u0001");
  return (!compare.tables || tables(a) === tables(b)) && (!compare.names || names(a) === names(b));
}

/**
 * Names Excel keeps for itself: `_xlfn.GROUPBY` and `_xleta.SUM` appear when a formula using a newer
 * function is written (seen in Excel for Mac, 2026-09-28), `_xlnm.` holds print areas and filters.
 * They change on every such Insert, and none is a name a formula can use to hide a call: the gate
 * compares names whole, so `_xleta.SUM` doesn't shadow SUM.
 */
export function isExcelInternalName(name: string): boolean {
  return /^_xl(?:fn|eta|ws|pm|nm|udf)\./iu.test(name.replace(/^.*!/u, ""));
}

const WHOLE_COLUMN_RE = /^\$?[A-Z]{1,3}:\$?[A-Z]{1,3}$/iu;

/** Whole-column references in `formula` ("E:E", "Orders!$B:$C"), as written, each once. */
export function wholeColumns(formula: string): string[] {
  try {
    return uniq(
      tokenize(formula)
        .filter((t) => t.type === "ref" && WHOLE_COLUMN_RE.test(t.text.slice(t.text.lastIndexOf("!") + 1)))
        .map((t) => t.text),
    );
  } catch {
    return [];
  }
}

/** Private values stored as numbers (dates included) or booleans, with the first column holding one. */
function notTextPrivateValues(s: Session): Map<string, string> {
  const out = new Map<string, string>();
  for (const spec of specsOf(s)) {
    for (const cell of privateCellsOf(spec, s.map)) {
      if (out.has(cell.value)) continue;
      const type = spec.data.valueTypes[cell.row]?.[columnIndex(cell.letter) - spec.data.columnIndex];
      if (type === "Double" || type === "Boolean") out.set(cell.value, cell.letter);
    }
  }
  return out;
}

/**
 * The user's header choice for a range, kept across a refresh when `data` is the same range (sheet
 * and address); undefined when they made none or the range is another one.
 */
function chosenHeaders(slot: RangeSlot | undefined, data: RangeData): boolean | undefined {
  if (!slot?.headersChosen || !sameRange(slot.data, data.sheet, data.address)) return undefined;
  return slot.hasHeaders;
}

function firstRowIsData(data: RangeData): boolean {
  try {
    return firstRowLooksLikeData(data);
  } catch {
    // Treat it as data: nothing from it goes out as a header.
    return true;
  }
}

/**
 * True when the header cell of column `i` looks like a data value (a number, an email, an ID),
 * compared with the data below it the same way row 1 is.
 */
function headerLooksLikeData(slot: Pick<RangeSlot, "data" | "hasHeaders">, i: number): boolean {
  if (!slot.hasHeaders) return false;
  try {
    return headerCellLooksLikeData(slot.data, i);
  } catch {
    return true;
  }
}

/**
 * A range's columns and default policies. `choice` is the header checkbox as the user set it;
 * without one, row 1 is a header row unless it looks like data.
 */
function makeSlot(data: RangeData, choice: boolean | undefined, previous?: RangeSlot): RangeSlot {
  const hasHeaders = choice ?? !firstRowIsData(data);
  const columns = inferSchema(data, hasHeaders);
  let suggestions: Suggestion[] = [];
  try {
    suggestions = suggestPolicies(data, columns, hasHeaders);
  } catch {
    suggestions = [];
  }
  const autoAliases = new Set<string>();
  const policies = columns.map((col, i) => {
    const kept = keptPolicy(previous, data, hasHeaders, col);
    if (kept) {
      if (previous?.autoAliases.has(col.letter) && kept.alias === autoAlias(col.letter)) autoAliases.add(col.letter);
      return kept;
    }
    const sug = suggestions.find((x) => x.letter === col.letter);
    // No suggestion means private: fail closed.
    const policy: ColumnPolicy = {
      letter: col.letter,
      private: sug?.private ?? true,
      treatment: sug?.treatment ?? defaultTreatment(col.type),
    };
    if (!treatmentsFor(policy.private).includes(policy.treatment)) {
      policy.treatment = policy.private ? defaultTreatment(col.type) : "as_is";
    }
    if (sug?.alias) policy.alias = sug.alias;
    // A private column whose header looks like data: send a neutral header unless the user changes it.
    else if (policy.private && headerLooksLikeData({ data, hasHeaders }, i)) {
      policy.alias = autoAlias(col.letter);
      autoAliases.add(col.letter);
    }
    return policy;
  });
  return { data, hasHeaders, headersChosen: choice !== undefined, columns, policies, suggestions, autoAliases };
}

function autoAlias(letter: string): string {
  return `Column ${letter}`;
}

/** On refresh, a column keeps the user's choices when it is the same column with the same header and type. */
function keptPolicy(previous: RangeSlot | undefined, data: RangeData, hasHeaders: boolean, col: ColumnInfo): ColumnPolicy | null {
  if (!previous || previous.hasHeaders !== hasHeaders || previous.data.sheet !== data.sheet) return null;
  const i = previous.columns.findIndex((c) => c.letter === col.letter);
  const old = previous.columns[i];
  const policy = previous.policies[i];
  if (!old || !policy || old.header !== col.header || old.type !== col.type) return null;
  return { ...policy };
}

function samePolicy(a: ColumnPolicy, b: ColumnPolicy): boolean {
  return a.private === b.private && a.treatment === b.treatment && (a.alias ?? "") === (b.alias ?? "") && (a.note ?? "") === (b.note ?? "");
}

function rangeView(slot: RangeSlot): RangeView {
  const d = slot.data;
  const rows = dataRows(slot);
  const columns = slot.columns.map((col, i): ColumnView => {
    const policy = slot.policies[i] ?? { letter: col.letter, private: true, treatment: defaultTreatment(col.type) };
    const sug = slot.suggestions.find((x) => x.letter === col.letter);
    return {
      letter: col.letter,
      header: col.header,
      type: col.type,
      stats: col.stats,
      policy: { ...policy },
      hint: sug && sug.reasons.length > 0 ? COPY.hint(sug.reasons) : null,
      suggestedPrivate: sug?.private ?? true,
    };
  });
  return {
    label: labelOf(d),
    sheet: d.sheet,
    address: d.address,
    table: d.table?.name ? d.table.name : null,
    hasHeaders: slot.hasHeaders,
    // Off without the user's choice only when row 1 looked like data.
    notice: !slot.headersChosen && !slot.hasHeaders ? COPY.headerRowData(d.rowIndex + 1) : null,
    dataRows: rows,
    summary: `${d.sheet}, ${d.address} · ${rows.toLocaleString("en-US")} ${rows === 1 ? "row" : "rows"}`,
    resolution: resolutionView(d.resolution),
    columns,
  };
}

/** The line under the summary that says how the range came from the selection. */
function resolutionView(r: Resolution | undefined): RangeView["resolution"] {
  switch (r?.kind) {
    case "table":
      return { kind: r.kind, text: COPY.fromTable(r.table ?? "", r.from), exact: true };
    case "filter":
      return { kind: r.kind, text: COPY.fromFilter(r.from), exact: true };
    case "region":
      return { kind: r.kind, text: COPY.fromRegion(r.from), exact: true };
    case "trimmed": {
      const to = trimmedTo(r.from, r.exact);
      // Excel calls the whole sheet 1:1048576, a name users never see.
      return { kind: r.kind, text: to === "sheet" ? COPY.trimmedSheet : COPY.trimmed(r.from, to), exact: false };
    }
    default:
      return null;
  }
}

/**
 * What whole rows or columns were trimmed to: 1:3 to the columns in use, A:F to the rows in use, the
 * whole sheet to its cells in use. When the kept part is also narrower along the selection's own
 * length (H:K kept as I2:J30), both were trimmed: to the cells in use.
 */
function trimmedTo(from: string, kept: string): "columns" | "rows" | "cells" | "sheet" {
  const rows = /^\$?(\d+):\$?(\d+)$/.exec(from);
  const cols = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(from);
  const k = parseLocalRange(kept);
  // A whole sheet is reported as 1:1048576; both are trimmed then.
  if (rows) {
    if (Number(rows[1]) === 1 && Number(rows[2]) === MAX_ROWS) return "sheet";
    return k && (k.r1 !== Number(rows[1]) - 1 || k.r2 !== Number(rows[2]) - 1) ? "cells" : "columns";
  }
  if (cols) {
    const c1 = columnIndex(cols[1]!.toUpperCase());
    const c2 = columnIndex(cols[2]!.toUpperCase());
    if (c1 === 0 && c2 === MAX_COLUMNS - 1) return "sheet";
    return k && (k.c1 !== c1 || k.c2 !== c2) ? "cells" : "rows";
  }
  return "cells";
}

function uniq<T>(items: T[]): T[] {
  return [...new Set(items)];
}

export function auditReasons(result: AuditResult): EvalRecord["auditReasons"] {
  if (result.ok) return [];
  const reasons: EvalRecord["auditReasons"] = uniq(result.hits.map((h) => h.variant));
  if ((result.structural ?? []).length > 0) reasons.push("structural");
  if (reasons.length === 0) reasons.push("error");
  return reasons;
}

/**
 * Plain status for the What gets sent screen. Names columns by letter only. `origins` says where
 * each hit's text came from (the audit's locations, or the Log's matches): when every hit is a
 * column's original header, renamed under Header sent, the headline says so instead of calling it
 * a value.
 */
export function auditStatus(
  result: AuditResult,
  origins: readonly { column: string; variant: AuditVariant; origin: AuditOrigin }[] = [],
): AuditStatus {
  if (result.ok) return { ok: true, headline: COPY.auditOk };
  const columns = uniq([
    ...(result.structural ?? []).map((c) => c.column),
    ...result.hits.filter((h) => h.variant !== "canary").map((h) => h.column),
  ]);
  // A hit is a renamed header's only when every match recorded for it came from the header: the
  // audit also records the column's values when they match the same way (see AuditOutcome).
  const fromHeader = (h: { column: string; variant: AuditVariant }) => {
    const mine = origins.filter((o) => o.column === h.column && o.variant === h.variant);
    return mine.length > 0 && mine.every((o) => o.origin === "alias");
  };
  const headersOnly = (result.structural ?? []).length === 0 && result.hits.length > 0 && result.hits.every(fromHeader);
  const out: AuditStatus = { ok: false, headline: "", action: COPY.auditAction };
  if (headersOnly) out.headline = COPY.auditBlockedAlias(columns);
  else if (columns.length > 0) out.headline = COPY.auditBlockedColumns(columns);
  else if (result.hits.some((h) => h.variant === "canary")) out.headline = COPY.auditBlockedCanary;
  else out.headline = COPY.auditBlockedError(result.error ?? "the check couldn't finish.");
  if ((result.hits.length > 0 || (result.structural ?? []).length > 0) && result.error) out.detail = result.error;
  return out;
}

export function combineGates(a: GateResult, b: GateResult): GateResult {
  const seen = new Set<string>();
  const reasons: GateResult["reasons"] = [];
  for (const r of [...a.reasons, ...b.reasons]) {
    const k = `${r.code}|${r.detail}`;
    if (seen.has(k)) continue;
    seen.add(k);
    reasons.push({ code: r.code, detail: r.detail });
  }
  return { ok: a.ok && b.ok, reasons };
}

export function gateStatus(gate: GateResult): GateStatus {
  if (gate.ok) return { ok: true, headline: COPY.gateOk, reasons: [] };
  return { ok: false, headline: COPY.gateBlocked, reasons: uniq(gate.reasons.map((r) => r.detail)) };
}

/** "g2" or "$G$2" -> "G2"; null for anything that isn't one cell on this sheet. */
export function normalizeCell(input: string): string | null {
  if (typeof input !== "string") return null;
  const t = input.trim().replace(/\$/g, "").toUpperCase();
  if (!/^[A-Z]{1,3}\d{1,7}$/.test(t)) return null;
  return parseCell(t) ? t : null;
}

function columnPart(cell: string): string {
  return cell.replace(/\d+$/, "");
}

/** The adapter's address without the sheet name when it is on `sheet`. */
function localAddress(address: string | undefined, sheet: string): string | null {
  if (typeof address !== "string" || address.trim() === "") return null;
  const { sheet: s, address: local } = splitSheetAddress(address.trim());
  if (s !== null && s !== sheet) return address.trim();
  return local.replace(/\$/g, "");
}

interface Area {
  sheet: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

function areas(s: Session): Area[] {
  return s.ranges.map((r) => ({
    sheet: r.data.sheet,
    r1: r.data.rowIndex,
    c1: r.data.columnIndex,
    r2: r.data.rowIndex + Math.max(1, rowCount(r.data)) - 1,
    c2: r.data.columnIndex + Math.max(1, columnCount(r.data)) - 1,
  }));
}

function insideRanges(s: Session, pos: { r: number; c: number }): boolean {
  return spanInsideRanges(s, pos.c, pos.r, pos.r);
}

/** True when any cell of column `c`, rows r1..r2 (0-based), is inside a selected range on this sheet. */
function spanInsideRanges(s: Session, c: number, r1: number, r2: number): boolean {
  const sheet = s.context.sheet.toLowerCase();
  return areas(s).some((a) => a.sheet.toLowerCase() === sheet && c >= a.c1 && c <= a.c2 && r1 <= a.r2 && r2 >= a.r1);
}

/**
 * The check an adapter reported. An adapter that doesn't say is read from `excelError`: true is an
 * error, false is fine, and neither means the result wasn't checked.
 */
export function insertCheckOf(out: { check?: unknown; excelError?: unknown; fillFailed?: unknown } | null | undefined): InsertCheck {
  if (out?.fillFailed === true) return "unverified";
  const c = out?.check;
  if (c === "ok" || c === "error" || c === "manual" || c === "unverified") return c;
  if (out?.excelError === true) return "error";
  if (out?.excelError === false) return "ok";
  return "unverified";
}

/** The cells Fill down writes from `cell` to `lastRow`, e.g. "G2:G501"; null when the cell isn't one cell or there is nothing to fill. */
export function fillSpan(cell: string, lastRow: number): string | null {
  const target = normalizeCell(cell);
  const pos = target ? parseCell(target) : null;
  if (!target || !pos || pos.r + 1 >= lastRow) return null;
  return `${target}:${columnPart(target)}${lastRow}`;
}

/** First column right of the selection, at the first data row. Null at the sheet's last column. */
export function defaultPlacementCell(context: SheetContext, selection: RangeData): string | null {
  const c = selection.columnIndex + Math.max(1, columnCount(selection));
  if (c >= MAX_COLUMNS) return null;
  return `${columnLetter(c)}${context.firstDataRow}`;
}

/**
 * The reply's cell when it is a real cell outside the selected ranges; otherwise the default. With
 * fill_down, only the reply's column is used: fill down starts at the first data row, and every cell
 * down to the last data row must be outside the ranges.
 */
function choosePlacement(
  s: Session,
  placement: ModelReply["placement"],
  fallback: string | null,
): { cell: string; fillDown: boolean } | null {
  const fillDown = placement?.fill_down === true;
  const proposed = placement ? normalizeCell(placement.cell) : null;
  const pos = proposed ? parseCell(proposed) : null;
  if (proposed && pos) {
    if (!fillDown && !insideRanges(s, pos)) return { cell: proposed, fillDown };
    const first = s.context.firstDataRow - 1;
    const last = Math.max(first, s.context.lastDataRow - 1);
    if (fillDown && !spanInsideRanges(s, pos.c, first, last)) return { cell: `${columnLetter(pos.c)}${first + 1}`, fillDown };
  }
  return fallback ? { cell: fallback, fillDown } : null;
}

// ---------------------------------------------------------------------------------------------
// The readable request view

/** Splits text into plain runs and stand-in tokens the session map knows. */
export function segments(text: string, isToken: (s: string) => boolean): Segment[] {
  const out: Segment[] = [];
  let last = 0;
  for (const m of text.matchAll(new RegExp(TOKEN_RE.source, "g"))) {
    if (!isToken(m[0])) continue;
    if (m.index > last) out.push({ kind: "text", text: text.slice(last, m.index) });
    out.push({ kind: "token", text: m[0] });
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push({ kind: "text", text: text.slice(last) });
  return out;
}

/** Objects without arrays inside stay on one line up to this length (a column's description fits). */
const INLINE_WIDTH = 200;
/** Nesting deeper than this is shown as the raw text, so a deeply nested reply can't exhaust the stack. */
const MAX_DEPTH = 20;

function isPlain(v: unknown): boolean {
  return v === null || typeof v !== "object";
}

function hasArray(v: unknown): boolean {
  if (Array.isArray(v)) return true;
  if (v === null || typeof v !== "object") return false;
  return Object.values(v as Record<string, unknown>).some(hasArray);
}

/** One line with a space after each separator, for reading. */
function spacedLine(v: unknown): string {
  if (isPlain(v)) return JSON.stringify(v) ?? "null";
  if (Array.isArray(v)) return `[${v.map(spacedLine).join(", ")}]`;
  const entries = Object.entries(v as Record<string, unknown>);
  if (entries.length === 0) return "{}";
  return `{ ${entries.map(([k, x]) => `${JSON.stringify(k)}: ${spacedLine(x)}`).join(", ")} }`;
}

/** True when arrays and objects nest more than `limit` levels deep. Iterative, so any depth is safe. */
function deeperThan(value: unknown, limit: number): boolean {
  const stack: [unknown, number][] = [[value, 0]];
  while (stack.length > 0) {
    const [v, depth] = stack.pop()!;
    if (isPlain(v)) continue;
    if (depth >= limit) return true;
    for (const x of Object.values(v as Record<string, unknown>)) stack.push([x, depth + 1]);
  }
  return false;
}

/**
 * JSON for reading: objects without arrays inside (a column, a range) and arrays of plain values
 * (a row) stay on one line; everything else is indented. Null when it nests deeper than MAX_DEPTH
 * levels; show the raw text then.
 */
export function prettyJson(value: unknown): string | null {
  return deeperThan(value, MAX_DEPTH) ? null : pretty(value, "");
}

function pretty(value: unknown, indent: string): string {
  if (isPlain(value)) return JSON.stringify(value) ?? "null";
  const inner = `${indent}  `;
  if (Array.isArray(value)) {
    if (value.length === 0) return "[]";
    if (value.every(isPlain)) return spacedLine(value);
    return `[\n${value.map((v) => inner + pretty(v, inner)).join(",\n")}\n${indent}]`;
  }
  const entries = Object.entries(value as Record<string, unknown>);
  if (entries.length === 0) return "{}";
  const line = spacedLine(value);
  if (!hasArray(value) && line.length <= INLINE_WIDTH) return line;
  return `{\n${entries.map(([k, v]) => `${inner}${JSON.stringify(k)}: ${pretty(v, inner)}`).join(",\n")}\n${indent}}`;
}

function tryJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

/**
 * True when `text` is exactly what JSON.stringify writes for `parsed`. Only then does a formatted
 * view show everything: JSON.parse drops a repeated key and rounds a long number.
 */
function readsBack(parsed: unknown, text: string): boolean {
  try {
    return JSON.stringify(parsed) === text;
  } catch {
    return false;
  }
}

function messageView(role: MessageView["role"], content: string, isToken: (s: string) => boolean): MessageView {
  const parsed = tryJson(content);
  const isObj = parsed !== null && typeof parsed === "object" && !Array.isArray(parsed);
  const text = isObj && readsBack(parsed, content) ? prettyJson(parsed) : null;
  if (text === null) return { role, description: false, segments: segments(content, isToken) };
  const description = role === "user" && "nymform" in (parsed as Record<string, unknown>);
  return { role, description, segments: segments(text, isToken) };
}

/** Parses the exact body into a readable view. Only what is in the body appears. */
export function formatRequest(body: string, bytes: number, isToken: (s: string) => boolean): RequestView {
  const view: RequestView = { model: "", params: [], system: "", earlier: [], current: null, bytes, exact: body };
  const parsed = tryJson(body);
  if (parsed === null || typeof parsed !== "object" || Array.isArray(parsed) || !readsBack(parsed, body)) {
    view.current = { role: "user", description: false, segments: segments(body, isToken) };
    return view;
  }
  const obj = parsed as Record<string, unknown>;
  view.model = typeof obj.model === "string" ? obj.model : "";
  for (const [key, value] of Object.entries(obj)) {
    if (key === "model" || key === "messages") continue;
    view.params.push({ key, value: JSON.stringify(value) });
  }
  const messages = Array.isArray(obj.messages) ? obj.messages : [];
  const list: MessageView[] = [];
  messages.forEach((m: unknown, i: number) => {
    if (m === null || typeof m !== "object") return;
    const { role, content } = m as { role?: unknown; content?: unknown };
    if (typeof content !== "string") return;
    if (i === 0 && role === "system") {
      view.system = content;
      return;
    }
    const r = role === "assistant" ? "assistant" : role === "system" ? "system" : "user";
    list.push(messageView(r, content, isToken));
  });
  view.current = list.pop() ?? null;
  view.earlier = list;
  return view;
}

// ---------------------------------------------------------------------------------------------
// What a blocked request matched, for the Log screen only (see LogLocal). Read from the blocked
// body while the flow still holds it; the text never goes into the log entry, an export, the
// history or a later request.

/** Characters of context shown on each side of a match. */
const CONTEXT_CHARS = 32;
/** Characters of the body read on each side before unescaping: an escaped character takes up to 7. */
const CONTEXT_WINDOW = CONTEXT_CHARS * 8;

const JSON_ESCAPES: Readonly<Record<string, string>> = { '"': '"', "\\": "\\", "/": "/", b: "\b", f: "\f", n: "\n", r: "\r", t: "\t" };

/** Undoes one level of JSON string escapes, best effort: an escape cut off at an edge stays as it is. */
function unescapeJson(s: string): string {
  return s.replace(/\\(?:u([0-9a-fA-F]{4})|(["\\/bfnrt]))/g, (_, hex: string | undefined, c: string | undefined) =>
    hex !== undefined ? String.fromCharCode(parseInt(hex, 16)) : JSON_ESCAPES[c!]!,
  );
}

/** The request's turns as sent, and which one is the current question (the last description). */
interface SentTurns {
  messages: { role: string; content: string }[];
  current: number;
  descriptions: Map<number, Record<string, unknown> | null>;
}

function readTurns(body: string): SentTurns | null {
  const parsed = tryJson(body) as { messages?: unknown } | undefined;
  if (!parsed || typeof parsed !== "object" || !Array.isArray(parsed.messages)) return null;
  const messages = (parsed.messages as unknown[]).map((m) => {
    const { role, content } = (m ?? {}) as { role?: unknown; content?: unknown };
    return { role: typeof role === "string" ? role : "", content: typeof content === "string" ? content : "" };
  });
  const turns: SentTurns = { messages, current: -1, descriptions: new Map() };
  messages.forEach((m, i) => {
    if (m.role === "user" && descriptionOf(turns, i) !== null) turns.current = i;
  });
  return turns;
}

function descriptionOf(turns: SentTurns, i: number): Record<string, unknown> | null {
  if (!turns.descriptions.has(i)) {
    const v = tryJson(turns.messages[i]?.content ?? "");
    const ok = v !== null && typeof v === "object" && !Array.isArray(v) && "nymform" in v;
    turns.descriptions.set(i, ok ? (v as Record<string, unknown>) : null);
  }
  return turns.descriptions.get(i) ?? null;
}

function letterAt(columns: unknown, i: unknown, sampleRow: boolean): string | null {
  if (!Array.isArray(columns) || typeof i !== "number") return null;
  // A sample row has a cell for each column that isn't left out, in order.
  const list = sampleRow ? columns.filter((c) => (c as { treatment?: unknown })?.treatment !== "exclude") : columns;
  const letter = (list[i] as { letter?: unknown } | undefined)?.letter;
  return typeof letter === "string" ? letter : null;
}

/** Where a match is, in words, and the column whose header holds it (see LocalMatch.inHeader). */
interface Place {
  where: string;
  inHeader?: string;
}

/**
 * Where a match is, in words, from the audit's field path ("your question", "the header of column
 * J"). Column letters and sheet names only, never text from a cell. A column is labelled as the
 * audit labels it: "J" in the selection, "Regions!J" in a lookup range.
 */
function describeField(field: AuditField | null, turns: SentTurns | null): Place {
  const fallback: Place = { where: "the request" };
  if (!field || !turns) return fallback;
  if (field.role === "assistant") return { where: "an earlier reply" };
  const d = field.role === "user" ? descriptionOf(turns, field.message) : null;
  if (!d) return fallback;
  const current = field.message === turns.current;
  const [first, second, ...rest] = field.path;
  if (first === "task") return { where: current ? "your question" : "an earlier question" };
  let place: Place | null;
  if (first === "allowed_ranges") place = { where: "the allowed ranges" };
  else if (first === "context_ranges") {
    const block: unknown = Array.isArray(d.context_ranges) && typeof second === "number" ? d.context_ranges[second] : undefined;
    const sheet = (block as { sheet?: unknown } | undefined)?.sheet;
    place = typeof sheet === "string" ? placeIn(block, rest, sheet) : null;
  } else {
    place = placeIn(d, field.path, null);
  }
  if (place === null) return fallback;
  return current ? place : { ...place, where: `${place.where} in an earlier request` };
}

/** Where a path inside one described range points: in the selection (`sheet` null) or in a lookup range on `sheet`. */
function placeIn(block: unknown, path: readonly (string | number)[], sheet: string | null): Place | null {
  const columns = (block as { columns?: unknown } | undefined)?.columns;
  const [key, i, part] = path;
  const label = (letter: string | null) => (letter === null || sheet === null ? letter : qualify(sheet, letter));
  if (key === "sheet") return { where: sheet === null ? "the sheet name" : "a lookup range's sheet name" };
  if (key === "table") return { where: sheet === null ? "the table name" : "a lookup range's table name" };
  if (key === "rows") {
    const col = label(letterAt(columns, part, true));
    if (col === null) return null;
    return { where: sheet === null ? `a sample row in column ${col}` : `a sample row in a lookup range (${col})` };
  }
  const col = key === "columns" ? label(letterAt(columns, i, false)) : null;
  if (col === null) return null;
  if (part === "header") return { where: sheet === null ? `the header of column ${col}` : `a lookup range's header (${col})`, inHeader: col };
  if (part === "note") return { where: sheet === null ? `the note on column ${col}` : `a lookup range's note (${col})` };
  return null;
}

/** The number of backslashes right before position i of `s`. */
function backslashesBefore(s: string, i: number): number {
  let n = 0;
  while (i - n - 1 >= 0 && s[i - n - 1] === "\\") n++;
  return n;
}

/**
 * The body's text on each side of a match, up to the ends of the JSON string it is in (and at most
 * CONTEXT_WINDOW characters), still escaped. At depth 1 (the message's own string) a quote with an
 * even number of backslashes before it is an end. At depth 2 (a description's or a reply's string
 * inside it) an end is written \" and a quote in the text \\\", each after the text's own
 * backslashes, written \\\\: so an end has 1, 5, 9… backslashes before it.
 */
function contextOf(body: string, start: number, end: number, depth: 1 | 2): { before: string; after: string } {
  const ends = (i: number) => body[i] === '"' && backslashesBefore(body, i) % (depth * 2) === depth - 1;
  let from = Math.max(0, start - CONTEXT_WINDOW);
  for (let i = start - 1; i >= from; i--) {
    if (ends(i)) {
      from = i + 1;
      break;
    }
  }
  let to = Math.min(body.length, end + CONTEXT_WINDOW);
  for (let i = end; i < to; i++) {
    if (ends(i)) {
      // At depth 2 the backslash of \" belongs to the quote.
      to = Math.max(end, i - (depth - 1));
      break;
    }
  }
  return { before: body.slice(from, start), after: body.slice(end, to) };
}

/**
 * Keeps the last `n` characters of the text before a match, with "…" when some were cut. A cut
 * inside a word moves to the next space, when there is one, so the context starts at a word.
 */
function tail(s: string, n: number): string {
  const chars = Array.from(s);
  if (chars.length <= n) return s;
  let kept = chars.slice(-n).join("");
  if (!/\s/u.test(chars[chars.length - n - 1]!)) {
    const space = kept.search(/\s/u);
    if (space >= 0 && space < kept.length - 1) kept = kept.slice(space + 1);
  }
  return `…${kept}`;
}

/** The first `n` characters of the text after a match, the same way. */
function head(s: string, n: number): string {
  const chars = Array.from(s);
  if (chars.length <= n) return s;
  let kept = chars.slice(0, n).join("");
  if (!/\s/u.test(chars[n]!)) {
    const space = kept.search(/\s\S*$/u);
    if (space > 0) kept = kept.slice(0, space);
  }
  return `${kept}…`;
}

/**
 * What each location of a blocked request matched: the text, up to 32 characters on each side
 * (within the same JSON string where it ends sooner), unescaped for reading, and where it is.
 * Best effort: the body is JSON, and a description's strings are escaped twice. Never throws, so
 * the diagnostics can't stop a preview; without them the block is shown as before.
 */
export function localMatches(body: string, locations: readonly AuditLocation[]): LocalMatch[] {
  try {
    return readMatches(body, locations);
  } catch {
    return [];
  }
}

function readMatches(body: string, locations: readonly AuditLocation[]): LocalMatch[] {
  const turns = readTurns(body);
  return locations.map((loc): LocalMatch => {
    // A description's or a reply's own strings are escaped twice in the body, the rest once. A match
    // in a number or other part of that JSON has no text around it worth showing.
    const within = loc.field?.within ?? "text";
    const depth = within === "string" ? 2 : 1;
    const unescape = (s: string) => (depth === 2 ? unescapeJson(unescapeJson(s)) : unescapeJson(s));
    const { before, after } = within === "json" ? { before: "", after: "" } : contextOf(body, loc.start, loc.end, depth);
    const place = describeField(loc.field, turns);
    // A match can end before accents or characters that can't be seen that follow it (the code and
    // digit projections stop at the last letter or digit): on screen they belong to the match.
    const rest = unescape(after);
    const marks = /^[\p{M}\p{Default_Ignorable_Code_Point}]+/u.exec(rest)?.[0] ?? "";
    const match: LocalMatch = {
      column: loc.column,
      variant: loc.variant,
      origin: loc.origin,
      text: unescape(body.slice(loc.start, loc.end)) + marks,
      before: tail(unescape(before), CONTEXT_CHARS),
      after: head(rest.slice(marks.length), CONTEXT_CHARS),
      where: place.where,
    };
    if (place.inHeader !== undefined) match.inHeader = place.inHeader;
    return match;
  });
}
