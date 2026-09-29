// The one builder of outbound request bodies (spec §7.6, invariant 1).
// Every body is built here and then checked by auditor.ts on the exact string before provider.ts sends it.
import { IS_OPENROUTER } from "../config";
import { buildMessages, CORRECTION_PROMPT, type ChatMessage } from "./prompt";
import { rowCount } from "./schema";
import { effectiveTreatment, transformRows, type StandInMap } from "./transform";
import type {
  CellValue,
  ColType,
  ColumnInfo,
  ColumnPolicy,
  HistoryMessage,
  Outbound,
  ProdMode,
  RangeSpec,
  SheetContext,
  Treatment,
} from "./types";

export interface PayloadInput {
  model: string;
  /** [0] is the selection; the rest are context ranges, in the order the user added them. */
  specs: readonly RangeSpec[];
  sheetContext: SheetContext;
  /** The user's question, already passed through substituteUserText. */
  question: string;
  history: readonly HistoryMessage[];
  /** Session stand-in map; substituted mode allocates tokens through it. */
  map: StandInMap;
  /** Substituted mode row cap: default 50, maximum 200. */
  rowCap?: number;
}

export interface BuiltPayload {
  outbound: Outbound;
  /** The exact user turn inside the body, for history. */
  userContent: string;
  messages: ChatMessage[];
  rowsSent?: number;
  rowsTotal?: number;
}

/** Version of the user-turn format ("nymform" key). */
export const PROTOCOL_VERSION = "0.1";

const REASONING_PREFIXES = ["openai/gpt-6", "openai/gpt-5", "openai/o1", "openai/o3", "openai/o4"];

/** Reasoning models take no temperature and need room for reasoning before the reply. */
export function isReasoningModel(model: string): boolean {
  const id = model.toLowerCase();
  return REASONING_PREFIXES.some((p) => id.startsWith(p));
}

/** Extra request parameters for a model (temperature, reasoning, token cap, response format). */
export function requestParams(model: string): Record<string, unknown> {
  if (isReasoningModel(model)) {
    return {
      max_tokens: 4000,
      reasoning: { effort: "low", exclude: true },
      response_format: { type: "json_object" },
    };
  }
  return { temperature: 0, max_tokens: 800, response_format: { type: "json_object" } };
}

export function buildStructureOnly(input: PayloadInput): BuiltPayload {
  return buildDescription(input, "structure_only");
}

export function buildSubstituted(input: PayloadInput): BuiltPayload {
  return buildDescription(input, "substituted");
}

/** The one retry after an unreadable reply (spec §7.9): same turn, the bad reply, then CORRECTION_PROMPT. */
export function buildCorrection(input: {
  model: string;
  mode: ProdMode;
  history: readonly HistoryMessage[];
  userContent: string;
  invalidReply: string;
}): BuiltPayload {
  const messages = buildMessages(input.history, input.userContent);
  // An empty reply (for example, reasoning used up the token cap) is left out: providers
  // reject empty assistant turns, and there is nothing in it to correct.
  if (typeof input.invalidReply === "string" && input.invalidReply.trim() !== "") {
    messages.push({ role: "assistant", content: input.invalidReply });
  }
  messages.push({ role: "user", content: CORRECTION_PROMPT });
  return { outbound: outbound(input.model, messages, input.mode), userContent: input.userContent, messages };
}

// ---------------------------------------------------------------------------------------------
// The user turn (key order is part of the format)

interface WireStats {
  blank: number;
  distinct: number;
  avg_words?: number;
  max_len?: number;
}

interface WireColumn {
  letter: string;
  header: string;
  type: ColType;
  private: boolean;
  treatment: Treatment;
  note: string | null;
  stats: WireStats;
}

interface WireRange {
  address: string;
  header_row: number | null;
  first_data_row: number;
  last_data_row: number;
}

interface WireRows {
  rows: CellValue[][];
  rows_sent: number;
  rows_total: number;
}

interface WireBlock extends Partial<WireRows> {
  sheet: string;
  table: string | null;
  range: WireRange;
  columns: WireColumn[];
}

function buildDescription(input: PayloadInput, mode: ProdMode): BuiltPayload {
  const [selection, ...contexts] = input.specs;
  if (!selection) throw new Error("There is no selection to describe.");
  const withRows = mode === "substituted";

  const main = describeRange(selection, withRows ? input : null);
  const contextRanges = contexts.map((spec) => describeRange(spec, withRows ? input : null));

  const turn = {
    nymform: PROTOCOL_VERSION,
    mode,
    sheet: main.block.sheet,
    table: main.block.table,
    range: main.block.range,
    columns: main.block.columns,
    ...(main.rows ?? {}),
    context_ranges: contextRanges.map((c) => ({ ...c.block, ...(c.rows ?? {}) })),
    allowed_ranges: [...input.sheetContext.allowedRanges],
    task: input.question,
  };
  const userContent = JSON.stringify(turn);
  const messages = buildMessages(input.history, userContent);
  const built: BuiltPayload = { outbound: outbound(input.model, messages, mode), userContent, messages };
  if (main.rows) {
    built.rowsSent = main.rows.rows_sent;
    built.rowsTotal = main.rows.rows_total;
  }
  return built;
}

/** One range, described the same way for the selection and for context ranges. */
function describeRange(spec: RangeSpec, rowsFrom: PayloadInput | null): { block: WireBlock; rows?: WireRows } {
  const d = spec.data;
  const top = d.rowIndex + 1;
  const block: WireBlock = {
    sheet: d.sheet,
    table: d.table?.name ? d.table.name : null,
    range: {
      address: d.address,
      header_row: spec.hasHeaders ? top : null,
      first_data_row: spec.hasHeaders ? top + 1 : top,
      last_data_row: top + rowCount(d) - 1,
    },
    columns: spec.columns.map((info) => describeColumn(info, policyFor(spec, info))),
  };
  if (!rowsFrom) return { block };
  const t = transformRows(spec, rowsFrom.map, rowsFrom.rowCap);
  return { block, rows: { rows: t.rows, rows_sent: t.rowsSent, rows_total: t.rowsTotal } };
}

/** A column without a policy counts as private (fail closed), matching transform.ts. */
function policyFor(spec: RangeSpec, info: ColumnInfo): ColumnPolicy {
  return (
    spec.policies.find((p) => p.letter === info.letter) ?? { letter: info.letter, private: true, treatment: "stand_in" }
  );
}

function describeColumn(info: ColumnInfo, policy: ColumnPolicy): WireColumn {
  const stats: WireStats = { blank: info.stats.blank, distinct: info.stats.distinct };
  if (info.type === "text") {
    if (typeof info.stats.avgWords === "number") stats.avg_words = info.stats.avgWords;
    if (typeof info.stats.maxLen === "number") stats.max_len = info.stats.maxLen;
  }
  const note = typeof policy.note === "string" ? policy.note.trim() : "";
  return {
    letter: info.letter,
    header: aliasOf(policy, info) ?? info.header,
    type: info.type,
    private: policy.private,
    treatment: effectiveTreatment(policy),
    note: note === "" ? null : note,
    stats,
  };
}

/** The alias sent instead of the header, if any. The original header is then never sent. */
function aliasOf(policy: ColumnPolicy, info: ColumnInfo): string | undefined {
  for (const a of [policy.alias, info.alias]) {
    if (typeof a === "string" && a.trim() !== "") return a.trim();
  }
  return undefined;
}

// ---------------------------------------------------------------------------------------------
// The body

function outbound(model: string, messages: ChatMessage[], mode: ProdMode): Outbound {
  const body = JSON.stringify({
    model,
    messages,
    ...requestParams(model),
    ...(IS_OPENROUTER ? { provider: { zdr: true } } : {}),
  });
  return { body, bytes: new TextEncoder().encode(body).byteLength, mode, createdAt: new Date().toISOString() };
}
