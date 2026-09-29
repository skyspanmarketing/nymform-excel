// Raw baseline sender for the benchmark (spec §11, invariant 8). It sends every value in the range,
// with no stand-ins and no auditor, because that is the baseline the other conditions are compared
// with. It runs only on workbooks carrying the bench marker, and only in bench builds: src/ may load
// it solely through the bench panel's `if (BENCH)` dynamic import. This file and
// src/core/provider.ts are the only callers of fetch.
import { qualify } from "../src/core/a1";
import { requestParams } from "../src/core/payload";
import { buildMessages } from "../src/core/prompt";
import { inferSchema, isDateFormat, rowCount } from "../src/core/schema";
import type { CellValue, ColType, RangeData } from "../src/core/types";

export const BENCH_MARKER = "NYMFORM_BENCH_V1";
/** The request mode of raw baseline runs. Release bundles must never contain this string (T-P4). */
export const RAW_MODE = "raw_bench" as const;
export const NOT_A_BENCH_WORKBOOK =
  "Not a bench workbook: the raw baseline runs only on workbooks with the bench marker.";
export const RAW_TIMEOUT_MS = 60_000;

export interface RawBodyInput {
  /** A1 of the hidden _nymform_bench sheet, as read from the open workbook. */
  marker: string | null | undefined;
  data: RangeData;
  question: string;
  model: string;
  /** Endpoint base URL, e.g. https://openrouter.ai/api/v1. */
  endpoint: string;
  /** Default true: the first row of `data` is the header row. */
  hasHeaders?: boolean;
  /** Default: every data row. */
  rowCap?: number;
}

export interface RawSendInput extends RawBodyInput {
  key: string;
  fetchImpl?: typeof fetch;
  signal?: AbortSignal;
  timeoutMs?: number;
  now?: () => number;
}

export interface RawBody {
  body: string;
  bytes: number;
  mode: typeof RAW_MODE;
  userContent: string;
  rowsSent: number;
  rowsTotal: number;
}

export interface RawResult {
  raw: string;
  status: number;
  latencyMs: number;
  /** The exact string sent. */
  body: string;
  bytes: number;
  mode: typeof RAW_MODE;
  tokens?: { prompt: number; completion: number };
  providerName?: string;
}

/** Throws unless the workbook carries the bench marker. */
export function assertBenchWorkbook(marker: string | null | undefined): void {
  if (marker !== BENCH_MARKER) throw new Error(NOT_A_BENCH_WORKBOOK);
}

function isOpenRouter(endpoint: string): boolean {
  return new URL(endpoint).host === "openrouter.ai";
}

function rawCell(value: CellValue | undefined, text: string | undefined, numberFormat: string | undefined, type: ColType): CellValue {
  if (value === null || value === undefined || value === "") return text ? text : null;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") {
    // Dates go as display text, like non-private dates on the production path.
    if (type === "date" || isDateFormat(numberFormat ?? "")) return text || String(value);
    return value;
  }
  return text ? text : value;
}

/**
 * The raw baseline body: the production request shape (same system prompt, request parameters and
 * provider field) with mode "raw_bench" and every value of the range in `rows`.
 */
export function buildRawBody(input: RawBodyInput): RawBody {
  assertBenchWorkbook(input.marker);
  const { data } = input;
  const hasHeaders = input.hasHeaders ?? true;
  const first = hasHeaders ? 1 : 0;
  const rowsTotal = Math.max(0, rowCount(data) - first);
  const cap = input.rowCap === undefined ? rowsTotal : Math.max(0, Math.floor(input.rowCap));
  const rowsSent = Math.min(rowsTotal, cap);
  const columns = inferSchema(data, hasHeaders);
  const top = data.rowIndex + 1;

  const rows: CellValue[][] = [];
  for (let r = first; r < first + rowsSent; r++) {
    rows.push(
      columns.map((col, c) =>
        rawCell(data.values[r]?.[c], data.text[r]?.[c], data.numberFormat[r]?.[c], col.type),
      ),
    );
  }

  const user = {
    nymform: "0.1",
    mode: RAW_MODE,
    sheet: data.sheet,
    table: data.table?.name ?? null,
    range: {
      address: data.address,
      header_row: hasHeaders ? top : null,
      first_data_row: top + first,
      last_data_row: top + rowCount(data) - 1,
    },
    columns: columns.map((c) => ({
      letter: c.letter,
      header: c.header,
      type: c.type,
      private: false,
      treatment: "as_is",
      note: null,
      stats: {
        blank: c.stats.blank,
        distinct: c.stats.distinct,
        ...(c.stats.avgWords !== undefined ? { avg_words: c.stats.avgWords } : {}),
        ...(c.stats.maxLen !== undefined ? { max_len: c.stats.maxLen } : {}),
      },
    })),
    rows,
    rows_sent: rowsSent,
    rows_total: rowsTotal,
    context_ranges: [],
    allowed_ranges: [qualify(data.sheet, data.address)],
    task: input.question,
  };
  const userContent = JSON.stringify(user);
  const body = JSON.stringify({
    model: input.model,
    messages: buildMessages([], userContent),
    ...requestParams(input.model),
    ...(isOpenRouter(input.endpoint) ? { provider: { zdr: true } } : {}),
  });
  return { body, bytes: new TextEncoder().encode(body).length, mode: RAW_MODE, userContent, rowsSent, rowsTotal };
}

function usageOf(raw: string): { prompt: number; completion: number } | undefined {
  try {
    const u = (JSON.parse(raw) as { usage?: { prompt_tokens?: unknown; completion_tokens?: unknown } }).usage;
    if (u && typeof u.prompt_tokens === "number" && typeof u.completion_tokens === "number") {
      return { prompt: u.prompt_tokens, completion: u.completion_tokens };
    }
  } catch {
    // Not JSON: no usage.
  }
  return undefined;
}

/**
 * Sends the raw baseline request. Throws before building anything when the marker is missing.
 * Returns the response whatever its status; the bench runner decides what a failure is.
 */
export async function sendRaw(input: RawSendInput): Promise<RawResult> {
  assertBenchWorkbook(input.marker);
  const built = buildRawBody(input);
  const now = input.now ?? (() => Date.now());
  const endpoint = input.endpoint.replace(/\/+$/, "");
  const host = new URL(endpoint).host;
  const controller = new AbortController();
  const onAbort = () => controller.abort();
  if (input.signal?.aborted) controller.abort();
  input.signal?.addEventListener("abort", onAbort, { once: true });
  let timedOut = false;
  const timer = setTimeout(() => {
    timedOut = true;
    controller.abort();
  }, input.timeoutMs ?? RAW_TIMEOUT_MS);

  const started = now();
  try {
    const send = input.fetchImpl ?? fetch;
    let response: Response;
    try {
      response = await send(`${endpoint}/chat/completions`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${input.key}`,
          "Content-Type": "application/json",
          "X-Title": "Nymform for Excel",
        },
        body: built.body,
        signal: controller.signal,
      });
    } catch {
      if (timedOut) throw new Error(`${host} didn't answer within ${Math.round((input.timeoutMs ?? RAW_TIMEOUT_MS) / 1000)} seconds.`);
      if (controller.signal.aborted) throw new Error("The raw request was stopped.");
      throw new Error(`Couldn't reach ${host}.`);
    }
    const raw = await response.text();
    const result: RawResult = {
      raw,
      status: response.status,
      latencyMs: Math.max(0, now() - started),
      body: built.body,
      bytes: built.bytes,
      mode: RAW_MODE,
    };
    const tokens = usageOf(raw);
    if (tokens) result.tokens = tokens;
    const providerName = response.headers?.get?.("x-provider-name");
    if (providerName) result.providerName = providerName;
    return result;
  } finally {
    clearTimeout(timer);
    input.signal?.removeEventListener("abort", onAbort);
  }
}
