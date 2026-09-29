// UI wording in one place (spec §7.13). Sentence case, plain verbs, the same verb from button to
// confirmation. Messages name columns by letter only, never by value.
import { buildLabel } from "../config";
import type { AuditVariant, ColType, ProdMode, Treatment } from "../core/types";

export const MODE_LABEL: Record<ProdMode, string> = {
  structure_only: "Structure only",
  substituted: "Sample rows",
};

export const TREATMENT_LABEL: Record<Treatment, string> = {
  as_is: "As is",
  stand_in: "Stand-in",
  range: "Range",
  month: "Month",
  exclude: "Leave out",
};

/** What a treatment sends, shown under the picker. */
export const TREATMENT_HINT: Record<Treatment, string> = {
  as_is: "Sent as shown in Excel.",
  stand_in: "Sent as a label like PERSON_001.",
  range: "Numbers sent as a range like 80000-89999.",
  month: "Dates sent as a month like 2026-03.",
  exclude: "Not sent in sample rows.",
};

/** How a match relates to the private value, named on the Log screen. */
export const VARIANT_LABEL: Record<AuditVariant, string> = {
  exact: "whole value",
  normalized: "another spelling",
  digits: "digits",
  json: "escaped form",
  word: "a word of a value",
  canary: "bench canary",
};

/** The same, as a phrase about a private value ("a word of a private value"). */
const VARIANT_PHRASE: Record<AuditVariant, string> = {
  exact: "a private value",
  normalized: "another spelling of a private value",
  digits: "the digits of a private value",
  json: "a private value in escaped form",
  word: "a word of a private value",
  canary: "a bench canary",
};

export const TYPE_LABEL: Record<ColType, string> = {
  text: "text",
  number: "number",
  date: "date",
  boolean: "true/false",
  mixed: "mixed",
  empty: "empty",
};

export const COPY = {
  appName: "Nymform for Excel",
  footer: (version: string, commit: string) => `Nymform for Excel ${buildLabel(version, commit)} · SkySpan`,
  notInExcel: "Open Nymform from the Home tab in Excel.",
  crashed: "Something went wrong in the task pane. Close it and open it again from the Home tab.",

  // Setup
  keyLine: (host: string) => `Your key is sent only to ${host} and is forgotten when you close this pane.`,
  keyMissing: "Add your API key in Setup to send.",
  modelInvalid: "Enter a model ID such as openai/gpt-6-luna.",

  // Columns
  tooLarge: "Select fewer cells (limit 20,000).",
  tooLargeRange: "That range is too large. Choose fewer cells (limit 20,000).",
  noSelection: "Select a range in Excel, then choose Refresh from selection.",
  showCellsUnavailable: "This version of Excel can't select cells for Nymform.",
  showCellsFailed: "Couldn't select those cells in Excel.",
  pickPartial: "Choose which value you meant above, and the question goes out with its stand-in instead.",
  readFailed: "Couldn't read the selection. Select a range in Excel and try again.",
  rangeReadFailed: "Couldn't read that range. Check the sheet name and address.",
  rangeFormat: "Enter a range such as Regions!A1:B5.",
  rangeDuplicate: "That range is already included.",
  suggestionsNote: "Private columns are suggested from headers and values. Your choice decides.",
  hint: (reasons: string[]) => `Suggested: ${reasons.map(lowerFirst).join("; ")}. Your choice decides.`,
  treatmentsOff: "Treatments apply to sample rows. Turn on Include sample rows in Ask to choose them.",
  headerRowData: (row: number) => `Row ${row} looks like data, so it's treated as data. Tick the box if it is a header row.`,

  // Columns: which range is in use (spec §7.1). `from` is the selection as Excel names it ("E4", "4:4").
  using: (summary: string) => `Using ${summary}`,
  fromTable: (table: string, from: string) => `From the table ${table ? `${table} ` : ""}around ${from}.`,
  fromFilter: (from: string) => `From the filtered range around ${from}.`,
  fromRegion: (from: string) => `From the block of filled cells around ${from}.`,
  trimmed: (from: string, to: "columns" | "rows" | "cells") => `${from} trimmed to the ${to} in use.`,
  trimmedSheet: "The whole sheet trimmed to the cells in use.",
  useExact: "Use exactly the selected cells",
  keepsRange: "Nymform keeps using this range until you choose Refresh from selection.",
  noDataRows: "This range has no data rows. Select your data, or one cell in it, and choose Refresh from selection.",
  // The adapter's own sentences (RESOLVE_MESSAGE, NO_SUCH_SHEET_MESSAGE), matched to say what applies here.
  resolveSelection: "Select your data, or one cell in it, and choose Refresh from selection.",
  resolveLookup: "Select the lookup range, or one cell in it, and choose Add selected cells. A one-row range can be typed, like Regions!A1:B1.",
  noSuchSheet: "There's no sheet with that name. Check the sheet name in the address.",
  exactSheetGone: "Couldn't find the sheet those cells were on. Select your data, or one cell in it, and choose Refresh from selection.",

  // Ask
  noQuestion: "Type a question first.",
  alreadySending: "A request is already on its way. Wait for its reply.",
  alreadyInserting: "A formula is already being inserted. Wait for it to finish.",
  stalePreview: "The columns or the conversation changed after this preview. Preview again.",
  dataChanged: "The data changed since Nymform read it. Select your data, or one cell in it, and choose Refresh from selection.",
  dataChangedSend:
    "The data changed after this preview (cells, headers, tables or defined names), so nothing was sent. Select your data, choose Refresh from selection, and preview again.",
  dataChangedInsert:
    "The data changed after this answer (cells, headers, tables or defined names), so the formula may not fit it any more. Nothing was written. Select your data, choose Refresh from selection, and ask again.",
  recheckFailedSend: "Couldn't check that the data is unchanged, so nothing was sent. Try again.",
  recheckFailedInsert: "Couldn't check that the data is unchanged, so nothing was written. Try again.",
  correctionReady:
    "The model's reply couldn't be read. Nymform can ask once more: the request below repeats your question with the model's reply and a note asking for the right format. It is checked like any request, and sent only if you choose Send.",
  correctionHeading: "Asking once more",

  // What gets sent
  formatted: (host: string) => `Formatted for reading. The check runs on the exact request sent to ${host}.`,
  auditOk: "Checked: no private values found",
  auditBlockedColumns: (columns: string[]) =>
    columns.length === 1
      ? `Blocked: a value from column ${columns[0]} is in the request.`
      : `Blocked: values from columns ${listJoin(columns)} are in the request.`,
  auditBlockedAlias: (columns: string[]) =>
    columns.length === 1
      ? `Blocked: the original header of column ${columns[0]} (renamed under Header sent) is in the request.`
      : `Blocked: the original headers of columns ${listJoin(columns)} (renamed under Header sent) are in the request.`,
  auditBlockedCanary: "Blocked: a bench canary is in the request.",
  auditFound: (places: string[]) => `Found in ${listJoin(places)}.`,
  auditBlockedError: (error: string) => `Blocked: ${lowerFirst(error)}`,
  auditAction: "Change the question or the column settings, then preview again.",
  notSentBody: "Not kept: this request was blocked, and its text holds a private value.",

  // Log: what a blocked request matched (pane only)
  logMatchAlias: (column: string) =>
    `It matches the original header of column ${column}, which you renamed under Header sent. A renamed header is checked like a private value.`,
  logMatchHeader: (header: string, letter: string, text: string, variant: AuditVariant, column: string) =>
    `The header of column ${header} contains “${text}”, ${VARIANT_PHRASE[variant]} in column ${column}. Rename ${header} under Header sent (for example “Column ${letter}”), then send again.`,
  logMatchValue: (variant: AuditVariant, column: string) =>
    variant === "canary" && column === "*"
      ? "It matches a bench canary."
      : `It matches a value (${VARIANT_LABEL[variant]}) from private column ${column}.`,
  logMatchNote: "Shown only here. Not in Export log or the evaluation report.",

  // Result
  unreadable: "The model's reply couldn't be read. Try rephrasing.",
  retryBlocked: "The model's reply couldn't be read, and the retry was blocked by the check. Try rephrasing.",
  gateOk: "Formula check passed: allowed functions and your ranges only.",
  gateBlocked: "Blocked: this formula can't be inserted.",
  unknownTokens: (tokens: string[]) =>
    `The reply mentions ${tokens.length === 1 ? "a stand-in" : "stand-ins"} this session didn't create: ${tokens.join(", ")}. ${
      tokens.length === 1 ? "It was" : "They were"
    } left as written.`,
  unresolvedTokens: (tokens: string[]) =>
    `Blocked: the formula uses ${tokens.length === 1 ? "a stand-in" : "stand-ins"} this session didn't create (${tokens.join(", ")}). ` +
    `Nymform can't put a real value back for ${tokens.length === 1 ? "it" : "them"}, so the formula would compare against the stand-in text. Ask again.`,
  insertUnresolved: "This formula uses a stand-in this session didn't create, so it can't be inserted. Ask again.",
  restoredNumber: (token: string, column: string) =>
    `${token} is put back as text, but column ${column} holds it as a number or date. An = comparison, XLOOKUP or MATCH may not find it. Check the result after inserting.`,
  restoredWildcard: (token: string) =>
    `The value put back for ${token} contains *, ? or ~. SUMIFS, COUNTIFS, MATCH and XLOOKUP's wildcard mode read these as wildcards, so it can match other values too. Check the result after inserting.`,
  wholeColumns: (refs: string[]) =>
    `Uses whole ${refs.length === 1 ? "column" : "columns"} (${refs.join(", ")}): in Excel, the formula also reads rows of ${
      refs.length === 1 ? "that column" : "those columns"
    } outside your selection. That happens in Excel only; nothing more is sent.`,
  inserted: (cell: string) => `Inserted in ${cell}.`,
  insertedError: (cell: string) =>
    `Inserted in ${cell}, but Excel shows an error in the result (such as #N/A or #SPILL!). Check the cells before relying on them.`,
  insertedManual: (cell: string) =>
    `Inserted in ${cell}. Excel is set to calculate manually, so the result isn't calculated or checked yet. Choose Formulas → Calculate Now to see it.`,
  insertedUnverified: (cell: string) =>
    `Inserted in ${cell}, but Nymform couldn't read the result back to check it. Look at the cells before relying on them.`,
  insertedFillFailed: (cell: string, span: string) =>
    `Inserted in ${cell} only: Excel couldn't fill it down to ${span}. Check ${cell}, then fill it down in Excel or insert again.`,
  overwriteAsk: (range: string) => `${range} already has values. Insert anyway to replace them; Excel can't undo this.`,
  overwriteUnknown: (range: string) =>
    `Couldn't check whether ${range} is empty. Insert anyway to replace anything there; Excel can't undo this.`,
  overwriteButton: "Replace and insert",
  insertBlocked: "This formula didn't pass the check, so it can't be inserted.",
  staleResult: "The selection changed after this reply. Ask the question again.",
  insertFailed: "Couldn't insert the formula. Check that the sheet isn't protected, then try again.",
  placementInvalid: "Enter a cell on this sheet, for example G2.",
  placementInside: "Choose a cell outside the ranges you selected, so no data is overwritten.",
  fillLabel: (lastRow: number, span: string | null) => `Fill down to row ${lastRow}${span ? ` (${span})` : ""}`,
  fillStart: (cell: string) => `Fill down starts at the first data row, ${cell}. Change the cell or turn off Fill down.`,
  fillBlocked: "Blocked: when filled down, this formula refers outside your ranges.",
  copied: "Copied.",
  copyFailed: "Couldn't copy. Select the text and copy it by hand.",
  unexpected: "Something went wrong. Try again.",
} as const;

export function lowerFirst(s: string): string {
  if (s.length < 2) return s.toLowerCase();
  // Keep acronyms and names that start with two capitals ("ID", "URL").
  if (/^\p{Lu}\p{Lu}/u.test(s)) return s;
  return s.charAt(0).toLowerCase() + s.slice(1);
}

/** "A", "A and C", "A, C and D". */
export function listJoin(items: string[]): string {
  if (items.length <= 1) return items.join("");
  return `${items.slice(0, -1).join(", ")} and ${items[items.length - 1]}`;
}
