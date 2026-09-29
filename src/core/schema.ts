// Schema inference (spec §7.2). Describes columns by header, type and counts only: never cell values.
import { columnIndex, columnLetter, qualify } from "./a1";
import type { CellValue, ColType, ColumnInfo, RangeData, SheetContext, TableInfo } from "./types";

/** The kind of one cell, before the 80% rule turns a column of them into a ColType. */
export type CellKind = "empty" | "text" | "number" | "date" | "boolean" | "error" | "other";

/** True when an Excel number format has d, m or y tokens outside quotes and brackets. */
export function isDateFormat(numberFormat: string): boolean {
  if (typeof numberFormat !== "string" || numberFormat === "") return false;
  const stripped = numberFormat
    .replace(/"[^"]*"?/g, " ") // quoted literals
    .replace(/\[[^\]]*\]?/g, " ") // colors, conditions, locale codes, elapsed time
    .replace(/\\./g, " ") // escaped characters
    .replace(/[_*]./g, " ") // padding and fill characters
    .replace(/am\/pm|a\/p/gi, " ");
  return /[dmy]/i.test(stripped);
}

/** True when a cell holds nothing. Formulas that return "" count as empty too. */
export function isEmptyCell(value: CellValue | undefined, text: string | undefined, valueType?: string): boolean {
  if (valueType === "Empty") return true;
  return (value === null || value === undefined || value === "") && (text === undefined || text === "");
}

/** Classifies one cell from its Excel value type (or, when that is missing, its JavaScript type). */
export function cellKind(
  value: CellValue | undefined,
  text: string | undefined,
  valueType: string | undefined,
  numberFormat: string | undefined,
): CellKind {
  if (isEmptyCell(value, text, valueType)) return "empty";
  switch (valueType) {
    case "Double":
    case "Integer":
      return isDateFormat(numberFormat ?? "") ? "date" : "number";
    case "String":
      return "text";
    case "Boolean":
      return "boolean";
    case "Error":
      return "error";
    case undefined:
    case "":
      break;
    default:
      return "other";
  }
  if (typeof value === "number") return isDateFormat(numberFormat ?? "") ? "date" : "number";
  if (typeof value === "boolean") return "boolean";
  return "text";
}

/** Number of rows in a range, from whichever array is longest. */
export function rowCount(data: RangeData): number {
  return Math.max(data.values.length, data.text.length);
}

/** Number of columns in a range. */
export function columnCount(data: RangeData): number {
  let n = 0;
  for (const row of data.values) n = Math.max(n, row.length);
  for (const row of data.text) n = Math.max(n, row.length);
  return n;
}

/** 0-based position of a column inside the range, from its letter; `fallback` when the letter is off. */
export function columnOffset(data: RangeData, letter: string, fallback: number): number {
  const abs = columnIndex(letter);
  const off = abs - data.columnIndex;
  return abs >= 0 && off >= 0 ? off : fallback;
}

/** Words in a display string, split on whitespace. */
export function wordCount(text: string): number {
  const t = text.trim();
  return t === "" ? 0 : t.split(/\s+/u).length;
}

const TYPE_SHARE = 0.8;

/** Describes each column: header, type, stats. Never contains cell values. */
export function inferSchema(data: RangeData, hasHeaders = true): ColumnInfo[] {
  const rows = rowCount(data);
  const cols = columnCount(data);
  const first = hasHeaders ? 1 : 0;
  const out: ColumnInfo[] = [];

  for (let c = 0; c < cols; c++) {
    const letter = columnLetter(data.columnIndex + c);
    const rawHeader = hasHeaders ? (data.text[0]?.[c] ?? stringOf(data.values[0]?.[c])) : "";
    const header = rawHeader.trim() === "" ? `Column ${letter}` : rawHeader.trim();

    const kinds = new Map<CellKind, number>();
    const distinct = new Set<string>();
    let blank = 0;
    let nonEmpty = 0;
    let words = 0;
    let maxLen = 0;

    for (let r = first; r < rows; r++) {
      const value = data.values[r]?.[c];
      const text = data.text[r]?.[c];
      const kind = cellKind(value, text, data.valueTypes[r]?.[c], data.numberFormat[r]?.[c]);
      if (kind === "empty") {
        blank++;
        continue;
      }
      nonEmpty++;
      kinds.set(kind, (kinds.get(kind) ?? 0) + 1);
      distinct.add(distinctKey(value, text));
      const shown = text ?? stringOf(value);
      words += wordCount(shown);
      maxLen = Math.max(maxLen, shown.length);
    }

    const type = columnType(kinds, nonEmpty);
    const stats: ColumnInfo["stats"] = { blank, distinct: distinct.size };
    if (type === "text") {
      stats.avgWords = nonEmpty === 0 ? 0 : Math.round((words / nonEmpty) * 10) / 10;
      stats.maxLen = maxLen;
    }
    out.push({ letter, header, type, stats });
  }
  return out;
}

function columnType(kinds: Map<CellKind, number>, nonEmpty: number): ColType {
  if (nonEmpty === 0) return "empty";
  for (const kind of ["text", "number", "date", "boolean"] as const) {
    if ((kinds.get(kind) ?? 0) >= TYPE_SHARE * nonEmpty) return kind;
  }
  return "mixed";
}

function distinctKey(value: CellValue | undefined, text: string | undefined): string {
  if (typeof value === "number") return `n:${value}`;
  if (typeof value === "boolean") return `b:${value}`;
  if (typeof value === "string" && value !== "") return `s:${value}`;
  return `t:${text ?? ""}`;
}

function stringOf(value: CellValue | undefined): string {
  return value === null || value === undefined ? "" : String(value);
}

/** Row numbers, sheets, tables and allowed ranges for the selection plus context ranges. */
export function buildSheetContext(
  main: RangeData,
  hasHeaders: boolean,
  contexts: readonly RangeData[],
  workbook: { sheets: string[]; tables: TableInfo[]; names?: string[] },
): SheetContext {
  const top = main.rowIndex + 1;
  const rows = rowCount(main);
  const sheets = uniq([...workbook.sheets, main.sheet, ...contexts.map((c) => c.sheet)]);
  const tables = [...workbook.tables];
  const tableNames = uniq([...tables.map((t) => t.name), ...(main.table?.name ? [main.table.name] : [])]);
  const allowedRanges = uniq([main, ...contexts].map((r) => qualify(r.sheet, r.address)));

  const ctx: SheetContext = {
    sheet: main.sheet,
    address: main.address,
    // 0 means the range has no header row (sent as header_row: null).
    headerRow: hasHeaders ? top : 0,
    firstDataRow: hasHeaders ? top + 1 : top,
    lastDataRow: top + rows - 1,
    sheets,
    tableNames,
    allowedRanges,
    tables,
  };
  if (main.table?.name) ctx.table = main.table.name;
  if (workbook.names) ctx.definedNames = uniq(workbook.names);
  return ctx;
}

function uniq(items: string[]): string[] {
  return [...new Set(items)];
}

// ---------------------------------------------------------------------------------------------
// A header row that is really data

// The same shapes as suggest.ts isEmailText, isPhoneText, isSsnText and isIdShaped. They are
// repeated here because suggest.ts imports this file; tests/schema.test.ts keeps the two in step.
const EMAIL_EXACT_RE = /^[^\s@<>()"',;:]+@[^\s@<>()"',;:]+\.[A-Za-z]{2,}$/u;
const PHONE_SHAPE_RE = /^[+(]?[\d\s().\-/+]+(?:\s*(?:x|ext\.?)\s*\d+)?$/iu;
const SSN_RE = /^(?:\d{3}-\d{2}-\d{4}|\d{3} \d{2} \d{4}|\d{9})$/u;
const ID_SHAPE_RE = /^[\p{L}\p{N}][\p{L}\p{N}\-_./#]*$/u;

// Numbers as Excel shows them: "$83,500", "€1.234,50", "83 500 ₽", "(1,200.00)", "12.5%", "83,500".
const GROUPED = String.raw`\d{1,3}(?:[,.'\u00a0\u202f ]\d{3})+(?:[.,]\d+)?`;
const DECIMAL = String.raw`(?:${GROUPED}|\d+(?:[.,]\d+)?)`;
const AMOUNT_RE = new RegExp(
  String.raw`^[-+(]?\s*(?:\p{Sc}\s*[-+]?${DECIMAL}|[-+]?${DECIMAL}\s*\p{Sc}|[-+]?${DECIMAL}\s?%|[-+]?${GROUPED})\s*\)?$`,
  "u",
);
/** A plain number written as text, in any digits: "2024", "02134", "-5", "12.5", "3,75", "２０２４". */
const PLAIN_NUMBER_RE = /^[-+]?\p{Nd}+(?:[.,]\p{Nd}+)?$/u;
// Dates and times: "15 Mar 2024", "March 15, 2024", "Wed, 15-Mar-24", "2024-03-15", "3/15/2026", "10:30 AM".
const MONTH =
  String.raw`(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|` +
  String.raw`sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\.?`;
const DAY = String.raw`\d{1,2}(?:st|nd|rd|th)?`;
const WEEKDAY = String.raw`(?:(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?,?\s+)?`;
const TIME = String.raw`\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:\s?[ap]\.?m\.?)?`;
const SEP = String.raw`[\s\-/]+`;
const DAY_MONTH = String.raw`${DAY}${SEP}${MONTH},?${SEP}\d{2,4}`;
const MONTH_DAY = String.raw`${MONTH}${SEP}${DAY},?${SEP}\d{2,4}`;
const NUMERIC_DATE = String.raw`\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.](?:\d{4}|\d{2})`;
// Other languages: a day, a month name and a year ("15. März 1985", "15 mars 1985"), and the
// Chinese, Japanese and Korean forms ("1985年3月15日", "1985년 3월 15일").
const LOCAL_DATE =
  String.raw`\d{1,2}\.?[\s\-/.]+\p{L}{3,}\.?,?[\s\-/.]+\d{4}|` +
  String.raw`\d{4}\s*[年년]\s*\d{1,2}\s*[月월](?:\s*\d{1,2}\s*[日일])?`;
const DATE_RE = new RegExp(
  String.raw`^(?:${WEEKDAY}(?:${DAY_MONTH}|${MONTH_DAY})|${NUMERIC_DATE}|${LOCAL_DATE})(?:(?:,?\s+|T)${TIME}Z?)?$`,
  "iu",
);
const TIME_RE = new RegExp(`^${TIME}$`, "iu");
/** Two or more letters in a row: "FY2024", "Address1" and "Test1" read as words; "B-201" and "E10234" don't. */
const LETTER_RUN_RE = /\p{L}{2,}/u;

/** How many data cells below a header cell are compared with it. */
const BELOW_SAMPLE = 200;

/** A data cell below a header cell. */
export interface BelowCell {
  text: string;
  value: CellValue;
  valueType: string;
}

/**
 * How a cell reads: "identifier" (an email address, phone number, SSN, IBAN, UK National Insurance
 * number or postcode), "address" (a street address with a street-type word, see isStreetAddress),
 * "value" (a number, a boolean, an amount, a percentage, a date or a time, as a value or as text),
 * "id" (ID-shaped text such as EMP-00123 or FY2024), "text" or "empty".
 */
type CellShape = "identifier" | "address" | "value" | "id" | "text" | "empty";

// IBANs: a country code, 2 check digits and up to 30 letters and digits, 15 to 34 in all, written
// together or in groups of 4 ("DE89 3704 0044 0532 0130 00").
const IBAN_RE = /^[A-Z]{2}\d{2}[A-Z\d]{11,30}$/u;
const IBAN_GROUPED_RE = /^[A-Z]{2}\d{2}(?: [A-Z\d]{4})+(?: [A-Z\d]{1,3})?$/u;

/** An IBAN, written together or in groups. A grouped one needs 15 characters too: "AB12 CD34" is a code. */
function isIban(t: string): boolean {
  if (IBAN_RE.test(t)) return true;
  return IBAN_GROUPED_RE.test(t) && t.replace(/ /gu, "").length >= 15;
}

/** A UK National Insurance number: "QQ 12 34 56 C", "QQ123456C". */
const NI_NUMBER_RE = /^[A-Z]{2} ?\d{2} ?\d{2} ?\d{2} ?[A-D]$/u;
/**
 * Postcodes that count on their own: UK ("SW1A 1AA", "M1 1AE"), Canadian ("K1A 0B1"), US ZIP+4
 * ("94107-1234"), Dutch ("1234 AB") and Swedish, Czech or Greek ("123 45") ones. Irish ones are
 * below.
 */
const POSTCODE_RES = [
  /^[A-Z]{1,2}\d[A-Z\d]? ?\d[A-Z]{2}$/u,
  /^[A-Z]\d[A-Z] ?\d[A-Z]\d$/u,
  /^\d{5}[- ]\d{4}$/u,
  /^\d{4} [A-Z]{2}$/u,
  /^\d{3} \d{2}$/u,
];
/** An Irish postcode: "D02 X285", "T12 X70A". */
const IRISH_POSTCODE_RE = /^[A-Z]\d[\dW] ([A-Z\d]{4})$/u;
const YEAR_RE = /^(?:19|20)\d{2}$/u;

/**
 * A postcode in one of the country shapes above. An Irish shape whose second part is a year
 * ("P12 2024", "W52 2024") reads as a period and a year, so it is a header, as "Q1 2024" is.
 */
function isPostcode(t: string): boolean {
  if (POSTCODE_RES.some((re) => re.test(t))) return true;
  const irish = IRISH_POSTCODE_RE.exec(t);
  return irish !== null && !YEAR_RE.test(irish[1]!);
}

/**
 * Other codes of capital letters and digits with a space ("A1 B2", "AB12 CD34"): two parts of 2 to
 * 4, each with a digit and one with a letter. Such a code in row 1 counts as data only when most
 * of the column below is postcodes or codes like it (see matchesColumn): on its own it is as
 * likely a header. A part that names a period ("Q1", "H2", "FY24", "W12", or a year) makes it a
 * header such as "Q1 2024" or "H1 FY25" whatever is below.
 */
const CODE_PARTS_RE = /^([A-Z\d]{2,4}) ([A-Z\d]{2,4})$/u;
const PERIOD_PART_RE = /^(?:[QHTSPWM]\d{1,2}|FY\d{2}(?:\d{2})?|CY\d{2}(?:\d{2})?|(?:19|20)\d{2})$/u;

function isCodeText(t: string): boolean {
  const m = CODE_PARTS_RE.exec(t);
  if (!m) return false;
  const parts = [m[1]!, m[2]!];
  return (
    parts.every((p) => /\d/u.test(p) && !PERIOD_PART_RE.test(p)) && parts.some((p) => /[A-Z]/u.test(p))
  );
}

// Street addresses. A street address counts as data on its own only when it has a street-type
// word, because "10 Largest Orders" and "5 Star Reviews" have the same shape as "12 Baker". Text
// that starts with a house number but has no street-type word counts only over a column of
// addresses (see matchesColumn).

/** Street-type words that follow the street's name: "12 Baker Street", "3 High St", "742 Evergreen Terrace". */
const STREET_TYPES = new Set([
  "street", "st", "avenue", "ave", "av", "road", "rd", "boulevard", "blvd", "lane", "ln", "drive", "dr",
  "way", "court", "ct", "place", "pl", "terrace", "ter", "terr", "parkway", "pkwy", "highway", "hwy",
  "square", "sq", "close", "crescent", "cres", "circle", "cir", "trail", "trl", "mews", "grove", "gardens",
  "gdns", "walk", "alley", "aly", "pike", "plaza", "hill", "green", "rise", "vale", "wynd", "parade",
  "promenade", "esplanade", "embankment", "causeway", "wharf", "quay", "expressway", "expy", "freeway",
  "fwy", "turnpike", "tpke",
]);
/**
 * Street-type words that come before the name: right after the house number ("5 Rue de Rivoli"),
 * or first with the house number last ("Calle Mayor 12", "Via Roma 10").
 */
const LEADING_STREET_TYPES = new Set([
  "rue", "avenue", "avenida", "avda", "boulevard", "blvd", "place", "chemin", "allée", "allee", "impasse",
  "quai", "cours", "calle", "carrer", "via", "viale", "vicolo", "piazza", "piazzale", "corso", "largo",
  "rua", "travessa", "praça", "praca", "paseo", "plaza", "camino", "carretera", "ronda", "rambla", "strada",
]);
/**
 * Street names written as one word, or a street word after the name, with the house number last:
 * "Hauptstraße 5", "Bahnhofstr. 3", "Kerkstraat 12", "Berliner Straße 12".
 */
const STREET_ENDING_RE = /(?:straße|strasse|str\.?|weg|platz|gasse|allee|straat|laan|gracht|plein|gatan|vägen|vagen|gata|veien|katu)$/iu;
/** A street-type word and the word before it that together name something else: "4 Wheel Drive". */
const NOT_STREETS = new Set(["wheel drive", "hard drive", "disk drive", "flash drive", "test drive"]);
/** Compass points after the street type: "1600 Pennsylvania Avenue NW". */
const DIRECTIONS = new Set(["n", "s", "e", "w", "ne", "nw", "se", "sw", "north", "south", "east", "west"]);
/** Words for a flat or suite after the street type: "123 Main St Apt 4B". */
const UNIT_WORDS = new Set(["apt", "apartment", "suite", "ste", "unit", "flat", "fl", "floor", "rm", "room", "bldg", "building"]);
/** A house number: "12", "221B", "12-14", "401k". */
const HOUSE_NUMBER_RE = /^\d{1,5}[A-Za-z]?(?:[-/]\d{1,5}[A-Za-z]?)?$/u;
/** A word of a street name: letters (with ' . -), or an ordinal such as "5th". */
const STREET_WORD_RE = /^(?:\p{L}[\p{L}\p{M}'\u2019.-]*|\d{1,3}(?:st|nd|rd|th))$/iu;
const UNIT_ID_RE = /^#?[\p{L}\p{N}-]+$/u;
const TIME_WORDS = new Set([
  "second", "seconds", "sec", "secs", "minute", "minutes", "min", "mins", "hour", "hours", "hr", "hrs",
  "day", "days", "week", "weeks", "wk", "wks", "month", "months", "mo", "mos", "quarter", "quarters",
  "qtr", "qtrs", "year", "years", "yr", "yrs",
]);

/** A word lowercased, without a trailing dot: "St." gives "st". */
function bareWord(word: string): string {
  return word.toLowerCase().replace(/\.$/u, "");
}

function isStreetWord(word: string): boolean {
  return STREET_WORD_RE.test(word);
}

/** True when the words after a street type end the street line: nothing, a compass point, a flat or suite, or both. */
function isStreetEnd(words: readonly string[]): boolean {
  let i = 0;
  if (i < words.length && DIRECTIONS.has(bareWord(words[i]!))) i++;
  if (i < words.length && /^#[\p{L}\p{N}-]+$/u.test(words[i]!)) i++;
  else if (i + 1 < words.length && UNIT_WORDS.has(bareWord(words[i]!)) && UNIT_ID_RE.test(words[i + 1]!)) i += 2;
  return i === words.length;
}

/** One street line, without commas, in one of the forms described at isStreetAddress. */
function isStreetLine(line: string): boolean {
  const w = line.split(/\s+/u).filter((x) => x !== "");
  if (w.length < 2) return false;
  if (HOUSE_NUMBER_RE.test(w[0]!)) {
    // "5 Rue de Rivoli": the type right after the number, then the name.
    if (w.length >= 3 && LEADING_STREET_TYPES.has(bareWord(w[1]!)) && w.slice(2).every(isStreetWord)) return true;
    // "12 Baker Street", "1600 Pennsylvania Avenue NW": the name, then the type. "12 Month Close"
    // is a span of time.
    if (TIME_WORDS.has(bareWord(w[1]!))) return false;
    for (let k = 2; k < w.length; k++) {
      if (!isStreetWord(w[k - 1]!)) return false;
      const type = bareWord(w[k]!);
      if (STREET_TYPES.has(type) && !NOT_STREETS.has(`${bareWord(w[k - 1]!)} ${type}`) && isStreetEnd(w.slice(k + 1))) return true;
    }
    return false;
  }
  if (!HOUSE_NUMBER_RE.test(w[w.length - 1]!)) return false;
  const name = w.slice(0, -1);
  if (!name.every(isStreetWord)) return false;
  // "Calle Mayor 12", "Via Roma 10": the type, the name, then the number. "Place 1" alone is not one.
  if (name.length >= 2 && LEADING_STREET_TYPES.has(bareWord(name[0]!))) return true;
  // "Hauptstraße 5", "Berliner Straße 12": a street word ending the name, then the number. The
  // word must be joined to a name ("Hauptstraße") or follow one, so "Platz 1" (first place) is not one.
  return name.some((word, i) => {
    const m = STREET_ENDING_RE.exec(word);
    return m !== null && (word.length - m[0].length >= 2 || i > 0);
  });
}

/**
 * A street address that counts on its own, in any comma-separated part of the text:
 * - a house number, the name, then a street-type word, maybe followed by a compass point or a flat
 *   ("12 Baker Street", "221B Baker Street", "1600 Pennsylvania Avenue NW", "10 Downing St");
 * - a house number, then a street type and the name ("5 Rue de Rivoli");
 * - a street type and the name, then the house number ("Calle Mayor 12", "Via Roma 10");
 * - a name ending in a street word, then the house number ("Hauptstraße 5", "Kerkstraat 12").
 * "4 Wheel Drive" is not an address.
 */
function isStreetAddress(t: string): boolean {
  if (!/\d/u.test(t)) return false;
  const parts = t.split(",").map((p) => p.trim());
  for (let i = 0; i < parts.length; i++) {
    if (isStreetLine(parts[i]!)) return true;
    // "5, rue de Rivoli" and "Calle Mayor, 12": a comma inside the street line.
    if (i + 1 < parts.length && isStreetLine(`${parts[i]!} ${parts[i + 1]!}`)) return true;
  }
  return false;
}

/**
 * Text that starts with a house number and a word ("12 Baker", "10 Largest Orders", "3D Model
 * Name"). It reads as an address only over a column of addresses (see matchesColumn). A year
 * ("2024 Sales Plan") or a span of time ("12 Month Total") is not a house number.
 */
function isNumberLed(t: string): boolean {
  const w = t.replace(/,/gu, " ").split(/\s+/u).filter((x) => x !== "");
  if (w.length < 2 || !HOUSE_NUMBER_RE.test(w[0]!) || !isStreetWord(w[1]!)) return false;
  if (/^\d{4}$/u.test(w[0]!) && Number(w[0]) >= 1900 && Number(w[0]) <= 2100) return false;
  return !TIME_WORDS.has(bareWord(w[1]!).replace(/[.'\u2019-]+$/u, ""));
}

function isIdentifierText(t: string): boolean {
  if (EMAIL_EXACT_RE.test(t) || SSN_RE.test(t)) return true;
  if (isIban(t) || NI_NUMBER_RE.test(t) || isPostcode(t)) return true;
  if (!PHONE_SHAPE_RE.test(t)) return false;
  const digits = t.replace(/\D/g, "").length;
  return digits >= 10 && digits <= 15;
}

function isIdText(t: string): boolean {
  return t.length >= 5 && t.length <= 15 && /\p{N}/u.test(t) && ID_SHAPE_RE.test(t);
}

/** How a piece of text reads (see CellShape); `t` is trimmed and not empty. */
function textShape(t: string): Exclude<CellShape, "empty"> {
  if (isIdentifierText(t)) return "identifier";
  if (isStreetAddress(t)) return "address";
  if (PLAIN_NUMBER_RE.test(t) || AMOUNT_RE.test(t) || DATE_RE.test(t) || TIME_RE.test(t)) return "value";
  // Digits with phone punctuation and too few digits for a full phone number ("555 0199").
  if (PHONE_SHAPE_RE.test(t) && t.replace(/\D/g, "").length >= 7) return "value";
  if (isIdText(t)) return "id";
  return "text";
}

interface Shape {
  shape: CellShape;
  /** The text that gave the shape: the display text, else the raw value. */
  form: string;
}

/**
 * How a cell reads. The display text decides first (a cell showing "€2,000" is an amount, one
 * showing "(831) 555-0199" a phone number), then the value: any number or boolean is a value,
 * whatever else its display looks like ("EMP-00123" from a custom number format).
 */
function readCell(text: string | undefined, value: CellValue | undefined, valueType: string | undefined): Shape {
  if (isEmptyCell(value, text, valueType)) return { shape: "empty", form: "" };
  const shown = typeof text === "string" ? text.trim() : "";
  const raw = value === null || value === undefined ? "" : String(value).trim();
  const shownShape = shown === "" ? "text" : textShape(shown);
  if (shownShape === "identifier" || shownShape === "address" || shownShape === "value") return { shape: shownShape, form: shown };
  // A number or boolean without a matching value type still counts: fail closed.
  const numberOrBoolean =
    valueType === "Double" || valueType === "Integer" || valueType === "Boolean" || typeof value === "number" || typeof value === "boolean";
  if (numberOrBoolean) return { shape: "value", form: raw };
  if (shownShape === "id") return { shape: "id", form: shown };
  if (raw !== "" && raw !== shown) {
    const shape = textShape(raw);
    if (shape !== "text") return { shape, form: raw };
  }
  return { shape: "text", form: shown || raw };
}

function cellShape(text: string | undefined, value: CellValue | undefined, valueType: string | undefined): CellShape {
  return readCell(text, value, valueType).shape;
}

function isTextCell(value: CellValue | undefined, valueType: string | undefined): boolean {
  return valueType === "String" || ((valueType === "" || valueType === undefined) && typeof value === "string");
}

/**
 * An ID in a text cell (digits stored as text included), or an identifier (email address, phone
 * number, SSN, IBAN, NI number, postcode) in any cell. A street address is not an ID: "Address1"
 * over a column of addresses is a header.
 */
function isIdCell(c: BelowCell): boolean {
  const { shape, form } = readCell(c.text, c.value, c.valueType);
  if (shape === "identifier") return true;
  return isTextCell(c.value, c.valueType) && isIdText(form);
}

/** True when most of the cells pass `test`, or when there are none to compare with (fail closed). */
function mostly(cells: readonly BelowCell[], test: (c: BelowCell) => boolean): boolean {
  if (cells.length === 0) return true;
  return cells.filter(test).length > cells.length / 2;
}

/** A street address, or text that starts with a house number and a word ("12 Baker", "3 Elm"). */
function isAddressCell(c: BelowCell): boolean {
  const { shape, form } = readCell(c.text, c.value, c.valueType);
  return shape === "address" || (shape === "text" && isNumberLed(form));
}

/** A postcode, or a code of capital letters and digits with a space ("AB12 CD34"). */
function isCodeCell(c: BelowCell): boolean {
  const { form } = readCell(c.text, c.value, c.valueType);
  return isPostcode(form) || isCodeText(form);
}

/**
 * True when text that reads as a header on its own is data once compared with the column below:
 * text that starts with a house number ("12 Baker", "10 Largest Orders") over a column that is
 * mostly addresses, or a code such as "A1 B2" over a column that is mostly postcodes or codes. An
 * empty column counts (fail closed), as it does for ID-shaped text.
 */
function matchesColumn(form: string, below: readonly BelowCell[]): boolean {
  if (isNumberLed(form) && mostly(below, isAddressCell)) return true;
  return isCodeText(form) && mostly(below, isCodeCell);
}

/** Up to 200 non-empty data cells below the header cell of column offset `c`. */
export function cellsBelow(data: RangeData, c: number, limit = BELOW_SAMPLE): BelowCell[] {
  const out: BelowCell[] = [];
  const rows = rowCount(data);
  for (let r = 1; r < rows && out.length < limit; r++) {
    const text = data.text[r]?.[c] ?? "";
    const value = data.values[r]?.[c] ?? null;
    const valueType = data.valueTypes[r]?.[c] ?? "";
    if (!isEmptyCell(value, text, valueType)) out.push({ text, value, valueType });
  }
  return out;
}

/**
 * True when a cell looks like a data value rather than a header (fail closed):
 * - an email address, phone number, SSN, IBAN, UK National Insurance number, postcode (UK,
 *   Canadian, ZIP+4, Irish, Dutch, "123 45") or street address with a street-type word ("12 Baker
 *   Street", "Hauptstraße 5"), whatever is below;
 * - any number, boolean, amount, percentage, date or time, a year included, whatever is below;
 * - ID-shaped text. It counts as a header only when it has a run of two or more letters ("FY2024",
 *   "Address1", "Test1") and, with `below` (the column's data cells, see cellsBelow), the column
 *   below is not mostly IDs. "B-201", "E10234" and "Q1_2026" are data. Without `below`, every
 *   ID-shaped cell counts;
 * - with `below`, text that starts with a house number ("12 Baker") over a column of addresses, or
 *   a code such as "A1 B2" over a column of postcodes or codes (see matchesColumn). Over other
 *   data, "10 Largest Orders", "4 Wheel Drive" and "A1 B2" are headers.
 */
export function cellLooksLikeData(text: string, value: CellValue, valueType: string, below?: readonly BelowCell[]): boolean {
  const { shape, form } = readCell(text, value, valueType);
  switch (shape) {
    case "identifier":
    case "address":
    case "value":
      return true;
    case "id":
      return below === undefined || !LETTER_RUN_RE.test(form) || mostly(below, isIdCell);
    case "text":
      return below !== undefined && matchesColumn(form, below);
    default:
      return false;
  }
}

/** cellLooksLikeData for the header cell of column offset `c`, compared with the data below it. */
export function headerCellLooksLikeData(data: RangeData, c: number): boolean {
  return cellLooksLikeData(data.text[0]?.[c] ?? "", data.values[0]?.[c] ?? null, data.valueTypes[0]?.[c] ?? "", cellsBelow(data, c));
}

/**
 * True when the header cell of column offset `c` identifies someone or something: an email address,
 * phone number, SSN, IBAN, National Insurance number, postcode or street address; text that
 * starts with a house number over a column of addresses, or a code over a column of codes (see
 * matchesColumn); or ID-shaped text (digits stored as text of 5 or more included) that counts as
 * data (see cellLooksLikeData). Bare numbers, amounts and dates don't: a year header such as 2024
 * would block every request that mentions the year.
 */
export function headerCellIsIdentifier(data: RangeData, c: number): boolean {
  const text = data.text[0]?.[c];
  const value = data.values[0]?.[c];
  const valueType = data.valueTypes[0]?.[c];
  const { shape, form } = readCell(text, value, valueType);
  if (shape === "identifier" || shape === "address") return true;
  if (shape === "text") return matchesColumn(form, cellsBelow(data, c));
  if (!isTextCell(value, valueType) || !isIdText(form)) return false;
  // Digits stored as text ("02134", "10234") are IDs; other numbers written as text are values.
  const idLike = shape === "id" || /^\d+$/u.test(form);
  return idLike && headerCellLooksLikeData(data, c);
}

/** True when the header cell of column offset `c` is a bare number, boolean, amount, date or time. */
export function headerCellIsBareValue(data: RangeData, c: number): boolean {
  return cellShape(data.text[0]?.[c], data.values[0]?.[c], data.valueTypes[0]?.[c]) === "value";
}

/** True when any non-empty cell of the first row looks like data, compared with the data below it. */
export function firstRowLooksLikeData(data: RangeData): boolean {
  const cols = columnCount(data);
  for (let c = 0; c < cols; c++) {
    if (headerCellLooksLikeData(data, c)) return true;
  }
  return false;
}
