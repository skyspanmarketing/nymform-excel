// Function allowlist (spec §7.11). Extend only with a test per addition.
// An allowlist rather than a blocklist: Excel keeps adding functions that reach the network or a
// cloud service, so anything not listed here is rejected by the formula gate.

const NAMES = `
ABS CEILING.MATH FLOOR.MATH INT MOD POWER PRODUCT QUOTIENT ROUND ROUNDDOWN ROUNDUP MROUND SIGN SQRT
SUM SUMIF SUMIFS SUMPRODUCT TRUNC EXP LN LOG LOG10 PI
AVERAGE AVERAGEIF AVERAGEIFS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS MAX MAXIFS MIN MINIFS MEDIAN
MODE.SNGL STDEV.S STDEV.P VAR.S VAR.P LARGE SMALL RANK.EQ RANK.AVG PERCENTILE.INC QUARTILE.INC CORREL
IF IFS IFERROR IFNA AND OR NOT XOR SWITCH TRUE FALSE
XLOOKUP XMATCH VLOOKUP HLOOKUP INDEX MATCH CHOOSE CHOOSECOLS CHOOSEROWS ROW ROWS COLUMN COLUMNS
TAKE DROP EXPAND VSTACK HSTACK TOCOL TOROW WRAPROWS WRAPCOLS
FILTER SORT SORTBY UNIQUE SEQUENCE GROUPBY PIVOTBY LET LAMBDA MAP REDUCE SCAN BYROW BYCOL MAKEARRAY
LEFT RIGHT MID LEN FIND SEARCH SUBSTITUTE REPLACE TRIM CLEAN UPPER LOWER PROPER CONCAT TEXTJOIN TEXT
VALUE NUMBERVALUE TEXTBEFORE TEXTAFTER TEXTSPLIT EXACT REPT REGEXTEST REGEXEXTRACT REGEXREPLACE
DATE DATEVALUE DAY MONTH YEAR TODAY NOW EDATE EOMONTH NETWORKDAYS NETWORKDAYS.INTL WORKDAY WORKDAY.INTL
WEEKDAY WEEKNUM ISOWEEKNUM DAYS DATEDIF YEARFRAC TIME HOUR MINUTE SECOND
ISBLANK ISNUMBER ISTEXT ISERROR ISNA ISLOGICAL NA N
`;

export const ALLOWLIST: ReadonlySet<string> = new Set<string>(NAMES.split(/\s+/).filter(Boolean));

/** Why each excluded function is rejected, in plain words ("Uses X, which {reason}."). */
export const EXCLUDED_REASONS: Readonly<Record<string, string>> = {
  WEBSERVICE: "can send data over the internet",
  FILTERXML: "works with content fetched from the web",
  ENCODEURL: "prepares values to be sent in a web address",
  IMAGE: "loads a picture from the internet and can send data with the request",
  HYPERLINK: "creates a link that can send data when clicked",
  RTD: "connects to a program outside Excel",
  CUBEVALUE: "connects to an outside data source",
  CUBEMEMBER: "connects to an outside data source",
  CUBESET: "connects to an outside data source",
  CUBESETCOUNT: "connects to an outside data source",
  CUBERANKEDMEMBER: "connects to an outside data source",
  CUBEMEMBERPROPERTY: "connects to an outside data source",
  CUBEKPIMEMBER: "connects to an outside data source",
  STOCKHISTORY: "fetches data from the internet",
  PY: "runs Python code in the Microsoft cloud",
  COPILOT: "sends data to an AI service",
  TRANSLATE: "sends text to an online service",
  DETECTLANGUAGE: "sends text to an online service",
  INDIRECT: "can build references outside the ranges you chose",
  OFFSET: "can reach cells outside the ranges you chose",
  CELL: "can read details about the workbook and where it is saved",
  INFO: "can read details about your computer",
  CALL: "can run code outside Excel",
  "REGISTER.ID": "can run code outside Excel",
  GETPIVOTDATA: "reads pivot tables outside the ranges you chose",
};

/** Functions explicitly excluded and covered by tests. */
export const EXCLUDED: readonly string[] = Object.keys(EXCLUDED_REASONS);

const PREFIX_RE = /^_xl(?:fn|ws|pm)\./i;

/**
 * Strips any number of `_xlfn.`, `_xlws.` and `_xlpm.` prefixes (any case) and upper-cases ASCII
 * letters only, so a look-alike letter from another script never turns into a Latin one.
 */
export function normalizeFunctionName(name: string): string {
  let n = name;
  while (PREFIX_RE.test(n)) n = n.replace(PREFIX_RE, "");
  return n.replace(/[a-z]+/g, (s) => s.toUpperCase());
}

/** True when `name`, after normalizing, is on the allowlist by exact ASCII match. */
export function isAllowedFunction(name: string): boolean {
  return ALLOWLIST.has(normalizeFunctionName(name));
}
