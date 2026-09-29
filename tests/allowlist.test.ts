import { describe, expect, it } from "vitest";
import { ALLOWLIST, EXCLUDED, EXCLUDED_REASONS, isAllowedFunction, normalizeFunctionName } from "../src/core/allowlist";

// Copied from spec §7.11. Extend only with a test per addition.
const SPEC_ALLOWLIST = `
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
`
  .split(/\s+/)
  .filter(Boolean);

// Spec §7.11: explicitly excluded, "every CUBE function" spelled out.
const SPEC_EXCLUDED = [
  "WEBSERVICE",
  "FILTERXML",
  "ENCODEURL",
  "IMAGE",
  "HYPERLINK",
  "RTD",
  "CUBEVALUE",
  "CUBEMEMBER",
  "CUBESET",
  "CUBESETCOUNT",
  "CUBERANKEDMEMBER",
  "CUBEMEMBERPROPERTY",
  "CUBEKPIMEMBER",
  "STOCKHISTORY",
  "PY",
  "COPILOT",
  "TRANSLATE",
  "DETECTLANGUAGE",
  "INDIRECT",
  "OFFSET",
  "CELL",
  "INFO",
  "CALL",
  "REGISTER.ID",
  "GETPIVOTDATA",
];

describe("ALLOWLIST", () => {
  it("equals the spec §7.11 list exactly", () => {
    expect(SPEC_ALLOWLIST).toHaveLength(154);
    expect(new Set(SPEC_ALLOWLIST).size).toBe(154);
    expect(ALLOWLIST.size).toBe(154);
    expect([...ALLOWLIST].sort()).toEqual([...SPEC_ALLOWLIST].sort());
  });

  it("contains a few expected members", () => {
    for (const name of ["SUM", "SUMIFS", "XLOOKUP", "LET", "LAMBDA", "CEILING.MATH", "STDEV.S", "NETWORKDAYS.INTL", "TRUE", "N", "REGEXREPLACE"]) {
      expect(ALLOWLIST.has(name), name).toBe(true);
    }
  });

  it("does not contain lower-case, prefixed or unlisted names", () => {
    for (const name of ["sum", "_xlfn.SUM", "WEBSERVICE", "INDIRECT", "OFFSET", "TRANSPOSE", "ISOMITTED", "T", "EVALUATE"]) {
      expect(ALLOWLIST.has(name), name).toBe(false);
    }
  });

  it("holds only upper-case ASCII names", () => {
    for (const name of ALLOWLIST) expect(name).toMatch(/^[A-Z][A-Z0-9.]*$/);
  });
});

describe("EXCLUDED", () => {
  it("equals the spec §7.11 exclusions, with every CUBE function", () => {
    expect([...EXCLUDED].sort()).toEqual([...SPEC_EXCLUDED].sort());
    expect(EXCLUDED.filter((n) => n.startsWith("CUBE"))).toHaveLength(7);
  });

  it("never overlaps the allowlist", () => {
    for (const name of EXCLUDED) expect(ALLOWLIST.has(name), name).toBe(false);
  });

  it("has a plain reason for every excluded function", () => {
    for (const name of EXCLUDED) {
      const reason = EXCLUDED_REASONS[name];
      expect(reason, name).toBeTruthy();
      expect(reason).not.toMatch(/encrypt|secure|compliant|anonymous/i);
    }
    expect(EXCLUDED_REASONS.WEBSERVICE).toBe("can send data over the internet");
  });
});

describe("normalizeFunctionName", () => {
  it("strips _xlfn., _xlws. and _xlpm. prefixes, case-insensitively and repeatedly", () => {
    expect(normalizeFunctionName("_xlfn.IMAGE")).toBe("IMAGE");
    expect(normalizeFunctionName("_xlfn._xlws.SORT")).toBe("SORT");
    expect(normalizeFunctionName("_XLFN.webservice")).toBe("WEBSERVICE");
    expect(normalizeFunctionName("_xlpm.x")).toBe("X");
    expect(normalizeFunctionName("_xlws.filter")).toBe("FILTER");
  });

  it("keeps other prefixes", () => {
    expect(normalizeFunctionName("_xludf.SUM")).toBe("_XLUDF.SUM");
    expect(isAllowedFunction("_xludf.SUM")).toBe(false);
  });

  it("upper-cases ASCII letters only, so look-alike letters never match", () => {
    expect(normalizeFunctionName("sum")).toBe("SUM");
    // Cyrillic small i, dotless i and long s all upper-case to Latin letters with toUpperCase().
    expect(normalizeFunctionName("webservіce")).not.toBe("WEBSERVICE");
    expect(isAllowedFunction("ſum")).toBe(false);
    expect(isAllowedFunction("mın")).toBe(false);
    expect(isAllowedFunction("_xlfn.xlookup")).toBe(true);
    expect(isAllowedFunction("ceiling.math")).toBe(true);
  });
});
