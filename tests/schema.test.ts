import { describe, expect, it } from "vitest";
import {
  buildSheetContext,
  cellLooksLikeData,
  cellsBelow,
  firstRowLooksLikeData,
  headerCellIsBareValue,
  headerCellIsIdentifier,
  headerCellLooksLikeData,
  inferSchema,
  isDateFormat,
} from "../src/core/schema";
import { isEmailText, isIdShaped, isPhoneText, isSsnText } from "../src/core/suggest";
import type { CellValue, RangeData } from "../src/core/types";

/** A range as the adapter would hand it over. Formats apply per column (default "General"). */
function range(
  rows: CellValue[][],
  opts: { sheet?: string; rowIndex?: number; columnIndex?: number; formats?: Record<number, string>; text?: (string | undefined)[][] } = {},
): RangeData {
  const rowIndex = opts.rowIndex ?? 0;
  const columnIndex = opts.columnIndex ?? 0;
  const width = Math.max(...rows.map((r) => r.length));
  const end = `${String.fromCharCode(65 + columnIndex + width - 1)}${rowIndex + rows.length}`;
  return {
    sheet: opts.sheet ?? "Sheet1",
    address: `${String.fromCharCode(65 + columnIndex)}${rowIndex + 1}:${end}`,
    rowIndex,
    columnIndex,
    values: rows,
    text: rows.map((r, i) =>
      r.map((v, j) => opts.text?.[i]?.[j] ?? (v === null ? "" : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v))),
    ),
    formulas: rows.map((r) => r.map((v) => v)),
    valueTypes: rows.map((r) =>
      r.map((v) =>
        v === null || v === ""
          ? "Empty"
          : typeof v === "number"
            ? Number.isInteger(v)
              ? "Integer"
              : "Double"
            : typeof v === "boolean"
              ? "Boolean"
              : v.startsWith("#")
                ? "Error"
                : "String",
      ),
    ),
    numberFormat: rows.map((r) => r.map((_, j) => opts.formats?.[j] ?? "General")),
  };
}

describe("isDateFormat", () => {
  it.each([
    "m/d/yyyy",
    "[$-409]mmm d, yyyy",
    "d-mmm-yy",
    "yyyy-mm-dd",
    "dddd, mmmm d",
    "mmm-yy",
    "[$-en-US]m/d/yy h:mm AM/PM",
    'yyyy "year"',
  ])("treats %s as a date", (fmt) => {
    expect(isDateFormat(fmt)).toBe(true);
  });

  it.each([
    "0.00",
    '"days" 0',
    "[Red]0",
    "General",
    "#,##0",
    "0%",
    "@",
    "0.00E+00",
    "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
    "\\d0",
    '0 "mm"',
    "[Color10]0;[Red]-0",
    "h AM/PM",
    "",
  ])("does not treat %s as a date", (fmt) => {
    expect(isDateFormat(fmt)).toBe(false);
  });

  it("counts minutes in a time format as a date token (time values are date serials)", () => {
    expect(isDateFormat("h:mm")).toBe(true);
  });
});

describe("inferSchema", () => {
  it("infers types with the 80% rule", () => {
    const data = range(
      [
        ["Name", "Amount", "Order date", "Paid", "Mixed", "Nothing", "Mostly numbers"],
        ["Ana", 10, 46085, true, 1, null, 1],
        ["Ben", 20.5, 46086, false, "x", null, 2],
        ["Cy", 30, 46087, true, 2, null, 3],
        ["Di", 40, 46088, true, "y", null, 4],
        ["Ed", 50, 46089, false, 3, null, "five"],
      ],
      { formats: { 2: "m/d/yyyy" } },
    );
    const cols = inferSchema(data);
    expect(cols.map((c) => c.type)).toEqual(["text", "number", "date", "boolean", "mixed", "empty", "number"]);
    expect(cols.map((c) => c.letter)).toEqual(["A", "B", "C", "D", "E", "F", "G"]);
  });

  it("falls back to mixed below 80%", () => {
    const data = range([["H"], [1], [2], [3], ["a"], ["b"]]);
    expect(inferSchema(data)[0]!.type).toBe("mixed");
    const eighty = range([["H"], [1], [2], [3], [4], ["a"]]);
    expect(inferSchema(eighty)[0]!.type).toBe("number");
  });

  it("counts blanks and distinct values; text columns also get avgWords (one decimal) and maxLen", () => {
    const data = range([
      ["Customer", "Amount"],
      ["Maria Lopez", 5],
      ["Maria Lopez", 5],
      [null, 7],
      ["Jo", null],
      ["Ann Marie Smith", 7],
    ]);
    const [a, b] = inferSchema(data);
    expect(a!.stats).toEqual({ blank: 1, distinct: 3, avgWords: 2, maxLen: 15 });
    expect(b!.stats).toEqual({ blank: 1, distinct: 2 });
    const words = range([["Note"], ["one two"], ["one two three"], ["one"]]);
    expect(inferSchema(words)[0]!.stats.avgWords).toBe(2);
    const third = range([["Note"], ["a b"], ["a"], ["a"]]);
    expect(inferSchema(third)[0]!.stats.avgWords).toBe(1.3);
  });

  it('names empty headers "Column {letter}" and uses sheet letters', () => {
    const data = range(
      [
        ["", "Amount", "  "],
        [1, 2, 3],
      ],
      { columnIndex: 2 },
    );
    expect(inferSchema(data).map((c) => [c.letter, c.header])).toEqual([
      ["C", "Column C"],
      ["D", "Amount"],
      ["E", "Column E"],
    ]);
  });

  it("treats every row as data when hasHeaders is false", () => {
    const data = range([
      ["Name", 1],
      ["Ana", 2],
    ]);
    const cols = inferSchema(data, false);
    expect(cols.map((c) => c.header)).toEqual(["Column A", "Column B"]);
    expect(cols[0]!.stats.distinct).toBe(2);
    expect(cols[1]!.type).toBe("number");
  });

  it("types an all-empty column as empty and error cells count against the 80% rule", () => {
    const data = range([["E", "Err"], [null, 1], [null, "#N/A"], [null, "#N/A"]]);
    const [e, err] = inferSchema(data);
    expect(e!.type).toBe("empty");
    expect(e!.stats).toEqual({ blank: 3, distinct: 0 });
    expect(err!.type).toBe("mixed");
  });

  it("never contains cell values", () => {
    const values = [
      "Quorbel Vantrisk",
      "zyxw.qvut@example.test",
      "(831) 555-0199",
      "Ana \"AJ\" Ruiz",
      "Diagnosed with something rare",
    ];
    const rows: CellValue[][] = [["Customer", "Email", "Phone", "Nick", "Notes", "Amount", "Visit", "Flag"]];
    for (let i = 0; i < values.length; i++) {
      rows.push([`${values[0]} ${i}`, values[1]!, values[2]!, values[3]!, values[4]!, 987654.321 + i, 46085 + i, i % 2 === 0]);
    }
    const data = range(rows, { formats: { 6: "yyyy-mm-dd" } });
    const serialized = JSON.stringify(inferSchema(data)) + JSON.stringify(inferSchema(data, false));
    for (const v of values) {
      expect(serialized).not.toContain(v);
      for (const w of v.split(/\s+/)) if (w.length >= 4) expect(serialized).not.toContain(w);
    }
    for (let i = 0; i < values.length; i++) {
      expect(serialized).not.toContain(String(987654.321 + i));
      expect(serialized).not.toContain(String(46085 + i));
    }
    expect(serialized).not.toContain("98765");
    expect(serialized).not.toContain("4608");
    expect(serialized).not.toMatch(/true|false/);
    // With headers off, the first row is data: its text must not become a header either.
    expect(JSON.stringify(inferSchema(data, false))).not.toContain("Customer");
  });
});

describe("buildSheetContext", () => {
  it("computes row numbers, sheets, tables and allowed ranges", () => {
    const main = range(
      [
        ["Customer", "Amount"],
        ["Ana", 1],
        ["Ben", 2],
      ],
      { sheet: "Orders", rowIndex: 4, columnIndex: 1 },
    );
    main.table = { name: "OrdersTable", columns: ["Customer", "Amount"] };
    const ctxRange = range([["Region", "Rate"], ["East", 1]], { sheet: "My Regions" });
    const ctx = buildSheetContext(main, true, [ctxRange], {
      sheets: ["Orders", "My Regions", "Other"],
      tables: [{ name: "OrdersTable", sheet: "Orders", address: "B5:C7" }],
    });
    expect(ctx).toEqual({
      sheet: "Orders",
      table: "OrdersTable",
      address: "B5:C7",
      headerRow: 5,
      firstDataRow: 6,
      lastDataRow: 7,
      sheets: ["Orders", "My Regions", "Other"],
      tableNames: ["OrdersTable"],
      allowedRanges: ["Orders!B5:C7", "'My Regions'!A1:B2"],
      tables: [{ name: "OrdersTable", sheet: "Orders", address: "B5:C7" }],
    });
  });

  it("uses header row 0 and starts data at the top when there are no headers", () => {
    const main = range([[1], [2], [3]], { sheet: "Data" });
    const ctx = buildSheetContext(main, false, [], { sheets: ["Data"], tables: [] });
    expect(ctx.headerRow).toBe(0);
    expect(ctx.firstDataRow).toBe(1);
    expect(ctx.lastDataRow).toBe(3);
    expect(ctx.table).toBeUndefined();
    expect(ctx.allowedRanges).toEqual(["Data!A1:A3"]);
  });
});

describe("cellLooksLikeData", () => {
  it("numbers, dates and booleans look like data", () => {
    expect(cellLooksLikeData("2024", 2024, "Double")).toBe(true);
    expect(cellLooksLikeData("3/15/2026", 46096, "Integer")).toBe(true);
    expect(cellLooksLikeData("TRUE", true, "Boolean")).toBe(true);
    // Without a value type, the raw value decides.
    expect(cellLooksLikeData("83500", 83500, "")).toBe(true);
  });

  it("text shaped like an email, phone number, SSN or ID looks like data; header words don't", () => {
    for (const t of ["maria.lopez@example.com", "(831) 555-0199", "+1 831 555 0199", "123-45-6789", "EMP-00123", "A1B2C3"]) {
      expect(cellLooksLikeData(t, t, "String"), t).toBe(true);
    }
    for (const t of ["Customer name", "Amount", "Email", "Phone", "Q1 total", "ID", "Maria Lopez", "  "]) {
      expect(cellLooksLikeData(t, t, "String"), t).toBe(false);
    }
    expect(cellLooksLikeData("", null, "Empty")).toBe(false);
  });

  it("uses the same shapes as suggest.ts", () => {
    const corpus = [
      "maria.lopez@example.com", "not an email@", "(831) 555-0199", "831-555-0199 x12", "555-0199", "+44 20 7946 0958",
      "123-45-6789", "123456789", "123 45 6789", "EMP-00123", "S1234567", "AB12", "ABCDEFG", "2024-Q1", "Region",
      "Customer name", "12345678901234567", "#N/A", "  ", "Ünïcode9", "a.b/c#1",
    ];
    for (const t of corpus) {
      // A number written as text is data too (third review), whatever suggest.ts calls it.
      const shaped = isEmailText(t) || isPhoneText(t) || isSsnText(t) || isIdShaped(t) || /^\d+$/.test(t);
      expect(cellLooksLikeData(t, t, "String"), t).toBe(shaped);
    }
  });
});

describe("firstRowLooksLikeData", () => {
  it("is false for a header row and true when the first row is data", () => {
    expect(firstRowLooksLikeData(range([["Customer", "Email", "Amount"], ["Maria Lopez", "maria@example.com", 100]]))).toBe(false);
    expect(firstRowLooksLikeData(range([["Maria Lopez", "maria@example.com", 100], ["Ana Ruiz", "ana@example.com", 200]]))).toBe(true);
  });

  it("needs one data-like cell; empty cells are skipped", () => {
    expect(firstRowLooksLikeData(range([["Name", null, 83500], ["Maria Lopez", null, 1]]))).toBe(true);
    expect(firstRowLooksLikeData(range([["Name", null, "Year"], ["Maria Lopez", null, 2024]]))).toBe(false);
    expect(firstRowLooksLikeData(range([["Name", "(831) 555-0199"]]))).toBe(true);
  });
});

describe("header detection (second review)", () => {
  it("treats amounts, percentages, grouped numbers, dates with month names and times as data", () => {
    for (const t of ["$83,500", "83,500", "€1.234,50", "83 500 ₽", "(1,200.00)", "12.5%", "-3%", "1.234.567"]) {
      expect(cellLooksLikeData(t, t, "String"), t).toBe(true);
    }
    for (const t of ["15 Mar 2024", "March 15, 2024", "Wed, 15-Mar-24", "4th July 2025", "2024-03-15 10:30", "15/03/26"]) {
      expect(cellLooksLikeData(t, t, "String"), t).toBe(true);
    }
    for (const t of ["10:30 AM", "14:05", "9:30:15 pm"]) expect(cellLooksLikeData(t, t, "String"), t).toBe(true);
    // Header words with numbers or month names are still headers.
    for (const t of ["Mar 2024", "Q1 2024", "Sales 2024", "May", "Score (0-100)", "Item #1"]) {
      expect(cellLooksLikeData(t, t, "String"), t).toBe(false);
    }
  });

  it("counts an ID-shaped header with a word in it as data only over a column of IDs", () => {
    const below = (rows: CellValue[][]) => firstRowLooksLikeData(range(rows));
    for (const h of ["FY2024", "Address1", "Test1"]) {
      expect(below([["Name", h], ["Maria Lopez", 1200], ["Ana Ruiz", 1300]]), h).toBe(false);
      expect(below([["Name", h], ["Maria Lopez", "12 Main St"], ["Ana Ruiz", "3 Elm Rd"]]), h).toBe(false);
    }
    expect(below([["Name", "EMP-00123"], ["Maria Lopez", "EMP-00124"], ["Ana Ruiz", "EMP-00125"]])).toBe(true);
    // Numbers below aren't IDs, even when their digits look like one.
    expect(below([["Name", "EMP-00123"], ["Maria Lopez", 100124], ["Ana Ruiz", 100125]])).toBe(false);
    // With nothing below to compare with, it counts (fail closed).
    expect(below([["Name", "EMP-00123"]])).toBe(true);
  });

  it("always counts emails, phone numbers and SSNs", () => {
    for (const h of ["maria@example.com", "(831) 555-0199", "123-45-6789"]) {
      expect(firstRowLooksLikeData(range([["Region", h], ["East", 1200], ["West", 1300]])), h).toBe(true);
    }
    expect(firstRowLooksLikeData(range([["Region", 8315550199], ["East", 1200]]))).toBe(true);
  });

  it("without the data below, cellLooksLikeData keeps counting IDs", () => {
    expect(cellLooksLikeData("FY2024", "FY2024", "String")).toBe(true);
    expect(cellLooksLikeData("2024", 2024, "Double")).toBe(true);
    expect(cellLooksLikeData("FY2024", "FY2024", "String", [{ text: "1200", value: 1200, valueType: "Double" }])).toBe(false);
    expect(cellLooksLikeData("2024", 2024, "Double", [{ text: "2023", value: 2023, valueType: "Double" }])).toBe(true);
    const d = range([["Name", "FY2024", 2024], ["Maria Lopez", "FY2023", 1200]]);
    expect(cellsBelow(d, 1)).toEqual([{ text: "FY2023", value: "FY2023", valueType: "String" }]);
    expect(headerCellLooksLikeData(d, 1)).toBe(true);
    // A year counts as data whatever is below it (third review).
    expect(headerCellLooksLikeData(d, 2)).toBe(true);
  });

  it("tells identifier headers from bare numbers and dates, for the auditor", () => {
    const d = range(
      [
        ["maria@example.com", "EMP-00123", "FY2024", 2024, "15 Mar 2024", "$83,500", "Customer"],
        ["ana@example.com", "EMP-00124", 1200, 1300, 1, 2, "Ana Ruiz"],
      ],
      {},
    );
    expect([0, 1, 2, 3, 4, 5, 6].map((c) => headerCellIsIdentifier(d, c))).toEqual([true, true, false, false, false, false, false]);
    expect([0, 1, 2, 3, 4, 5, 6].map((c) => headerCellIsBareValue(d, c))).toEqual([false, false, false, true, true, true, false]);
  });
});

describe("header detection (third review: fail closed)", () => {
  const below = (rows: CellValue[][], text?: (string | undefined)[][]) => firstRowLooksLikeData(range(rows, text ? { text } : {}));

  it("counts any number in row 1 as data, a year included, and numbers stored as text", () => {
    expect(below([["Account", 2023, 2024], ["Rent", 10000, 11000], ["Travel", 11375, 12450]])).toBe(true);
    expect(below([["Maria Lopez", "Sales", 2000], ["Ana Ruiz", "Ops", 2350]])).toBe(true);
    expect(below([["Maria Lopez", 2045, "Sales"], ["Ana Ruiz", 3120, "Ops"]])).toBe(true);
    expect(below([["2001", "Maria Lopez"], ["2002", "Ana Ruiz"]])).toBe(true);
    for (const t of ["2000", "02134", "12.5", "3,75", "-5", "555 0199", "555-0199", "２０２４", "٢٠٢٤"]) {
      expect(cellLooksLikeData(t, t, "String"), t).toBe(true);
    }
  });

  it("reads the display text first: a cell showing €2,000 is an amount", () => {
    const d = range([["Maria Lopez", "Sales", 2000], ["Ana Ruiz", "Ops", 2350]], { text: [["Maria Lopez", "Sales", "€2,000"], ["Ana Ruiz", "Ops", "€2,350"]] });
    expect(firstRowLooksLikeData(d)).toBe(true);
    expect(headerCellIsBareValue(d, 2)).toBe(true);
    expect(headerCellIsIdentifier(d, 2)).toBe(false);
    for (const t of ["€2,000", "2.000 €", "20 %", "15. März 1985", "15 mars 1985", "1985年3月15日", "1985년 3월 15일"]) {
      expect(cellLooksLikeData(t, t, "String"), t).toBe(true);
    }
    // A number shown through a custom format is still a number, even over a column of text.
    expect(cellLooksLikeData("EMP-00123", 123, "Double", [{ text: "Remote", value: "Remote", valueType: "String" }])).toBe(true);
  });

  it("counts an ID-shaped cell as a header only with a run of 2+ letters and not over a column of IDs", () => {
    for (const h of ["B-201", "E10234", "Q1_2026", "12-345"]) {
      expect(below([["Name", h], ["Maria Lopez", "Remote"], ["Ana Ruiz", "Home office"]]), h).toBe(true);
      expect(below([["Name", h], ["Maria Lopez", 1200], ["Ana Ruiz", 1300]]), h).toBe(true);
    }
    for (const h of ["FY2024", "Address1", "Test1"]) {
      expect(below([["Name", h], ["Maria Lopez", "Remote"], ["Ana Ruiz", "Home office"]]), h).toBe(false);
      expect(below([["Name", h], ["Maria Lopez", "EMP-00124"], ["Ana Ruiz", "EMP-00125"]]), h).toBe(true);
      expect(below([["Name", h], ["Maria Lopez", "10234"], ["Ana Ruiz", "10235"]]), h).toBe(true);
    }
  });

  it("tells the auditor which header cells identify someone: ID-shaped data and digits stored as text, not years", () => {
    const d = range([
      ["B-201", "FY2024", "10234", "2024", 10234],
      ["Remote", "Remote", "10235", "2025", 10235],
    ]);
    expect([0, 1, 2, 3, 4].map((c) => headerCellIsIdentifier(d, c))).toEqual([true, false, true, false, false]);
    expect([0, 1, 2, 3, 4].map((c) => headerCellIsBareValue(d, c))).toEqual([false, false, true, true, true]);
  });
});

describe("header detection (fourth review: IBANs, NI numbers, postcodes, addresses)", () => {
  const below = (rows: CellValue[][]) => firstRowLooksLikeData(range(rows));

  it("counts IBANs, UK National Insurance numbers, postcodes and street addresses as data", () => {
    for (const t of [
      "DE89 3704 0044 0532 0130 00",
      "GB29 NWBK 6016 1331 9268 19",
      "FR14 2004 1010 0505 0001 3M02 606",
      "DE89370400440532013000",
      "QQ 12 34 56 C",
      "QQ123456C",
      "SW1A 1AA",
      "M1 1AE",
      "LS1 4AP",
      "EC1A1BB",
      "94107-1234",
      "K1A 0B1",
      "D02 X285",
      "1234 AB",
      "114 55",
      "12 Baker Street",
      "3 High St",
      "221B Baker Street",
      "12-14 Baker St",
      "123 Main St",
      "1600 Pennsylvania Avenue NW",
    ]) {
      expect(cellLooksLikeData(t, t, "String"), t).toBe(true);
    }
  });

  it("keeps headers with periods, years, spans of time and codes as headers", () => {
    for (const t of ["Q1 2024", "H1 FY25", "Q1 FY24", "FY 2024", "YTD 2024", "Sales 2024", "12 Month Total", "30 Day Average", "3 Year CAGR", "2024 Sales Plan", "Top 10 Customers"]) {
      expect(cellLooksLikeData(t, t, "String"), t).toBe(false);
    }
  });

  it("finds a selection made below its header row, whatever the private column holds", () => {
    for (const cell of ["DE89 3704 0044 0532 0130 00", "QQ 12 34 56 C", "SW1A 1AA", "12 Baker Street"]) {
      const d = range([["Maria Lopez", cell, "London"], ["Ana Ruiz", "GB29 NWBK 6016 1331 9268 19", "Leeds"]]);
      expect(firstRowLooksLikeData(d), cell).toBe(true);
      expect(headerCellLooksLikeData(d, 1), cell).toBe(true);
      // Fail closed: the auditor scans such a header cell of a private column.
      expect(headerCellIsIdentifier(d, 1), cell).toBe(true);
      expect(headerCellIsBareValue(d, 1), cell).toBe(false);
    }
  });

  it("still reads an ID-shaped header over a column of addresses as a header", () => {
    expect(below([["Name", "Address1"], ["Maria Lopez", "12 Main St"], ["Ana Ruiz", "3 Elm Rd"]])).toBe(false);
  });
});

describe("header detection (fourth review, second pass: headers that start with a number)", () => {
  // Ordinary data columns: numbers, amounts shown with a currency, short text, dates, yes/no.
  const ordinary: [string, CellValue[], (string | undefined)[]?][] = [
    ["numbers", [1200, 800, 450, 975]],
    ["amounts", [1200, 800, 450, 975], ["$1,200", "$800", "$450", "$975"]],
    ["text", ["Great service", "Loved it", "Fast delivery", "Would buy again"]],
    ["dates", [46096, 46097, 46100, 46105], ["3/15/2026", "3/16/2026", "3/19/2026", "3/24/2026"]],
    ["yes/no", ["Yes", "No", "Yes", "Yes"]],
  ];
  /** Row 1 is Customer, Dept and `header`; column C below holds `cells`. */
  const sheet = (header: CellValue, cells: CellValue[], shown?: (string | undefined)[]) => {
    const names = ["Maria Lopez", "Ana Ruiz", "Kenji Watanabe", "Omar Haddad"];
    const rows: CellValue[][] = [["Customer", "Dept", header], ...cells.map((v, i) => [names[i % names.length]!, "Ops", v])];
    const text = shown ? [[undefined, undefined, undefined], ...shown.map((t) => [undefined, undefined, t])] : undefined;
    return range(rows, text ? { text } : {});
  };

  const HEADERS = [
    "10 Largest Orders",
    "401k Match Amount",
    "5 Star Reviews",
    "3D Model Name",
    "360 Review Score",
    "2 Factor Enabled",
    "3 Point Rating",
    "4 Wheel Drive",
    "7 Eleven Stores",
    "A1 B2",
    "Q1 2024",
    "2024 Sales Plan",
    "12 Month Total",
    "Address1",
  ];

  it("keeps headers that start with a number, codes and periods as headers over ordinary data", () => {
    for (const h of HEADERS) {
      for (const [kind, cells, shown] of ordinary) {
        const d = sheet(h, cells, shown);
        expect(firstRowLooksLikeData(d), `${h} over ${kind}`).toBe(false);
        expect(headerCellLooksLikeData(d, 2), `${h} over ${kind}`).toBe(false);
        // So a question that names the header is not blocked when the column is private.
        expect(headerCellIsIdentifier(d, 2), `${h} over ${kind}`).toBe(false);
      }
    }
  });

  it("keeps a period and a year shaped like an Irish postcode as a header", () => {
    for (const h of ["P12 2024", "W52 2024", "M01 2025"]) {
      expect(cellLooksLikeData(h, h, "String"), h).toBe(false);
      expect(firstRowLooksLikeData(sheet(h, [1200, 800, 450])), h).toBe(false);
    }
  });

  const ADDRESSES = [
    "12 Baker Street",
    "221B Baker Street",
    "1600 Pennsylvania Avenue NW",
    "10 Downing St",
    "742 Evergreen Terrace",
    "5 Rue de Rivoli",
    "Hauptstraße 5",
    "Calle Mayor 12",
  ];
  const IDENTIFIERS = [
    // IBANs
    "DE89 3704 0044 0532 0130 00",
    "GB29NWBK60161331926819",
    // UK National Insurance numbers
    "QQ 12 34 56 C",
    "AB123456C",
    // UK, Canadian, ZIP+4, Irish and Dutch postcodes
    "SW1A 1AA",
    "M1 1AE",
    "EC1A 1BB",
    "K1A 0B1",
    "M5V3L9",
    "94107-1234",
    "D02 X285",
    "T12 X70A",
    "1012 AB",
  ];

  it("counts street addresses, IBANs, NI numbers and postcodes as data on their own", () => {
    for (const t of [...ADDRESSES, ...IDENTIFIERS]) {
      expect(cellLooksLikeData(t, t, "String"), t).toBe(true);
      for (const [kind, cells, shown] of ordinary) {
        const d = sheet(t, cells, shown);
        expect(firstRowLooksLikeData(d), `${t} over ${kind}`).toBe(true);
        // Fail closed: the auditor scans such a header cell of a private column.
        expect(headerCellIsIdentifier(d, 2), `${t} over ${kind}`).toBe(true);
        expect(headerCellIsBareValue(d, 2), `${t} over ${kind}`).toBe(false);
      }
    }
  });

  it("reads other street address forms, and only a street-type word makes one on its own", () => {
    for (const t of [
      "12 Oak Drive",
      "350 5th Avenue",
      "123 Main St Apt 4B",
      "Flat 2, 14 Elm Grove",
      "12 Baker Street, London NW1 6XE",
      "5, rue de Rivoli",
      "Calle Mayor, 12",
      "Via Roma 10",
      "Berliner Straße 12",
      "Bahnhofstr. 3",
      "Kerkstraat 12",
    ]) {
      expect(cellLooksLikeData(t, t, "String"), t).toBe(true);
    }
    // No street-type word, a drive that is not a street, first place, a span of time.
    for (const t of ["12 Baker", "4 Wheel Drive", "Platz 1", "Place 1", "10 Day Close", "Via 2"]) {
      expect(cellLooksLikeData(t, t, "String"), t).toBe(false);
    }
  });

  const ADDRESS_COLUMN = ["3 High St", "77 Queen Road", "Hauptstraße 5", "14 Elm"];

  it("counts text that starts with a house number as data over a column of addresses (fail closed)", () => {
    for (const h of ["12 Baker", "221B Baker", "10 Largest Orders", "4 Wheel Drive"]) {
      const d = sheet(h, ADDRESS_COLUMN);
      expect(firstRowLooksLikeData(d), h).toBe(true);
      expect(headerCellLooksLikeData(d, 2), h).toBe(true);
      // A private column's header cell is scanned like its values.
      expect(headerCellIsIdentifier(d, 2), h).toBe(true);
      expect(cellLooksLikeData(h, h, "String", cellsBelow(d, 2)), h).toBe(true);
    }
    // Over a column of text that starts with house numbers but has no street-type word.
    expect(firstRowLooksLikeData(sheet("12 Baker", ["3 High", "77 Queen", "5 Mill"]))).toBe(true);
    // With nothing below to compare with, it counts, as ID-shaped text does.
    const alone = range([["Maria Lopez", "12 Baker", "London"]]);
    expect(firstRowLooksLikeData(alone)).toBe(true);
    expect(headerCellIsIdentifier(alone, 1)).toBe(true);
    expect(firstRowLooksLikeData(sheet("10 Largest Orders", [null, null, null]))).toBe(true);
    // A year or a span of time is never a house number.
    expect(firstRowLooksLikeData(sheet("2024 Sales Plan", ADDRESS_COLUMN))).toBe(false);
    expect(firstRowLooksLikeData(sheet("12 Month Total", ADDRESS_COLUMN))).toBe(false);
    // "Address1" over addresses is still a header.
    expect(firstRowLooksLikeData(sheet("Address1", ADDRESS_COLUMN))).toBe(false);
  });

  it("counts a two-part code as data only over a column of postcodes or codes like it", () => {
    for (const h of ["A1 B2", "AB12 CD34"]) {
      expect(cellLooksLikeData(h, h, "String"), h).toBe(false);
      const d = sheet(h, ["C3 D4", "EF56 GH78", "SW1A 1AA", "K1A 0B1"]);
      expect(firstRowLooksLikeData(d), h).toBe(true);
      expect(headerCellIsIdentifier(d, 2), h).toBe(true);
      expect(firstRowLooksLikeData(sheet(h, [1200, 800, 450])), h).toBe(false);
      expect(firstRowLooksLikeData(sheet(h, ADDRESS_COLUMN)), h).toBe(false);
    }
    // A period and a year stay a header even over codes.
    expect(firstRowLooksLikeData(sheet("Q1 2024", ["C3 D4", "EF56 GH78", "SW1A 1AA"]))).toBe(false);
    // Grouped like an IBAN but far too short for one.
    expect(cellLooksLikeData("AB12 CD34", "AB12 CD34", "String")).toBe(false);
    expect(cellLooksLikeData("NO93 8601 1117 947", "NO93 8601 1117 947", "String")).toBe(true);
  });
});
