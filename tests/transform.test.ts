import { describe, expect, it } from "vitest";
import { inferSchema } from "../src/core/schema";
import {
  COLLISION_RE,
  DEFAULT_ROW_CAP,
  MAX_ROW_CAP,
  MONTH_RE,
  RANGE_RE,
  StandInMap,
  TOKEN_RE,
  codeLetters,
  codeWithoutSeparators,
  createStandInMap,
  dateKey,
  eightDigitDates,
  emailLocalPart,
  findDates,
  foldUnit,
  hasTokenCollision,
  isCodeSeparatorCode,
  joinedNames,
  letterSeparatedParts,
  privateDates,
  privateLooseDates,
  privateNumbers,
  privateValues,
  renderMonth,
  renderRange,
  shortNumbers,
  shownNumbers,
  stripDropped,
  substituteUserText,
  textDates,
  tokenKindFor,
  transformRows,
  withoutCountryCode,
} from "../src/core/transform";
import type { CellValue, ColumnInfo, ColumnPolicy, RangeData, RangeSpec, Treatment } from "../src/core/types";

interface Opts {
  sheet?: string;
  formats?: Record<number, string>;
  /** Display text overrides, by [row][column]. */
  text?: Record<number, Record<number, string>>;
}

function range(rows: CellValue[][], opts: Opts = {}): RangeData {
  return {
    sheet: opts.sheet ?? "Orders",
    address: `A1:${String.fromCharCode(64 + (rows[0]?.length ?? 1))}${rows.length}`,
    rowIndex: 0,
    columnIndex: 0,
    values: rows,
    text: rows.map((r, i) =>
      r.map((v, j) => opts.text?.[i]?.[j] ?? (v === null ? "" : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v))),
    ),
    formulas: rows.map((r) => [...r]),
    valueTypes: rows.map((r) =>
      r.map((v) =>
        v === null
          ? "Empty"
          : typeof v === "number"
            ? "Double"
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

/** A spec with policies given per letter; unspecified columns are not private. */
function spec(
  data: RangeData,
  policies: Record<string, { private: boolean; treatment: Treatment; alias?: string }> = {},
  hasHeaders = true,
): RangeSpec {
  const columns = inferSchema(data, hasHeaders);
  return {
    data,
    hasHeaders,
    columns,
    policies: columns.map((c): ColumnPolicy => ({
      letter: c.letter,
      private: policies[c.letter]?.private ?? false,
      treatment: policies[c.letter]?.treatment ?? "as_is",
      ...(policies[c.letter]?.alias ? { alias: policies[c.letter]!.alias } : {}),
    })),
  };
}

const col = (header: string, type: ColumnInfo["type"] = "text"): ColumnInfo => ({
  letter: "A",
  header,
  type,
  stats: { blank: 0, distinct: 1 },
});

const ORDERS: CellValue[][] = [
  ["Customer Name", "Region", "Amount", "Order date", "Paid", "Email"],
  ["Maria Lopez", "East", 83500, 46085, true, "maria@example.test"],
  ["Ben Ode", "West", 120.5, 46100, false, "ben@example.test"],
  ["Maria Lopez", "East", -4, 46115, true, "maria@example.test"],
  [null, "North", 0, null, null, null],
];
const ORDERS_FMT = { formats: { 3: "m/d/yyyy" }, text: { 1: { 3: "3/4/2026" }, 2: { 3: "3/19/2026" }, 3: { 3: "4/3/2026" } } };

describe("StandInMap", () => {
  it("T-T1: the same value twice in a session gets the same token; a new session gets a fresh map", () => {
    const data = range(ORDERS, ORDERS_FMT);
    const s = spec(data, { A: { private: true, treatment: "stand_in" } });
    const map = createStandInMap([data]);
    const { rows } = transformRows(s, map);
    expect(rows[0]![0]).toBe("PERSON_001");
    expect(rows[1]![0]).toBe("PERSON_002");
    expect(rows[2]![0]).toBe("PERSON_001");
    expect(map.tokenFor("Maria Lopez", "PERSON")).toBe("PERSON_001");
    expect(map.valueOf("PERSON_001")).toBe("Maria Lopez");
    expect(map.isToken("PERSON_001")).toBe(true);
    expect(map.isToken("PERSON_003")).toBe(false);
    expect(map.size).toBe(2);

    const next = createStandInMap([data]);
    expect(next.size).toBe(0);
    expect(next.tokenFor("Ben Ode", "PERSON")).toBe("PERSON_001");
    expect(next.valueOf("PERSON_002")).toBeUndefined();
  });

  it("counts per kind and keeps a value's first kind", () => {
    const map = new StandInMap();
    expect(map.tokenFor("a@x.test", "EMAIL")).toBe("EMAIL_001");
    expect(map.tokenFor("Ana", "PERSON")).toBe("PERSON_001");
    expect(map.tokenFor("X-1", "ID")).toBe("ID_001");
    expect(map.tokenFor("blue", "TEXT")).toBe("TEXT_001");
    expect(map.tokenFor("Ana", "TEXT")).toBe("PERSON_001");
    expect(map.tokenFor("Ben", "PERSON")).toBe("PERSON_002");
  });

  it("pads to at least 3 digits and keeps counting past 999", () => {
    const map = new StandInMap();
    let last = "";
    for (let i = 1; i <= 1001; i++) last = map.tokenFor(`v${i}`, "PERSON");
    expect(map.tokenFor("v9", "PERSON")).toBe("PERSON_009");
    expect(map.tokenFor("v999", "PERSON")).toBe("PERSON_999");
    expect(map.tokenFor("v1000", "PERSON")).toBe("PERSON_1000");
    expect(last).toBe("PERSON_1001");
    const re = new RegExp(TOKEN_RE.source);
    expect(re.test("PERSON_1000")).toBe(true);
    expect(map.isToken("PERSON_1000")).toBe(true);
  });

  it("T-T2: a selection that already contains PERSON_001 gets NYMFORM_-prefixed tokens", () => {
    const data = range([["Customer Name", "Note"], ["Maria Lopez", "see PERSON_001"], ["Ben Ode", "ok"]]);
    expect(hasTokenCollision([data])).toBe(true);
    const map = createStandInMap([data]);
    expect(map.prefix).toBe("NYMFORM_");
    const { rows } = transformRows(spec(data, { A: { private: true, treatment: "stand_in" } }), map);
    expect(rows[0]![0]).toBe("NYMFORM_PERSON_001");
    expect(rows[1]![0]).toBe("NYMFORM_PERSON_002");
    expect(map.isToken("PERSON_001")).toBe(false);

    const clean = range([["Customer Name"], ["Maria Lopez"]]);
    expect(hasTokenCollision([clean])).toBe(false);
    expect(createStandInMap([clean]).prefix).toBe("");
    // A context range counts too.
    expect(createStandInMap([clean, data]).prefix).toBe("NYMFORM_");
  });

  it("never allocates a token that already appears in the workbook", () => {
    const data = range([["Code"], ["PERSON_001"], ["NYMFORM_PERSON_001"], ["x"]]);
    const map = createStandInMap([data]);
    expect(map.prefix).toBe("NYMFORM_");
    expect(map.tokenFor("Maria Lopez", "PERSON")).toBe("NYMFORM_PERSON_002");
    expect(COLLISION_RE.test("NYMFORM_PERSON_001")).toBe(false);
  });

  it("does not serialize its contents", () => {
    const map = new StandInMap();
    map.tokenFor("Quorbel Vantrisk", "PERSON");
    const json = JSON.stringify({ map });
    expect(json).not.toContain("Quorbel");
    expect(json).not.toContain("PERSON_001");
  });
});

describe("tokenKindFor", () => {
  it("picks the kind from the header", () => {
    expect(tokenKindFor(col("Customer Name"))).toBe("PERSON");
    expect(tokenKindFor(col("Client name"))).toBe("PERSON");
    expect(tokenKindFor(col("Customer"))).toBe("PERSON");
    expect(tokenKindFor(col("Instructor"))).toBe("PERSON");
    expect(tokenKindFor(col("E-mail"))).toBe("EMAIL");
    expect(tokenKindFor(col("Employee ID"))).toBe("ID");
    expect(tokenKindFor(col("Phone"))).toBe("ID");
    expect(tokenKindFor(col("Student Number"))).toBe("ID");
    expect(tokenKindFor(col("Student"))).toBe("PERSON");
    expect(tokenKindFor(col("Product name"))).toBe("TEXT");
    expect(tokenKindFor(col("Region"))).toBe("TEXT");
  });

  it("uses display texts when the header doesn't decide", () => {
    expect(tokenKindFor(col("Contact"), ["Maria Lopez", "Ben Ode"])).toBe("PERSON");
    expect(tokenKindFor(col("Contact"), ["a@x.test", "b@y.test", ""])).toBe("EMAIL");
    expect(tokenKindFor(col("Reach"), ["a@x.test", "b@y.test", "n/a"])).toBe("EMAIL");
    expect(tokenKindFor(col("Ref"), ["AB12345", "AB12346", "AB12347"])).toBe("ID");
    expect(tokenKindFor(col("Amount", "number"), ["83500", "12345"])).toBe("TEXT");
    expect(tokenKindFor(col("Status"), ["open", "closed"])).toBe("TEXT");
  });
});

describe("renderRange", () => {
  it.each([
    [83500, "80000-89999"],
    [0, "0"],
    [-0, "0"],
    [-83500, "-89999--80000"],
    [5, "0-9"],
    [0.5, "0-9"],
    [-5, "-9-0"],
    [1e6, "1000000-1999999"],
    [10, "10-19"],
    [99.5, "90-99"],
    [100, "100-199"],
    [999.999, "900-999"],
    [1e21, `1${"0".repeat(21)}-1${"9".repeat(21)}`],
  ])("renders %d as %s", (x, expected) => {
    expect(renderRange(x)).toBe(expected);
    expect(RANGE_RE.test(renderRange(x))).toBe(true);
  });

  it("rejects numbers it can't render", () => {
    expect(() => renderRange(Number.NaN)).toThrow();
    expect(() => renderRange(Number.POSITIVE_INFINITY)).toThrow();
  });
});

describe("renderMonth", () => {
  it("renders Excel date serials", () => {
    expect(renderMonth(46085, "3/4/2026")).toBe("2026-03");
    expect(renderMonth(46085.75, "3/4/2026 18:00")).toBe("2026-03");
    expect(renderMonth(1, "1/1/1900")).toBe("1900-01");
    expect(renderMonth(32, "2/1/1900")).toBe("1900-02");
    expect(renderMonth(60, "2/29/1900")).toBe("1900-02");
    expect(renderMonth(61, "3/1/1900")).toBe("1900-03");
    expect(renderMonth(2958465, "12/31/9999")).toBe("9999-12");
    expect(renderMonth(2958466, "")).toBeNull();
    expect(renderMonth(-1, "")).toBeNull();
  });

  it("renders date text and refuses anything else", () => {
    expect(renderMonth("2026-03-04", "2026-03-04")).toBe("2026-03");
    expect(renderMonth("2026-03-04T10:00:00Z", "")).toBe("2026-03");
    expect(renderMonth("March 4, 2026", "")).toBe("2026-03");
    expect(renderMonth("4 Mar 2026", "")).toBe("2026-03");
    expect(renderMonth("13/04/2026", "")).toBe("2026-04");
    expect(renderMonth("04/13/2026", "")).toBe("2026-04");
    expect(renderMonth("03/04/2026", "")).toBeNull(); // day and month can't be told apart
    expect(renderMonth("2026-13-01", "")).toBeNull();
    expect(renderMonth("2026-02-30", "")).toBeNull();
    expect(renderMonth("Maria Lopez", "")).toBeNull();
    expect(renderMonth(true, "TRUE")).toBeNull();
    expect(renderMonth(null, "")).toBeNull();
    for (const m of ["2026-03", "1900-01", "9999-12"]) expect(MONTH_RE.test(m)).toBe(true);
  });
});

describe("transformRows", () => {
  it("renders private cells by treatment and passes other cells through", () => {
    const data = range(ORDERS, ORDERS_FMT);
    const s = spec(data, {
      A: { private: true, treatment: "stand_in" },
      C: { private: true, treatment: "range" },
      D: { private: true, treatment: "month" },
      F: { private: true, treatment: "exclude" },
    });
    const out = transformRows(s, new StandInMap());
    expect(out.rowsTotal).toBe(4);
    expect(out.rowsSent).toBe(4);
    expect(out.rows).toEqual([
      ["PERSON_001", "East", "80000-89999", "2026-03", true],
      ["PERSON_002", "West", "100-199", "2026-03", false],
      ["PERSON_001", "East", "-9-0", "2026-04", true],
      [null, "North", "0", null, null],
    ]);
  });

  it("omits excluded columns from rows, private or not", () => {
    const data = range(ORDERS, ORDERS_FMT);
    const s = spec(data, { B: { private: false, treatment: "exclude" }, F: { private: true, treatment: "exclude" } });
    const out = transformRows(s, new StandInMap());
    expect(out.rows.every((r) => r.length === 4)).toBe(true);
    expect(out.rows[0]).toEqual(["Maria Lopez", 83500, "3/4/2026", true]);
    expect(JSON.stringify(out)).not.toContain("example.test");
  });

  it("sends numbers and booleans raw, text and dates as display text, errors as display text", () => {
    const data = range(
      [
        ["Name", "Amount", "When", "Ok", "Calc"],
        ["Ana", 1234.5, 46085, true, "#DIV/0!"],
      ],
      { formats: { 2: "yyyy-mm-dd" }, text: { 1: { 1: "$1,234.50", 2: "2026-03-04" } } },
    );
    const out = transformRows(spec(data), new StandInMap());
    expect(out.rows[0]).toEqual(["Ana", 1234.5, "2026-03-04", true, "#DIV/0!"]);
  });

  it("falls back to a token for cells a range or month can't render", () => {
    const data = range([
      ["Score", "Seen"],
      ["n/a", "sometime"],
      [7, "2026-05-01"],
      [true, 42],
    ]);
    const s = spec(data, { A: { private: true, treatment: "range" }, B: { private: true, treatment: "month" } });
    const map = new StandInMap();
    const out = transformRows(s, map);
    expect(map.valueOf(out.rows[0]![0] as string)).toBe("n/a");
    expect(map.valueOf(out.rows[0]![1] as string)).toBe("sometime");
    expect(out.rows[1]).toEqual(["0-9", "2026-05"]);
    expect(map.valueOf(out.rows[2]![0] as string)).toBe("TRUE");
    // 42 in a text column without a date format is not a date serial.
    expect(map.valueOf(out.rows[2]![1] as string)).toBe("42");
  });

  it("caps rows at 50 by default and at most 200", () => {
    const rows: CellValue[][] = [["Customer Name", "N"]];
    for (let i = 0; i < 300; i++) rows.push([`Person ${i}`, i]);
    const data = range(rows);
    const s = spec(data, { A: { private: true, treatment: "stand_in" } });
    expect(DEFAULT_ROW_CAP).toBe(50);
    expect(MAX_ROW_CAP).toBe(200);
    const def = transformRows(s, new StandInMap());
    expect(def).toMatchObject({ rowsSent: 50, rowsTotal: 300 });
    expect(def.rows).toHaveLength(50);
    expect(def.rows[49]![1]).toBe(49);
    expect(transformRows(s, new StandInMap(), 500).rowsSent).toBe(200);
    expect(transformRows(s, new StandInMap(), 120).rows).toHaveLength(120);
    expect(transformRows(s, new StandInMap(), 0).rowsSent).toBe(1);
    const small = transformRows(spec(range(rows.slice(0, 4))), new StandInMap(), 200);
    expect(small).toMatchObject({ rowsSent: 3, rowsTotal: 3 });
  });

  it("treats every row as data without headers", () => {
    const data = range([["Maria Lopez"], ["Ben Ode"]]);
    const out = transformRows(spec(data, { A: { private: true, treatment: "stand_in" } }, false), new StandInMap());
    expect(out.rows).toEqual([["TEXT_001"], ["TEXT_002"]]);
    expect(out.rowsTotal).toBe(2);
  });

  it.each<Treatment>(["as_is", "stand_in", "range", "month", "exclude"])(
    "never passes a private cell through raw (treatment %s)",
    (treatment) => {
      const rows: CellValue[][] = [["Customer Name", "Amount", "Visit", "Flag", "Code"]];
      for (let i = 0; i < 30; i++) {
        rows.push([
          i % 7 === 0 ? null : `Person ${i % 11}`,
          i % 5 === 0 ? `note ${i}` : 1000 * i + 0.5,
          i % 4 === 0 ? `2026-0${(i % 9) + 1}-01` : 46000 + i,
          i % 2 === 0,
          i % 3 === 0 ? "#N/A" : `C-${i}`,
        ]);
      }
      const data = range(rows, { formats: { 2: "m/d/yyyy" } });
      const all = { private: true, treatment };
      const s = spec(data, { A: all, B: all, C: all, D: all, E: all });
      const map = new StandInMap();
      const out = transformRows(s, map, 200);
      const originals = new Set<string>();
      for (const r of rows.slice(1)) for (const v of r) if (v !== null) originals.add(String(v));
      for (const t of data.text.slice(1)) for (const v of t) if (v !== "") originals.add(v);
      expect(out.rows.length).toBe(30);
      for (const row of out.rows) {
        expect(row.length).toBe(treatment === "exclude" ? 0 : 5);
        for (const cell of row) {
          if (cell === null) continue;
          expect(typeof cell).toBe("string");
          const ok =
            map.isToken(cell as string) ||
            (treatment === "range" && RANGE_RE.test(cell as string)) ||
            (treatment === "month" && MONTH_RE.test(cell as string));
          expect(ok).toBe(true);
          expect(originals.has(cell as string)).toBe(false);
        }
      }
    },
  );
});

describe("privateValues", () => {
  it("covers every row, not just the rows sent, with display text and raw value", () => {
    const rows: CellValue[][] = [["Customer Name", "Salary", "Region"]];
    for (let i = 0; i < 300; i++) rows.push([`Person ${i}`, 50000 + i, "East"]);
    rows.push([null, null, "West"]);
    const text: Record<number, Record<number, string>> = { 1: { 1: "$50,000.00" } };
    const data = range(rows, { text });
    const s = spec(data, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "range" } });
    transformRows(s, new StandInMap());
    const values = privateValues(s);
    expect([...values.keys()]).toEqual(["A", "B"]);
    expect(values.get("A")).toHaveLength(300);
    expect(values.get("A")).toContain("Person 299");
    expect(values.get("A")).not.toContain("Customer Name");
    expect(values.get("B")).toContain("$50,000.00");
    expect(values.get("B")).toContain("50000");
    expect(values.get("B")).toContain("50299");
    expect(values.get("B")).not.toContain("");
    expect(values.has("C")).toBe(false);
  });

  it("includes excluded private columns and both forms of dates", () => {
    const data = range(ORDERS, ORDERS_FMT);
    const s = spec(data, { D: { private: true, treatment: "month" }, F: { private: true, treatment: "exclude" } });
    const values = privateValues(s);
    expect(values.get("D")).toEqual(expect.arrayContaining(["3/4/2026", "46085", "4/3/2026", "46115"]));
    expect(values.get("F")).toEqual(["maria@example.test", "ben@example.test"]);
  });
});

describe("substituteUserText", () => {
  const data = () => range(ORDERS, ORDERS_FMT);

  it("T-T3: replaces a private name in the question with its token", () => {
    const d = data();
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    const map = createStandInMap([d]);
    transformRows(s, map);
    // Ben Ode was allocated second, so the question uses his existing token, not a new one.
    expect(substituteUserText("total for Maria Lopez", [s], map)).toBe("total for PERSON_001");
    expect(substituteUserText("orders by Ben Ode?", [s], map)).toBe("orders by PERSON_002?");
    expect(map.size).toBe(2);
  });

  it("allocates a token when the value wasn't sent in rows", () => {
    const d = data();
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    const map = createStandInMap([d]);
    expect(substituteUserText("total for Ben Ode", [s], map)).toBe("total for PERSON_001");
    expect(map.valueOf("PERSON_001")).toBe("Ben Ode");
  });

  it("matches values of 4+ characters case-insensitively and across whitespace", () => {
    const d = data();
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("total for MARIA  lopez.", [s], map)).toBe("total for PERSON_001.");
    expect(substituteUserText("total for Maria\nLopez", [s], map)).toBe("total for PERSON_001");
  });

  it("T-T7: a short private value is matched case-exactly, so the word 'an' is untouched", () => {
    const d = range([["Initials"], ["An"], ["Bo"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("an apple for An and AN", [s], map)).toBe("an apple for TEXT_001 and AN");
    expect(map.valueOf("TEXT_001")).toBe("An");
  });

  it("replaces whole words only", () => {
    const d = range([["Word"], ["Lopez"], ["5"], ["Ann"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("Lopezes and Ann's 5.5 or 15 or 1,5", [s], map)).toBe("Lopezes and TEXT_001's 5.5 or 15 or 1,5");
    expect(substituteUserText("top 5 for Lopez", [s], map)).toBe("top TEXT_002 for TEXT_003");
  });

  it("replaces the longest match first, without overlaps", () => {
    const d = range([["Customer Name", "First"], ["Maria Lopez", "Maria"], ["Ben Ode", "Ben"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    const out = substituteUserText("Maria Lopez and Maria", [s], map);
    expect(out).toBe("PERSON_001 and PERSON_002");
    expect(map.valueOf("PERSON_001")).toBe("Maria Lopez");
    expect(map.valueOf("PERSON_002")).toBe("Maria");
  });

  it("replaces raw and display forms of private numbers and dates", () => {
    const d = range(
      [
        ["Salary", "Hired"],
        [83500, 46085],
      ],
      { formats: { 1: "m/d/yyyy" }, text: { 1: { 0: "$83,500", 1: "3/4/2026" } } },
    );
    const s = spec(d, { A: { private: true, treatment: "range" }, B: { private: true, treatment: "month" } });
    const map = new StandInMap();
    const out = substituteUserText("above 83500 or $83,500, hired 3/4/2026", [s], map);
    expect(out).not.toMatch(/83500|83,500|3\/4\/2026/);
    const tokens = out.match(new RegExp(TOKEN_RE.source, "g"));
    expect(tokens).toHaveLength(3);
    expect(tokens![0]).toBe(tokens![1]);
  });

  it("replaces the original header of an aliased column with its alias", () => {
    const d = range([["Region", "Diagnosis"], ["East", "Flu"], ["West", "Cold"]]);
    const s = spec(d, { B: { private: true, treatment: "stand_in", alias: "Field B" } });
    const map = new StandInMap();
    expect(substituteUserText("count by diagnosis and Region", [s], map)).toBe("count by Field B and Region");
    expect(substituteUserText("count Flu cases", [s], map)).toBe("count TEXT_001 cases");
    // Short headers match in any case too: they are sent only under the alias.
    const short = range([["Region", "Sex"], ["East", "F"], ["West", "M"]]);
    const s2 = spec(short, { B: { private: true, treatment: "stand_in", alias: "Field B" } });
    expect(substituteUserText("count by sex", [s2], new StandInMap())).toBe("count by Field B");
  });

  it("does not rewrite inside already-substituted tokens", () => {
    const d = range([["Customer Name", "Kind"], ["Maria Lopez", "Person"], ["PERSON_001", "001"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "stand_in" } });
    // A map made without the collision guard, so a sheet value can look exactly like a token.
    const map = new StandInMap();
    const once = substituteUserText("total for Maria Lopez", [s], map);
    expect(once).toBe("total for PERSON_001");
    expect(substituteUserText(once, [s], map)).toBe(once);
    expect(substituteUserText(`${once} per Person`, [s], map)).toBe("total for PERSON_001 per TEXT_001");
  });

  it("uses every spec, including context ranges", () => {
    const main = range([["Region"], ["East"]]);
    const lookup = range([["Manager"], ["Quorbel Vantrisk"]], { sheet: "Staff" });
    const specs = [spec(main), spec(lookup, { A: { private: true, treatment: "stand_in" } })];
    const out = substituteUserText("regions run by Quorbel Vantrisk", specs, new StandInMap());
    expect(out).toBe("regions run by PERSON_001");
  });

  it("leaves text alone when nothing private matches", () => {
    const d = data();
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("Total Amount per Region", [s], map)).toBe("Total Amount per Region");
    expect(substituteUserText("", [s], map)).toBe("");
    expect(map.size).toBe(0);
  });
});

describe("substituteUserText: other spellings of a private value", () => {
  it("folds Greek final sigma, so a name stored in capitals is replaced when typed normally", () => {
    const d = range([["Name"], ["ΓΙΏΡΓΟΣ ΠΑΠΠΆΣ"], ["ΝΊΚΟΣ ΠΑΠΑΔΌΠΟΥΛΟΣ"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("Τι μισθό παίρνει ο Γιώργος Παππάς;", [s], map)).toBe("Τι μισθό παίρνει ο PERSON_001;");
    expect(substituteUserText("Salary of Νίκος Παπαδόπουλος", [s], map)).toBe("Salary of PERSON_002");
    expect(map.valueOf("PERSON_001")).toBe("ΓΙΏΡΓΟΣ ΠΑΠΠΆΣ");
  });

  it("matches either form of accents and either apostrophe; the question comes back NFC-normalized", () => {
    const d = range([["Name", "Last name"], ["José Muñoz".normalize("NFD"), "O’Sullivan"], ["Ana Ruiz", "D'Alessandro"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("total for José Muñoz".normalize("NFC"), [s], map)).toBe("total for PERSON_001");
    expect(substituteUserText("total for José Muñoz".normalize("NFD"), [s], map)).toBe("total for PERSON_001");
    // Apostrophes outside a match stay as typed.
    expect(substituteUserText("Sean’s total for O'Sullivan and D’Alessandro", [s], map)).toBe(
      "Sean’s total for PERSON_002 and PERSON_003",
    );
    expect(substituteUserText("the café".normalize("NFD"), [s], map)).toBe("the café".normalize("NFC"));
  });

  it("replaces a name written without the space it is stored with, next to Japanese text", () => {
    const d = range([["Name", "Dept"], ["山田 太郎", "営業"], ["佐藤　花子", "営業"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("山田太郎さんの給与は?", [s], map)).toBe("PERSON_001さんの給与は?");
    expect(substituteUserText("佐藤花子の部署", [s], map)).toBe("PERSON_002の部署");
    expect(substituteUserText("Total for 山田 太郎 and 佐藤 花子", [s], map)).toBe("Total for PERSON_001 and PERSON_002");
  });

  it("ends a word at a Chinese, Japanese or Korean character, but not inside a Latin word", () => {
    const d = range([["Name"], ["Lopez"], ["王小明明"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("Lopezさんの売上 and Lopezes", [s], map)).toBe("PERSON_001さんの売上 and Lopezes");
    expect(substituteUserText("王小明明的销售额", [s], map)).toBe("PERSON_002的销售额");
  });

  it("replaces a phone number typed with other punctuation", () => {
    const d = range([["Phone"], ["831-555-0199"], ["4155550123"], ["212.555.0147"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("orders from (831) 555-0199?", [s], map)).toBe("orders from ID_001?");
    expect(substituteUserText("call 415-555-0123 or +1 212 555 0147", [s], map)).toBe("call ID_002 or +1 ID_003");
    // Not inside a longer number.
    expect(substituteUserText("ref 98315550199", [s], map)).toBe("ref 98315550199");
  });

  it("replaces an amount typed with thousands separators", () => {
    const d = range([["Salary"], [83500], [41250.5]], { text: { 1: { 0: "$83,500" }, 2: { 0: "$41,250.50" } } });
    const s = spec(d, { A: { private: true, treatment: "range" } });
    const map = new StandInMap();
    expect(substituteUserText("between 83,500 and 83.500 or 83 500 or $83,500.00", [s], map)).toBe(
      "between TEXT_001 and TEXT_001 or TEXT_001 or $TEXT_001",
    );
    expect(substituteUserText("not 41,250.5 or 41.250,5, and not 183,500 or 83,5001", [s], map)).toBe(
      "not TEXT_002 or TEXT_002, and not 183,500 or 83,5001",
    );
  });
});

// ---------------------------------------------------------------------------------------------
// Second review: the numbers a user sees, more spellings

describe("shownNumbers", () => {
  it("reads the number the user sees, and the raw value rounded to the displayed decimals", () => {
    expect(shownNumbers(83500.4, "$83,500", true)).toEqual(["83500", "83500.4"]);
    expect(shownNumbers(0.0525, "5.25%", true)).toEqual(["5.25", "0.0525"]);
    expect(shownNumbers(1234.2000000000003, "1,234.20", true).sort()).toEqual(["1234", "1234.2", "1234.20"]);
    expect(shownNumbers(-83500, "(83,500)", true)).toEqual(["83500"]);
    expect(shownNumbers(83500, "83.500 €", true)).toEqual(["83500"]);
    expect(shownNumbers(835000, "₹8,35,000", true)).toEqual(["835000"]);
    expect(shownNumbers(1.234, "1.234", true)).toEqual(["1.234"]);
    expect(shownNumbers(83500, "", true)).toEqual(["83500"]);
    // A display that isn't the value rounded (a custom format in thousands) is not read, but the
    // stored value still counts.
    expect(shownNumbers(83500, "84", true)).toEqual(["83500"]);
    // Shorter than 4 characters: left to the other variants.
    expect(shownNumbers(0.031, "3.10%", true)).toEqual(["3.10", "0.031"]);
    expect(shownNumbers(12, "12", true)).toEqual([]);
  });

  it("reads text cells that are numbers or amounts (third review: digits stored as text count)", () => {
    expect(shownNumbers("$83,500", "$83,500", false)).toEqual(["83500"]);
    expect(shownNumbers("12.5%", "12.5%", false)).toEqual(["12.5"]);
    expect(shownNumbers("1,250", "1,250", false)).toEqual(["1250"]);
    expect(shownNumbers("83500", "83500", false)).toEqual(["83500"]);
    expect(shownNumbers("1234.50", "1234.50", false).sort()).toEqual(["1234.5", "1234.50"]);
    expect(shownNumbers("02134", "02134", false)).toEqual(["2134"]);
    for (const t of ["EMP-00123", "3/4/2026", "555-0199", "Maria", "TRUE", "123"]) expect(shownNumbers(t, t, false), t).toEqual([]);
  });
});

describe("privateNumbers", () => {
  it("covers every private column but dates (third review: ZIP codes, IDs and phone numbers too)", () => {
    const d = range(
      [
        ["Salary", "ZIP", "Hired", "Employee ID", "Phone", "Paid"],
        [83500.4, 93496, 46085, 1000123, 8315550199, "$91,250"],
        [41250, 2134, 46100, 1000124, 4155550123, "$7,000"],
      ],
      { formats: { 2: "m/d/yyyy" }, text: { 1: { 0: "$83,500", 2: "3/4/2026" }, 2: { 0: "$41,250", 2: "3/19/2026" } } },
    );
    const s = spec(d, Object.fromEntries(["A", "B", "C", "D", "E", "F"].map((l) => [l, { private: true, treatment: "stand_in" as const }])));
    const numbers = privateNumbers(s);
    expect(numbers.get("A")).toEqual(["83500", "83500.4", "41250"]);
    expect(numbers.get("F")).toEqual(["91250", "7000"]);
    expect(numbers.get("B")).toEqual(["93496", "2134"]);
    expect(numbers.get("D")).toEqual(["1000123", "1000124"]);
    expect(numbers.get("E")).toEqual(["8315550199", "4155550123"]);
    expect(numbers.get("C")).toEqual([]);
  });
});

describe("substituteUserText: second review", () => {
  it("replaces the number the user sees in a cell, in other formats", () => {
    const d = range(
      [
        ["Employee", "Salary", "Raise", "Paid"],
        ["Maria Lopez", 83500.4, 0.0525, "$91,250"],
        ["Ana Ruiz", 1234.2000000000003, 0.031, "$7,000"],
        ["Kenji Watanabe", 835000, 0.1, "$12,50,000"],
      ],
      { text: { 1: { 1: "$83,500", 2: "5.25%" }, 2: { 1: "1,234.20", 2: "3.10%" }, 3: { 1: "₹8,35,000", 2: "10.00%" } } },
    );
    const s = spec(d, { B: { private: true, treatment: "range" }, C: { private: true, treatment: "range" }, D: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    const out = substituteUserText(
      "who earns 83,500 or 83500, paid 1234.2 or 1.234,20, got 5.25 or a 5,25 raise, 91,250, 1250000, 8,35,000",
      [s],
      map,
    );
    expect(out).not.toMatch(/83,?500|1234|1\.234|5[.,]25|91,250|1250000|8,35,000/);
    // Not inside other numbers, and "3.1" is shorter than 4 characters.
    expect(substituteUserText("rows 183,500 or 83,5001 or 5.255 and 3.1", [s], map)).toBe("rows 183,500 or 83,5001 or 5.255 and 3.1");
  });

  it("folds Turkish İ and ı in the question and the values", () => {
    const d = range([["Name", "Last name"], ["İbrahim", "Yıldız"], ["İsmail Ak", "Kaya"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("total for ibrahim, IBRAHIM and İBRAHİM", [s], map)).toBe("total for PERSON_001, PERSON_001 and PERSON_001");
    expect(substituteUserText("total for ISMAIL AK and ismail ak", [s], map)).toBe("total for PERSON_002 and PERSON_002");
    expect(substituteUserText("orders by YILDIZ or Yildiz", [s], map)).toBe("orders by PERSON_003 or PERSON_003");
    // A dotted i typed with its combining dot is one letter; text after it keeps its place.
    expect(substituteUserText("i̇brahim's total, then Kaya", [s], map)).toBe("PERSON_001's total, then PERSON_004");
  });

  it("folds apostrophe look-alikes, full-width digits and the minus sign", () => {
    const d = range([["Last name", "Phone"], ["O'Neill", "090-1234-5678"], ["Murphy", "080-2345-6789"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    for (const typed of ["O´Neill", "O`Neill", "O＇Neill", "O’Neill"]) {
      expect(substituteUserText(`total for ${typed}`, [s], map), typed).toBe("total for PERSON_001");
    }
    expect(substituteUserText("０９０−１２３４−５６７８の注文", [s], map)).toBe("ID_001の注文");
    expect(substituteUserText("０８０－２３４５－６７８９ and 080−2345−6789", [s], map)).toBe("ID_002 and ID_002");
    // The auditor joins digits across the same full-width punctuation.
    expect(substituteUserText("電話 ０９０（１２３４）５６７８", [s], map)).toBe("電話 ID_001");
  });

  it("replaces a Chinese, Japanese or Korean value typed with spaces between its characters", () => {
    const d = range([["Name"], ["山田太郎"], ["김민준"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("山田 太郎さん and 山 田 太 郎", [s], map)).toBe("PERSON_001さん and PERSON_001");
    expect(substituteUserText("김 민준", [s], map)).toBe("PERSON_002");
    // Latin letters next to each other still need to be typed together.
    const latin = spec(range([["Name"], ["Lopez"]]), { A: { private: true, treatment: "stand_in" } });
    expect(substituteUserText("L o p e z", [latin], new StandInMap())).toBe("L o p e z");
  });
});

// ---------------------------------------------------------------------------------------------
// Third review: regressions and gaps

describe("substituteUserText: third review", () => {
  const TOKEN = /^(?:NYMFORM_)?(?:PERSON|EMAIL|ID|TEXT)_\d{3,}$/;

  it("replaces an amount stored as text in other formats, whatever the column's token kind", () => {
    const d = range([["Employee", "Salary"], ["Maria Lopez", "83500"], ["Ana Ruiz", "125000"], ["Kenji Watanabe", "1234.50"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "stand_in" } });
    expect(tokenKindFor(s.columns[1]!, ["83500", "125000", "1234.50"])).toBe("ID");
    const map = new StandInMap();
    const out = substituteUserText("who earns 83,500, $125,000, 1,25,000, 125.000,00 or 1,234.50?", [s], map);
    expect(out).not.toMatch(/83|125|1,?234/);
    expect(out.match(new RegExp(TOKEN_RE.source, "g"))).toHaveLength(5);
  });

  it("replaces a phone number stored with its trunk 0 typed in international form, and the reverse", () => {
    const d = range([["Name", "Phone"], ["Maria Lopez", "020 7946 0958"], ["Ana Ruiz", "+18315550199"]]);
    const s = spec(d, { B: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("calls from +44 20 7946 0958, +442079460958 or 0044 20 7946 0958", [s], map)).toBe("calls from ID_001, ID_001 or ID_001");
    expect(substituteUserText("and (831) 555-0199", [s], map)).toBe("and ID_002");
    expect(map.valueOf("ID_001")).toBe("020 7946 0958");
    // Parts of a number are left alone: the auditor blocks them in the question.
    expect(substituteUserText("the one ending 0958", [s], map)).toBe("the one ending 0958");
  });

  it("folds accents, ß and full-width letters, and replaces a name joined with _", () => {
    const d = range([["Name"], ["José Muñoz"], ["François Müller"], ["ΝΊΚΟΣ ΠΑΠΑΔΌΠΟΥΛΟΣ"], ["Gerhard Weiß"], ["Maria Lopez"], ["Łukasz Wójcik"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    const typed: [string, string][] = [
      ["Jose Munoz", "José Muñoz"],
      ["JOSE MUNOZ", "José Muñoz"],
      ["Francois Muller", "François Müller"],
      ["Νικος Παπαδοπουλος", "ΝΊΚΟΣ ΠΑΠΑΔΌΠΟΥΛΟΣ"],
      ["Gerhard Weiss", "Gerhard Weiß"],
      ["Ｍａｒｉａ Ｌｏｐｅｚ", "Maria Lopez"],
      ["maria_lopez", "Maria Lopez"],
      ["Lukasz Wojcik", "Łukasz Wójcik"],
    ];
    for (const [question, value] of typed) {
      const out = substituteUserText(question, [s], map);
      expect(out, question).toMatch(TOKEN);
      expect(map.valueOf(out), question).toBe(value);
    }
    expect(substituteUserText("use Bonus_Maria_Lopez", [s], map)).toBe(`use Bonus_${map.peek("Maria Lopez")}`);
    // Still whole words only.
    expect(substituteUserText("mariax_lopez", [s], map)).toBe("mariax_lopez");
  });

  it("folds Arabic-Indic, Persian and Devanagari digits", () => {
    const d = range([["Name", "Phone", "Salary"], ["Maria Lopez", "0501234567", 83500], ["Ana Ruiz", "9876543210", 91250]]);
    const s = spec(d, { B: { private: true, treatment: "stand_in" }, C: { private: true, treatment: "range" } });
    const map = new StandInMap();
    expect(substituteUserText("call ٠٥٠١٢٣٤٥٦٧ or ۰۵۰۱۲۳۴۵۶۷ or ९८७६५४३२१०", [s], map)).toBe("call ID_001 or ID_001 or ID_002");
    expect(substituteUserText("earns ٨٣٬٥٠٠ or ٨٣٥٠٠", [s], map)).toMatch(/^earns TEXT_\d{3} or TEXT_\d{3}$/);
  });

  it("replaces a private date written in other formats, as a whole date only", () => {
    const d = range([["Name", "Born"], ["Maria Lopez", 31121], ["Ana Ruiz", "4/2/1990"]], { formats: { 1: "m/d/yyyy" }, text: { 1: { 1: "3/15/1985" } } });
    const s = spec(d, { B: { private: true, treatment: "month" } });
    const map = new StandInMap();
    for (const q of ["1985-03-15", "15/03/1985", "03/15/1985", "15.03.1985", "15.3.1985", "March 15, 1985", "15 March 1985", "Mar 15 1985", "Mar 15, 1985", "1990-04-02", "2 April 1990", "1990-02-04"]) {
      expect(substituteUserText(`born ${q}?`, [s], map), q).toMatch(/^born (?:NYMFORM_)?[A-Z]+_\d{3}\?$/);
    }
    expect(substituteUserText("born in 1985, in March 1985 or on 1985-03-16", [s], map)).toBe("born in 1985, in March 1985 or on 1985-03-16");
  });
});

describe("dates as people write them (third review)", () => {
  it("reads date text, with both readings when day and month could swap", () => {
    expect(textDates("3/15/1985")).toEqual([{ y: 1985, m: 3, d: 15 }]);
    expect(textDates("4/2/1990")).toEqual([
      { y: 1990, m: 4, d: 2 },
      { y: 1990, m: 2, d: 4 },
    ]);
    expect(textDates("1985-03-15")).toEqual([{ y: 1985, m: 3, d: 15 }]);
    expect(textDates("15.03.1985")).toEqual([{ y: 1985, m: 3, d: 15 }]);
    expect(textDates("Wed, March 4, 2026 10:30 AM")).toEqual([{ y: 2026, m: 3, d: 4 }]);
    expect(textDates("4 Mar 2026")).toEqual([{ y: 2026, m: 3, d: 4 }]);
    for (const t of ["1985", "March 1985", "13/13/2020", "2/30/2024", "Maria", "123-45-6789"]) expect(textDates(t), t).toEqual([]);
  });

  it("privateDates gives the ISO dates of date-formatted numbers and date text in private columns only", () => {
    const d = range(
      [
        ["Born", "Hired", "Amount"],
        [31121, "4/2/1990", 31121],
        [null, "not a date", 5],
      ],
      { formats: { 0: "m/d/yyyy" } },
    );
    const s = spec(d, { A: { private: true, treatment: "month" }, B: { private: true, treatment: "stand_in" }, C: { private: true, treatment: "range" } });
    const dates = privateDates(s);
    expect(dates.get("A")).toEqual(["1985-03-15"]);
    expect(dates.get("B")).toEqual(["1990-04-02", "1990-02-04"]);
    expect(dates.get("C")).toEqual([]);
  });
});

describe("foldUnit (third review)", () => {
  it("folds accents, full-width ASCII, other digits and look-alikes one unit for one, keeping case", () => {
    const fold = (s: string) => [...s].map((c) => String.fromCharCode(foldUnit(c.charCodeAt(0)))).join("");
    expect(fold("José Muñoz ÉÀÇ")).toBe("Jose Munoz EAC");
    expect(fold("ΝΊΚΟΣ ά ς")).toBe("ΝΙΚΟΣ α σ");
    expect(fold("Ｍａｒｉａ（１２）")).toBe("Maria(12)");
    expect(fold("٠١٢٣٤٥٦٧٨٩ ۰۱۲۳ ०१२३ ٫٬")).toBe("0123456789 0123 0123 .,");
    expect(fold("Łódź Øre đ ħ ŧ İı")).toBe("Lodz Ore d h t Ii");
    expect(fold("O’Neill\u00a0−5")).toBe("O'Neill -5");
    // Hangul and kana with voicing marks are left alone.
    expect(fold("김민준 が")).toBe("김민준 が");
  });

  it("finds the national number of a phone stored in international form", () => {
    expect(withoutCountryCode("+18315550199")).toEqual(["8315550199", "315550199", "15550199"]);
    expect(withoutCountryCode("0044 20 7946 0958")).toEqual(["42079460958", "2079460958", "079460958"]);
    expect(withoutCountryCode("(831) 555-0199")).toEqual([]);
  });
});

describe("shownNumbers: the stored value counts whatever the format shows", () => {
  it("reads amounts with currency codes, units and scale letters, and always includes the raw value", async () => {
    const { shownNumbers } = await import("../src/core/transform");
    expect(shownNumbers(83500, "83 500,00 kr", true)).toEqual(expect.arrayContaining(["83500"]));
    expect(shownNumbers(91250, "CHF 91,250", true)).toEqual(expect.arrayContaining(["91250"]));
    expect(shownNumbers(83500, "83.5K", true)).toEqual(expect.arrayContaining(["83500"]));
    expect(shownNumbers(1234.56, "$1,235", true)).toEqual(expect.arrayContaining(["1235", "1234.56"]));
    expect(shownNumbers(null, "83,500 EUR", false)).toEqual(expect.arrayContaining(["83500"]));
  });
});


// ---------------------------------------------------------------------------------------------
// Fourth review: date phrases parsed, not listed; grouped numbers with zero decimals

describe("date phrases in free text (fourth review)", () => {
  const read = (text: string) => findDates(text).map((f) => [text.slice(f.start, f.end), f.dates.map(dateKey).join(" ")]);

  it("reads numeric dates in y-m-d, m-d-y and d-m-y order, with 2- or 4-digit years and every reading", () => {
    expect(read("born 3/15/85")).toEqual([["3/15/85", "1985-03-15 2085-03-15"]]);
    expect(read("born 15.3.85")).toEqual([["15.3.85", "1985-03-15 2085-03-15"]]);
    expect(read("born 4/2/1990")).toEqual([["4/2/1990", "1990-04-02 1990-02-04"]]);
    for (const t of ["1985-03-15", "1985/03/15", "1985/3/15", "1985.03.15", "1985 03 15", "15 03 1985", "3.15.1985", "03-15-1985", "15 / 03 / 1985"]) {
      expect(read(`on ${t}.`), t).toEqual([[t, "1985-03-15"]]);
    }
    expect(read("born 85-03-15")).toEqual([["85-03-15", "1985-03-15 2085-03-15"]]);
    expect(read("id 19850315 ok")).toEqual([["19850315", "1985-03-15"]]);
    expect(findDates("id 19850315", false)).toEqual([]);
  });

  it("reads month names and abbreviations, with ordinals and 'of'", () => {
    for (const t of ["4th of July 1988", "July 4th, 1988", "4 Jul 1988", "Jul. 4, 1988", "4-Jul-1988", "Jul-4-1988", "4th July, 1988", "JULY 4 1988"]) {
      expect(read(`born ${t}?`), t).toEqual([[t, "1988-07-04"]]);
    }
    expect(read("Friday, March 15th, 1985")).toEqual([["March 15th, 1985", "1985-03-15"]]);
    expect(read("15 Mar. 1985")).toEqual([["15 Mar. 1985", "1985-03-15"]]);
    expect(read("15-Mar-85")).toEqual([["15-Mar-85", "1985-03-15 2085-03-15"]]);
    expect(read("Sept 3 '85")).toEqual([["Sept 3 '85", "1985-09-03 2085-09-03"]]);
  });

  it("finds whole dates only, not inside other numbers, and nothing that is no date", () => {
    for (const t of [
      "born in 1985, in March 1985",
      "rows 11985-03-15",
      "(831) 555-0199",
      "123-45-6789",
      "4111 1111 1111 1234",
      "Orders!A1:F501",
      "Q1 2024",
      "13/13/85",
      "2/30/2024",
      "the mayor 5 2024",
      "total 5,390,000",
      "version 1.2.3",
      "192.168.1.15",
      "3/15/1985.5",
    ]) {
      expect(read(t), t).toEqual([]);
    }
  });

  it("reads text dates with 2-digit years", () => {
    expect(textDates("03/15/85").map(dateKey)).toEqual(["1985-03-15", "2085-03-15"]);
    expect(textDates("15-Mar-85").map(dateKey)).toEqual(["1985-03-15", "2085-03-15"]);
    expect(textDates("1/2/03").map(dateKey)).toEqual(["1903-01-02", "1903-02-01", "2003-01-02", "2003-02-01"]);
    // Nothing with other characters. (8 digits are read since the sixth review; see below.)
    for (const t of ["(831) 555-0199", "maria85@example.com", "$3/15/85"]) expect(textDates(t), t).toEqual([]);
  });

  it("gives a private column of 2-digit-year text dates every date it can mean", () => {
    const d = range([["Name", "Born"], ["Maria Lopez", "03/15/85"], ["Ana Ruiz", "04/02/90"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "stand_in" } });
    expect(privateDates(s).get("B")).toEqual(["1985-03-15", "2085-03-15", "1990-04-02", "1990-02-04", "2090-04-02", "2090-02-04"]);
  });

  it("replaces any date phrase that means a private date with that date's stand-in", () => {
    const serial = (y: number, m: number, day: number) => Math.round((Date.UTC(y, m - 1, day) - Date.UTC(1899, 11, 30)) / 86400000);
    const d = range([["Name", "Born"], ["Maria Lopez", serial(1985, 3, 15)], ["Ines Duarte", serial(1988, 7, 4)]], {
      formats: { 1: "m/d/yyyy" },
      text: { 1: { 1: "3/15/1985" }, 2: { 1: "7/4/1988" } },
    });
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "month" } });
    const map = new StandInMap();
    const maria = substituteUserText("born 3/15/1985?", [s], map);
    expect(maria).toBe(`born ${map.peek("3/15/1985")!}?`);
    for (const q of [
      "3/15/85",
      "03/15/85",
      "15/03/85",
      "15.3.85",
      "15-Mar-1985",
      "15-Mar-85",
      "Mar-15-1985",
      "1985/03/15",
      "1985.03.15",
      "March 15th, 1985",
      "15th March 1985",
      "Mar. 15, 1985",
      "15 Mar. 1985",
      "19850315",
    ]) {
      expect(substituteUserText(`born ${q}?`, [s], map), q).toBe(maria);
    }
    expect(substituteUserText("born on the 4th of July 1988 or July 4th, 1988", [s], map)).toBe(
      `born on the ${map.peek("7/4/1988")!} or ${map.peek("7/4/1988")!}`,
    );
    expect(substituteUserText("Friday, March 15, 1985", [s], map)).toBe(`Friday, ${map.peek("3/15/1985")!}`);
    // Dates that aren't private stay, and so do years and months alone.
    for (const q of ["born 3/16/85", "born 4th of July 1989", "born in 1985 or March 1985", "id 119850315"]) {
      expect(substituteUserText(q, [s], map), q).toBe(q);
    }
  });

  it("lets a longer private text that holds a date win over the date alone", () => {
    const d = range([["Note", "Born"], ["Visited on 3/15/85 with family", "3/15/85"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    const out = substituteUserText("Visited on 3/15/85 with family", [s], map);
    expect(map.valueOf(out)).toBe("Visited on 3/15/85 with family");
  });
});

describe("grouped numbers with zero decimals (fourth review)", () => {
  it("doesn't read the separator of a grouped number as its decimal mark: 5,390,000 is not 5,390", () => {
    const d = range([["Salary"], [5390], [125000]]);
    const s = spec(d, { A: { private: true, treatment: "range" } });
    const map = new StandInMap();
    for (const q of ["total 5,390,000", "total 5.390.000", "total 125,000,000"]) expect(substituteUserText(q, [s], map), q).toBe(q);
    const t = substituteUserText("total 5390", [s], map);
    for (const q of ["total 5,390.00", "total 5.390,00", "total 5390,000", "total 5 390,00", "total 5'390.00", "total 5,390"]) {
      expect(substituteUserText(q, [s], map), q).toBe(t);
    }
    // Only a digit, or the same separator and 3 digits, stop it (eighth review): "5,390,00" and
    // "125,000.000,0" hold 5,390 and 125,000.
    expect(substituteUserText("total 5,390,00", [s], map)).toBe(`${t},00`);
    expect(substituteUserText("total 125,000.000,0", [s], map)).toBe(`total ${map.peek("125000")!},0`);
  });
});

// ---------------------------------------------------------------------------------------------
// Fifth review: dates read in more forms, more spellings of names, numbers and phone numbers

describe("date phrases (fifth review)", () => {
  const read = (text: string) => findDates(text).map((f) => [text.slice(f.start, f.end), f.dates.map(dateKey).join(" ")]);
  const found = (text: string) => findDates(text).map((f) => text.slice(f.start, f.end));

  it("finds each date of two joined without spaces", () => {
    expect(found("3/15/85-4/2/90")).toEqual(["3/15/85", "4/2/90"]);
    expect(found("3/15/85–4/2/90")).toEqual(["3/15/85", "4/2/90"]);
    expect(found("3/15/85,4/2/90")).toEqual(["3/15/85", "4/2/90"]);
    expect(read("15.03.85-02.04.90")).toEqual([
      ["15.03.85", "1985-03-15 2085-03-15"],
      ["02.04.90", "1990-02-04 1990-04-02 2090-02-04 2090-04-02"],
    ]);
    expect(read("1985/03/15-1990/04/02")).toEqual([
      ["1985/03/15", "1985-03-15"],
      ["1990/04/02", "1990-04-02"],
    ]);
    expect(read("15-Mar-85/02-Apr-90")).toEqual([
      ["15-Mar-85", "1985-03-15 2085-03-15"],
      ["02-Apr-90", "1990-04-02 2090-04-02"],
    ]);
  });

  it("finds a date with a stray space or comma", () => {
    expect(read("born 15.3. 1985")).toEqual([["15.3. 1985", "1985-03-15"]]);
    expect(read("born 3/15, 1985")).toEqual([["3/15, 1985", "1985-03-15"]]);
    expect(read("born 1985. 3. 15.")).toEqual([["1985. 3. 15", "1985-03-15"]]);
    expect(read("born March 15 of 1985")).toEqual([["March 15 of 1985", "1985-03-15"]]);
  });

  it("reads month names in German, French, Spanish, Italian, Portuguese and Dutch, and Chinese, Japanese and Korean dates", () => {
    for (const t of ["15. März 1985", "15 mars 1985", "15 de marzo de 1985", "15 marzo 1985", "15 de março de 1985", "15 maart 1985", "15-MÄR-85"]) {
      expect(findDates(t).map((f) => f.dates.map(dateKey)[0]), t).toEqual(["1985-03-15"]);
    }
    expect(read("1er décembre 1985")).toEqual([["1er décembre 1985", "1985-12-01"]]);
    expect(read("15-dic-85")).toEqual([["15-dic-85", "1985-12-15 2085-12-15"]]);
    expect(read("1985年3月15日")).toEqual([["1985年3月15日", "1985-03-15"]]);
    expect(read("1985년 3월 15일")).toEqual([["1985년 3월 15일", "1985-03-15"]]);
    // A month and a year is still no date.
    for (const t of ["März 1985", "1985年3月", "de marzo de 1985"]) expect(findDates(t), t).toEqual([]);
  });

  it("marks the loosest forms, 8 digits and numbers separated only by spaces, as weak", () => {
    expect(findDates("19850315").map((f) => f.weak)).toEqual([true]);
    expect(findDates("15 03 85").map((f) => f.weak)).toEqual([true]);
    expect(findDates("06 15 03 85 12").every((f) => f.weak)).toBe(true);
    // Three numbers inside a longer run with the same separator: a phone number, an IP address.
    for (const t of ["06.45.23.12.80", "06-45-23-12-80", "10.12.1.85"]) {
      expect(findDates(t).length, t).toBeGreaterThan(0);
      expect(findDates(t).every((f) => f.weak), t).toBe(true);
    }
    expect(findDates("15.03.85-02.04.90").map((f) => f.weak)).toEqual([false, false]);
    // A 4-digit year inside a longer run is still a date: "SO-1985-03-15-001".
    expect(findDates("SO-1985-03-15-001").map((f) => f.weak)).toEqual([false]);
    for (const t of ["3/15/85", "15.3. 1985", "15-Mar-85", "March 15, 1985"]) expect(findDates(t).map((f) => f.weak), t).toEqual([false]);
  });
});

describe("textDates (fifth review)", () => {
  const keys = (t: string) => textDates(t).map(dateKey);

  it("reads ISO text with a time and an offset or zone, and a zone right after the date", () => {
    expect(keys("1985-03-15T00:00:00-05:00")).toEqual(["1985-03-15"]);
    expect(keys("1990-04-02 00:00:00+00")).toEqual(["1990-04-02"]);
    expect(keys("1982-02-18-08:00")).toEqual(["1982-02-18"]);
    expect(keys("2001-01-09 00:00:00 UTC")).toEqual(["2001-01-09"]);
    expect(keys("1988-07-04T00:00:00.000+01:00")).toEqual(["1988-07-04"]);
    expect(keys("1985-03-15Z")).toEqual(["1985-03-15"]);
    expect(keys("1985-03-15T10:30:00+0530")).toEqual(["1985-03-15"]);
    // The last part of a plain ISO date is not a zone.
    expect(keys("1985-03-15")).toEqual(["1985-03-15"]);
    for (const t of ["+1 212 555 0147", "(831) 555-0199", "15 Foo 1985"]) expect(textDates(t), t).toEqual([]);
  });

  it("reads dates written in other languages", () => {
    for (const t of ["15. März 1985", "15 mars 1985", "15 de marzo de 1985", "1985年3月15日", "1985년 3월 15일", "1985. 3. 15."]) {
      expect(keys(t), t).toEqual(["1985-03-15"]);
    }
  });

  it("reads a date cell's display text too, so a 1904 workbook's dates are right", () => {
    const serial = (y: number, m: number, day: number) => Math.round((Date.UTC(y, m - 1, day) - Date.UTC(1899, 11, 30)) / 86400000);
    const d = range(
      [
        ["Born 1904", "Born 1900", "Hired"],
        [serial(1985, 3, 15) - 1462, serial(1985, 3, 15), serial(1990, 4, 2)],
      ],
      { formats: { 0: "m/d/yyyy", 1: "m/d/yyyy", 2: "m/d/yyyy" }, text: { 1: { 0: "3/15/1985", 1: "3/15/1985", 2: "4/2/1990" } } },
    );
    const s = spec(d, { A: { private: true, treatment: "month" }, B: { private: true, treatment: "month" }, C: { private: true, treatment: "month" } });
    const dates = privateDates(s);
    // 1904: the serial read in the 1900 system names March 14, 1981; the display names the date.
    expect(dates.get("A")).toEqual(["1981-03-14", "1985-03-15"]);
    // 1900: the display agrees with the serial, which alone is kept (not also February 4, 1990).
    expect(dates.get("B")).toEqual(["1985-03-15"]);
    expect(dates.get("C")).toEqual(["1990-04-02"]);
    const map = new StandInMap();
    const token = substituteUserText("born 1985-03-15", [s], map);
    expect(token).not.toContain("1985");
    // A number cell showing "#####" is not a date's display.
    expect(privateDates(spec(range([["Born"], [serial(1985, 3, 15)]], { formats: { 0: "m/d/yyyy" }, text: { 1: { 0: "#####" } } }), { A: { private: true, treatment: "month" } })).get("A")).toEqual([
      "1985-03-15",
    ]);
  });
});

describe("more spellings in the question (fifth review)", () => {
  it("helpers: joined names, email local parts, parts between letters, short numbers", () => {
    expect(joinedNames("O'Brien")).toEqual(["OBrien"]);
    expect(joinedNames("Anne-Marie")).toEqual(["AnneMarie", "Anne Marie"]);
    expect(joinedNames("Maria Lopez-Ruiz")).toEqual(["Maria LopezRuiz"]);
    expect(joinedNames("D’Angelo")).toEqual(["DAngelo"]);
    for (const v of ["e-mail", "don't", "1985-03-15", "SO-10001"]) expect(joinedNames(v), v).toEqual([]);
    expect(emailLocalPart("sunflower77@gmail.com")).toBe("sunflower77");
    expect(emailLocalPart("kw.tokyo.1988@example.jp")).toBe("kw.tokyo.1988");
    for (const v of ["maria@example.com", "ab1@x.com", "not an email"]) expect(emailLocalPart(v), v).toBeNull();
    expect(letterSeparatedParts("(831) 555-0199 x204")).toEqual(["(831) 555-0199 ", "204"]);
    expect(letterSeparatedParts("GB29 NWBK 6016 1331 9268 19")).toEqual(["29 ", " 6016 1331 9268 19"]);
    expect(letterSeparatedParts("(831) 555-0199")).toEqual([]);
    expect(shortNumbers("83500")).toEqual([["83.5", "k"]]);
    expect(shortNumbers("1200000")).toEqual([["1200", "k"], ["1.2", "m"]]);
    expect(shortNumbers("83512")).toEqual([]);
    expect(shortNumbers("5.25")).toEqual([]);
  });

  it("replaces a phone number typed without its extension, and an IBAN's account number", () => {
    const d = range([["Name", "Phone", "IBAN"], ["Maria Lopez", "(831) 555-0199 x204", "GB29 NWBK 6016 1331 9268 19"], ["Ana Ruiz", "+1 212 555 0147 ext 3", ""]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "stand_in" }, C: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("Orders for 8315550199?", [s], map)).toBe(`Orders for ${map.peek("(831) 555-0199 x204")!}?`);
    expect(substituteUserText("Orders for 2125550147?", [s], map)).toBe(`Orders for ${map.peek("+1 212 555 0147 ext 3")!}?`);
    expect(substituteUserText("account 60161331926819", [s], map)).toBe(`account ${map.peek("GB29 NWBK 6016 1331 9268 19")!}`);
  });

  it("replaces a name typed without its apostrophe or hyphen, with a ligature written out, or with ae oe ue", () => {
    const d = range([["Name"], ["O'Brien"], ["Anne-Marie"], ["Mary O’Neill"], ["Griﬃths"], ["Müller"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    const cases: [string, string][] = [
      ["rows for OBrien", "O'Brien"],
      ["rows for obrien", "O'Brien"],
      ["rows for AnneMarie", "Anne-Marie"],
      ["rows for Anne Marie", "Anne-Marie"],
      ["rows for Mary ONeill", "Mary O’Neill"],
      ["rows for Griffiths", "Griﬃths"],
      ["rows for Mueller", "Müller"],
    ];
    for (const [q, value] of cases) expect(substituteUserText(q, [s], map), q).toBe(`rows for ${map.peek(value)!}`);
  });

  it("replaces a number written exactly in thousands or millions, or grouped with _", () => {
    const d = range([["Salary"], [83500], [1200000]]);
    const s = spec(d, { A: { private: true, treatment: "range" } });
    const map = new StandInMap();
    const t83 = substituteUserText("salary 83500", [s], map);
    for (const q of ["salary 83.5k", "salary 83.5 K", "salary 83,5k", "salary 83_500"]) expect(substituteUserText(q, [s], map), q).toBe(t83);
    const t12 = substituteUserText("salary 1200000", [s], map);
    for (const q of ["salary 1.2M", "salary 1200k", "salary 1_200_000"]) expect(substituteUserText(q, [s], map), q).toBe(t12);
    for (const q of ["distance 83.5km", "salary 83.51k", "salary 1.25M"]) expect(substituteUserText(q, [s], map), q).toBe(q);
  });

  it("replaces each date of two joined without spaces, and a date stored in another language typed with numbers", () => {
    const d = range([["Name", "Born"], ["Maria Lopez", "15. März 1985"], ["Ana Ruiz", "1990-04-02T00:00:00+02:00"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    const maria = map.tokenFor("15. März 1985", "TEXT");
    const ana = map.tokenFor("1990-04-02T00:00:00+02:00", "TEXT");
    expect(substituteUserText("born 3/15/85-4/2/90", [s], map)).toBe(`born ${maria}-${ana}`);
    expect(substituteUserText("born 1985-03-15 or April 2, 1990", [s], map)).toBe(`born ${maria} or ${ana}`);
    expect(substituteUserText("born 15.3. 1985", [s], map)).toBe(`born ${maria}`);
  });

  it("doesn't use a number's '#####' display as a value", () => {
    const d = range([["Phone"], [8315550199]], { text: { 1: { 0: "##########" } } });
    const s = spec(d, { A: { private: true, treatment: "stand_in" } });
    expect(substituteUserText("rows with ##########", [s], new StandInMap())).toBe("rows with ##########");
  });
});

// ---------------------------------------------------------------------------------------------
// Sixth review: plain numbers next to a comma, a decimal mark or zeros; dates written together

describe("plain numbers in the question (sixth review)", () => {
  const d = range([
    ["Customer", "ZIP", "Card last 4", "Salary", "Employee no", "Rate"],
    ["Maria Lopez", 93496, "0199", 83500, 4521, 12.75],
    ["Ana Ruiz", 94103, "4242", 91250, 1187, 18.5],
    ["Kenji Watanabe", 2139, "1881", 69000, 30422, 22.25],
  ]);
  const s = spec(d, {
    A: { private: true, treatment: "stand_in" },
    B: { private: true, treatment: "range" },
    C: { private: true, treatment: "stand_in" },
    D: { private: true, treatment: "range" },
    E: { private: true, treatment: "stand_in" },
    F: { private: true, treatment: "range" },
  });

  it("replaces a private number next to a comma or a decimal mark, and zeros before it", () => {
    const map = new StandInMap();
    const t = (q: string) => substituteUserText(q, [s], map);
    const zip = t("ZIP 93496").slice(4);
    const zip2 = t("ZIP 94103").slice(4);
    const cases: [string, string][] = [
      ["ZIPs 93496,94103", `ZIPs ${zip},${zip2}`],
      ["=IF(B2=93496,1,0)", `=IF(B2=${zip},1,0)`],
      ["{93496,94103}", `{${zip},${zip2}}`],
      ["ZIP 93496.5", `ZIP ${zip}.5`],
      ["ZIP 93496,00", `ZIP ${zip}`],
      ["ZIP 093496", `ZIP ${zip}`],
    ];
    for (const [q, want] of cases) expect(t(q), q).toBe(want);
    // Zeros before it go with it: "02139" for the 2139 Excel keeps, "004521" and "#04521" for 4521.
    const boston = t("ZIP 2139").slice(4);
    expect(t("ZIP 02139")).toBe(`ZIP ${boston}`);
    const emp = t("no 4521").slice(3);
    expect(t("employee 004521 and #04521")).toBe(`employee ${emp} and #${emp}`);
    expect(t("salaries 83500,91250")).toBe(`salaries ${t("salary 83500").slice(7)},${t("salary 91250").slice(7)}`);
    expect(t("cards 4242,1881")).toBe(`cards ${t("card 4242").slice(5)},${t("card 1881").slice(5)}`);
    expect(t("call 831.555.0199")).toBe(`call 831.555.${t("card 0199").slice(5)}`);
    // Zero decimals after its own are the same number.
    expect(t("rate 12.750")).toBe(t("rate 12.75"));
    expect(t("rate 12.75")).not.toContain("12.75");
  });

  it("leaves a number with a digit right next to the private one", () => {
    const map = new StandInMap();
    // Nor in the decimals of a number with one decimal point (seventh review).
    for (const q of ["id 934961", "lon -120.934961", "code 102139", "ref 45210", "ref 1004521", "rate 12.751", "total 835001", "id 10199"]) {
      expect(substituteUserText(q, [s], map), q).toBe(q);
    }
  });
});

describe("dates written together, with _ or other spaces next to the month name (sixth review)", () => {
  const read = (text: string) => findDates(text).map((f) => [text.slice(f.start, f.end), f.dates.map(dateKey).join(" ")]);

  it("reads a month name with no separator next to it, or any mix of _ . , / dashes and spaces", () => {
    for (const t of [
      "15MAR1985",
      "15Mar1985",
      "15March1985",
      "15thMarch1985",
      "March15, 1985",
      "Mar15 1985",
      "15_March_1985",
      "March_15,_1985",
      "15_Mar_1985",
      "15_of_March_1985",
      "15 -_ Mar ._ 1985",
      "1985MAR15",
      "1985_Mar_15",
    ]) {
      expect(read(`born ${t}?`), t).toEqual([[t, "1985-03-15"]]);
    }
    expect(read("born 15MAR85")).toEqual([["15MAR85", "1985-03-15 2085-03-15"]]);
  });

  it("reads every space and invisible separator between the parts", () => {
    for (const sep of ["\u2028", "\u2029", "\u205f", "\ufeff", "\u200b", "\u200c", "\u200d", "\u2060", "\u3000", "\u1680", "\u00a0"]) {
      const t = `15${sep}Mar${sep}1985`;
      expect(read(t), JSON.stringify(t)).toEqual([[t, "1985-03-15"]]);
      const u = `March${sep}15,${sep}1985`;
      expect(read(u), JSON.stringify(u)).toEqual([[u, "1985-03-15"]]);
    }
  });

  it("reads 'March the 15th, 1985'", () => {
    expect(read("born March the 15th, 1985")).toEqual([["March the 15th, 1985", "1985-03-15"]]);
  });

  it("reads a slash look-alike as a slash", () => {
    for (const t of ["3\u22152\u22151990", "3\u20442\u20441990", "3\uff0f2\uff0f1990"]) {
      expect(read(t), t).toEqual([[t, "1990-03-02 1990-02-03"]]);
    }
  });

  it("still needs a separator between numbers alone, and between the day and the year", () => {
    // "1985_03_15" is read since the seventh review.
    for (const t of ["031585", "15031985", "March 151985", "Mar1985", "rows 3may", "15-III-1985"]) expect(read(t), t).toEqual([]);
  });

  it("reads such a date as a whole stored text, and 8 digits with eightDigitDates (seventh review)", () => {
    const keys = (t: string) => textDates(t).map(dateKey);
    for (const t of ["15MAR1985", "15_Mar_1985", "March15, 1985", "15\u2028Mar\u20281985"]) expect(keys(t), JSON.stringify(t)).toEqual(["1985-03-15"]);
    expect(keys("3\u22152\u22151990")).toEqual(["1990-03-02", "1990-02-03"]);
    // 8 digits are as often an ID: their dates are eightDigitDates, which count only in free text.
    for (const t of ["19910919", "19850315"]) expect(textDates(t), t).toEqual([]);
    const loose = (t: string) => eightDigitDates(t).map(dateKey);
    expect(loose("19910919")).toEqual(["1991-09-19"]);
    expect(loose("21000101")).toEqual(["2100-01-01"]);
    for (const t of ["18991231", "21010101", "20230231", "20231301", "1234567", "123456789"]) expect(eightDigitDates(t), t).toEqual([]);
  });

  it("replaces such a date in the question with the stand-in of the private date", () => {
    const serial = (y: number, m: number, day: number) => Math.round((Date.UTC(y, m - 1, day) - Date.UTC(1899, 11, 30)) / 86400000);
    const d = range([["Name", "Born"], ["Maria Lopez", serial(1985, 3, 15)]], { formats: { 1: "m/d/yyyy" }, text: { 1: { 1: "3/15/1985" } } });
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "month" } });
    const map = new StandInMap();
    const maria = substituteUserText("born 3/15/1985?", [s], map);
    for (const q of ["15MAR1985", "15March1985", "March15, 1985", "Mar15 1985", "15_March_1985", "March_15,_1985", "15\u2028Mar\u20281985", "15\ufeffMar 1985", "March the 15th, 1985"]) {
      expect(substituteUserText(`born ${q}?`, [s], map), JSON.stringify(q)).toBe(maria);
    }
  });

  it("gives a private column of dates stored written together, or as 8 digits, their dates", () => {
    const d = range([["Name", "Born", "Hired"], ["Maria Lopez", "15MAR1985", "19910919"], ["Ana Ruiz", "02APR1990", "20010109"]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "stand_in" }, C: { private: true, treatment: "stand_in" } });
    expect(privateDates(s).get("B")).toEqual(["1985-03-15", "1990-04-02"]);
    // Since the seventh review, 8 digits give dates the auditor counts only in free text.
    expect(privateDates(s).get("C")).toEqual([]);
    expect(privateLooseDates(s).get("C")).toEqual(["1991-09-19", "2001-01-09"]);
    const map = new StandInMap();
    expect(substituteUserText("born 1985-03-15, hired 9/19/1991", [s], map)).toBe(`born ${map.peek("15MAR1985")!}, hired ${map.peek("19910919")!}`);
  });
});

// ---------------------------------------------------------------------------------------------
// Seventh review

describe("two dates next to each other in the question (seventh review)", () => {
  const serial = (y: number, m: number, day: number) => Math.round((Date.UTC(y, m - 1, day) - Date.UTC(1899, 11, 30)) / 86400000);
  // Maria was born on March 15, 1985 and Ana on September 2, 1985: "1985 - September 2" also
  // reads as Ana's date.
  const d = range(
    [
      ["Name", "Born"],
      ["Maria Lopez", serial(1985, 3, 15)],
      ["Ana Ruiz", serial(1985, 9, 2)],
    ],
    { formats: { 1: "m/d/yyyy" }, text: { 1: { 1: "3/15/1985" }, 2: { 1: "9/2/1985" } } },
  );
  const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "month" } });
  const tokens = () => {
    const map = new StandInMap();
    const maria = substituteUserText("born 3/15/1985", [s], map).slice(5);
    const ana = substituteUserText("born 9/2/1985", [s], map).slice(5);
    return { map, maria, ana };
  };

  it("replaces both dates, not the phrase that runs from one into the other", () => {
    const { map, maria, ana } = tokens();
    const cases: [string, string][] = [
      ["born March 15, 1985 - September 2, 1985?", `born ${maria} - ${ana}?`],
      ["born 15 March 1985 - September 2 1985?", `born ${maria} - ${ana}?`],
      ["born 15 Mar 1985 – Sep 2 1985", `born ${maria} – ${ana}`],
      ["born March 15, 1985, September 2, 1985", `born ${maria}, ${ana}`],
      ["born 15.03.1985 / September 2 1985", `born ${maria} / ${ana}`],
      ["born 15MAR1985, MAR_15_1985", `born ${maria}, ${maria}`],
      ["born 3/15/1985 - 9/2/1985", `born ${maria} - ${ana}`],
    ];
    for (const [q, want] of cases) expect(substituteUserText(q, [s], map), q).toBe(want);
  });

  it("replaces the whole stretch as one stand-in when a part of a private date would be left", () => {
    // Only Ana's date is in this sheet: "1985 - September 2" and "September 2, 1985" both mean it.
    const only = spec(range([["Name", "Born"], ["Ana Ruiz", serial(1985, 9, 2)]], { formats: { 1: "m/d/yyyy" }, text: { 1: { 1: "9/2/1985" } } }), {
      B: { private: true, treatment: "month" },
    });
    const map = new StandInMap();
    const ana = substituteUserText("born 9/2/1985", [only], map).slice(5);
    expect(substituteUserText("born 1985 - September 2, 1985?", [only], map)).toBe(`born ${ana}?`);
  });

  it("still lets a longer private text that holds a date win", () => {
    const notes = range([["Note"], ["Moved on 3/15/1985 to Lyon"]]);
    const n = spec(notes, { A: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("What about Moved on 3/15/1985 to Lyon?", [n], map)).toBe(`What about ${map.peek("Moved on 3/15/1985 to Lyon")!}?`);
  });
});

describe("plain numbers in the decimals of another number (seventh review)", () => {
  const d = range([
    ["Customer", "ZIP", "Card last 4"],
    ["Maria Lopez", 81939, "0199"],
    ["Ana Ruiz", 60112, "4242"],
    ["Kenji Watanabe", 93496, "1881"],
  ]);
  const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "range" }, C: { private: true, treatment: "stand_in" } });

  it("replaces the decimals of a number with one decimal point too: the question is free text (eighth review)", () => {
    const map = new StandInMap();
    for (const q of ["lon -116.081939?", "lat 43.060112", "x 1.93496", "at 0.0093496 or $12.93496"]) expect(substituteUserText(q, [s], map), q).not.toBe(q);
    expect(substituteUserText("x 1.93496", [s], map)).toBe(`x 1.${map.peek("93496")!}`);
  });

  it("still replaces the last group of a dotted number, a list, and zeros that follow no digit", () => {
    const map = new StandInMap();
    const t = (q: string) => substituteUserText(q, [s], map);
    const card = t("card 0199").slice(5);
    const zip = t("ZIP 93496").slice(4);
    const lat = t("ZIP 60112").slice(4);
    expect(t("call 831.555.0199")).toBe(`call 831.555.${card}`);
    expect(t("call 1.831.555.0199")).toBe(`call 1.831.555.${card}`);
    expect(t("ZIPs 10001,93496")).toBe(`ZIPs 10001,${zip}`);
    expect(t("ZIP 060112")).toBe(`ZIP ${lat}`);
    expect(t("ZIP 93496.0")).toBe(`ZIP ${zip}`);
  });
});

describe("dates stored as 8 digits, and SAS date-times (seventh review)", () => {
  it("reads every date 8 digits can mean with a year from 1900 to 2100", () => {
    const keys = (t: string) => eightDigitDates(t).map(dateKey);
    expect(keys("19850315")).toEqual(["1985-03-15"]);
    expect(keys("03151985")).toEqual(["1985-03-15"]);
    expect(keys("15031985")).toEqual(["1985-03-15"]);
    expect(keys("01022001")).toEqual(["2001-01-02", "2001-02-01"]);
    expect(keys("20240115")).toEqual(["2024-01-15"]);
    for (const t of ["1985031", "198503150", "19851315", "ab850315", "2024-01-15"]) expect(eightDigitDates(t), t).toEqual([]);
  });

  it("gives an 8-digit column its dates separately, from text and from a whole number without a date format", () => {
    const d = range(
      [
        ["Name", "Born text", "Born number", "Born mdy", "Hired"],
        ["Maria Lopez", "19850315", 19850315, "03151985", 19850315],
        ["Ana Ruiz", "19910919", 19910919, "09191991", 20.5],
      ],
      { formats: { 4: "m/d/yyyy" } },
    );
    const s = spec(d, {
      B: { private: true, treatment: "stand_in" },
      C: { private: true, treatment: "range" },
      D: { private: true, treatment: "stand_in" },
      E: { private: true, treatment: "range" },
    });
    const loose = privateLooseDates(s);
    expect(loose.get("B")).toEqual(["1985-03-15", "1991-09-19"]);
    expect(loose.get("C")).toEqual(["1985-03-15", "1991-09-19"]);
    expect(loose.get("D")).toEqual(["1985-03-15", "1991-09-19"]);
    // A number with a date format is a date serial, not 8 digits.
    expect(loose.get("E")).toEqual([]);
    for (const letter of ["B", "C", "D"]) expect(privateDates(s).get(letter), letter).toEqual([]);
  });

  it("replaces a date typed in the question for a date stored as 8 digits, text or number", () => {
    const d = range([
      ["Name", "Born"],
      ["Maria Lopez", 19910919],
      ["Ana Ruiz", "03151985"],
    ]);
    const s = spec(d, { B: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    const q = substituteUserText("born 9/19/1991 or March 15, 1985?", [s], map);
    expect(q).toBe(`born ${map.peek("19910919")!} or ${map.peek("03151985")!}?`);
  });

  it("reads a SAS date-time stored as text", () => {
    expect(textDates("15MAR1985:00:00:00").map(dateKey)).toEqual(["1985-03-15"]);
    expect(textDates("19SEP1991:10:30:00").map(dateKey)).toEqual(["1991-09-19"]);
    const s = spec(range([["Name", "Born"], ["Maria Lopez", "15MAR1985:00:00:00"]]), { B: { private: true, treatment: "stand_in" } });
    expect(privateDates(s).get("B")).toEqual(["1985-03-15"]);
    const map = new StandInMap();
    expect(substituteUserText("born 3/15/1985?", [s], map)).toBe(`born ${map.peek("15MAR1985:00:00:00")!}?`);
  });
});

describe("characters that can't be seen (seventh review)", () => {
  const read = (text: string) => findDates(text).map((f) => [text.slice(f.start, f.end), f.dates.map(dateKey).join(" ")]);
  const marks = ["\u200e", "\u200f", "\u061c", "\u00ad", "\u202a", "\u202c", "\u202e", "\u2066", "\u2069"];

  it("reads a date with them between its parts, where it is in the text", () => {
    for (const m of marks) {
      expect(read(`born ${m}3/${m}15/${m}1985?`), JSON.stringify(m)).toEqual([[`3/${m}15/${m}1985`, "1985-03-15"]]);
      expect(read(`born 15${m}Mar${m}1985`), JSON.stringify(m)).toEqual([[`15${m}Mar${m}1985`, "1985-03-15"]]);
    }
    expect(textDates("\u200e3/\u200e15/\u200e1985").map(dateKey)).toEqual(["1985-03-15"]);
    // A form feed counts as a space.
    expect(read("born 15\fMar\f1985")).toEqual([["15\fMar\f1985", "1985-03-15"]]);
  });

  it("replaces a private date or name with them inside, in the question", () => {
    const serial = (y: number, m: number, day: number) => Math.round((Date.UTC(y, m - 1, day) - Date.UTC(1899, 11, 30)) / 86400000);
    const d = range([["Name", "Born"], ["Maria Lopez", serial(1985, 3, 15)]], { formats: { 1: "m/d/yyyy" }, text: { 1: { 1: "3/15/1985" } } });
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "month" } });
    const map = new StandInMap();
    const dob = substituteUserText("born 3/15/1985", [s], map).slice(5);
    const maria = substituteUserText("Maria Lopez", [s], map);
    for (const m of [...marks, "\f"]) {
      expect(substituteUserText(`Who was born ${m}3/${m}15/${m}1985?`, [s], map), JSON.stringify(m)).toBe(`Who was born ${m}${dob}?`);
    }
    for (const m of marks) expect(substituteUserText(`Who is Ma${m}ria Lo${m}pez?`, [s], map), JSON.stringify(m)).toBe(`Who is ${maria}?`);
    // A form feed counts as a space (eighth review).
    expect(substituteUserText("Who is Maria\fLopez?", [s], map)).toBe(`Who is ${maria}?`);
  });
});

describe("numeric dates with _ or \\ (seventh review)", () => {
  const read = (text: string) => findDates(text).map((f) => [text.slice(f.start, f.end), f.dates.map(dateKey).join(" ")]);

  it("reads numbers separated by _ or \\ as a date", () => {
    expect(read("born 1985_03_15")).toEqual([["1985_03_15", "1985-03-15"]]);
    expect(read("born 15_03_1985")).toEqual([["15_03_1985", "1985-03-15"]]);
    expect(read("born 3\\15\\1985")).toEqual([["3\\15\\1985", "1985-03-15"]]);
    // A doubled backslash, as it is in a request body.
    expect(read("born 3\\\\15\\\\1985")).toEqual([["3\\\\15\\\\1985", "1985-03-15"]]);
    expect(read("born 15\\Mar\\1985")).toEqual([["15\\Mar\\1985", "1985-03-15"]]);
    expect(read("born 15\\Nov\\1985")).toEqual([["15\\Nov\\1985", "1985-11-15"]]);
    expect(textDates("1985_03_15").map(dateKey)).toEqual(["1985-03-15"]);
    expect(textDates("3\\15\\1985").map(dateKey)).toEqual(["1985-03-15"]);
  });

  it("replaces such a date in the question", () => {
    const serial = (y: number, m: number, day: number) => Math.round((Date.UTC(y, m - 1, day) - Date.UTC(1899, 11, 30)) / 86400000);
    const d = range([["Name", "Born"], ["Maria Lopez", serial(1985, 3, 15)]], { formats: { 1: "m/d/yyyy" }, text: { 1: { 1: "3/15/1985" } } });
    const s = spec(d, { B: { private: true, treatment: "month" } });
    const map = new StandInMap();
    const dob = substituteUserText("born 3/15/1985", [s], map).slice(5);
    for (const q of ["1985_03_15", "15_03_1985", "3\\15\\1985"]) expect(substituteUserText(`born ${q}?`, [s], map), q).toBe(`born ${dob}?`);
  });
});

describe("codes typed without their separators (seventh review)", () => {
  it("gives a code of letters and digits without its - / . and spaces", () => {
    expect(codeWithoutSeparators("AB-1234-CD")).toBe("AB1234CD");
    expect(codeWithoutSeparators("X7 42.B")).toBe("X742B");
    expect(codeWithoutSeparators("GB/29/NW")).toBe("GB29NW");
    expect(codeWithoutSeparators("AB-\u200e1234-CD")).toBe("AB1234CD");
    for (const v of ["Maria Lopez", "1234-5678", "AB1234CD", "A-1", "a-b-1", "(831) 555-0199 x204", "AB_1234", "12 Main St, Apt 4"]) {
      expect(codeWithoutSeparators(v), v).toBeNull();
    }
  });

  it("replaces a private code typed without them in the question", () => {
    const s = spec(range([["Driver", "Plate"], ["Maria Lopez", "AB-1234-CD"]]), { B: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    const plate = substituteUserText("Who drives AB-1234-CD?", [s], map);
    expect(substituteUserText("Who drives AB1234CD?", [s], map)).toBe(plate);
    expect(substituteUserText("Who drives ab1234cd?", [s], map)).toBe(plate);
    // A letter touching a code of 6+ letters and digits doesn't hide it from the auditor (eighth
    // review), but only a whole code is replaced (ninth review): the auditor blocks this question.
    expect(substituteUserText("Who drives XAB1234CD?", [s], map)).toBe("Who drives XAB1234CD?");
  });
});

// ---------------------------------------------------------------------------------------------
// Eighth review

describe("a plain number after digits and a period in the question (eighth review)", () => {
  const d = range([
    ["Customer", "ZIP", "Employee no"],
    ["Maria Lopez", 93496, 4521],
    ["Ana Ruiz", 94103, 4528],
    ["Kenji Watanabe", 60614, 4535],
  ]);
  const s = spec(d, { B: { private: true, treatment: "range" }, C: { private: true, treatment: "stand_in" } });

  it("replaces a private number typed as a numbered item or after another number and a period", () => {
    const map = new StandInMap();
    const t = (q: string) => substituteUserText(q, [s], map);
    const [a, b, c] = ["93496", "94103", "60614"].map((z) => t(`ZIP ${z}`).slice(4));
    expect(t("Orders per ZIP for:\n1.93496\n2.94103\n3.60614")).toBe(`Orders per ZIP for:\n1.${a}\n2.${b}\n3.${c}`);
    expect(t("ZIPs 93496.94103?")).toBe(`ZIPs ${a}.${b}?`);
    expect(t("ZIPs #1.93496 and 60614.93496.94103")).toBe(`ZIPs #1.${a} and ${c}.${a}.${b}`);
    const emp = t("employee 4528").slice(9);
    for (const q of ["Hours for #1.4528?", "Hours for v1.4528?", "Hours for employees 4521.4528?"]) expect(t(q), q).toContain(`.${emp}?`);
  });
});

describe("characters that can't be seen, all of them (eighth review)", () => {
  const read = (text: string) => findDates(text).map((f) => [text.slice(f.start, f.end), f.dates.map(dateKey).join(" ")]);
  // Zero-width space, non-joiner and joiner, word joiner, BOM, U+180E, U+2061 to U+206F, a
  // variation selector, a tag character, Hangul fillers, the combining grapheme joiner, and
  // control characters that aren't whitespace (backspace, DEL, C0 and C1 ones).
  const hidden = ["\u200b", "\u200c", "\u200d", "\u2060", "\ufeff", "\u180e", "\u2061", "\u2064", "\u206a", "\u206f", "\ufe0f", "\u{e0020}", "\u3164", "\uffa0", "\u115f", "\u034f", "\b", "\u007f", "\u0001", "\u001f", "\u009f"];

  it("reads a date with them anywhere in it, where it is in the text", () => {
    for (const m of hidden) {
      expect(read(`born 3${m}/${m}15${m}/${m}1985?`), JSON.stringify(m)).toEqual([[`3${m}/${m}15${m}/${m}1985`, "1985-03-15"]]);
      expect(read(`born Mar${m}ch 1${m}5, 19${m}85`), JSON.stringify(m)).toEqual([[`Mar${m}ch 1${m}5, 19${m}85`, "1985-03-15"]]);
    }
    expect(textDates("\u200b3/\ufe0f15/\u{e0020}1985").map(dateKey)).toEqual(["1985-03-15"]);
  });

  it("reads a whitespace control as a space in a date", () => {
    for (const m of ["\t", "\n", "\v", "\f", "\r", "\u0085"]) {
      expect(read(`born 15${m}Mar${m}1985`), JSON.stringify(m)).toEqual([[`15${m}Mar${m}1985`, "1985-03-15"]]);
      // Not skipped: "19 85" is not 1985.
      expect(read(`born 3/15/19${m}85`).map(([, dates]) => dates).join(" "), JSON.stringify(m)).not.toContain("1985-03-15");
    }
  });

  it("replaces a private value with them inside, in the question", () => {
    const serial = (y: number, m: number, day: number) => Math.round((Date.UTC(y, m - 1, day) - Date.UTC(1899, 11, 30)) / 86400000);
    const d = range(
      [
        ["Name", "Born", "ZIP", "Plate"],
        ["Maria Lopez", serial(1985, 3, 15), 93496, "AB-1234-CD"],
      ],
      { formats: { 1: "m/d/yyyy" }, text: { 1: { 1: "3/15/1985" } } },
    );
    const s = spec(d, {
      A: { private: true, treatment: "stand_in" },
      B: { private: true, treatment: "month" },
      C: { private: true, treatment: "range" },
      D: { private: true, treatment: "stand_in" },
    });
    const map = new StandInMap();
    const one = (q: string) => substituteUserText(q, [s], map);
    const [maria, dob, zip, plate] = ["Maria Lopez", "3/15/1985", "93496", "AB-1234-CD"].map(one);
    for (const m of hidden) {
      const k = JSON.stringify(m);
      expect(one(`Ma${m}ria Lo${m}pez`), k).toBe(maria);
      expect(one(`3${m}/15/19${m}85`), k).toBe(dob);
      expect(one(`93${m}496`), k).toBe(zip);
      expect(one(`AB${m}1234${m}CD`), k).toBe(plate);
    }
  });

  it("counts a whitespace control as a space in the question", () => {
    const d = range([["Name", "ZIP"], ["Maria Lopez", 93496]]);
    const s = spec(d, { A: { private: true, treatment: "stand_in" }, B: { private: true, treatment: "range" } });
    const map = new StandInMap();
    const maria = substituteUserText("Maria Lopez", [s], map);
    const zip = substituteUserText("93496", [s], map);
    for (const m of ["\t", "\v", "\f", "\u0085"]) {
      expect(substituteUserText(`Maria${m}Lopez`, [s], map), JSON.stringify(m)).toBe(maria);
      // "93 496", grouped with a space.
      expect(substituteUserText(`93${m}496`, [s], map), JSON.stringify(m)).toBe(zip);
    }
  });
});

describe("numeric dates with two separators, and look-alikes of - and \\ (eighth review)", () => {
  const read = (text: string) => findDates(text).map((f) => [text.slice(f.start, f.end), f.dates.map(dateKey).join(" ")]);
  const found = (text: string) => findDates(text).map((f) => text.slice(f.start, f.end));

  it("reads numbers with two different separators when the year has 4 digits", () => {
    for (const t of ["3_15/1985", "3_15-1985", "3/15_1985", "3.15-1985", "3\\15/1985", "1985-03/15", "1985_3.15"]) {
      expect(read(`born ${t}`), t).toEqual([[t, "1985-03-15"]]);
    }
    // Not with a 2-digit year, nor inside a longer run of numbers.
    expect(read("born 3_15/85")).toEqual([]);
    expect(found("1985/03/15-1990/04/02")).toEqual(["1985/03/15", "1990/04/02"]);
    expect(found("3/15/1985-4/2/1990")).toEqual(["3/15/1985", "4/2/1990"]);
  });

  it("reads a dash look-alike as - and a backslash look-alike as \\", () => {
    for (const dash of ["\u2010", "\u2011", "\u2012", "\u2013", "\u2014", "\u2015", "\u2212", "\ufe58", "\ufe63", "\uff0d"]) {
      expect(read(`3${dash}15${dash}1985`), JSON.stringify(dash)).toEqual([[`3${dash}15${dash}1985`, "1985-03-15"]]);
      expect(read(`15${dash}Mar${dash}1985`), JSON.stringify(dash)).toEqual([[`15${dash}Mar${dash}1985`, "1985-03-15"]]);
    }
    for (const backslash of ["\ufe68", "\u29f5", "\uff3c"]) {
      expect(read(`3${backslash}15${backslash}1985`), JSON.stringify(backslash)).toEqual([[`3${backslash}15${backslash}1985`, "1985-03-15"]]);
      expect(textDates(`15${backslash}Mar${backslash}1985`).map(dateKey), JSON.stringify(backslash)).toEqual(["1985-03-15"]);
    }
  });

  it("replaces such a date, and a number typed with a dash look-alike, in the question", () => {
    const serial = (y: number, m: number, day: number) => Math.round((Date.UTC(y, m - 1, day) - Date.UTC(1899, 11, 30)) / 86400000);
    const d = range([["Name", "Born", "Phone"], ["Maria Lopez", serial(1985, 3, 15), "(831) 555-0199"]], { formats: { 1: "m/d/yyyy" }, text: { 1: { 1: "3/15/1985" } } });
    const s = spec(d, { B: { private: true, treatment: "month" }, C: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    const dob = substituteUserText("3/15/1985", [s], map);
    const phone = substituteUserText("(831) 555-0199", [s], map);
    for (const q of ["3_15/1985", "3_15-1985", "3\ufe6815\ufe681985", "3\u29f515\u29f51985", "3\ufe6315\ufe631985"]) expect(substituteUserText(q, [s], map), q).toBe(dob);
    for (const q of ["831\ufe63555\ufe630199", "831\ufe58555\ufe580199", "831_555_0199"]) expect(substituteUserText(q, [s], map), q).toBe(phone);
  });
});

describe("codes typed with some or other separators (eighth review)", () => {
  it("gives the letters and digits of a value of 6 or more that mixes both", () => {
    expect(codeLetters("ab-1234-cd")).toBe("ab1234cd");
    expect(codeLetters("x7 42.b9")).toBe("x742b9");
    expect(codeLetters("pol/2024/00123")).toBe("pol202400123");
    for (const v of ["a-1030", "ab-12", "x7-42-b", "maria lopez", "123-45-6789", ""]) expect(codeLetters(v), v).toBeNull();
  });

  it("replaces a private code typed with some of its separators or other ones, as a whole word", () => {
    const s = spec(range([["Driver", "Plate", "Part"], ["Maria Lopez", "AB-1234-CD", "X7-42-B9"]]), {
      B: { private: true, treatment: "stand_in" },
      C: { private: true, treatment: "stand_in" },
    });
    const map = new StandInMap();
    const plate = substituteUserText("AB-1234-CD", [s], map);
    const part = substituteUserText("X7-42-B9", [s], map);
    for (const q of ["AB-1234CD", "AB 1234CD", "AB\u20131234CD", "AB.1234.CD", "ab 12 34 cd", "AB1234-CD", "AB\u00b71234\u00b7CD"]) {
      expect(substituteUserText(q, [s], map), q).toBe(plate);
    }
    for (const q of ["X7_42_B9", "x7 42 b9", "X7\u201342\u2013B9"]) expect(substituteUserText(q, [s], map), q).toBe(part);
    // A letter touching it: the auditor blocks these (ninth review).
    expect(substituteUserText("Who has AB1234CDs?", [s], map)).toBe("Who has AB1234CDs?");
    expect(substituteUserText("Who has xAB 1234CD?", [s], map)).toBe("Who has xAB 1234CD?");
    // Only where nothing but separators is between its letters and digits.
    expect(substituteUserText("AB, 1234CD", [s], map)).toBe("AB, 1234CD");
  });
});

describe("8-digit values are dates only in a column of dates (eighth review)", () => {
  const invoices = (asText: boolean) =>
    range([["Invoice no", "Amount"], ...Array.from({ length: 200 }, (_, i): CellValue[] => [asText ? String(20240001 + i) : 20240001 + i, 100 + i])]);

  it("gives no dates for numbers that are dates only now and then, such as invoice numbers", () => {
    for (const asText of [false, true]) {
      const s = spec(invoices(asText), { A: { private: true, treatment: "stand_in" } });
      expect(privateLooseDates(s).get("A"), String(asText)).toEqual([]);
      const map = new StandInMap();
      for (const q of ["Total after 3/15/2024", "Amount on March 15, 2024", "Since 1 Feb 2024"]) {
        expect(substituteUserText(q, [s], map), q).toBe(q);
      }
      // The invoice number itself is still replaced, typed as a date too (tenth review).
      expect(substituteUserText("Amount of 20240101?", [s], map)).toBe(`Amount of ${map.peek("20240101")!}?`);
      expect(substituteUserText("Invoices between 2024-01-01 and 2024-06-30", [s], map)).toBe(`Invoices between ${map.peek("20240101")!} and 2024-06-30`);
    }
  });

  it("still reads a column of dates of birth stored as 8 digits, as text or numbers", () => {
    // Everyone born on March 15, from 1960 on; the first `bad` values are no date (month 13).
    const born = (asText: boolean, bad: number) => {
      const rows: CellValue[][] = [["Name", "Born"]];
      for (let i = 0; i < 20; i++) {
        const v = i < bad ? `${1990 + i}1340` : `${1960 + i}0315`;
        rows.push([`Person ${i}`, asText ? v : Number(v)]);
      }
      return spec(range(rows), { B: { private: true, treatment: "stand_in" } });
    };
    for (const asText of [false, true]) {
      // 1 of 20 is no date: 95% are.
      const s = born(asText, 1);
      expect(privateLooseDates(s).get("B"), String(asText)).toContain("1965-03-15");
      const map = new StandInMap();
      expect(substituteUserText("Who was born 3/15/1965?", [s], map), String(asText)).toBe(`Who was born ${map.peek("19650315")!}?`);
      // Half of them are dates: enough (ninth review). 11 of 20 are no date: not enough.
      expect(privateLooseDates(born(asText, 10)).get("B"), String(asText)).toContain("1975-03-15");
      expect(privateLooseDates(born(asText, 11)).get("B"), String(asText)).toEqual([]);
    }
  });
});

describe("grouped numbers in the question (eighth review)", () => {
  it("replaces a grouped private number with a decimal part after it, but not one another group continues", () => {
    const s = spec(range([["Customer", "ZIP"], ["Maria Lopez", 93496], ["Ana Ruiz", 94103]]), { B: { private: true, treatment: "range" } });
    const map = new StandInMap();
    const [a, b] = ["93496", "94103"].map((z) => substituteUserText(z, [s], map));
    expect(substituteUserText("ZIPs 93 496,94 103?", [s], map)).toBe(`ZIPs ${a},${b}?`);
    expect(substituteUserText("ZIPs 93,496.5 or 93.496,5", [s], map)).toBe(`ZIPs ${a}.5 or ${a},5`);
    for (const q of ["total 1 093 496", "total 93,496,000"]) expect(substituteUserText(q, [s], map), q).toBe(q);
    // Grouped with a space, only a digit right next to it continues it (ninth review).
    expect(substituteUserText("total 93 496 000", [s], map)).toBe(`total ${a} 000`);
  });
});

// ---------------------------------------------------------------------------------------------
// Ninth review

describe("characters matching drops still separate words and numbers (ninth review)", () => {
  const customers = () =>
    spec(range([["Customer", "ZIP"], ["Maria Lopez", 93496], ["Ana Ruiz", 94103], ["Kenji Watanabe", 10001]]), {
      A: { private: true, treatment: "stand_in" },
      B: { private: true, treatment: "stand_in" },
    });
  const hidden = ["\u200b", "\u2060", "\ufeff", "\u3164", "\ufe0f", "\u034f", "\u0001", "\u007f", "\u00ad", "\u200e", "\f", "\v"];

  it("says where characters were dropped", () => {
    expect(stripDropped("Mr\u200bWata\u200b\u200bnabe\u2060")).toEqual({ text: "MrWatanabe", at: [0, 1, 3, 4, 5, 6, 9, 10, 11, 12, 14], gap: Uint8Array.from([0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1]) });
    expect(stripDropped("934\f96")).toMatchObject({ text: "93496", gap: Uint8Array.from([0, 0, 0, 1, 0, 0]) });
    expect(stripDropped("Maria Lopez")).toEqual({ text: "Maria Lopez", at: null, gap: null });
    // A tab or a line break is not dropped: it is a space.
    expect(stripDropped("93\t496\n").gap).toBeNull();
  });

  it("replaces a value with a dropped character as the boundary next to it", () => {
    for (const x of hidden) {
      const map = new StandInMap();
      const s = customers();
      const [maria, a, b] = ["Maria Lopez", "93496", "94103"].map((v) => substituteUserText(v, [s], map));
      const k = JSON.stringify(x);
      expect(substituteUserText(`ZIPs 93496${x}94103?`, [s], map), k).toBe(`ZIPs ${a}${x}${b}?`);
      expect(substituteUserText(`ZIP 93496${x}2?`, [s], map), k).toBe(`ZIP ${a}${x}2?`);
      expect(substituteUserText(`Ask${x}Maria Lopez${x}s`, [s], map), k).toBe(`Ask${x}${maria}${x}s`);
    }
  });

  it("reads one as the end of a date, and between its numbers as a space", () => {
    const keys = (t: string) => findDates(t).flatMap((f) => f.dates.map(dateKey));
    for (const x of ["\u200b", "\u2060", "\u00ad", "\u200e"]) {
      expect(keys(`Born 3/15/1985${x}5`), JSON.stringify(x)).toContain("1985-03-15");
      expect(keys(`Born 12${x}3/15/1985`), JSON.stringify(x)).toContain("1985-03-15");
      expect(findDates(`Born 15${x}03${x}1985`).filter((f) => f.weak).flatMap((f) => f.dates.map(dateKey)), JSON.stringify(x)).toContain("1985-03-15");
    }
    const s = spec(range([["Name", "Born"], ["Maria Lopez", 31121]], { formats: { 1: "m/d/yyyy" }, text: { 1: { 1: "3/15/1985" } } }), { B: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    expect(substituteUserText("Born 3/15/1985\u200b5?", [s], map)).toBe(`Born ${map.peek("3/15/1985")!}\u200b5?`);
  });

  it("still finds a value with one inside it, a form feed or a vertical tab too", () => {
    for (const x of hidden) {
      const map = new StandInMap();
      const s = customers();
      const [maria, a] = ["Maria Lopez", "93496"].map((v) => substituteUserText(v, [s], map));
      expect(substituteUserText(`ZIP 934${x}96?`, [s], map), JSON.stringify(x)).toBe(`ZIP ${a}?`);
      expect(substituteUserText(`Who is Ma${x}ria Lopez?`, [s], map), JSON.stringify(x)).toBe(`Who is ${maria}?`);
    }
  });
});

describe("space-grouped numbers in a list (ninth review)", () => {
  it("replaces a private number grouped with a space whatever number is next to it", () => {
    const s = spec(range([["Employee", "Salary"], ["Maria Lopez", 83500], ["Ana Ruiz", 390000]], { text: { 1: { 1: "$83,500" }, 2: { 1: "$390,000" } }, formats: { 1: "$#,##0" } }), {
      B: { private: true, treatment: "stand_in" },
    });
    const map = new StandInMap();
    const [a, b] = ["83500", "390000"].map((v) => substituteUserText(v, [s], map));
    expect(substituteUserText("Which exist?\n40 000\n83 500\n120 000", [s], map)).toBe(`Which exist?\n40 000\n${a}\n120 000`);
    expect(substituteUserText("Bands: 40 000 83 500 120 000", [s], map)).toBe(`Bands: 40 000 ${a} 120 000`);
    expect(substituteUserText("Paid 1 83\u00a0500?", [s], map)).toBe(`Paid 1 ${a}?`);
    expect(substituteUserText("Total 5 390 000", [s], map)).toBe(`Total 5 ${b}`);
    // A digit right next to it still continues it; so does "," or "." with 3 digits.
    for (const q of ["Total 183 500", "Total 83 5001", "Total 5,390,000 and 5.390.000"]) expect(substituteUserText(q, [s], map), q).toBe(q);
  });
});

describe("8-digit columns with placeholders (ninth review)", () => {
  const born = (placeholder: CellValue, k: number, asText = true) => {
    const rows: CellValue[][] = [["Name", "Born"]];
    for (let i = 0; i < 40; i++) {
      const v = `${1950 + i}0${1 + (i % 9)}${10 + (i % 18)}`;
      rows.push([`Person ${i}`, i >= 40 - k ? placeholder : asText ? v : Number(v)]);
    }
    return spec(range(rows), { B: { private: true, treatment: "stand_in" } });
  };

  it("reads a column of dates of birth with a few placeholders for unknown dates", () => {
    for (const [ph, asText] of [["00000000", true], [99999999, false], ["19000000", true], ["20240100", true]] as [CellValue, boolean][]) {
      for (const k of [3, 8, 20]) {
        const s = born(ph, k, asText);
        expect(privateLooseDates(s).get("B"), `${ph} ${k}`).toContain("1953-04-13");
        const map = new StandInMap();
        expect(substituteUserText("Who was born 1953-04-13?", [s], map), `${ph} ${k}`).toBe(`Who was born ${map.peek("19530413")!}?`);
      }
    }
  });

  it("reads a column where at least half are dates, and not one where fewer are", () => {
    // Order IDs 20240115 ..., 7 of 8 of them dates: read as dates (as round 7 did).
    const ids = spec(range([["Order ID", "Customer"], ...["20240115", "20240116", "20231299", "20240301", "20240302", "20240303", "20240304", "20240305"].map((v, i) => [v, `C${i}`])]), {
      A: { private: true, treatment: "stand_in" },
    });
    const map = new StandInMap();
    expect(substituteUserText("Orders placed on 1/15/2024?", [ids], map)).toBe(`Orders placed on ${map.peek("20240115")!}?`);
    // 3 of 8 dates, the rest no dates and not placeholders: not a column of dates.
    const other = spec(range([["Lot", "Qty"], ...["20240115", "20240116", "20240301", "20241399", "20241398", "20241397", "20241396", "20241395"].map((v, i): CellValue[] => [v, i])]), {
      A: { private: true, treatment: "stand_in" },
    });
    expect(privateLooseDates(other).get("A")).toEqual([]);
    expect(substituteUserText("Lots made on 1/15/2024?", [other], new StandInMap())).toBe("Lots made on 1/15/2024?");
  });
});

describe("codes: whole words, not across ranges or stand-ins (ninth review)", () => {
  const bins = () =>
    spec(range([["Bin", "Member", "Qty"], ["A10-B20", "A10023", 5], ["C12-D22", "A10150", 7]]), {
      A: { private: true, treatment: "stand_in" },
      B: { private: true, treatment: "stand_in" },
    });

  it("doesn't read the colon of a cell range as a code's separator", () => {
    expect(isCodeSeparatorCode(0x3a)).toBe(false);
    for (const q of ["Sum of A10:B20", "What is the total of C12:D22?", "=SUM(A10:B20)", "=SUM(Stock!A10:B20)", "Qty in A10\uff1aB20?"]) {
      expect(substituteUserText(q, [bins()], new StandInMap()), q).toBe(q);
    }
    // Any other colon still is one.
    const s = spec(range([["Driver", "Plate"], ["Maria Lopez", "AB-1234-CD"]]), { B: { private: true, treatment: "stand_in" } });
    const map = new StandInMap();
    const plate = substituteUserText("AB-1234-CD", [s], map);
    for (const q of ["AB:1234:CD", "AB\uff1a1234\uff1aCD", "AB12:34CD"]) expect(substituteUserText(`Who has ${q}?`, [s], map), q).toBe(`Who has ${plate}?`);
  });

  it("replaces a code only as a whole word", () => {
    const map = new StandInMap();
    const s = bins();
    const member = substituteUserText("A10023", [s], map);
    expect(substituteUserText("Members in area 10023?", [s], map)).not.toContain("are" + member);
    expect(substituteUserText("Members like A 10023?", [s], map)).toBe(`Members like ${member}?`);
    expect(substituteUserText("Members like A\u200b10023\u200bx?", [s], map)).toBe(`Members like ${member}\u200bx?`);
  });

  it("replaces an aliased header typed without its separators", () => {
    const s = spec(range([["Employee", "Falcon X9-2024"], ["Maria Lopez", 100]]), { B: { private: false, treatment: "as_is", alias: "Project" } });
    for (const q of ["Budget for FalconX92024?", "Budget for falconx92024?", "Budget for Falcon X9-2024?"]) {
      expect(substituteUserText(q, [s], new StandInMap()), q).toBe("Budget for Project?");
    }
  });
});

// ---------------------------------------------------------------------------------------------
// Tenth review

describe("an 8-digit value typed as a date (tenth review)", () => {
  // Invoice numbers 20240001 ... 20240200: only now and then a date, so the column isn't read as dates.
  const invoices = (asText: boolean) =>
    spec(range([["Invoice no", "Amount"], ...Array.from({ length: 200 }, (_, i): CellValue[] => [asText ? String(20240001 + i) : 20240001 + i, 100 + i])]), {
      A: { private: true, treatment: "stand_in" },
    });

  it("replaces the invoice number typed as a date, a character that can't be seen after it or not", () => {
    for (const asText of [false, true]) {
      const s = invoices(asText);
      const map = new StandInMap();
      const id = substituteUserText("20240115", [s], map);
      expect(id, String(asText)).toMatch(/^ID_\d{3}$/u);
      for (const t of ["2024-01-15", "2024/01/15", "2024.01.15"]) expect(substituteUserText(`Total for invoice ${t}?`, [s], map), t).toBe(`Total for invoice ${id}?`);
      for (const mark of ["\u200b", "\u200e", "\u2060"]) {
        expect(substituteUserText(`Total for invoice 2024-01-15${mark}?`, [s], map), JSON.stringify(mark)).toBe(`Total for invoice ${id}${mark}?`);
      }
      // A date whose digits are no invoice's stays.
      expect(substituteUserText("Total after 1/1/2024", [s], map)).toBe("Total after 1/1/2024");
    }
  });
});

describe("a plain number with a digit next to it (tenth review)", () => {
  const s = spec(range([["Customer", "ZIP", "Account"], ["Maria Lopez", 93496, 483920], ["Ana Ruiz", 94103, 571046]]), {
    B: { private: true, treatment: "range" },
    C: { private: true, treatment: "range" },
  });

  it("replaces only a clean occurrence, never part of a longer number", () => {
    const map = new StandInMap();
    for (const q of ["How many orders came from 934961234?", "Payments for reference 4839201?", "Payments for 48392001?", "REF4839201"]) {
      expect(substituteUserText(q, [s], map), q).toBe(q);
    }
    const zip = substituteUserText("93496", [s], map);
    expect(substituteUserText("From 93496-1234 or 934961234?", [s], map)).toBe(`From ${zip}-1234 or 934961234?`);
  });
});
