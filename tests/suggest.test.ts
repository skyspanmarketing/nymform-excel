import { describe, expect, it } from "vitest";
import { inferSchema } from "../src/core/schema";
import { formulaRefs, headerKeywords, headerWords, suggestPolicies, type Suggestion } from "../src/core/suggest";
import type { CellValue, RangeData } from "../src/core/types";

function range(
  rows: CellValue[][],
  opts: { formats?: Record<number, string>; formulas?: (string | undefined)[][]; columnIndex?: number } = {},
): RangeData {
  const columnIndex = opts.columnIndex ?? 0;
  return {
    sheet: "Sheet1",
    address: "A1",
    rowIndex: 0,
    columnIndex,
    values: rows,
    text: rows.map((r) => r.map((v) => (v === null ? "" : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v)))),
    formulas: rows.map((r, i) => r.map((v, j) => opts.formulas?.[i]?.[j] ?? v)),
    valueTypes: rows.map((r) =>
      r.map((v) =>
        v === null ? "Empty" : typeof v === "number" ? "Double" : typeof v === "boolean" ? "Boolean" : "String",
      ),
    ),
    numberFormat: rows.map((r) => r.map((_, j) => opts.formats?.[j] ?? "General")),
  };
}

/** A sheet from column-major data: header plus values per column. */
function columns(cols: Record<string, CellValue[]>, opts: Parameters<typeof range>[1] = {}): RangeData {
  const headers = Object.keys(cols);
  const n = Math.max(...Object.values(cols).map((c) => c.length));
  const rows: CellValue[][] = [headers];
  for (let i = 0; i < n; i++) rows.push(headers.map((h) => cols[h]![i] ?? null));
  return range(rows, opts);
}

function suggest(data: RangeData, hasHeaders = true): Record<string, Suggestion> {
  const out: Record<string, Suggestion> = {};
  for (const s of suggestPolicies(data, inferSchema(data, hasHeaders), hasHeaders)) out[s.letter] = s;
  return out;
}

const repeat = <T,>(items: T[], n: number): T[] => Array.from({ length: n }, (_, i) => items[i % items.length]!);
const REGIONS = repeat(["East", "West", "North", "South"], 40);

describe("header words and keywords", () => {
  it("splits on punctuation, camelCase and digits", () => {
    expect(headerWords("Customer_Name")).toEqual(["customer", "name"]);
    expect(headerWords("emailAddress")).toEqual(["email", "address"]);
    expect(headerWords("StudentID")).toEqual(["student", "id"]);
    expect(headerWords("Address2")).toEqual(["address", "2"]);
    expect(headerWords("E-mail")).toContain("email");
  });

  it("matches keywords as whole words only", () => {
    expect(headerKeywords("Customer Name")).toEqual(["customer", "name"]);
    expect(headerKeywords("Customers")).toEqual([]);
    expect(headerKeywords("E-mail")).toEqual(["email"]);
    expect(headerKeywords("Order ID")).toEqual(["id"]);
    expect(headerKeywords("Paid")).toEqual([]);
    expect(headerKeywords("Identity")).toEqual([]);
    expect(headerKeywords("Width")).toEqual([]);
    expect(headerKeywords("Named range")).toEqual([]);
    expect(headerKeywords("Date of Birth")).toEqual(["birth"]);
    expect(headerKeywords("DOB")).toEqual(["dob"]);
    expect(headerKeywords("FirstName")).toEqual(["first", "name"]);
    expect(headerKeywords("firstname")).toEqual(["first"]);
  });
});

describe("suggestPolicies", () => {
  it("suggests private columns from header keywords, with treatments by type", () => {
    const data = columns(
      {
        "Customer Name": repeat(["Ana", "Ben"], 160),
        "E-mail": repeat(["x"], 160),
        Salary: repeat([50000, 60000], 160),
        Birthday: repeat([30000, 31000], 160),
        Region: REGIONS,
        Amount: REGIONS.map((_, i) => (i % 5) + 0.25),
      },
      { formats: { 3: "m/d/yyyy" } },
    );
    const s = suggest(data);
    expect(s.A).toMatchObject({ private: true, treatment: "stand_in" });
    expect(s.A!.reasons[0]).toMatch(/^Header mentions (customer, )?name$/);
    expect(s.B).toMatchObject({ private: true, treatment: "stand_in" });
    expect(s.C).toMatchObject({ private: true, treatment: "range" });
    expect(s.D).toMatchObject({ private: true, treatment: "month" });
    expect(s.E).toMatchObject({ private: false, treatment: "as_is", reasons: [] });
    expect(s.F).toMatchObject({ private: false, treatment: "as_is" });
  });

  it("ignores header keywords when the range has no header row", () => {
    const data = range([["Name"], ["East"], ["West"], ["East"], ["West"], ["East"]]);
    // Private because there is no header row, not because row 1 says "Name".
    expect(suggest(data, false).A!.private).toBe(true);
    expect(suggest(data, false).A!.reasons.some((r) => r.startsWith("Header mentions"))).toBe(false);
  });

  it("finds emails, phones, 9-digit numbers and unique IDs by value", () => {
    const n = 60;
    const data = columns({
      Contact: Array.from({ length: n }, (_, i) => `user${i % 3}@example.test`),
      Tel: Array.from({ length: n }, (_, i) => `(831) 555-${String(1000 + (i % 4)).padStart(4, "0")}`),
      Ref: Array.from({ length: n }, (_, i) => `123-45-${String(6000 + (i % 3))}`),
      Code: Array.from({ length: n }, (_, i) => `AB${10000 + i}`),
      Status: repeat(["open", "closed"], n),
    });
    const s = suggest(data);
    expect(s.A).toMatchObject({ private: true });
    expect(s.A!.reasons).toContain("Values contain email addresses");
    expect(s.B!.reasons).toContain("Values look like phone numbers");
    expect(s.C!.reasons).toContain("Values look like 9-digit ID numbers");
    expect(s.D!.reasons).toContain("Values look like unique IDs");
    expect(s.E!.private).toBe(false);
  });

  it("does not take unique decimal amounts or dates for IDs", () => {
    const data = columns(
      {
        Total: Array.from({ length: 100 }, (_, i) => 10000.5 + i * 3.25),
        "Order date": Array.from({ length: 100 }, (_, i) => 46000 + i),
        Shipped: Array.from({ length: 100 }, (_, i) => `2026-01-${String((i % 28) + 1).padStart(2, "0")}`),
      },
      { formats: { 1: "m/d/yyyy" } },
    );
    data.text = data.text.map((row, r) => row.map((t, c) => (c === 1 && r > 0 ? `1/${r}/2026` : t)));
    const s = suggest(data);
    expect(s.A!.private).toBe(false);
    expect(s.B!.private).toBe(false);
    expect(s.C!.reasons).not.toContain("Values look like unique IDs");
  });

  it("checks value patterns on at most 200 non-empty cells", () => {
    const vals: CellValue[] = [...repeat(["open", "closed"], 200), "someone@example.test"];
    const data = columns({ Status: vals });
    expect(suggest(data).A!.private).toBe(false);
    const early = columns({ Status: [null, null, "someone@example.test", ...repeat(["open", "closed"], 300)] });
    expect(suggest(early).A!.private).toBe(true);
  });

  it("keeps free text private and excluded from sample rows", () => {
    const long = "the customer called twice about the late delivery and asked for a refund";
    const data = columns({ Remarks: repeat([long, `${long} again`], 30), Region: repeat(["East", "West"], 30) });
    const s = suggest(data);
    expect(s.A).toMatchObject({ private: true, treatment: "exclude" });
    expect(s.B!.private).toBe(false);
  });

  it("defaults unsure text columns to private", () => {
    const data = columns({
      Account: Array.from({ length: 100 }, (_, i) => `Acct ${i}`),
      Product: repeat(["Widget", "Gadget", "Gizmo"], 100),
      Mixed: Array.from({ length: 100 }, (_, i) => (i % 2 === 0 ? i : `v${i}`)),
    });
    const s = suggest(data);
    expect(s.A).toMatchObject({ private: true, treatment: "stand_in" });
    expect(s.A!.reasons).toEqual(["Many different values, so it is kept private to be safe"]);
    expect(s.B!.private).toBe(false);
    expect(s.C).toMatchObject({ private: true, treatment: "stand_in" });
  });

  it("suggests a neutral alias for revealing headers", () => {
    const data = columns({
      Region: REGIONS,
      Amount: REGIONS.map((_, i) => i % 3),
      Units: REGIONS.map((_, i) => i % 2),
      Diagnosis: repeat(["Flu", "Cold"], 160),
      Gender: repeat(["F", "M"], 160),
    });
    const s = suggest(data);
    expect(s.D).toMatchObject({ private: true, alias: "Field D" });
    expect(s.E).toMatchObject({ private: true, alias: "Field E" });
    expect(s.A!.alias).toBeUndefined();
  });

  it("T-S2: a column of formulas referencing a private column is suggested private", () => {
    const n = 40;
    const names = Array.from({ length: n }, (_, i) => `Person ${i % 5}`);
    const rows: CellValue[][] = [["Customer Name", "Region", "Label", "Label length", "Region code"]];
    const formulas: (string | undefined)[][] = [[]];
    for (let i = 0; i < n; i++) {
      const r = i + 2;
      const region = REGIONS[i]!;
      rows.push([names[i]!, region, names[i]!.toUpperCase().slice(0, 3), 3, region.length]);
      formulas.push([undefined, undefined, `=LEFT(UPPER(A${r}),3)`, `=LEN(C${r})`, `=LEN(B${r})`]);
    }
    const s = suggest(range(rows, { formulas }));
    expect(s.A!.private).toBe(true);
    expect(s.B!.private).toBe(false);
    // C reads A directly; D reads only C, which becomes private first (fixpoint).
    expect(s.C).toMatchObject({ private: true });
    expect(s.C!.reasons).toContain("Formulas use a private column");
    expect(s.D).toMatchObject({ private: true, treatment: "range" });
    expect(s.E!.private).toBe(false);
  });

  it("follows ranges, whole columns and structured references to private columns", () => {
    const n = 20;
    const rows: CellValue[][] = [["Region", "Salary", "Sum", "Share", "Tag"]];
    const formulas: (string | undefined)[][] = [[]];
    for (let i = 0; i < n; i++) {
      rows.push([REGIONS[i]!, 50000 + (i % 2), 1, 2, 3]);
      formulas.push([undefined, undefined, `=SUMIFS(B:B,A:A,A${i + 2})`, "=[@Salary]/SUM(Staff[Salary])", "=Staff[@Region]"]);
    }
    const data = range(rows, { formulas });
    data.table = { name: "Staff", columns: ["Region", "Salary", "Sum", "Share", "Tag"] };
    const s = suggest(data);
    expect(s.B!.private).toBe(true);
    expect(s.C!.private).toBe(true);
    expect(s.D!.private).toBe(true);
    expect(s.E!.private).toBe(false);
  });
});

describe("formulaRefs", () => {
  it("reads cells, ranges, whole columns, rows and structured references", () => {
    const cols = (f: string) => [...formulaRefs(f).columns].sort((a, b) => a - b);
    expect(cols("=A2*2")).toEqual([0]);
    expect(cols("=SUM($B$2:D9)")).toEqual([1, 2, 3]);
    expect(cols("=SUMIFS(E:E,B:B,B2)")).toEqual([1, 4]);
    expect(cols("=Regions!B2")).toEqual([1]);
    expect(cols("='Q1 data'!C3")).toEqual([2]);
    expect(cols("=LOG10(F2)+ATAN2(1,2)")).toEqual([5]);
    expect(cols('=IF(A2="B7","x","y")')).toEqual([0]);
    expect(formulaRefs("=SUM(2:2)").all).toBe(true);
    expect(formulaRefs("=ROWS(Orders[#All])").all).toBe(true);
    expect(formulaRefs('=INDIRECT("A"&ROW())').all).toBe(true);
    expect(formulaRefs("=A2*2").all).toBe(false);
    expect([...formulaRefs("=[@[Customer Name]]&Orders[[#This Row],[Amount]]").names]).toEqual(["customer name", "amount"]);
  });
});

describe("person headers beyond the spec's keyword list", () => {
  it("suggests Customer, Client name and Instructor private, but not Region or Product", async () => {
    const { suggestPolicies } = await import("../src/core/suggest");
    const { inferSchema } = await import("../src/core/schema");
    const headers = ["Customer", "Client name", "Instructor", "Region", "Product"];
    const rows = [headers, ["A B", "C D", "E F", "West", "Pens"], ["G H", "I J", "K L", "East", "Pens"], ["M N", "O P", "Q R", "West", "Ink"], ["S T", "U V", "W X", "East", "Ink"]];
    const data = {
      sheet: "S", address: "A1:E5", rowIndex: 0, columnIndex: 0,
      values: rows, text: rows, formulas: rows,
      valueTypes: rows.map((r) => r.map(() => "String")),
      numberFormat: rows.map((r) => r.map(() => "General")),
    };
    const suggestions = suggestPolicies(data, inferSchema(data));
    expect(suggestions.map((x) => [x.letter, x.private])).toEqual([["A", true], ["B", true], ["C", true], ["D", false], ["E", false]]);
  });
});

describe("a range without a header row", () => {
  it("starts every non-empty column private, since nothing says what it holds", async () => {
    const { suggestPolicies } = await import("../src/core/suggest");
    const { inferSchema } = await import("../src/core/schema");
    const rows = [["Ann Lee", "West", 10], ["Bo Chan", "East", 20], ["Ann Lee", "West", 30]];
    const data = {
      sheet: "S", address: "A1:C3", rowIndex: 0, columnIndex: 0,
      values: rows, text: rows.map((r) => r.map(String)), formulas: rows,
      valueTypes: rows.map((r) => r.map((v) => (typeof v === "number" ? "Double" : "String"))),
      numberFormat: rows.map((r) => r.map(() => "General")),
    };
    const suggestions = suggestPolicies(data, inferSchema(data, false), false);
    expect(suggestions.map((x) => x.private)).toEqual([true, true, true]);
    expect(suggestions[2]!.treatment).toBe("range");
  });
});

