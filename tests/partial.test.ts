import { describe, expect, it } from "vitest";
import { MAX_CANDIDATES, applyChoices, findPartialMatches, foldKey } from "../src/core/partial";
import { inferSchema } from "../src/core/schema";
import { StandInMap, substituteUserText } from "../src/core/transform";
import type { CellValue, ColumnPolicy, RangeData, RangeSpec } from "../src/core/types";

function range(rows: CellValue[][], sheet = "Orders", rowIndex = 0): RangeData {
  return {
    sheet,
    address: `A${rowIndex + 1}:${String.fromCharCode(64 + (rows[0]?.length ?? 1))}${rowIndex + rows.length}`,
    rowIndex,
    columnIndex: 0,
    values: rows,
    text: rows.map((r) => r.map((v) => (v === null ? "" : String(v)))),
    formulas: rows.map((r) => [...r]),
    valueTypes: rows.map((r) => r.map((v) => (v === null ? "Empty" : typeof v === "number" ? "Double" : "String"))),
    numberFormat: rows.map((r) => r.map(() => "General")),
  };
}

/** Columns whose letter is listed are private (stand-in); the rest are sent as they are. */
function spec(data: RangeData, privateLetters: string[]): RangeSpec {
  const columns = inferSchema(data, true);
  return {
    data,
    hasHeaders: true,
    columns,
    policies: columns.map((c): ColumnPolicy => ({
      letter: c.letter,
      private: privateLetters.includes(c.letter),
      treatment: privateLetters.includes(c.letter) ? "stand_in" : "as_is",
    })),
  };
}

const ROWS: CellValue[][] = [
  ["Customer", "Email", "Phone", "City", "Region", "Amount"],
  ["Felix Bianchi", "felix.b@example.com", "(831) 555-0199", "Boston", "East", 120],
  ["Felix Kim", "fkim@example.org", "(415) 555-0142", "Kansas City", "Central", 80],
  ["Rosa Felix", "sunflower77@example.net", "(212) 555-0177", "Denver", "West", 60],
  ["Felix Bianchi", "felix.b@example.com", "(831) 555-0199", "Boston", "East", 45],
  ["José Muñoz", "jose@example.com", "(303) 555-0101", "Denver", "West", 30],
  ["Siobhán O'Brien", "sob@example.com", "(617) 555-0120", "Boston", "East", 25],
];
const SPEC = spec(range(ROWS), ["A", "B", "C", "D"]);

function find(question: string, map = new StandInMap()) {
  const sent = substituteUserText(question, [SPEC], map);
  return { sent, matches: findPartialMatches(sent, [SPEC], map), map };
}

describe("partial values: finding them", () => {
  it("lists every private value a first name could be, most frequent first, with its cells and stand-in", () => {
    const { sent, matches, map } = find("Orders from Felix last month");
    expect(sent).toBe("Orders from Felix last month");
    expect(matches).toHaveLength(1);
    const m = matches[0]!;
    expect(m).toMatchObject({ fragment: "Felix", key: "felix", spans: [[12, 17]], more: 0, autoMap: false });
    expect(m.candidates.map((c) => [c.value, c.count, c.places[0]!.cells])).toEqual([
      ["Felix Bianchi", 2, ["A2", "A5"]],
      ["Felix Kim", 1, ["A3"]],
      ["Rosa Felix", 1, ["A4"]],
    ]);
    // The stand-ins are the ones substitution uses for the whole values.
    for (const c of m.candidates) {
      expect(c.token).toMatch(/^PERSON_\d{3}$/u);
      expect(map.valueOf(c.token)).toBe(c.value);
      expect(substituteUserText(`Total for ${c.value}`, [SPEC], map)).toBe(`Total for ${c.token}`);
    }
    expect(m.candidates[0]!.places[0]).toMatchObject({ sheet: "Orders", column: "A", header: "Customer" });
  });

  it("marks one candidate that names a person, email or ID for mapping without asking, but not free text", () => {
    const [bianchi] = find("Orders from Bianchi").matches;
    expect(bianchi).toMatchObject({ fragment: "Bianchi", autoMap: true });
    expect(bianchi!.candidates.map((c) => c.value)).toEqual(["Felix Bianchi"]);
    const [kansas] = find("How many orders from Kansas?").matches;
    expect(kansas!.candidates.map((c) => c.value)).toEqual(["Kansas City"]);
    expect(kansas!.autoMap).toBe(false);
  });

  it("finds an email by the part before the @, and not by its domain", () => {
    const [m] = find("Orders for sunflower77").matches;
    expect(m!.candidates.map((c) => c.value)).toEqual(["sunflower77@example.net"]);
    expect(find("Orders from example customers").matches).toEqual([]);
  });

  it("finds a phone by a group of its digits, but not by a year or a short group", () => {
    const [m] = find("Orders from the phone ending 0199").matches;
    expect(m!.candidates.map((c) => [c.value, c.count])).toEqual([["(831) 555-0199", 2]]);
    expect(find("Orders in 2026 from area 555").matches).toEqual([]);
  });

  it("folds case and accents, and treats O'Brien as one part", () => {
    expect(find("orders for jose").matches[0]!.candidates.map((c) => c.value)).toEqual(["José Muñoz"]);
    const obrien = find("Orders for O'Brien").matches;
    expect(obrien.map((m) => m.fragment)).toEqual(["O'Brien"]);
    expect(obrien[0]!.candidates.map((c) => c.value)).toEqual(["Siobhán O'Brien"]);
    expect(foldKey("Siobhán")).toBe("siobhan");
  });

  it("finds nothing for whole values (substitution replaced them), stand-ins, short words or unmarked columns", () => {
    expect(find("Total Amount for Felix Bianchi").matches).toEqual([]);
    expect(find("Compare PERSON_001 and Kim").matches).toEqual([]);
    expect(find("Orders in the East region").matches).toEqual([]);
    const allOpen = spec(range(ROWS), []);
    expect(findPartialMatches("Orders from Felix", [allOpen], new StandInMap())).toEqual([]);
  });

  it("doesn't read a word of a column's name as a part of a value (Unit price vs Jordan Price)", () => {
    const rows: CellValue[][] = [
      ["Customer", "Unit price", "Region"],
      ["Jordan Price", 12.5, "East"],
      ["Ana Ruiz", 8, "West"],
    ];
    const s = spec(range(rows), ["A"]);
    expect(findPartialMatches("Average Unit price per Region", [s], new StandInMap())).toEqual([]);
    expect(findPartialMatches("Average UNIT  PRICE per Region", [s], new StandInMap())).toEqual([]);
    // On its own, "Price" is still read as the customer.
    const [price] = findPartialMatches("Orders from Price, by Unit price", [s], new StandInMap());
    expect(price!.spans).toEqual([[12, 17]]);
    expect(price!.candidates.map((c) => c.value)).toEqual(["Jordan Price"]);
    // A hidden header was replaced by its alias in the question: the alias is the column's name.
    const aliased: RangeSpec = { ...s, policies: s.policies.map((p) => (p.letter === "B" ? { ...p, alias: "Price each" } : p)) };
    expect(findPartialMatches("Average Price each per Region", [aliased], new StandInMap())).toEqual([]);
  });

  it("masks column names with any characters as whole words only, and survives very long ones", () => {
    const rows: CellValue[][] = [
      ["Customer", "Cost (USD", "Price [EUR", "Growth +/-", "*Notes", "Unit", "Rice"],
      ["Ana Ruiz", 1, 2, 3, "x", 4, 5],
      ["Lena Unity", 1, 2, 3, "y", 4, 5],
      ["Jordan Price", 1, 2, 3, "z", 4, 5],
    ];
    const s = spec(range(rows), ["A"]);
    for (const h of ["Cost (USD", "Price [EUR", "Growth +/-", "*Notes"]) {
      expect(findPartialMatches(`Total ${h} for Ruiz`, [s], new StandInMap()).map((m) => m.fragment), h).toEqual(["Ruiz"]);
    }
    // "Unit" and "Rice" are column names; "Unity" and "Price" aren't them.
    expect(findPartialMatches("Orders of Unity", [s], new StandInMap()).map((m) => m.fragment)).toEqual(["Unity"]);
    expect(findPartialMatches("Orders of Price", [s], new StandInMap()).map((m) => m.fragment)).toEqual(["Price"]);
    // A name glued after a letter isn't masked there, but the next whole occurrence is, even when
    // it starts inside the rejected one, or with a character outside the basic plane.
    const glued = spec(range([["Customer", "Bora Bora", "𠮷野家 sales", "Region"], ["Ana Bora", 1, 2, "East"], ["Lena Ruiz", 1, 2, "West"]]), ["A"]);
    expect(findPartialMatches("Sales in xBora Bora Bora by Region", [glued], new StandInMap())).toEqual([]);
    expect(findPartialMatches("Total x𠮷野家 sales for Ruiz", [glued], new StandInMap()).map((m) => m.fragment)).toEqual(["Ruiz"]);
    expect(findPartialMatches("Total x𠮷野家 sales, 𠮷野家 sales for Ruiz", [glued], new StandInMap()).map((m) => m.fragment)).toEqual(["Ruiz"]);
    // Spacing inside a name doesn't matter, and 3-character names aren't looked for.
    const spaced = spec(range([["Customer", "Unit      price"], ["Jordan Price", 1]]), ["A"]);
    expect(findPartialMatches("Average Unit price per Region", [spaced], new StandInMap())).toEqual([]);
    expect(findPartialMatches("Unit price", [spaced], new StandInMap())).toEqual([]);
    const short = spec(range([["Customer", "Ana"], ["Bo-Ana Li", 1], ["Lena Ana-Maria", 2]]), ["A"]);
    expect(findPartialMatches("Sales for Bo-Ana", [short], new StandInMap()).map((m) => m.fragment)).toEqual(["Bo-Ana"]);
    expect(findPartialMatches("Orders from Ana-Maria", [short], new StandInMap()).map((m) => m.fragment)).toEqual(["Ana-Maria"]);
    // A first-row cell can hold a whole paragraph (32,767 characters).
    const long = spec(range([["Customer", "a ".repeat(6000).trim()], ["Ana Ruiz", 1]]), ["A"]);
    expect(findPartialMatches("Orders from Ruiz", [long], new StandInMap()).map((m) => m.fragment)).toEqual(["Ruiz"]);
    const question = `${"a ".repeat(6000)}Ruiz`;
    expect(findPartialMatches(question, [long], new StandInMap()).map((m) => m.fragment)).toEqual(["Ruiz"]);
  });

  it("lists several parts in the order they appear, each with every place it occurs", () => {
    const { matches } = find("Felix or Muñoz, and Felix again");
    expect(matches.map((m) => [m.fragment, m.spans])).toEqual([
      ["Felix", [[0, 5], [20, 25]]],
      ["Muñoz", [[9, 14]]],
    ]);
  });

  it("counts candidates beyond the list limit, and keeps cell addresses on the sheet's rows", () => {
    const many: CellValue[][] = [["Customer"], ...Array.from({ length: MAX_CANDIDATES + 5 }, (_, i) => [`Ana Surname${i}`])];
    const s = spec(range(many, "Crm", 9), ["A"]);
    // The range starts on row 10: its header is row 10, so data row 3 is sheet row 14.
    const [ana] = findPartialMatches("Orders from Surname3", [s], new StandInMap());
    expect(ana!.candidates.map((c) => [c.value, c.places[0]!.cells])).toEqual([["Ana Surname3", ["A14"]]]);
    expect(findPartialMatches("Orders for Anas", [s], new StandInMap())).toEqual([]);
    const shared: CellValue[][] = [["Customer"], ...Array.from({ length: MAX_CANDIDATES + 5 }, (_, i) => [`Given${i} Shared`])];
    const [big] = findPartialMatches("Orders for Shared", [spec(range(shared), ["A"])], new StandInMap());
    expect(big!.candidates).toHaveLength(MAX_CANDIDATES);
    expect(big!.more).toBe(5);
  });
});

describe("partial values: applying a choice", () => {
  it("replaces every occurrence with the chosen stand-ins, joined with or", () => {
    const { sent, matches } = find("Felix orders, then Felix returns");
    const [m] = matches;
    const [bianchi, kim] = m!.candidates;
    expect(applyChoices(sent, matches, new Map([["felix", [bianchi!.token]]]))).toBe(`${bianchi!.token} orders, then ${bianchi!.token} returns`);
    expect(applyChoices(sent, matches, new Map([["felix", [bianchi!.token, kim!.token]]]))).toBe(
      `${bianchi!.token} or ${kim!.token} orders, then ${bianchi!.token} or ${kim!.token} returns`,
    );
  });

  it("leaves a part as typed when nothing, or a stand-in that isn't a candidate, is chosen", () => {
    const { sent, matches } = find("Orders from Felix");
    expect(applyChoices(sent, matches, new Map([["felix", []]]))).toBe(sent);
    expect(applyChoices(sent, matches, new Map([["felix", ["PERSON_999"]]]))).toBe(sent);
    expect(applyChoices(sent, matches, new Map())).toBe(sent);
  });

  it("keeps the rest of the question, including stand-ins already there", () => {
    const { sent, matches } = find("Felix Bianchi versus Kim");
    expect(sent).toMatch(/^PERSON_\d{3} versus Kim$/u);
    expect(matches).toEqual([]);
    const two = find("Total for Felix Bianchi and Muñoz");
    const [munoz] = two.matches;
    expect(applyChoices(two.sent, two.matches, new Map([[munoz!.key, [munoz!.candidates[0]!.token]]]))).toMatch(/^Total for PERSON_\d{3} and PERSON_\d{3}$/u);
  });
});
