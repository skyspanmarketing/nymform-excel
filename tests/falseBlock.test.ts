// False blocks found in real use, pinned with the real Flow, builder and auditor.
//
// Jorge's first live session (Orders!B1:G1401 of the playground, column B "Order date"): he typed
// the header again under "Header sent" in another case or spacing. The check then treated
// "Order date" as a hidden header and found it in the header Nymform itself sends for B, so every
// request blocked with [B exact, B word], in both modes, with B private or not.
import { describe, expect, it } from "vitest";
import { parseLocalRange } from "../src/core/a1";
import { aliasRevealsHeader, substituteUserText, StandInMap } from "../src/core/transform";
import type { CellValue, RangeData } from "../src/core/types";
import { Flow, type HostAdapter } from "../src/taskpane/flow";

const MODEL = "openai/gpt-6-luna";

const NAMES = ["Felix Bianchi", "Ana Ruiz", "Kenji Watanabe", "Priya Natarajan", "Omar Haddad", "Lena Fischer", "Tomas Novak", "Ines Duarte"];
const REGIONS = ["East", "West", "Central", "South"];
const PRODUCTS = ["Desk lamp", "Office chair", "Standing desk", "Monitor arm"];

/** Orders!B1:G31 shaped like the playground: B dates (serials shown m/d/yyyy), C-G text. */
function orders(): RangeData {
  const header: CellValue[] = ["Order date", "Customer ID", "Customer", "Region", "SKU", "Product"];
  const rows: CellValue[][] = [header];
  const text: string[][] = [header.map(String)];
  for (let i = 0; i < 30; i++) {
    const serial = 45658 + i * 3; // 1/1/2025 onwards
    const day = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
    const shown = `${day.getUTCMonth() + 1}/${day.getUTCDate()}/${day.getUTCFullYear()}`;
    const c = i % NAMES.length;
    const row: CellValue[] = [serial, `C-${10001 + c}`, NAMES[c]!, REGIONS[i % REGIONS.length]!, `P-${101 + (i % 4)}`, PRODUCTS[i % 4]!];
    rows.push(row);
    text.push([shown, ...row.slice(1).map(String)]);
  }
  const rect = parseLocalRange("B1:G31")!;
  return {
    sheet: "Orders",
    address: "B1:G31",
    rowIndex: rect.r1,
    columnIndex: rect.c1,
    values: rows,
    text,
    formulas: rows.map((r) => [...r]),
    valueTypes: rows.map((r, ri) => r.map((v) => (ri > 0 && typeof v === "number" ? "Double" : "String"))),
    numberFormat: rows.map((r, ri) => r.map((_, ci) => (ri > 0 && ci === 0 ? "m/d/yyyy" : "General"))),
  };
}

function hostFor(data: RangeData): HostAdapter {
  return {
    readSelection: async () => data,
    listSheets: async () => [data.sheet],
    listTables: async () => [],
    readRange: async () => {
      throw new Error("no such range");
    },
    insertFormula: async () => ({ address: "", excelError: false }),
    isRangeEmpty: async () => true,
    isTooLarge: () => false,
  };
}

async function flowFor(data: RangeData) {
  const flow = new Flow({ adapter: hostFor(data), send: async () => ({ ok: false, message: "no network in tests" }) as never, model: MODEL, host: "openrouter.ai", build: "test" });
  expect((await flow.refresh()).ok).toBe(true);
  // As in Jorge's session: Customer ID and Customer kept private.
  flow.setPolicy(0, "C", { private: true, treatment: "stand_in" });
  flow.setPolicy(0, "D", { private: true, treatment: "stand_in" });
  return flow;
}

function hitsOf(flow: Flow, question: string, mode: "structure_only" | "substituted") {
  const r = flow.prepare(question, mode);
  if (!r.ok) throw new Error(r.message);
  const o = r.prepared.outcome;
  return o.ok ? [] : o.result.hits.map((h) => `${h.column} ${h.variant}`);
}

const QUESTIONS = ["What did Felix Bianchi buy and when", "what did buy and when", "Total orders per Region by Order date"];

describe("Jorge's column B block: the header typed again under Header sent", () => {
  for (const alias of ["Order Date", "Order date ", " order date", "Order  date", "ORDER DATE", "Order date (month)", "Month"]) {
    for (const ticked of [true, false]) {
      it(`passes with alias ${JSON.stringify(alias)}, B ${ticked ? "private (month)" : "not private"}`, async () => {
        const flow = await flowFor(orders());
        if (ticked) flow.setPolicy(0, "B", { private: true });
        flow.setPolicy(0, "B", { alias });
        for (const q of QUESTIONS) {
          expect(hitsOf(flow, q, "structure_only"), q).toEqual([]);
          expect(hitsOf(flow, q, "substituted"), q).toEqual([]);
        }
      });
    }
  }

  it("still hides a header renamed to something else: a note naming it blocks", async () => {
    const flow = await flowFor(orders());
    flow.setPolicy(0, "B", { alias: "Field B" });
    expect(hitsOf(flow, "Total per Region", "structure_only")).toEqual([]);
    // The question is rewritten to the alias, so naming the column there is fine.
    expect(hitsOf(flow, "Total per Region by order date", "structure_only")).toEqual([]);
    flow.setPolicy(0, "E", { note: "sorted by order date" });
    expect(hitsOf(flow, "Total per Region", "structure_only")).toEqual(["B exact", "B word"]);
  });

  it("an alias with the header inside a longer word still hides it, and that alias blocks by itself", async () => {
    const data = orders();
    data.values[0]![3] = "Race";
    data.text[0]![3] = "Race";
    const flow = await flowFor(data);
    // "Tracer" doesn't show "Race" as a word, so "Race" stays hidden, and the scan finds it inside
    // the alias Nymform sends. Fail-closed: pick a name without the header in it.
    flow.setPolicy(0, "E", { alias: "Tracer" });
    expect(hitsOf(flow, "Total per Region", "structure_only")).toEqual(["E exact"]);
  });

  it("a genuine rename keeps the header hidden: the question is rewritten, a note naming it blocks", async () => {
    const data = orders();
    data.values[0]![3] = "Race";
    data.text[0]![3] = "Race";
    const flow = await flowFor(data);
    flow.setPolicy(0, "E", { alias: "Group" });
    expect(hitsOf(flow, "Count by race", "structure_only")).toEqual([]);
    flow.setPolicy(0, "F", { note: "grouped by race" });
    expect(hitsOf(flow, "Count per Group", "structure_only")).toEqual(["E exact"]);
  });

  it("a rename in another script stays hidden when it differs by a tone mark, vowel sign or dakuten", async () => {
    for (const [header, alias] of [
      ["ビザ番号", "ピザ番号"],
      ["ข่าว", "ขาว"],
      ["बुखार", "बखार"],
      ["बीमार", "बीमारी"],
    ] as const) {
      const data = orders();
      data.values[0]![3] = header;
      data.text[0]![3] = header;
      const flow = await flowFor(data);
      flow.setPolicy(0, "E", { alias });
      flow.setPolicy(0, "F", { note: `${header} 2024` });
      expect(hitsOf(flow, "Total per Region", "structure_only"), `${header} / ${alias}`).toContain("E exact");
    }
  });
});

describe("aliasRevealsHeader", () => {
  it("is true when the alias shows the header: another case, spacing, width or accents, or the header's words whole", () => {
    for (const alias of ["Order Date", "Order date ", "  order   date", "ＯＲＤＥＲ ＤＡＴＥ", "Order date (month)", "Month of order date", "Order​ date"]) {
      expect(aliasRevealsHeader(alias, "Order date"), alias).toBe(true);
    }
    expect(aliasRevealsHeader("Cafe", "Café")).toBe(true);
    expect(aliasRevealsHeader("Straße", "Straße")).toBe(true);
    expect(aliasRevealsHeader("Tracer race", "Race")).toBe(true);
    expect(aliasRevealsHeader("Last update date", "Date")).toBe(true);
    // Typed again unchanged, even a header of marks or characters that can't be seen.
    expect(aliasRevealsHeader("\u0e48\u0e49\u0e4a\u0e4b", "\u0e48\u0e49\u0e4a\u0e4b")).toBe(true);
    expect(aliasRevealsHeader("\u200b\u200b", "\u200b\u200b")).toBe(true);
    // Folded like the scan: a backtick in an ASCII header, a next-line character at its end.
    expect(aliasRevealsHeader("Driver`s id №", "Driver`s id")).toBe(true);
    expect(aliasRevealsHeader("Order Date", "Order date\u0085")).toBe(true);
    // A lone low surrogate isn't part of a pair with the letter before it: it separates words, for
    // the scan too, which finds "race" in this alias.
    expect(aliasRevealsHeader("\u03c9\udc28Race", "Race")).toBe(true);
  });

  it("is false when the header is hidden: another name, a part of it, or the header inside a longer word", () => {
    for (const [alias, header] of [
      ["Month", "Order date"],
      ["Order", "Order date"],
      ["Tracer", "Race"],
      ["Races", "Race"],
      ["Orderdate", "Order date"],
      ["Field F", "Diagnosis"],
      ["", "Diagnosis"],
      // The header inside a longer word, after a letter, an astral letter or a mark.
      ["Embrace", "Race"],
      ["Basesalary", "Salary"],
      ["\u{10330}Race", "Race"],
      ["\u{20BB7}Race", "Race"],
      ["Race\u064eball", "Race"],
      ["बीमारी", "बीमार"],
      // Other words, not other spellings: a dakuten, a Thai tone mark, an Indic vowel sign.
      ["ピザ番号", "ビザ番号"],
      ["ขาว", "ข่าว"],
      ["बखार", "बुखार"],
      ["カン検診", "ガン検診"],
    ] as const) {
      expect(aliasRevealsHeader(alias, header), `${alias} / ${header}`).toBe(false);
    }
  });

  it("keeps the question as typed for a header the alias shows, and rewrites a hidden one", () => {
    const data = orders();
    const flowSpec = (alias: string) => ({
      data,
      hasHeaders: true,
      columns: ["B", "C", "D", "E", "F", "G"].map((letter, i) => ({ letter, header: String(data.values[0]![i]), type: "text" as const, stats: { blank: 0, distinct: 1 } })),
      policies: ["B", "C", "D", "E", "F", "G"].map((letter) => ({ letter, private: false, treatment: "as_is" as const, ...(letter === "B" ? { alias } : {}) })),
    });
    expect(substituteUserText("Count per Order date", [flowSpec("Order Date")], new StandInMap())).toBe("Count per Order date");
    expect(substituteUserText("Count per Order date", [flowSpec("Field B")], new StandInMap())).toBe("Count per Field B");
  });
});
