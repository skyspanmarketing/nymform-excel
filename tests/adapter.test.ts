// Office host adapter (spec §7.1, §2 invariant 9, §4 S4, §11 scratch sheet), run against the fake
// Excel host in tests/e2e/fakeOffice.js, which enforces Office's load/sync rules.
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { runInThisContext } from "node:vm";
import ExcelJS from "exceljs";
import { beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import {
  BENCH_SHEET,
  CELL_LIMIT,
  ExcelHostError,
  NOT_IN_EXCEL_MESSAGE,
  RESOLVE_MESSAGE,
  SCRATCH_SHEET,
  SelectionTooLargeError,
  TOO_LARGE_MESSAGE,
  firstEmptyColumn,
  insertFormula,
  isRangeEmpty,
  listNames,
  listSheets,
  listTables,
  officeTheme,
  readBenchInfo,
  readRange,
  readScratch,
  readSelection,
  readWorkbookInfo,
  ready,
  selectCells,
  selectRange,
  toPlainError,
  writeScratch,
} from "../src/office/adapter";
import { generateData, toRangeData, workbookBuffer } from "../bench/generate";
import { columnLetter, parseCell, parseLocalRange } from "../src/core/a1";
import type { RangeData } from "../src/core/types";
import { displayText, stringifyFake, workbookToFake, xlsxToFake, type FakeWorkbook } from "./e2e/workbookToFake";

// Tests build the release shape (BENCH false); the scratch helpers are exercised as a bench build.
const bench = vi.hoisted(() => ({ on: true }));
vi.mock("../src/config", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../src/config")>();
  return {
    ...actual,
    get BENCH() {
      return bench.on;
    },
  };
});

type Json = Record<string, unknown>;

interface FakeCell {
  value: string | number | boolean | null;
  text: string;
  type: string;
  format: string;
  formula?: string;
  spill?: string;
  spilledFrom?: string;
}

interface FakeHost {
  writes: Json[];
  selections: { sheet: string; address: string }[];
  loads: { type: string; sheet?: string; address?: string; props: string[] }[];
  evaluate: ((formula: string, sheet: string, addr: string) => unknown) | null;
  host: string | null;
  platform: string;
  maxApi: string | null;
  usedRange: "cells" | "sheet";
  theme: { isDarkTheme?: boolean; bodyBackgroundColor?: string } | null;
  reset(workbook?: unknown): void;
  getCell(sheet: string, addr: string): FakeCell | null;
  setCell(sheet: string, addr: string, cell: Partial<FakeCell> | null): void;
  setSelection(sheet: string, address: string): void;
  getSelection(): { sheet: string; address: string };
  setCalculationMode(mode: string): void;
  failNext(code: string, message?: string): void;
  failOnSync(n: number, code: string, message?: string): void;
  workbook(): FakeWorkbook;
  sheetNames(): string[];
  shiftFormula(formula: string, dr: number, dc: number): string;
}

const read = (rel: string) => readFileSync(new URL(rel, import.meta.url), "utf8");
const FAKE_SRC = read("./e2e/fakeOffice.js");
const MINI = JSON.parse(read("./e2e/fixtures/mini.json")) as FakeWorkbook;
const G = globalThis as unknown as { __NYMFORM_FAKE__: FakeHost; Office?: unknown; Excel?: unknown };

let fake: FakeHost;

function withSheet(wb: FakeWorkbook, name: string, cells: Record<string, Partial<FakeCell>>, visibility = "Visible"): FakeWorkbook {
  const copy = structuredClone(wb);
  copy.sheets.push({ name, visibility: visibility as "Visible", cells: cells as FakeWorkbook["sheets"][number]["cells"] });
  return copy;
}

function loadsSince(i: number) {
  return fake.loads.slice(i);
}

/** What readSelection may load before it knows the size: where things are, never what they hold. */
const GEOMETRY_PROPS = [
  "address",
  "rowIndex",
  "columnIndex",
  "rowCount",
  "columnCount",
  "isEntireRow",
  "isEntireColumn",
  "items/name",
  "items/showTotals",
];

/**
 * True when everything loaded since `i` is geometry, except the value types of one clicked row
 * (empty or not), which the adapter reads to find where that row holds anything: one row, at most
 * a full sheet row, and nothing else in that load.
 */
function onlyGeometrySince(i: number): boolean {
  return loadsSince(i).every((l) => {
    if (!l.props.includes("valueTypes")) return l.props.every((p) => GEOMETRY_PROPS.includes(p));
    const rect = l.type === "Range" && l.address ? parseLocalRange(l.address.replace(/^.*!/u, "").replace(/\$/gu, "")) : null;
    return l.props.length === 1 && rect !== null && rect.r1 === rect.r2 && rect.c2 - rect.c1 + 1 <= 16384;
  });
}

function contentLoads(i: number): string[] {
  return loadsSince(i)
    .filter((l) => l.type === "Range")
    .flatMap((l) => l.props)
    .filter((p) => ["values", "text", "formulas", "numberFormat"].includes(p));
}

beforeAll(() => {
  runInThisContext(FAKE_SRC, { filename: "tests/e2e/fakeOffice.js" });
  fake = G.__NYMFORM_FAKE__;
});

beforeEach(() => {
  fake.reset(structuredClone(MINI));
  bench.on = true;
});

// ---------------------------------------------------------------------------------------------

describe("ready", () => {
  it("resolves with the host and platform inside Excel", async () => {
    await expect(ready()).resolves.toEqual({ host: "Excel", platform: "OfficeOnline" });
  });

  it("rejects with a plain message when the host isn't Excel", async () => {
    fake.host = null;
    await expect(ready()).rejects.toThrow(NOT_IN_EXCEL_MESSAGE);
    fake.host = "Word";
    await expect(ready()).rejects.toThrow(NOT_IN_EXCEL_MESSAGE);
  });

  it("rejects when Office.js isn't loaded at all", async () => {
    const office = G.Office;
    delete G.Office;
    try {
      await expect(ready()).rejects.toThrow(NOT_IN_EXCEL_MESSAGE);
    } finally {
      G.Office = office;
    }
  });

  it("rejects when Excel lacks ExcelApi 1.9", async () => {
    fake.maxApi = "1.8";
    await expect(ready()).rejects.toThrow(/too old/);
  });

  it("reports Office's theme when the host has one", () => {
    expect(officeTheme()).toBeNull();
    fake.theme = { isDarkTheme: true, bodyBackgroundColor: "#1f1f1f" };
    expect(officeTheme()).toEqual({ isDarkTheme: true, bodyBackgroundColor: "#1f1f1f" });
  });
});

describe("readSelection", () => {
  it("reads values, text, formulas, value types and formats, with empty cells as null", async () => {
    const d = await readSelection();
    expect(d.sheet).toBe("Orders");
    expect(d.address).toBe("A1:F9");
    expect(d.rowIndex).toBe(0);
    expect(d.columnIndex).toBe(0);
    expect(d.values[0]).toEqual(["Customer", "Email", "Region", "Order date", "Product", "Amount"]);
    expect(d.values[1]).toEqual(["Ana Ruiz", "ana.ruiz@example.com", "East", 46027, "Widget", 1250.5]);
    expect(d.text[1]).toEqual(["Ana Ruiz", "ana.ruiz@example.com", "East", "1/5/2026", "Widget", "1,250.50"]);
    expect(d.valueTypes[1]).toEqual(["String", "String", "String", "Double", "String", "Double"]);
    expect(d.numberFormat[1]).toEqual(["General", "General", "General", "m/d/yyyy", "General", "#,##0.00"]);
    // Dev Patel's email (B5) and Erik Lund's product (E7) are empty.
    expect(d.values[4]![1]).toBeNull();
    expect(d.text[4]![1]).toBe("");
    expect(d.valueTypes[4]![1]).toBe("Empty");
    expect(d.formulas[4]![1]).toBe("");
    expect(d.values[6]![4]).toBeNull();
    expect(d.values).toHaveLength(9);
    for (const grid of [d.values, d.text, d.formulas, d.valueTypes, d.numberFormat]) {
      expect(grid.every((row) => row.length === 6)).toBe(true);
    }
  });

  it("strips the sheet from the address, including quoted sheet names", async () => {
    fake.reset(withSheet(MINI, "Q1 Sales", { B2: { value: "x" }, C3: { value: 2 } }));
    fake.setSelection("Q1 Sales", "B2:C3");
    const d = await readSelection();
    expect(d.sheet).toBe("Q1 Sales");
    expect(d.address).toBe("B2:C3");
    expect(d.rowIndex).toBe(1);
    expect(d.columnIndex).toBe(1);
    expect(d.values).toEqual([
      ["x", null],
      [null, 2],
    ]);
  });

  it("keeps formulas and error cells as Excel reports them", async () => {
    fake.setCell("Orders", "G2", { formula: "=F2/0", value: "#DIV/0!", type: "Error", text: "#DIV/0!" });
    // One row is a pointer; exact reads the cells themselves.
    fake.setSelection("Orders", "F2:G2");
    const d = await readSelection({ exact: true });
    expect(d.formulas).toEqual([[1250.5, "=F2/0"]]);
    expect(d.values).toEqual([[1250.5, "#DIV/0!"]]);
    expect(d.valueTypes).toEqual([["Double", "Error"]]);
  });

  it("detects the table the selection sits in, with its header names", async () => {
    const d = await readSelection();
    expect(d.table).toEqual({ name: "Orders", columns: ["Customer", "Email", "Region", "Order date", "Product", "Amount"] });
  });

  it("aligns table column names with the selected columns", async () => {
    fake.setSelection("Orders", "C3:D5");
    expect((await readSelection()).table).toEqual({ name: "Orders", columns: ["Region", "Order date"] });
    fake.setSelection("Orders", "E1:H3");
    expect((await readSelection()).table).toEqual({ name: "Orders", columns: ["Product", "Amount", "", ""] });
  });

  it("reports no table outside tables", async () => {
    fake.setSelection("Regions", "A1:B5");
    expect((await readSelection()).table).toBeNull();
    fake.setSelection("Orders", "H1:I3");
    expect((await readSelection()).table).toBeNull();
  });

  it(`throws SelectionTooLargeError above ${CELL_LIMIT} cells, before loading any cell content`, async () => {
    fake.setSelection("Orders", "A1:T1001"); // 20 x 1001 = 20,020 cells
    fake.setCell("Orders", "T1001", { value: 1, text: "1", type: "Double" }); // the used range is that large too
    const before = fake.loads.length;
    const err = await readSelection().catch((e: unknown) => e);
    expect(err).toBeInstanceOf(SelectionTooLargeError);
    expect((err as Error).message).toBe(TOO_LARGE_MESSAGE);
    expect((err as SelectionTooLargeError).cellCount).toBe(20020);
    expect(onlyGeometrySince(before)).toBe(true);
    expect(contentLoads(before)).toEqual([]);
  });

  it("accepts exactly 20,000 cells", async () => {
    fake.setSelection("Orders", "A1:T1000");
    const d = await readSelection();
    expect(d.values).toHaveLength(1000);
    expect(d.values[999]).toHaveLength(20);
  });

  it("trims a whole-column selection to the used range", async () => {
    fake.setSelection("Orders", "A:A");
    expect((await readSelection()).address).toBe("A1:A9");
  });
});

describe("listSheets and listTables", () => {
  it("lists every sheet in tab order, hidden ones included", async () => {
    expect(await listSheets()).toEqual(["Orders", "Regions"]);
    fake.reset(withSheet(MINI, BENCH_SHEET, { A1: { value: "NYMFORM_BENCH_V1" } }, "Hidden"));
    expect(await listSheets()).toEqual(["Orders", "Regions", BENCH_SHEET]);
  });

  it("lists tables with their sheet and local address", async () => {
    expect(await listTables()).toEqual([{ name: "Orders", sheet: "Orders", address: "A1:F9" }]);
    const wb = withSheet(MINI, "Price list", { A1: { value: "Item" }, B1: { value: "Price" }, A2: { value: "x" }, B2: { value: 1 } });
    wb.tables.push({ name: "Prices", sheet: "Price list", address: "A1:B2" });
    fake.reset(wb);
    expect(await listTables()).toEqual([
      { name: "Orders", sheet: "Orders", address: "A1:F9" },
      { name: "Prices", sheet: "Price list", address: "A1:B2" },
    ]);
  });
});

describe("readSelection with whole columns", () => {
  it("trims A:F to the used part of the sheet instead of refusing it", async () => {
    fake.setSelection("Orders", "A:F");
    const data = await readSelection();
    expect(data.address).toBe("A1:F9");
    expect(data.values).toHaveLength(9);
  });

  it("still refuses a used range over the limit", async () => {
    fake.setCell("Orders", "Z1000", { value: 1, text: "1", type: "Double" });
    fake.setSelection("Orders", "A:Z");
    await expect(readSelection()).rejects.toBeInstanceOf(SelectionTooLargeError);
  });
});

describe("readSelection: what the selection points at (spec §7.1)", () => {
  /** Cells for rows of values with their top-left at `topLeft`; null leaves a cell empty. */
  function block(topLeft: string, rows: (string | number | null)[][]): Record<string, Partial<FakeCell>> {
    const at = parseCell(topLeft)!;
    const cells: Record<string, Partial<FakeCell>> = {};
    rows.forEach((row, r) =>
      row.forEach((v, c) => {
        if (v !== null) cells[`${columnLetter(at.c + c)}${at.r + r + 1}`] = { value: v };
      }),
    );
    return cells;
  }

  /** MINI plus a sheet "Log" holding `cells`, with an optional AutoFilter range, selected at `address`. */
  function logSheet(cells: Record<string, Partial<FakeCell>>, address: string, autoFilter?: string): void {
    const wb = withSheet(MINI, "Log", cells);
    if (autoFilter) wb.sheets.at(-1)!.autoFilter = autoFilter;
    fake.reset(wb);
    fake.setSelection("Log", address);
  }

  const LOG = [
    ["Day", "Team", "Hours"],
    ["Mon", "Red", 4],
    ["Tue", "Blue", 6],
    ["Wed", "Red", 5],
  ];
  const LATER = [
    ["Fri", "Blue", 3],
    ["Sat", "Red", 2],
  ];

  async function resolveError(): Promise<unknown> {
    return readSelection().catch((e: unknown) => e);
  }

  it("finds the data around a clicked row under either reading of Excel's used range of a range", async () => {
    // Data in C1:F50, and something else on the sheet in column A (row 70), so the sheet's used
    // range starts in column A. Row 4 then starts on an empty cell under the second reading.
    const data = block("C1", [["Day", "Team", "Hours", "Rate"], ...Array.from({ length: 49 }, (_, i) => [`D${i}`, "Red", i, i * 2])]);
    for (const mode of ["cells", "sheet"] as const) {
      logSheet({ ...data, A70: { value: "note" } }, "4:4");
      fake.usedRange = mode;
      const d = await readSelection();
      expect(d.address, mode).toBe("C1:F50");
      expect(d.resolution?.kind, mode).toBe("region");
      // Data from column B, a note further down in column A.
      logSheet({ ...block("B1", [["Day", "Team"], ...Array.from({ length: 49 }, (_, i) => [`D${i}`, "Red"])]), A60: { value: "note" } }, "4:4");
      fake.usedRange = mode;
      expect((await readSelection()).address, mode).toBe("B1:C50");
      // A clicked row with nothing in it still asks for the data.
      logSheet({ ...data, A70: { value: "note" } }, "60:60");
      fake.usedRange = mode;
      expect(((await resolveError()) as Error).message, mode).toBe(RESOLVE_MESSAGE);
      // A table the row crosses where it is blank (data only in H1:J10 of H1:J100) isn't the data.
      const wb = withSheet(MINI, "Mixed", {
        ...block("B1", [["Name", "Score"], ...Array.from({ length: 49 }, (_, i) => [`N${i}`, i])]),
        ...block("H1", [["Key", "Value", "Note"], ...Array.from({ length: 9 }, (_, i) => [`K${i}`, i, "x"])]),
      });
      wb.tables.push({ name: "KeyTbl", sheet: "Mixed", address: "H1:J100" });
      fake.reset(wb);
      fake.setSelection("Mixed", "30:30");
      fake.usedRange = mode;
      const mixed = await readSelection();
      expect(mixed.address, mode).toBe("B1:C50");
      expect(mixed.resolution?.kind, mode).toBe("region");
      // The mirror: a table blank in the row, left of the row's first filled cell.
      const left = withSheet(MINI, "Left", {
        ...block("A1", [["Key", "Value", "Note"], ...Array.from({ length: 49 }, (_, i) => (i === 28 ? [null, null, null] : [`K${i}`, i, "x"]))]),
        ...block("E1", [["Name", "Score"], ...Array.from({ length: 49 }, (_, i) => [`N${i}`, i])]),
      });
      left.tables.push({ name: "LeftTbl", sheet: "Left", address: "A1:C50" });
      fake.reset(left);
      fake.setSelection("Left", "30:30");
      fake.usedRange = mode;
      expect((await readSelection()).address, mode).toBe("E1:F50");
      // Nor one lying between the row's filled cells.
      const mid = withSheet(MINI, "Mid", {
        ...block("B1", [["Name", "Score"], ...Array.from({ length: 49 }, (_, i) => [`N${i}`, i])]),
        ...block("E1", [["Key", "Value"], ...Array.from({ length: 49 }, (_, i) => (i === 28 ? [null, null] : [`K${i}`, i]))]),
        ...block("H1", [["Team", "Hours"], ...Array.from({ length: 49 }, (_, i) => [`T${i}`, i])]),
      });
      mid.tables.push({ name: "Mid", sheet: "Mid", address: "E1:F50" });
      fake.reset(mid);
      fake.setSelection("Mid", "30:30");
      fake.usedRange = mode;
      expect((await readSelection()).address, mode).toBe("B1:C50");
      // A selected one-row range with nothing in it stands for the data around it, as one empty cell does.
      logSheet(data, "B51:G51");
      fake.usedRange = mode;
      expect((await readSelection()).address, mode).toBe("C1:F50");
    }
  });

  it("grows one cell inside a table to the table", async () => {
    fake.setSelection("Orders", "E4");
    const d = await readSelection();
    expect(d.address).toBe("A1:F9");
    expect(d.values).toHaveLength(9);
    expect(d.table?.name).toBe("Orders");
    expect(d.resolution).toEqual({ kind: "table", from: "E4", exact: "E4", table: "Orders" });
  });

  it("grows part of one row inside a table to the table", async () => {
    fake.setSelection("Orders", "B4:D4");
    const d = await readSelection();
    expect(d.address).toBe("A1:F9");
    expect(d.resolution).toEqual({ kind: "table", from: "B4:D4", exact: "B4:D4", table: "Orders" });
  });

  it("leaves a table's totals row out", async () => {
    const wb = structuredClone(MINI);
    wb.sheets[0]!.cells.F10 = { value: 14000, text: "14,000.00", type: "Double", format: "#,##0.00" };
    wb.tables = [{ name: "Orders", sheet: "Orders", address: "A1:F10", showTotals: true }];
    fake.reset(wb);
    fake.setSelection("Orders", "C3");
    const d = await readSelection();
    expect(d.address).toBe("A1:F9");
    expect(d.values.flat()).not.toContain(14000);
  });

  it("uses the only table a row touches when its first cell is outside it", async () => {
    const wb = withSheet(MINI, "Log", { ...block("C1", LOG), A3: { value: "note" } });
    wb.tables.push({ name: "Hours", sheet: "Log", address: "C1:E4" });
    fake.reset(wb);
    fake.setSelection("Log", "A3:E3");
    const d = await readSelection();
    expect(d.address).toBe("C1:E4");
    expect(d.resolution).toMatchObject({ kind: "table", table: "Hours" });
  });

  it("picks among the tables on a row: the one holding the anchor, else the only one where the row holds something", async () => {
    const col = (top: string, header: [string, string], blankRow: number | null) =>
      block(top, [header, ...Array.from({ length: 49 }, (_, i) => (i + 2 === blankRow ? [null, null] : [`${header[0]}${i}`, i]))]);
    const tables = (sheet: string, cells: Record<string, Partial<FakeCell>>, list: [string, string][]) => {
      const wb = withSheet(MINI, sheet, cells);
      for (const [name, address] of list) wb.tables.push({ name, sheet, address });
      return wb;
    };
    for (const mode of ["cells", "sheet"] as const) {
      const run = async (wb: FakeWorkbook, sheet: string, address: string) => {
        fake.reset(wb);
        fake.setSelection(sheet, address);
        fake.usedRange = mode;
        return readSelection().catch((e: unknown) => e as Error);
      };
      // Row 30 holds something in two tables; the anchor (its first filled cell) is in TA.
      const two = tables("Two", { ...col("B1", ["A", "B"], 20), ...col("E1", ["C", "D"], null) }, [["TA", "B1:C50"], ["TB", "E1:F50"]]);
      for (const sel of ["30:30", "A30:F30"]) {
        const d = (await run(two, "Two", sel)) as RangeData;
        expect(d.address, `${mode} ${sel}`).toBe("B1:C50");
        expect(d.resolution?.table, `${mode} ${sel}`).toBe("TA");
      }
      // The anchor is a note outside both, and the row holds something in both: no table, and the
      // note's own block is one row.
      const note = tables("Note", { ...col("C1", ["A", "B"], null), ...col("F1", ["C", "D"], null), A30: { value: "note" } }, [["T1", "C1:D50"], ["T2", "F1:G50"]]);
      expect(((await run(note, "Note", "30:30")) as Error).message, mode).toBe(RESOLVE_MESSAGE);
      // Only the table's last column, or only its first, holds something on the row: that table.
      for (const filled of ["F30", "C30"]) {
        const edge: Record<string, Partial<FakeCell>> = { ...block("C1", [["W", "X", "Y", "Z"]]), A30: { value: "note" }, [filled]: { value: 1 } };
        for (let r = 2; r <= 50; r++) if (r !== 30) edge[`C${r}`] = { value: r };
        const d = (await run(tables("Edge", edge, [["Wide", "C1:F50"]]), "Edge", "30:30")) as RangeData;
        expect(d.address, `${mode} ${filled}`).toBe("C1:F50");
        expect(d.resolution?.kind, `${mode} ${filled}`).toBe("table");
      }
      // A selected row that ends, or starts, inside a table blank on that row: not the table.
      const keyed = tables(
        "Keyed",
        { ...col("B1", ["Name", "Score"], null), ...block("H1", [["Key", "Value", "Note"], ...Array.from({ length: 9 }, (_, i) => [`K${i}`, i, "x"])]) },
        [["KeyTbl", "H1:J100"]],
      );
      const ends = (await run(keyed, "Keyed", "B30:I30")) as RangeData;
      expect([ends.address, ends.resolution?.kind], mode).toEqual(["B1:C50", "region"]);
      const starts = (await run(tables("Keyed2", { ...block("H1", [["Key", "Value", "Note"], ...Array.from({ length: 9 }, (_, i) => [`K${i}`, i, "x"])]), ...col("K1", ["P", "Q"], null) }, [["KeyTbl", "H1:J100"]]), "Keyed2", "I30:L30")) as RangeData;
      expect([starts.address, starts.resolution?.kind], mode).toEqual(["H1:L50", "region"]);
      // An empty selected row inside a table stands for the table, as one empty cell there does.
      const inside = tables("Inside", col("C1", ["A", "B"], null), [["Big", "C1:F60"]]);
      const empty = (await run(inside, "Inside", "D55:E55")) as RangeData;
      expect([empty.address, empty.resolution?.table], mode).toEqual(["C1:F60", "Big"]);
    }
  });

  it("grows a cell outside tables to the block of filled cells around it, which stops at a blank row", async () => {
    logSheet({ ...block("A1", LOG), ...block("A6", LATER) }, "B2");
    const d = await readSelection();
    expect(d.address).toBe("A1:C4");
    expect(d.table).toBeNull();
    expect(d.resolution).toEqual({ kind: "region", from: "B2", exact: "B2" });
  });

  it("uses the sheet's AutoFilter range when it holds the cell, across a blank row", async () => {
    logSheet({ ...block("A1", LOG), ...block("A6", LATER) }, "B2", "A1:C7");
    const d = await readSelection();
    expect(d.address).toBe("A1:C7");
    expect(d.values[4]).toEqual([null, null, null]);
    expect(d.resolution).toEqual({ kind: "filter", from: "B2", exact: "B2" });
    // A cell outside the filtered range still gets its own block.
    fake.setCell("Log", "G10", { value: "Total" });
    fake.setCell("Log", "G11", { value: 20 });
    fake.setSelection("Log", "G10");
    expect((await readSelection()).address).toBe("G10:G11");
  });

  it("cuts the AutoFilter range to its filled cells: whole columns, or rows past the data", async () => {
    for (const autoFilter of ["A:F", "A1:F2000"]) {
      const wb = structuredClone(MINI);
      wb.tables = [];
      wb.sheets[0]!.autoFilter = autoFilter;
      fake.reset(wb);
      fake.setSelection("Orders", "C3");
      const d = await readSelection();
      expect(d.address, autoFilter).toBe("A1:F9");
      expect(d.values, autoFilter).toHaveLength(9);
      expect(d.resolution, autoFilter).toEqual({ kind: "filter", from: "C3", exact: "C3" });
    }
    // A cell in the filter's empty rows below the data isn't in its filled cells: it gets its own block.
    fake.setSelection("Orders", "C1500");
    expect(((await resolveError()) as Error).message).toBe(RESOLVE_MESSAGE);
  });

  it("ignores an AutoFilter set on the header row alone", async () => {
    logSheet(block("A1", LOG), "B1", "A1:C1");
    const d = await readSelection();
    expect(d.address).toBe("A1:C4");
    expect(d.resolution?.kind).toBe("region");
  });

  it("leaves out a title row above the data when a blank row separates them", async () => {
    logSheet({ A1: { value: "Hours by team" }, ...block("A3", LOG) }, "B4");
    expect((await readSelection()).address).toBe("A3:C6");
  });

  it("cuts the empty column off the block around an empty cell next to the data", async () => {
    logSheet(block("A1", LOG), "D3");
    const d = await readSelection();
    expect(d.address).toBe("A1:C4");
    expect(d.resolution).toEqual({ kind: "region", from: "D3", exact: "D3" });
  });

  it("trims a whole row to the cells in use, then grows it like a pointer (4:4)", async () => {
    fake.setSelection("Regions", "4:4");
    const before = fake.loads.length;
    const d = await readSelection();
    expect(d.address).toBe("A1:B5");
    expect(d.values).toHaveLength(5);
    expect(d.resolution).toEqual({ kind: "region", from: "4:4", exact: "A4:B4" });
    // Nothing as wide as the row itself was read.
    const content = loadsSince(before).filter((l) => l.type === "Range" && l.props.some((p) => ["values", "text", "formulas"].includes(p)));
    expect(content.map((l) => l.address)).toEqual(["A1:B5"]);
  });

  it("trims whole columns to the rows in use and uses them as selected (A:F)", async () => {
    fake.setSelection("Orders", "A:F");
    const d = await readSelection();
    expect(d.address).toBe("A1:F9");
    expect(d.resolution).toEqual({ kind: "trimmed", from: "A:F", exact: "A1:F9" });
    fake.setSelection("Orders", "1:3");
    expect((await readSelection()).resolution).toEqual({ kind: "trimmed", from: "1:3", exact: "A1:F3" });
  });

  it("trims the whole sheet to the cells in use", async () => {
    fake.setSelection("Regions", "1:1048576");
    const d = await readSelection();
    expect(d.address).toBe("A1:B5");
    expect(d.resolution).toEqual({ kind: "trimmed", from: "1:1048576", exact: "A1:B5" });
  });

  it("uses a selection of several rows as selected, a single column too", async () => {
    fake.setSelection("Orders", "B2:B5");
    const d = await readSelection();
    expect(d.address).toBe("B2:B5");
    expect(d.resolution).toEqual({ kind: "selection", from: "B2:B5", exact: "B2:B5" });
  });

  it("asks for the data when there is nothing to grow to", async () => {
    for (const [sheet, address] of [
      ["Regions", "20:20"], // an empty row
      ["Regions", "Z100"], // an empty cell far from the data
      ["Regions", "D:E"], // empty columns
    ] as const) {
      fake.setSelection(sheet, address);
      const err = await resolveError();
      expect(err, address).toBeInstanceOf(ExcelHostError);
      expect((err as Error).message, address).toBe(RESOLVE_MESSAGE);
    }
    // An isolated cell is a header with no row under it.
    fake.setCell("Regions", "H20", { value: "alone" });
    fake.setSelection("Regions", "H20");
    expect(((await resolveError()) as Error).message).toBe(RESOLVE_MESSAGE);
    // So is a table-less row of headers with nothing under it.
    logSheet(block("A1", [LOG[0]!]), "A1:C1");
    expect(((await resolveError()) as Error).message).toBe(RESOLVE_MESSAGE);
    expect(fake.writes).toEqual([]);
  });

  it("with exact, reads the selected cells after trimming, one cell or one row included", async () => {
    fake.setSelection("Orders", "E4");
    const cell = await readSelection({ exact: true });
    expect(cell.address).toBe("E4");
    expect(cell.values).toEqual([["Widget"]]);
    expect(cell.resolution).toEqual({ kind: "selection", from: "E4", exact: "E4" });
    fake.setSelection("Regions", "4:4");
    const row = await readSelection({ exact: true });
    expect(row.address).toBe("A4:B4");
    expect(row.values).toEqual([["North", 3500]]);
    expect(row.resolution).toEqual({ kind: "trimmed", from: "4:4", exact: "A4:B4" });
    fake.setSelection("Regions", "20:20");
    await expect(readSelection({ exact: true })).rejects.toThrow(RESOLVE_MESSAGE);
  });

  it("returns an address the flow can parse for every kind of selection", async () => {
    const cases: [string, string][] = [
      ["Orders", "E4"],
      ["Orders", "4:4"],
      ["Orders", "A:F"],
      ["Orders", "C:C"],
      ["Orders", "1:1048576"],
      ["Regions", "B3"],
      ["Regions", "2:3"],
      ["Regions", "A1:B5"],
    ];
    for (const [sheet, address] of cases) {
      fake.setSelection(sheet, address);
      for (const exact of [false, true]) {
        const d = await readSelection({ exact });
        expect(parseLocalRange(d.address), `${address} exact=${exact}`).not.toBeNull();
        expect(parseLocalRange(d.resolution!.exact), `${address} exact=${exact}`).not.toBeNull();
      }
    }
    // A row filled from A to XFD trims to all of itself, which Excel names "4:4": refused, not returned.
    logSheet({ A4: { value: "from" }, XFD4: { value: "to" } }, "4:4");
    await expect(readSelection({ exact: true })).rejects.toThrow(RESOLVE_MESSAGE);
  });

  it("reads a table's columns while growing to it, so the read itself asks Excel for no tables", async () => {
    fake.setSelection("Orders", "E4");
    const ctx = (G.Excel as { RequestContext: { prototype: { sync(): Promise<void> } } }).RequestContext.prototype;
    const sync = vi.spyOn(ctx, "sync");
    try {
      const d = await readSelection();
      expect(d.table).toEqual({ name: "Orders", columns: ["Customer", "Email", "Region", "Order date", "Product", "Amount"] });
      // The selection, the table/filter/region lookup, the table's range and columns, the content,
      // and Excel.run's closing sync, which has nothing left to send.
      expect(sync).toHaveBeenCalledTimes(5);
    } finally {
      sync.mockRestore();
    }
  });

  it("names the pointer when the data around it is over the limit, before loading any cell content", async () => {
    // 101 columns x 200 rows = 20,200 cells.
    const rows = Array.from({ length: 200 }, (_, r) => Array.from({ length: 101 }, (_, c) => (r === 0 ? `H${c}` : r * c)));
    logSheet(block("A1", rows), "B2");
    const before = fake.loads.length;
    const err = await resolveError();
    expect(err).toBeInstanceOf(SelectionTooLargeError);
    expect((err as Error).message).toBe("The data around B2 has 20,200 cells. Select fewer cells (limit 20,000).");
    expect((err as SelectionTooLargeError).cellCount).toBe(20200);
    expect(onlyGeometrySince(before)).toBe(true);
    expect(contentLoads(before)).toEqual([]);

    // The same for a table, and for a whole row that grows.
    const wb = fake.workbook();
    wb.tables.push({ name: "Big", sheet: "Log", address: "A1:CW200" });
    fake.reset(wb);
    fake.setSelection("Log", "C5");
    const beforeTable = fake.loads.length;
    expect(((await resolveError()) as Error).message).toBe("The data around C5 has 20,200 cells. Select fewer cells (limit 20,000).");
    // A table's column names are its header cells: not loaded before the size check either.
    expect(loadsSince(beforeTable).filter((l) => l.type === "TableColumnCollection")).toEqual([]);
    expect(contentLoads(beforeTable)).toEqual([]);
    expect(onlyGeometrySince(beforeTable)).toBe(true);
    fake.setSelection("Log", "7:7");
    const beforeRow = fake.loads.length;
    expect(((await resolveError()) as Error).message).toBe("The data around 7:7 has 20,200 cells. Select fewer cells (limit 20,000).");
    expect(onlyGeometrySince(beforeRow)).toBe(true);
    // A row that starts on an empty cell grows from its first filled cell: nothing of that region
    // loads before the check either, under either reading of Excel's used range.
    for (const mode of ["cells", "sheet"] as const) {
      for (const sel of ["7:7", "A7:CY7"]) {
        logSheet({ ...block("C1", rows.map((r) => r.slice(0, 101))), A300: { value: "note" } }, sel);
        fake.usedRange = mode;
        const beforeEmpty = fake.loads.length;
        expect(((await resolveError()) as Error).message, `${mode} ${sel}`).toMatch(/^The data around .+ has 20,200 cells\./u);
        expect(onlyGeometrySince(beforeEmpty), `${mode} ${sel}`).toBe(true);
      }
    }
    // The same for a filter over too many cells.
    logSheet(block("A1", rows), "C5", "A:CW");
    const beforeFilter = fake.loads.length;
    expect(((await resolveError()) as Error).message).toBe("The data around C5 has 20,200 cells. Select fewer cells (limit 20,000).");
    expect(onlyGeometrySince(beforeFilter)).toBe(true);
  });

  it("reads the resolved range with the count it already has (no second size check)", async () => {
    fake.setSelection("Orders", "E4");
    const before = fake.loads.length;
    await readSelection();
    const counts = loadsSince(before).filter((l) => l.type === "Range" && l.props.join() === "rowCount,columnCount");
    expect(counts).toEqual([]);
    // readRange doesn't know the size, so it checks it first.
    const at = fake.loads.length;
    await readRange("Regions!A1:B5");
    expect(loadsSince(at)[0]).toMatchObject({ type: "Range", address: "A1:B5", props: ["rowCount", "columnCount"] });
  });

  it("changes neither the workbook nor the selection", async () => {
    for (const address of ["E4", "4:4", "A:F", "B2:B5"]) {
      fake.setSelection("Orders", address);
      await readSelection();
      expect(fake.getSelection().address).toBe(address);
    }
    expect(fake.writes).toEqual([]);
    expect(fake.selections).toEqual([]);
  });
});

describe("isRangeEmpty", () => {
  it("reads value types only and reports whether every cell is empty", async () => {
    expect(await isRangeEmpty("Orders", "H2:H9")).toBe(true);
    expect(await isRangeEmpty("Orders", "A2")).toBe(false);
    expect(await isRangeEmpty("Orders", "F9:G12")).toBe(false);
  });
});

describe("selectCells", () => {
  it("selects several cells together, shows their sheet, and writes nothing", async () => {
    await selectCells("Orders", ["A2", "A5", "C3"]);
    expect(fake.getSelection()).toEqual({ sheet: "Orders", address: "A2, A5, C3" });
    expect(fake.writes).toEqual([]);
  });

  it("selects the first cell only where Excel can't select several (before ExcelApi 1.18)", async () => {
    fake.maxApi = "1.17";
    await selectCells("Orders", ["A2", "A5"]);
    expect(fake.getSelection()).toEqual({ sheet: "Orders", address: "A2" });
  });

  it("selects a whole range, as restoring the selection does", async () => {
    await selectCells("Orders", ["A1:F501"]);
    expect(fake.getSelection()).toEqual({ sheet: "Orders", address: "A1:F501" });
  });

  it("skips anything that isn't a cell address, and refuses when nothing is left", async () => {
    await selectCells("Orders", ["A2", "Sheet2!B3", "=HYPERLINK(1)", "B4"]);
    expect(fake.getSelection()).toEqual({ sheet: "Orders", address: "A2, B4" });
    await expect(selectCells("Orders", ["not a cell"])).rejects.toBeInstanceOf(ExcelHostError);
  });
});

describe("listNames", () => {
  it("lists workbook names and sheet-scoped names as Sheet!Name", async () => {
    expect(await listNames()).toEqual([]);
    const wb = structuredClone(MINI) as FakeWorkbook & { names?: { name: string; sheet?: string }[] };
    wb.names = [{ name: "TaxRate" }, { name: "SUM", sheet: "Regions" }];
    fake.reset(wb);
    expect(await listNames()).toEqual(["TaxRate", "Regions!SUM"]);
  });
});

describe("readWorkbookInfo", () => {
  it("reads sheets, tables and names in one run, as the three separate reads do", async () => {
    const wb = withSheet(MINI, "Price list", { A1: { value: "Item" }, B1: { value: "Price" }, A2: { value: "x" }, B2: { value: 1 } }) as FakeWorkbook & {
      names?: { name: string; sheet?: string }[];
    };
    wb.tables.push({ name: "Prices", sheet: "Price list", address: "A1:B2" });
    wb.names = [{ name: "TaxRate" }, { name: "SUM", sheet: "Regions" }];
    fake.reset(wb);
    const info = await readWorkbookInfo();
    expect(info).toEqual({ sheets: await listSheets(), tables: await listTables(), names: await listNames() });
    expect(info).toEqual({
      sheets: ["Orders", "Regions", "Price list"],
      tables: [
        { name: "Orders", sheet: "Orders", address: "A1:F9" },
        { name: "Prices", sheet: "Price list", address: "A1:B2" },
      ],
      names: ["TaxRate", "Regions!SUM"],
    });
  });

  it("fails as a whole, with a plain message", async () => {
    fake.failNext("GeneralException", "Internal error with Ana Ruiz");
    await expect(readWorkbookInfo()).rejects.toThrow("Excel couldn't finish that. Try again.");
  });
});

describe("readRange", () => {
  it("reads a context range on another sheet", async () => {
    const d = await readRange("Regions!A1:B5");
    expect(d.sheet).toBe("Regions");
    expect(d.address).toBe("A1:B5");
    expect(d.values).toEqual([
      ["Region", "Target"],
      ["East", 5000],
      ["West", 4000],
      ["North", 3500],
      ["South", 2500],
    ]);
    expect(d.text[1]).toEqual(["East", "5,000"]);
    expect(d.table).toBeNull();
  });

  it("reads quoted sheet names, including doubled quotes", async () => {
    let wb = withSheet(MINI, "Q1 Sales", { B2: { value: "a" }, C3: { value: 3 } });
    wb = withSheet(wb, "Bob's list", { A1: { value: "b" } });
    fake.reset(wb);
    const q = await readRange("'Q1 Sales'!B2:C3");
    expect(q.sheet).toBe("Q1 Sales");
    expect(q.address).toBe("B2:C3");
    expect(q.values).toEqual([
      ["a", null],
      [null, 3],
    ]);
    const b = await readRange("'Bob''s list'!A1:A2");
    expect(b.sheet).toBe("Bob's list");
    expect(b.values).toEqual([["b"], [null]]);
  });

  it("normalizes absolute addresses and detects tables", async () => {
    const d = await readRange("Orders!$A$1:$B$3");
    expect(d.address).toBe("A1:B3");
    expect(d.table).toEqual({ name: "Orders", columns: ["Customer", "Email"] });
  });

  it("uses the active sheet when the address has none", async () => {
    const d = await readRange("C1:C2");
    expect(d.sheet).toBe("Orders");
    expect(d.values).toEqual([["Region"], ["East"]]);
  });

  it("rejects unknown sheets and unreadable addresses with plain messages", async () => {
    const err = (await readRange("Nowhere!A1:B2").catch((e: unknown) => e)) as Error;
    expect(err).toBeInstanceOf(ExcelHostError);
    expect(err.message).toBe("There's no sheet with that name. Check the sheet name in the address.");
    expect(err.message).not.toMatch(/Nowhere|RichApi/);
    for (const bad of ["Regions!A:B", "Regions!Totals", "", "Regions!"]) {
      await expect(readRange(bad), bad).rejects.toThrow("Enter a range like Regions!A1:B5.");
    }
  });

  it("applies the cell limit", async () => {
    await expect(readRange("Orders!A1:Z1000")).rejects.toBeInstanceOf(SelectionTooLargeError);
  });
});

describe("firstEmptyColumn", () => {
  it("finds the first empty column right of the selection, reading value types only", async () => {
    const before = fake.loads.length;
    expect(await firstEmptyColumn("Orders", 6, 1, 9)).toBe("G");
    expect(contentLoads(before)).toEqual([]);
    fake.setCell("Orders", "G5", { value: "note" });
    fake.setCell("Orders", "H1", { value: "Total" });
    expect(await firstEmptyColumn("Orders", 6, 1, 9)).toBe("I");
    expect(await firstEmptyColumn("Orders", 6, 2, 9)).toBe("H");
  });

  it("scans in chunks under the cell limit and falls back to the first column", async () => {
    for (let c = 0; c < 10; c++) fake.setCell("Regions", `${String.fromCharCode(67 + c)}1`, { value: c });
    const before = fake.loads.length;
    expect(await firstEmptyColumn("Regions", 2, 1, 5000)).toBe("M");
    const sizes = loadsSince(before)
      .filter((l) => l.type === "Range")
      .map((l) => parseLocalRange(l.address ?? ""))
      .map((r) => (r ? (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) : Infinity));
    expect(sizes).toEqual([20000, 20000, 20000]);
    expect(await firstEmptyColumn("Regions", 2, 1, 5, 3)).toBe("C");
    expect(await firstEmptyColumn("Regions", Number.NaN, 1, Number.NaN)).toBe("M");
  });
});

describe("insertFormula", () => {
  it("sets the formula on one cell and reads back only value types", async () => {
    const before = fake.loads.length;
    const res = await insertFormula("Orders", "G2", "=F2*2", false, 9);
    expect(res).toEqual({ address: "G2", check: "ok", excelError: false });
    expect(fake.writes).toEqual([{ kind: "formulas", sheet: "Orders", address: "G2", formulas: [["=F2*2"]] }]);
    expect(fake.getCell("Orders", "G2")).toMatchObject({ formula: "=F2*2", value: 0, type: "Double" });
    expect(contentLoads(before)).toEqual([]);
    const rangeLoads = loadsSince(before).filter((l) => l.type === "Range");
    expect(rangeLoads.flatMap((l) => l.props).filter((p) => p !== "valueTypes" && !/Index|Count/.test(p))).toEqual([]);
    expect(rangeLoads.some((l) => l.props.includes("valueTypes") && l.address === "G2")).toBe(true);
  });

  it("autofills down to lastRow with FillDefault (spec S4)", async () => {
    const res = await insertFormula("Orders", "G2", "=F2*2", true, 11);
    expect(res).toEqual({ address: "G2:G11", check: "ok", excelError: false });
    expect(fake.writes).toEqual([
      { kind: "formulas", sheet: "Orders", address: "G2", formulas: [["=F2*2"]] },
      { kind: "autoFill", sheet: "Orders", source: "G2", destination: "G2:G11", type: "FillDefault" },
    ]);
    expect(fake.getCell("Orders", "G3")?.formula).toBe("=F3*2");
    expect(fake.getCell("Orders", "G11")?.formula).toBe("=F11*2");
    expect(fake.getCell("Orders", "G12")).toBeNull();
  });

  it("doesn't fill when lastRow isn't below the cell", async () => {
    expect(await insertFormula("Orders", "G9", "=F9*2", true, 9)).toEqual({ address: "G9", check: "ok", excelError: false });
    expect(await insertFormula("Orders", "H9", "=F9*2", true, 3)).toEqual({ address: "H9", check: "ok", excelError: false });
    expect(fake.writes.filter((w) => w.kind === "autoFill")).toEqual([]);
  });

  it("reports excelError when an inserted cell shows an Excel error", async () => {
    fake.evaluate = (_f, _s, addr) => (addr === "G5" ? { value: "#DIV/0!", type: "Error" } : { value: 1, type: "Double" });
    const res = await insertFormula("Orders", "G2", "=F2/C2", true, 9);
    expect(res).toEqual({ address: "G2:G9", check: "error", excelError: true });
    expect(fake.getCell("Orders", "G5")).toMatchObject({ formula: "=F5/C5", type: "Error" });
  });

  it("leaves excelError unset when the workbook calculates manually", async () => {
    fake.setCalculationMode("Manual");
    fake.evaluate = () => ({ value: "#N/A", type: "Error" });
    const before = fake.loads.length;
    const res = await insertFormula("Orders", "G2", "=F2*2", true, 9);
    expect(res).toEqual({ address: "G2:G9", check: "manual" });
    expect("excelError" in res).toBe(false);
    expect(loadsSince(before).filter((l) => l.type === "Range")).toEqual([]);
    expect(fake.getCell("Orders", "G9")?.formula).toBe("=F9*2");
  });

  it("checks the spill range of a spilling formula", async () => {
    fake.evaluate = () => ({
      spill: [[{ value: "East", type: "String" }], [{ value: "West", type: "String" }], [{ value: "#N/A", type: "Error" }]],
    });
    const res = await insertFormula("Orders", "H2", "=UNIQUE(C2:C9)", false, 9);
    expect(res).toEqual({ address: "H2", spill: "H2:H4", check: "error", excelError: true });
    expect(fake.getCell("Orders", "H3")).toMatchObject({ value: "West", spilledFrom: "H2" });

    fake.reset(structuredClone(MINI));
    fake.evaluate = () => ({ spill: [[{ value: "East", type: "String" }], [{ value: "West", type: "String" }]] });
    expect(await insertFormula("Orders", "H2", "=UNIQUE(C2:C9)", false, 9)).toEqual({
      address: "H2",
      spill: "H2:H3",
      check: "ok",
      excelError: false,
    });
  });

  it("treats a blocked spill as an error", async () => {
    fake.setCell("Orders", "H3", { value: "in the way" });
    fake.evaluate = () => ({ spill: [[{ value: 1 }], [{ value: 2 }]] });
    expect(await insertFormula("Orders", "H2", "=SEQUENCE(2)", false, 9)).toEqual({ address: "H2", check: "error", excelError: true });
    expect(fake.getCell("Orders", "H2")).toMatchObject({ value: "#SPILL!", type: "Error" });
  });

  it("skips spill readback on hosts without ExcelApi 1.12", async () => {
    fake.maxApi = "1.11";
    fake.evaluate = () => ({ spill: [[{ value: 1 }], [{ value: "#N/A", type: "Error" }]] });
    // Only the anchor is checked; the fake throws if the 1.12 API is called.
    expect(await insertFormula("Orders", "H2", "=SEQUENCE(2)", false, 9)).toEqual({ address: "H2", check: "ok", excelError: false });
  });

  it("rejects a multi-cell target and turns Office errors into plain messages", async () => {
    await expect(insertFormula("Orders", "G2:G3", "=1", false, 9)).rejects.toThrow("Pick a single cell for the formula, like G2.");
    const err = (await insertFormula("Nowhere", "G2", "=1", false, 9).catch((e: unknown) => e)) as ExcelHostError;
    expect(err).toBeInstanceOf(ExcelHostError);
    expect(err.code).toBe("ItemNotFound");
    expect(err.message).not.toMatch(/RichApi|resource/);
    expect(err.stack).toBe(`ExcelHostError: ${err.message}`);
    expect(fake.writes).toEqual([]);
  });

  it("converts any failed sync into a plain message", async () => {
    fake.failNext("GeneralException", "Internal error at line 1\n    at stack frame");
    const err = (await insertFormula("Orders", "G2", "=1", false, 9).catch((e: unknown) => e)) as Error;
    expect(err.message).toBe("Excel couldn't finish that. Try again.");
    const odd = toPlainError(new Error("Something odd with Ana Ruiz\n    at x (y.js:1:1)"));
    expect(odd.message).toBe("Excel reported a problem. Try again.");
    expect(odd.stack).toBe("ExcelHostError: Excel reported a problem. Try again.");
    expect(toPlainError(Object.assign(new Error("x"), { code: "WeirdCode" })).message).toBe(
      "Excel reported a problem (WeirdCode). Try again.",
    );
  });
});

describe("insertFormula once the formula is written (review 2026-09-28)", () => {
  // Sync 1 writes the formula. Without fill down: 2 reads the calculation mode, 3 the spill range,
  // 4 the value types. With fill down: 2 fills, 3 the calculation mode, 4 the value types.
  it("a readback that fails after the write reports the cell as written but unchecked, not as a failure", async () => {
    for (const n of [2, 3, 4]) {
      fake.reset(structuredClone(MINI));
      fake.failOnSync(n, "GeneralException");
      expect(await insertFormula("Orders", "G2", "=F2*2", false, 9)).toEqual({ address: "G2", check: "unverified" });
      expect(fake.getCell("Orders", "G2")?.formula).toBe("=F2*2");
      expect(fake.writes).toHaveLength(1);
    }
  });

  it("a spill readback that fails after the write is reported the same way", async () => {
    fake.evaluate = () => ({ spill: [[{ value: "East", type: "String" }], [{ value: "West", type: "String" }]] });
    fake.failOnSync(4, "GeneralException");
    expect(await insertFormula("Orders", "H2", "=UNIQUE(C2:C9)", false, 9)).toEqual({ address: "H2", check: "unverified" });
    expect(fake.getCell("Orders", "H2")?.formula).toBe("=UNIQUE(C2:C9)");
  });

  it("a fill down that fails after the first cell is written says only that cell was written", async () => {
    fake.failOnSync(2, "InvalidOperation");
    expect(await insertFormula("Orders", "G2", "=F2*2", true, 9)).toEqual({ address: "G2", check: "unverified", fillFailed: true });
    expect(fake.getCell("Orders", "G2")?.formula).toBe("=F2*2");
    expect(fake.getCell("Orders", "G3")).toBeNull();
    expect(fake.writes.filter((w) => w.kind === "autoFill")).toEqual([]);
  });

  it("the value-type check after a fill down failing is reported as written but unchecked", async () => {
    fake.failOnSync(4, "GeneralException");
    expect(await insertFormula("Orders", "G2", "=F2*2", true, 9)).toEqual({ address: "G2:G9", check: "unverified" });
    expect(fake.getCell("Orders", "G9")?.formula).toBe("=F9*2");
  });

  it("a write that fails throws, and nothing is written", async () => {
    fake.failOnSync(1, "AccessDenied");
    await expect(insertFormula("Orders", "G2", "=F2*2", true, 9)).rejects.toThrow(
      "Excel didn't allow the change. The sheet or workbook may be protected.",
    );
    expect(fake.writes).toEqual([]);
    expect(fake.getCell("Orders", "G2")).toBeNull();
  });
});

describe("bench helpers", () => {
  it("selectRange changes the selection without writing to the workbook", async () => {
    await selectRange("Regions", "A1:B5");
    expect(fake.getSelection()).toEqual({ sheet: "Regions", address: "A1:B5" });
    expect((await readSelection()).address).toBe("A1:B5");
    expect(fake.writes).toEqual([]);
  });

  it("writeScratch refuses to write outside bench builds (invariant 9)", async () => {
    bench.on = false;
    await expect(writeScratch("=1", false, { row: 2, lastRow: 9 })).rejects.toThrow("only used in bench builds");
    expect(fake.writes).toEqual([]);
  });

  it("writeScratch creates the scratch sheet when missing and writes in column A", async () => {
    const address = await writeScratch("=Orders!F2*2", false, { row: 2, lastRow: 9 });
    expect(address).toBe("A2");
    expect(fake.sheetNames()).toContain(SCRATCH_SHEET);
    expect(fake.writes[0]).toEqual({ kind: "addSheet", sheet: SCRATCH_SHEET });
    expect(fake.getCell(SCRATCH_SHEET, "A2")?.formula).toBe("=Orders!F2*2");
  });

  it("writeScratch clears the sheet first and fills down", async () => {
    fake.reset(withSheet(MINI, SCRATCH_SHEET, { Z99: { value: "old" }, A5: { value: 7 } }));
    const address = await writeScratch("=Orders!F2*2", true, { row: 2, lastRow: 9 });
    expect(address).toBe("A2:A9");
    expect(fake.getCell(SCRATCH_SHEET, "Z99")).toBeNull();
    expect(fake.getCell(SCRATCH_SHEET, "A9")?.formula).toBe("=Orders!F9*2");
    expect(fake.writes.map((w) => w.kind)).toEqual(["clear", "formulas", "autoFill"]);
  });

  it("writeScratch returns the spill range and doesn't fill a spilling formula", async () => {
    fake.evaluate = () => ({ spill: [[{ value: "East" }, { value: 3 }], [{ value: "West" }, { value: 2 }]] });
    expect(await writeScratch("=GROUPBY(Orders!C2:C9,Orders!F2:F9,SUM)", true, { row: 2, lastRow: 9 })).toBe("A2:B3");
    expect(fake.writes.filter((w) => w.kind === "autoFill")).toEqual([]);
  });

  it("readScratch reads values back with empty cells as null", async () => {
    fake.evaluate = (_f, _s, addr) => (addr === "A4" ? { value: "", type: "String" } : { value: Number(addr.slice(1)) * 10 });
    const address = await writeScratch("=ROW()*10", true, { row: 2, lastRow: 4 });
    expect(await readScratch(address)).toEqual([[20], [30], [""]]);
    expect(await readScratch("A1:A2")).toEqual([[null], [20]]);
  });

  it("readBenchInfo returns nothing when the workbook has no bench sheet", async () => {
    expect(await readBenchInfo()).toEqual({ marker: null, privateLetters: [], canaries: [] });
  });

  it("readBenchInfo reads the marker, private letters and canaries from the hidden sheet", async () => {
    const orders = JSON.parse(read("./e2e/fixtures/orders.json")) as FakeWorkbook;
    fake.reset(orders);
    const info = await readBenchInfo();
    expect(info.marker).toBe("NYMFORM_BENCH_V1");
    expect(info.privateLetters).toEqual(["A", "B"]);
    expect(info.canaries).toHaveLength(10);
    expect(info.canaries[0]).toBe("Quorbel Vantrisk");
    expect(fake.writes).toEqual([]);
  });
});

describe("fake Excel host", () => {
  it("throws PropertyNotLoaded for properties that weren't loaded and synced, like Office", async () => {
    await Excel.run(async (ctx) => {
      const r = ctx.workbook.getSelectedRange();
      expect(() => r.values).toThrow(/PropertyNotLoaded|not available/);
      r.load("address");
      expect(() => r.address).toThrow(/not available/);
      await ctx.sync();
      expect(r.address).toBe("Orders!A1:F9");
      expect(() => r.values).toThrow(/not available/);
      const missing = ctx.workbook.worksheets.getItemOrNullObject("Nope");
      expect(() => missing.isNullObject).toThrow(/not available/);
      await ctx.sync();
      expect(missing.isNullObject).toBe(true);
    });
  });

  it("queues writes until sync and quotes sheet names in addresses", async () => {
    fake.reset(withSheet(MINI, "Q1 Sales", {}));
    await Excel.run(async (ctx) => {
      const r = ctx.workbook.worksheets.getItem("Q1 Sales").getRange("B2");
      r.formulas = [["=1+1"]];
      expect(fake.getCell("Q1 Sales", "B2")).toBeNull();
      r.load("address");
      await ctx.sync();
      expect(fake.getCell("Q1 Sales", "B2")?.formula).toBe("=1+1");
      expect(r.address).toBe("'Q1 Sales'!B2");
    });
  });

  it("reports whole rows and columns, and the used range inside a range", async () => {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem("Orders");
      const row = ws.getRange("4:4");
      const cols = ws.getRange("B:C");
      const cell = ws.getRange("B4");
      const used = cols.getUsedRangeOrNullObject(true);
      const none = ws.getRange("H1:J5").getUsedRangeOrNullObject(true);
      for (const r of [row, cols, cell, used]) r.load("address,isEntireRow,isEntireColumn");
      none.load("address");
      await ctx.sync();
      expect([row.address, row.isEntireRow, row.isEntireColumn]).toEqual(["Orders!4:4", true, false]);
      expect([cols.address, cols.isEntireRow, cols.isEntireColumn]).toEqual(["Orders!B:C", false, true]);
      expect([cell.isEntireRow, cell.isEntireColumn]).toEqual([false, false]);
      expect(used.address).toBe("Orders!B1:C9");
      expect(none.isNullObject).toBe(true);
    });
  });

  it("finds the surrounding region like Excel's current region: diagonals count, blank rows and columns bound it", async () => {
    fake.reset(
      withSheet(MINI, "Grid", {
        A1: { value: "a" }, B1: { value: "b" },
        A2: { value: 1 }, B2: { value: 2 },
        C3: { value: "diagonal" }, // touches B2 at a corner only
        E1: { value: "apart" }, // column D is blank
        A5: { value: "below" }, // row 4 is blank
      }),
    );
    const region = async (address: string, sheet = "Grid") =>
      Excel.run(async (ctx) => {
        const r = ctx.workbook.worksheets.getItem(sheet).getRange(address).getSurroundingRegion();
        r.load("address");
        await ctx.sync();
        return r.address;
      });
    expect(await region("A1")).toBe("Grid!A1:C3");
    expect(await region("C3")).toBe("Grid!A1:C3");
    expect(await region("E1")).toBe("Grid!E1");
    expect(await region("A5")).toBe("Grid!A5");
    // An empty cell touching data takes it in, with its own row or column; here it bridges column D.
    expect(await region("D2")).toBe("Grid!A1:E3");
    // An empty cell with nothing around it is its own region; the top-left cell of a range decides.
    expect(await region("H10:J12")).toBe("Grid!H10");
    // Clipped at the sheet's edges.
    fake.setCell("Grid", "XFD1048576", { value: "corner" });
    fake.setCell("Grid", "XFC1048575", { value: "next" });
    expect(await region("XFD1048576")).toBe("Grid!XFC1048575:XFD1048576");
  });

  it("gives a sheet's AutoFilter range, or a null object without one, and a table's totals setting", async () => {
    const wb = structuredClone(MINI);
    wb.sheets[1]!.autoFilter = "A1:B5";
    wb.tables[0]!.showTotals = true;
    fake.reset(wb);
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      const filtered = sheets.getItem("Regions").autoFilter.getRangeOrNullObject();
      const unfiltered = sheets.getItem("Orders").autoFilter.getRangeOrNullObject();
      filtered.load("address");
      unfiltered.load("address");
      const tables = ctx.workbook.tables;
      tables.load("items/name,items/showTotals");
      await ctx.sync();
      expect(filtered.address).toBe("Regions!A1:B5");
      expect(unfiltered.isNullObject).toBe(true);
      expect(tables.items.map((t) => [t.name, t.showTotals])).toEqual([["Orders", true]]);
    });
    expect(fake.workbook().sheets[1]).toMatchObject({ name: "Regions", autoFilter: "A1:B5" });
  });

  it("shifts relative references when filling, leaving strings, absolute parts and names alone", () => {
    expect(fake.shiftFormula('=SUM($A$1:A2)+Regions!B2&"A1"+Orders[@Amount]+LOG10(A2)+$C2+D$2', 3, 0)).toBe(
      '=SUM($A$1:A5)+Regions!B5&"A1"+Orders[@Amount]+LOG10(A5)+$C5+D$2',
    );
    expect(fake.shiftFormula("='My Sheet'!A1+B1", 0, 1)).toBe("='My Sheet'!B1+C1");
  });
});

// ---------------------------------------------------------------------------------------------

describe("workbookToFake", () => {
  it("converts an .xlsx with types, display text, hidden sheets, tables and formulas", async () => {
    const book = new ExcelJS.Workbook();
    const ws = book.addWorksheet("Data");
    ws.addTable({ name: "Items", ref: "A1", headerRow: true, columns: [{ name: "Name" }, { name: "Price" }], rows: [["pen", 1.5], ["ink", 1234.5]] });
    ws.getCell("A3").value = "ink";
    ws.getCell("B2").numFmt = "#,##0.00";
    ws.getCell("B3").numFmt = "#,##0.00";
    ws.getCell("D1").value = new Date(Date.UTC(2025, 0, 31));
    ws.getCell("D1").numFmt = "m/d/yyyy";
    ws.getCell("D2").value = 615 / 1440;
    ws.getCell("D2").numFmt = "h:mm AM/PM";
    ws.getCell("D3").value = true;
    ws.getCell("E1").value = { formula: "SUM(B2:B3)", result: 1236 };
    ws.getCell("E2").value = { error: "#N/A" } as ExcelJS.CellErrorValue;
    const hidden = book.addWorksheet("_nymform_bench", { state: "hidden" });
    hidden.getCell("A1").value = "NYMFORM_BENCH_V1";
    const wb = await xlsxToFake(Buffer.from(await book.xlsx.writeBuffer()));

    expect(wb.sheets.map((s) => [s.name, s.visibility])).toEqual([
      ["Data", "Visible"],
      ["_nymform_bench", "Hidden"],
    ]);
    expect(wb.tables).toEqual([{ name: "Items", sheet: "Data", address: "A1:B3" }]);
    expect(wb.selection).toEqual({ sheet: "Data", address: "A1" });
    const c = wb.sheets[0]!.cells;
    expect(c.A2).toEqual({ value: "pen", text: "pen", type: "String", format: "General" });
    expect(c.B3).toEqual({ value: 1234.5, text: "1,234.50", type: "Double", format: "#,##0.00" });
    expect(c.D1).toEqual({ value: 45688, text: "1/31/2025", type: "Double", format: "m/d/yyyy" });
    expect(c.D2).toEqual({ value: 615 / 1440, text: "10:15 AM", type: "Double", format: "h:mm AM/PM" });
    expect(c.D3).toEqual({ value: true, text: "TRUE", type: "Boolean", format: "General" });
    expect(c.E1).toEqual({ value: 1236, text: "1236", type: "Double", format: "General", formula: "=SUM(B2:B3)" });
    expect(c.E2).toEqual({ value: "#N/A", text: "#N/A", type: "Error", format: "General" });
    expect(c.C1).toBeUndefined();

    // The converted workbook drives the adapter like a real one.
    fake.reset(JSON.parse(stringifyFake(wb)));
    fake.setSelection("Data", "A1:B3");
    const d = await readSelection();
    expect(d.table).toEqual({ name: "Items", columns: ["Name", "Price"] });
    expect(d.text).toEqual([
      ["Name", "Price"],
      ["pen", "1.50"],
      ["ink", "1,234.50"],
    ]);
    expect((await readBenchInfo()).marker).toBe("NYMFORM_BENCH_V1");
  });

  it("carries a sheet's AutoFilter range and a table's totals row across", async () => {
    const book = new ExcelJS.Workbook();
    const ws = book.addWorksheet("Log");
    ws.addRows([
      ["Day", "Hours"],
      ["Mon", 4],
      ["Tue", 6],
    ]);
    ws.autoFilter = "A1:B3";
    const other = book.addWorksheet("Totals");
    other.addTable({
      name: "Sums",
      ref: "A1",
      headerRow: true,
      totalsRow: true,
      columns: [{ name: "Team", totalsRowLabel: "Total" }, { name: "Hours", totalsRowFunction: "sum" }],
      rows: [["Red", 4], ["Blue", 6]],
    });
    const wb = await xlsxToFake(Buffer.from(await book.xlsx.writeBuffer()));
    expect(wb.sheets.map((s) => s.autoFilter)).toEqual(["A1:B3", undefined]);
    expect(wb.tables).toEqual([{ name: "Sums", sheet: "Totals", address: "A1:B4", showTotals: true }]);

    fake.reset(JSON.parse(stringifyFake(wb)));
    fake.setSelection("Log", "B2");
    expect((await readSelection()).resolution?.kind).toBe("filter");
    fake.setSelection("Totals", "A2");
    const d = await readSelection();
    expect(d.address).toBe("A1:B3");
    expect(d.resolution).toMatchObject({ kind: "table", table: "Sums" });
  });

  it("approximates Excel display text for common formats", () => {
    expect(displayText(-1234.5, "#,##0.00")).toBe("-1,234.50");
    expect(displayText(-5, "#,##0.00;(#,##0.00)")).toBe("(5.00)");
    expect(displayText(0.123, "0.0%")).toBe("12.3%");
    expect(displayText(45658.75, "h:mm AM/PM")).toBe("6:00 PM");
    expect(displayText(45658, "yyyy-mm-dd")).toBe("2025-01-01");
    expect(displayText(0.1 + 0.2, "General")).toBe("0.3");
    expect(displayText(false, "General")).toBe("FALSE");
    expect(displayText(null, "General")).toBe("");
  });

  it("round-trips every bench workbook to what the adapter would hand over (bench/generate.ts)", async () => {
    const data = generateData();
    for (const wb of data.workbooks) {
      const json = await xlsxToFake(await workbookBuffer(wb), { selection: { sheet: wb.sheet, address: wb.range } });
      fake.reset(json);
      const { resolution, ...data } = await readSelection();
      expect(data, wb.file).toEqual(toRangeData(wb));
      expect(resolution, wb.file).toEqual({ kind: "selection", from: wb.range, exact: wb.range });
      expect(await readBenchInfo(), wb.file).toEqual({
        marker: "NYMFORM_BENCH_V1",
        privateLetters: wb.privateLetters,
        canaries: wb.canaries,
      });
    }
  });

  it("tests/e2e/fixtures/orders.json matches bench/workbooks/orders.xlsx", async () => {
    const book = new ExcelJS.Workbook();
    await book.xlsx.readFile(fileURLToPath(new URL("../bench/workbooks/orders.xlsx", import.meta.url)));
    const expected = stringifyFake(workbookToFake(book, { selection: { sheet: "Orders", address: "A1:F501" } }));
    expect(
      read("./e2e/fixtures/orders.json") === expected,
      "orders.json is stale: npx tsx tests/e2e/workbookToFake.ts bench/workbooks/orders.xlsx tests/e2e/fixtures/orders.json --select=Orders!A1:F501",
    ).toBe(true);
  });
});
