// End-to-end tests of selection resolution (spec §7.1): one clicked cell, a row header, a filtered
// range, a table and "Use exactly the selected cells", against the built pane (dist/) under its
// real CSP with the fake Excel host. Build first: `npm run build`.
import { expect, test, type Page } from "@playwright/test";
import { ORDERS, enterKey, openPane, preview, reply, userTurn, type FakeWorkbook } from "./harness";

interface Fake {
  writes: unknown[];
  selections: unknown[];
  loads: { type: string; address?: string; props: string[] }[];
  setSelection(sheet: string, address: string): void;
  getSelection(): { sheet: string; address: string };
}

const KEEPS = "Nymform keeps using this range until you choose Refresh from selection.";
const USE_EXACT = "Use exactly the selected cells";
const RESOLVE = "Select your data, or one cell in it, and choose Refresh from selection.";

function writes(page: Page): Promise<unknown[]> {
  return page.evaluate(() => (window as unknown as { __NYMFORM_FAKE__: Fake }).__NYMFORM_FAKE__.writes);
}

/** Selections the pane made itself (Show in sheet); reading the selection makes none. */
function selectionsMade(page: Page): Promise<unknown[]> {
  return page.evaluate(() => (window as unknown as { __NYMFORM_FAKE__: Fake }).__NYMFORM_FAKE__.selections);
}

/** Clicks cells in Excel, as the user would while the pane is open. */
function select(page: Page, address: string): Promise<void> {
  return page.evaluate((a) => (window as unknown as { __NYMFORM_FAKE__: Fake }).__NYMFORM_FAKE__.setSelection("Orders", a), address);
}

function selection(page: Page): Promise<{ sheet: string; address: string }> {
  return page.evaluate(() => (window as unknown as { __NYMFORM_FAKE__: Fake }).__NYMFORM_FAKE__.getSelection());
}

function selecting(address: string, workbook: FakeWorkbook = ORDERS): FakeWorkbook {
  return { ...structuredClone(workbook), selection: { sheet: "Orders", address } };
}

/** Largest range, in cells, the pane loaded any cell content of. */
async function largestContentLoad(page: Page): Promise<number> {
  const loads = await page.evaluate(() => (window as unknown as { __NYMFORM_FAKE__: Fake }).__NYMFORM_FAKE__.loads);
  let largest = 0;
  for (const l of loads) {
    if (l.type !== "Range" || !l.props.some((p) => ["values", "text", "formulas", "valueTypes", "numberFormat"].includes(p))) continue;
    const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(l.address ?? "");
    if (!m) return Infinity;
    const col = (s: string) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    const rows = m[3] ? Number(m[4]) - Number(m[2]) + 1 : 1;
    const cols = m[3] ? col(m[3]) - col(m[1]!) + 1 : 1;
    largest = Math.max(largest, rows * cols);
  }
  return largest;
}

test("one clicked cell stands for the block of filled cells around it", async ({ page }) => {
  const pane = await openPane(page, { workbook: selecting("E4") });
  await enterKey(page);
  await expect(page.getByText("Using Orders, A1:F501 · 500 rows")).toBeVisible();
  await expect(page.getByText("From the block of filled cells around E4.")).toBeVisible();
  await expect(page.getByRole("button", { name: USE_EXACT })).toBeVisible();
  await expect(page.getByText(KEEPS)).toBeVisible();
  // Customer and Email are suggested private, as for the whole range selected by hand.
  await expect(page.getByRole("listitem", { name: "Column A" }).getByRole("checkbox", { name: "Keep private" })).toBeChecked();

  // Clicking elsewhere in Excel changes nothing until Refresh from selection.
  await select(page, "H20");
  await page.getByRole("tab", { name: "Ask" }).click();
  await page.getByRole("tab", { name: "Columns" }).click();
  await expect(page.getByText("Using Orders, A1:F501 · 500 rows")).toBeVisible();
  await page.getByRole("button", { name: "Refresh from selection" }).click();
  await expect(page.getByText(RESOLVE)).toBeVisible();

  // Reading the selection is not a write, and the pane leaves the selection where it was.
  expect(await writes(page)).toEqual([]);
  expect(await selectionsMade(page)).toEqual([]);
  expect(await selection(page)).toEqual({ sheet: "Orders", address: "H20" });
  expect(pane.consoleErrors).toEqual([]);
});

test("a row header (4:4) is trimmed to the columns in use and read as the data around it, not 16,384 columns", async ({ page }) => {
  const pane = await openPane(page, {
    workbook: selecting("4:4"),
    model: () => reply({ kind: "formula", formula: "=GROUPBY(C2:C501,F2:F501,SUM)", explanation: "Totals Amount by Region." }),
  });
  await enterKey(page);
  await expect(page.getByText("Using Orders, A1:F501 · 500 rows")).toBeVisible();
  await expect(page.getByText("From the block of filled cells around 4:4.")).toBeVisible();
  expect(await largestContentLoad(page)).toBeLessThanOrEqual(3006);

  await preview(page, "Total Amount per Region");
  await expect(page.getByText("Checked: no private values found")).toBeVisible();
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByRole("heading", { name: "Formula" })).toBeVisible();
  expect(pane.sent).toHaveLength(1);
  const turn = userTurn(pane.sent[0]!.body) as { range: { address: string }; columns: unknown[] };
  expect(turn.range.address).toBe("A1:F501");
  expect(turn.columns).toHaveLength(6);
  expect(pane.sent[0]!.body.length).toBeLessThan(20_000);
  expect(pane.consoleErrors).toEqual([]);
});

/** Orders with row 250 blank, selected at C10; with `filtered`, an AutoFilter on A1:F501. */
function withBlankRow(filtered: boolean): FakeWorkbook {
  const workbook = selecting("C10");
  const sheet = workbook.sheets.find((s) => s.name === "Orders")! as FakeWorkbook["sheets"][number] & { autoFilter?: string };
  for (const col of "ABCDEF") delete sheet.cells[`${col}250`];
  if (filtered) sheet.autoFilter = "A1:F501";
  return workbook;
}

test("a cell in a filtered range stands for the whole filtered range, across a blank row", async ({ page }) => {
  await openPane(page, { workbook: withBlankRow(true) });
  await enterKey(page);
  await expect(page.getByText("Using Orders, A1:F501 · 500 rows")).toBeVisible();
  await expect(page.getByText("From the filtered range around C10.")).toBeVisible();
});

test("without a filter, a blank row ends the block of filled cells", async ({ page }) => {
  await openPane(page, { workbook: withBlankRow(false) });
  await enterKey(page);
  await expect(page.getByText("Using Orders, A1:F249 · 248 rows")).toBeVisible();
  await expect(page.getByText("From the block of filled cells around C10.")).toBeVisible();
});

test("a cell in a table stands for the table, without its totals row", async ({ page }) => {
  const workbook = selecting("B7");
  const sheet = workbook.sheets.find((s) => s.name === "Orders")!;
  sheet.cells.E502 = { value: "Total", text: "Total", type: "String", format: "General" };
  sheet.cells.F502 = { value: 1, text: "1.00", type: "Double", format: "#,##0.00" };
  workbook.tables = [{ name: "Sales", sheet: "Orders", address: "A1:F502", showTotals: true }];
  await openPane(page, { workbook });
  await enterKey(page);
  await expect(page.getByText("Using Orders, A1:F501 · 500 rows")).toBeVisible();
  await expect(page.getByText("From the table Sales around B7.")).toBeVisible();
  // Named once, in the line above.
  await expect(page.getByText("· table Sales")).toHaveCount(0);
});

test("Use exactly the selected cells reads the clicked cells themselves", async ({ page }) => {
  const pane = await openPane(page, { workbook: selecting("E4") });
  await enterKey(page);
  await expect(page.getByText("Using Orders, A1:F501 · 500 rows")).toBeVisible();
  // Excel's selection has moved on; the exact cells are the ones the range came from.
  await select(page, "A1:B3");
  await page.getByRole("button", { name: USE_EXACT }).click();
  await expect(page.getByText("Using Orders, E4 · 0 rows")).toBeVisible();
  await expect(page.getByText("From the block of filled cells around E4.")).toHaveCount(0);
  await expect(page.getByRole("button", { name: USE_EXACT })).toHaveCount(0);
  await expect(page.getByText(KEEPS)).toBeVisible();

  // One header cell and no rows: nothing to compute on, so nothing is prepared.
  await page.getByRole("tab", { name: "Ask" }).click();
  await page.getByLabel("Question").fill("Total Amount");
  await page.getByRole("button", { name: "Preview" }).click();
  await expect(page.getByText("This range has no data rows. Select your data, or one cell in it, and choose Refresh from selection.")).toBeVisible();
  await expect(page.getByRole("heading", { name: "What gets sent" })).toHaveCount(0);
  expect(pane.sent).toEqual([]);
  expect(await writes(page)).toEqual([]);
  expect(pane.consoleErrors).toEqual([]);
});
