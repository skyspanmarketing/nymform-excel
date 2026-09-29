// End-to-end checks for the release-hardening review (2026-09-28), against the built pane (dist/)
// under its real CSP, with the fake Excel host and a scripted model. Build first: `npm run build`.
import { expect, test, type Page } from "@playwright/test";
import { enterKey, openPane, preview, reply } from "./harness";

interface Fake {
  writes: unknown[];
  evaluate: ((formula: string, sheet: string, addr: string) => unknown) | null;
  setCell(sheet: string, addr: string, cell: unknown): void;
  getCell(sheet: string, addr: string): { formula?: string } | null;
}

const fake = <T>(page: Page, fn: (f: Fake) => T) =>
  page.evaluate(`(${fn.toString()})(window.__NYMFORM_FAKE__)`) as Promise<T>;

test("an inserted formula that Excel shows as an error is reported as a warning, not as success", async ({ page }) => {
  await openPane(page, {
    model: () => reply({ kind: "formula", formula: "=F2/0", placement: { cell: "G2", fill_down: false }, explanation: "Divides." }),
  });
  await fake(page, (f) => {
    f.evaluate = () => ({ value: "#DIV/0!", type: "Error" });
  });
  await enterKey(page);
  await preview(page, "Divide each amount by zero");
  await page.getByRole("button", { name: "Send" }).click();
  await page.getByRole("button", { name: "Insert formula" }).click();
  await expect(page.getByText("Inserted in G2, but Excel shows an error in the result (such as #N/A or #SPILL!). Check the cells before relying on them.")).toBeVisible();
  await expect(page.getByText("Inserted in G2.", { exact: true })).toHaveCount(0);
  expect((await fake(page, (f) => f.getCell("Orders", "G2")))?.formula).toBe("=F2/0");
});

test("two clicks on Insert formula before the first insert finishes write once and report the insert", async ({ page }) => {
  await openPane(page, {
    model: () => reply({ kind: "formula", formula: "=SUMIFS(F2:F501,C2:C501,C2)", placement: { cell: "H2", fill_down: false }, explanation: "x" }),
  });
  await enterKey(page);
  await preview(page, "Total Amount for this row's Region");
  await page.getByRole("button", { name: "Send" }).click();
  // The fake host completes an insert within a few milliseconds, so a Playwright double click is
  // two inserts in a row (the second rightly asks before replacing H2). Real Excel takes far longer,
  // and the second click of a double click lands while the first insert is still in flight: two
  // clicks in one task model that.
  await page.getByRole("button", { name: "Insert formula" }).evaluate((b: HTMLButtonElement) => {
    b.click();
    b.click();
  });
  await expect(page.getByText("Inserted in H2.", { exact: true })).toBeVisible();
  await expect(page.getByRole("button", { name: "Insert formula" })).toBeEnabled();
  expect(await fake(page, (f) => (f.writes as { kind: string }[]).filter((w) => w.kind === "formulas").length)).toBe(1);
  await expect(page.getByText(/already has values/)).toHaveCount(0);
});

test("a formula that needs a stand-in this session didn't create can't be inserted or copied", async ({ page }) => {
  await openPane(page, {
    model: () => reply({ kind: "formula", formula: '=SUMIFS(F2:F501,A2:A501,"PERSON_099")', placement: { cell: "H2", fill_down: false }, explanation: "x" }),
  });
  await enterKey(page);
  await preview(page, "Total Amount per Region");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText(/^Blocked: the formula uses a stand-in this session didn't create \(PERSON_099\)/)).toBeVisible();
  await expect(page.getByRole("button", { name: "Insert formula" })).toBeDisabled();
  await expect(page.getByRole("button", { name: "Copy" })).toBeDisabled();
  expect(await fake(page, (f) => f.writes)).toEqual([]);
});

test("an unreadable reply comes back as a correction to preview; it is sent only on Send", async ({ page }) => {
  const pane = await openPane(page, { model: () => "Sure! Here is your formula: SUM(F2:F501)" });
  await enterKey(page);
  await preview(page, "Total Amount");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("Asking once more")).toBeVisible();
  await expect(page.getByText("Checked: no private values found")).toBeVisible();
  expect(pane.sent).toHaveLength(1);
  await page.getByText("Show exact request").click();
  const exact = await page.getByLabel("Exact request", { exact: true }).innerText();
  expect(exact).toMatch(/Your previous reply could not be read/);

  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("The model's reply couldn't be read. Try rephrasing.", { exact: false })).toBeVisible();
  expect(pane.sent).toHaveLength(2);
  expect(pane.sent[1]!.body).toBe(exact);
});

test("a cell changed in Excel after the preview: Send refuses, nothing is sent, and Columns says to refresh", async ({ page }) => {
  const pane = await openPane(page);
  await enterKey(page);
  await preview(page, "Total Amount per Region");
  await fake(page, (f) => f.setCell("Orders", "A3", { value: "Zed Quux", text: "Zed Quux", type: "String" }));
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText(/^The data changed after this preview/)).toBeVisible();
  expect(pane.sent).toEqual([]);
  await page.getByRole("tab", { name: "Columns" }).click();
  await expect(page.getByText("The data changed since Nymform read it. Select your data, or one cell in it, and choose Refresh from selection.")).toBeVisible();
  await page.getByRole("button", { name: "Refresh from selection" }).click();
  await expect(page.getByText(/^The data changed since Nymform read it/)).toHaveCount(0);
  // The old preview is still on Ask, marked out of date, as after any change to the columns.
  await page.getByRole("tab", { name: "Ask" }).click();
  await expect(page.getByText("The columns or the conversation changed after this preview. Preview again.")).toBeVisible();
  await page.getByRole("button", { name: "Back" }).click();
  await preview(page, "Total Amount per Region");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByRole("heading", { name: "Formula" })).toBeVisible();
  expect(pane.sent).toHaveLength(1);
});

test("a whole-column formula says it reads rows outside the selection, in Excel only", async ({ page }) => {
  await openPane(page, {
    model: () => reply({ kind: "formula", formula: "=SUMIFS(F:F,C:C,C2)", placement: { cell: "H2", fill_down: false }, explanation: "x" }),
  });
  await enterKey(page);
  await preview(page, "Total Amount for this row's Region");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText(/^Uses whole columns \(F:F, C:C\): in Excel, the formula also reads rows of those columns outside your selection/)).toBeVisible();
  await expect(page.getByRole("button", { name: "Insert formula" })).toBeEnabled();
});
