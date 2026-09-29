// End-to-end regressions from the red-team review, against the built pane (dist/) under its real
// CSP, with the fake Excel host and a scripted model. Build first: `npm run build`.
import { expect, test, type Locator, type Page } from "@playwright/test";
import { ORDERS, enterKey, leakedValues, openPane, preview, reply, userTurn, type FakeWorkbook } from "./harness";

interface Fake {
  writes: unknown[];
  getCell(sheet: string, addr: string): { value?: unknown; formula?: string } | null;
  setCell(sheet: string, addr: string, cell: unknown): void;
}

function writes(page: Page): Promise<unknown[]> {
  return page.evaluate(() => (window as unknown as { __NYMFORM_FAKE__: Fake }).__NYMFORM_FAKE__.writes);
}

function cellOf(page: Page, addr: string) {
  return page.evaluate((a) => (window as unknown as { __NYMFORM_FAKE__: Fake }).__NYMFORM_FAKE__.getCell("Orders", a), addr);
}

function setText(page: Page, addr: string, text: string) {
  return page.evaluate(
    ([a, t]) => (window as unknown as { __NYMFORM_FAKE__: Fake }).__NYMFORM_FAKE__.setCell("Orders", a!, { value: t, text: t, type: "String" }),
    [addr, text] as const,
  );
}

function column(page: Page, letter: string) {
  return page.getByRole("listitem", { name: `Column ${letter}` });
}

test("a selection without its header row: row 1 is treated as data and none of it is sent", async ({ page }) => {
  const workbook: FakeWorkbook = { ...ORDERS, selection: { sheet: "Orders", address: "A2:F501" } };
  const pane = await openPane(page, { workbook });
  await enterKey(page);
  await expect(page.getByText("Using Orders, A2:F501 · 500 rows")).toBeVisible();
  await expect(page.getByLabel("This range has a header row")).not.toBeChecked();
  await expect(page.getByText("Row 2 looks like data, so it's treated as data. Tick the box if it is a header row.")).toBeVisible();
  // Keep the personal columns private, whatever was suggested.
  for (const letter of ["A", "B"]) {
    const box = column(page, letter).getByRole("checkbox", { name: "Keep private" });
    if (!(await box.isChecked())) await box.check();
  }

  await preview(page, "Total Amount per Region");
  await expect(page.getByText("Checked: no private values found")).toBeVisible();
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByRole("heading", { name: "Formula" })).toBeVisible();
  expect(pane.sent).toHaveLength(1);
  expect(leakedValues(pane.sent[0]!.body)).toEqual([]);
  const turn = userTurn(pane.sent[0]!.body) as { range: { header_row: unknown }; columns: { header: string }[] };
  expect(turn.range.header_row).toBeNull();
  expect(turn.columns.map((c) => c.header)).toEqual(["Column A", "Column B", "Column C", "Column D", "Column E", "Column F"]);

  // Ticking the box is the user's call; the email column's header, which looks like data, gets a neutral one.
  await page.getByRole("tab", { name: "Columns" }).click();
  await page.getByLabel("This range has a header row").check();
  await expect(page.getByText(/looks like data, so it's treated as data/)).toHaveCount(0);
  await expect(column(page, "B").getByRole("textbox").first()).toHaveValue("Column B");
});

test("Fill down starts at the first data row, whatever row the model proposed", async ({ page }) => {
  await openPane(page, {
    model: () => reply({ kind: "formula", formula: "=F2*2", placement: { cell: "G1", fill_down: true }, explanation: "Doubles Amount on each row." }),
  });
  await enterKey(page);
  await preview(page, "Double the amount on each row");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("Formula check passed: allowed functions and your ranges only.")).toBeVisible();
  await expect(page.getByLabel("Insert in cell")).toHaveValue("G2");
  await expect(page.getByRole("switch", { name: "Fill down to row 501 (G2:G501)" })).toBeChecked();

  // The model's row, typed back in, is refused and nothing is written.
  await page.getByLabel("Insert in cell").fill("G1");
  await expect(page.getByRole("switch", { name: "Fill down to row 501 (G1:G501)" })).toBeVisible();
  await page.getByRole("button", { name: "Insert formula" }).click();
  await expect(page.getByText("Fill down starts at the first data row, G2. Change the cell or turn off Fill down.")).toBeVisible();
  expect(await writes(page)).toEqual([]);

  await page.getByLabel("Insert in cell").fill("G2");
  await page.getByRole("button", { name: "Insert formula" }).click();
  await expect(page.getByText("Inserted in G2:G501.")).toBeVisible();
  expect((await cellOf(page, "G1"))?.formula).toBeUndefined();
  expect((await cellOf(page, "G2"))?.formula).toBe("=F2*2");
  expect((await cellOf(page, "G501"))?.formula).toBe("=F501*2");
});

test("a formula that reaches outside the ranges once filled down is blocked at Insert", async ({ page }) => {
  await openPane(page, {
    model: () => reply({ kind: "formula", formula: "=F2/SUM(F2:F501)", placement: { cell: "G2", fill_down: true }, explanation: "Share of the total." }),
  });
  await enterKey(page);
  await preview(page, "Share of the total amount on each row");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("Formula check passed: allowed functions and your ranges only.")).toBeVisible();
  await page.getByRole("button", { name: "Insert formula" }).click();
  await expect(page.getByText("Blocked: when filled down, this formula refers outside your ranges.")).toBeVisible();
  await expect(page.getByText(/^Filled down/).first()).toBeVisible();
  expect(await writes(page)).toEqual([]);
});

test("a confirmation to replace cells covers only the cells it named", async ({ page }) => {
  await openPane(page, {
    model: () => reply({ kind: "formula", formula: "=F2*2", placement: { cell: "H2", fill_down: false }, explanation: "Doubles Amount." }),
  });
  for (let r = 2; r <= 6; r++) await setText(page, `H${r}`, `keep H${r}`);
  await setText(page, "I2", "keep I2");
  await enterKey(page);
  await preview(page, "Double the amount");
  await page.getByRole("button", { name: "Send" }).click();
  await page.getByRole("button", { name: "Insert formula" }).click();
  await expect(page.getByText("H2 already has values. Insert anyway to replace them; Excel can't undo this.")).toBeVisible();

  // Turning on Fill down withdraws the question about H2 alone.
  await page.getByRole("switch", { name: /Fill down to row 501/ }).check();
  await expect(page.getByText(/already has values/)).toHaveCount(0);
  await expect(page.getByRole("button", { name: "Replace and insert" })).toHaveCount(0);
  await page.getByRole("button", { name: "Insert formula" }).click();
  await expect(page.getByText("H2:H501 already has values. Insert anyway to replace them; Excel can't undo this.")).toBeVisible();

  // So does changing the cell.
  await page.getByLabel("Insert in cell").fill("I2");
  await expect(page.getByRole("button", { name: "Replace and insert" })).toHaveCount(0);
  await page.getByRole("switch", { name: /Fill down to row 501/ }).uncheck();
  await page.getByRole("button", { name: "Insert formula" }).click();
  await expect(page.getByText("I2 already has values. Insert anyway to replace them; Excel can't undo this.")).toBeVisible();
  expect(await writes(page)).toEqual([]);
  for (const a of ["H2", "H3", "I2"]) expect((await cellOf(page, a))?.value).toBe(`keep ${a}`);

  await page.getByRole("button", { name: "Replace and insert" }).click();
  await expect(page.getByText("Inserted in I2.")).toBeVisible();
  for (const a of ["H2", "H3"]) expect((await cellOf(page, a))?.value).toBe(`keep ${a}`);
});

test("the API key never appears in a value attribute or the page's markup", async ({ page }) => {
  // A fake key in the OpenRouter shape, assembled at runtime so the secret scanner still flags real ones.
  const key = ["sk", "or", "v1", "0123456789abcdef".repeat(3)].join("-");
  const pane = await openPane(page);
  const keyInDom = (k: string) =>
    page.evaluate(
      (needle) => ({
        attribute: [...document.querySelectorAll("*")].some((el) => [...el.attributes].some((a) => a.value.includes(needle))),
        markup: document.documentElement.outerHTML.includes(needle),
      }),
      k,
    );
  await page.getByLabel("API key").fill(key);
  expect(await keyInDom(key)).toEqual({ attribute: false, markup: false });
  await page.getByRole("button", { name: "Continue" }).click();
  await expect(page.getByRole("heading", { name: "Columns" })).toBeVisible();

  // Back in Setup the key is still there, as the field's value only.
  await page.getByRole("tab", { name: "Setup" }).click();
  await expect(page.getByLabel("API key")).toHaveValue(key);
  expect(await keyInDom(key)).toEqual({ attribute: false, markup: false });

  await page.getByRole("tab", { name: "Columns" }).click();
  await preview(page, "Total Amount per Region");
  expect(await keyInDom(key)).toEqual({ attribute: false, markup: false });
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByRole("heading", { name: "Formula" })).toBeVisible();
  expect(pane.sent[0]!.headers["authorization"]).toBe(`Bearer ${key}`);
  expect(await keyInDom(key)).toEqual({ attribute: false, markup: false });
});

test("text fields opt out of browser spell check and writing aids, and the page out of translation", async ({ page }) => {
  await openPane(page, {
    model: () => reply({ kind: "clarify", formula: null, explanation: "Per region or one grand total?" }),
  });
  const optedOut = (field: Locator) =>
    field
      .first()
      .evaluate((el) => ({
        spellcheck: el.getAttribute("spellcheck"),
        autocorrect: el.getAttribute("autocorrect"),
        autocapitalize: el.getAttribute("autocapitalize"),
        gramm: el.getAttribute("data-gramm"),
        suggestions: el.getAttribute("writingsuggestions"),
      }));
  const off = { spellcheck: "false", autocorrect: "off", autocapitalize: "off", gramm: "false", suggestions: "false" };
  expect(await page.evaluate(() => document.documentElement.getAttribute("translate"))).toBe("no");
  expect(await page.locator('meta[name="google"][content="notranslate"]').count()).toBe(1);
  await enterKey(page);
  expect(await optedOut(page.getByPlaceholder("For example: USD"))).toEqual(off);
  await preview(page, "Total Amount");
  await page.getByRole("button", { name: "Back" }).click();
  expect(await optedOut(page.getByLabel("Question"))).toEqual(off);
  await page.getByRole("button", { name: "Preview" }).click();
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("Per region or one grand total?")).toBeVisible();
  expect(await optedOut(page.getByLabel("Your answer"))).toEqual(off);
});
