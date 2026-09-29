# Demo script for a tester

About 10 minutes. Everything uses `bench/workbooks/orders.xlsx`, which is synthetic: 500 orders with
made-up customers at example.com addresses. Install first (README, "Install").

## 1. Set up (1 minute)

1. Open `orders.xlsx` in Excel and select `A1:F501` (click A1, then Ctrl+Shift+End).
2. Choose **Home → Nymform**. The pane opens on **Setup**.
3. Paste an OpenRouter key with a spending limit. Point out the line under it: the key goes only to
   openrouter.ai and is forgotten when the pane closes. The endpoint can't be edited.
4. Choose **Continue**.

## 2. What Nymform thinks is private (1 minute)

**Columns** shows `Using Orders, A1:F501 · 500 rows`. Customer and Email are marked **Keep private**, with
the layered-line pattern and a note saying why. Region, Order date, Product and Amount are not.

Say: the suggestions are only defaults; the checkbox decides. Tick Amount if you want to show that
numbers can be private too.

## 3. The main flow: structure only (3 minutes)

1. Go to **Ask**. Keep **Structure only**.
2. Type: `Total Amount per Region, one row per region`. Choose **Preview**.
3. **What gets sent**: read it together. Headers, types, counts and the question. No customer, no
   email, no amount. "Checked: no private values found", and the exact size in bytes.
4. Turn on **Show exact request**: this is the literal text that is checked and sent.
5. Choose **Send**. The **Result** shows a formula (often `=GROUPBY(C2:C501,F2:F501,SUM)`), "Formula
   check passed", and a short explanation.
6. Choose **Insert formula**. It lands in the first empty column to the right, and Excel computes it
   on the real data.

## 4. A private name in the question (2 minutes)

1. Choose **Ask another question**. Type: `Total Amount for Faustino Torphy` (a customer in row 2).
2. **Preview**: the name has been replaced by a stand-in chip such as `PERSON_001`. The model never
   sees the name.
3. **Send**. The formula comes back with `"PERSON_001"`, and Nymform puts the real name back inside
   the quotes: `=SUMIFS(F2:F501,A2:A501,"Faustino Torphy")`. The explanation shows the name too, but
   that restored text stays in the pane; the conversation kept for the model holds only stand-ins.
4. Insert it, and compare with a manual filter on Customer.

## 5. Part of a name, and what a block looks like (2 minutes)

1. **Ask another question**: `Orders from Faustino` (the first name only). Only one customer has it,
   so the preview says "Read “Faustino” as Faustino Torphy" and sends a stand-in such as `PERSON_001`.
2. Choose **Keep as typed**: now the preview shows "Blocked: a value from column A is in the
   request." and "Found in your question.", and Send is off. Nothing was sent: the check found part
   of a private value. The preview says only where; **Log** shows what (see 8).
3. Ask about a family name several customers share (in `orders.xlsx`, `Total Amount for Crist`).
   The pane lists every match with its stand-in and cell count. **Show in sheet** selects that
   customer's cells in Excel. Tick one or several and choose **Use this value**: the question goes
   out with their stand-ins, and Excel's selection goes back to your range.

## 6. Sample rows (1 minute)

1. In **Ask**, choose **Include sample rows** (50 rows by default).
2. **Preview**: the rows are there, with Customer and Email as stand-ins, and Region, Product and
   Amount as they are. On **Columns**, each private column now has a treatment picker: Stand-in,
   Range (numbers become `80000-89999`), Month (dates become `2026-03`) or Leave out.

## 7. Fill down, and a second insert (1 minute)

1. Ask: `Flag orders over 500 with Yes, otherwise No, in a new column`. The result usually has
   **Fill down to row 501** on.
2. Insert it. Then choose **Insert formula** again on the same cell: the pane asks before replacing
   values ("… already has values. Insert anyway to replace them; Excel can't undo this.").

## 8. The log (30 seconds)

**Log** lists every request with its size, check results and the provider that served it (Azure for
Luna with zero data retention). **Exact request** shows what was sent. For the request blocked in 5,
it shows what the check found and where ("Found “Faustino” in your question"), with the text around
it and why it counts. That is shown only in the pane: neither export includes it. **Export evaluation
report** holds only counts, sizes, timings and check results. Ask the tester to rate results with
**Worked / Partly / Wrong** as they go.

## What to ask the tester afterwards

- Was it clear what would leave your machine before you clicked Send?
- Did the private-column suggestions match what you'd have chosen?
- Did any formula come back wrong, or get blocked when it shouldn't have? (For a wrong formula, Log →
  Export log helps us see what happened; it holds only stand-ins, never the key. For a block, Export
  log leaves out what matched, so send a screenshot of the Log's "Found" lines instead. With
  `orders.xlsx` that's fine, since its data is made up; with real data, tell us only where it was
  found and the line under it, not the text.)
- Anything confusing in the wording?

## If something goes wrong

- The pane stays blank: the dev server isn't running, or the localhost certificate isn't trusted
  (`npm run certs`, then reopen Excel).
- "Your API key was rejected": check the key in Setup.
- "No zero-data-retention route can serve openai/gpt-6-luna right now": OpenRouter has no ZDR route
  up for Luna. Try again later, or choose another ZDR-eligible model in Setup. Nymform never falls
  back to a non-ZDR route.
- "Select fewer cells (limit 20,000)": select the data only, not whole sheets.
- "The data changed after this preview": a cell, header, table or defined name changed in Excel after
  Nymform read the range. Nothing was sent (or written). Select the data, choose Refresh from
  selection and preview again.
- "Asking once more": the model's reply couldn't be read. The follow-up is shown like any request and
  goes out only if you choose Send.
