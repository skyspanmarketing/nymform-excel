# Walkthrough parity

The walkthrough `nymform-walkthrough-final.mp4` (58 s, 1920×1032) labels itself "Scripted, not a live
recording · Synthetic data". It was reviewed from frames taken every 2 seconds on 2026-09-28. It is a
reference for what the user should understand and control, not evidence of what the add-in does.
Parity means the same workflow and the same correct result, not the same pixels or formula text.

Classification: **implemented** (in the add-in and checked in real Excel), **implemented, fake host
only**, **differs** (the add-in does it another way, on purpose or not), **conceptual only** (the
video shows something the add-in doesn't and shouldn't do as shown).

| Time | Video | Add-in | Evidence | Class |
| --- | --- | --- | --- | --- |
| 0:00–0:03 | Title: "Ask for the formula. Decide what AI sees." | — | — | Marketing |
| 0:04 | Five rows A1:D6 (Customer, Email, Region, Amount); Columns: "Using Orders, A1:D6 · 5 rows" | `readSelection` resolves the block around the selected cell; the Columns summary | REAL_EXCEL_ACCEPTANCE.md #2: "Using Orders, A1:D6 · 5 rows. From the block of filled cells around A1." | Implemented |
| 0:04–0:18 | No column ticked at first; the user ticks Customer, then Email | Customer and Email are suggested private, so they start ticked; the checkbox decides | #3 | Differs (suggestions on by default; the video's manual ticking is the same control) |
| 0:04 | Region and Amount "Left visible for this example" | On this 5-row sheet Region is also suggested private ("Many different values, so it is kept private to be safe") | #3 | **Differs.** A 3-in-5 distinct ratio trips the high-cardinality rule. The user can untick it. Changing the rule for small samples is a Type 2 decision (vision.md) for Jorge |
| 0:06–0:12 | Ask: question "Total Amount per Region, one row per region"; Structure only / Include sample rows | Ask screen with the same two modes and question box | #7 | Implemented |
| 0:16–0:18 | Names and emails turn into PERSON_001… / EMAIL_001… **in the worksheet cells**, with coloured cells | Nothing in the sheet changes. Stand-ins exist only in the request (What gets sent shows them as chips) | #7 (A1:D6 unchanged after the run) | **Conceptual only, and must stay so.** Rewriting or recolouring source cells would be a write (invariant 9) that Excel can't undo, and vision.md rules it out ("Selection, not recolouring"). Don't build it to match the video |
| 0:22 | Region and Amount cells highlighted green as "visible" | No highlighting | — | Conceptual only (same reason) |
| 0:24–0:34 | What gets sent: "Checked: no private values found", 5 sample rows with stand-ins, Send and Back | Same screen; formatted view plus **Show exact request** (the literal string) | #4, #5 | Implemented (the add-in shows more: size in bytes, exact string, model and settings) |
| 0:34–0:36 | "Sending is simulated in this preview"; "Preparing the formula…" | A real request to the configured endpoint, with Cancel | #5 (fixture); live: pending | Implemented (fixture); live provider pending |
| 0:38–0:42 | Formula `=LET(r,C2:C6,a,D2:D6,u,UNIQUE(r),HSTACK(u,MAP(u,LAMBDA(x,SUMIF(r,x,a)))))`, "Formula check passed", Insert in cell F1 | Same (the fixture returned this formula); every function on the allowlist | #7 | Implemented (fixture) |
| 0:44–0:48 | Inserted in F1: West 3,420, East 2,480, North 940; "Formula inserted in F1. Excel calculates the result."; "The original cells stay unchanged." | "Inserted in F1." as a success only when Excel's result has no error; F1:G3 = West 3,420, East 2,480, North 940; A1:D6 unchanged | #7 | Implemented |
| 0:44 | Names shown back in the cells, greyed during insert | The sheet never showed stand-ins, so there is nothing to put back | — | Conceptual only |
| 0:50–0:58 | "AI for Excel. Control what you share."; "No live model request is made." | — | — | Marketing |

## What the video leaves out, which the add-in does

- The exact request string and its size; the Log with every request, its audit result and the
  provider that served it.
- A private value typed in the question goes out as its stand-in; a part of one is resolved or asked
  about.
- An Excel error or blocked spill in the result is a warning, not a success; a formula that needs a
  stand-in the session didn't create can't be inserted.
- If the sheet changes after the preview, nothing is sent until Refresh.

## Recommendation

Keep the video as a workflow illustration with its "scripted" label. If it is re-cut, show the
stand-ins in What gets sent rather than in the cells, since the in-sheet animation is the one thing
the add-in deliberately never does.
