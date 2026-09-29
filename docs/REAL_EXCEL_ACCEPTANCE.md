# Real Excel acceptance record

First run in a real Excel host: 2026-09-28, on Jorge's Mac, with synthetic workbooks only. This file
separates three kinds of evidence and never relabels one as another:

- **Real Excel, fixture model**: the add-in in Excel for Mac, its whole production path (payload,
  audit, the one sender's `fetch`, reply parsing, gate, restore, Insert), with the model endpoint
  pointed at a local fixture server that returns scripted replies. It proves what Excel does with a
  reply. It says nothing about what a model would reply.
- **Real Excel, live provider**: the same, against OpenRouter. See "Live provider" below.
- **Fake host**: the unit and end-to-end suites (`tests/`), which drive a fake Excel
  (`tests/e2e/fakeOffice.js`). They are the regression net, not evidence about Excel.

## Environment

| | |
| --- | --- |
| Host | Excel for Mac 16.113.2 (Microsoft 365), macOS 26.6.2 (25G83), Apple silicon |
| Add-in builds | `9860b17` (5-row run) and `e2b1d22` (orders run), `npm run dev-server` (development mode, bench off) |
| Endpoint (fixture runs) | `https://localhost:41952/api/v1`, a build-time setting like any other (`NYMFORM_ENDPOINT`); CSP `connect-src https://localhost:41952` only |
| Fixture model | A local HTTPS server outside this repository, labelled as a fixture in every reply ("FIXTURE REPLY (scripted, not a model)"). It saves each request body it receives, never headers. |
| Sideloading | A separate acceptance manifest (its own add-in ID, port 41961, label "Nymform test") placed in Excel's add-in folder (`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`); Excel restarted. `npm start` itself wasn't used. |
| Workbooks | `sales-summary-5row.xlsx`: the walkthrough's five rows (Customer, Email, Region, Amount; example.com addresses), generated with exceljs. `orders-acceptance.xlsx`: a copy of `bench/workbooks/orders.xlsx` (500 synthetic rows). Expected answers computed independently with exceljs, not with Excel. |

## Results: real Excel, fixture model

| # | Check | Expected | Observed | Result |
| --- | --- | --- | --- | --- |
| 1 | Pane loads under the release CSP in Excel for Mac's webview (spike S1, S3) | Setup screen, endpoint read-only, footer with the build | "Nymform for Excel 0.1.0-alpha · build 9860b17 · SkySpan"; endpoint shown read-only | Pass |
| 2 | Selection resolution, 5-row sheet, A1 selected | A1:D6, 5 rows | "Using Orders, A1:D6 · 5 rows. From the block of filled cells around A1." | Pass |
| 3 | Default private columns, 5-row sheet | Customer, Email (as in the video) | Customer, Email **and Region** ("Many different values, so it is kept private to be safe": 3 distinct in 5 rows) | Differs from the video; see DEMO_PARITY.md |
| 4 | Nothing sent before Send | No request at the endpoint | Fixture received nothing until Send | Pass |
| 5 | Sent body = previewed body, byte for byte | Identical | "Show exact request" copied from the pane and the fixture's received body: identical, 3,115 bytes, SHA-256 `371d2d57…b374` | Pass |
| 6 | Network failure | Plain message, nothing inserted | The fixture crashed on the first try (a path bug in the fixture itself); the pane said "Couldn't reach localhost:41952." Send again worked. | Pass (failure path) |
| 7 | Region totals, 5 rows, sample-rows mode, `LET/UNIQUE/MAP` inserted at F1 | West 3,420 · East 2,480 · North 940 | Same; spill F1:G3; "Inserted in F1." as a success. Saved file read back: A1:D6 unchanged, F1 an array formula over F1:G3 | Pass |
| 8 | Orders, 500 rows, structure only, a customer named in the question | Sent as a stand-in; restored; total 3,617.68 | Task sent as "Total Amount for PERSON_001"; formula `=SUMIFS(F2:F501,A2:A501,"Faustino Torphy")`; H2 = 3,617.68 | Pass |
| 9 | GROUPBY by Region | Central 39,657.34 · East 56,437.17 · North 44,508.24 · South 37,601.09 · West 35,741.64 | Same (plus Excel's own Total row, 213,945.50) | Pass |
| 10 | Insert, then ask again (build `9860b17`) | Next Send goes out | **Blocked: "The data changed after this preview".** A1:F501 was unchanged (saved workbook compared with its source). Cause: Excel adds hidden names (`_xlfn.SUMIFS`, `_xlfn.GROUPBY`, `_xleta.SUM`) when such a formula is written, and the new check compared them. | **Fail → fixed in `e2b1d22`** |
| 11 | Same sequence on `e2b1d22` | Next Send goes out | Insert at L2, then the next Send went out | Pass |
| 12 | Formula whose result is an Excel error (`=F2/0` at M2) | Written, reported as a warning | M2 shows `#DIV/0!`; "Inserted in M2, but Excel shows an error in the result (such as #N/A or #SPILL!). Check the cells before relying on them." (warning) | Pass (finding A) |
| 13 | Blocked spill (`=UNIQUE(C2:C501)` at P2, text in P3) | Written, reported as a warning | P2 shows `#SPILL!`; the same warning | Pass (finding A) |
| 14 | Fill down (`=IF(F2>500,"Yes","No")`, O2:O501) | 111 Yes, 389 No, nothing below row 501 | "Inserted in O2:O501."; 111 / 389; every row agrees with Amount > 500; O502 empty | Pass |
| 15 | A formula needing a stand-in the session never created (`"PERSON_099"`) | Blocked; Insert and Copy off | "Formula check passed" and "Blocked: the formula uses a stand-in this session didn't create (PERSON_099)…"; both buttons disabled | Pass (finding B) |
| 16 | Unreadable reply | One correction, shown first, sent only on Send | "Asking once more" preview, checked, 8,428 bytes; nothing sent until Send; then exactly one request of 8,428 bytes | Pass (finding D) |
| 17 | A private value typed into the sheet after the preview ("Zed Quux" into A3), then Send | Refused, nothing sent | "The data changed after this preview … so nothing was sent."; Columns showed the stale banner; no received body contains "Zed Quux" | Pass (finding C) |
| 18 | Refresh, then the same question | The new name goes out as a stand-in | Task sent as "Total Amount for PERSON_001"; the formula restored with "Zed Quux" | Pass |
| 19 | A cell edited outside the range (P3) | No stale verdict | The next Send went out | Pass |
| 20 | Independent leak check over every body the fixture received | No private value or canary | 11 bodies, 368 private values and canaries of 4+ characters (the generator's list plus the 5-row names and emails): 0 found | Pass |

## Not verified in real Excel

- Manual calculation (the "isn't calculated or checked yet" outcome), a formula refused by the gate
  in the pane (WEBSERVICE), the whole-column note, a lookup range, and a table as the source range:
  fake host only.
- Excel on the web and Excel for Windows: not run here (the website notes earlier use of Excel on the
  web; this run adds nothing to it).
- `npm start` sideloading, and the release build from GitHub Pages (`manifest.release.xml`).
- Performance: not measured. Selection, preview, audit, request and insert times on small, medium and
  near-limit ranges still need timing in Excel before any claim about speed.
- Table auto-expansion: whether inserting into the column next to an Excel table grows the table
  (which the Send/Insert check would then report as a change) wasn't tried.

## Live provider

Run on 2026-09-28, 12:37–12:44 PT, build `e2b1d22` served with the release endpoint
(`https://openrouter.ai/api/v1`, CSP `connect-src https://openrouter.ai`), in Excel for Mac 16.113.2.
Jorge entered his capped, short-lived OpenRouter key in the pane himself. The same key had also been
pasted into the chat that ran this session, so it is **treated as exposed**: it is not to be used
again, and no further paid calls are made without Jorge's authorization and a fresh key entered
outside the chat. Closing the pane afterwards cleared only the pane's own memory; it neither revoked
the key nor removed copies from the chat transcript. No part of the key appears in this record.
Workbook: five synthetic rows (the walkthrough's), A1:D6, Customer and Email private. Models: the
default `openai/gpt-6-luna` and two Anthropic models listed with zero-data-retention endpoints that day
(`anthropic/claude-sonnet-5.5`, `anthropic/claude-opus-5.5`). The conversation was cleared between
models, so no model saw another's exchange.

| # | Model | Mode | Question | Reply (formula as returned) | Provider | Bytes sent | Inserted | Excel result | Expected |
| --- | --- | --- | --- | --- | --- | ---: | --- | --- | --- |
| L1 | gpt-6-luna | Structure only | Total Amount per Region, one row per region | `=LET(r,UNIQUE(C2:C6),HSTACK(r,SUMIF(C2:C6,r,D2:D6)))` | Azure | 2,855 | F1 | West 3,420 · East 2,480 · North 940 | Same |
| L2 | gpt-6-luna | Sample rows | Total Amount for Maya Chen | An answer, not a formula: "Maya Chen has a total Amount of 1,240." (the name restored in the pane only; the request carried `PERSON_001`) | Azure | 4,450 | — | — | 1,240 |
| L3 | claude-sonnet-5.5 | Structure only | as L1 | `=LET(regions,UNIQUE(C2:C6),HSTACK(regions,SUMIF(C2:C6,regions,D2:D6)))` | Google | 2,836 | I1 | West 3,420 · East 2,480 · North 940 | Same |
| L4 | claude-opus-5.5 | Structure only | as L1 | `=LET(r,UNIQUE(Orders!C2:C6),HSTACK(r,SUMIFS(Orders!D2:D6,Orders!C2:Orders!C6,r)))` | Amazon Bedrock | 2,834 | L1 | West 3,420 · East 2,480 · North 940 | Same |
| L5 | claude-opus-5.5 | Structure only | Which region has the highest total Amount? Return just the region name. | `=LET(r,C2:C6,a,D2:D6,u,UNIQUE(r),t,SUMIFS(a,r,u),INDEX(SORTBY(u,t,-1),1))` | Amazon Bedrock | 2,862 | O1 | West | West |
| L6 | gpt-6-luna | Structure only | as L5 | `=LET(r,UNIQUE(C2:C6),t,SUMIF(C2:C6,r,D2:D6),INDEX(r,XMATCH(MAX(t),t)))` | Azure | 2,883 | P1 | West | West |
| L7 | claude-sonnet-5.5 | Structure only | as L5 | `=LET(r,UNIQUE(C2:C6),t,SUMIF(C2:C6,r,D2:D6),INDEX(r,MATCH(MAX(t),t,0)))` | Google | 2,864 | Q1 | West | West |

Every formula passed the gate before and after restoring and was inserted with "Inserted in …" as a
success. The saved workbook, read back with exceljs, holds the six formulas with those results, and
A1:D6 unchanged. Seven live requests in all.

What this shows: seven successful requests across three task types (totals by region, the total for
one named customer, and the region with the highest total) on the tested synthetic workbook, in Excel
for Mac 16.113.2. It is not an accuracy rate, not a benchmark or comparison between models, and no
evidence about Excel on the web or Windows. The benchmark (`bench/`, 20 tasks) has still not been run
against a live model. Luna answered L2 directly from the sample rows, which the spec allows in that
mode; an answer is text in the pane, not a checked formula.

### Zero-data-retention evidence: captured, inferred, unavailable

OpenRouter documents that with `provider.zdr` set to `true` "the request will only be routed to
endpoints that have a Zero Data Retention policy", that the policy is tracked per endpoint rather
than per provider, and that in-memory prompt caching is still allowed under ZDR
(<https://openrouter.ai/docs/guides/features/zdr>, read 2026-09-28). A provider name on its own says
nothing about an endpoint's retention policy.

| Item | Status | Where |
| --- | --- | --- |
| `provider: { zdr: true }` in each request body | Captured for L1–L4: read on What gets sent (request settings line) before Send. For L5–L7 the preview wasn't read before Send; the builder adds the field unconditionally for the `openrouter.ai` endpoint, so it is expected there, not observed | `src/core/payload.ts` (`provider: { zdr: true }` when the endpoint host is OpenRouter); tests T-P2 (`tests/payload.test.ts`, `tests/provider.test.ts`) |
| Model ID of each request | Captured (What gets sent and the table above) | this record |
| Serving provider per request | Captured as the provider name OpenRouter returned (`X-Provider-Name`), read from the pane's Log and transcribed; the Log wasn't exported | table above |
| ZDR endpoint list at the time | Captured: `https://openrouter.ai/api/v1/endpoints/zdr` fetched at 12:36:31 PT, one minute before the first request; full snapshot SHA-256 `65dc74e9…ad67`. It lists `openai/gpt-6-luna` on `azure`, `azure/us`, `azure/eu`; `anthropic/claude-sonnet-5.5` on `google-vertex/global`, `google-vertex/us`, `google-vertex/europe`, `amazon-bedrock`; `anthropic/claude-opus-5.5` on `amazon-bedrock`, `amazon-bedrock/us-east-1`, `amazon-bedrock/eu-west-1`, `google-vertex/global`, `google-vertex/us`, `google-vertex/europe` | review handoff, `evidence/zdr/` |
| That each request was served by a ZDR-listed endpoint | **Inferred** from OpenRouter's documented `provider.zdr` routing and the snapshot above, not observed | — |
| The exact endpoint (region, tag) that served each request | Unavailable: the returned provider name gives the brand only | — |
| OpenRouter generation IDs | Unavailable here: the pane's Log held each raw response (which carries the ID) in memory only; it wasn't exported before the pane closed. The owner can look them up on OpenRouter's Activity page for 12:37–12:44 PT on 2026-09-28 | — |
| Account-level ZDR setting | Not checked (the request-level setting applies either way; OpenRouter combines them with OR) | — |

## A hidden-name experiment (security, 2026-09-28)

Question: Excel stores the `SUM` passed to `GROUPBY(..., SUM)` as `_xleta.SUM`. Could a crafted
workbook that defines a name `_xleta.SUM` make that formula call something else, past the gate (which
allows `SUM` unless a defined name `SUM` exists)?

Test: a synthetic workbook generated with exceljs, with `definedName` entries added to
`xl/workbook.xml`: `_xleta.SUM` = `LAMBDA(x, 999)`, and a control `Canary` = `LAMBDA(x, 777)`. Opened
in Excel for Mac 16.113.2 (no repair prompt), then typed into cells:

| Formula | Result |
| --- | --- |
| `=GROUPBY(C2:C6,D2:D6,SUM)` | East 2,480 · North 940 · West 3,420 · Total 6,840: the real sums, not 999 |
| `=Canary(1)` (control) | 777: names defined in the file are live |
| `=SUM(1,2)` | 3 |

After saving, the file's `definedNames` held only `Canary`: Excel dropped `_xleta.SUM`.

The conclusion is limited to what was tested: Excel for Mac 16.113.2, one crafted workbook defining
`_xleta.SUM`, and `GROUPBY` with the eta-reduced `SUM`. There, the crafted name neither redirected the
function nor survived a save. It says nothing about other reserved or hidden names (other `_xleta.` or
`_xlfn.` functions, sheet-scoped names), other functions that take a function argument, Excel on the
web or Windows, or other Excel versions.
