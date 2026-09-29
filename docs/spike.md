# Day-one spike (spec §4)

Recorded 2026-09-24. The build ran in a Linux cloud container with no Excel and no OpenRouter key,
so each item says what was verified here and what still needs a person at Excel.

| Item | Status | Evidence |
| --- | --- | --- |
| S1. Sideload into Excel on the web and desktop | **Excel for Mac: verified 2026-09-28** (16.113.2, macOS 26.6.2; manifest in Excel's add-in folder, Excel restarted, add-in under Home → Add-ins → Developer Add-ins). Excel on the web and Windows: not run in this record. | Both manifests pass Microsoft's validator (`npm run validate`). The dev server serves the pane at `https://localhost:41951/taskpane.html`. Sideloading steps are in the README. |
| S2. POST to `${NYMFORM_ENDPOINT}/chat/completions` from the pane | **Verified in Chromium, and from Excel for Mac's webview to a local HTTPS fixture endpoint (CORS preflight, then POST; the body received was the audited one, byte for byte). OpenRouter from the Office webview: see `docs/REAL_EXCEL_ACCEPTANCE.md`, "Live provider".** | The built pane, under its real Content Security Policy in headless Chromium, sent a real request to `https://openrouter.ai/api/v1/chat/completions` (CORS preflight, then POST with `Authorization`, `Content-Type` and `X-Title`). An invalid key came back as a JSON 401, which the pane showed as "Your API key was rejected. Check it in Setup." (end-to-end test "live OpenRouter", `NYMFORM_E2E_LIVE=1`). With curl: preflight 204, `access-control-allow-origin: *`; OpenRouter exposes `X-Provider-Name` to the browser, which the Log shows. |
| S2, Luna | **Checked against OpenRouter's live data; a real model call needs a key.** | `openai/gpt-6-luna` is served by OpenAI, Azure and Amazon Bedrock. Only its Azure routes are on the zero-data-retention list (`/api/v1/endpoints/zdr`), so every `provider: { zdr: true }` Luna request goes to Azure. Luna is a reasoning model and doesn't list `temperature`; the Azure routes list `max_completion_tokens` rather than `max_tokens`. The request parameters were changed accordingly (`docs/decisions.md`). With a key, run `NYMFORM_E2E_LIVE=1 OPENROUTER_API_KEY=… npm run test:e2e`, and confirm in the pane's Log that the provider is Azure and that the reply parsed (finish reason `stop`, non-empty content). |
| S3. Content Security Policy | **Verified outside Excel, and in Excel for Mac: Office.js loads and the pane works under the policy.** | The release and dev pages carry the §7.13 policy with `connect-src` set to the endpoint origin only (`npm run check:release` fails otherwise). The real `office.js` loads under the policy (`office.js` and `o15apptofilemappingtable.js` from appsforoffice.microsoft.com). The one violation: Office.js tries to frame `https://telemetryservice.firstpartyapps.oaspapps.com/`, Microsoft's telemetry, and the policy refuses it. That is intended (invariant 7) and Office.js carried on. A test confirms the page can't connect to any other host. |
| S4. Read values, text, valueTypes, numberFormat; write a formula and autoFill 10 rows | **Verified in Excel for Mac: 500-row reads, a formula written and filled down O2:O501, spill and error readback (`docs/REAL_EXCEL_ACCEPTANCE.md`).** | `src/office/adapter.ts` implements it with the Excel API (ExcelApi 1.9; spill readback uses 1.12 when available). It is exercised against a fake Excel host that enforces Office's load/sync rules: the adapter tests, and the end-to-end tests that insert formulas. |

## First thing to do at Excel

1. `npm ci && npm run certs && npm run dev-server`, then sideload `manifest.dev.xml` (README).
2. Open the pane from **Home → Nymform**. If it stays blank, open the webview's developer tools
   (Excel on the web: the browser's own; Windows: right-click the pane → Inspect) and look for CSP
   errors on Office.js. Record them here before changing the policy. `connect-src` must stay the
   endpoint only (invariant 10). A refused `telemetryservice.firstpartyapps.oaspapps.com` frame is
   expected.
3. Open `bench/workbooks/orders.xlsx`, select `A1:F501`, and run the demo in the README with a capped
   key. Check the Log shows provider Azure.
4. Insert a fill-down formula (for example "Flag orders over 500 in a new column") to check autoFill.
