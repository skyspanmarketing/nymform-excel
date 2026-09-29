# NymForm for Excel

Ask a model for the formula you need without sending it your spreadsheet's private data. **Alpha.**

Protected values stay in Excel. NymForm replaces private data before AI analysis and shows you what will be sent.

> In structure-only mode, no cell value is included in the request. In substituted mode, every value from a column you mark private is replaced before sending. Before anything is sent, NymForm checks the exact request: every private cell is verified as replaced, and the full text is scanned for private values of 4 or more characters.

You select a range, mark private columns, and ask a question. NymForm sends the model a description
of the data's structure, or sample rows with private values swapped for stand-ins such as
`PERSON_014`. The model returns a formula. NymForm checks it, puts your real values back into it, and
Excel computes it on your data. Nothing is written until you click **Insert formula**.

## Watch the walkthrough

[![NymForm for Excel walkthrough](docs/walkthrough-poster.jpg)](https://vimeo.com/1230731041)

▶ [Watch the NymForm for Excel walkthrough](https://vimeo.com/1230731041) (58 seconds, on Vimeo).
It is a scripted product preview with synthetic data. Its in-sheet masking is a conceptual
animation: the add-in leaves your worksheet cells unchanged, and stand-ins appear only in the
request ([docs/DEMO_PARITY.md](docs/DEMO_PARITY.md)).

## What it doesn't protect

Headers, sheet names, notes you write and your question (with private values replaced) are always
sent. Values under 4 characters in your typed question, values encoded in ways the scan doesn't look
for, columns you don't mark private, re-identification from counts in small groups, and what the
model provider does while processing a request are not covered. Details: [SECURITY.md](SECURITY.md).

## Install the hosted release

The release build is served from GitHub Pages at <https://skyspanmarketing.github.io/nymform-excel/>.
Its manifest is <https://skyspanmarketing.github.io/nymform-excel/manifest.xml>, and
`build-info.json` at the same address names the commit it was built from. Nothing runs on your
computer except Excel: no local server and no Node.js. You need
Excel on the web or Excel for Windows or Mac (Excel JavaScript API 1.9 or later) and an
[OpenRouter](https://openrouter.ai) API key.

- **Excel for Mac** (the method used in the acceptance run below): download `manifest.xml` from the
  address above, copy it into `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` (create
  the `wef` folder if it isn't there), then quit and reopen Excel. Choose **Home → Add-ins**; the
  add-in is listed under **Developer Add-ins** as **Nymform for Excel**. Once it has been opened, a
  **Nymform** button also appears at the right end of the **Home** tab.
- **Excel on the web:** open a workbook, choose **Home → Add-ins → More Add-ins**, open **My
  Add-ins**, choose **Upload My Add-in**, and pick the downloaded `manifest.xml`. This route hasn't
  been exercised with the hosted build yet.
- **Excel for Windows:** not tested yet.

The pane's footer shows the version and the build commit (`0.1.0-alpha · build bfe578f` for this
release), so a hosted pane can be told apart from a local build.

## Hosted acceptance

On 2026-09-28 the hosted `v0.1.0-alpha` build (commit `bfe578f`) was installed by the Mac method above
in Excel for Mac 16.113.2 on macOS 26.6.2 and run once against a synthetic five-row workbook. The
substituted preview contained none of the workbook's names or emails, the structure-only preview
contained no cell values, one live request to the default model returned
`=SUMIF($C$2:$C$6,C2,$D$2:$D$6)`, the formula passed the gate and, filled down, gave the expected
region totals, the source cells were unchanged, and the key was gone after the pane was closed and
reopened. That is one smoke test on one host with synthetic data: not an accuracy measurement, not
a check of Excel for Windows or Excel on the web, and not a fitness claim for sensitive data (see
[SECURITY.md](SECURITY.md)).

## Run from source (developer setup)

You need Node.js 22.12 or later (`engines` in `package.json`),
Excel on the web or Excel for Windows or Mac, and an [OpenRouter](https://openrouter.ai) API key.
The commands below contain no comments, so each block can be pasted into a terminal as it is.

1. Open a terminal in the project folder (the one that holds `package.json`), then install and
   create the local certificate. This is needed once; answer yes, or enter your password, when asked
   to trust the certificate.

   ```sh
   npm ci
   npm run certs
   ```

2. Start the add-in's local server and leave that terminal window open while you use NymForm.
   It is ready when it prints `compiled successfully`. To check it, open
   <https://localhost:41951/taskpane.html> in a browser: it should load with no certificate warning.

   ```sh
   npm run dev-server
   ```

3. Load the add-in into Excel.

   **Excel on the web** (on the same computer): open a workbook, choose **Home → Add-ins → More
   Add-ins**, open **My Add-ins**, choose **Upload My Add-in**, and pick `manifest.dev.xml` from the
   project folder.

   **Excel for Windows or Mac:** in a second terminal window, in the same folder, run the command
   below. It sideloads `manifest.dev.xml` and opens Excel; `npm stop` removes it again. On a Mac you
   can instead copy `manifest.dev.xml` into the `wef` folder named above and restart Excel, which is
   how the Mac acceptance runs were sideloaded; `npm start` and `npm stop` haven't been verified on
   Mac or Windows.

   ```sh
   npm start
   ```

4. In Excel, choose **Nymform** at the right end of the **Home** tab. If it isn't there, look under
   **Home → Add-ins**. On a Mac, Excel picks up a newly sideloaded add-in when it starts (quit and
   reopen Excel if it was already open), and lists it there under **Developer Add-ins**; **Insert →
   My Add-ins** doesn't show it. The pane opens on **Setup**.

## OpenRouter setup

1. In your OpenRouter account's privacy settings, turn on zero data retention.
2. Create a key with a spending limit. For a team, give each person their own capped key rather than
   sharing one.
3. Paste the key into NymForm's **Setup** screen. It stays in memory until you close the pane.

The default model is `openai/gpt-6-luna`, requested with `provider: { zdr: true }` so OpenRouter only
routes it to zero-data-retention endpoints (today that means Luna's Azure routes). Zero retention
limits storage, not processing: the request still reaches the provider.

## Try it

1. Open `bench/workbooks/orders.xlsx` (synthetic data) and select `A1:F501`.
2. Open NymForm, paste your key in **Setup**, choose **Continue**.
3. **Columns** shows `Using Orders, A1:F501 · 500 rows`, with Customer and Email marked private.
4. In **Ask**, keep **Structure only**, type `Total Amount per Region, one row per region` and choose
   **Preview**.
5. **What gets sent** shows the exact request and "Checked: no private values found". No names, no
   emails, no amounts. Choose **Send**.
6. **Result** shows the formula and "Formula check passed". Choose **Insert formula**.

A longer script for walking a tester through it, including what a blocked request looks like, is in
[docs/demo.md](docs/demo.md).

## Benchmark

`bench/` holds four synthetic workbooks, 20 tasks with expected answers, a bench build that runs each
task with raw data, with sample rows in stand-in form, and with structure only, and an independent
leak check.

**Results:** not yet run against a live model. The pipeline and the leak check have been run end to
end against a scripted model (zero leaks); the accuracy table needs Excel and a key.

To reproduce, regenerate the workbooks and tasks (seed 42, byte-identical), then start the bench
build on https://localhost:41951:

```sh
npm run bench:generate
npm run dev-server:bench
```

Sideload `manifest.bench.xml` (or run `npm run start:bench`), open a bench workbook, choose
**Nymform bench** on the **Home** tab, paste a key, and choose **Run benchmark**. Then check the
exported requests with the independent leak check:

```sh
npm run leakcheck -- exported-requests.json
```

## Development

| Command | What it does |
| --- | --- |
| `npm test` | Unit tests (Vitest) |
| `npm run lint && npm run typecheck` | ESLint and TypeScript |
| `npm run build && npm run check:release` | Release build and release guard (T-P4, CSP, branding) |
| `npm run battery` | Regression battery: ordinary questions on the bench workbooks, with the table |
| `npm run test:e2e` | End-to-end tests against a fake Excel host (build first). `NYMFORM_E2E_PORT` picks another port; `NYMFORM_E2E_REUSE=1` reuses a server already running |
| `NYMFORM_E2E_LIVE=1 npm run test:e2e` | Also calls the real endpoint (add `OPENROUTER_API_KEY` for the full flow, `NYMFORM_E2E_MODEL` for another model) |
| `npm run validate` | Microsoft's manifest validator |
| `npm run hooks` | Uses the gitleaks pre-commit hook |

Build settings: `NYMFORM_ENDPOINT` (default `https://openrouter.ai/api/v1`) and
`NYMFORM_DEFAULT_MODEL` (default `openai/gpt-6-luna`). The spec is [docs/spec.md](docs/spec.md);
departures from it are in [docs/decisions.md](docs/decisions.md); what has been checked in Excel is
in [docs/spike.md](docs/spike.md).

## License

AGPL-3.0-only: the GNU Affero General Public License, version 3 only. See [LICENSE](LICENSE) and
[NOTICE](NOTICE).

Created and maintained by Jorge Garcia.
