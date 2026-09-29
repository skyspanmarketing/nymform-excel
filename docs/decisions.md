# Decisions

Where the build departs from, or adds detail to, `docs/spec.md`. Newest last. Each entry says what
was decided and why.

## 2026-09-24 — Setup

- **Scaffold.** `yo office` (generator-office 3.0.4) refuses to run as root in the build container,
  so the project was assembled by hand to match the React task pane template: webpack 5 with
  ts-loader, HtmlWebpackPlugin, CopyWebpackPlugin, `office-addin-dev-certs` for the local HTTPS
  certificate, `office-addin-debugging` to sideload on desktop and `office-addin-manifest` to validate.
  The template's `office-addin-lint` was left out; ESLint with typescript-eslint is used directly.
- **Node.** Node 20 reached end of life in April 2026. The project targets Node 22 LTS
  (`engines: >=22.12`, CI on 22). Vitest 5 requires it.
- **TypeScript 5.9.** TypeScript 7 (the native compiler) doesn't yet expose the API that ts-loader
  and typescript-eslint need.
- **Dependencies beyond §3.** Approved-list libraries plus build and test tooling only:
  `@types/*`, `typescript-eslint`, `@eslint/js`, `globals`, `tsx` (runs the bench scripts),
  `@playwright/test` 1.56 (end-to-end tests; matches the preinstalled Chromium). None ships in the
  add-in bundle.
- **Microsoft CLI usage data.** `office-addin-dev-certs`, `office-addin-debugging` and
  `office-addin-manifest` include Microsoft's opt-in usage-data module for the command line. It
  concerns the developer tools, not the add-in; the add-in itself has no telemetry (invariant 7). Opt
  out with `npx office-addin-usage-data off`.
- **Manifest version 1.0.0.0.** Microsoft's validator rejects manifest versions below 1.0. The
  package version stays `0.1.0-alpha`.
- **Two manifest IDs.** The dev and release manifests use different add-in IDs so both can be
  sideloaded at once.
- **Publisher.** `ProviderName` is "Sky Span Marketing" (Jorge's decision). Copyright stays with Jorge
  Garcia in LICENSE and NOTICE.
- **Function file.** The manifest's ribbon button needs a function file; `commands.html` only waits
  for Office.js. It carries the same Content Security Policy.
- **Dev server without a live-reload client.** webpack-dev-server's injected client needs a websocket
  that the Content Security Policy doesn't allow, so dev builds reload by hand and use the same policy
  as release builds.
- **`build-info.json`.** Every build writes its version, commit, endpoint, model and bench flag next
  to the page, so anyone can match served code to source and `npm run check:release` can confirm the
  bench flag is off (T-P4).
- **`src/core/a1.ts`.** A small shared module for A1 address arithmetic used by schema, payload, the
  gate and the adapter. Not in the §5 tree; pure TypeScript.
- **Icons** are drawn by `scripts/make-icons.mjs` (no image dependencies): the accent-colored square
  with an offset second group of lines, the "original → representation" idea from §7.13.
- **Legacy name.** No tracked file carries the legacy working name; `npm run check:release` fails if
  one does. Git history is not rewritten.
- **Security reporting.** GitHub Private Vulnerability Reporting is the initial channel (Jorge's
  decision); the release doesn't wait for a dedicated address.

## 2026-09-24 — Luna and OpenRouter

From OpenRouter's live model data on 2026-09-24 (see `docs/spike.md`):

- **Routing.** Only Luna's Azure routes are on the zero-data-retention list, so `zdr: true` sends
  every Luna request to Azure. If no Azure route is up, OpenRouter refuses the request and the pane
  says "No zero-data-retention route can serve {model} right now…". Nymform never retries without
  `zdr` (invariant 3 and §7.8).
- **Request parameters.** Luna is a reasoning model and doesn't support `temperature`; reasoning
  tokens count toward the output cap. For reasoning models (IDs starting `openai/gpt-6`,
  `openai/gpt-5`, `openai/o1`, `openai/o3`, `openai/o4`) the body sends no `temperature`,
  `reasoning: { effort: "low", exclude: true }` and `max_tokens: 4000`. Other models get
  `temperature: 0` and `max_tokens: 800`. Both send `response_format: { type: "json_object" }`.
  `provider.require_parameters` is not set, because the Azure routes list `max_completion_tokens`
  rather than `max_tokens` and requiring parameters could leave no route; OpenRouter maps
  `max_tokens` for the upstream provider.
- **Benchmark temperature.** Luna runs at its default sampling; the results say so (§11).

## 2026-09-24 — Build decisions (Jorge)

- **README line.** "Protected values stay in Excel. Nymform replaces private data before AI analysis
  and shows you what will be sent." The §1 claim is unchanged and still literally true.
- **Cut order.** Evaluation metadata and its export go first if anything has to be cut; the Worked /
  Partly / Wrong rating stays if cheap. Nothing has been cut.
- **G-PARSE** stays: a formula the tokenizer can't read is rejected.
- **Structural reason.** `AuditResult.hits` keeps the six text-scan detection kinds from the spec.
  Structural-check failures are reported separately in `AuditResult.structural` (column letters only)
  and as the `structural` reason code in evaluation metadata. Detection is unchanged.
- **Publisher, security reporting, branding** as recorded above.

## 2026-09-24 — Sessions, stand-ins and transforms

- A session is one selection context (the selection plus its lookup ranges). Refresh from selection,
  toggling the header-row checkbox, or adding or removing a lookup range starts a new session: a fresh
  stand-in map (collision guard re-run over every range) and empty history. Any column policy edit
  clears history and keeps the map. A preview can be sent only in the session and version it was
  built for.
- Tokens are keyed by the cell's exact display text, per-kind counters, at least three digits
  (PERSON_1000 after PERSON_999). A value keeps the kind of its first allocation. The collision guard
  checks display text, string values and formulas; token-shaped text already in the sheet is reserved
  so a new token never repeats it.
- Token kinds, in order: EMAIL (header or mostly email values), ID (id/ssn/emplid headers), PERSON
  (name headers, or role words such as customer, client, instructor), ID (phone, number, code, zip
  headers, or ID-shaped values), TEXT.
- A private column is never sent as is: `as_is` on a private column is treated as `stand_in`.
  Defaults: text/mixed/boolean → stand-in, number → range, date → month, free text → leave out.
- `range`: bins of width 10^floor(log10|x|), at least 10, so |x| < 10 renders "0-9"; 0 renders "0";
  negatives mirror ("-89999--80000"). Non-numeric cells in a range column get a token. `month`: Excel
  date serials (with Excel's 1900 leap-year quirk) or unambiguous date text → "YYYY-MM"; ambiguous
  text such as 03/04/2026 gets a token.
- Non-private cells in sample rows: numbers and booleans raw, dates and errors as display text.
- Question substitution matches private values (display and raw forms) longest first, whole words,
  case-insensitive from 4 characters and case-sensitive below; the original header of an aliased
  column is replaced by its alias. Short private numbers are replaced too ("top 5" can become a
  stand-in when 5 is a private value); this is the safe reading of §7.4.
- Row cap: first N data rows, 1–200, default 50. History: at most 6 messages (3 exchanges).

## 2026-09-24 — Suggestions and schema

- Header keywords are matched as whole words after splitting on non-alphanumerics and camelCase, with
  run-together forms (firstname, emailaddress, dateofbirth, ...). Beyond the spec's list, person
  headers are private by default: customer, client, patient, instructor, teacher, contact, member,
  applicant, person, tenant, donor, guardian, parent, recipient.
- Value patterns on up to 200 non-empty cells: any email address; phone shapes with 10–15 digits in
  half the cells; 9-digit IDs in half the cells; ID-shaped unique values (80% ID-shaped, distinct
  ratio above 0.9). Free text (average words above 6) → private, left out of rows. Unsure (text or
  mixed, distinct ratio above 0.5, no rule matched) → private.
- Derived columns: a column whose formulas read a private column (A1 references, whole columns,
  structured references by header) is suggested private, iterated to a fixpoint.
- Schema: error and rich-value cells count as non-empty for the 80% rule; distinct counts use raw
  values; `avg_words` and `max_len` only for text columns. `header_row` is null when there is no
  header row.

## 2026-09-24 — Request path

- Body key order is fixed (`model`, `messages`, parameters, `provider`); `bytes` is the UTF-8 length.
- The correction retry omits the assistant turn when the bad reply was empty.
- The provider refuses anything that isn't the exact object the auditor issued (a WeakSet in
  auditor.ts; `isAudited`). It sends with `credentials: "omit"` and `redirect: "error"`, trims the key,
  refuses keys with non-printable characters, and replaces the key (and anything shaped like
  `Bearer …`/`sk-…`) with `[key]` in every message.
- Error mapping beyond §7.8: timeout and 408 → "took too long"; caller abort → "Request cancelled.";
  a 4xx whose message is about data policy, zero data retention or no endpoints → the no-ZDR-route
  message; other 4xx → "refused ({status})" with at most 160 characters of the provider's message;
  a 200 carrying an `error` object is an error.
- `providerName` comes from the `X-Provider-Name` header (OpenRouter exposes it to browsers), falling
  back to the response body.
- Replies: extra keys are dropped; missing formula/placement/assumptions default to null/null/[];
  an invalid placement cell becomes null (the default placement is used); a reply cut off at the
  token cap gets its own message; `content_filter` and refusals get plain messages.
- Blocked requests are logged without their body, because the body holds the value that blocked it.

## 2026-09-24 — Auditor

- Preconditions: canonical JSON (`JSON.stringify(JSON.parse(body)) === body`, also for each nymform
  user turn, so `\u` escapes can't hide a value), a fixed set of top-level keys with exact parameter
  shapes, exactly `{role, content}` messages, the system prompt first and verbatim, every user turn
  either a description or the correction prompt, and the model equal to the configured one.
- Descriptions are validated strictly: unknown keys block at every level, so a value can't hide under
  an unexpected key. Columns, private flags and treatments come from the auditor's own context, not
  from the payload.
- Range renderings are checked as real bins, stricter than the pattern alone.
- The text scan's values are the context's private values plus the auditor's own read of the data
  rows. Hits name the column letter (or `Sheet!Letter` for lookup ranges) and never the value; error
  messages are fixed sentences.
- Fixed text is masked (spec §7.7, SECURITY.md): the system prompt, the correction prompt, known key
  names, literals, the configured model ID, fixed enum values, request parameters, key names and
  `kind` values of earlier model replies, session stand-in tokens anywhere, and private row cells the
  structural check accepted (so a round value such as 90000 doesn't block inside "90000-99999").
- indexOf is used unless there are more than 5,000 variants and the body is longer than 16 KB, when Aho-Corasick is faster; both are tested to give the same hits.

## 2026-09-24 — Formula gate

- The gate parses into a tree after tokenizing, so it knows each function's arguments, LET and
  LAMBDA scopes, and the operands of `:`. Nesting deeper than 100 levels is G-PARSE.
- Function names are matched by exact ASCII: only a–z are upper-cased, so look-alike letters never
  become Latin. `_xlfn.`, `_xlws.` and `_xlpm.` are stripped repeatedly.
- LET and LAMBDA names only count inside their own scope. A local may shadow an allowed function
  (`LET(n, …)` is common and whichever Excel calls is allowed or gated) but never an excluded one.
- An allowed function passed by name, as `GROUPBY(C2:C501, F2:F501, SUM)` is normally written, is
  allowed unless a workbook or sheet defined name shadows it. The pane reads defined names from
  Excel; if it can't, every allowed name counts as shadowed (fail closed). An excluded function
  passed by name is G-FUNC.
- `:` with computed operands is checked against the box every side could reach; a reference is
  allowed when it lies in the union of allowed ranges on its sheet. Whole rows are rejected; whole
  columns are allowed only inside allowed columns.
- Structured references are read strictly; a table in `tableNames` without a known location fails
  G-REF; a bare `[…]` fails G-EXT. Functions called through a sheet prefix fail G-FUNC.
- Non-breaking spaces, zero-width characters and full-width punctuation are G-PARSE; array constants
  hold only numbers, text, booleans and errors.
- Details can name functions, references, sheets and tables, cut to 60 characters, never string
  contents.
- Restoration works on the raw text between quotes, so everything outside replaced tokens stays
  byte-identical; unknown tokens are reported once each.

## 2026-09-24 — Office host and pane

- `readSelection` checks the cell count before loading any content. Whole-column or whole-row
  selections are trimmed to the sheet's used range; the 20,000-cell limit applies to what's left.
- Office errors become plain messages carrying Office's error code only, never Office's own text
  (which could quote workbook content).
- `ready()` rejects when Office.js isn't loaded, the host isn't Excel, onReady takes more than 15
  seconds, or ExcelApi 1.9 isn't supported. Spill readback uses ExcelApi 1.12 when available.
- `writeScratch` refuses to run outside bench builds (invariant 9 at runtime too).
- Insert: re-runs the gate against the current context; refuses a cell inside any selected range or
  on another sheet; asks before replacing cells that hold values (value types only; a failed check
  asks rather than overwrite), because Office.js writes can't be undone with Ctrl+Z.
- Default placement: the first empty column right of the selection (value types only) at the first
  data row; the reply's own cell is used when it is valid and outside the selected ranges.
- Copy is allowed only when both gate passes succeed, like Insert.
- The pane opens on Setup; tabs are Setup, Columns, Ask and Log, with What gets sent and Result as
  steps inside Ask. Accent color indigo #4B3FD1 (dark theme #9D94FF); the layered-line pattern marks
  private columns and stand-in chips.
- Exports offer a download and always a "show as text" fallback, because a failed download in an
  Office webview can't be detected.

## 2026-09-24 — Benchmark

- Seven categories (the design doc isn't in the repository): conditional_totals, counting, lookup,
  dates, text, ranking (3 tasks each) and row_anomalies (2). Orders has 7 tasks, roster 5, sections
  4, appointments 4.
- Data planted so every expected answer is unique (the generator fails otherwise), including one
  salary of exactly 90,000 to keep the round-number case under test. Generated names avoid 4+ letter
  words that appear elsewhere in the workbook (a surname "West" next to a West region would block by
  design).
- Workbook metadata and zip timestamps are fixed, so regenerating gives byte-identical files.
- Compare: numbers within tolerance, text trimmed and case-insensitive, tables order-insensitive with
  an optional header and total row ignored.
- The headline counts a task as matched when its best production condition's pass rate is at least
  its raw pass rate, and not when every condition scored 0%.
- Raw runs have no correction retry (the raw sender builds only the first request); a raw run whose
  reply can't be read is recorded as a parse failure.
- Runs are interleaved (run 1 of every task and condition, then run 2) so a bad provider period
  affects all conditions alike.

## 2026-09-24 — After the red-team review

An adversarial review (five attack lenses, each finding independently reproduced) found no way to
insert a network-capable formula and no path around the auditor for well-formed input, but it did
find the gaps below. All are fixed with regression tests; the limits that remain are in SECURITY.md.

- **Header row that is really data.** With the header checkbox on by default, a selection without its
  header row sent row 1 verbatim as the headers. Now: when row 1 looks like data (a number, date or
  boolean, or text shaped like an email, phone, SSN or ID), the checkbox defaults to off with a
  notice; a private column whose header cell looks like data gets the alias "Column {letter}"; the
  auditor adds such header cells to that column's values, so a request that would send one blocks.
  With no header row, every column starts private. A first row of ordinary words that is really data
  can't be detected; the Columns screen shows every header that will be sent.
- **Other spellings.** The scan and question substitution now fold Greek final sigma and curly
  apostrophes, add NFC and NFD forms, match CJK names typed without their space (Han, Kana and Hangul
  count as word boundaries), match 7+ digit numbers across punctuation (a digits projection of the
  body, mapped back so the fixed-text mask still applies) and amounts of 1,000 or more with thousands
  separators. These report as the existing detection kinds (`normalized`, `digits`); the kinds are
  unchanged.
- **Fill down.** `checkFill(formula, ctx, rowShift)` sweeps every relative row reference down the
  fill and requires the swept area to stay inside the allowed ranges; Insert runs it. Fill down
  starts only at the first data row, and no cell of the filled span may be inside a selected range.
- **Spill references** (`A2#`) are refused everywhere: their extent isn't known to the gate.
- **Local names** may carry only the `_xlpm.` prefix.
- **Overwrite confirmation** now covers only the exact span it named; a host that can't check whether
  cells are empty always asks.
- **A stand-in whose `NYMFORM_` prefix the model dropped** is flagged and left as written.
- **Hardening:** `form-action 'none'` in the CSP; spell check, writing aids and translation off for
  fields that can hold private text; the API key is never written to a DOM attribute; provider
  messages drop bidi and zero-width characters; `isAudited` uses functions captured at load; Pages
  deploys only from tags; G-LEN no longer reports a length.
- **Accepted, documented:** whole-column references for the chosen columns (spec T-G18); restored
  values are text (numeric IDs, wildcards); Excel's 255-character limit on text constants; each lookup
  range has its own private columns; unaccented Greek typed for an accented stored value.

## 2026-09-24 — Second and third review rounds

Two more find-fix-verify rounds on the auditor, substitution and header detection. The first fixes
caught more spellings but made requests block for no reason; these rounds balanced the two, keeping
the rule that anything the auditor doesn't recognize as fixed text is scanned.

- **Where each variant counts.** Every variant has a match rule. Whole values and 7+ digit runs count
  anywhere. Grouped number forms (`5,390`) count only inside JSON string content, with no digit or
  decimal mark continuing them, so `5,390` doesn't match inside `15,390.5` or across a JSON value
  separator (`5,390` in `[5,390]`). A 4 to 6 digit run of a value that isn't a number or a date (the
  `0199` of a phone) counts only in free text: the question, notes, headers, sheet and table names,
  allowed ranges and earlier replies' text. In row cells and JSON numbers it matched counts and amounts
  so often that half of all sample-row requests on a sheet with a phone column blocked. Year-like runs
  (1900 to 2100) are skipped. Recorded as a stated limit.
- **Computed numbers are fixed text once validated.** Row numbers must match the range address and
  the header row; counts (blanks, distinct, rows) must be integers within the row count. After that
  they are masked like key names. Previously a private 3-digit or 4-digit value blocked every
  request whose selection had that many rows.
- **Amounts.** The scan and question substitution read a number the way the cell shows it and the
  way it is stored: currency symbols and codes of up to 4 letters (`$`, `kr`, `CHF`, `Rs.`), `K`/`M`
  suffixes, text amounts from CSV imports, zero decimals, `,` or `.` as the decimal mark, lakh
  grouping. The stored value always counts, whatever the format shows. Float noise
  (`1234.2000000000003`) is read to two decimals.
- **Scripts and spellings.** Turkish dotted I, accents folded on letters, `ß` as `ss`, full-width
  forms, Arabic-Indic, Persian and Devanagari digits, look-alike apostrophes (`ʼ`, `＇`, `´`), and `_`
  as a word break. They report under the existing kinds (`normalized`, `digits`, `word`); the kinds are
  unchanged (decision 7).
- **Phones.** A stored number is also scanned for without its trunk 0 (`+44 20 7946 0958` for
  `020 7946 0958`, 9+ digits) and without a country code (`(831) 555-0199` for `+18315550199`).
- **Header row.** Row 1 counts as data for any number, date, amount, percent or time, and for
  ID-shaped text unless it has a run of 2+ letters over a column that isn't ID-shaped. A bare number
  or year header is not added to the scan (it would block every request that mentions the year); an
  identifier header is. The header choice carries over only for the same range; the automatic
  `Column X` alias is dropped when the column is un-marked.
- **Accepted, documented:** false blocks when a private value also appears in an unmarked column;
  workbooks on the 1904 date system (date forms from the stored value are off by four years; the
  displayed dates are still scanned).

## 2026-09-24 — Fourth to tenth review rounds

Seven more find-fix-verify rounds. In each, an agent that hadn't written the fix verified it against
the committed code and the previous round, then hunted for new leaks. Four fixes opened a new leak,
each closed in the next round. That history is why these rules look the way they do. The work was
committed only once a round's verification found nothing that leaks here but not on the committed
code, apart from the one accepted trade-off below.

- **Dates are parsed, not listed.** Listing 13 written forms for every reading of every private date
  cost up to 3.7 s for 20,000 text dates with 2-digit years. Now the auditor keeps each private
  column's dates as ISO dates, and `findDates` reads every date written in the body and compares.
  That covers numeric dates in any order (mixed separators, `_` and `\` included), month names in
  several languages, ordinals, run-together forms such as `15MAR1985`, CJK forms, zones after a
  time, and invisible marks between parts. A 20,000-cell `prepare()` stays under 1 s in every
  scenario measured, down from up to 3.7 s. Dates are read from the stored value and the display
  text, which covers 1904 workbooks when the cells show a full date.
- **Overlapping dates.** In "born March 15, 1985 - September 2, 1985", replacing the longest date
  reading ("1985 - September 2") left pieces of both real dates in the question (round 5). Date
  phrases are now resolved leftmost first, and a guard replaces a whole overlapping stretch if any
  letter or digit of a private date would be left outside a stand-in.
- **8-digit values.** 8-digit numbers or text are read as dates (every valid order) only when at least
  half of the column's 8-digit values, placeholders such as `00000000` aside, are valid dates, and
  those dates count only in free text. So invoice numbers `20240001…` don't turn every date typed in
  a question into an invoice stand-in. Round 8 also stopped their digits counting where the body spelled
  them as a date. That let "invoice 2024-01-15" through, so round 10 reverted it: the digits of every
  8-digit private value count wherever they appear.
- **Round hundreds are scanned (reverted in round 5).** Round 4 skipped 4-digit runs that are
  multiples of 100, so that "top 1000 rows" wouldn't block on a sheet of 20,000 private phone
  numbers. That let "SSN ends in 6700" through. Only years are skipped; the false block is stated.
- **Numeric boundaries.** Inside a cell stored as a number (a JSON number in the rows), a private
  number under 7 digits doesn't match where another digit is directly next to it (ZIP `93496` inside
  `934961`), and not as the decimals of a number with one decimal point (a longitude `-120.93496`).
  This is the one accepted trade-off against the committed code: an unmarked number cell whose
  digits or decimals contain a private number is sent. Without it a 10,000-row contacts sheet with a
  private ZIP and unmarked coordinates blocked every sample-rows request (30 of 30); with it, 0.
  Everywhere else (the question, notes, headers, earlier replies, and text cells), a private number
  counts even inside a longer number: `934961234` (a ZIP+4 typed without its hyphen) or
  `Transfer to 48392001` in an open Memo column. Earlier rounds narrowed this too far three times:
  commas (round 5, `93496,94103`), decimals in lists (round 7, `1.93496`), and text cells (rounds
  5–10, the Memo case).
- **Grouped amounts.** A grouped form uses one separator throughout, and its decimal mark differs from
  that separator (`5,390,000` isn't `5,390`). In free text a space-grouped form is rejected only
  when a digit touches it, so a list `40 000 83 500` matches both (round 8 had missed it).
- **Invisible characters.** Default-ignorable characters, non-whitespace controls, form feed and
  vertical tab are dropped for matching (`Wata\u200Bnabe`), and the spot counts as a boundary for
  every rule (`Mr\u200BWatanabe`, `93496\u200B94103`). Round 8 dropped them without the boundary and
  leaked 840 of 840 such checks.
- **Codes.** A value of 6+ letters and digits that mixes both also counts in a letters-and-digits
  projection of the body (`AB-1234CD`, `AB 1234CD` for `AB-1234-CD`). A match must lie entirely in
  unmarked text, so a stand-in followed by a number can't form a code. The colon of a cell range isn't
  a separator, so `Sum of A10:B20` isn't a private `A10-B20`. Substitution replaces a code only as a
  whole word.
- **Range addresses** that equal a range the auditor derives itself are fixed text (the cell part,
  never the sheet name), so a private amount of 1000 no longer blocks on `A1:J1000`.
- **Header row.** IBANs, UK National Insurance numbers, postcodes (UK, Canadian, ZIP+4, Irish, Dutch,
  `123 45`) and street addresses with a street-type word count as data. Text that starts with a house
  number but has no street-type word, and generic two-part codes, count as data only when the column
  below agrees, so "10 Largest Orders", "401k Match Amount" and "4 Wheel Drive" stay headers.
- **More spellings, same kinds:** names without an apostrophe or hyphen before a capital (`OBrien`),
  NFKC (ligatures), ä ö ü as ae oe ue, an email's local part when it has a digit or `.` `_` `-`, the
  digits of each part of a value between letters (phone extensions, IBAN account numbers), exact
  `k`/`M` short forms, `_` grouping, and the last 4 digits of a number shown as `#####`.
- **False blocks in sample-rows mode** on the synthetic sheets, committed code → now: contacts 42 → 0 of
  100 (and 30 → 0 of 30 at 10,000 rows), sales 90 → 9, payroll 26 → 1, orders and roster 0 → 0.
  Structure-only mode stays at 0 of 100 on every sheet. The remaining blocks are real collisions
  with unmarked columns.

## 2026-09-25 — Dev server port

The dev server moved from port 3000 to 41951, because 3000 is the default of many local tools and
was already taken on Jorge's machine (`EADDRINUSE`). 41951 is outside the ports common tools use and
below macOS's temporary-port range (49152 and up). It is set once, in `package.json`
(`config.dev_server_port`), which webpack and `office-addin-debugging` read. The dev and bench
manifests carry the same port, and `tests/manifests.test.ts` fails if they drift apart. The release
manifest is unaffected: it points at GitHub Pages.

## 2026-09-25 — Parts of private values: pick instead of block

Typing part of a private value ("Orders from Felix" when the sheet has Felix Bianchi, Felix Kim and
Rosa Felix) used to block the request, because substitution replaces whole values only. Jorge asked
for the question to be mapped instead, with a choice when there are several matches and the cells
shown. Now:

- `findPartialMatches` (src/core/partial.ts) finds words of the question that are parts of private
  values, as the auditor would: words of multi-word values, email local parts, 4 to 6 digit groups
  (years skipped), folded for case and accents.
- **One candidate** naming a person, email or ID is read as its stand-in without asking; the preview
  says "Read “Kenji” as Kenji Watanabe" with **Keep as typed** to undo. A free-text candidate (a
  city, a note) always asks, because mapping a common word to it would change the question's meaning.
- **Several candidates** are listed with their stand-in, column and number of cells (at most 30,
  the rest counted). The user ticks one or several, sent as `PERSON_014 or PERSON_022`, or keeps the
  part as typed, and the auditor then blocks it.
- **Show in sheet** *selects* the candidate's cells in Excel rather than colouring them. Colouring
  would be a workbook write, which breaks invariant 9 (writes only through Insert); Excel can't undo
  an add-in's writes, and colours left behind by a closed pane would stay. Selecting writes nothing.
  Several cells are selected together from ExcelApi 1.18 (RangeAreas.select), else the first one;
  the user's range is selected again once they choose.
- The candidates (real values and their cells) stay in the pane: they aren't logged, exported or
  sent. The choice only puts stand-ins into the question, and the request is audited as before, so
  a wrong or missing choice still blocks.

## 2026-09-25 — Jorge's column B block: a header typed again under Header sent

Jorge's first live session blocked 7 requests, 4 of them structure-only, with "a value from column B
is in the request" (hits `[B exact, B word]`). The exported Log and evaluation report were enough to
reproduce it with the real Flow and auditor on the playground workbook (Orders!B1:G1401). The
reproduction regenerates his one logged sent body byte for byte, and three independent agents
checked it. The date values in B played no part: date matches report kind `normalized`.

He had typed the header again under Header sent in another case or spacing ("Order Date"). The
auditor decided a header was hidden with a raw `alias !== header`, while the payload sends the
trimmed alias and the scan folds case. So "Order date" became a hidden header, scanned whole and
word by word. That applies to columns that aren't private too, and a typed alias survives unticking.
The scan then found it in the header Nymform itself sends for B, and every question blocked, in both
modes, with B ticked or not.

Fix, and nothing wider: `aliasRevealsHeader` (transform.ts) says an alias still shows its header when
it is the same text in another case, spacing, width or accents, or holds the header's words whole
("Order date (month)"). The auditor doesn't treat such a header as hidden, and the question's header
rewrite skips it. Only whole words count: "Tracer" keeps "Race" hidden. Plain containment would let
"race" through in a note. A header renamed to something else is scanned as before, words included.
The remaining false blocks from that are listed in SECURITY.md, and the header, sheet and table-name
scans are unchanged.

The same run found that the picker read "price" in "Average Unit price per Region" as the customer
Jordan Price and sent "Average Unit PERSON_001 per Region". The picker now skips words inside a
column's name as typed in the question (its header, or the alias of a hidden one). The auditor still
blocks a header word that is a word of a private value: "Unit price" with a private customer Jordan
Price blocks every request. The way out is renaming that header under Header sent. Jorge asked for no
header exemption unless the log showed one was the cause, and it didn't.

Pinned in `tests/falseBlock.test.ts`, which fails 12 of its cases on the old comparison.

### Independent audit of that fix (same day)

The fix went onto `Live` before an independent audit. Jorge asked whether one had been done, and
it hadn't, so it was run after the fact: three auditors (a 72,000-run differential fuzz against the
previous commit, Unicode and word-boundary attacks, and the picker plus a mutation run of the tests),
with each finding reproduced or refuted by a separate agent. Confirmed and fixed here:

- **Leak.** `aliasRevealsHeader` dropped every nonspacing mark, not only accents. A rename that
  differs by a Thai tone mark, an Indic vowel sign or a dakuten (ビザ番号 "visa number" renamed
  ピザ番号 "pizza number") counted as showing the header, so the hidden header stopped being
  scanned and the question stopped being rewritten. It now folds only what the auditor's scan
  folds: U+0300–036F accents, case, width, look-alikes and characters that can't be seen.
- **Leak.** Its whole-word check read a mark as a boundary, so the Hindi equivalent of "Races"
  showed "Race". Marks now count as part of their word, as in the auditor.
- **Regression.** A header of about 11,000 characters or more (a first-row cell can hold 32,767) made
  the picker's column-name pattern overflow the regex engine's stack, and every Preview failed. Names
  longer than the question are skipped, and a pattern the engine can't build masks nothing.
- **Regression.** Each column's pattern used Unicode lookarounds that cost about a millisecond to
  compile, again after every garbage collection: about 1 s on the first Preview at 1,000 columns.
  The whole-word check now runs outside the pattern.
- **Regression (contrived).** A header made only of marks, typed again unchanged, blocked. The same
  text typed again now always counts as showing the header.
- **Test gaps.** Eleven of 27 mutants of the new code survived the suite, and the "Tracer/Race" test
  passed because of the alias itself, not the note. The tests now cover whole words after letters,
  astral letters and marks, other scripts, regex characters in names, long names, and a genuine
  rename with a note. Each new test fails on the audited commit.

Found by the same audit, and older than this fix: a private name written in styled Unicode letters
(math bold 𝐋𝐞𝐧𝐚, circled Ⓕⓔⓛⓘⓧ, or the ﬃ ligature in the question) is neither replaced nor
blocked. That needs a change to the auditor's fold, done next with its own audit.

### Re-audit of the audit fix

The same harnesses were rerun on the fix (72,000-run fuzz, probes, perf, long headers, a
38-mutant run, and a fresh read of the diff). Every confirmed finding is fixed, with no regression
against the code before the alias fix and no new leak class. First-Preview time at 1,000 columns
is back to about 190 ms, and no header length makes Preview fail. Closed after the re-audit:

- **Tests.** One surviving mutant mattered: if the step over a surrogate pair after a rejected match
  were removed, Chrome would loop forever on a column name that starts outside the basic plane
  (𠮷野家) typed right after a letter. That step is now tested, and the loop also stops if the
  engine returns a match it already turned down. Also tested now: restarting inside a rejected match ("xBora Bora Bora"), spacing
  inside a name, and fold details (a backtick in an ASCII header, a trailing U+0085).
- **Lone surrogates.** The whole-word check stepped back two units after any low surrogate. It now
  does that only for a real pair, as the scan reads it.
- **Rule written down.** `aliasRevealsHeader` folds what the auditor's scan folds, look-alikes
  included. An alias that differs from the header only in a stroke letter, a dotless i, digits of
  another script, a dash or apostrophe look-alike, or width ("Lodz" for "Łódź", "Q1–Q2" for "Q1-Q2")
  shows the header: the scan itself finds the header in that alias. The auditors called these
  borderline, not leaks, and this is the rule now.

Pre-existing auditor gaps the audits found, to be fixed next with their own audit:

- private names in styled letters (𝐋𝐞𝐧𝐚, Ⓕⓔⓛⓘⓧ, the ﬃ ligature in the question);
- no case folding for astral scripts (Adlam, Osage, Deseret), so a private value in the other case
  in a note passes;
- a kept part whose pieces are all shorter than 4 characters ("O'Bob") passes;
- a very long hidden header typed in the question makes Preview fail. That fails closed.
## 2026-09-25 — Publisher, build label, CP2 hooks

- **Publisher "SkySpan"** (Jorge's decision) in all three manifests (`ProviderName`), the README
  credit and the pane footer. Copyright stays with Jorge. If a company is formed, its legal name
  replaces this; that doesn't block the alpha. No skyspanmarketing.com address is given as a
  contact. security@nymform.com is added to SECURITY.md once its forwarding is set up and tested;
  until then GitHub private vulnerability reporting is the channel. `tests/manifests.test.ts` checks
  both.
- **Build label.** A bundle built without git (from a downloaded ZIP) showed "build unknown". The
  footer and the evaluation records now say `0.1.0-alpha · local build`, or `0.1.0-alpha · build
  abc1234`.
- **CP2 live check.** The live primary-flow test takes `NYMFORM_E2E_MODEL`. A reply that asks a
  question now fails with the model's question instead of timing out: "Formula check: none" in
  Jorge's first session was a clarify reply, not a gate result.
- **Parallel e2e runs.** `NYMFORM_E2E_PORT` moves the test server off 4173, so two checkouts can run
  the suite at once. A server already on the port fails the run instead of being reused (it may be
  another checkout's build); `NYMFORM_E2E_REUSE=1` reuses it on purpose.

## 2026-09-25 — The check is frozen after this round

After this round (the alias fix and its audits, pane diagnostics, the word-boundary rule and the
confirmed ordinary word), the check is frozen: `src/core/auditor.ts` and the matchers it imports
from `src/core/transform.ts`. It changes only for a reproducible failure from the benchmark or real
use, pinned with a test and independently audited before it reaches `Live`. Round 12 showed why:
the alias fix went out before its audit, and the audit found two leaks in it. `CLAUDE.md` says the
same for coding agents.

## 2026-09-25 — Selection resolution

What Nymform thinks it is looking at. In Excel for the web, clicking one cell (E4) made the data one
cell (header "East", no rows), and clicking a row header (4:4, 16,384 cells, under the limit) was
read untrimmed: a 2.7 MB request with the address "4:4", which the flow can't parse, so the check
blocked it. `readSelection` now resolves the selection first (spec §7.1).

- **Order.** Whole rows or columns of any size are trimmed to the filled cells inside them. A
  pointer (one cell, or one row after trimming) is anchored on its cell, or on the row's first filled
  cell, and stands for the table holding the anchor (else the only table where the row holds
  something) without the totals row, else the sheet's AutoFilter range when it holds the anchor, else
  Excel's current region around the anchor cut to its filled cells. A table the row crosses where it
  is blank doesn't count, wherever it lies along the row. Anything else is
  used as selected. Fewer than two rows, nothing filled, or an address that isn't a plain A1 range
  gives one message: "Select your data, or one cell in it, and choose Refresh from selection."
- **Trimming uses the range's own used range.** The spec's first draft intersected with the sheet's
  used range. `Range.getUsedRangeOrNullObject(true)` on the selection, and on the region, asks Excel
  in the same round trip. Excel documents it only as "the used range of the given range": whether
  that is the box around the filled cells inside the range or the range cut to the sheet's used
  range, it lies inside the intersection, so it is never looser. On the first reading it is tighter
  where it matters: a row trimmed to its own filled cells starts at a filled cell, so a title in A1
  over data starting in column B doesn't make the region start in an empty column A; and the empty
  row or column Excel adds to the region of an empty cell next to the data is cut off even when
  something else on the sheet uses that column. The fake host can take either reading
  (`fake.usedRange`). On the second, a clicked row can start on an empty cell, so the adapter
  anchors the region on the row's first filled cell, found from the row's value types (empty or
  not, never the values), and looks for tables only where the row holds anything, so a table the
  row crosses where it is blank isn't taken for the data. The clicked-row case is tested under both
  readings, with a table on the row. What the second reading still costs, only if Excel turns out to
  use it:
  - an empty row or column next to the data can come with an empty clicked cell's, or an empty
    selected row's, region;
  - whole columns can bring empty rows down to the sheet's last used row, and rows 3:5 empty
    columns before the data;
  - a filter range is cut to the sheet's used range instead of its own filled cells, so it can take
    in blank rows. A filter just under the limit plus a far-off note could then be refused as too
    large.

  These widen the read to blank cells, and at worst refuse a range the first reading would take. A 30-second check in Excel for the web
  settles it: data in C1:F50, a value in A70, click the row 4 header, and see whether the pane says
  "Using …, C1:F50".
- **The filter range is cut to its filled cells.** Excel reports the range the filter was set on:
  whole columns (`A:F`) when the user selected columns before turning on Filter, or rows far past
  the data. Used as reported, a click in the data was refused as 6 million cells, or read hundreds
  of blank rows. Its own used range is asked for in one more round trip, only when the filter holds
  the clicked cell (shared with the table step's round trip when there is a table), and the checks
  below run on what's left.
- **A filter on the header row alone** is skipped (it says nothing about where the data ends), and
  the region is used.
- **A table** found while growing is passed to the read, which then asks Excel for no tables: a click
  in a table costs four round trips, not five. Its column names are its header cells, so the read
  loads them with the cell contents, after the size check (an audit caught them loaded before it).
- **Known limits, left as they are.** A trimmed row is anchored on its first filled cell, so a note
  in A4 apart from the data makes row 4 the note's own block and gives the resolve message; the
  user clicks a cell in the data. A table with its header row turned off grows to its body, whose
  first row is then the header unless it looks like data (the pane's usual header check and
  checkbox apply); the block around the cell would be the same body when the row above is blank,
  and would take in a totals row. A row filled from A to XFD, read exactly, is called "4:4" by
  Excel and gives the resolve message.
- **Size.** Only positions and sizes are loaded before the check, plus the value types (empty or not)
  of one clicked row, to find where it holds anything; the count comes from the row and
  column counts (Office reports `cellCount` as -1 past 2^31-1 cells, a whole sheet). A range grown
  from a pointer says what was large: "The data around E4 has 28,000 cells. Select fewer cells
  (limit 20,000)." The read then passes the count it has, so `readRangeData` skips its own size
  round trip.
- **Exact.** `readSelection({ exact: true })` skips the pointer rule and may return one row. "Use
  exactly the selected cells" reads the resolution's `exact` address with `readRange` on the range's
  sheet, not the current selection, which may have moved; it starts a new session like Refresh.
  A range with no data rows (one clicked cell read exactly) can't be previewed: "This range has no
  data rows. ..."; unticking the header row makes it a data row.
- **Fixed until Refresh.** The pane still reads the selection only when it opens and on Refresh
  from selection. Columns shows "Using Orders, A1:F501 · 500 rows" (en-US thousands), the line
  saying how it was found, the exact-cells button for a table, filtered range or block, and
  "Nymform keeps using this range until you choose Refresh from selection." A row header that grew
  says so with the row: "From the block of filled cells around 4:4." The whole sheet is "The whole
  sheet trimmed to the cells in use.", not Excel's "1:1048576". A table the range grew to isn't
  named again after the summary, and the checkbox reads "This range has a header row", since the
  range is often not the selection. A lookup range added from the selection resolves the same way
  and shows its line, without the button; when there is nothing to read, the pane asks for the
  lookup range and says a one-row range can be typed. Putting the selection back after Show in
  sheet selects the cells the range came from (E4), so a later Refresh resolves it the same way.
  "Use exactly the selected cells" on a sheet renamed or deleted since says so instead of asking
  the user to check an address they didn't type.
- **Where it lives.** `RangeData.resolution` is optional, so existing fakes and tests stay valid.
  It holds addresses and a table name, is shown in the pane and never goes into a request.
- **Workbook info in one run.** `readWorkbookInfo()` reads sheets, tables and names in one
  `Excel.run` of two round trips instead of three runs of five. The flow uses it when the host has
  it and falls back to the three reads, each with its own default, when it fails.
- **API sets.** `isEntireRow`, `isEntireColumn` and `getSurroundingRegion` are ExcelApi 1.7,
  `Range.getUsedRangeOrNullObject` 1.4, `Worksheet.autoFilter` and `AutoFilter.getRangeOrNullObject`
  1.9, `Table.showTotals` 1.1: all within ExcelApi 1.9, which the add-in already requires (checked
  against `@types/office-js`). The fake host implements them, with Excel's current-region rule
  (grow while any cell of the one-cell ring, corners included, is filled; clipped at the sheet's
  edges), and `workbookToFake` carries an ExcelJS sheet's AutoFilter and a table's totals row across.

## 2026-09-27 — License: AGPL-3.0-only

Jorge's decision: Nymform for Excel is relicensed from Apache-2.0 to the GNU Affero General Public
License, version 3 only (SPDX `AGPL-3.0-only`). "Only" is deliberate: the notice names version 3 and
leaves out the "or any later version" option, so a later version of the license can't be applied
without a new decision.

- `LICENSE` is the text of https://www.gnu.org/licenses/agpl-3.0.txt as published, byte for byte
  (SHA-256 `0d96a4ff68ad6d4b6f1f30f713b18d5184912ba8dd389f86aa7710db079abcb0`, checked against a
  second download).
- `NOTICE` keeps the product name and "Copyright 2026 Jorge Garcia", and carries GNU's standard
  notice with the version clause for version 3 only.
- `package.json` and the root entry of `package-lock.json` say `"license": "AGPL-3.0-only"`. The
  README's License section and spec §15 say the same.
- Copyright stays with Jorge, as before. The repository has been private and every commit is his,
  so there are no outside contributions to relicense. The dependencies bundled into the add-in are
  under MIT (92), Apache-2.0 (1) and 0BSD (1), per `package-lock.json`; all allow their use in an
  AGPL-3.0 program.

## 2026-09-28 — Release-hardening review: insert outcomes, unresolved stand-ins, fresh data, the correction

An outside review of `3813b20` found places where the pane said more, or checked less, than the
code did. Each was reproduced with a test before it was fixed. None of it touches the frozen check
(`auditor.ts` and its matchers): what is scanned, and the README claim, are unchanged.

- **What Insert reports.** `insertFormula` wrote, filled, read the calculation mode, the spill range
  and the value types in one or two syncs, and any failure threw. A readback that failed after the
  write had gone through was shown as "Couldn't insert the formula", with the formula already in the
  sheet; an Excel error in the result was recorded for the report but shown as a plain "Inserted in
  G2." Now the write has its own sync. If it fails, it throws and nothing was written. After it, the
  adapter returns `check` (`InsertCheck` in `core/types.ts`): `ok`, `error`, `manual` or `unverified`,
  plus `fillFailed` when the fill down failed after the first cell. The pane shows only `ok` as a
  success; the others are warnings that say what to look at. Nothing is written again on the user's
  behalf: a second Insert asks before replacing, as for any cells that hold something.
  `EvalRecord.excelError` is set for `ok` and `error` only, as before for the manual case; no field
  was added.
- **A safe formula isn't necessarily a correct one.** `restoreFormula` already reported token-shaped
  strings it couldn't put a value back for, and the pane showed a warning, but `canInsert` came only
  from the gate. A criterion such as `"PERSON_099"`, never issued, passed the gate and would have
  compared the data against the stand-in text. Such a formula now can't be inserted or copied
  (`FlowResult.unresolvedTokens`). The prefix rule is unchanged: with `NYMFORM_` tokens, an
  unprefixed look-alike is sheet text, not an unresolved stand-in. Unknown stand-ins only in the
  explanation stay a warning.
- **Restored values Excel may read differently.** A stand-in comes back as the cell's shown text,
  so a private number or date is text in the formula (leading zeros as shown), and `*`, `?` and `~`
  act as wildcards in criteria (SECURITY.md already said so). Blind coercion or escaping would change
  what the user asked for, so neither is done; the Result screen now marks these before Insert
  (`FlowResult.cautions`).
- **Fresh data (Jorge, 2026-09-28, Decision Desk option A).** The pane reads the range at Refresh
  only (§7.1). A name typed into a private column after that isn't known to the session: a question
  naming it went out unsubstituted and passed the check, which works from the same snapshot. Send
  and Insert now read the session's ranges again at their addresses, and the workbook's tables and
  defined names, and compare them with what the session was built from (`sameCells`,
  `sameWorkbook`). A difference stops the action ("The data changed … nothing was sent" / "nothing
  was written") and marks the session stale until Refresh from selection; a read that fails stops
  it too, and trying again works. These are the reads Refresh already makes, local only. No event
  listeners, and no reading on every click. A formula's own result isn't compared, since volatile
  functions change it on every read, nor is display text (a narrow column shows ####). Not
  noticed: rows added below a plain range (not a table), which Refresh picks up.
- **Whole columns (option A).** G-REF keeps allowing `E:E` for chosen columns (SECURITY.md). The
  Result screen now names such references and says the formula reads those columns' other rows in
  Excel, and that nothing more is sent.
- **The correction request (option A).** §7.9 sent one correction request automatically when a
  reply couldn't be read. It was audited, but never shown, while §7.13 says the request is always
  shown before sending. It is now built and audited the same way and returned as a preview
  (`SendOutcome.correction`, `Prepared.correction`), with a line saying what it is. It goes out only
  when the user chooses Send, updates the same evaluation record (`retries: 1`), and never leads to
  another correction.
- **Tests.** `failOnSync(n)` in the fake host fails a later sync, so a failure after the write can be
  tested. The flow tests' fake host serves the session's own range at its address, as Excel does.
  One adapter test read its workbook through `URL.pathname`, which failed in a folder whose name has
  a space; it uses `fileURLToPath`.

### Independent review of that change, and the first real-Excel run (same day)

An independent review of the commit, and the first run in Excel for Mac 16.113.2, found these; each
is now pinned with a test.

- **Real Excel: hidden names made every later Send look stale.** Inserting a formula with a newer
  function makes Excel add hidden defined names (`_xlfn.SUMIFS`, `_xlfn.GROUPBY`, `_xleta.SUM`),
  which the names read returns. The check compared them, so after any Insert the next Send said
  "The data changed" although no cell, table or name of the user's had changed (confirmed by saving
  the workbook and comparing A1:F501 with the source file: no difference). The fake host doesn't
  model those names. They are now left out of the comparison (`isExcelInternalName`), with `_xlnm.`
  (print areas, filters). The gate was never affected: it compares names whole, so `_xleta.SUM`
  doesn't shadow SUM.
- **A read that fell back at Refresh** (names or tables, to their fail-closed defaults) no longer
  compares against a later successful read (`WorkbookRead`).
- **Stale stays stale** through header and lookup-range changes; only a fresh read clears it.
- **The correction** is offered only while the preview it corrects still stands (columns and
  conversation unchanged while the reply was out), and sending it again after a failure doesn't
  count a second retry.
- **Sheet text shaped like a stand-in** that is in the ranges stays a warning, not a block, even when
  its prefixed form is a stand-in: the formula may mean the sheet's own value, and asking again
  couldn't help.
- **Insert** checks once more, after its reads and just before writing, that the session is the one
  the answer came from and isn't stale.

## 2026-09-28 — Dependency audit: two overrides, dev tooling only

`npm audit` reported 10 vulnerabilities (2 moderate, 8 high); `npm audit --omit=dev` reported none.
Every finding is in development tooling, none in the add-in bundle (React, Fluent UI, zod):

- `adm-zip` below 0.6.1 (crafted-ZIP memory allocation, extraction following symlinks), pulled in by
  Microsoft's `office-addin-manifest`, `office-addin-debugging` and the Microsoft 365 Agents Toolkit
  CLI they depend on. Used here to validate XML manifests and sideload; no ZIP from outside is opened.
- `uuid` below 11.1.1 (a missing bounds check when a buffer is passed to v3/v5/v6), pulled in by
  `exceljs`, which the bench generator and the tests use. `exceljs` calls only `v4`, without a buffer.

`npm audit fix --force` would downgrade `office-addin-debugging` to 6.0.4, a major step back, so it
wasn't used. Instead `package.json` carries two `overrides`: `adm-zip` `^0.6.1` everywhere, and `uuid`
`^11.1.1` under `exceljs` (uuid 11 still offers `v4` to CommonJS). After a clean install `npm audit`
reports 0. Lint, typecheck, unit tests (apart from the path-with-space test fixed on the hardening
branch), both builds, the release guard and its bench rejection, manifest validation and the
end-to-end suite pass, and `npm run bench:generate` reproduces the committed workbooks byte for byte.
Revisit the overrides when Microsoft's tooling moves to adm-zip 0.6.1 itself.

Follow-up the same day: the supported upstream releases come first. `office-addin-manifest` 3.0.1 and
`office-addin-debugging` 7.0.1 (with `office-addin-dev-settings` 4.0.1), all inside the ranges in
`package.json`, already require `adm-zip ^0.6.1`, so the direct tools no longer need the override.
It remains only where Microsoft's Agents Toolkit chain pins old versions with no newer release:
`@microsoft/teamsfx-core` 3.1.3 pins `@microsoft/kiota` 1.31.1 (`adm-zip ^0.5.17`) and its own
`office-addin-manifest` 2.1.6 and `office-addin-project` 1.0.10 (`adm-zip` 0.5.x). There the override
crosses the declared range (0.5.x to 0.6.1), so compatibility was tested, not assumed: every adm-zip
call those packages make (`new AdmZip(path?)`, `addLocalFile`, `addLocalFolder`, `addFile`,
`writeZip`, `writeZipPromise`, `getEntries`, `readFile`, `extractAllTo`) was exercised against 0.6.1:
`exportMetadataPackage` of all three `office-addin-manifest` copies was called for real, and the call
sequences of office-addin-project's backup/restore, teamsfx-core's app-package writer and readers, and
kiota's installer were replayed with the adm-zip each of them resolves (`convertProject` runs
`npm install`, so it wasn't called). `exceljs` 4.4.0 (latest) still asks for
`uuid ^8.3.0`; it calls only `v4()`, which uuid 11 keeps for CommonJS, and a workbook with an
extended conditional-format rule (the one path that calls it) was written and read back. None of these
tools runs in the add-in's bundle. The production dependency tree is unchanged.

## Independent-review follow-up, 2026-09-28

The local follow-up keeps the release candidate dependency overrides and closes the collision-mode
unresolved-token gap: an unprefixed token counts as workbook text only when the session confirms
it was observed in the workbook. A regression checks that an unknown token cannot be copied or
inserted. Existing workbook-text and dropped-prefix cases remain covered.

The limited synthetic-data alpha defers the live benchmark, with its absence stated explicitly,
as now recorded in spec §15. No paid run is authorized by this change. No accuracy or sensitive-use
claim follows from the seven-request Mac smoke run.

The archived dependency checks recorded versions; they did not assert them. Their seven checks
cover selected call sequences, not every API used by the Microsoft tooling chain (for example,
Teams tooling also uses extractAllToAsync). Independent reruns resolved adm-zip 0.6.1 and uuid
11.1.1 and passed. This is targeted compatibility evidence, not full tooling certification.

Website wording now distinguishes the tested manual Mac sideload from unverified npm start/stop,
and the walkthrough transcript identifies in-sheet masking as conceptual. The original review
archives remain unchanged.
