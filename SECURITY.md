# Security

Nymform for Excel is alpha software. This page says what it protects, what it doesn't, and how to
report a problem.

## Reporting a problem

Use GitHub's private vulnerability reporting: on this repository, open **Security → Report a
vulnerability**. Please don't open a public issue for a security problem. Include the footer line of
the task pane (the version and build, or "local build"), the steps, and what you expected. Use synthetic data in any
example; never send real personal data.

## What Nymform does

| Threat | What Nymform does |
| --- | --- |
| Private values reach the model provider | Structure-only mode (the default) sends headers, types and counts, never data cells. Sample-rows mode replaces every value in a column you mark private with a stand-in (`PERSON_014`), a coarse range (`80000-89999`), a month (`2026-03`), or leaves the column out. Private values you type in your question are replaced too. |
| A bug puts a private value into a request | The auditor checks the exact request text before anything is sent: every private cell in the rows must be a stand-in or a rendering (structural check), and the full text is scanned for private values of 4 or more characters, including JSON-escaped, digits-only and word forms. Any hit, error or exception blocks the send. The sender only accepts a request the auditor passed. |
| You can't see what leaves your machine | "What gets sent" shows the exact request, byte for byte, before you click Send. That includes the one follow-up Nymform offers when a reply can't be read: it is shown the same way and sent only if you choose Send. The Log keeps what was sent for this session and can be exported. |
| The model's formula sends data somewhere | Every formula passes a gate before and after stand-ins are filled back in: an allowlist of worksheet functions (no WEBSERVICE, IMAGE, HYPERLINK, STOCKHISTORY, COPILOT, PY, CUBE functions, INDIRECT, OFFSET and more), no URLs in strings, no external workbooks, no defined names, and references only to the ranges you chose (whole columns of those ranges included; see the limits). A formula that is filled down is checked for every row it reaches. Nothing is written until you click Insert formula. |
| Insert damages your data | Insert refuses cells inside the ranges you selected, fills down only from the first data row, and asks before replacing cells that already hold values, because Excel can't undo an add-in's writes. After writing it says whether the result was checked, shows an Excel error, isn't calculated yet (manual calculation) or couldn't be read back, and never writes again on its own. |
| The workbook changes after Nymform read it | Nymform reads your range when the pane opens and on Refresh from selection. Before sending and before inserting it reads the same ranges, and the workbook's tables and defined names, again; if anything changed, nothing is sent or written until you refresh. Otherwise a name typed into a private column after the read wouldn't be known as private. |
| Text in the sheet tries to instruct the model (prompt injection) | The model is told that headers, notes and cell text are data. The gate still enforces every rule above whatever the model was told. |
| The add-in is pointed at a hostile server | The endpoint is fixed when the add-in is built and shown read-only. The page's Content Security Policy only allows connections to that endpoint. |
| Your API key leaks | The key is kept in memory for the session only, sent only to the endpoint, never logged or stored, and gone when you close the pane. |
| The provider keeps your prompts | Requests to OpenRouter ask for zero-data-retention routing (`provider: { zdr: true }`). |
| Someone collects usage data | There is no Nymform server and no telemetry, analytics or error reporting. Evaluation metadata (counts and pass/fail codes, never content) stays in the pane unless you export it. |
| The served code differs from this repository | The release is built by CI from a tagged commit. The pane footer and `build-info.json` show the commit. |

## Stated limits

- **Short values.** The text scan skips values under 4 characters, because they would match
  everywhere. In rows they are still covered by the structural check. In free text (your question,
  earlier turns) a short private value such as `Al` is only replaced if you type it exactly; the scan
  won't catch it otherwise.
- **Numbers and parts of IDs.** A whole private value, and any run of 7 or more digits in it, is
  scanned for everywhere. A private number under 7 digits is found even inside a longer number
  (`934961234` holds ZIP `93496`, `Transfer to 48392001` holds account `483920`), except in a
  sample-row cell that holds a number, where it doesn't count with another digit directly next to it
  or as the decimals (ZIP `93496` in a longitude `-120.93496`). Numbers in cells matched unrelated
  numbers so often that sample rows were unusable; a number cell in a column you didn't mark private
  is that column's data. A run of 4 to 6 digits
  inside a value (the `0199` of `(831) 555-0199`) is scanned for only in your question, notes,
  headers, names and earlier replies, and a run that reads as a year (1900 to 2100) is skipped there,
  because questions name years all the time.
- **Dates.** Private dates are found wherever they are written in the request, in the forms people
  use: numbers in any order separated by `/`, `-`, `.`, `_`, `\` or look-alikes, two- or four-digit
  years, month names and abbreviations in English and several European languages, ordinals,
  run-together forms such as `15MAR1985`, and Chinese, Japanese and Korean forms. A date is read from
  the stored value and from what the cell shows, including dates stored as text and as 8-digit
  numbers or text (`19850315`) when most of the column reads as dates. Looser forms count only in
  your question, notes, headers, names and earlier replies, not in sample-row cells: numbers
  separated only by spaces (`15 03 1985`) and eight digits run together. Day-first or month-first
  digits run together in a question (`15031985`, `031585`), 6-digit stored dates (`910919`),
  Roman-numeral months and dates without a year are not caught.
- **Encoded forms.** The scan looks for exact, normalized, digits-only, JSON-escaped and word forms,
  amounts as they are commonly written, and dates as above. It does not catch values that are
  abbreviated, misspelled, split, reversed, translated or encoded another way.
- **Parts of private values in your question.** When a word you type is part of a private value
  (a first name, an email's part before the @, a phone's last digits), Nymform lists the values it
  could be, in the pane only, and puts the stand-in of the one you choose in the question. A part with
  a single candidate that names a person, email or ID is read as it automatically, and the preview
  says so. Show in sheet only selects cells; it doesn't change the workbook. If you keep the part as
  typed, the check blocks the request as before. A wrong choice sends the wrong stand-in, never the
  value.
- **Unmarked columns.** Columns you don't mark private are sent as they are in sample-rows mode.
  Suggestions are only defaults; your checkbox decides.
- **What is always sent.** Headers (or their aliases), sheet and table names, addresses, notes you
  write, simple counts (blanks, distinct values, average words), and your question with private values
  replaced. Structure-only mode sends no data cells.
- **The header row.** Row 1 of the selection is sent as the headers. When row 1 looks like data (a
  number, a date, an amount, an email address, a phone number, an ID, an IBAN, a postcode or a street
  address), Nymform treats it as data, and a private column whose header looks like data is not sent
  under that header. A first row of ordinary text that is really data (a selection that starts with
  a name) can't be told apart: check the headers on the Columns screen.
- **Other spellings and scripts.** The scan folds case (including Greek final sigma and Turkish
  dotted I), accents (`Jose Munoz` for `José Muñoz`), `ß` as `ss`, `ä` `ö` `ü` as `ae` `oe` `ue`,
  ligatures, Unicode forms (NFC and NFD), curly and look-alike apostrophes and dashes, full-width
  letters and digits, Arabic-Indic, Persian and Devanagari digits, `_` as a word break, and CJK names
  written with or without their space. Invisible characters (zero-width spaces and joiners, direction
  marks, soft hyphens, control characters) are ignored inside a value and count as a break between
  two. Names are also found without their apostrophe or hyphen (`OBrien`, `AnneMarie`), an email
  address by its part before the `@` when that part has a digit or `.`, `_` or `-`, and a code that
  mixes letters and digits with some or all of its separators left out (`AB1234CD`, `AB 1234CD` for
  `AB-1234-CD`). Amounts are matched plain, grouped (`83,500`, `83.500`, `83 500`, `83_500`, lakh),
  with a decimal comma and as exact short forms (`83.5k`), whatever the cell's format shows. Phone
  numbers are matched with or without a country code, trunk `0` or extension. Other
  transliterations, nicknames and spellings, and a code typed like a cell range (`AB12:CD34`), are
  not caught.
- **Whole columns.** A formula may use whole-column references such as `E:E` for the columns you
  chose, so it can read rows of those columns below your selection. The Result screen names them.
  This happens in Excel only: those rows aren't sent. It can't reach other columns, sheets or
  workbooks, and spill references such as `A2#` are refused because their size isn't known.
- **Restored values are text.** A stand-in comes back as text inside the formula, as the cell shows it
  (`00123`, `3/15/1985`). A private numeric ID or date compared with numbers may not match (Excel
  shows #N/A), and a restored value containing `*`, `?` or `~` acts as a wildcard in SUMIFS, COUNTIF
  and MATCH. The Result screen marks both before you insert. A formula that needs a stand-in this
  session didn't create can't be inserted or copied. A restored value longer than 255 characters
  inside a text constant passes the check and is written, but Excel can't calculate a text constant
  that long: the cell shows #VALUE!, and the Result screen reports that Excel shows an error in the
  result.
- **Changes the check doesn't notice.** Rows added below a range that isn't an Excel table aren't
  noticed before Send or Insert, and neither is a new result of a formula cell inside the range: its
  formula is compared, not its value, so a private column filled by formulas that read other sheets
  can change unnoticed. Refresh from selection picks both up.
- **Lookup ranges.** Each lookup range has its own private columns. If a lookup range overlaps the
  selection, mark the same columns private in both.
- **Small groups.** Counts, ranges and months can still identify a person in a small group, for
  example the only row in a region.
- **Fixed text.** A match that lies entirely inside text Nymform itself writes (the system prompt, JSON
  key names, the configured model ID, fixed values such as `true` or `structure_only`) is not treated
  as a leak, because it would be sent whatever your data is. Counts and row numbers Nymform computes
  are checked against the selection (a row number must match the address, a count must be within the
  number of rows) and then treated the same way, as is a range address equal to the selection's (the
  cell part only, never the sheet name). Everything else is scanned.
- **False blocks.** A private value that also appears somewhere Nymform can't tell apart from it
  blocks the request: the same amount or code in a column you didn't mark private, a word of a private
  name that is also a word in your question, a number in your question or in a text cell that contains
  a private number (an order ID `SO-67912` holding a private store code `7912`) or ends a private
  phone number or ID, or a date you type that can be read as a private date
  (`3/4/2026` is read both day-first and month-first). Nymform blocks rather than guess: mark that
  column private too, rephrase, or use structure only. A header renamed under Header sent stays
  hidden, words included: with Order date renamed to Month, the word "order" in another header or
  in your question blocks, and so does "order dates", which the question's rewrite to the new name
  doesn't catch. Typing the same header again in another case or spacing ("Order Date") doesn't
  hide it, so that doesn't block.
- **Dates in 1904 workbooks.** Nymform reads a date both from the stored value, assuming Excel's
  default 1900 date system, and from what the cell displays. In a workbook on the 1904 date system the
  dates worked out from the stored value are four years off. So when a private date cell doesn't
  display a full date (`#####`, or a format such as `15-Mar`), the check may not recognise that date
  typed in full in your question, and it can be sent; and a private date column sent in sample rows
  as months shows months four years off (a wrong description, not a leak). Nymform can't tell which
  date system a workbook uses: the Office.js property that reports it, `Workbook.use1904DateSystem`,
  is preview-only (ExcelApi BETA) in Microsoft's reference as of 2026-09-28, and Nymform uses released
  API sets only (ExcelApi 1.9; 1.12 and 1.18 where available). So it gives no warning. A value put back
  into a formula uses the text the cell shows, so restoring isn't affected.
- **Prompt injection** can still mislead the model's answer, even though it can't get a dangerous
  formula past the gate.
- **Provider processing.** Zero data retention limits storage, not processing: the request still
  reaches the provider.
- **Your workbook.** Nymform doesn't protect the workbook itself, your machine, or Excel.

## Scope

In scope: the code in this repository and the release build it produces (the task pane, the
manifests, the auditor, the formula gate, the request path).

Out of scope: OpenRouter and model providers, Microsoft Excel and Office, your device and browser,
social engineering, denial of service, and findings that need real personal data to demonstrate.
