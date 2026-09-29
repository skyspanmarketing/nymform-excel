// Shared setup for the end-to-end tests: the built task pane (dist/) in Chromium, with the fake
// Excel host (fakeOffice.js) served in place of Office.js and a scripted model endpoint in place
// of OpenRouter. Every request that reaches the endpoint is kept, so tests can check it
// independently of the auditor.
import { expect, type Page, type Route } from "@playwright/test";
import { readFileSync } from "node:fs";
import { join } from "node:path";

// Playwright runs from the repository root (playwright.config.ts).
const read = (rel: string) => readFileSync(join(process.cwd(), rel), "utf8");

export const FAKE_OFFICE = read("tests/e2e/fakeOffice.js");
export const ORDERS = JSON.parse(read("tests/e2e/fixtures/orders.json")) as FakeWorkbook;
export const PRIVATE = JSON.parse(read("bench/workbooks/private-values.json")) as Record<
  string,
  { privateLetters: string[]; values: string[]; canaries: string[] }
>;

export const CHAT_URL = "https://openrouter.ai/api/v1/chat/completions";

export interface FakeWorkbook {
  sheets: { name: string; cells: Record<string, { value: unknown; text: string; type: string; format: string }> }[];
  selection: { sheet: string; address: string };
  [k: string]: unknown;
}

export interface Sent {
  body: string;
  headers: Record<string, string>;
}

/** A scripted reply: model content (string), or a full HTTP answer. */
export type ModelAnswer = string | { status: number; json: unknown; headers?: Record<string, string> };

export interface PaneOptions {
  workbook?: FakeWorkbook;
  /** Called for each request that reaches the endpoint, in order. Default: a GROUPBY formula. */
  model?: (sent: Sent, index: number) => ModelAnswer;
  /** Serve no Office.js at all, as when the page is opened outside Excel. */
  noOffice?: boolean;
  /** Let requests through to the real endpoint instead of the scripted model. */
  live?: boolean;
}

export interface Pane {
  sent: Sent[];
  consoleErrors: string[];
  cspViolations: string[];
}

const CORS = {
  "access-control-allow-origin": "*",
  "access-control-allow-headers": "authorization,content-type,x-title",
  "access-control-expose-headers": "X-Provider-Name",
};

export function reply(content: Record<string, unknown>): string {
  return JSON.stringify({ placement: null, assumptions: [], ...content });
}

export function chatJson(content: string): unknown {
  return {
    id: "gen-test",
    choices: [{ message: { role: "assistant", content }, finish_reason: "stop" }],
    usage: { prompt_tokens: 900, completion_tokens: 80 },
  };
}

export async function openPane(page: Page, opts: PaneOptions = {}): Promise<Pane> {
  const pane: Pane = { sent: [], consoleErrors: [], cspViolations: [] };
  page.on("console", (m) => {
    const text = m.text();
    if (/Content Security Policy|Refused to connect|Refused to load/i.test(text)) pane.cspViolations.push(text);
    else if (m.type() === "error" && !/Failed to load resource/.test(text)) pane.consoleErrors.push(text);
  });
  page.on("pageerror", (e) => pane.consoleErrors.push(`pageerror: ${e.message}`));

  await page.addInitScript((wb) => {
    (window as unknown as { __NYMFORM_FAKE_WORKBOOK__: unknown }).__NYMFORM_FAKE_WORKBOOK__ = wb;
  }, opts.workbook ?? ORDERS);
  await page.route("https://appsforoffice.microsoft.com/**", (route) =>
    route.fulfill({ contentType: "text/javascript", body: opts.noOffice ? "/* no Office.js outside Excel */" : FAKE_OFFICE }),
  );
  if (!opts.live) {
    const model = opts.model ?? (() => reply({ kind: "formula", formula: "=GROUPBY(C2:C501,F2:F501,SUM)", explanation: "Totals Amount by Region." }));
    await page.route(CHAT_URL, async (route: Route) => {
      const req = route.request();
      if (req.method() === "OPTIONS") return route.fulfill({ status: 204, headers: CORS });
      const sent: Sent = { body: req.postData() ?? "", headers: await req.allHeaders() };
      const index = pane.sent.push(sent) - 1;
      const answer = model(sent, index);
      if (typeof answer === "string") {
        return route.fulfill({
          status: 200,
          headers: { ...CORS, "content-type": "application/json", "X-Provider-Name": "Azure" },
          body: JSON.stringify(chatJson(answer)),
        });
      }
      return route.fulfill({
        status: answer.status,
        headers: { ...CORS, "content-type": "application/json", ...(answer.headers ?? {}) },
        body: JSON.stringify(answer.json),
      });
    });
  } else {
    page.on("request", (req) => {
      if (req.url() === CHAT_URL && req.method() === "POST") pane.sent.push({ body: req.postData() ?? "", headers: req.headers() });
    });
  }
  await page.goto("/taskpane.html");
  return pane;
}

/** Setup: paste a key (and a model, when given) and continue to Columns. */
export async function enterKey(page: Page, key = "fake-key-for-tests", model?: string): Promise<void> {
  await page.getByLabel("API key").fill(key);
  if (model !== undefined) await page.getByLabel("Model").fill(model);
  await page.getByRole("button", { name: "Continue" }).click();
  await expect(page.getByRole("heading", { name: "Columns" })).toBeVisible();
}

/** Ask: pick the mode, type the question and preview. */
export async function preview(page: Page, question: string, mode: "structure" | "rows" = "structure"): Promise<void> {
  await page.getByRole("tab", { name: "Ask" }).click();
  if (mode === "rows") await page.getByText("Include sample rows", { exact: true }).click();
  else await page.getByText("Structure only", { exact: true }).click();
  await page.getByLabel("Question").fill(question);
  await page.getByRole("button", { name: "Preview" }).click();
  await expect(page.getByRole("heading", { name: "What gets sent" })).toBeVisible();
}

/** The user turn of a request body, parsed. */
export function userTurn(body: string): Record<string, unknown> {
  const parsed = JSON.parse(body) as { messages: { role: string; content: string }[] };
  const last = parsed.messages[parsed.messages.length - 1]!;
  return JSON.parse(last.content) as Record<string, unknown>;
}

/**
 * An independent leak check (not the auditor's code): every private value and canary of 4+
 * characters from the generator's own list, searched case-insensitively in the body, plain and
 * JSON-escaped once and twice. Returns the offending values' indexes, never the values.
 */
export function leakedValues(body: string, workbook = "orders"): number[] {
  const info = PRIVATE[workbook]!;
  const hay = body.toLowerCase();
  const found: number[] = [];
  [...info.values, ...info.canaries].forEach((v, i) => {
    if (v.trim().length < 4) return;
    const once = JSON.stringify(v).slice(1, -1);
    const twice = JSON.stringify(once).slice(1, -1);
    if ([v, once, twice].some((f) => hay.includes(f.toLowerCase()))) found.push(i);
  });
  return found;
}

export function cellText(wb: FakeWorkbook, sheet: string, addr: string): string {
  const s = wb.sheets.find((x) => x.name === sheet)!;
  return s.cells[addr]!.text;
}
