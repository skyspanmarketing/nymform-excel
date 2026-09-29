// The one sender (spec §7.8, invariant 2). It sends the exact string the auditor checked, byte for
// byte, to the build-time endpoint. No retries, no storage; the key is used for this call only.
import { ENDPOINT, ENDPOINT_HOST } from "../config";
import { isAudited, type AuditedOutbound } from "./auditor";

export interface SendOk {
  ok: true;
  /** Raw response text, unchanged. */
  raw: string;
  status: number;
  latencyMs: number;
  tokens?: { prompt: number; completion: number };
  /** OpenRouter's X-Provider-Name response header, when present. */
  providerName?: string;
}

export interface SendErr {
  ok: false;
  /** Plain message for the UI. Never contains the key. */
  message: string;
  status?: number;
  latencyMs: number;
}

export interface SendDeps {
  fetch?: typeof globalThis.fetch;
  timeoutMs?: number;
  now?: () => number;
}

export const TIMEOUT_MS = 60_000;
export const APP_TITLE = "Nymform for Excel";
export const CHAT_URL = `${ENDPOINT}/chat/completions`;

const DETAIL_MAX = 160;
const PROVIDER_NAME_MAX = 64;
const MODEL_ID_RE = /^[\w.~-]+\/[\w.:~-]+$/;
/** OpenRouter's refusal when routing with `provider: { zdr: true }` finds no endpoint. */
const NO_ZDR_RE = /zero[\s_-]*data[\s_-]*retention|\bzdr\b|data[\s_-]*polic|no\s+endpoints?\b/i;

export const MESSAGES = {
  unchecked: "This request wasn't checked, so it wasn't sent.",
  noKey: "Add your API key in Setup.",
  badKey: "Your API key has characters that can't be sent. Paste it again in Setup.",
  unauthorized: "Your API key was rejected. Check it in Setup.",
  noCredits: "Your account is out of credits.",
  rateLimited: "Rate limited. Wait a minute and try again.",
  unavailable: "The model provider isn't responding. Try again shortly.",
  timeout: "The model provider took too long to reply. Try again.",
  cancelled: "Request cancelled.",
  network: (host: string) => `Couldn't reach ${host}.`,
  noZdrRoute: (model: string) =>
    `No zero-data-retention route can serve ${model} right now. Try again later or choose another model in Setup.`,
  refused: (status: number) => `The model provider refused the request (${status}).`,
  providerError: "The model provider returned an error.",
  unexpected: (status: number) => `The model provider sent an unexpected response (${status}).`,
} as const;

export async function send(
  outbound: AuditedOutbound,
  key: string,
  signal?: AbortSignal,
  deps: SendDeps = {},
): Promise<SendOk | SendErr> {
  const now = deps.now ?? defaultNow;
  const started = now();
  const elapsed = () => Math.max(0, Math.round(now() - started));
  const secret = typeof key === "string" ? key.trim() : "";
  const fail = (message: string, status?: number): SendErr => {
    const err: SendErr = { ok: false, message: redact(message, secret), latencyMs: elapsed() };
    if (status !== undefined) err.status = status;
    return err;
  };

  // Read once: the string sent is the string the auditor returned. The auditor freezes what it
  // returns, so an altered copy (for example `{ ...audited, body }`) is refused here.
  const body: unknown = outbound?.body;
  if (!isAudited(outbound) || typeof body !== "string" || outbound.audit?.ok !== true || !Object.isFrozen(outbound)) {
    return fail(MESSAGES.unchecked);
  }
  if (secret === "") return fail(MESSAGES.noKey);
  if (!/^[\x21-\x7e]+$/.test(secret)) return fail(MESSAGES.badKey);
  if (signal?.aborted) return fail(MESSAGES.cancelled);

  const controller = new AbortController();
  const state: { stop: "timeout" | "cancelled" | null } = { stop: null };
  const onCallerAbort = () => {
    state.stop ??= "cancelled";
    controller.abort();
  };
  signal?.addEventListener("abort", onCallerAbort, { once: true });
  const timer = setTimeout(() => {
    state.stop ??= "timeout";
    controller.abort();
  }, deps.timeoutMs ?? TIMEOUT_MS);

  const doFetch = deps.fetch ?? ((input: RequestInfo | URL, init?: RequestInit) => globalThis.fetch(input, init));
  let res: Response;
  let text: string;
  try {
    res = await untilAborted(
      doFetch(CHAT_URL, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${secret}`,
          "Content-Type": "application/json",
          "X-Title": APP_TITLE,
        },
        body,
        signal: controller.signal,
        credentials: "omit",
        // The audited body goes to the build-time endpoint and nowhere else.
        redirect: "error",
      }),
      controller.signal,
    );
    text = await untilAborted(res.text(), controller.signal);
  } catch {
    if (state.stop === "cancelled") return fail(MESSAGES.cancelled);
    if (state.stop === "timeout") return fail(MESSAGES.timeout);
    return fail(MESSAGES.network(ENDPOINT_HOST));
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener("abort", onCallerAbort);
  }

  const json = parseJson(text);
  const error = isObject(json) && isObject(json.error) ? json.error : null;
  if (!res.ok || error) {
    const code = res.ok && error && typeof error.code === "number" ? error.code : res.status;
    return fail(errorMessage(code, res.ok, detailOf(error, secret), body), res.status);
  }

  const ok: SendOk = { ok: true, raw: text, status: res.status, latencyMs: elapsed() };
  const tokens = tokensOf(json);
  if (tokens) ok.tokens = tokens;
  const providerName = providerNameOf(res, json, secret);
  if (providerName) ok.providerName = providerName;
  return ok;
}

// ---------------------------------------------------------------------------------------------
// Errors

function errorMessage(code: number, httpOk: boolean, detail: string | undefined, body: string): string {
  if (code === 401) return MESSAGES.unauthorized;
  if (code === 402) return MESSAGES.noCredits;
  if (code === 429) return MESSAGES.rateLimited;
  if (code === 408) return MESSAGES.timeout;
  if (code >= 500 && code <= 599) return MESSAGES.unavailable;
  if (code >= 400 && code <= 499) {
    if (detail && NO_ZDR_RE.test(detail)) return MESSAGES.noZdrRoute(modelOf(body));
    return withDetail(MESSAGES.refused(code), detail);
  }
  // A 200 whose JSON carries an error object without a usable code.
  if (httpOk) return withDetail(MESSAGES.providerError, detail);
  return withDetail(MESSAGES.unexpected(code), detail);
}

function withDetail(message: string, detail: string | undefined): string {
  return detail ? `${message} ${detail}` : message;
}

/** The provider's own short error message: one line, at most DETAIL_MAX characters, key removed. */
function detailOf(error: Record<string, unknown> | null, secret: string): string | undefined {
  if (!error || typeof error.message !== "string") return undefined;
  const line = oneLine(redact(error.message, secret));
  return line === "" ? undefined : clip(line, DETAIL_MAX);
}

/** The model named in the body, when it has the shape of a model ID. */
function modelOf(body: string): string {
  const parsed = parseJson(body);
  const model = isObject(parsed) ? parsed.model : undefined;
  return typeof model === "string" && MODEL_ID_RE.test(model) ? model : "this model";
}

/** Removes the key, and anything shaped like a key or bearer credential, from text shown to the user. */
function redact(text: string, secret: string): string {
  let out = text;
  // A very short "key" would mangle ordinary words; real keys are far longer.
  if (secret.length >= 8) out = out.split(secret).join("[key]");
  return out.replace(/\bBearer\s+\S+/gi, "Bearer [key]").replace(/\bsk-[\w-]{6,}/g, "[key]");
}

function oneLine(text: string): string {
  return (
    text
      // eslint-disable-next-line no-control-regex
      .replace(/[\u0000-\u001f\u007f-\u009f\u2028\u2029]+/g, " ")
      // Bidi overrides and zero-width characters could make the line read differently from what it says.
      .replace(/[\u200b-\u200f\u202a-\u202e\u2060-\u2064\u2066-\u2069\ufeff]/g, "")
      .replace(/\s+/g, " ")
      .trim()
  );
}

function clip(text: string, max: number): string {
  if (text.length <= max) return text;
  let cut = text.slice(0, max - 1);
  // Don't split a surrogate pair.
  if (/[\ud800-\udbff]$/.test(cut)) cut = cut.slice(0, -1);
  return `${cut.trimEnd()}…`;
}

// ---------------------------------------------------------------------------------------------
// Response metadata

function tokensOf(json: unknown): SendOk["tokens"] {
  const usage = isObject(json) ? json.usage : undefined;
  if (!isObject(usage)) return undefined;
  const prompt = usage.prompt_tokens;
  const completion = usage.completion_tokens;
  if (!isCount(prompt) || !isCount(completion)) return undefined;
  return { prompt, completion };
}

/** X-Provider-Name, or the "provider" field OpenRouter puts in the JSON when the header isn't readable. */
function providerNameOf(res: Response, json: unknown, secret: string): string | undefined {
  let name: unknown;
  try {
    name = res.headers?.get("x-provider-name");
  } catch {
    name = null;
  }
  if (typeof name !== "string" || name.trim() === "") name = isObject(json) ? json.provider : undefined;
  if (typeof name !== "string") return undefined;
  const line = oneLine(redact(name, secret));
  return line === "" ? undefined : clip(line, PROVIDER_NAME_MAX);
}

// ---------------------------------------------------------------------------------------------
// Helpers

function defaultNow(): number {
  return typeof performance !== "undefined" && typeof performance.now === "function" ? performance.now() : Date.now();
}

/** Settles with `p`, or rejects as soon as `signal` aborts (even if `p` ignores the signal). */
function untilAborted<T>(p: Promise<T>, signal: AbortSignal): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    if (signal.aborted) {
      reject(new Error("aborted"));
      return;
    }
    const onAbort = () => reject(new Error("aborted"));
    signal.addEventListener("abort", onAbort, { once: true });
    p.then(
      (v) => {
        signal.removeEventListener("abort", onAbort);
        resolve(v);
      },
      (e: unknown) => {
        signal.removeEventListener("abort", onAbort);
        reject(e instanceof Error ? e : new Error("failed"));
      },
    );
  });
}

function parseJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

function isObject(v: unknown): v is Record<string, unknown> {
  return v !== null && typeof v === "object" && !Array.isArray(v);
}

function isCount(v: unknown): v is number {
  return typeof v === "number" && Number.isFinite(v) && v >= 0;
}
