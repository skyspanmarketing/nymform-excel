// provider.ts (spec §7.8): the one sender. The AuditedOutbound used here comes from the real
// auditor, fed by payload.ts, so these tests also cover the build -> audit -> send path.
import { describe, expect, it, vi } from "vitest";
import { audit, buildAuditContext, type AuditedOutbound } from "../src/core/auditor";
import { buildStructureOnly, buildSubstituted } from "../src/core/payload";
import { APP_TITLE, CHAT_URL, MESSAGES, send, type SendDeps } from "../src/core/provider";
import { buildSheetContext, inferSchema } from "../src/core/schema";
import { createStandInMap } from "../src/core/transform";
import type { CellValue, ColumnPolicy, RangeData, RangeSpec } from "../src/core/types";

const MODEL = "openai/gpt-6-luna";
// A fake key in the OpenRouter shape, assembled at runtime so the secret scanner still flags real ones.
const KEY = ["sk", "or", "v1", "0123456789abcdef".repeat(4)].join("-");

// ---------------------------------------------------------------------------------------------
// Fixtures

function rangeData(sheet: string, rows: CellValue[][]): RangeData {
  return {
    sheet,
    address: `A1:C${rows.length}`,
    rowIndex: 0,
    columnIndex: 0,
    values: rows,
    text: rows.map((r) => r.map((v) => (v === null ? "" : String(v)))),
    formulas: rows,
    valueTypes: rows.map((r) => r.map((v) => (v === null ? "Empty" : typeof v === "number" ? "Double" : "String"))),
    numberFormat: rows.map((r) => r.map(() => "General")),
  };
}

function audited(mode: "structure_only" | "substituted" = "structure_only"): AuditedOutbound {
  const data = rangeData("Orders", [
    ["Customer", "Region", "Amount"],
    ["Maria Lopez", "East", 83500],
    ["Kenji Watanabe", "West", 41250],
  ]);
  const policies: ColumnPolicy[] = [
    { letter: "A", private: true, treatment: "stand_in" },
    { letter: "B", private: false, treatment: "as_is" },
    { letter: "C", private: true, treatment: "range" },
  ];
  const spec: RangeSpec = { data, hasHeaders: true, columns: inferSchema(data), policies };
  const map = createStandInMap([data]);
  const input = {
    model: MODEL,
    specs: [spec],
    sheetContext: buildSheetContext(data, true, [], { sheets: ["Orders"], tables: [] }),
    question: "Total Amount per Region",
    history: [],
    map,
  };
  const built = mode === "structure_only" ? buildStructureOnly(input) : buildSubstituted(input);
  const outcome = audit(built.outbound, buildAuditContext([spec], map, { model: MODEL }));
  if (!outcome.ok) throw new Error(`fixture blocked: ${outcome.result.error}`);
  return outcome.outbound;
}

const REPLY = {
  id: "gen-1",
  provider: "Azure",
  model: MODEL,
  choices: [
    {
      finish_reason: "stop",
      message: {
        role: "assistant",
        content: '{"kind":"formula","formula":"=SUMIFS(C2:C3,B2:B3,B2)","placement":null,"explanation":"x","assumptions":[]}',
      },
    },
  ],
  usage: { prompt_tokens: 812, completion_tokens: 64, total_tokens: 876 },
};

interface Call {
  url: string;
  init: RequestInit;
}

type Responder = (call: Call) => Response | Promise<Response>;

function fake(respond: Responder) {
  const calls: Call[] = [];
  const fn = vi.fn((input: RequestInfo | URL, init?: RequestInit) => {
    const call = { url: String(input), init: init ?? {} };
    calls.push(call);
    return Promise.resolve().then(() => respond(call));
  });
  return { fetch: fn as unknown as typeof globalThis.fetch, calls, fn };
}

function json(status: number, body: unknown, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json", ...headers } });
}

/** A fetch that never settles on its own; rejects with AbortError when its signal fires if `honor`. */
function hanging(honor: boolean) {
  return fake(
    (call) =>
      new Promise<Response>((_, reject) => {
        if (!honor) return;
        call.init.signal?.addEventListener("abort", () => reject(new DOMException("aborted", "AbortError")));
      }),
  );
}

function deps(f: { fetch: typeof globalThis.fetch }, extra: SendDeps = {}): SendDeps {
  return { fetch: f.fetch, ...extra };
}

// ---------------------------------------------------------------------------------------------

describe("send: request", () => {
  it("T-P1: fetch receives exactly the audited string, with the spec's URL, method and headers", async () => {
    const outbound = audited("substituted");
    const f = fake(() => json(200, REPLY));
    const res = await send(outbound, KEY, undefined, deps(f));

    expect(res.ok).toBe(true);
    expect(f.calls).toHaveLength(1);
    const { url, init } = f.calls[0]!;
    expect(url).toBe("https://openrouter.ai/api/v1/chat/completions");
    expect(CHAT_URL).toBe(url);
    expect(init.method).toBe("POST");
    expect(init.body).toBe(outbound.body);
    expect(typeof init.body).toBe("string");
    expect(new TextEncoder().encode(init.body as string).byteLength).toBe(outbound.bytes);
    expect(init.headers).toEqual({
      Authorization: `Bearer ${KEY}`,
      "Content-Type": "application/json",
      "X-Title": "Nymform for Excel",
    });
    expect(APP_TITLE).toBe("Nymform for Excel");
    expect(init.signal).toBeInstanceOf(AbortSignal);
    expect(init.redirect).toBe("error");
    expect(init.credentials).toBe("omit");
  });

  it("T-P2: an OpenRouter request carries provider zdr in the body that is sent", async () => {
    const outbound = audited();
    const f = fake(() => json(200, REPLY));
    await send(outbound, KEY, undefined, deps(f));
    const sent = f.calls[0]!.init.body as string;
    expect(sent).toContain('"provider":{"zdr":true}');
    expect((JSON.parse(sent) as { provider: unknown }).provider).toEqual({ zdr: true });
  });

  it("trims whitespace around a pasted key", async () => {
    const f = fake(() => json(200, REPLY));
    await send(audited(), `  ${KEY}\n`, undefined, deps(f));
    expect((f.calls[0]!.init.headers as Record<string, string>).Authorization).toBe(`Bearer ${KEY}`);
  });

  it("does not send without a key, or with a key that can't go in a header", async () => {
    for (const key of ["", "   ", "sk-or-v1-abc déf", "sk-or\u0000v1"]) {
      const f = fake(() => json(200, REPLY));
      const res = await send(audited(), key, undefined, deps(f));
      expect(res.ok).toBe(false);
      expect(f.calls).toHaveLength(0);
      if (!res.ok) expect([MESSAGES.noKey, MESSAGES.badKey]).toContain(res.message);
    }
  });

  it("refuses an altered copy of an audited request", async () => {
    const real = audited();
    // A spread copy keeps the brand in its type, so the check has to happen at run time too.
    const copy: AuditedOutbound = { ...real, body: real.body.replace("Total", "Grand total") };
    const f = fake(() => json(200, REPLY));
    const res = await send(copy, KEY, undefined, deps(f));
    expect(res).toMatchObject({ ok: false, message: MESSAGES.unchecked });
    expect(f.calls).toHaveLength(0);
  });
});

describe("send: success", () => {
  it("returns the raw text unchanged, status, token usage, provider name and latency", async () => {
    const raw = JSON.stringify(REPLY, null, 1);
    const f = fake(() => new Response(raw, { status: 200, headers: { "X-Provider-Name": "Azure" } }));
    const times = [1000, 1250.4];
    const res = await send(audited(), KEY, undefined, deps(f, { now: () => times.shift() ?? 1250.4 }));
    expect(res).toEqual({
      ok: true,
      raw,
      status: 200,
      latencyMs: 250,
      tokens: { prompt: 812, completion: 64 },
      providerName: "Azure",
    });
  });

  it("falls back to the provider field in the JSON when the header isn't readable", async () => {
    const f = fake(() => json(200, { ...REPLY, provider: "Azure" }));
    const res = await send(audited(), KEY, undefined, deps(f));
    expect(res.ok && res.providerName).toBe("Azure");
  });

  it("leaves tokens and provider name out when the response doesn't have them", async () => {
    const f = fake(() => json(200, { choices: REPLY.choices }));
    const res = await send(audited(), KEY, undefined, deps(f));
    expect(res.ok).toBe(true);
    if (res.ok) {
      expect(res.tokens).toBeUndefined();
      expect(res.providerName).toBeUndefined();
    }
  });
});

describe("send: errors map to plain messages", () => {
  const cases: [string, () => Response, string][] = [
    ["401", () => json(401, { error: { code: 401, message: "No auth credentials found" } }), MESSAGES.unauthorized],
    ["402", () => json(402, { error: { code: 402, message: "Insufficient credits" } }), MESSAGES.noCredits],
    ["429", () => json(429, { error: { code: 429, message: "Rate limit exceeded" } }), MESSAGES.rateLimited],
    ["500", () => json(500, { error: { code: 500, message: "Internal Server Error" } }), MESSAGES.unavailable],
    ["502", () => new Response("<html>Bad gateway</html>", { status: 502 }), MESSAGES.unavailable],
    ["503", () => json(503, { error: { message: "No instances available" } }), MESSAGES.unavailable],
    ["408", () => json(408, { error: { code: 408, message: "Timed out" } }), MESSAGES.timeout],
    [
      "404 no zero-data-retention endpoint",
      () =>
        json(404, {
          error: {
            code: 404,
            message: "No endpoints found matching your data policy (Zero data retention). Configure: https://openrouter.ai/settings/privacy",
          },
        }),
      "No zero-data-retention route can serve openai/gpt-6-luna right now. Try again later or choose another model in Setup.",
    ],
    [
      "404 no endpoints",
      () => json(404, { error: { code: 404, message: "No endpoints found for openai/gpt-6-luna." } }),
      "No zero-data-retention route can serve openai/gpt-6-luna right now. Try again later or choose another model in Setup.",
    ],
    [
      "400 zdr",
      () => json(400, { error: { code: 400, message: "Provider routing: ZDR is not available for this model" } }),
      "No zero-data-retention route can serve openai/gpt-6-luna right now. Try again later or choose another model in Setup.",
    ],
    [
      "400 with a message",
      () => json(400, { error: { code: 400, message: "  Invalid value\nfor 'response_format'.  " } }),
      "The model provider refused the request (400). Invalid value for 'response_format'.",
    ],
    ["403 without JSON", () => new Response("Forbidden", { status: 403 }), "The model provider refused the request (403)."],
    ["404 without a message", () => json(404, { error: { code: 404 } }), "The model provider refused the request (404)."],
    [
      "200 with an error object and a status code",
      () => json(200, { error: { code: 429, message: "Rate limited upstream" } }),
      MESSAGES.rateLimited,
    ],
    [
      "200 with an error object and a 5xx code",
      () => json(200, { error: { code: 502, message: "Upstream error" } }),
      MESSAGES.unavailable,
    ],
    [
      "200 with an error object and no usable code",
      () => json(200, { error: { code: "server_error", message: "Upstream failed" } }),
      "The model provider returned an error. Upstream failed",
    ],
    ["unexpected status", () => new Response(null, { status: 304 }), "The model provider sent an unexpected response (304)."],
  ];

  for (const [name, respond, message] of cases) {
    it(`${name}`, async () => {
      const f = fake(respond);
      const res = await send(audited(), KEY, undefined, deps(f));
      expect(res.ok).toBe(false);
      if (!res.ok) {
        expect(res.message).toBe(message);
        expect(res.latencyMs).toBeGreaterThanOrEqual(0);
      }
      // No automatic retries.
      expect(f.calls).toHaveLength(1);
    });
  }

  it("keeps the HTTP status on the error", async () => {
    const res = await send(audited(), KEY, undefined, deps(fake(() => json(429, {}))));
    expect(res).toMatchObject({ ok: false, status: 429 });
  });

  it("shortens a long provider message to 160 characters", async () => {
    const long = "Bad request: " + "parameter problem ".repeat(30);
    const res = await send(audited(), KEY, undefined, deps(fake(() => json(400, { error: { message: long } }))));
    expect(res.ok).toBe(false);
    if (!res.ok) {
      const prefix = "The model provider refused the request (400). ";
      expect(res.message.startsWith(prefix)).toBe(true);
      const detail = res.message.slice(prefix.length);
      expect(detail.length).toBeLessThanOrEqual(160);
      expect(detail.endsWith("…")).toBe(true);
    }
  });

  it("says it couldn't reach the host when the network fails", async () => {
    const f = fake(() => Promise.reject(new TypeError("Failed to fetch")));
    const res = await send(audited(), KEY, undefined, deps(f));
    expect(res).toMatchObject({ ok: false, message: "Couldn't reach openrouter.ai." });
    expect(res).not.toHaveProperty("status");
    expect(f.calls).toHaveLength(1);
  });

  it("treats a fetch that throws synchronously as a network failure", async () => {
    const throwing = (() => {
      throw new TypeError("Invalid header");
    }) as unknown as typeof globalThis.fetch;
    const res = await send(audited(), KEY, undefined, { fetch: throwing });
    expect(res).toMatchObject({ ok: false, message: "Couldn't reach openrouter.ai." });
  });
});

describe("send: timeout and cancel", () => {
  it("times out after the limit even when fetch ignores the signal", async () => {
    const f = hanging(false);
    const res = await send(audited(), KEY, undefined, deps(f, { timeoutMs: 20 }));
    expect(res).toMatchObject({ ok: false, message: "The model provider took too long to reply. Try again." });
    expect(f.calls[0]!.init.signal?.aborted).toBe(true);
  });

  it("times out when fetch honors the signal", async () => {
    const res = await send(audited(), KEY, undefined, deps(hanging(true), { timeoutMs: 20 }));
    expect(res).toMatchObject({ ok: false, message: MESSAGES.timeout });
  });

  it("times out while reading a slow body", async () => {
    const slow = {
      ok: true,
      status: 200,
      headers: new Headers(),
      text: () => new Promise<string>(() => {}),
    } as unknown as Response;
    const res = await send(audited(), KEY, undefined, deps(fake(() => slow), { timeoutMs: 20 }));
    expect(res).toMatchObject({ ok: false, message: MESSAGES.timeout });
  });

  it("uses a 60-second limit by default", async () => {
    vi.useFakeTimers();
    try {
      const pending = send(audited(), KEY, undefined, deps(hanging(false)));
      await vi.advanceTimersByTimeAsync(59_000);
      let settled = false;
      void pending.then(() => (settled = true));
      await vi.advanceTimersByTimeAsync(0);
      expect(settled).toBe(false);
      await vi.advanceTimersByTimeAsync(1_000);
      await expect(pending).resolves.toMatchObject({ ok: false, message: MESSAGES.timeout });
    } finally {
      vi.useRealTimers();
    }
  });

  it("says 'Request cancelled.' when the caller aborts mid-flight", async () => {
    const controller = new AbortController();
    const f = hanging(true);
    const pending = send(audited(), KEY, controller.signal, deps(f));
    await new Promise((r) => setTimeout(r, 5));
    controller.abort();
    await expect(pending).resolves.toMatchObject({ ok: false, message: "Request cancelled." });
    expect(f.calls[0]!.init.signal?.aborted).toBe(true);
  });

  it("does not send when the caller's signal is already aborted", async () => {
    const controller = new AbortController();
    controller.abort();
    const f = fake(() => json(200, REPLY));
    const res = await send(audited(), KEY, controller.signal, deps(f));
    expect(res).toMatchObject({ ok: false, message: MESSAGES.cancelled });
    expect(f.calls).toHaveLength(0);
  });

  it("the caller's signal is not the one given to fetch, and a finished send leaves no listener behind", async () => {
    const controller = new AbortController();
    const add = vi.spyOn(controller.signal, "addEventListener");
    const remove = vi.spyOn(controller.signal, "removeEventListener");
    const f = fake(() => json(200, REPLY));
    await send(audited(), KEY, controller.signal, deps(f));
    expect(f.calls[0]!.init.signal).not.toBe(controller.signal);
    expect(add).toHaveBeenCalledTimes(1);
    expect(remove).toHaveBeenCalledTimes(1);
  });
});

describe("send: the key never appears in a message", () => {
  const echoes: [string, () => Response][] = [
    ["400 echo", () => json(400, { error: { message: `Key ${KEY} is malformed` } })],
    ["400 bearer echo", () => json(400, { error: { message: `Header was Authorization: Bearer ${KEY}` } })],
    ["404 zdr echo", () => json(404, { error: { message: `No endpoints found matching your data policy for ${KEY}` } })],
    ["200 error echo", () => json(200, { error: { message: `bad key ${KEY.slice(0, 20)}` } })],
    ["provider name echo", () => json(200, { ...REPLY, provider: KEY })],
    ["401", () => json(401, { error: { message: KEY } })],
    ["network", () => Promise.reject(new TypeError(KEY)) as unknown as Response],
  ];

  for (const [name, respond] of echoes) {
    it(name, async () => {
      const res = await send(audited(), KEY, undefined, deps(fake(respond)));
      const shown = res.ok ? (res.providerName ?? "") : res.message;
      expect(shown).not.toContain(KEY);
      expect(shown).not.toContain(KEY.slice(0, 16));
      expect(shown).not.toContain(KEY.slice(-16));
    });
  }

  it("also for a key without the usual prefix", async () => {
    const key = "custom-key-7f3a9b2c1d";
    const res = await send(audited(), key, undefined, deps(fake(() => json(400, { error: { message: `bad ${key}` } }))));
    expect(res.ok).toBe(false);
    if (!res.ok) {
      expect(res.message).not.toContain(key);
      expect(res.message).toBe("The model provider refused the request (400). bad [key]");
    }
  });
});

describe("provider: message hygiene", () => {
  it("drops bidi overrides and zero-width characters from the provider's detail and name", async () => {
    const f = fake(() => json(400, { error: { message: "bad\u202Eguest\u200b request" } }));
    const res = await send(audited(), KEY, undefined, deps(f));
    expect(res.ok).toBe(false);
    if (!res.ok) {
      expect(res.message).not.toMatch(/[\u200b\u202e]/);
      expect(res.message).toContain("badguest request");
    }
  });
});

