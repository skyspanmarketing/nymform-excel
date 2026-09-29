// reply.ts (spec §7.9): pull the content out of a chat-completions response, strip fences, parse,
// validate against ModelReply with zod.
import { describe, expect, it } from "vitest";
import { extractContent, ModelReplySchema, parseContent, parseReply, REPLY_ERRORS, stripFences } from "../src/core/reply";

const FORMULA = {
  kind: "formula",
  formula: "=SUMIFS(F2:F501,C2:C501,C2)",
  placement: { cell: "G2", fill_down: true },
  explanation: "Adds Amount for each row's Region.",
  assumptions: ["Amounts are in USD."],
};

function response(content: unknown, finishReason: string | null = "stop", extra: Record<string, unknown> = {}): string {
  return JSON.stringify({
    id: "gen-1",
    choices: [{ index: 0, finish_reason: finishReason, message: { role: "assistant", content, ...extra } }],
    usage: { prompt_tokens: 10, completion_tokens: 5 },
  });
}

describe("extractContent", () => {
  it("returns choices[0].message.content and the finish reason", () => {
    const content = JSON.stringify(FORMULA);
    expect(extractContent(response(content))).toEqual({ content, finishReason: "stop" });
  });

  it("joins the text parts of a content array", () => {
    const raw = response([
      { type: "text", text: '{"kind":"clarify",' },
      { type: "text", text: '"explanation":"Which year?"}' },
    ]);
    expect(extractContent(raw).content).toBe('{"kind":"clarify","explanation":"Which year?"}');
  });

  it("treats empty content after the token cap (finish_reason length) as an error with a plain message", () => {
    for (const content of [null, "", "   "]) {
      const out = extractContent(response(content, "length"));
      expect(out.content).toBeNull();
      expect(out.finishReason).toBe("length");
      expect(out.error).toBe(REPLY_ERRORS.lengthCap);
    }
  });

  it("treats other empty content as an error", () => {
    expect(extractContent(response(null, "stop")).error).toBe(REPLY_ERRORS.empty);
    expect(extractContent(response(null, "content_filter")).error).toBe(REPLY_ERRORS.filtered);
    expect(extractContent(response(null, "stop", { refusal: "I can't help with that." })).error).toBe(REPLY_ERRORS.declined);
  });

  it("reports unreadable responses, error objects and missing choices", () => {
    expect(extractContent("<html>oops</html>")).toEqual({ content: null, error: REPLY_ERRORS.unreadable });
    expect(extractContent("[1,2]")).toEqual({ content: null, error: REPLY_ERRORS.unreadable });
    expect(extractContent(JSON.stringify({ error: { code: 502, message: "x" } })).error).toBe(REPLY_ERRORS.providerError);
    expect(extractContent(JSON.stringify({ choices: [] })).error).toBe(REPLY_ERRORS.noChoice);
    expect(extractContent(JSON.stringify({ id: "x" })).error).toBe(REPLY_ERRORS.noChoice);
  });
});

describe("stripFences", () => {
  it("removes ```json fences, bare fences and an unclosed fence", () => {
    expect(stripFences('```json\n{"a":1}\n```')).toBe('{"a":1}');
    expect(stripFences('```\n{"a":1}\n```')).toBe('{"a":1}');
    expect(stripFences('  ```JSON\n{"a":1}```  ')).toBe('{"a":1}');
    expect(stripFences('```json\n{"a":1}')).toBe('{"a":1}');
    expect(stripFences('{"a":1}')).toBe('{"a":1}');
  });
});

describe("parseContent", () => {
  it("accepts a valid formula reply", () => {
    const content = JSON.stringify(FORMULA);
    const out = parseContent(content, "structure_only");
    expect(out).toEqual({ ok: true, reply: FORMULA, content });
  });

  it("accepts a reply inside code fences", () => {
    const content = "```json\n" + JSON.stringify(FORMULA, null, 2) + "\n```";
    const out = parseContent(content, "structure_only");
    expect(out.ok).toBe(true);
    if (out.ok) {
      expect(out.reply).toEqual(FORMULA);
      // The content kept for history is the reply as the model wrote it.
      expect(out.content).toBe(content);
    }
  });

  it("falls back to the text from the first { to the last } when JSON.parse fails", () => {
    const content = `Here is the formula you asked for:\n${JSON.stringify(FORMULA)}\nLet me know if it works.`;
    const out = parseContent(content, "substituted");
    expect(out.ok && out.reply).toEqual(FORMULA);
  });

  it("rejects text that isn't a JSON object", () => {
    for (const content of ["=SUM(A1:A9)", "[1,2,3]", "{not json}", "", "null"]) {
      const out = parseContent(content, "structure_only");
      expect(out).toEqual({ ok: false, error: REPLY_ERRORS.notJson, content });
    }
  });

  it("kind formula needs a formula starting with '='", () => {
    for (const formula of [null, "", "SUM(A1:A9)", "  "]) {
      const out = parseContent(JSON.stringify({ ...FORMULA, formula }), "structure_only");
      expect(out.ok).toBe(false);
      if (!out.ok) expect(out.error).toBe(REPLY_ERRORS.noFormula);
    }
    const missing: Record<string, unknown> = { ...FORMULA };
    delete missing.formula;
    expect(parseContent(JSON.stringify(missing), "structure_only").ok).toBe(false);
  });

  it("trims whitespace around the formula", () => {
    const out = parseContent(JSON.stringify({ ...FORMULA, formula: "  =SUM(F2:F501)\n" }), "structure_only");
    expect(out.ok && out.reply.formula).toBe("=SUM(F2:F501)");
  });

  it("structure-only rejects kind answer; substituted accepts it", () => {
    const answer = { kind: "answer", formula: null, placement: null, explanation: "PERSON_003 has the largest order.", assumptions: [] };
    const so = parseContent(JSON.stringify(answer), "structure_only");
    expect(so.ok).toBe(false);
    if (!so.ok) expect(so.error).toBe(REPLY_ERRORS.answerInStructureOnly);
    const sub = parseContent(JSON.stringify(answer), "substituted");
    expect(sub.ok && sub.reply).toEqual(answer);
  });

  it("accepts clarify in both modes", () => {
    const clarify = { kind: "clarify", formula: null, placement: null, explanation: "Which year?", assumptions: [] };
    for (const mode of ["structure_only", "substituted"] as const) {
      expect(parseContent(JSON.stringify(clarify), mode)).toMatchObject({ ok: true, reply: clarify });
    }
  });

  it("turns an invalid placement into null instead of rejecting the reply", () => {
    const bad: unknown[] = ["G2", { cell: "G2:G501", fill_down: true }, { cell: "Orders!G2" }, { cell: "hello" }, { cell: 7 }, { cell: "ZZZZ1" }, { cell: "A0" }, [], 3];
    for (const placement of bad) {
      const out = parseContent(JSON.stringify({ ...FORMULA, placement }), "structure_only");
      expect(out.ok).toBe(true);
      if (out.ok) expect(out.reply.placement).toBeNull();
    }
  });

  it("normalizes a valid placement cell and defaults fill_down to false", () => {
    const out = parseContent(JSON.stringify({ ...FORMULA, placement: { cell: " $g$2 " } }), "structure_only");
    expect(out.ok && out.reply.placement).toEqual({ cell: "G2", fill_down: false });
    const yes = parseContent(JSON.stringify({ ...FORMULA, placement: { cell: "H10", fill_down: true } }), "structure_only");
    expect(yes.ok && yes.reply.placement).toEqual({ cell: "H10", fill_down: true });
    const str = parseContent(JSON.stringify({ ...FORMULA, placement: { cell: "H10", fill_down: "true" } }), "structure_only");
    expect(str.ok && str.reply.placement).toEqual({ cell: "H10", fill_down: false });
  });

  it("defaults missing placement and assumptions, and wraps a single assumption string", () => {
    const content = JSON.stringify({ kind: "formula", formula: "=1+1", explanation: "Adds." });
    expect(parseContent(content, "structure_only")).toMatchObject({
      ok: true,
      reply: { kind: "formula", formula: "=1+1", placement: null, explanation: "Adds.", assumptions: [] },
    });
    const one = parseContent(JSON.stringify({ ...FORMULA, assumptions: "Amounts are in USD." }), "structure_only");
    expect(one.ok && one.reply.assumptions).toEqual(["Amounts are in USD."]);
  });

  it("rejects wrong shapes: unknown kind, missing explanation, non-string assumptions", () => {
    const bad = [
      { ...FORMULA, kind: "table" },
      { ...FORMULA, kind: undefined },
      { ...FORMULA, explanation: undefined },
      { ...FORMULA, explanation: 42 },
      { ...FORMULA, assumptions: [1, 2] },
      { ...FORMULA, formula: 12 },
    ];
    for (const reply of bad) {
      const out = parseContent(JSON.stringify(reply), "substituted");
      expect(out.ok).toBe(false);
      if (!out.ok) expect(out.error).toBe(REPLY_ERRORS.shape);
    }
  });

  it("drops extra keys and returns keys in ModelReply order", () => {
    const shuffled = { extra: "x", assumptions: FORMULA.assumptions, explanation: FORMULA.explanation, kind: "formula", formula: FORMULA.formula, confidence: 0.9 };
    const out = parseContent(JSON.stringify(shuffled), "structure_only");
    expect(out.ok).toBe(true);
    if (out.ok) expect(Object.keys(out.reply)).toEqual(["kind", "formula", "placement", "explanation", "assumptions"]);
  });

  it("error messages never quote the reply", () => {
    const secretish = "PERSON_001 lives at 12 Elm Street";
    const out = parseContent(JSON.stringify({ kind: "nope", explanation: secretish }), "structure_only");
    expect(out.ok).toBe(false);
    if (!out.ok) expect(out.error).not.toContain("Elm");
  });

  it("the zod schema is exported for reuse", () => {
    expect(ModelReplySchema.safeParse(FORMULA).success).toBe(true);
  });
});

describe("parseReply", () => {
  it("parses a full chat-completions response", () => {
    const content = "```json\n" + JSON.stringify(FORMULA) + "\n```";
    const out = parseReply(response(content), "substituted");
    expect(out).toEqual({ ok: true, reply: FORMULA, content });
  });

  it("returns the extraction error with null content", () => {
    expect(parseReply(response(null, "length"), "structure_only")).toEqual({
      ok: false,
      error: REPLY_ERRORS.lengthCap,
      content: null,
    });
    expect(parseReply("not json", "structure_only")).toEqual({ ok: false, error: REPLY_ERRORS.unreadable, content: null });
  });

  it("says the length limit was reached when a cut-off reply can't be parsed", () => {
    const cut = '{"kind":"formula","formula":"=SUMIFS(F2:F501,';
    expect(parseReply(response(cut, "length"), "structure_only")).toEqual({
      ok: false,
      error: REPLY_ERRORS.lengthCap,
      content: cut,
    });
  });

  it("returns the parse error with the content, for the correction retry", () => {
    const out = parseReply(response("Sure! Use =SUM(F2:F501)."), "structure_only");
    expect(out).toEqual({ ok: false, error: REPLY_ERRORS.notJson, content: "Sure! Use =SUM(F2:F501)." });
  });
});
