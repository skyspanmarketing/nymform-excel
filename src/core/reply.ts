// Reply parsing (spec §7.9). Model output is untrusted: this only checks its shape. Formulas still
// go through formulaGate.ts before and after restoration (invariant 4).
import { z } from "zod";
import { parseCell } from "./a1";
import type { ModelReply, ProdMode } from "./types";

export type ParseResult =
  | { ok: true; reply: ModelReply; content: string }
  | { ok: false; error: string; content: string | null };

export const REPLY_ERRORS = {
  unreadable: "The model provider's reply couldn't be read.",
  providerError: "The model provider returned an error instead of a reply.",
  noChoice: "The model provider's reply had no answer in it.",
  lengthCap: "The model reached its length limit before it replied. Try again, or ask a simpler question.",
  filtered: "The model provider withheld the reply. Try rephrasing.",
  declined: "The model declined to answer. Try rephrasing.",
  empty: "The model sent an empty reply. Try again.",
  notJson: "The model's reply isn't a JSON object.",
  shape: "The model's reply doesn't follow the expected format.",
  noFormula: "The model's reply is a formula without a formula that starts with \"=\".",
  answerInStructureOnly: "The model answered from data it wasn't sent. Structure-only requests can only get a formula or a question back.",
} as const;

/** A usable placement, or null. A bad cell is not a reason to reject the reply. */
function toPlacement(v: unknown): ModelReply["placement"] {
  if (v === null || typeof v !== "object" || Array.isArray(v)) return null;
  const p = v as Record<string, unknown>;
  if (typeof p.cell !== "string") return null;
  const cell = p.cell.trim().replace(/\$/g, "").toUpperCase();
  if (!parseCell(cell)) return null;
  return { cell, fill_down: p.fill_down === true };
}

function toAssumptions(v: string[] | string | null | undefined): string[] {
  if (v === null || v === undefined) return [];
  if (typeof v === "string") return v.trim() === "" ? [] : [v];
  return v;
}

/** The ModelReply contract (spec §6). Extra keys are dropped; missing optional parts get defaults. */
export const ModelReplySchema = z
  .object({
    kind: z.enum(["formula", "answer", "clarify"]),
    formula: z
      .string()
      .nullable()
      .optional()
      .transform((f) => (typeof f === "string" && f.trim() !== "" ? f.trim() : null)),
    placement: z.unknown().optional().transform(toPlacement),
    explanation: z.string(),
    assumptions: z.union([z.array(z.string()), z.string(), z.null()]).optional().transform(toAssumptions),
  })
  .transform(
    (r): ModelReply => ({
      kind: r.kind,
      formula: r.formula,
      placement: r.placement,
      explanation: r.explanation,
      assumptions: r.assumptions,
    }),
  );

/** Pulls choices[0].message.content out of the raw chat-completions response text. */
export function extractContent(raw: string): { content: string | null; finishReason?: string; error?: string } {
  let data: unknown;
  try {
    data = JSON.parse(raw);
  } catch {
    return { content: null, error: REPLY_ERRORS.unreadable };
  }
  if (!isObject(data)) return { content: null, error: REPLY_ERRORS.unreadable };
  if (isObject(data.error)) return { content: null, error: REPLY_ERRORS.providerError };

  const choice = Array.isArray(data.choices) ? (data.choices[0] as unknown) : undefined;
  if (!isObject(choice)) return { content: null, error: REPLY_ERRORS.noChoice };
  const finishReason = typeof choice.finish_reason === "string" ? choice.finish_reason : undefined;
  const message = isObject(choice.message) ? choice.message : {};
  const content = textOf(message.content);

  if (content === null || content.trim() === "") {
    const out: { content: null; finishReason?: string; error: string } = { content: null, error: REPLY_ERRORS.empty };
    if (finishReason !== undefined) out.finishReason = finishReason;
    if (finishReason === "length") out.error = REPLY_ERRORS.lengthCap;
    else if (finishReason === "content_filter") out.error = REPLY_ERRORS.filtered;
    else if (typeof message.refusal === "string" && message.refusal.trim() !== "") out.error = REPLY_ERRORS.declined;
    return out;
  }
  return finishReason === undefined ? { content } : { content, finishReason };
}

/** Strips code fences, parses JSON, validates with zod. Structure-only mode rejects kind "answer". */
export function parseContent(content: string, mode: ProdMode): ParseResult {
  const fail = (error: string): ParseResult => ({ ok: false, error, content });
  if (typeof content !== "string") return { ok: false, error: REPLY_ERRORS.empty, content: null };

  const json = parseObject(stripFences(content));
  if (json === undefined) return fail(REPLY_ERRORS.notJson);
  const parsed = ModelReplySchema.safeParse(json);
  if (!parsed.success) return fail(REPLY_ERRORS.shape);
  const reply = parsed.data;

  if (reply.kind === "formula" && !(reply.formula ?? "").startsWith("=")) return fail(REPLY_ERRORS.noFormula);
  if (reply.kind === "answer" && mode === "structure_only") return fail(REPLY_ERRORS.answerInStructureOnly);
  return { ok: true, reply, content };
}

/** extractContent + parseContent on the raw response text. */
export function parseReply(raw: string, mode: ProdMode): ParseResult {
  const extracted = extractContent(raw);
  if (extracted.content === null) return { ok: false, error: extracted.error ?? REPLY_ERRORS.empty, content: null };
  const parsed = parseContent(extracted.content, mode);
  // A reply cut off by the token cap usually isn't valid JSON; say why.
  if (!parsed.ok && extracted.finishReason === "length") return { ...parsed, error: REPLY_ERRORS.lengthCap };
  return parsed;
}

// ---------------------------------------------------------------------------------------------
// Helpers

/** Removes a surrounding Markdown code fence (```json ... ```), if there is one. */
export function stripFences(text: string): string {
  const t = text.trim();
  const closed = /^```[^\n`]*\n?([\s\S]*?)\n?[ \t]*```$/.exec(t);
  if (closed) return (closed[1] ?? "").trim();
  const open = /^```[^\n`]*\n([\s\S]*)$/.exec(t);
  if (open) return (open[1] ?? "").trim();
  return t;
}

/** JSON.parse, then the text from the first "{" to the last "}". Only objects count. */
function parseObject(text: string): Record<string, unknown> | undefined {
  const direct = tryParse(text);
  if (isObject(direct)) return direct;
  const first = text.indexOf("{");
  const last = text.lastIndexOf("}");
  if (first < 0 || last <= first) return undefined;
  const inner = tryParse(text.slice(first, last + 1));
  return isObject(inner) ? inner : undefined;
}

function tryParse(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

/** Message content as text: a string, or the text parts of a content array. */
function textOf(content: unknown): string | null {
  if (typeof content === "string") return content;
  if (!Array.isArray(content)) return null;
  const parts = content
    .map((p: unknown) => (isObject(p) && typeof p.text === "string" ? p.text : typeof p === "string" ? p : ""))
    .filter((s) => s !== "");
  return parts.length === 0 ? null : parts.join("");
}

function isObject(v: unknown): v is Record<string, unknown> {
  return v !== null && typeof v === "object" && !Array.isArray(v);
}
