// Restoration (spec §7.12). Puts real values back in place of stand-ins, for the formula that is
// inserted and for display. Neither function ever touches history (invariant 5).
import { tokenize, type FToken } from "./formulaGate";
import { TOKEN_RE } from "./transform";

export interface TokenLookup {
  valueOf(token: string): string | undefined;
  /**
   * The map's token prefix, when known ("" or "NYMFORM_"). With "NYMFORM_", an unprefixed look-alike
   * such as PERSON_050 is ordinary text only if isReserved confirms it came from the sheet,
   * unless its prefixed form is a stand-in: then it is reported as a possible dropped prefix.
   */
  readonly prefix?: string;
  /** Token-shaped text observed in the workbook; absence fails closed. */
  isReserved?(token: string): boolean;
}

function tokenPattern(): RegExp {
  return new RegExp(TOKEN_RE.source, "g");
}

/**
 * Replaces every token in `text` with `encode(value)`, collecting tokens the map doesn't know and,
 * when `restored` is given, the tokens it replaced.
 */
function replaceTokens(
  text: string,
  map: TokenLookup,
  encode: (value: string) => string,
  unknown: Set<string>,
  restored?: Set<string>,
): string {
  return text.replace(tokenPattern(), (token) => {
    const value = map.valueOf(token);
    if (value === undefined) {
      const sheetText = map.prefix === "NYMFORM_" && !token.startsWith("NYMFORM_") && map.isReserved?.(token) === true;
      if (!sheetText || map.valueOf(`NYMFORM_${token}`) !== undefined) unknown.add(token);
      return token;
    }
    restored?.add(token);
    return encode(value);
  });
}

/**
 * Replaces tokens only inside string literals, doubling quotes. Nothing else changes.
 * String literals are found with the gate's tokenizer; a formula it can't read comes back unchanged
 * (the gate rejects it as G-PARSE anyway). `unknownTokens` are token-shaped strings in the literals
 * that the map doesn't know: the formula would compare against them as text, so the flow refuses to
 * insert it. `restoredTokens` are the ones replaced.
 */
export function restoreFormula(
  formula: string,
  map: TokenLookup,
): { formula: string; unknownTokens: string[]; restoredTokens: string[] } {
  let tokens: FToken[];
  try {
    tokens = tokenize(formula);
  } catch {
    return { formula, unknownTokens: [], restoredTokens: [] };
  }
  const unknown = new Set<string>();
  const restored = new Set<string>();
  let out = "";
  let last = 0;
  for (const t of tokens) {
    if (t.type !== "string") continue;
    // The raw content between the quotes: quotes in it are already doubled, and tokens never contain one.
    const inner = formula.slice(t.start + 1, t.end - 1);
    out += formula.slice(last, t.start + 1) + replaceTokens(inner, map, (v) => v.replace(/"/g, '""'), unknown, restored);
    last = t.end - 1;
  }
  out += formula.slice(last);
  return { formula: out, unknownTokens: [...unknown], restoredTokens: [...restored] };
}

/** Replaces tokens anywhere, for display only. Never goes into history. */
export function restoreText(text: string, map: TokenLookup): { text: string; unknownTokens: string[] } {
  const unknown = new Set<string>();
  const out = replaceTokens(text, map, (v) => v, unknown);
  return { text: out, unknownTokens: [...unknown] };
}
