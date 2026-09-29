// Parts of private values typed in a question: "Felix" of "Felix Bianchi", the "0199" of a phone
// number, the part of an email address before the @. Substitution replaces only whole values, and
// the auditor blocks a request that still holds a part of one. Here the pane finds which private
// values a part belongs to, so the user can say which one they meant; the question then carries
// that value's stand-in instead.
//
// Nothing here is sent. The candidates (real values and where they are) stay in the pane and are
// never logged. The question that goes out holds stand-ins only, and it is audited as usual.
import { TOKEN_RE, aliasRevealsHeader, emailLocalPart, nameCharAt, nameCharBefore, privateCellsOf, withoutMarks, withoutSharpS, type StandInMap } from "./transform";
import type { RangeSpec, TokenKind } from "./types";

/** Parts shorter than this aren't looked for (the auditor skips them too). */
export const MIN_PART = 4;
/** At most this many candidates are listed per part; the rest are counted in `more`. */
export const MAX_CANDIDATES = 30;
/** At most this many cells are kept per candidate and place, for showing them in the sheet. */
export const MAX_CELLS = 150;

/** Where a candidate value is: its sheet and the cells holding it (local addresses, "D2"). */
export interface PartialPlace {
  sheet: string;
  column: string;
  header: string;
  cells: string[];
}

export interface PartialCandidate {
  /** The private value as the cell shows it. Pane only. */
  value: string;
  /** The stand-in it goes out as. */
  token: string;
  kind: TokenKind;
  /** How many cells hold it. */
  count: number;
  places: PartialPlace[];
}

export interface PartialMatch {
  /** The part as typed. */
  fragment: string;
  /** Its folded form: the key choices are made under. */
  key: string;
  /** Where it is in the question, as [start, end) positions. */
  spans: [number, number][];
  /** Most frequent first, then alphabetical. */
  candidates: PartialCandidate[];
  /** Candidates not listed (beyond MAX_CANDIDATES). */
  more: number;
  /** One candidate that names a person, an email or an ID: the part can stand for it without asking. */
  autoMap: boolean;
}

/** Case, accents and ß folded, as the auditor folds them. */
export function foldKey(s: string): string {
  return withoutSharpS(withoutMarks(s.normalize("NFC"))).toLowerCase();
}

// A word of the question: letters, marks and digits, with . _ - or an apostrophe between them
// ("maria.lopez", "O'Brien", "Anne-Marie", "sunflower77").
const WORD_RE = /[\p{L}\p{M}\p{N}]+(?:[._\-'’][\p{L}\p{M}\p{N}]+)*/gu;
const INNER_RE = /[._\-'’]/u;
const PLAIN_NUMBER_RE = /^[-+]?\d+(?:[.,]\d+)?$/u;

function yearLike(digits: string): boolean {
  return digits.length === 4 && Number(digits) >= 1900 && Number(digits) <= 2100;
}

/** The parts of a value a question can hold on their own, folded. */
function partsOf(value: string): Set<string> {
  const out = new Set<string>();
  const folded = foldKey(value);
  const local = emailLocalPart(value);
  if (local !== null) {
    // An email address: its part before the @ only (the domain is shared by many).
    out.add(foldKey(local));
    return out;
  }
  const words = folded.split(/\s+/u).map((w) => w.replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$/gu, "")).filter(Boolean);
  if (words.length > 1) {
    // Words as typed ("o'brien") and their pieces ("brien"), as the auditor splits them.
    for (const w of words) {
      if (w.length >= MIN_PART && !/^\d+$/u.test(w)) out.add(w);
      for (const piece of w.split(/[^\p{L}\p{M}\p{N}]+/u)) if (piece.length >= MIN_PART && !/^\d+$/u.test(piece)) out.add(piece);
    }
  }
  // The digit groups of a value that isn't a number: a phone's "0199", an ID's "6789".
  if (!PLAIN_NUMBER_RE.test(value.trim())) {
    for (const g of folded.split(/\D+/u)) if (g.length >= MIN_PART && g.length <= 6 && !yearLike(g)) out.add(g);
  }
  out.delete(folded);
  return out;
}

interface Found {
  fragment: string;
  key: string;
  spans: [number, number][];
}

/**
 * Where `text` names a column by the header it is sent under ("Unit price", or the alias of a
 * hidden header), as whole words in any case. A word there is the column's name, not a part of a
 * value: "price" in "Average Unit price" isn't the customer Jordan Price.
 */
function headerSpans(text: string, specs: readonly RangeSpec[]): [number, number][] {
  const names = new Set<string>();
  for (const spec of specs) {
    const policies = new Map(spec.policies.map((p) => [p.letter, p]));
    spec.columns.forEach((info, i) => {
      const policy = policies.get(info.letter) ?? spec.policies[i];
      const alias = typeof policy?.alias === "string" && policy.alias.trim() !== "" ? policy.alias.trim() : null;
      // A hidden header was replaced by its alias in the question; one the alias shows was not.
      if (alias === null || aliasRevealsHeader(alias, info.header)) names.add(info.header.trim());
      if (alias !== null) names.add(alias);
    });
  }
  const spans: [number, number][] = [];
  for (const name of names) {
    // A name longer than the question can't be in it (a first-row cell can hold 32,767 characters).
    if (name.length < MIN_PART || name.replace(/\s+/gu, " ").length > text.length) continue;
    // The name as a plain pattern, any case; whole words are checked here rather than with
    // lookarounds, which cost about a millisecond per column to compile, again after every GC.
    try {
      const re = new RegExp(name.split(/\s+/u).map(escapeRegExp).join("\\s+"), "giu");
      let rejected = -1;
      for (let m = re.exec(text); m !== null; m = re.exec(text)) {
        const start = m.index;
        // The engine went back to a match already turned down: stop rather than loop.
        if (start <= rejected) break;
        const end = start + m[0].length;
        if (!nameCharBefore(text, start) && !nameCharAt(text, end)) spans.push([start, end]);
        else {
          // Not a whole word: look again from the next character, as a lookaround would have.
          // With the u flag the engine moves a lastIndex inside a surrogate pair back to its
          // start, so the step is a whole code point.
          rejected = start;
          re.lastIndex = start + ((text.codePointAt(start) ?? 0) > 0xffff ? 2 : 1);
        }
      }
    } catch {
      // A pattern the engine can't build (it runs out of stack on some very long names): that
      // name isn't masked, so its words can be read as parts of values again. The check decides.
    }
  }
  return spans;
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\/]/gu, "\\$&");
}

/** The words of `text` (outside stand-ins and column names) that could be parts of a value, by folded key. */
function wordsOf(text: string, names: readonly [number, number][]): Map<string, Found> {
  const tokens: [number, number][] = [...names];
  for (const m of text.matchAll(new RegExp(TOKEN_RE.source, "g"))) tokens.push([m.index, m.index + m[0].length]);
  const inToken = (a: number, b: number) => tokens.some(([s, e]) => a < e && b > s);
  const out = new Map<string, Found>();
  const add = (fragment: string, start: number) => {
    const end = start + fragment.length;
    if (fragment.length < MIN_PART || inToken(start, end)) return;
    const key = foldKey(fragment);
    if (/^\d+$/u.test(key) && (key.length > 6 || yearLike(key))) return;
    const f = out.get(key) ?? { fragment, key, spans: [] };
    if (!f.spans.some(([s]) => s === start)) f.spans.push([start, end]);
    out.set(key, f);
  };
  for (const m of text.matchAll(WORD_RE)) {
    add(m[0], m.index);
    // "O'Brien" as a whole and as its pieces ("Brien"); a piece inside a whole that matched is
    // dropped in findPartialMatches.
    if (INNER_RE.test(m[0])) {
      let at = m.index;
      for (const piece of m[0].split(INNER_RE)) {
        const start = text.indexOf(piece, at);
        at = start + piece.length;
        add(piece, start);
      }
    }
  }
  return out;
}

/**
 * The parts of private values in `text` (a question after substitution), each with the values it
 * could stand for. A part that is also a whole value doesn't count: substitution replaced it.
 */
export function findPartialMatches(text: string, specs: readonly RangeSpec[], map: StandInMap): PartialMatch[] {
  if (typeof text !== "string" || text === "") return [];
  const words = wordsOf(text, headerSpans(text, specs));
  if (words.size === 0) return [];

  interface Acc {
    value: string;
    kind: () => TokenKind;
    token: () => string;
    count: number;
    places: Map<string, PartialPlace>;
  }
  // folded part -> value -> what we know of it
  const byPart = new Map<string, Map<string, Acc>>();
  const partsCache = new Map<string, Set<string>>();
  for (const spec of specs) {
    const d = spec.data;
    for (const cell of privateCellsOf(spec, map)) {
      let parts = partsCache.get(cell.value);
      if (!parts) {
        parts = partsOf(cell.value);
        partsCache.set(cell.value, parts);
      }
      for (const part of parts) {
        if (!words.has(part)) continue;
        const values = byPart.get(part) ?? new Map<string, Acc>();
        byPart.set(part, values);
        const acc = values.get(cell.value) ?? { value: cell.value, kind: cell.kind, token: cell.token, count: 0, places: new Map() };
        values.set(cell.value, acc);
        acc.count++;
        const placeKey = `${d.sheet}\u0000${cell.letter}`;
        const place = acc.places.get(placeKey) ?? { sheet: d.sheet, column: cell.letter, header: cell.header, cells: [] };
        acc.places.set(placeKey, place);
        if (place.cells.length < MAX_CELLS) place.cells.push(`${cell.letter}${d.rowIndex + cell.row + 1}`);
      }
    }
  }

  const out: PartialMatch[] = [];
  for (const [key, found] of words) {
    const values = byPart.get(key);
    if (!values || values.size === 0) continue;
    const all = [...values.values()].sort((a, b) => b.count - a.count || a.value.localeCompare(b.value));
    const listed = all.slice(0, MAX_CANDIDATES).map((a) => ({
      value: a.value,
      token: a.token(),
      kind: a.kind(),
      count: a.count,
      places: [...a.places.values()],
    }));
    const only = listed.length === 1 && all.length === 1 ? listed[0]! : null;
    out.push({
      fragment: found.fragment,
      key,
      spans: found.spans.sort((a, b) => a[0] - b[0]),
      candidates: listed,
      more: all.length - listed.length,
      autoMap: only !== null && only.kind !== "TEXT",
    });
  }
  // A piece inside a longer part that matched ("Brien" inside "O'Brien") is the same question.
  const inside = (m: PartialMatch) =>
    m.spans.every(([s, e]) => out.some((o) => o !== m && o.fragment.length > m.fragment.length && o.spans.some(([os, oe]) => os <= s && e <= oe)));
  // In the order the parts appear in the question.
  return out.filter((m) => !inside(m)).sort((a, b) => a.spans[0]![0] - b.spans[0]![0]);
}

/**
 * `text` with each chosen part replaced by the stand-ins chosen for it, joined with " or ". A part
 * chosen with no stand-ins is left as typed (the auditor then blocks it). Only stand-ins listed
 * as candidates for that part are used; anything else leaves the part as typed.
 */
export function applyChoices(text: string, matches: readonly PartialMatch[], chosen: ReadonlyMap<string, readonly string[]>): string {
  const edits: { start: number; end: number; replacement: string }[] = [];
  for (const m of matches) {
    const tokens = chosen.get(m.key);
    if (!tokens || tokens.length === 0) continue;
    const allowed = new Set(m.candidates.map((c) => c.token));
    if (!tokens.every((t) => allowed.has(t))) continue;
    const replacement = [...new Set(tokens)].join(" or ");
    for (const [start, end] of m.spans) edits.push({ start, end, replacement });
  }
  // Right to left, skipping any span that overlaps one already replaced.
  edits.sort((a, b) => b.start - a.start);
  let out = text;
  let limit = Infinity;
  for (const e of edits) {
    if (e.end > limit) continue;
    out = out.slice(0, e.start) + e.replacement + out.slice(e.end);
    limit = e.start;
  }
  return out;
}
