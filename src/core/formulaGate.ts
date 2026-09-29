// Formula gate (spec §7.10). Model output is untrusted: every formula passes check() before and
// after restore.ts fills in string literals. Fails closed: anything the tokenizer or parser can't
// read is G-PARSE, and any unexpected error inside check() is G-PARSE too.
import { MAX_COLUMNS, MAX_ROWS, columnIndex, parseCell, splitSheetAddress } from "./a1";
import { ALLOWLIST, EXCLUDED_REASONS, normalizeFunctionName } from "./allowlist";
import type { GateCode, GateResult, SheetContext, TableInfo } from "./types";

export type GateContext = Pick<SheetContext, "sheet" | "sheets" | "tableNames" | "tables" | "allowedRanges" | "definedNames">;

export interface FToken {
  /** e.g. "string", "number", "bool", "error", "ref", "name", "func", "op", "paren", "sep", "array", "struct", "ws" */
  type: string;
  /** Raw source text of the token. */
  text: string;
  /** Decoded content for string literals (doubled quotes undone). */
  value?: string;
  /** Offsets into the formula, end exclusive. For strings, the span includes the quotes. */
  start: number;
  end: number;
}

export class FormulaParseError extends Error {
  constructor(message = "") {
    super(message);
    this.name = "FormulaParseError";
  }
}

export const MAX_FORMULA_LENGTH = 2000;
/** Nesting limit for parentheses, calls and operators. Excel itself stops at 64 nested functions. */
const MAX_DEPTH = 100;

// ---------------------------------------------------------------------------------------------
// Tokenizer

interface SheetPrefix {
  /** As written, including quotes and the "!". */
  raw: string;
  /** Decoded sheet part (quotes and '' escapes undone), without the "!". */
  sheet: string;
  kind: "sheet" | "workbook" | "3d";
}

interface LexToken extends FToken {
  /** ref, name and func tokens written with a sheet or workbook prefix. */
  prefix?: SheetPrefix;
  /** ref, name and func tokens: the text after the prefix, without a trailing spill "#". */
  body?: string;
  /** ref tokens followed by the spill operator, e.g. A2#. */
  spill?: boolean;
  /** struct tokens: the table name, or null for a reference with no table (e.g. [@Amount]). */
  table?: string | null;
}

const WS_RE = /[ \t\r\n]/;
const DIGIT_RE = /[0-9]/;
const WORD_START_RE = /[\p{L}_$]/u;
const WORD_RE = /[\p{L}\p{M}\p{N}_.$]+/uy;
const WORD_CHAR_RE = /[\p{L}\p{M}\p{N}_.$]/u;
const NAME_SHAPE_RE = /^[\p{L}_][\p{L}\p{M}\p{N}_.]*$/u;
const NUMBER_RE = /(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/y;
const ROWS_RE = /\$?\d+:\$?\d+/y;
const ERROR_RE = /#(?:N\/A|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)/iy;
// Unquoted prefix: optional [workbook], optional sheet or Sheet1:Sheet3, then "!".
const SHEET_PREFIX_RE =
  /(?:\[[^[\]'"!]*\])?(?:[\p{L}_][\p{L}\p{M}\p{N}_.]*(?::[\p{L}_][\p{L}\p{M}\p{N}_.]*)?)?!/uy;
const SINGLE_OPS = "+-*/^&=<>%:@";

const PARENS = "The parentheses don't match.";
const BRACES = "The curly braces don't match.";
const BRACKETS = "The square brackets don't match.";
const BAD_TABLE_REF = "A table reference has characters it can't read.";

function at(i: number): string {
  return `at position ${i + 1}`;
}

function execAt(re: RegExp, src: string, i: number): string | null {
  re.lastIndex = i;
  const m = re.exec(src);
  return m ? m[0] : null;
}

function skipWs(src: string, i: number): number {
  let j = i;
  while (j < src.length && WS_RE.test(src[j]!)) j++;
  return j;
}

function classifyPrefix(sheet: string, quoted: boolean): SheetPrefix["kind"] {
  if (/[[\]\\/]/.test(sheet)) return "workbook";
  if (sheet.includes(":")) return "3d";
  return quoted || sheet.length > 0 ? "sheet" : "workbook";
}

type AreaShape = { r1: number; c1: number; r2: number; c2: number; kind: "cells" | "cols" | "rows" };

/** "A1", "$A$1:$B$9", "A:C" or "2:5" (no sheet) to 0-based coordinates, or null. */
function parseArea(address: string): AreaShape | null {
  const parts = address.split(":");
  if (parts.length === 1) {
    const c = parseCell(parts[0]!);
    return c ? { r1: c.r, c1: c.c, r2: c.r, c2: c.c, kind: "cells" } : null;
  }
  if (parts.length !== 2) return null;
  const [a, b] = parts as [string, string];
  const ca = parseCell(a);
  const cb = parseCell(b);
  if (ca && cb) {
    return {
      r1: Math.min(ca.r, cb.r),
      c1: Math.min(ca.c, cb.c),
      r2: Math.max(ca.r, cb.r),
      c2: Math.max(ca.c, cb.c),
      kind: "cells",
    };
  }
  const colA = /^\$?([A-Za-z]{1,3})$/.exec(a);
  const colB = /^\$?([A-Za-z]{1,3})$/.exec(b);
  if (colA && colB) {
    const x = columnIndex(colA[1]!);
    const y = columnIndex(colB[1]!);
    if (x < 0 || y < 0) return null;
    return { r1: 0, c1: Math.min(x, y), r2: MAX_ROWS - 1, c2: Math.max(x, y), kind: "cols" };
  }
  const rowA = /^\$?(\d{1,7})$/.exec(a);
  const rowB = /^\$?(\d{1,7})$/.exec(b);
  if (rowA && rowB) {
    const x = Number(rowA[1]) - 1;
    const y = Number(rowB[1]) - 1;
    if (x < 0 || y < 0 || x >= MAX_ROWS || y >= MAX_ROWS) return null;
    return { r1: Math.min(x, y), c1: 0, r2: Math.max(x, y), c2: MAX_COLUMNS - 1, kind: "rows" };
  }
  return null;
}

const SPECIAL_ITEMS = new Set(["#ALL", "#DATA", "#HEADERS", "#TOTALS", "#THIS ROW"]);

function validStructItem(item: string): boolean {
  if (SPECIAL_ITEMS.has(item.trim().toUpperCase())) return true;
  if (item.length === 0) return false;
  for (let k = 0; k < item.length; k++) {
    const c = item[k]!;
    if (c === "'") k++;
    else if (c === "[" || c === "]" || c === "#") return false;
  }
  return true;
}

/** The part between a table reference's outer brackets: "", "@", "Amount", "@Amount", "#All", "[#This Row],[Amount]", "@[Unit Price]", "[A]:[B]". */
function validStructInner(inner: string): boolean {
  let s = inner;
  if (s === "") return true;
  if (s.startsWith("@")) {
    s = s.slice(1);
    if (s === "") return true;
  }
  if (!s.startsWith("[")) return validStructItem(s);
  let k = 0;
  for (;;) {
    if (s[k] !== "[") return false;
    let e = k + 1;
    while (e < s.length && s[e] !== "]") e += s[e] === "'" ? 2 : 1;
    if (e >= s.length) return false;
    if (!validStructItem(s.slice(k + 1, e))) return false;
    k = e + 1;
    while (s[k] === " ") k++;
    if (k === s.length) return true;
    if (s[k] !== "," && s[k] !== ":") return false;
    k++;
    while (s[k] === " ") k++;
  }
}

/** Reads a bracket group starting at src[i] === "[" and returns the index after its closing "]". */
function readBrackets(src: string, i: number): number {
  let depth = 0;
  let j = i;
  while (j < src.length) {
    const c = src[j]!;
    if (c === "'") {
      const n = src[j + 1];
      if (n === undefined || !"[]#'".includes(n)) throw new FormulaParseError(BAD_TABLE_REF);
      j += 2;
      continue;
    }
    if (c === "[") {
      depth++;
      if (depth > 2) throw new FormulaParseError(BAD_TABLE_REF);
    } else if (c === "]") {
      depth--;
      if (depth === 0) {
        if (!validStructInner(src.slice(i + 1, j))) throw new FormulaParseError(BAD_TABLE_REF);
        return j + 1;
      }
    } else if (c === '"' || c < " ") {
      throw new FormulaParseError(BAD_TABLE_REF);
    }
    j++;
  }
  throw new FormulaParseError(BRACKETS);
}

class Lexer {
  private readonly out: LexToken[] = [];
  constructor(private readonly src: string) {}

  private push(type: string, start: number, end: number, extra: Partial<LexToken> = {}): number {
    this.out.push({ type, text: this.src.slice(start, end), start, end, ...extra });
    return end;
  }

  run(): LexToken[] {
    const src = this.src;
    let i = 0;
    let parens = 0;
    let braces = 0;
    while (i < src.length) {
      const ch = src[i]!;
      if (WS_RE.test(ch)) {
        i = this.push("ws", i, skipWs(src, i));
        continue;
      }
      switch (ch) {
        case '"':
          i = this.readString(i);
          continue;
        case "'":
          i = this.readQuotedPrefix(i);
          continue;
        case "(":
          parens++;
          i = this.push("paren", i, i + 1);
          continue;
        case ")":
          if (--parens < 0) throw new FormulaParseError(PARENS);
          i = this.push("paren", i, i + 1);
          continue;
        case "{":
          if (braces > 0) throw new FormulaParseError("An array constant can't contain another array.");
          braces++;
          i = this.push("array", i, i + 1);
          continue;
        case "}":
          if (--braces < 0) throw new FormulaParseError(BRACES);
          i = this.push("array", i, i + 1);
          continue;
        case ",":
          i = this.push("sep", i, i + 1);
          continue;
        case ";":
          if (braces === 0) throw new FormulaParseError("A semicolon can only separate rows in an array constant.");
          i = this.push("sep", i, i + 1);
          continue;
        case "#": {
          const m = execAt(ERROR_RE, src, i);
          if (!m) throw new FormulaParseError(`There's a # it can't read ${at(i)}.`);
          i = this.push("error", i, i + m.length);
          continue;
        }
        case "[":
          i = this.readBracketStart(i);
          continue;
      }
      const two = src.slice(i, i + 2);
      if (two === "<=" || two === ">=" || two === "<>") {
        i = this.push("op", i, i + 2);
        continue;
      }
      if (SINGLE_OPS.includes(ch)) {
        i = this.push("op", i, i + 1);
        continue;
      }
      if (DIGIT_RE.test(ch) || (ch === "." && DIGIT_RE.test(src[i + 1] ?? ""))) {
        i = this.readNumberOrRows(i);
        continue;
      }
      if (WORD_START_RE.test(ch)) {
        i = this.readWordStart(i);
        continue;
      }
      throw new FormulaParseError(`There's a character it can't read ${at(i)}.`);
    }
    if (parens !== 0) throw new FormulaParseError(PARENS);
    if (braces !== 0) throw new FormulaParseError(BRACES);
    return this.out;
  }

  private readString(i: number): number {
    const src = this.src;
    let j = i + 1;
    let value = "";
    for (;;) {
      if (j >= src.length) throw new FormulaParseError("A text value is missing its closing quote.");
      const c = src[j]!;
      if (c === '"') {
        if (src[j + 1] === '"') {
          value += '"';
          j += 2;
          continue;
        }
        j++;
        break;
      }
      value += c;
      j++;
    }
    return this.push("string", i, j, { value });
  }

  private readQuotedPrefix(i: number): number {
    const src = this.src;
    let j = i + 1;
    let sheet = "";
    for (;;) {
      if (j >= src.length) throw new FormulaParseError("A sheet name is missing its closing quote.");
      const c = src[j]!;
      if (c === "'") {
        if (src[j + 1] === "'") {
          sheet += "'";
          j += 2;
          continue;
        }
        j++;
        break;
      }
      sheet += c;
      j++;
    }
    if (src[j] !== "!") throw new FormulaParseError("A quoted sheet name isn't followed by an exclamation mark.");
    j++;
    const prefix: SheetPrefix = { raw: src.slice(i, j), sheet, kind: classifyPrefix(sheet, true) };
    return this.readRefBody(i, j, prefix);
  }

  private readBracketStart(i: number): number {
    const m = execAt(SHEET_PREFIX_RE, this.src, i);
    if (m && m.length > 1) {
      const prefix: SheetPrefix = { raw: m, sheet: m.slice(0, -1), kind: "workbook" };
      return this.readRefBody(i, i + m.length, prefix);
    }
    const end = readBrackets(this.src, i);
    return this.push("struct", i, end, { table: null });
  }

  private readNumberOrRows(i: number): number {
    const src = this.src;
    const rows = execAt(ROWS_RE, src, i);
    if (rows) {
      if (!parseArea(rows)) throw new FormulaParseError(`There's a row number outside the sheet ${at(i)}.`);
      const end = i + rows.length;
      if (WORD_CHAR_RE.test(src[end] ?? "")) throw new FormulaParseError(`There's a reference it can't read ${at(i)}.`);
      return this.push("ref", i, end, { body: rows });
    }
    const num = execAt(NUMBER_RE, src, i);
    if (!num) throw new FormulaParseError(`There's a number it can't read ${at(i)}.`);
    const end = i + num.length;
    if (WORD_CHAR_RE.test(src[end] ?? "")) throw new FormulaParseError(`There's a number joined to letters ${at(i)}.`);
    return this.push("number", i, end);
  }

  private readWordStart(i: number): number {
    const src = this.src;
    const m = execAt(SHEET_PREFIX_RE, src, i);
    if (m && m.length > 1) {
      const sheet = m.slice(0, -1);
      return this.readRefBody(i, i + m.length, { raw: m, sheet, kind: classifyPrefix(sheet, false) });
    }
    const word = execAt(WORD_RE, src, i)!;
    const end = i + word.length;
    if (src[skipWs(src, end)] === "(") return this.push("func", i, end, { body: word });
    if (src[end] === "[") {
      if (!NAME_SHAPE_RE.test(word)) throw new FormulaParseError(`There's a table name it can't read ${at(i)}.`);
      const close = readBrackets(src, end);
      return this.push("struct", i, close, { table: word });
    }
    return this.readRefOrName(i, i, undefined);
  }

  /** After a sheet prefix: a reference, a sheet-level name, or a function called through the prefix. */
  private readRefBody(start: number, bodyStart: number, prefix: SheetPrefix): number {
    const src = this.src;
    const word = execAt(WORD_RE, src, bodyStart);
    if (!word) throw new FormulaParseError("A sheet name must be followed by a cell reference.");
    const end = bodyStart + word.length;
    if (src[skipWs(src, end)] === "(") return this.push("func", start, end, { prefix, body: word });
    if (src[end] === "[") throw new FormulaParseError("A table reference can't follow a sheet name.");
    return this.readRefOrName(start, bodyStart, prefix);
  }

  private readRefOrName(start: number, bodyStart: number, prefix: SheetPrefix | undefined): number {
    const src = this.src;
    const w1 = execAt(WORD_RE, src, bodyStart)!;
    const end1 = bodyStart + w1.length;
    if (src[end1] === ":") {
      const w2 = execAt(WORD_RE, src, end1 + 1);
      if (w2) {
        const end2 = end1 + 1 + w2.length;
        const next = src[end2];
        const joined = `${w1}:${w2}`;
        if (next !== "!" && next !== "[" && next !== "#" && src[skipWs(src, end2)] !== "(" && parseArea(joined)) {
          return this.push("ref", start, end2, { prefix, body: joined });
        }
      }
    }
    const area = parseArea(w1);
    if (area) {
      if (src[end1] === "#") return this.push("ref", start, end1 + 1, { prefix, body: w1, spill: true });
      return this.push("ref", start, end1, { prefix, body: w1 });
    }
    if (!prefix && /^(?:TRUE|FALSE)$/i.test(w1)) return this.push("bool", start, end1);
    if (NAME_SHAPE_RE.test(w1)) return this.push("name", start, end1, { prefix, body: w1 });
    throw new FormulaParseError(`There's a word it can't read ${at(bodyStart)}.`);
  }
}

/** Tokenizes an Excel formula (en-US invariant syntax, as Office.js `formulas` uses). Throws FormulaParseError. */
export function tokenize(formula: string): FToken[] {
  return new Lexer(formula).run();
}

// ---------------------------------------------------------------------------------------------
// Parser

type Node =
  | { t: "lit"; tok: LexToken }
  | { t: "array"; tok: LexToken; items: LexToken[] }
  | { t: "ref"; tok: LexToken }
  | { t: "struct"; tok: LexToken }
  | { t: "name"; tok: LexToken }
  | { t: "call"; tok: LexToken; args: (Node | null)[] }
  | { t: "invoke"; callee: Node; args: (Node | null)[] }
  | { t: "paren"; items: Node[] }
  | { t: "unary"; op: string; arg: Node }
  | { t: "postfix"; op: string; arg: Node }
  | { t: "bin"; op: string; left: Node; right: Node };

const COMPARE_OPS = ["=", "<>", "<", ">", "<=", ">="];

class Parser {
  private readonly toks: LexToken[] = [];
  /** wsBefore[k]: whitespace sits right before toks[k]. */
  private readonly wsBefore: boolean[] = [];
  private pos = 0;
  private depth = 0;

  constructor(all: readonly LexToken[]) {
    let ws = false;
    for (const t of all) {
      if (t.type === "ws") {
        ws = true;
        continue;
      }
      this.toks.push(t);
      this.wsBefore.push(ws);
      ws = false;
    }
  }

  private peek(): LexToken | undefined {
    return this.toks[this.pos];
  }

  private take(): LexToken {
    const t = this.toks[this.pos];
    if (!t) throw new FormulaParseError("The formula ends too early.");
    this.pos++;
    return t;
  }

  private isOp(t: LexToken | undefined, ops: readonly string[]): boolean {
    return t !== undefined && t.type === "op" && ops.includes(t.text);
  }

  private is(t: LexToken | undefined, type: string, text: string): boolean {
    return t !== undefined && t.type === type && t.text === text;
  }

  private unexpected(t: LexToken | undefined): FormulaParseError {
    if (!t) return new FormulaParseError("The formula ends too early.");
    return new FormulaParseError(`Something is out of place ${at(t.start)}.`);
  }

  parseFormula(): Node {
    const first = this.toks[0];
    if (!first || !this.is(first, "op", "=") || first.start !== 0) throw this.unexpected(first);
    this.pos = 1;
    if (!this.peek()) throw new FormulaParseError("There's nothing after the equals sign.");
    const node = this.parseExpr();
    if (this.peek()) throw this.unexpected(this.peek());
    return node;
  }

  private parseExpr(): Node {
    if (++this.depth > MAX_DEPTH) throw new FormulaParseError("It's nested too deeply.");
    try {
      return this.parseBinary(0);
    } finally {
      this.depth--;
    }
  }

  // Levels, loosest first: comparison, &, + -, * /, ^.
  private static readonly LEVELS: readonly (readonly string[])[] = [COMPARE_OPS, ["&"], ["+", "-"], ["*", "/"], ["^"]];

  private parseBinary(level: number): Node {
    const ops = Parser.LEVELS[level];
    if (!ops) return this.parsePercent();
    let left = this.parseBinary(level + 1);
    while (this.isOp(this.peek(), ops)) {
      const op = this.take().text;
      const right = this.parseBinary(level + 1);
      left = { t: "bin", op, left, right };
    }
    return left;
  }

  private parsePercent(): Node {
    let node = this.parseUnary();
    while (this.isOp(this.peek(), ["%"])) {
      this.take();
      node = { t: "postfix", op: "%", arg: node };
    }
    return node;
  }

  private parseUnary(): Node {
    const ops: string[] = [];
    while (this.isOp(this.peek(), ["-", "+", "@"])) ops.push(this.take().text);
    let node = this.parseIntersection();
    for (let k = ops.length - 1; k >= 0; k--) node = { t: "unary", op: ops[k]!, arg: node };
    return node;
  }

  private startsReference(t: LexToken | undefined): boolean {
    if (!t) return false;
    return t.type === "ref" || t.type === "struct" || t.type === "name" || t.type === "func" || this.is(t, "paren", "(");
  }

  private parseIntersection(): Node {
    let left = this.parseRange();
    while (this.wsBefore[this.pos] && this.startsReference(this.peek())) {
      const right = this.parseRange();
      left = { t: "bin", op: " ", left, right };
    }
    return left;
  }

  private parseRange(): Node {
    let left = this.parsePostfix();
    while (this.isOp(this.peek(), [":"]) && !this.wsBefore[this.pos]) {
      this.take();
      if (this.wsBefore[this.pos]) throw this.unexpected(this.peek());
      const right = this.parsePostfix();
      left = { t: "bin", op: ":", left, right };
    }
    return left;
  }

  private parsePostfix(): Node {
    let node = this.parsePrimary();
    while (
      this.is(this.peek(), "paren", "(") &&
      !this.wsBefore[this.pos] &&
      (node.t === "call" || node.t === "invoke" || node.t === "paren")
    ) {
      this.take();
      node = { t: "invoke", callee: node, args: this.parseArgs() };
    }
    return node;
  }

  private parsePrimary(): Node {
    const t = this.peek();
    if (!t) throw this.unexpected(t);
    switch (t.type) {
      case "number":
      case "string":
      case "bool":
      case "error":
        this.take();
        return { t: "lit", tok: t };
      case "ref":
        this.take();
        return { t: "ref", tok: t };
      case "struct":
        this.take();
        return { t: "struct", tok: t };
      case "name":
        this.take();
        return { t: "name", tok: t };
      case "func": {
        this.take();
        if (!this.is(this.peek(), "paren", "(")) throw this.unexpected(this.peek());
        this.take();
        return { t: "call", tok: t, args: this.parseArgs() };
      }
      case "paren": {
        if (t.text !== "(") throw this.unexpected(t);
        this.take();
        const items = [this.parseExpr()];
        while (this.is(this.peek(), "sep", ",")) {
          this.take();
          items.push(this.parseExpr());
        }
        if (!this.is(this.peek(), "paren", ")")) throw this.unexpected(this.peek());
        this.take();
        return { t: "paren", items };
      }
      case "array":
        if (t.text !== "{") throw this.unexpected(t);
        return this.parseArray();
      default:
        throw this.unexpected(t);
    }
  }

  /** Arguments after an opening parenthesis, up to and including the closing one. Empty arguments are null. */
  private parseArgs(): (Node | null)[] {
    const args: (Node | null)[] = [];
    if (this.is(this.peek(), "paren", ")")) {
      this.take();
      return args;
    }
    for (;;) {
      const t = this.peek();
      args.push(this.is(t, "sep", ",") || this.is(t, "paren", ")") ? null : this.parseExpr());
      const sep = this.take();
      if (this.is(sep, "sep", ",")) continue;
      if (this.is(sep, "paren", ")")) return args;
      throw this.unexpected(sep);
    }
  }

  private parseArray(): Node {
    const open = this.take();
    const items: LexToken[] = [];
    for (;;) {
      let signed = false;
      if (this.isOp(this.peek(), ["-", "+"])) {
        this.take();
        signed = true;
      }
      const v = this.take();
      const ok = signed ? v.type === "number" : ["number", "string", "bool", "error"].includes(v.type);
      if (!ok) {
        throw new FormulaParseError("An array constant can only hold numbers, text, TRUE, FALSE and errors.");
      }
      items.push(v);
      const sep = this.take();
      if (sep.type === "sep") continue;
      if (this.is(sep, "array", "}")) return { t: "array", tok: open, items };
      throw this.unexpected(sep);
    }
  }
}

// ---------------------------------------------------------------------------------------------
// Rules

interface Area extends AreaShape {
  sheet: string;
}

type Binding = { kind: "let"; def: Node | null } | { kind: "param" };
type Scope = ReadonlyMap<string, Binding>;

const URL_RE = /https?:|ftp:|file:|\\\\/i;
const LOCAL_NAME_RE = /^[A-Za-z_][A-Za-z0-9_.]*$/;
/** A cell corner of a reference: its row part, and whether a $ fixes that row. */
const CORNER_RE = /^\$?[A-Za-z]{1,3}(\$?)(\d{1,7})$/;
const CORNERS_RE = /(\$?[A-Za-z]{1,3})(\$?)(\d+)/g;

function upper(s: string): string {
  return s.toUpperCase();
}

/**
 * Scope key for a LET or LAMBDA name as written, or null. Only _xlpm. marks a local name; any other
 * _xl prefix (_xlfn., _xlws., a second _xlpm.) belongs to functions, so such a name is never a local.
 */
function localKey(text: string): string | null {
  const rest = text.replace(/^_xlpm\./i, "");
  return /^_xl[a-z]*\./i.test(rest) ? null : normalizeFunctionName(rest);
}

/** Formula text shown in a detail: shortened so a long name can't flood the UI. */
function show(text: string): string {
  const flat = text.replace(/\s+/g, " ");
  return flat.length > 60 ? `${flat.slice(0, 57)}...` : flat;
}

function thousands(n: number): string {
  return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function intervalCovered(lo: number, hi: number, intervals: [number, number][]): boolean {
  let cur = lo;
  for (const [s, e] of [...intervals].sort((x, y) => x[0] - y[0])) {
    if (s > cur) break;
    if (e >= cur) cur = e + 1;
    if (cur > hi) return true;
  }
  return cur > hi;
}

/** True when every cell of `a` lies in the union of `rects`. */
function rectCovered(a: AreaShape, rects: AreaShape[]): boolean {
  const rs = rects.filter((r) => r.c1 <= a.c2 && r.c2 >= a.c1 && r.r1 <= a.r2 && r.r2 >= a.r1);
  const xs = new Set([a.c1, a.c2 + 1]);
  const ys = new Set([a.r1, a.r2 + 1]);
  for (const r of rs) {
    xs.add(Math.max(r.c1, a.c1)).add(Math.min(r.c2, a.c2) + 1);
    ys.add(Math.max(r.r1, a.r1)).add(Math.min(r.r2, a.r2) + 1);
  }
  const xl = [...xs].sort((p, q) => p - q);
  const yl = [...ys].sort((p, q) => p - q);
  for (let i = 0; i < xl.length - 1; i++) {
    for (let j = 0; j < yl.length - 1; j++) {
      const x = xl[i]!;
      const y = yl[j]!;
      if (!rs.some((r) => r.c1 <= x && x <= r.c2 && r.r1 <= y && y <= r.r2)) return false;
    }
  }
  return true;
}

class Rules {
  readonly reasons: GateResult["reasons"] = [];
  private readonly seen = new Set<string>();
  private readonly allowed: Area[] = [];
  private readonly bindings = new Map<Node, Binding>();
  private readonly tableRefs = new Map<Node, TableInfo | undefined>();
  private readonly rangeOps: Extract<Node, { t: "bin" }>[] = [];
  private readonly allRefsCache = new Map<number, { areas: Area[]; unknown: boolean }>();

  constructor(
    private readonly ctx: GateContext,
    private readonly tokens: readonly LexToken[],
    private readonly root: Node,
    /** Rows the formula is filled down below its cell, 0 when it isn't. */
    private readonly shift = 0,
  ) {
    for (const entry of ctx.allowedRanges ?? []) {
      const { sheet, address } = splitSheetAddress(entry);
      const shape = parseArea(address);
      if (shape) this.allowed.push({ ...shape, sheet: sheet ?? ctx.sheet });
    }
  }

  run(): void {
    this.walk(this.root, new Map());
    for (const t of this.tokens) this.checkToken(t);
    const inside = this.rangeOps.filter((op) => this.checkRangeOp(op, 0));
    if (this.shift <= 0) return;
    // Filled down, Excel evaluates a moved copy in every row. Only what passed as written is checked
    // again, so nothing is reported twice.
    for (const t of this.tokens) this.checkFilledRef(t);
    for (const op of inside) this.checkRangeOp(op, this.shift);
  }

  private add(code: GateCode, detail: string): void {
    const key = `${code}\u0000${detail}`;
    if (this.seen.has(key)) return;
    this.seen.add(key);
    this.reasons.push({ code, detail });
  }

  // --- sheets, tables and areas

  /** Canonical sheet name for a qualified or unqualified token, or null when it isn't a known sheet. */
  private sheetOf(t: LexToken): string | null {
    if (!t.prefix) return this.ctx.sheet;
    if (t.prefix.kind !== "sheet") return null;
    const key = upper(t.prefix.sheet);
    return (this.ctx.sheets ?? []).find((s) => upper(s) === key) ?? null;
  }

  private checkPrefix(t: LexToken): void {
    const p = t.prefix;
    if (!p) return;
    const shown = show(p.raw.slice(0, -1));
    if (p.kind === "workbook") this.add("G-EXT", `Refers to another workbook or file (${shown}).`);
    else if (p.kind === "3d") this.add("G-EXT", `Refers to several sheets at once (${shown}).`);
    else if (this.sheetOf(t) === null) this.add("G-EXT", `Refers to sheet ${shown}, which isn't in this workbook.`);
  }

  private refArea(t: LexToken): Area | null {
    const sheet = this.sheetOf(t);
    const shape = sheet === null ? null : parseArea(t.body ?? t.text);
    return shape && sheet !== null ? { ...shape, sheet } : null;
  }

  private shadowedByDefinedName(fn: string): boolean {
    return (this.ctx.definedNames ?? []).some((n) => upper(n) === fn || upper(n.replace(/^.*!/, "")) === fn);
  }

  private findTable(name: string): { known: boolean; info: TableInfo | undefined } {
    const key = upper(name);
    const info = (this.ctx.tables ?? []).find((x) => upper(x.name) === key);
    const known = info !== undefined || (this.ctx.tableNames ?? []).some((n) => upper(n) === key);
    return { known, info };
  }

  private tableArea(info: TableInfo | undefined): Area | null {
    if (!info) return null;
    const shape = parseArea(info.address);
    return shape ? { ...shape, sheet: info.sheet } : null;
  }

  private covered(a: Area): boolean {
    if (a.kind === "rows") return false;
    const key = upper(a.sheet);
    const same = this.allowed.filter((x) => upper(x.sheet) === key);
    if (a.kind === "cols") return intervalCovered(a.c1, a.c2, same.map((x): [number, number] => [x.c1, x.c2]));
    return rectCovered(a, same);
  }

  private checkTableUse(name: string, info: TableInfo | undefined): void {
    const area = this.tableArea(info);
    if (!info) this.add("G-REF", `Uses table ${show(name)}, whose location isn't known.`);
    else if (!area || !this.covered(area)) this.add("G-REF", `Uses table ${show(name)}, which lies outside the ranges you chose.`);
  }

  // --- token rules: G-URL, G-EXT, G-REF

  private checkToken(t: LexToken): void {
    switch (t.type) {
      case "string":
        if (URL_RE.test(t.value ?? "")) this.add("G-URL", "Contains a web address or file path in a text value.");
        return;
      case "name":
      case "func":
        this.checkPrefix(t);
        return;
      case "ref": {
        this.checkPrefix(t);
        if (t.spill) {
          this.add("G-REF", `Uses the spill range of ${show(t.text.slice(0, -1))}, whose size isn't known.`);
          return;
        }
        const area = this.refArea(t);
        if (!area) return;
        if (area.kind === "rows") {
          this.add("G-REF", `Refers to ${show(t.text)}, a whole row. Use cells inside the ranges you chose.`);
        } else if (!this.covered(area)) {
          this.add("G-REF", `Refers to ${show(t.text)}, outside the ranges you chose.`);
        }
        return;
      }
      case "struct": {
        if (t.table === null || t.table === undefined) {
          this.add("G-EXT", `Uses ${show(t.text)} without a table name.`);
          return;
        }
        const { known, info } = this.findTable(t.table);
        if (!known) this.add("G-EXT", `Uses ${show(t.text)}, but ${show(t.table)} isn't a table in this workbook.`);
        else this.checkTableUse(t.table, info);
        return;
      }
    }
  }

  // --- fill-down: Excel adds one to every row number without a $ for each row it fills

  /** Row (0-based) of each corner of a cell or range reference, and whether it moves when filled. */
  private corners(t: LexToken): { r: number; moves: boolean }[] {
    return (t.body ?? t.text).split(":").map((part) => {
      const m = CORNER_RE.exec(part);
      if (!m) throw new FormulaParseError();
      return { r: Number(m[2]) - 1, moves: m[1] === "" };
    });
  }

  /** The area a cell or range reference covers in the copy `k` rows down. */
  private areaAt(t: LexToken, a: Area, k: number): Area {
    const rows = this.corners(t).map((c) => c.r + (c.moves ? k : 0));
    return { ...a, r1: Math.min(...rows), r2: Math.max(...rows) };
  }

  /** Every cell a reference covers in any copy up to `shift` rows down. Whole columns and rows keep their area. */
  private sweptArea(t: LexToken, shift: number): Area | null {
    const a = this.refArea(t);
    if (!a || a.kind !== "cells" || shift <= 0) return a;
    return { ...a, r2: Math.max(...this.corners(t).map((c) => c.r + (c.moves ? shift : 0))) };
  }

  /**
   * The first copy down that reaches outside the allowed ranges or the sheet, or null. Coverage only
   * changes where a moving row crosses the edge of an allowed range or of the sheet, so only those
   * copies (and the last one) are tried.
   */
  private firstCopyOutside(t: LexToken, a: Area): number | null {
    const moving = this.corners(t).filter((c) => c.moves);
    const edges = [MAX_ROWS];
    for (const x of this.allowed) if (upper(x.sheet) === upper(a.sheet)) edges.push(x.r1, x.r2 + 1);
    const steps = new Set([this.shift]);
    for (const e of edges) for (const c of moving) if (e - c.r > 0 && e - c.r < this.shift) steps.add(e - c.r);
    for (const k of [...steps].sort((p, q) => p - q)) {
      const moved = this.areaAt(t, a, k);
      if (moved.r2 >= MAX_ROWS || !this.covered(moved)) return k;
    }
    return null;
  }

  private checkFilledRef(t: LexToken): void {
    if (t.type !== "ref" || t.spill) return;
    const a = this.refArea(t);
    // Whole columns don't move; whole rows and references outside as written are already reported.
    if (!a || a.kind !== "cells" || !this.covered(a)) return;
    const k = this.firstCopyOutside(t, a);
    if (k === null) return;
    const rows = `${thousands(k)} ${k === 1 ? "row" : "rows"}`;
    if (this.areaAt(t, a, k).r2 >= MAX_ROWS) {
      this.add("G-REF", `Filled down ${rows}, ${show(t.text)} would move past the last row of the sheet.`);
      return;
    }
    const body = (t.body ?? t.text).replace(CORNERS_RE, (m, col: string, fixed: string, row: string) =>
      fixed ? m : `${col}${Number(row) + k}`,
    );
    this.add("G-REF", `Filled down ${rows}, it would refer to ${show((t.prefix?.raw ?? "") + body)}, outside the ranges you chose.`);
  }

  // --- AST rules: G-FUNC, G-NAME, LET and LAMBDA scopes

  private walk(node: Node | null, scope: Scope): void {
    if (!node) return;
    switch (node.t) {
      case "lit":
      case "array":
      case "ref":
      case "struct":
        return;
      case "name":
        this.walkName(node, scope);
        return;
      case "call":
        this.walkCall(node, scope);
        return;
      case "invoke":
        this.walk(node.callee, scope);
        for (const a of node.args) this.walk(a, scope);
        return;
      case "paren":
        for (const item of node.items) this.walk(item, scope);
        return;
      case "unary":
      case "postfix":
        this.walk(node.arg, scope);
        return;
      case "bin":
        if (node.op === ":") this.rangeOps.push(node);
        this.walk(node.left, scope);
        this.walk(node.right, scope);
        return;
    }
  }

  private local(text: string, scope: Scope): Binding | undefined {
    const key = localKey(text);
    return key === null ? undefined : scope.get(key);
  }

  private walkName(node: Extract<Node, { t: "name" }>, scope: Scope): void {
    const t = node.tok;
    const text = t.body ?? t.text;
    if (t.prefix) {
      this.add("G-NAME", `Uses the name ${show(t.text)}, which isn't a cell reference. Names can hide other formulas.`);
      return;
    }
    const binding = this.local(text, scope);
    if (binding) {
      this.bindings.set(node, binding);
      return;
    }
    const { known, info } = this.findTable(text);
    if (known) {
      this.tableRefs.set(node, info);
      this.checkTableUse(text, info);
      return;
    }
    const fn = normalizeFunctionName(text);
    if (Object.hasOwn(EXCLUDED_REASONS, fn)) {
      this.add("G-FUNC", `Uses ${fn}, which ${EXCLUDED_REASONS[fn]}.`);
      return;
    }
    if (ALLOWLIST.has(fn)) {
      // An allowed function passed by name, as GROUPBY(..., SUM) or MAP(..., ABS) do. It can only
      // call that same allowed function, unless a workbook name shadows it.
      if (this.shadowedByDefinedName(fn)) {
        this.add("G-NAME", `Uses ${show(text)} on its own, and the workbook also has a name ${fn}. Names can hide other formulas.`);
      }
      return;
    }
    this.add("G-NAME", `Uses the name ${show(text)}, which isn't a cell reference. Names can hide other formulas.`);
  }

  private walkCall(node: Extract<Node, { t: "call" }>, scope: Scope): void {
    const t = node.tok;
    const text = t.body ?? t.text;
    const walkArgs = (s: Scope) => {
      for (const a of node.args) this.walk(a, s);
    };
    if (t.prefix) {
      this.add("G-FUNC", `Calls ${show(t.text)} through a sheet or workbook name.`);
      walkArgs(scope);
      return;
    }
    const binding = this.local(text, scope);
    if (binding) {
      this.bindings.set(node, binding);
      walkArgs(scope);
      return;
    }
    const fn = normalizeFunctionName(text);
    if (!ALLOWLIST.has(fn)) {
      if (Object.hasOwn(EXCLUDED_REASONS, fn)) {
        this.add("G-FUNC", `Uses ${fn}, which ${EXCLUDED_REASONS[fn]}.`);
      } else if (/[^\x20-\x7e]/.test(text)) {
        this.add("G-FUNC", `Uses ${show(text)}, which isn't on the list of allowed functions. Its name has characters outside A to Z.`);
      } else {
        this.add("G-FUNC", `Uses ${show(text)}, which isn't on the list of allowed functions.`);
      }
      walkArgs(scope);
      return;
    }
    if (fn === "LET") this.walkLet(node, scope);
    else if (fn === "LAMBDA") this.walkLambda(node, scope);
    else walkArgs(scope);
  }

  /**
   * Validates a LET name or LAMBDA parameter and returns its scope key, or null when it isn't a name
   * at all. A name that hides a function is rejected but still returned, so later uses of it don't
   * add a second, confusing reason.
   */
  private declare(decl: Node | null, fn: "LET" | "LAMBDA"): string | null {
    if (!decl) {
      this.add("G-FUNC", `Uses ${fn} with an empty name.`);
      return null;
    }
    if (decl.t === "ref") {
      this.add("G-FUNC", `Uses ${fn} with ${show(decl.tok.text)} as a name, but that is a cell reference.`);
      return null;
    }
    if (decl.t !== "name" || decl.tok.prefix) {
      this.add("G-FUNC", `Uses ${fn} with something other than a plain name where a name belongs.`);
      return null;
    }
    const text = decl.tok.body ?? decl.tok.text;
    const key = localKey(text);
    if (key === null) {
      this.add("G-FUNC", `Uses ${fn} with the name ${show(text)}, which has a prefix only functions use.`);
      return null;
    }
    if (!LOCAL_NAME_RE.test(text) || !LOCAL_NAME_RE.test(key) || parseCell(key)) {
      this.add("G-FUNC", `Uses ${fn} with the name ${show(text)}, which isn't a plain name.`);
      return null;
    }
    // Shadowing an allowed function (LET(n, ...), LAMBDA(value, ...)) is harmless: whichever Excel
    // calls is allowed or gated. An excluded function name is never a local.
    if (Object.hasOwn(EXCLUDED_REASONS, key)) {
      this.add("G-FUNC", `Uses ${show(text)} as a ${fn} name, but ${key} is also a function that isn't allowed. Pick another name.`);
    }
    return key;
  }

  private walkLet(node: Extract<Node, { t: "call" }>, scope: Scope): void {
    const args = node.args;
    if (args.length < 3 || args.length % 2 === 0) {
      this.add("G-FUNC", "Uses LET without a name and value for each pair, or without a final calculation.");
      for (const a of args) this.walk(a, scope);
      return;
    }
    let inner = new Map(scope);
    for (let k = 0; k < args.length - 1; k += 2) {
      const key = this.declare(args[k] ?? null, "LET");
      const value = args[k + 1] ?? null;
      this.walk(value, inner);
      if (key !== null) {
        inner = new Map(inner);
        inner.set(key, { kind: "let", def: value });
      }
    }
    this.walk(args[args.length - 1] ?? null, inner);
  }

  private walkLambda(node: Extract<Node, { t: "call" }>, scope: Scope): void {
    const args = node.args;
    const body = args[args.length - 1];
    if (!body) {
      this.add("G-FUNC", "Uses LAMBDA without a calculation.");
      for (const a of args) this.walk(a, scope);
      return;
    }
    const inner = new Map(scope);
    for (let k = 0; k < args.length - 1; k++) {
      const key = this.declare(args[k] ?? null, "LAMBDA");
      if (key !== null) inner.set(key, { kind: "param" });
    }
    this.walk(body, inner);
  }

  // --- the range operator: A1:INDEX(...), A2:F2:I20, A1:r

  /**
   * Every area a node's value could reach as a reference. Reference values only come from reference
   * literals, tables and functions that pass part of a reference on, so the union of the literals
   * inside a node (following LET names to their values) bounds it. A LAMBDA parameter can hold any
   * reference in the formula. `unknown` means a spill range, whose size isn't known. With `shift`,
   * cell references cover every row they move through when filled down that far.
   */
  private refsOf(node: Node | null, seen: Set<Node>, shift: number): { areas: Area[]; unknown: boolean } {
    const acc = { areas: [] as Area[], unknown: false };
    const merge = (r: { areas: Area[]; unknown: boolean }) => {
      acc.areas.push(...r.areas);
      acc.unknown ||= r.unknown;
    };
    const follow = (n: Node) => {
      const b = this.bindings.get(n);
      if (!b) return;
      if (b.kind === "param") merge(this.allRefs(shift));
      else if (b.def && !seen.has(b.def)) {
        seen.add(b.def);
        merge(this.refsOf(b.def, seen, shift));
      }
    };
    if (!node) return acc;
    switch (node.t) {
      case "lit":
      case "array":
        break;
      case "ref": {
        const a = this.sweptArea(node.tok, shift);
        if (a) acc.areas.push(a);
        if (node.tok.spill) acc.unknown = true;
        break;
      }
      case "struct": {
        const a = node.tok.table ? this.tableArea(this.findTable(node.tok.table).info) : null;
        if (a) acc.areas.push(a);
        break;
      }
      case "name": {
        follow(node);
        const a = this.tableRefs.has(node) ? this.tableArea(this.tableRefs.get(node)) : null;
        if (a) acc.areas.push(a);
        break;
      }
      case "call":
        follow(node);
        for (const a of node.args) merge(this.refsOf(a, seen, shift));
        break;
      case "invoke":
        merge(this.refsOf(node.callee, seen, shift));
        for (const a of node.args) merge(this.refsOf(a, seen, shift));
        break;
      case "paren":
        for (const item of node.items) merge(this.refsOf(item, seen, shift));
        break;
      case "unary":
      case "postfix":
        merge(this.refsOf(node.arg, seen, shift));
        break;
      case "bin":
        merge(this.refsOf(node.left, seen, shift));
        merge(this.refsOf(node.right, seen, shift));
        break;
    }
    return acc;
  }

  /** Every reference literal and table in the formula, without following names. */
  private allRefs(shift: number): { areas: Area[]; unknown: boolean } {
    const cached = this.allRefsCache.get(shift);
    if (cached) return cached;
    const acc = { areas: [] as Area[], unknown: false };
    for (const t of this.tokens) {
      if (t.type === "ref") {
        const a = this.sweptArea(t, shift);
        if (a) acc.areas.push(a);
        if (t.spill) acc.unknown = true;
      } else if (t.type === "struct" && t.table) {
        const a = this.tableArea(this.findTable(t.table).info);
        if (a) acc.areas.push(a);
      }
    }
    for (const info of this.tableRefs.values()) {
      const a = this.tableArea(info);
      if (a) acc.areas.push(a);
    }
    this.allRefsCache.set(shift, acc);
    return acc;
  }

  /** Checks the box a : operator builds, as written or swept `shift` rows down. False when it adds a reason. */
  private checkRangeOp(node: Extract<Node, { t: "bin" }>, shift: number): boolean {
    const literalSheets = new Set<string>();
    const collect = (n: Node) => {
      if (n.t === "bin" && n.op === ":") {
        collect(n.left);
        collect(n.right);
      } else if (n.t === "ref") {
        const s = this.sheetOf(n.tok);
        if (s !== null) literalSheets.add(upper(s));
      }
    };
    collect(node);
    if (literalSheets.size > 1) {
      this.add("G-REF", "Joins references from different sheets with a colon.");
      return false;
    }
    const { areas, unknown } = this.refsOf(node, new Set(), shift);
    if (unknown) {
      this.add("G-REF", "Builds a range with a colon from a spill range, which can reach outside the ranges you chose.");
      return false;
    }
    const bySheet = new Map<string, Area[]>();
    for (const a of areas) {
      const key = upper(a.sheet);
      bySheet.set(key, [...(bySheet.get(key) ?? []), a]);
    }
    for (const group of bySheet.values()) {
      const first = group[0]!;
      const box: Area = { ...first };
      for (const a of group) {
        box.r1 = Math.min(box.r1, a.r1);
        box.c1 = Math.min(box.c1, a.c1);
        box.r2 = Math.max(box.r2, a.r2);
        box.c2 = Math.max(box.c2, a.c2);
        if (a.kind === "rows" || box.kind === "rows") box.kind = "rows";
        else if (a.kind === "cols") box.kind = "cols";
      }
      if (!this.covered(box)) {
        this.add(
          "G-REF",
          shift > 0
            ? "When filled down, it builds a range with a colon that reaches outside the ranges you chose."
            : "Builds a range with a colon that reaches outside the ranges you chose.",
        );
        return false;
      }
    }
    return true;
  }
}

function fail(code: GateCode, detail: string): GateResult {
  return { ok: false, reasons: [{ code, detail }] };
}

function parseFailure(e: unknown): GateResult {
  const message = e instanceof FormulaParseError ? e.message : "";
  return fail("G-PARSE", message ? `Couldn't read this formula. ${message}` : "Couldn't read this formula.");
}

function run(formula: string, ctx: GateContext, shift: number): GateResult {
  if (typeof formula !== "string" || !formula.startsWith("=")) {
    return fail("G-START", "Doesn't start with =, so it isn't a formula.");
  }
  if (formula.length > MAX_FORMULA_LENGTH) {
    // Never the exact length: after restore it would give away the length of a private value.
    return fail("G-LEN", `This formula is longer than ${thousands(MAX_FORMULA_LENGTH)} characters.`);
  }
  const tokens = new Lexer(formula).run();
  const root = new Parser(tokens).parseFormula();
  const rules = new Rules(ctx, tokens, root, shift);
  rules.run();
  return rules.reasons.length === 0 ? { ok: true, reasons: [] } : { ok: false, reasons: rules.reasons };
}

/** Runs every gate rule. Fails closed: any parse failure is G-PARSE. */
export function check(formula: string, ctx: GateContext): GateResult {
  try {
    return run(formula, ctx, 0);
  } catch (e) {
    return parseFailure(e);
  }
}

/**
 * check() for a formula that Insert fills `rowShift` rows down from its cell. Excel moves every row
 * number without a $ by one per filled row, so each such reference must stay inside the allowed
 * ranges in every filled cell (G-REF). Never throws; a distance it can't read fails closed.
 */
export function checkFill(formula: string, ctx: GateContext, rowShift: number): GateResult {
  try {
    if (typeof rowShift !== "number" || !(Number.isInteger(rowShift) || Math.abs(rowShift) === Infinity)) {
      const r = check(formula, ctx);
      return { ok: false, reasons: [...r.reasons, { code: "G-REF", detail: "Couldn't tell how far it would be filled down." }] };
    }
    if (rowShift <= 0) return check(formula, ctx);
    // A fill as tall as the sheet moves every row without a $ past its end; a longer one adds nothing.
    return run(formula, ctx, Math.min(rowShift, MAX_ROWS));
  } catch (e) {
    return parseFailure(e);
  }
}
