// Source checks for the spec §2 invariants, run over the repository's own files. They parse each
// file with the TypeScript compiler, so comments, strings and JSX text don't cause false alarms.
// The ESLint rules in eslint.config.js enforce the same things while editing; these tests are the
// second line and also cover bench/ and tests/.
import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { dirname, join, posix, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import ts from "typescript";
import { describe, expect, it } from "vitest";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const CODE_FILE = /\.(?:ts|tsx|mts|cts|js|jsx|mjs|cjs)$/;
const SKIP_DIRS = new Set(["node_modules", "dist", "coverage", ".git", "playwright-report", "test-results"]);

// Built by concatenation so this file never contains the patterns it looks for.
const AUDITED = "Audited" + "Outbound";
const RAW_MARKER = "raw" + "_bench";

/** Repo-relative paths (forward slashes) of code files under `dir`. */
function listCode(dir: string): string[] {
  const abs = join(ROOT, dir);
  if (!existsSync(abs)) return [];
  const out: string[] = [];
  const walkDir = (d: string) => {
    for (const name of readdirSync(d)) {
      const p = join(d, name);
      if (statSync(p).isDirectory()) {
        if (!SKIP_DIRS.has(name)) walkDir(p);
      } else if (CODE_FILE.test(name)) {
        out.push(relative(ROOT, p).split(sep).join("/"));
      }
    }
  };
  walkDir(abs);
  return out.sort();
}

function scriptKind(name: string): ts.ScriptKind {
  if (name.endsWith(".tsx")) return ts.ScriptKind.TSX;
  if (name.endsWith(".jsx")) return ts.ScriptKind.JSX;
  if (/\.[cm]?js$/.test(name)) return ts.ScriptKind.JS;
  return ts.ScriptKind.TS;
}

function parse(name: string, text: string): ts.SourceFile {
  return ts.createSourceFile(name, text, ts.ScriptTarget.Latest, true, scriptKind(name));
}

const cache = new Map<string, { text: string; sf: ts.SourceFile; code: string }>();

/** The file's text, its syntax tree, and its code with comments removed. */
function load(rel: string): { text: string; sf: ts.SourceFile; code: string } {
  let entry = cache.get(rel);
  if (!entry) {
    const text = readFileSync(join(ROOT, rel), "utf8");
    const sf = parse(rel, text);
    entry = { text, sf, code: stripComments(sf) };
    cache.set(rel, entry);
  }
  return entry;
}

function stripComments(sf: ts.SourceFile): string {
  return ts.createPrinter({ removeComments: true }).printFile(sf);
}

function walk(node: ts.Node, visit: (n: ts.Node) => void): void {
  visit(node);
  node.forEachChild((child) => walk(child, visit));
}

function lineOf(sf: ts.SourceFile, node: ts.Node): number {
  return sf.getLineAndCharacterOfPosition(node.getStart(sf)).line + 1;
}

function entityLast(name: ts.EntityName): string {
  return ts.isIdentifier(name) ? name.text : name.right.text;
}

function within(node: ts.Node, outer: ts.Node): boolean {
  return node.pos >= outer.pos && node.end <= outer.end;
}

// ---------------------------------------------------------------------------------------------
// Detectors. Each returns the lines (1-based) where it found something.

/** `x as AuditedOutbound`, `<AuditedOutbound>x`, and assertions to any type that mentions it. */
function auditedCasts(sf: ts.SourceFile, code: string): number[] {
  const lines: number[] = [];
  walk(sf, (n) => {
    if (ts.isAsExpression(n) || ts.isTypeAssertionExpression(n)) {
      let mentions = false;
      walk(n.type, (t) => {
        if (ts.isTypeReferenceNode(t) && entityLast(t.typeName) === AUDITED) mentions = true;
        if (ts.isImportTypeNode(t) && t.qualifier && entityLast(t.qualifier) === AUDITED) mentions = true;
      });
      if (mentions) lines.push(lineOf(sf, n));
    }
  });
  // Plain-text backstop on the comment-free code.
  const asCast = new RegExp(`\\bas\\s+(?:[\\w$]+\\.)*${AUDITED}\\b`);
  const angleCast = new RegExp(`(?<![\\w$)\\]>.]\\s*)<\\s*${AUDITED}\\s*>`);
  if (lines.length === 0 && (asCast.test(code) || angleCast.test(code))) lines.push(0);
  return lines;
}

/** Calls to fetch: `fetch(...)`, `window.fetch(...)`, `x["fetch"](...)`, and the text `fetch(`. */
function fetchCalls(sf: ts.SourceFile, code: string): number[] {
  const lines: number[] = [];
  walk(sf, (n) => {
    if (!ts.isCallExpression(n)) return;
    const callee = n.expression;
    const named =
      (ts.isIdentifier(callee) && callee.text === "fetch") ||
      (ts.isPropertyAccessExpression(callee) && callee.name.text === "fetch") ||
      (ts.isElementAccessExpression(callee) &&
        ts.isStringLiteralLike(callee.argumentExpression) &&
        callee.argumentExpression.text === "fetch");
    if (named) lines.push(lineOf(sf, n));
  });
  if (lines.length === 0 && /\bfetch\s*\(/.test(code)) lines.push(0);
  return lines;
}

/** References to the Office.js globals: `Office.x`, `Excel.x`, `Excel.Range` in types, `window.Office`. */
function officeRefs(sf: ts.SourceFile): number[] {
  const lines: number[] = [];
  const globals = new Set(["Office", "Excel", "OfficeExtension"]);
  walk(sf, (n) => {
    if (!ts.isIdentifier(n) || !globals.has(n.text)) return;
    const p = n.parent;
    const hit =
      (ts.isPropertyAccessExpression(p) && p.expression === n) ||
      (ts.isQualifiedName(p) && p.left === n) ||
      (ts.isPropertyAccessExpression(p) &&
        p.name === n &&
        ts.isIdentifier(p.expression) &&
        ["window", "globalThis", "self"].includes(p.expression.text)) ||
      (ts.isTypeQueryNode(p) && p.exprName === n);
    if (hit) lines.push(lineOf(sf, n));
  });
  return lines;
}

const STORAGE_RE = /\b(?:localStorage|sessionStorage|indexedDB|sendBeacon)\b|\bdocument\s*\.\s*cookie\b/;

/** Browser storage and beacons, searched in comment-free code (strings included). */
function storageUses(code: string): number[] {
  return code.split("\n").flatMap((line, i) => (STORAGE_RE.test(line) ? [i + 1] : []));
}

function mentionsBench(expr: ts.Expression): boolean {
  if (ts.isParenthesizedExpression(expr)) return mentionsBench(expr.expression);
  if (ts.isIdentifier(expr)) return expr.text === "BENCH" || expr.text === "__NYMFORM_BENCH__";
  if (ts.isBinaryExpression(expr) && expr.operatorToken.kind === ts.SyntaxKind.AmpersandAmpersandToken) {
    return mentionsBench(expr.left) || mentionsBench(expr.right);
  }
  return false;
}

/** True when `node` only runs when BENCH is true: `if (BENCH) {..}`, `BENCH ? .. :`, `BENCH && ..`. */
function benchGuarded(node: ts.Node): boolean {
  for (let p: ts.Node | undefined = node.parent; p; p = p.parent) {
    if (ts.isIfStatement(p) && within(node, p.thenStatement) && mentionsBench(p.expression)) return true;
    if (ts.isConditionalExpression(p) && within(node, p.whenTrue) && mentionsBench(p.condition)) return true;
    if (
      ts.isBinaryExpression(p) &&
      p.operatorToken.kind === ts.SyntaxKind.AmpersandAmpersandToken &&
      within(node, p.right) &&
      mentionsBench(p.left)
    ) {
      return true;
    }
  }
  return false;
}

function pointsIntoBench(fromRel: string, specifier: string): boolean {
  if (specifier.startsWith(".")) {
    const target = posix.normalize(posix.join(posix.dirname(fromRel), specifier));
    return target === "bench" || target.startsWith("bench/");
  }
  return specifier === "bench" || specifier.startsWith("bench/");
}

/** Imports of bench/ code that would land in a release bundle. `import type` is erased and allowed. */
function benchImports(sf: ts.SourceFile, rel: string): number[] {
  const lines: number[] = [];
  walk(sf, (n) => {
    if (ts.isImportDeclaration(n) && ts.isStringLiteral(n.moduleSpecifier)) {
      if (!n.importClause?.isTypeOnly && pointsIntoBench(rel, n.moduleSpecifier.text)) lines.push(lineOf(sf, n));
    } else if (ts.isExportDeclaration(n) && n.moduleSpecifier && ts.isStringLiteral(n.moduleSpecifier)) {
      if (!n.isTypeOnly && pointsIntoBench(rel, n.moduleSpecifier.text)) lines.push(lineOf(sf, n));
    } else if (ts.isImportEqualsDeclaration(n) && ts.isExternalModuleReference(n.moduleReference)) {
      const e = n.moduleReference.expression;
      if (ts.isStringLiteral(e) && pointsIntoBench(rel, e.text)) lines.push(lineOf(sf, n));
    } else if (ts.isCallExpression(n)) {
      const isRequire = ts.isIdentifier(n.expression) && n.expression.text === "require";
      const isDynamic = n.expression.kind === ts.SyntaxKind.ImportKeyword;
      if (!isRequire && !isDynamic) return;
      const arg = n.arguments[0];
      if (!arg || !ts.isStringLiteralLike(arg)) {
        // A computed specifier could point anywhere, bench/ included.
        if (isDynamic) lines.push(lineOf(sf, n));
        return;
      }
      if (!pointsIntoBench(rel, arg.text)) return;
      if (isRequire || !benchGuarded(n)) lines.push(lineOf(sf, n));
    }
  });
  return lines;
}

function identifierUses(sf: ts.SourceFile, name: string): number[] {
  const lines: number[] = [];
  walk(sf, (n) => {
    if (ts.isIdentifier(n) && n.text === name) lines.push(lineOf(sf, n));
  });
  return lines;
}

/** Occurrences of the raw-bench marker in comment-free code that sit outside a `type X = ...` declaration. */
function markerOutsideTypeAliases(rel: string, code: string): number {
  const sf = parse(rel, code);
  const spans: [number, number][] = [];
  walk(sf, (n) => {
    if (ts.isTypeAliasDeclaration(n)) spans.push([n.getStart(sf), n.end]);
  });
  let outside = 0;
  for (let i = code.indexOf(RAW_MARKER); i !== -1; i = code.indexOf(RAW_MARKER, i + 1)) {
    if (!spans.some(([a, b]) => i >= a && i + RAW_MARKER.length <= b)) outside += 1;
  }
  return outside;
}

function report(files: string[], find: (rel: string) => number[]): string[] {
  return files.flatMap((rel) => find(rel).map((line) => `${rel}:${line}`));
}

const SRC = () => listCode("src");
const BENCH = () => listCode("bench");
const EVERYTHING = () => [...listCode("src"), ...listCode("tests"), ...listCode("bench"), ...listCode("scripts")];

// ---------------------------------------------------------------------------------------------

describe("invariants: source checks", () => {
  it("finds the source files it checks", () => {
    expect(SRC()).toContain("src/core/auditor.ts");
    expect(SRC()).toContain("src/core/types.ts");
    expect(listCode("tests")).toContain("tests/invariants.test.ts");
  });

  it(`T-A10: only src/core/auditor.ts asserts a value to ${AUDITED} (invariant 3)`, () => {
    const found = report(
      EVERYTHING().filter((f) => f !== "src/core/auditor.ts"),
      (rel) => {
        const { sf, code } = load(rel);
        return auditedCasts(sf, code);
      },
    );
    expect(found).toEqual([]);
  });

  it("only src/core/provider.ts and bench/rawSender.ts call fetch (invariant 2)", () => {
    const allowed = new Set(["src/core/provider.ts", "bench/rawSender.ts"]);
    const found = report(
      [...SRC(), ...BENCH()].filter((f) => !allowed.has(f)),
      (rel) => {
        const { sf, code } = load(rel);
        return fetchCalls(sf, code);
      },
    );
    expect(found).toEqual([]);
  });

  it("in src/, only src/office/adapter.ts references the Office and Excel globals", () => {
    const found = report(
      SRC().filter((f) => f !== "src/office/adapter.ts"),
      (rel) => officeRefs(load(rel).sf),
    );
    expect(found).toEqual([]);
  });

  it("no src/ file uses localStorage, sessionStorage, indexedDB, document.cookie or sendBeacon (invariants 6, 7, 11)", () => {
    expect(report(SRC(), (rel) => storageUses(load(rel).code))).toEqual([]);
  });

  it("no src/ file imports bench code, except dynamic imports guarded by if (BENCH) (invariant 8)", () => {
    expect(report(SRC(), (rel) => benchImports(load(rel).sf, rel))).toEqual([]);
  });

  it("only prompt.ts, payload.ts and auditor.ts use SYSTEM_PROMPT, so only payload.ts builds request bodies (invariant 1)", () => {
    const allowed = new Set(["src/core/prompt.ts", "src/core/payload.ts", "src/core/auditor.ts"]);
    const found = report(
      SRC().filter((f) => !allowed.has(f)),
      (rel) => identifierUses(load(rel).sf, "SYSTEM_PROMPT"),
    );
    expect(found).toEqual([]);
  });

  it(`the ${RAW_MARKER} marker appears in src/ only in src/core/types.ts type declarations (T-P4)`, () => {
    const holders = SRC().filter((rel) => load(rel).text.includes(RAW_MARKER));
    expect(holders.filter((f) => f !== "src/core/types.ts")).toEqual([]);
    if (holders.includes("src/core/types.ts")) {
      expect(markerOutsideTypeAliases("src/core/types.ts", load("src/core/types.ts").code)).toBe(0);
    }
  });

  it("src/core imports nothing from the Office host, the UI or React", () => {
    const found = report(
      SRC().filter((f) => f.startsWith("src/core/")),
      (rel) => {
        const { sf } = load(rel);
        const lines: number[] = [];
        walk(sf, (n) => {
          if ((ts.isImportDeclaration(n) || ts.isExportDeclaration(n)) && n.moduleSpecifier && ts.isStringLiteral(n.moduleSpecifier)) {
            const s = n.moduleSpecifier.text;
            if (/^\.\.\/(?:office|taskpane)\b/.test(s) || /^(?:react|react-dom|@fluentui\/)/.test(s)) lines.push(lineOf(sf, n));
          }
        });
        return lines;
      },
    );
    expect(found).toEqual([]);
  });
});

// ---------------------------------------------------------------------------------------------
// The detectors must actually detect: run them on small snippets.

function snippet(code: string, name = "src/core/x.ts") {
  const sf = parse(name, code);
  return { sf, code: stripComments(sf) };
}

describe("invariants: detectors", () => {
  it(`flags assertions to ${AUDITED} and ignores other uses`, () => {
    const bad = [
      `const a = b as ${AUDITED};`,
      `const a = b as unknown as ${AUDITED};`,
      `const a = <${AUDITED}>b;`,
      `const a = [b] as ${AUDITED}[];`,
      `const a = b as import("./auditor").${AUDITED};`,
      `const a = b as Auditor.${AUDITED};`,
    ];
    for (const c of bad) {
      const s = snippet(c);
      expect(auditedCasts(s.sf, s.code), c).not.toEqual([]);
    }
    const good = [
      `function f(x: ${AUDITED}): Promise<${AUDITED}> { return Promise.resolve(x); }`,
      `const [s, set] = useState<${AUDITED} | null>(null);`,
      `// never write b as ${AUDITED}`,
      `const o = r.outbound satisfies ${AUDITED};`,
    ];
    for (const c of good) {
      const s = snippet(c, "src/taskpane/x.tsx");
      expect(auditedCasts(s.sf, s.code), c).toEqual([]);
    }
  });

  it("flags fetch calls and ignores comments and types", () => {
    for (const c of ["fetch(u);", "await window.fetch(u);", 'globalThis["fetch"](u);', "deps.fetch (u)"]) {
      const s = snippet(c);
      expect(fetchCalls(s.sf, s.code), c).not.toEqual([]);
    }
    for (const c of ["// only provider.ts calls fetch (invariant 2)", "type F = typeof globalThis.fetch;", 'const s = "fetched";']) {
      const s = snippet(c);
      expect(fetchCalls(s.sf, s.code), c).toEqual([]);
    }
  });

  it("flags Office and Excel globals and ignores comments, strings and JSX text", () => {
    for (const c of ["Office.onReady();", "await Excel.run(async () => {});", "let r: Excel.Range;", "window.Office;"]) {
      expect(officeRefs(snippet(c).sf), c).not.toEqual([]);
    }
    const good = [
      "// no Office.js here",
      'const s = "Open this in Excel.";',
      "export const A = () => <p>Open this in Excel. Office.js loads first.</p>;",
      "import { ready } from '../office/adapter';",
    ];
    for (const c of good) expect(officeRefs(snippet(c, "src/taskpane/x.tsx").sf), c).toEqual([]);
  });

  it("flags storage APIs outside comments", () => {
    for (const c of ["localStorage.setItem('k', v);", 'window["sessionStorage"];', "document.cookie = x;", "navigator.sendBeacon(u);", "indexedDB.open('x');"]) {
      expect(storageUses(snippet(c).code), c).not.toEqual([]);
    }
    expect(storageUses(snippet("// no localStorage in v0\nconst x = 1;").code)).toEqual([]);
  });

  it("flags static bench imports and unguarded dynamic ones", () => {
    const rel = "src/taskpane/flow.ts";
    const bad = [
      'import { sendRaw } from "../../bench/rawSender";',
      'import "../../bench/rawSender";',
      'export * from "../../bench/rawSender";',
      'const m = await import("../../bench/rawSender");',
      'if (!BENCH) { await import("../../bench/rawSender"); }',
      'const r = require("../../bench/rawSender");',
      "const m = await import(path);",
    ];
    for (const c of bad) expect(benchImports(snippet(c, rel).sf, rel), c).not.toEqual([]);
    const good = [
      'if (BENCH) { const m = await import("../../bench/rawSender"); }',
      'if (BENCH && ready) { await import("../../bench/runner"); }',
      'const m = BENCH ? await import("../../bench/rawSender") : null;',
      'import type { RawResult } from "../../bench/rawSender";',
      'import { x } from "../core/benchmarks";',
    ];
    for (const c of good) expect(benchImports(snippet(c, rel).sf, rel), c).toEqual([]);
  });

  it("finds the raw-bench marker outside type declarations", () => {
    const inType = `export type Mode = "structure_only" | "substituted" | "${RAW_MARKER}";`;
    expect(markerOutsideTypeAliases("src/core/types.ts", snippet(inType).code)).toBe(0);
    const inValue = `export const MODES = ["${RAW_MARKER}"];`;
    expect(markerOutsideTypeAliases("src/core/types.ts", snippet(inValue).code)).toBe(1);
    const inComment = `// ${RAW_MARKER} lives in bench only\nexport const x = 1;`;
    expect(markerOutsideTypeAliases("src/core/types.ts", snippet(inComment).code)).toBe(0);
  });
});
