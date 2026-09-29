// Release guard (spec T-P4 and the §15 pre-publication checks that code can verify).
// Run after `npm run build`. Fails when the release bundle contains bench code, was built
// with __NYMFORM_BENCH__ true, carries the legacy product name, allows connections to any
// host other than the endpoint, or points the release manifest anywhere but GitHub Pages.
import { readFileSync, readdirSync, statSync } from "node:fs";
import { join } from "node:path";
import { execFileSync } from "node:child_process";

const DIST = "dist";
// The legacy working name, spelled so this file doesn't trip its own check.
const LEGACY = new RegExp("v" + "eil", "i");
const failures = [];
const fail = (msg) => failures.push(msg);

function walk(dir) {
  return readdirSync(dir).flatMap((name) => {
    const p = join(dir, name);
    return statSync(p).isDirectory() ? walk(p) : [p];
  });
}

let files;
try {
  files = walk(DIST);
} catch {
  console.error("No dist/ folder. Run `npm run build` first.");
  process.exit(1);
}

const info = JSON.parse(readFileSync(join(DIST, "build-info.json"), "utf8"));
if (info.bench !== false) fail("build-info.json: bench is not false (__NYMFORM_BENCH__ must be false in release builds).");
if (info.mode !== "production") fail("build-info.json: not a production build.");

// Strings that only bench code contains. "raw_bench" is the spec's marker string (T-P4).
const benchMarkers = ["raw_bench", "NYMFORM_BENCH_V1", "_nymform_scratch", "sendRaw", "Not a bench workbook", "Quorbel"];
const textFiles = files.filter((f) => /\.(js|html|json|css|txt|map)$/.test(f));
for (const f of textFiles) {
  const text = readFileSync(f, "utf8");
  for (const m of benchMarkers) if (text.includes(m)) fail(`${f}: contains bench marker "${m}".`);
  if (LEGACY.test(text)) fail(`${f}: contains the legacy name.`);
  if (/sk-or-v1-[0-9a-f]{16,}/i.test(text)) fail(`${f}: contains something shaped like an OpenRouter key.`);
  if (f.endsWith(".map")) fail(`${f}: source maps are not shipped in release builds.`);
}

// CSP: connect-src must be exactly the endpoint origin.
const origin = new URL(info.endpoint).origin;
for (const page of ["taskpane.html", "commands.html"]) {
  const html = readFileSync(join(DIST, page), "utf8");
  // The HTML minifier drops quotes around attribute values that don't need them.
  const m = /http-equiv=["']?Content-Security-Policy["']?\s+content=(?:"([^"]+)"|'([^']+)')/i.exec(html);
  if (!m) {
    fail(`${page}: no Content Security Policy meta tag.`);
    continue;
  }
  const policy = m[1] ?? m[2] ?? "";
  const connect = /connect-src ([^;]+)/.exec(policy);
  if (!connect || connect[1].trim() !== origin) fail(`${page}: connect-src is "${connect?.[1]}", expected "${origin}".`);
  if (/unsafe-eval/.test(policy)) fail(`${page}: CSP allows unsafe-eval.`);
  for (const directive of ["default-src 'self'", "object-src 'none'", "base-uri 'self'", "form-action 'none'"]) {
    if (!policy.split(";").map((d) => d.trim()).includes(directive)) fail(`${page}: CSP lacks ${directive}.`);
  }
}

// Release manifest: GitHub Pages only, no dev URLs, no legacy name.
const manifest = readFileSync("manifest.release.xml", "utf8");
if (/localhost/.test(manifest)) fail("manifest.release.xml: points at localhost.");
if (LEGACY.test(manifest)) fail("manifest.release.xml: contains the legacy name.");
const urls = [...manifest.matchAll(/DefaultValue="(https?:[^"]+)"/g)].map((m) => m[1]);
for (const u of urls) {
  if (!u.startsWith("https://skyspanmarketing.github.io/nymform-excel/") && !u.startsWith("https://github.com/skyspanmarketing/nymform-excel")) {
    fail(`manifest.release.xml: unexpected URL ${u}`);
  }
}

// Public-facing files: no legacy product name anywhere in the tracked tree. History is not rewritten.
let tracked = [];
try {
  tracked = execFileSync("git", ["ls-files", "-z"], { encoding: "utf8" }).split("\0").filter(Boolean);
} catch {
  fail("Couldn't list tracked files with git.");
}
for (const f of tracked) {
  if (/\.(png|xlsx|ico|jpg|gif)$/i.test(f)) continue;
  let text;
  try {
    text = readFileSync(f, "utf8");
  } catch {
    continue;
  }
  if (LEGACY.test(text) || LEGACY.test(f)) fail(`${f}: contains the legacy name.`);
}

if (failures.length) {
  console.error(`Release check failed (${failures.length}):`);
  for (const f of failures) console.error(`  - ${f}`);
  process.exit(1);
}
console.log(`Release check passed: ${textFiles.length} files, bench off, CSP connect-src ${origin}, commit ${info.commit}.`);
