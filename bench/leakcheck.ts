// Leak check (spec §11): re-scans exported request logs against the generator's own list of private
// values and canaries. Deliberately independent of src/core/auditor.ts, with its own simple scan,
// so a bug in the auditor can't hide itself here. Output names value indexes, never values,
// unless --verbose is given.
//
//   npx tsx bench/leakcheck.ts <exported-log.json>... [--workbook orders] [--values path] [--verbose]
//
// Exits 1 when any substituted or structure_only request contains a private value or canary.
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { PROTECTED_MODES, extractEntries, leakCheck, type PrivateValuesFile } from "./leakscan";

export * from "./leakscan";

interface CliOptions {
  files: string[];
  workbook?: string;
  valuesPath: string;
  verbose: boolean;
}

function parseArgs(argv: readonly string[], defaultValues: string): CliOptions {
  const opts: CliOptions = { files: [], valuesPath: defaultValues, verbose: false };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i]!;
    if (a === "--verbose" || a === "-v") opts.verbose = true;
    else if (a === "--workbook") opts.workbook = argv[++i];
    else if (a.startsWith("--workbook=")) opts.workbook = a.slice("--workbook=".length);
    else if (a === "--values") opts.valuesPath = argv[++i] ?? defaultValues;
    else if (a.startsWith("--values=")) opts.valuesPath = a.slice("--values=".length);
    else if (a.startsWith("--")) throw new Error(`Unknown option ${a}.`);
    else opts.files.push(a);
  }
  if (opts.files.length === 0) throw new Error("Name at least one exported log file.");
  return opts;
}

/** Runs the CLI; returns the exit code. `print` receives each output line. */
export function runCli(argv: readonly string[], print: (line: string) => void, defaultValues: string): number {
  let opts: CliOptions;
  try {
    opts = parseArgs(argv, defaultValues);
  } catch (e) {
    print(e instanceof Error ? e.message : String(e));
    print("Usage: npx tsx bench/leakcheck.ts <exported-log.json>... [--workbook orders] [--values path] [--verbose]");
    return 2;
  }
  const file = JSON.parse(readFileSync(opts.valuesPath, "utf8")) as PrivateValuesFile;
  const names = opts.workbook ? [opts.workbook] : Object.keys(file);
  const values: string[] = [];
  const canaries: string[] = [];
  const valueOwner: string[] = [];
  const canaryOwner: string[] = [];
  for (const name of names) {
    const wb = file[name];
    if (!wb) {
      print(`No workbook "${name}" in ${opts.valuesPath}. Known: ${Object.keys(file).join(", ")}.`);
      return 2;
    }
    wb.values.forEach((v, i) => {
      values.push(v);
      valueOwner.push(`${name} value #${i}`);
    });
    wb.canaries.forEach((c, i) => {
      canaries.push(c);
      canaryOwner.push(`${name} canary #${i}`);
    });
  }

  const all: { mode: string; body: string }[] = [];
  const source: string[] = [];
  for (const f of opts.files) {
    let parsed: ReturnType<typeof extractEntries>;
    try {
      parsed = extractEntries(JSON.parse(readFileSync(f, "utf8")));
    } catch (e) {
      print(`${f}: ${e instanceof Error ? e.message : String(e)}`);
      return 2;
    }
    parsed.entries.forEach((e, i) => {
      all.push(e);
      source.push(`${f} entry ${i}`);
    });
    if (parsed.skipped > 0) print(`${f}: skipped ${parsed.skipped} entries without a mode and body.`);
  }

  const report = leakCheck(all, values, canaries);
  print(`Scanned ${all.length} requests against ${values.length} private values and ${canaries.length} canaries (${names.join(", ")}).`);
  print("mode              requests  leaking  leaks");
  for (const [mode, c] of Object.entries(report.byMode).sort()) {
    print(`${mode.padEnd(18)}${String(c.requests).padStart(8)}${String(c.leakingRequests).padStart(9)}${String(c.leaks).padStart(7)}`);
  }
  for (const l of report.leaks) {
    if (!PROTECTED_MODES.includes(l.mode) && !opts.verbose) continue;
    const what = l.kind === "canary" ? canaryOwner[l.index] : valueOwner[l.index];
    const shown = opts.verbose ? ` ${JSON.stringify(l.kind === "canary" ? canaries[l.index] : values[l.index])}` : "";
    print(`  leak: ${source[l.entry]} (${l.mode}): ${what} (${l.variant})${shown}`);
  }
  if (report.protectedLeaks > 0) {
    print(`FAIL: ${report.protectedLeaks} leaks in substituted or structure_only requests.`);
    return 1;
  }
  print("OK: no private value or canary in any substituted or structure_only request.");
  return 0;
}

if (process.argv[1] && /leakcheck\.ts$/.test(process.argv[1])) {
  const defaultValues = resolve(process.argv[1], "..", "workbooks", "private-values.json");
  process.exit(runCli(process.argv.slice(2), (line) => console.log(line), defaultValues));
}
