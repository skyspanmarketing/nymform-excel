// Bench panel (spec §11). Bench builds only: App.tsx loads it with a dynamic import guarded by
// __NYMFORM_BENCH__, so release builds contain none of this (invariant 8, T-P4). Opened with
// ?bench=1 on a workbook that carries the bench marker.
//
// For each task on the open workbook it runs the raw, substituted and structure-only conditions.
// Substituted and structure-only runs use the production path (schema, stand-ins, payload, auditor,
// provider). Raw runs go through bench/rawSender.ts, which checks the marker and sends without the
// auditor, because sending raw values is the point of that baseline. Results are written to the
// scratch sheet _nymform_scratch, read back and compared with the expected values.
import { Button, Checkbox, Field, Input, Spinner, Textarea, makeStyles, tokens } from "@fluentui/react-components";
import { useEffect, useMemo, useRef, useState } from "react";
import { ENDPOINT } from "../src/config";
import { audit, buildAuditContext, type AuditedOutbound } from "../src/core/auditor";
import { check } from "../src/core/formulaGate";
import { buildCorrection, buildStructureOnly, buildSubstituted } from "../src/core/payload";
import { send as providerSend } from "../src/core/provider";
import { parseReply } from "../src/core/reply";
import { restoreFormula } from "../src/core/restore";
import { buildSheetContext, inferSchema } from "../src/core/schema";
import { defaultTreatment } from "../src/core/suggest";
import { createStandInMap, substituteUserText, type StandInMap } from "../src/core/transform";
import type { ColumnPolicy, ProdMode, RangeData, RangeSpec, SheetContext, TableInfo } from "../src/core/types";
import {
  listNames,
  listSheets,
  listTables,
  readBenchInfo,
  readRange,
  readScratch,
  writeScratch,
} from "../src/office/adapter";
import type { BenchTask } from "./compare";
import { leakCheck } from "./leakscan";
import privateValuesFile from "./workbooks/private-values.json";
import { BENCH_MARKER, buildRawBody, sendRaw } from "./rawSender";
import {
  CONDITIONS,
  DEFAULT_RUNS,
  resultsFileName,
  runBench,
  summarize,
  toJsonl,
  toMarkdown,
  type Condition,
  type Prepared,
  type RunRecord,
  type RunnerDeps,
} from "./runner";
import { TASKS } from "./tasks";

const useStyles = makeStyles({
  panel: {
    marginTop: "24px",
    paddingTop: "12px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    flexDirection: "column",
    gap: "10px",
  },
  row: { display: "flex", gap: "8px", flexWrap: "wrap", alignItems: "center" },
  mono: { fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200, whiteSpace: "pre-wrap" },
});

interface Handle {
  kind: "raw" | "prod";
  context: SheetContext;
  audited?: AuditedOutbound;
  map?: StandInMap;
  userContent?: string;
  mode?: ProdMode;
  spec?: RangeSpec;
  canaries: string[];
}

type WorkbookKey = keyof typeof privateValuesFile;

export function BenchPanel({ apiKey, model }: { apiKey: string; model: string }) {
  const s = useStyles();
  const [info, setInfo] = useState<{ marker: string | null; canaries: string[] } | null>(null);
  const [sheets, setSheets] = useState<string[]>([]);
  const [conditions, setConditions] = useState<Condition[]>([...CONDITIONS]);
  const [runs, setRuns] = useState(DEFAULT_RUNS);
  const [records, setRecords] = useState<RunRecord[]>([]);
  const [running, setRunning] = useState(false);
  const [message, setMessage] = useState<string | null>(null);
  const [exported, setExported] = useState<string | null>(null);
  const abort = useRef<AbortController | null>(null);
  const bodies = useRef<{ mode: string; body: string }[]>([]);

  useEffect(() => {
    readBenchInfo().then(setInfo, () => setInfo({ marker: null, canaries: [] }));
    listSheets().then(setSheets, () => setSheets([]));
  }, []);

  const tasks = useMemo(() => TASKS.filter((t) => sheets.includes(t.sheet)), [sheets]);
  const workbook = tasks[0]?.workbook.replace(/\.xlsx$/, "") as WorkbookKey | undefined;

  if (!info) return <Spinner size="tiny" label="Reading the bench marker…" />;
  if (info.marker !== BENCH_MARKER) {
    return (
      <section className={s.panel} aria-label="Benchmark">
        <strong>Benchmark</strong>
        <p>This workbook has no bench marker. Open one of the workbooks in bench/workbooks.</p>
      </section>
    );
  }

  async function start() {
    if (!apiKey.trim()) {
      setMessage("Add your API key in Setup first.");
      return;
    }
    setMessage(null);
    setExported(null);
    setRecords([]);
    bodies.current = [];
    const controller = new AbortController();
    abort.current = controller;
    setRunning(true);
    try {
      const deps = await makeDeps(apiKey, model, info!.canaries, controller.signal, bodies.current);
      const out = await runBench(tasks, conditions, runs, { ...deps, onRecord: (r) => setRecords((prev) => [...prev, r]) });
      setMessage(`Finished ${out.length} runs.`);
    } catch (e) {
      setMessage(e instanceof Error ? e.message : "The benchmark stopped on an unexpected problem.");
    } finally {
      setRunning(false);
      abort.current = null;
    }
  }

  const leaks = workbook
    ? leakCheck(bodies.current, privateValuesFile[workbook].values, privateValuesFile[workbook].canaries).protectedLeaks
    : undefined;
  const summary = records.length ? summarize(records, { leaks, model }) : null;

  function exportResults() {
    const name = resultsFileName(new Date().toISOString().slice(0, 10), model);
    const text = toJsonl(records);
    offerDownload(name, text);
    setExported(text);
  }

  function exportBodies() {
    const text = JSON.stringify(bodies.current, null, 2);
    offerDownload(`bench-requests-${workbook ?? "workbook"}.json`, text);
    setExported(text);
  }

  return (
    <section className={s.panel} aria-label="Benchmark">
      <strong>Benchmark</strong>
      <p>
        {tasks.length} tasks on this workbook. Each run selects the task's range, builds and checks the request, sends it,
        writes the formula on the scratch sheet and compares the result.
      </p>
      <div className={s.row} role="group" aria-label="Conditions">
        {CONDITIONS.map((c) => (
          <Checkbox
            key={c}
            label={c}
            checked={conditions.includes(c)}
            disabled={running}
            onChange={(_, d) => setConditions((prev) => (d.checked ? [...prev, c] : prev.filter((x) => x !== c)))}
          />
        ))}
      </div>
      <Field label="Runs per task and condition">
        <Input
          type="number"
          min={1}
          max={10}
          value={String(runs)}
          disabled={running}
          onChange={(_, d) => setRuns(Math.max(1, Math.min(10, Number(d.value) || 1)))}
        />
      </Field>
      <div className={s.row}>
        <Button appearance="primary" disabled={running || tasks.length === 0 || conditions.length === 0} onClick={start}>
          Run benchmark
        </Button>
        <Button disabled={!running} onClick={() => abort.current?.abort()}>
          Stop
        </Button>
        {running && <Spinner size="tiny" label={`${records.length} runs done`} />}
      </div>
      {message && <p role="status">{message}</p>}
      {summary && <div className={s.mono}>{toMarkdown(summary)}</div>}
      {records.length > 0 && (
        <div className={s.row}>
          <Button onClick={exportResults}>Export results (JSONL)</Button>
          <Button onClick={exportBodies}>Export requests for leakcheck</Button>
        </div>
      )}
      {exported && (
        <Field label="Exported text (if the download didn't start, copy it from here)">
          <Textarea readOnly value={exported} resize="vertical" rows={6} />
        </Field>
      )}
    </section>
  );
}

function offerDownload(name: string, text: string): void {
  try {
    const url = URL.createObjectURL(new Blob([text], { type: "application/json" }));
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
  } catch {
    // The textarea below the buttons always shows the text.
  }
}

/** Wires the runner to the adapter and the production modules. */
async function makeDeps(
  key: string,
  model: string,
  canaries: string[],
  signal: AbortSignal,
  bodies: { mode: string; body: string }[],
): Promise<RunnerDeps> {
  const workbook = {
    sheets: await listSheets(),
    tables: await listTables().catch((): TableInfo[] => []),
    names: await listNames().catch(() => undefined),
  };
  const ranges = new Map<string, RangeData>();
  const dataFor = async (task: BenchTask): Promise<RangeData> => {
    const k = `${task.sheet}!${task.range}`;
    let d = ranges.get(k);
    if (!d) {
      d = await readRange(k);
      ranges.set(k, d);
    }
    return d;
  };
  const handleOf = (p: Prepared) => p.handle as Handle;

  return {
    model,
    signal,
    async prepare(task, condition) {
      const data = await dataFor(task);
      const context = buildSheetContext(data, true, [], workbook);
      if (condition === "raw") {
        const info = await readBenchInfo();
        const raw = buildRawBody({ marker: info.marker, data, question: task.prompt, model, endpoint: ENDPOINT });
        bodies.push({ mode: "raw_bench", body: raw.body });
        return { body: raw.body, mode: raw.mode, bytes: raw.bytes, handle: { kind: "raw", context, canaries } satisfies Handle };
      }
      const columns = inferSchema(data, true);
      const privateLetters = new Set(task.private);
      const policies: ColumnPolicy[] = columns.map((c) => ({
        letter: c.letter,
        private: privateLetters.has(c.letter),
        treatment: privateLetters.has(c.letter) ? defaultTreatment(c.type) : "as_is",
      }));
      const spec: RangeSpec = { data, hasHeaders: true, columns, policies };
      const map = createStandInMap([data]);
      const question = substituteUserText(task.prompt, [spec], map);
      const input = { model, specs: [spec], sheetContext: context, question, history: [], map };
      const built = condition === "substituted" ? buildSubstituted(input) : buildStructureOnly(input);
      const outcome = audit(built.outbound, buildAuditContext([spec], map, { model, canaries }));
      if (!outcome.ok) {
        const reasons = [...new Set([...outcome.result.hits.map((h) => h.variant), ...(outcome.result.structural.length ? ["structural"] : [])])];
        return { blocked: true, reasons: reasons.length ? reasons : ["error"], bytes: built.outbound.bytes };
      }
      bodies.push({ mode: condition, body: outcome.outbound.body });
      const handle: Handle = { kind: "prod", context, audited: outcome.outbound, map, userContent: built.userContent, mode: condition, spec, canaries };
      return { body: outcome.outbound.body, mode: condition, bytes: outcome.outbound.bytes, handle };
    },
    async send(prepared, task, _condition, sig) {
      const h = handleOf(prepared);
      if (h.kind === "raw") {
        const info = await readBenchInfo();
        const res = await sendRaw({ marker: info.marker, data: await dataFor(task), question: task.prompt, model, endpoint: ENDPOINT, key, signal: sig });
        if (res.status < 200 || res.status >= 300) return { ok: false, message: `The provider answered ${res.status}.`, latencyMs: res.latencyMs, status: res.status };
        return res.tokens ? { ok: true, raw: res.raw, latencyMs: res.latencyMs, tokens: res.tokens } : { ok: true, raw: res.raw, latencyMs: res.latencyMs };
      }
      const res = await providerSend(h.audited!, key, sig);
      if (!res.ok) return { ok: false, message: res.message, latencyMs: res.latencyMs, ...(res.status !== undefined ? { status: res.status } : {}) };
      return res.tokens ? { ok: true, raw: res.raw, latencyMs: res.latencyMs, tokens: res.tokens } : { ok: true, raw: res.raw, latencyMs: res.latencyMs };
    },
    parse(raw, condition) {
      const r = parseReply(raw, condition === "structure_only" ? "structure_only" : "substituted");
      return r.ok ? { ok: true, reply: r.reply } : { ok: false, error: r.error, content: r.content };
    },
    gate(formula, _task, _phase, prepared) {
      return check(formula, handleOf(prepared).context);
    },
    restore(formula, prepared) {
      const h = handleOf(prepared);
      return h.kind === "raw" || !h.map ? formula : restoreFormula(formula, h.map);
    },
    async writeAndRead(formula, placement) {
      const address = await writeScratch(formula, placement.fillDown, { row: placement.row, lastRow: placement.lastDataRow });
      return readScratch(address);
    },
    correct(prepared, invalidReply) {
      const h = handleOf(prepared);
      // Raw runs have no correction turn: the raw baseline sender builds only the first request.
      // Throwing makes the runner record the run as a parse failure.
      if (h.kind === "raw" || !h.spec || !h.map || !h.mode || h.userContent === undefined) {
        throw new Error("the raw baseline has no correction retry");
      }
      const built = buildCorrection({ model, mode: h.mode, history: [], userContent: h.userContent, invalidReply });
      const outcome = audit(built.outbound, buildAuditContext([h.spec], h.map, { model, canaries: h.canaries }));
      if (!outcome.ok) return { blocked: true, reasons: ["correction-blocked"], bytes: built.outbound.bytes };
      bodies.push({ mode: h.mode, body: outcome.outbound.body });
      return { body: outcome.outbound.body, mode: h.mode, bytes: outcome.outbound.bytes, handle: { ...h, audited: outcome.outbound } satisfies Handle };
    },
  };
}
