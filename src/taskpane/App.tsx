// The task pane. Screens in flow order (spec §7.13): Setup, Columns, Ask -> What gets sent ->
// Result, and Log. All orchestration lives in flow.ts; this file holds UI state only. The API key
// lives only in this component's memory (invariant 11): never stored, never logged.
import {
  Button,
  FluentProvider,
  Spinner,
  Tab,
  TabList,
  mergeClasses,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import { useCallback, useEffect, useRef, useState, useSyncExternalStore, type ComponentType, type ReactNode } from "react";
import { COMMIT, ENDPOINT, VERSION } from "../config";
import { DEFAULT_ROW_CAP } from "../core/transform";
import type { ProdMode } from "../core/types";
import { AskScreen } from "./components/AskScreen";
import { ColumnsScreen, type Notice } from "./components/ColumnsScreen";
import { ErrorBoundary } from "./components/ErrorBoundary";
import { LogScreen } from "./components/LogScreen";
import { PreviewScreen } from "./components/PreviewScreen";
import { ResultScreen } from "./components/ResultScreen";
import { SetupScreen } from "./components/SetupScreen";
import { StatusLine } from "./components/StatusLine";
import { COPY } from "./copy";
import { Flow, type ActionResult, type FlowResult, type Prepared } from "./flow";
import { isDarkTheme, type OfficeHost } from "./host";
import { useAppStyles, usePageStyles } from "./styles";

declare const __NYMFORM_BENCH__: boolean;

/** How long to wait for Office.js before saying the pane isn't inside Excel. */
const READY_TIMEOUT_MS = 20_000;

type Tab = "setup" | "columns" | "ask" | "log";
type Stage = { name: "ask" } | { name: "preview"; prepared: Prepared } | { name: "result"; result: FlowResult };

export function App({ host }: { host: OfficeHost }) {
  const [phase, setPhase] = useState<"loading" | "ready" | "outside">("loading");
  const [dark, setDark] = useState(false);

  useEffect(() => {
    let cancelled = false;
    const timer = setTimeout(() => {
      if (!cancelled) setPhase((p) => (p === "loading" ? "outside" : p));
    }, READY_TIMEOUT_MS);
    host.ready().then(
      () => {
        if (cancelled) return;
        clearTimeout(timer);
        setDark(isDarkTheme(safeTheme(host)));
        setPhase("ready");
      },
      () => {
        if (cancelled) return;
        clearTimeout(timer);
        setPhase("outside");
      },
    );
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [host]);

  return (
    <FluentProvider theme={dark ? webDarkTheme : webLightTheme}>
      <Frame dark={dark}>
        {phase === "loading" && <Centered><Spinner size="small" label="Starting…" /></Centered>}
        {phase === "outside" && (
          <Centered>
            <StatusLine intent="info" title={COPY.notInExcel} />
          </Centered>
        )}
        {phase === "ready" && (
          <ErrorBoundary
            fallback={(retry) => (
              <Centered>
                <StatusLine intent="error" title={COPY.crashed} />
                <Button onClick={retry}>Try again</Button>
              </Centered>
            )}
          >
            <Workspace host={host} />
          </ErrorBoundary>
        )}
      </Frame>
    </FluentProvider>
  );
}

function safeTheme(host: OfficeHost) {
  try {
    return host.theme();
  } catch {
    return null;
  }
}

function Frame({ dark, children }: { dark: boolean; children: ReactNode }) {
  usePageStyles();
  const s = useAppStyles();
  return (
    <div className={mergeClasses(s.root, dark ? s.dark : s.light)}>
      {children}
      <footer className={s.footer}>{COPY.footer(VERSION, COMMIT)}</footer>
    </div>
  );
}

function Centered({ children }: { children: ReactNode }) {
  const s = useAppStyles();
  return <div className={s.centered}>{children}</div>;
}

function Workspace({ host }: { host: OfficeHost }) {
  const s = useAppStyles();
  const [flow] = useState(() => new Flow({ adapter: host }));
  const state = useSyncExternalStore(flow.subscribe, flow.getState);

  const [tab, setTab] = useState<Tab>("setup");
  // Memory only (invariant 11). Never put in storage, the log, or anything that is exported.
  const [apiKey, setApiKey] = useState("");
  const [model, setModel] = useState(flow.model);
  const [BenchPanel, setBenchPanel] = useState<ComponentType<{ apiKey: string; model: string }> | null>(null);
  useEffect(() => {
    if (__NYMFORM_BENCH__) {
      if (new URLSearchParams(window.location.search).get("bench") === "1") {
        import("../../bench/BenchPanel").then(
          (m) => setBenchPanel(() => m.BenchPanel),
          () => undefined,
        );
      }
    }
  }, []);
  const [mode, setMode] = useState<ProdMode>("structure_only");
  const [rowCap, setRowCap] = useState(DEFAULT_ROW_CAP);
  const [question, setQuestion] = useState("");
  // Stand-ins chosen for parts of private values typed in the question, for this question only.
  const [choices, setChoices] = useState<ReadonlyMap<string, readonly string[]>>(new Map());
  // True after Show in sheet moved Excel's selection, until it is put back.
  const shownCells = useRef(false);
  const [stage, setStage] = useState<Stage>({ name: "ask" });
  const [askError, setAskError] = useState<string | null>(null);
  const [sendError, setSendError] = useState<string | null>(null);
  const [sending, setSending] = useState(false);
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState<Notice | null>(null);
  const abortRef = useRef<AbortController | null>(null);
  const started = useRef(false);

  const run = useCallback(async (action: () => Promise<ActionResult>, success?: string): Promise<ActionResult> => {
    setBusy(true);
    setNotice(null);
    let r: ActionResult;
    try {
      r = await action();
    } catch {
      r = { ok: false, message: COPY.unexpected };
    }
    setBusy(false);
    if (!r.ok) setNotice({ intent: "error", text: r.message });
    else if (r.message) setNotice({ intent: "info", text: r.message });
    else if (success) setNotice({ intent: "success", text: success });
    return r;
  }, []);

  const refresh = useCallback(() => void run(() => flow.refresh()), [run, flow]);

  // Read the selection once when the pane opens. Reading is local; nothing is sent.
  useEffect(() => {
    if (started.current) return;
    started.current = true;
    refresh();
  }, [refresh]);

  // Abort a request in flight if the pane goes away.
  useEffect(() => () => abortRef.current?.abort(), []);

  const onModel = (m: string) => {
    setModel(m);
    flow.setModel(m);
  };

  const onQuestion = (q: string) => {
    setQuestion(q);
    setChoices(new Map());
  };

  const restoreSelection = () => {
    if (!shownCells.current) return;
    shownCells.current = false;
    void flow.restoreSelection();
  };

  const preview = (text: string, picks: ReadonlyMap<string, readonly string[]> = choices) => {
    setAskError(null);
    setSendError(null);
    const r = flow.prepare(text, mode, rowCap, picks);
    if (!r.ok) {
      setAskError(r.message);
      setStage({ name: "ask" });
      return;
    }
    setStage({ name: "preview", prepared: r.prepared });
  };

  const send = async (prepared: Prepared) => {
    setSendError(null);
    setSending(true);
    const controller = new AbortController();
    abortRef.current = controller;
    try {
      const r = await flow.send(prepared, apiKey, controller.signal);
      // An unreadable reply comes back as the one correction request: preview it like any other.
      if (!r.ok && r.correction) setStage({ name: "preview", prepared: r.correction });
      else if (!r.ok) setSendError(r.detail ? `${r.message} ${r.detail}` : r.message);
      else setStage({ name: "result", result: r.result });
    } catch {
      setSendError(COPY.unexpected);
    } finally {
      abortRef.current = null;
      setSending(false);
    }
  };

  let screen: ReactNode;
  if (tab === "setup") {
    screen = (
      <SetupScreen
        apiKey={apiKey}
        onApiKey={setApiKey}
        model={model}
        onModel={onModel}
        endpoint={ENDPOINT}
        host={flow.host}
        onContinue={() => setTab("columns")}
      />
    );
  } else if (tab === "columns") {
    screen = (
      <ColumnsScreen
        flow={flow}
        state={state}
        sampleRows={mode === "substituted"}
        busy={busy}
        notice={notice}
        onRefresh={refresh}
        run={run}
      />
    );
  } else if (tab === "log") {
    screen = <LogScreen flow={flow} state={state} />;
  } else if (stage.name === "preview") {
    screen = (
      <PreviewScreen
        key={stage.prepared.id}
        prepared={stage.prepared}
        host={flow.host}
        hasKey={apiKey.trim() !== ""}
        current={flow.isCurrent(stage.prepared)}
        sending={sending}
        error={sendError}
        onSend={() => void send(stage.prepared)}
        onCancel={() => abortRef.current?.abort()}
        onBack={() => {
          restoreSelection();
          setSendError(null);
          setStage({ name: "ask" });
        }}
        canShowCells={flow.canShowCells()}
        onChoose={(key, tokens) => {
          restoreSelection();
          const next = new Map(choices).set(key, tokens);
          setChoices(next);
          preview(question, next);
        }}
        onReset={(key) => {
          const next = new Map(choices);
          next.delete(key);
          setChoices(next);
          preview(question, next);
        }}
        onShowCells={(sheet, cells) => {
          shownCells.current = true;
          void run(() => flow.showCells(sheet, cells));
        }}
      />
    );
  } else if (stage.name === "result") {
    screen = (
      <ResultScreen
        key={stage.result.id}
        flow={flow}
        result={stage.result}
        onAnswer={(answer) => preview(answer, new Map())}
        onAskAnother={() => {
          onQuestion("");
          setStage({ name: "ask" });
        }}
      />
    );
  } else {
    screen = (
      <AskScreen
        flow={flow}
        state={state}
        mode={mode}
        onMode={setMode}
        rowCap={rowCap}
        onRowCap={setRowCap}
        question={question}
        onQuestion={onQuestion}
        error={askError}
        onPreview={() => preview(question)}
      />
    );
  }

  // ---- Bench panel ----------------------------------------------------------------------------
  // Bench builds only (invariant 8): __NYMFORM_BENCH__ is a compile-time constant, so in a release
  // build this branch and the dynamic import inside it are removed and no bench chunk exists (T-P4).
  const benchPanel: ReactNode = BenchPanel ? <BenchPanel apiKey={apiKey} model={model} /> : null;

  return (
    <>
      <header className={s.header}>
        <TabList selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as Tab)} size="small" aria-label="Nymform screens">
          <Tab value="setup">Setup</Tab>
          <Tab value="columns">Columns</Tab>
          <Tab value="ask">Ask</Tab>
          <Tab value="log">Log</Tab>
        </TabList>
      </header>
      <main className={s.main}>
        {screen}
        {benchPanel}
      </main>
    </>
  );
}
