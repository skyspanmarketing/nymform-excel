// Log (spec §7.13, screen 6): past requests with their exact bodies, Export log and Export
// evaluation report. Both exports happen only on a click; nothing is transmitted. For a blocked
// request it also shows what the check found and where (LogLocal): that text is kept in the pane
// only, beside the entry, and neither export includes it.
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Button,
  Textarea,
  mergeClasses,
} from "@fluentui/react-components";
import { ArrowDownload20Regular } from "@fluentui/react-icons";
import { useState } from "react";
import type { LocalMatch, LogLocal } from "../../core/log";
import type { LogEntry } from "../../core/types";
import { copyToClipboard, downloadText, fileStamp } from "../browser";
import { COPY, MODE_LABEL } from "../copy";
import { auditStatus, type Flow, type FlowState } from "../flow";
import { usePattern, useUi } from "../styles";
import { StatusLine } from "./StatusLine";

interface Exported {
  label: string;
  filename: string;
  json: string;
  downloaded: boolean;
}

export function LogScreen({ flow, state }: { flow: Flow; state: FlowState }) {
  const ui = useUi();
  const [exported, setExported] = useState<Exported | null>(null);
  const items = [...flow.logItems()].reverse();

  const doExport = (label: string, prefix: string, json: string) => {
    const filename = `${prefix}-${fileStamp()}.json`;
    setExported({ label, filename, json, downloaded: downloadText(filename, json) });
  };

  return (
    <section className={ui.stack} aria-labelledby="log-heading">
      <h2 id="log-heading" className={ui.heading}>
        Log
      </h2>
      <p className={ui.muted}>
        Requests from this pane, kept in memory until you close it. The log never holds your key or the stand-in map.
      </p>
      <div className={ui.row}>
        <Button
          icon={<ArrowDownload20Regular />}
          disabled={state.logSize === 0}
          onClick={() => doExport("Log", "nymform-log", flow.exportLog())}
        >
          Export log
        </Button>
        <Button
          icon={<ArrowDownload20Regular />}
          onClick={() => doExport("Evaluation report", "nymform-evaluation-report", flow.exportReport())}
        >
          Export evaluation report
        </Button>
      </div>
      <p className={ui.muted}>
        The evaluation report holds counts, sizes, timings and check results only: no values, headers, questions or
        formulas.
      </p>
      {exported && <ExportResult key={exported.filename} exported={exported} />}
      {items.length === 0 ? (
        <p className={ui.muted}>No requests yet.</p>
      ) : (
        <ol className={mergeClasses(ui.stack, ui.plainList)} aria-label="Past requests, newest first">
          {items.map((item, i) => (
            <LogItem key={`${item.entry.at}-${items.length - i}`} entry={item.entry} local={item.local} />
          ))}
        </ol>
      )}
    </section>
  );
}

function ExportResult({ exported }: { exported: Exported }) {
  const ui = useUi();
  const [show, setShow] = useState(!exported.downloaded);
  const [copied, setCopied] = useState<boolean | null>(null);
  return (
    <div className={ui.stack}>
      <StatusLine
        intent={exported.downloaded ? "success" : "warning"}
        title={exported.downloaded ? `Exported ${exported.filename}.` : "Couldn't save a file here. Copy the text below instead."}
        live
      />
      {exported.downloaded && !show && (
        <div className={ui.row}>
          <Button appearance="subtle" size="small" onClick={() => setShow(true)}>
            No file? Show it as text
          </Button>
        </div>
      )}
      {show && (
        <>
          <Textarea
            readOnly
            value={exported.json}
            aria-label={`${exported.label} as text`}
            resize="vertical"
            rows={8}
            textarea={{ className: ui.exportBox }}
          />
          <div className={ui.row}>
            <Button
              size="small"
              onClick={() => {
                void copyToClipboard(exported.json).then(setCopied);
              }}
            >
              Copy
            </Button>
            <span role="status" aria-live="polite" className={ui.muted}>
              {copied === null ? "" : copied ? COPY.copied : COPY.copyFailed}
            </span>
          </div>
        </>
      )}
    </div>
  );
}

function timeOf(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
}

/** Why the matched text counts: the column's renamed header, another column's header, or a private value. */
function matchSource(m: LocalMatch): string {
  if (m.origin === "alias") return COPY.logMatchAlias(m.column);
  if (m.inHeader !== undefined && m.inHeader !== m.column) {
    const letter = m.inHeader.slice(m.inHeader.lastIndexOf("!") + 1);
    return COPY.logMatchHeader(m.inHeader, letter, m.text, m.variant, m.column);
  }
  return COPY.logMatchValue(m.variant, m.column);
}

function Matches({ matches }: { matches: readonly LocalMatch[] }) {
  const ui = useUi();
  return (
    <div className={ui.tight}>
      <ul className={mergeClasses(ui.tight, ui.plainList)} aria-label="What the check found">
        {matches.map((m, i) => (
          <li key={i} className={ui.tight}>
            <span className={ui.text}>
              Found “<strong>{m.text}</strong>” in {m.where}
            </span>
            <span className={ui.code} aria-label="Text around it">
              {m.before}
              <strong>{m.text}</strong>
              {m.after}
            </span>
            <span className={ui.muted}>{matchSource(m)}</span>
          </li>
        ))}
      </ul>
      <span className={ui.muted}>{COPY.logMatchNote}</span>
    </div>
  );
}

function LogItem({ entry, local }: { entry: LogEntry; local: LogLocal | undefined }) {
  const ui = useUi();
  const p = usePattern();
  const matches = entry.audit.ok ? [] : (local?.matches ?? []);
  const audit = entry.audit.ok ? "passed" : "blocked";
  const gate = entry.gate ? (entry.gate.ok ? "passed" : "blocked") : "none";
  const mode = entry.mode === "structure_only" || entry.mode === "substituted" ? MODE_LABEL[entry.mode] : entry.mode;
  return (
    <li className={p.card}>
      <span className={ui.text}>
        <strong>{timeOf(entry.at)}</strong> · {mode} · {entry.bytes.toLocaleString("en-US")} bytes
      </span>
      <span className={ui.muted}>
        Check: {audit} · Formula check: {gate} · Provider: {entry.providerName ?? "not reported"}
        {entry.inserted ? " · Inserted" : ""}
      </span>
      {!entry.audit.ok && <span className={ui.muted}>{auditStatus(entry.audit, matches).headline}</span>}
      {matches.length > 0 && <Matches matches={matches} />}
      {entry.error && entry.audit.ok && <span className={ui.muted}>{entry.error}</span>}
      <Accordion collapsible multiple>
        <AccordionItem value="body">
          <AccordionHeader size="small">Exact request</AccordionHeader>
          <AccordionPanel>
            <pre className={ui.code} tabIndex={0} aria-label="Exact request">
              {entry.body !== "" ? entry.body : COPY.notSentBody}
            </pre>
          </AccordionPanel>
        </AccordionItem>
        {entry.replyRaw !== undefined && (
          <AccordionItem value="reply">
            <AccordionHeader size="small">Reply</AccordionHeader>
            <AccordionPanel>
              <pre className={ui.code} tabIndex={0} aria-label="Reply">
                {entry.replyRaw}
              </pre>
            </AccordionPanel>
          </AccordionItem>
        )}
      </Accordion>
    </li>
  );
}
