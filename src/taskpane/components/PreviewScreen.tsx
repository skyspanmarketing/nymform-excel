// What gets sent (spec §7.13, screen 4): a readable view of the exact request, its size, the
// literal string on request, and the audit status. Always shown before sending.
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Button,
  Spinner,
  Switch,
} from "@fluentui/react-components";
import { useState } from "react";
import { COPY, MODE_LABEL } from "../copy";
import type { Prepared, RequestView } from "../flow";
import { usePattern, useUi } from "../styles";
import { useFocusOnMount } from "./focus";
import { PartialPicker } from "./PartialPicker";
import { StatusLine } from "./StatusLine";
import { TokenText } from "./TokenText";

interface Props {
  prepared: Prepared;
  host: string;
  hasKey: boolean;
  current: boolean;
  sending: boolean;
  error: string | null;
  onSend: () => void;
  onCancel: () => void;
  onBack: () => void;
  /** Show in sheet is offered only when the host can select cells. */
  canShowCells: boolean;
  onChoose: (key: string, tokens: readonly string[]) => void;
  onReset: (key: string) => void;
  onShowCells: (sheet: string, cells: readonly string[]) => void;
}

export function PreviewScreen({ prepared, host, hasKey, current, sending, error, onSend, onCancel, onBack, canShowCells, onChoose, onReset, onShowCells }: Props) {
  const ui = useUi();
  const heading = useFocusOnMount<HTMLHeadingElement>();
  const [exact, setExact] = useState(false);
  const { status, view } = prepared;
  const canSend = status.ok && hasKey && current && !sending;
  // A part of a private value waiting for a choice: the picker comes first, since choosing fixes the block.
  const pending = prepared.partial.pending.length > 0;
  const picker = (
    <PartialPicker
      partial={prepared.partial}
      canShow={canShowCells}
      busy={sending}
      onChoose={onChoose}
      onReset={onReset}
      onShow={onShowCells}
    />
  );
  return (
    <section className={ui.stack} aria-labelledby="sent-heading">
      <h2 id="sent-heading" ref={heading} tabIndex={-1} className={ui.heading}>
        What gets sent
      </h2>
      {prepared.correction && <StatusLine intent="info" title={COPY.correctionHeading}>{COPY.correctionReady}</StatusLine>}
      {pending && picker}
      <StatusLine intent={status.ok ? "success" : "error"} title={status.headline} live>
        {status.detail && <span>{status.detail}</span>}
        {status.found && <span>{status.found}</span>}
        {pending ? <span>{COPY.pickPartial}</span> : status.action && <span>{status.action}</span>}
      </StatusLine>
      {!pending && picker}
      {!current && <StatusLine intent="warning" title={COPY.stalePreview} />}
      <p className={ui.text}>
        <strong>{view.bytes.toLocaleString("en-US")} bytes</strong> to {host} · {MODE_LABEL[prepared.mode]}
      </p>
      <p className={ui.muted}>{COPY.formatted(host)}</p>
      <Switch label="Show exact request" checked={exact} onChange={(_, d) => setExact(d.checked)} />
      {exact ? (
        <pre className={ui.code} tabIndex={0} aria-label="Exact request">
          {view.exact}
        </pre>
      ) : (
        <FormattedRequest view={view} />
      )}
      {error && <StatusLine intent="error" title={error} live />}
      {!hasKey && <p className={ui.muted}>{COPY.keyMissing}</p>}
      <div className={ui.row}>
        <Button appearance="primary" disabled={!canSend} onClick={onSend}>
          Send
        </Button>
        <Button onClick={onBack} disabled={sending}>
          Back
        </Button>
        {sending && (
          <>
            <Spinner size="tiny" label="Sending…" labelPosition="after" />
            <Button appearance="subtle" onClick={onCancel}>
              Cancel
            </Button>
          </>
        )}
      </div>
    </section>
  );
}

function roleLabel(role: string, description: boolean): string {
  if (role === "assistant") return "Model's reply";
  if (role === "system") return "System";
  return description ? "Your earlier request" : "Your message";
}

function FormattedRequest({ view }: { view: RequestView }) {
  const ui = useUi();
  const p = usePattern();
  return (
    <div className={ui.stack}>
      <div className={ui.tight}>
        <span className={ui.subheading}>Model</span>
        <span className={ui.text}>{view.model}</span>
        {view.params.length > 0 && (
          <span className={ui.muted}>{view.params.map((x) => `${x.key}: ${x.value}`).join(" · ")}</span>
        )}
      </div>
      <Accordion collapsible multiple>
        <AccordionItem value="system">
          <AccordionHeader size="small">System prompt</AccordionHeader>
          <AccordionPanel>
            <pre className={ui.code} tabIndex={0} aria-label="System prompt">
              {view.system}
            </pre>
          </AccordionPanel>
        </AccordionItem>
        {view.earlier.length > 0 && (
          <AccordionItem value="earlier">
            <AccordionHeader size="small">Earlier messages ({view.earlier.length})</AccordionHeader>
            <AccordionPanel>
              <div className={ui.stack}>
                {view.earlier.map((m, i) => (
                  <div key={i} className={ui.tight}>
                    <span className={ui.muted}>{roleLabel(m.role, m.description)}</span>
                    <pre className={ui.code} tabIndex={0} aria-label={roleLabel(m.role, m.description)}>
                      <TokenText segments={m.segments} lines />
                    </pre>
                  </div>
                ))}
              </div>
            </AccordionPanel>
          </AccordionItem>
        )}
      </Accordion>
      {view.current && (
        <div className={ui.tight}>
          <span className={ui.subheading}>Your request</span>
          <pre className={ui.code} tabIndex={0} aria-label="Your request, formatted">
            <TokenText segments={view.current.segments} lines />
          </pre>
        </div>
      )}
      <p className={ui.muted}>
        <span className={p.swatch} aria-hidden="true" />
        Marked labels are stand-ins for private values. The originals stay in Excel.
      </p>
    </div>
  );
}
