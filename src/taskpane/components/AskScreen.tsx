// Ask (spec §7.13, screen 3): mode switch, row count for sample rows, the question and Preview.
import { Button, Field, Radio, RadioGroup, SpinButton, Textarea } from "@fluentui/react-components";
import { DEFAULT_ROW_CAP, MAX_ROW_CAP, clampRowCap } from "../../core/transform";
import type { ProdMode } from "../../core/types";
import type { Flow, FlowState } from "../flow";
import { useUi } from "../styles";
import { StatusLine } from "./StatusLine";
import { COPY } from "../copy";

interface Props {
  flow: Flow;
  state: FlowState;
  mode: ProdMode;
  onMode: (mode: ProdMode) => void;
  rowCap: number;
  onRowCap: (n: number) => void;
  question: string;
  onQuestion: (q: string) => void;
  error: string | null;
  onPreview: () => void;
}

export function AskScreen(props: Props) {
  const ui = useUi();
  const { flow, state, mode, onMode, rowCap, onRowCap, question, onQuestion, error, onPreview } = props;
  const exchanges = Math.floor(state.historyMessages / 2);
  const hasSelection = state.ranges.length > 0;
  return (
    <section className={ui.stack} aria-labelledby="ask-heading">
      <h2 id="ask-heading" className={ui.heading}>
        Ask
      </h2>
      {!hasSelection && <StatusLine intent="info" title={COPY.noSelection} />}
      {state.stale && <StatusLine intent="warning" title={COPY.dataChanged} />}
      <Field label="What to send">
        <RadioGroup value={mode} onChange={(_, d) => onMode(d.value === "substituted" ? "substituted" : "structure_only")}>
          <Radio
            value="structure_only"
            label={
              <span className={ui.tight}>
                <span>Structure only</span>
                <span className={ui.muted}>Headers, types and counts. No cell values.</span>
              </span>
            }
          />
          <Radio
            value="substituted"
            label={
              <span className={ui.tight}>
                <span>Include sample rows</span>
                <span className={ui.muted}>The first rows, with private values replaced by stand-ins.</span>
              </span>
            }
          />
        </RadioGroup>
      </Field>
      {mode === "substituted" && (
        <Field label="Rows to include" hint={`From the top of the selection. Up to ${MAX_ROW_CAP}.`}>
          <SpinButton
            min={1}
            max={MAX_ROW_CAP}
            value={rowCap}
            onChange={(_, d) => {
              const raw = d.value ?? (d.displayValue !== undefined ? Number(d.displayValue) : NaN);
              onRowCap(Number.isFinite(raw) ? clampRowCap(raw as number) : DEFAULT_ROW_CAP);
            }}
          />
        </Field>
      )}
      <Field label="Question">
        {/* No browser spell check or writing aids: some send the text to a cloud service. */}
        <Textarea
          value={question}
          onChange={(_, d) => onQuestion(d.value)}
          placeholder="For example: total Amount per Region"
          resize="vertical"
          rows={3}
          spellCheck={false}
          autoCapitalize="off"
          data-gramm="false"
          // Textarea doesn't pass these through from its own props; its textarea slot does.
          textarea={{ autoCorrect: "off", writingsuggestions: "false" }}
          onKeyDown={(e) => {
            if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
              e.preventDefault();
              onPreview();
            }
          }}
        />
      </Field>
      {exchanges > 0 && (
        <div className={ui.spread}>
          <span className={ui.muted}>
            Follow-up: your last {exchanges === 1 ? "exchange goes" : `${exchanges} exchanges go`} with the question, with
            stand-ins only.
          </span>
          <Button size="small" appearance="subtle" onClick={() => flow.clearHistory()}>
            Clear conversation
          </Button>
        </div>
      )}
      {error && <StatusLine intent="error" title={error} live />}
      <div className={ui.row}>
        <Button appearance="primary" onClick={onPreview} disabled={!hasSelection}>
          Preview
        </Button>
        <span className={ui.muted}>Nothing is sent until you choose Send.</span>
      </div>
    </section>
  );
}
