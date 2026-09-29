// Columns (spec §7.13, screen 2): the range in use and how it came from the selection (§7.1), its
// columns with "Keep private", the header sent (editable alias), inferred type, treatment
// (sample-rows mode only), an optional note, and lookup ranges. Suggestions are defaults; the
// checkbox is the decision.
import { Button, Checkbox, Field, Input, Select, Spinner, mergeClasses } from "@fluentui/react-components";
import { Add20Regular, ArrowSync20Regular, Dismiss20Regular } from "@fluentui/react-icons";
import { useState } from "react";
import type { Treatment } from "../../core/types";
import { COPY, TREATMENT_HINT, TREATMENT_LABEL, TYPE_LABEL } from "../copy";
import { treatmentsFor, type ActionResult, type ColumnView, type Flow, type FlowState, type RangeView } from "../flow";
import { usePattern, useUi } from "../styles";
import { StatusLine, type Intent } from "./StatusLine";

export interface Notice {
  intent: Intent;
  text: string;
}

interface Props {
  flow: Flow;
  state: FlowState;
  sampleRows: boolean;
  busy: boolean;
  notice: Notice | null;
  onRefresh: () => void;
  run: (action: () => Promise<ActionResult>, success?: string) => Promise<ActionResult>;
}

export function ColumnsScreen({ flow, state, sampleRows, busy, notice, onRefresh, run }: Props) {
  const ui = useUi();
  const p = usePattern();
  const main = state.ranges[0];
  return (
    <section className={ui.stack} aria-labelledby="columns-heading">
      <div className={ui.spread}>
        <h2 id="columns-heading" className={ui.heading}>
          Columns
        </h2>
        <Button icon={<ArrowSync20Regular />} onClick={onRefresh} disabled={busy}>
          Refresh from selection
        </Button>
      </div>
      {busy && <Spinner size="tiny" label="Reading from Excel…" labelPosition="after" />}
      {notice && <StatusLine intent={notice.intent} title={notice.text} live />}
      {state.stale && <StatusLine intent="warning" title={COPY.dataChanged} live />}
      {!main ? (
        !notice && <p className={ui.text}>{COPY.noSelection}</p>
      ) : (
        <>
          <div className={ui.tight}>
            <p className={ui.text}>
              <strong>{COPY.using(main.summary)}</strong>
              {/* The line below already names a table the range grew to. */}
              {main.table && main.resolution?.kind !== "table" ? ` · table ${main.table}` : ""}
            </p>
            {main.resolution && <p className={ui.muted}>{main.resolution.text}</p>}
          </div>
          {main.resolution?.exact && (
            <div className={ui.row}>
              <Button size="small" onClick={() => void run(() => flow.useExactSelection())} disabled={busy}>
                {COPY.useExact}
              </Button>
            </div>
          )}
          <p className={ui.muted}>{COPY.keepsRange}</p>
          <Checkbox
            label="This range has a header row"
            checked={main.hasHeaders}
            onChange={(_, d) => flow.setHasHeaders(0, d.checked === true)}
          />
          {main.notice && <StatusLine intent="info" title={main.notice} />}
          <p className={ui.muted}>
            <span className={p.swatch} aria-hidden="true" />
            {COPY.suggestionsNote}
          </p>
          {!sampleRows && <p className={ui.muted}>{COPY.treatmentsOff}</p>}
          <ColumnList flow={flow} rangeIndex={0} range={main} sampleRows={sampleRows} />

          <h3 className={ui.subheading}>Lookup ranges</h3>
          <p className={ui.muted}>
            Add a range the formula may read, such as a lookup table on another sheet. It is described like the
            selection, with its own private columns.
          </p>
          {state.ranges.slice(1).map((range, i) => (
            <LookupRange key={range.label} flow={flow} range={range} rangeIndex={i + 1} sampleRows={sampleRows} />
          ))}
          <AddLookup busy={busy} run={run} flow={flow} />
        </>
      )}
    </section>
  );
}

function ColumnList({ flow, rangeIndex, range, sampleRows }: { flow: Flow; rangeIndex: number; range: RangeView; sampleRows: boolean }) {
  const ui = useUi();
  return (
    <ul className={mergeClasses(ui.stack, ui.plainList)} aria-label={`Columns in ${range.label}`}>
      {range.columns.map((col) => (
        <ColumnCard key={col.letter} flow={flow} rangeIndex={rangeIndex} col={col} sampleRows={sampleRows} />
      ))}
    </ul>
  );
}

function ColumnCard({ flow, rangeIndex, col, sampleRows }: { flow: Flow; rangeIndex: number; col: ColumnView; sampleRows: boolean }) {
  const ui = useUi();
  const p = usePattern();
  const policy = col.policy;
  const set = (patch: Parameters<Flow["setPolicy"]>[2]) => flow.setPolicy(rangeIndex, col.letter, patch);
  return (
    <li className={mergeClasses(p.card, policy.private && p.cardPrivate)} aria-label={`Column ${col.letter}`}>
      {policy.private && <span className={p.strip} aria-hidden="true" />}
      <div className={ui.row}>
        <span className={p.letter} aria-hidden="true">
          {col.letter}
        </span>
        <span className={ui.columnName}>
          {col.header}
        </span>
        <span className={ui.muted}>{TYPE_LABEL[col.type]}</span>
      </div>
      <Checkbox
        label="Keep private"
        checked={policy.private}
        onChange={(_, d) => set({ private: d.checked === true })}
      />
      {col.hint && <span className={p.hint}>{col.hint}</span>}
      <div className={ui.grid}>
        <Field label="Header sent" size="small" className={ui.field}>
          <Input
            size="small"
            className={ui.fill}
            value={policy.alias ?? ""}
            placeholder={col.header}
            onChange={(_, d) => set({ alias: d.value })}
            spellCheck={false}
          />
        </Field>
        {sampleRows && (
          <Field label="In sample rows" size="small" className={ui.field} hint={TREATMENT_HINT[policy.treatment]}>
            <Select
              size="small"
              className={ui.fill}
              value={policy.treatment}
              onChange={(_, d) => set({ treatment: d.value as Treatment })}
            >
              {treatmentsFor(policy.private).map((t) => (
                <option key={t} value={t}>
                  {TREATMENT_LABEL[t]}
                </option>
              ))}
            </Select>
          </Field>
        )}
        <Field label="Note (optional)" size="small" className={ui.field}>
          <Input
            size="small"
            className={ui.fill}
            value={policy.note ?? ""}
            placeholder="For example: USD"
            onChange={(_, d) => set({ note: d.value })}
            spellCheck={false}
            autoCorrect="off"
            autoCapitalize="off"
            data-gramm="false"
            // Input drops attributes it doesn't know; its input slot passes them on.
            input={{ writingsuggestions: "false" }}
          />
        </Field>
      </div>
    </li>
  );
}

function LookupRange({ flow, range, rangeIndex, sampleRows }: { flow: Flow; range: RangeView; rangeIndex: number; sampleRows: boolean }) {
  const ui = useUi();
  return (
    <div className={ui.stack}>
      <div className={ui.spread}>
        <h4 className={ui.subheading}>{range.summary}</h4>
        <Button
          appearance="subtle"
          size="small"
          icon={<Dismiss20Regular />}
          onClick={() => flow.removeContextRange(rangeIndex)}
          aria-label={`Remove lookup range ${range.label}`}
        >
          Remove
        </Button>
      </div>
      {range.resolution && <p className={ui.muted}>{range.resolution.text}</p>}
      <Checkbox
        label="This range has a header row"
        checked={range.hasHeaders}
        onChange={(_, d) => flow.setHasHeaders(rangeIndex, d.checked === true)}
      />
      {range.notice && <StatusLine intent="info" title={range.notice} />}
      <ColumnList flow={flow} rangeIndex={rangeIndex} range={range} sampleRows={sampleRows} />
    </div>
  );
}

function AddLookup({ flow, busy, run }: { flow: Flow; busy: boolean; run: Props["run"] }) {
  const ui = useUi();
  const [open, setOpen] = useState(false);
  const [address, setAddress] = useState("");
  // Undefined until the user sets the box: the range's row 1 then decides (a header unless it looks like data).
  const [headers, setHeaders] = useState<boolean | undefined>(undefined);

  if (!open) {
    return (
      <div className={ui.row}>
        <Button icon={<Add20Regular />} onClick={() => setOpen(true)} disabled={busy}>
          Add a lookup range
        </Button>
      </div>
    );
  }
  const done = (r: ActionResult) => {
    if (r.ok) {
      setOpen(false);
      setAddress("");
      setHeaders(undefined);
    }
  };
  return (
    <div className={ui.stack}>
      <Field label="Lookup range" hint="For example Regions!A1:B5, or select the cells in Excel and choose Add selected cells.">
        <Input
          value={address}
          onChange={(_, d) => setAddress(d.value)}
          placeholder="Regions!A1:B5"
          spellCheck={false}
          onKeyDown={(e) => {
            if (e.key === "Enter" && address.trim() !== "") void run(() => flow.addContextRange(address, headers), "Added the lookup range.").then(done);
          }}
        />
      </Field>
      <Checkbox label="This range has a header row" checked={headers ?? true} onChange={(_, d) => setHeaders(d.checked === true)} />
      <div className={ui.row}>
        <Button
          appearance="primary"
          disabled={busy || address.trim() === ""}
          onClick={() => void run(() => flow.addContextRange(address, headers), "Added the lookup range.").then(done)}
        >
          Add
        </Button>
        <Button disabled={busy} onClick={() => void run(() => flow.addSelectionAsContext(headers), "Added the lookup range.").then(done)}>
          Add selected cells
        </Button>
        <Button appearance="subtle" onClick={() => setOpen(false)}>
          Cancel
        </Button>
      </div>
    </div>
  );
}
