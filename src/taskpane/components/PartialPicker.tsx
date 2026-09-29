// A part of a private value typed in the question ("Felix" of "Felix Bianchi"): the pane lists the
// values it could be, each with the stand-in it goes out as and where it is, so the user can pick.
// Show in sheet selects the value's cells in Excel (selection only; nothing is written). The real
// values shown here stay in the pane: only the chosen stand-ins go into the request.
import { Button, Checkbox } from "@fluentui/react-components";
import { useState } from "react";
import type { PartialCandidate, PartialMatch } from "../../core/partial";
import type { PartialPick, PartialView } from "../flow";
import { usePattern, useUi } from "../styles";
import { StatusLine } from "./StatusLine";

interface Props {
  partial: PartialView;
  canShow: boolean;
  busy: boolean;
  /** Use these stand-ins for the part (by key); an empty list keeps the part as typed. */
  onChoose: (key: string, tokens: readonly string[]) => void;
  /** Forget the choice for the part, so the pane asks again. */
  onReset: (key: string) => void;
  onShow: (sheet: string, cells: readonly string[]) => void;
}

export function PartialPicker({ partial, canShow, busy, onChoose, onReset, onShow }: Props) {
  const ui = useUi();
  if (partial.pending.length === 0 && partial.auto.length === 0 && partial.chosen.length === 0 && partial.kept.length === 0) return null;
  return (
    <div className={ui.stack}>
      {partial.pending.map((m) => (
        <PendingPart key={m.key} match={m} canShow={canShow} busy={busy} onChoose={onChoose} onShow={onShow} />
      ))}
      {partial.auto.map((p) => (
        <Picked key={p.key} pick={p} automatic busy={busy} action="Keep as typed" onAction={() => onChoose(p.key, [])} />
      ))}
      {partial.chosen.map((p) => (
        <Picked key={p.key} pick={p} busy={busy} action="Change" onAction={() => onReset(p.key)} />
      ))}
      {partial.kept.map((k) => (
        <StatusLine key={k.key} intent="info" title={`“${k.fragment}” is kept as typed, so the check blocks it.`}>
          <span>
            <Button size="small" appearance="transparent" disabled={busy} onClick={() => onReset(k.key)}>
              Choose a value instead
            </Button>
          </span>
        </StatusLine>
      ))}
    </div>
  );
}

function Picked({ pick, automatic = false, busy, action, onAction }: { pick: PartialPick; automatic?: boolean; busy: boolean; action: string; onAction: () => void }) {
  const p = usePattern();
  const title = automatic ? `Read “${pick.fragment}” as ${pick.values.join(" or ")}.` : `“${pick.fragment}” is ${pick.values.join(" or ")}.`;
  return (
    <StatusLine intent="info" title={title}>
      <span>
        Sent as{" "}
        {pick.tokens.map((t, i) => (
          <span key={t}>
            {i > 0 && " or "}
            <span className={p.chip}>{t}</span>
          </span>
        ))}
        .{" "}
        <Button size="small" appearance="transparent" disabled={busy} onClick={onAction}>
          {action}
        </Button>
      </span>
    </StatusLine>
  );
}

function PendingPart({
  match,
  canShow,
  busy,
  onChoose,
  onShow,
}: {
  match: PartialMatch;
  canShow: boolean;
  busy: boolean;
  onChoose: (key: string, tokens: readonly string[]) => void;
  onShow: (sheet: string, cells: readonly string[]) => void;
}) {
  const ui = useUi();
  const p = usePattern();
  const [picked, setPicked] = useState<readonly string[]>([]);
  const n = match.candidates.length + match.more;
  const toggle = (token: string, on: boolean) => setPicked((cur) => (on ? [...cur, token] : cur.filter((t) => t !== token)));
  return (
    <section className={`${p.card} ${ui.tight}`} aria-label={`Which value is ${match.fragment}?`}>
      <p className={ui.text}>
        <strong>“{match.fragment}”</strong> is part of {n === 1 ? "a private value" : `${n} private values`}.{" "}
        {n === 1 ? "Is this the one you meant?" : "Which did you mean?"}
      </p>
      <ul className={`${ui.plainList} ${ui.stack}`} aria-label={`Values for ${match.fragment}`}>
        {match.candidates.map((c) => (
          <Candidate key={c.token} candidate={c} canShow={canShow} busy={busy} checked={picked.includes(c.token)} onToggle={toggle} onShow={onShow} />
        ))}
      </ul>
      {match.more > 0 && (
        <p className={ui.muted}>
          {match.more} more not listed. Type more of the value to narrow it down.
        </p>
      )}
      <div className={ui.row}>
        <Button appearance="primary" disabled={busy || picked.length === 0} onClick={() => onChoose(match.key, picked)}>
          {picked.length > 1 ? `Use these ${picked.length}` : "Use this value"}
        </Button>
        <Button disabled={busy} onClick={() => onChoose(match.key, [])}>
          Keep as typed
        </Button>
      </div>
    </section>
  );
}

function Candidate({
  candidate: c,
  canShow,
  busy,
  checked,
  onToggle,
  onShow,
}: {
  candidate: PartialCandidate;
  canShow: boolean;
  busy: boolean;
  checked: boolean;
  onToggle: (token: string, on: boolean) => void;
  onShow: (sheet: string, cells: readonly string[]) => void;
}) {
  const ui = useUi();
  const p = usePattern();
  const place = c.places[0];
  const where = c.places.map((pl) => `${pl.header} (${pl.sheet}!${pl.column})`).join(", ");
  return (
    <li className={ui.tight}>
      <Checkbox checked={checked} disabled={busy} onChange={(_, d) => onToggle(c.token, d.checked === true)} label={c.value} />
      <div className={ui.row}>
        <span className={ui.muted}>
          <span className={p.chip}>{c.token}</span> · {where} · {c.count === 1 ? "1 cell" : `${c.count} cells`}
        </span>
        {canShow && place && (
          <Button
            size="small"
            appearance="secondary"
            disabled={busy}
            onClick={() => onShow(place.sheet, place.cells)}
            aria-label={`Show ${c.value} in the sheet`}
          >
            Show in sheet
          </Button>
        )}
      </div>
    </li>
  );
}
