// Result (spec §7.13, screen 5): the formula, gate status, explanation and assumptions (restored
// for display), placement, Insert formula and Copy, and the optional rating.
import { Button, Field, Input, Switch, Textarea, ToggleButton } from "@fluentui/react-components";
import { Copy20Regular } from "@fluentui/react-icons";
import { useRef, useState } from "react";
import { copyToClipboard } from "../browser";
import { COPY } from "../copy";
import { fillSpan, type Flow, type FlowResult } from "../flow";
import { useUi } from "../styles";
import { useFocusOnMount } from "./focus";
import { StatusLine, type Intent } from "./StatusLine";

interface Props {
  flow: Flow;
  result: FlowResult;
  onAnswer: (answer: string) => void;
  onAskAnother: () => void;
}

type Rating = "worked" | "partly" | "wrong";

const RATINGS: { value: Rating; label: string }[] = [
  { value: "worked", label: "Worked" },
  { value: "partly", label: "Partly" },
  { value: "wrong", label: "Wrong" },
];

export function ResultScreen({ flow, result, onAnswer, onAskAnother }: Props) {
  const ui = useUi();
  const heading = useFocusOnMount<HTMLHeadingElement>();
  const title = result.kind === "formula" ? "Formula" : result.kind === "clarify" ? "The model asks" : "Answer";
  return (
    <section className={ui.stack} aria-labelledby="result-heading">
      <h2 id="result-heading" ref={heading} tabIndex={-1} className={ui.heading}>
        {title}
      </h2>
      {result.kind === "formula" && <FormulaResult flow={flow} result={result} />}
      {result.kind === "clarify" && <Clarify result={result} onAnswer={onAnswer} />}
      {result.kind === "answer" && <Explanation result={result} />}
      {result.kind !== "formula" && <UnknownTokens result={result} />}
      {result.kind !== "clarify" && <RatingRow flow={flow} resultId={result.id} />}
      <div className={ui.row}>
        <Button onClick={onAskAnother}>Ask another question</Button>
      </div>
    </section>
  );
}

function UnknownTokens({ result }: { result: FlowResult }) {
  // Those inside the formula block it and have their own line (Unresolved).
  const elsewhere = result.unknownTokens.filter((t) => !result.unresolvedTokens.includes(t));
  if (elsewhere.length === 0) return null;
  return <StatusLine intent="warning" title={COPY.unknownTokens(elsewhere)} />;
}

function Unresolved({ result }: { result: FlowResult }) {
  if (result.unresolvedTokens.length === 0) return null;
  return <StatusLine intent="error" title={COPY.unresolvedTokens([...result.unresolvedTokens])} live />;
}

function Cautions({ result }: { result: FlowResult }) {
  if (!result.canInsert) return null;
  return (
    <>
      {result.cautions.map((c, i) => (
        <StatusLine key={i} intent="warning" title={c} />
      ))}
      {result.wholeColumns.length > 0 && <StatusLine intent="info" title={COPY.wholeColumns([...result.wholeColumns])} />}
    </>
  );
}

function Explanation({ result }: { result: FlowResult }) {
  const ui = useUi();
  return (
    <>
      {result.explanation.trim() !== "" && <p className={ui.text}>{result.explanation}</p>}
      {result.assumptions.length > 0 && (
        <div className={ui.tight}>
          <span className={ui.subheading}>Assumptions</span>
          <ul className={ui.list}>
            {result.assumptions.map((a, i) => (
              <li key={i} className={ui.text}>
                {a}
              </li>
            ))}
          </ul>
        </div>
      )}
    </>
  );
}

function FormulaResult({ flow, result }: { flow: Flow; result: FlowResult }) {
  const ui = useUi();
  const [cell, setCell] = useState(result.placement?.cell ?? "");
  const [fillDown, setFillDown] = useState(result.placement?.fillDown ?? false);
  const [busy, setBusy] = useState(false);
  // The second click of a double click can arrive before `busy` has disabled the button; it is
  // dropped here so it never reaches the flow, whose own guard refuses it anyway.
  const inserting = useRef(false);
  // `confirm` is the span the warning asks about; Replace and insert confirms that span only.
  const [insertMsg, setInsertMsg] = useState<{ intent: Intent; text: string; reasons?: string[]; confirm?: string } | null>(null);
  const [copyMsg, setCopyMsg] = useState<{ intent: Intent; text: string } | null>(null);
  const gate = result.gateStatus;
  const span = fillSpan(cell, result.lastRow);

  const insert = async (confirmedSpan?: string) => {
    if (inserting.current) return;
    inserting.current = true;
    setBusy(true);
    setInsertMsg(null);
    try {
      const r = await flow.insert(result.id, cell, fillDown, confirmedSpan !== undefined ? { confirmedSpan } : {});
      // Written. Only a checked result is a success; the rest say what to look at.
      if (r.ok) setInsertMsg({ intent: r.check === "ok" ? "success" : "warning", text: r.message });
      else if (r.needsConfirm && r.span) setInsertMsg({ intent: "warning", text: r.message, confirm: r.span });
      else setInsertMsg({ intent: "error", text: r.message, ...(r.reasons ? { reasons: r.reasons } : {}) });
    } catch {
      setInsertMsg({ intent: "error", text: COPY.insertFailed });
    } finally {
      inserting.current = false;
      setBusy(false);
    }
  };

  // A confirmation is for the cells it named; changing the cell or Fill down asks again.
  const dropConfirm = () => setInsertMsg((m) => (m?.confirm ? null : m));

  const copy = async () => {
    const text = flow.copyText(result.id);
    if (text === null) return;
    const ok = await copyToClipboard(text);
    if (ok) flow.markCopied(result.id);
    setCopyMsg(ok ? { intent: "success", text: COPY.copied } : { intent: "error", text: COPY.copyFailed });
  };

  return (
    <>
      <pre className={`${ui.code} ${ui.formula}`} tabIndex={0} aria-label="Formula">
        {result.formula ?? ""}
      </pre>
      {gate && (
        <StatusLine intent={gate.ok ? "success" : "error"} title={gate.headline} live>
          {gate.reasons.length > 0 && (
            <ul className={ui.list}>
              {gate.reasons.map((r, i) => (
                <li key={i}>{r}</li>
              ))}
            </ul>
          )}
        </StatusLine>
      )}
      <Unresolved result={result} />
      <UnknownTokens result={result} />
      <Cautions result={result} />
      <Explanation result={result} />
      {result.canInsert && (
        <div className={ui.tight}>
          <Field label="Insert in cell" hint={`On sheet ${result.sheet}. Nothing is written until you choose Insert formula.`}>
            <Input
              value={cell}
              onChange={(_, d) => {
                setCell(d.value);
                dropConfirm();
              }}
              spellCheck={false}
            />
          </Field>
          <Switch
            label={COPY.fillLabel(result.lastRow, span)}
            checked={fillDown}
            onChange={(_, d) => {
              setFillDown(d.checked);
              dropConfirm();
            }}
          />
        </div>
      )}
      <div className={ui.row}>
        <Button appearance="primary" disabled={!result.canInsert || busy} onClick={() => void insert()}>
          Insert formula
        </Button>
        <Button icon={<Copy20Regular />} disabled={!result.canInsert} onClick={() => void copy()}>
          Copy
        </Button>
      </div>
      <div role="status" aria-live="polite" aria-atomic="true">
        {insertMsg && (
          <StatusLine intent={insertMsg.intent} title={insertMsg.text}>
            {insertMsg.confirm && (
              <Button size="small" disabled={busy} onClick={() => void insert(insertMsg.confirm)}>
                {COPY.overwriteButton}
              </Button>
            )}
            {insertMsg.reasons && insertMsg.reasons.length > 0 && (
              <ul className={ui.list}>
                {insertMsg.reasons.map((r, i) => (
                  <li key={i}>{r}</li>
                ))}
              </ul>
            )}
          </StatusLine>
        )}
        {copyMsg && <StatusLine intent={copyMsg.intent} title={copyMsg.text} />}
      </div>
    </>
  );
}

function Clarify({ result, onAnswer }: { result: FlowResult; onAnswer: (answer: string) => void }) {
  const ui = useUi();
  const [answer, setAnswer] = useState("");
  return (
    <>
      <p className={ui.text}>{result.explanation}</p>
      <Field label="Your answer" hint="It goes out with the earlier question, in stand-in form.">
        {/* No browser spell check or writing aids: some send the text to a cloud service. */}
        <Textarea
          value={answer}
          onChange={(_, d) => setAnswer(d.value)}
          resize="vertical"
          rows={2}
          spellCheck={false}
          autoCapitalize="off"
          data-gramm="false"
          // Textarea doesn't pass these through from its own props; its textarea slot does.
          textarea={{ autoCorrect: "off", writingsuggestions: "false" }}
          onKeyDown={(e) => {
            if (e.key === "Enter" && (e.ctrlKey || e.metaKey) && answer.trim() !== "") {
              e.preventDefault();
              onAnswer(answer);
            }
          }}
        />
      </Field>
      <div className={ui.row}>
        <Button appearance="primary" disabled={answer.trim() === ""} onClick={() => onAnswer(answer)}>
          Preview
        </Button>
      </div>
    </>
  );
}

function RatingRow({ flow, resultId }: { flow: Flow; resultId: number }) {
  const ui = useUi();
  const [rating, setRating] = useState<Rating | null>(null);
  return (
    <div className={ui.tight}>
      <span className={ui.subheading} id={`rating-${resultId}`}>
        Did this work?
      </span>
      <div className={ui.row} role="group" aria-labelledby={`rating-${resultId}`}>
        {RATINGS.map((r) => (
          <ToggleButton
            key={r.value}
            size="small"
            checked={rating === r.value}
            onClick={() => {
              setRating(r.value);
              flow.rate(resultId, r.value);
            }}
          >
            {r.label}
          </ToggleButton>
        ))}
      </div>
      {rating && <span className={ui.muted}>Rated. It goes into the evaluation report, which stays here unless you export it.</span>}
    </div>
  );
}
