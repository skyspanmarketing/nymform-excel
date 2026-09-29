// Setup (spec §7.13, screen 1): API key (memory only), endpoint (read-only), model ID.
import {
  Button,
  Field,
  Input,
  renderInput_unstable,
  useInputStyles_unstable,
  useInput_unstable,
  type InputProps,
} from "@fluentui/react-components";
import { useEffect, useRef, type Ref } from "react";
import { MODEL_ID_RE } from "../../core/metrics";
import { COPY } from "../copy";
import { useUi } from "../styles";

interface Props {
  apiKey: string;
  onApiKey: (key: string) => void;
  model: string;
  onModel: (model: string) => void;
  endpoint: string;
  host: string;
  onContinue: () => void;
}

/**
 * Fluent's Input without a value prop, so the input is uncontrolled: React never writes the key
 * into the element's value attribute, where selectors, extensions and page copies can read it.
 */
function KeyInput({ inputRef, ...props }: InputProps & { inputRef: Ref<HTMLInputElement> }) {
  const state = useInput_unstable(props, inputRef);
  delete state.input.value;
  useInputStyles_unstable(state);
  return renderInput_unstable(state);
}

export function SetupScreen({ apiKey, onApiKey, model, onModel, endpoint, host, onContinue }: Props) {
  const ui = useUi();
  const modelOk = MODEL_ID_RE.test(model.trim());
  const keyRef = useRef<HTMLInputElement>(null);
  const initialKey = useRef(apiKey);
  // Show the key kept in memory when the screen opens again: the DOM property only, never the attribute.
  useEffect(() => {
    if (keyRef.current) keyRef.current.value = initialKey.current;
  }, []);
  return (
    <section className={ui.stack} aria-labelledby="setup-heading">
      <h2 id="setup-heading" className={ui.heading}>
        Setup
      </h2>
      <Field label="API key" hint={COPY.keyLine(host)}>
        {/* Kept in memory only (invariant 11): no form, no saving, no autofill. */}
        <KeyInput
          inputRef={keyRef}
          type="password"
          onChange={(_, d) => onApiKey(d.value)}
          autoComplete="off"
          spellCheck={false}
          placeholder="Paste your key"
        />
      </Field>
      <Field label="Endpoint" hint="Set when Nymform was built. It can't be changed here.">
        <Input value={endpoint} readOnly />
      </Field>
      <Field
        label="Model"
        validationState={modelOk ? "none" : "error"}
        validationMessage={modelOk ? undefined : COPY.modelInvalid}
      >
        <Input value={model} onChange={(_, d) => onModel(d.value)} spellCheck={false} autoComplete="off" />
      </Field>
      <div className={ui.row}>
        <Button appearance="primary" onClick={onContinue}>
          Continue
        </Button>
      </div>
    </section>
  );
}
