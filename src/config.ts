// Build-time configuration. Values are fixed when the bundle is built (webpack DefinePlugin,
// or vitest's `define` in tests) from the NYMFORM_* environment variables. Nothing here can be
// changed at runtime: the endpoint is shown read-only, and the Content Security Policy only
// allows connections to its origin (spec §2 invariant 10).

declare const __NYMFORM_BENCH__: boolean;
declare const __NYMFORM_CONFIG__: {
  endpoint: string;
  defaultModel: string;
  version: string;
  commit: string;
};

/** True only in bench builds. Release builds compile bench code out entirely (invariant 8). */
export const BENCH: boolean = __NYMFORM_BENCH__;

/** Model endpoint base URL, e.g. https://openrouter.ai/api/v1 (no trailing slash). */
export const ENDPOINT: string = __NYMFORM_CONFIG__.endpoint.replace(/\/+$/, "");

/** Host of the endpoint, e.g. openrouter.ai. Shown in the UI and used to decide the `provider` field. */
export const ENDPOINT_HOST: string = new URL(ENDPOINT).host;

/** Default model ID. Must be served through a zero-data-retention-eligible route (spec §3). */
export const DEFAULT_MODEL: string = __NYMFORM_CONFIG__.defaultModel;

export const VERSION: string = __NYMFORM_CONFIG__.version;
export const COMMIT: string = __NYMFORM_CONFIG__.commit;

/**
 * "0.1.0-alpha · build abc1234", or "0.1.0-alpha · local build" for a build made without git (from
 * a downloaded ZIP), whose commit is unknown. Shown in the footer and kept in evaluation records.
 */
export function buildLabel(version: string, commit: string): string {
  const c = commit.trim();
  return c === "" || c === "unknown" ? `${version} · local build` : `${version} · build ${c}`;
}

/** True when the endpoint is OpenRouter, which takes `provider: { zdr: true }`. */
export const IS_OPENROUTER: boolean = ENDPOINT_HOST === "openrouter.ai";
