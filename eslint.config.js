// Lint rules. Besides the usual TypeScript checks, these enforce spec invariants in code:
// - only src/core/provider.ts (and the bench-only raw sender) may call fetch (invariant 2);
// - nothing outside src/core/auditor.ts may assert a value to AuditedOutbound (invariant 3);
// - no browser storage anywhere in the add-in (invariants 6 and 11).
const js = require("@eslint/js");
const tseslint = require("typescript-eslint");
const globals = require("globals");

const noAuditedCast = [
  {
    selector: "TSAsExpression > TSTypeReference[typeName.name='AuditedOutbound']",
    message: "Only auditor.ts can create an AuditedOutbound.",
  },
  {
    selector: "TSTypeAssertion > TSTypeReference[typeName.name='AuditedOutbound']",
    message: "Only auditor.ts can create an AuditedOutbound.",
  },
];

const noStorage = [
  { name: "localStorage", message: "No storage in v0: keys and stand-ins live in memory only." },
  { name: "sessionStorage", message: "No storage in v0: keys and stand-ins live in memory only." },
  { name: "indexedDB", message: "No storage in v0: keys and stand-ins live in memory only." },
];

module.exports = tseslint.config(
  { ignores: ["dist/**", "node_modules/**", "coverage/**", "playwright-report/**", "test-results/**", ".scratch/**"] },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    languageOptions: { globals: { ...globals.browser, ...globals.node } },
    rules: {
      "@typescript-eslint/no-unused-vars": ["error", { argsIgnorePattern: "^_", varsIgnorePattern: "^_" }],
    },
  },
  {
    files: ["src/**/*.ts", "src/**/*.tsx"],
    rules: {
      "no-restricted-syntax": ["error", ...noAuditedCast],
      "no-restricted-globals": [
        "error",
        { name: "fetch", message: "Only src/core/provider.ts sends requests (invariant 2)." },
        { name: "XMLHttpRequest", message: "Only src/core/provider.ts sends requests (invariant 2)." },
        { name: "WebSocket", message: "Only src/core/provider.ts sends requests (invariant 2)." },
        { name: "EventSource", message: "Only src/core/provider.ts sends requests (invariant 2)." },
        ...noStorage,
      ],
      "no-restricted-properties": [
        "error",
        { object: "window", property: "fetch", message: "Only src/core/provider.ts sends requests." },
        { object: "globalThis", property: "fetch", message: "Only src/core/provider.ts sends requests." },
        { object: "navigator", property: "sendBeacon", message: "No telemetry (invariant 7)." },
        { object: "document", property: "cookie", message: "No storage in v0." },
        { object: "window", property: "localStorage", message: "No storage in v0." },
        { object: "window", property: "sessionStorage", message: "No storage in v0." },
      ],
    },
  },
  {
    files: ["src/core/provider.ts"],
    rules: {
      "no-restricted-globals": ["error", ...noStorage],
      "no-restricted-properties": "off",
    },
  },
  {
    files: ["src/core/auditor.ts"],
    rules: { "no-restricted-syntax": "off" },
  },
  {
    files: ["tests/**/*.ts", "tests/**/*.tsx", "bench/**/*.ts", "scripts/**/*.mjs", "*.js", "*.ts"],
    rules: { "no-restricted-syntax": ["error", ...noAuditedCast] },
  },
  {
    files: ["*.js", "scripts/**/*.mjs", "tests/e2e/**/*.js"],
    rules: { "@typescript-eslint/no-require-imports": "off" },
  },
);
