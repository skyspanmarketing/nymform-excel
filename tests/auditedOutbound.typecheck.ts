// T-A10 (spec §10): code outside src/core/auditor.ts can't construct an AuditedOutbound.
// `tsc --noEmit` checks this file: each line after an @ts-expect-error must fail to type-check,
// or the directive itself is an error. tests/auditor.test.ts runs the same check through the
// TypeScript compiler API, and also confirms each marked line fails without its directive.
import type { AuditedOutbound } from "../src/core/auditor";
import type { Outbound } from "../src/core/types";

const plain: Outbound = { body: "{}", bytes: 2, mode: "structure_only", createdAt: "2026-09-24T12:00:00.000Z" };

// An object literal with every public field is still missing the auditor's brand.
// @ts-expect-error the brand symbol is private to auditor.ts
export const forged: AuditedOutbound = { ...plain, audit: { ok: true, hits: [] } };

// A plain Outbound is not an AuditedOutbound.
// @ts-expect-error an unaudited Outbound is not an AuditedOutbound
export const unaudited: AuditedOutbound = plain;

// A look-alike symbol is a different key.
const lookalike = Symbol("nymform.audited");
// @ts-expect-error a new Symbol is not the auditor's brand
export const spoofed: AuditedOutbound = { ...plain, [lookalike]: true, audit: { ok: true, hits: [] } };

// A parameter typed like provider.send's accepts nothing else.
declare function sendOnly(outbound: AuditedOutbound): void;
export function trySend(): void {
  // @ts-expect-error provider.send accepts only an AuditedOutbound
  sendOnly({ ...plain, audit: { ok: true, hits: [] } });
}
