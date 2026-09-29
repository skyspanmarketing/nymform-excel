# Notes for coding agents

- `docs/spec.md` is the build contract. Its invariants (§2) are non-negotiable. If a task seems to
  need breaking one, stop and ask Jorge.
- Read `docs/vision.md` once. It is direction, not scope: don't build toward it, and avoid
  one-way doors that would close it off. If a fix needs a new abstraction (a new module boundary,
  a new kind of write, a new thing that reads values), stop and ask before adding it.
- The check is frozen after the 2026-09-25 round: `src/core/auditor.ts` and the matchers it imports
  from `src/core/transform.ts` (private values, numbers and dates, date reading, the fold and
  normalize helpers, `aliasRevealsHeader`) and from `src/core/schema.ts` (the header-cell tests). Change them only for a reproducible failure, pinned
  with a test, and have the change independently audited before it reaches `Live`. See
  `docs/decisions.md` (2026-09-25).
- The README claim must stay literally true. If what the check does changes, the claim, `SECURITY.md`
  and the spec change in the same commit.
- Plain wording in code, UI and docs: never "encrypt", "secure", "compliant" or "anonymous".
- No new dependencies without asking. Commits are authored by Jorge's account only.
- Before a commit: `npm run typecheck`, `npm run lint`, `npm test`. For UI changes: `npm run build`,
  then `npx playwright test` (it serves `dist/`, so build first). For manifest changes:
  `npm run validate`. For release-affecting changes: `npm run build && npm run check:release`.
  When matching or suggestions change: `npm run battery`.
