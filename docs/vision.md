# Nymform: direction

Written 2026-09-25. **This is direction, not v0 scope.** Nothing here is built in v0, and nothing
in v0 should be built toward it. It exists so that choices made now don't close doors we may want
later, and so that anyone working on the repo knows which doors those are. The build contract is
`docs/spec.md`, and its invariants (§2) win over anything written here.

## The thesis

A model can write the recipe without seeing the ingredients. Nymform sends a minimized description
of the data: headers, types and simple counts, and when needed sample rows with private values
swapped for stand-ins. The model returns a formula, and Excel computes it locally on the real
values. The values of private columns aren't sent, the formula runs on them in Excel, and
references survive the round trip: a stand-in in the reply comes back as the value it stands for.
What is sent, and the stated limits of the check, are listed in the README and SECURITY.md.

v0 limits itself to formulas. That limit is the v0 boundary, not the product's: a formula is
something the formula gate can read, check and refuse before anything is written, and the user
inserts it with one click they can see.

## What decides whether the thesis holds

**Utility under minimization** is the first question: how much of a model's usefulness survives
when it sees less. The benchmark (`bench/`) is the one artifact that can prove this wrong, so it comes
before any feature that assumes the answer.

The representations form a ladder, from most to least revealing:

1. raw values (bench builds only, as the baseline);
2. stand-ins for private values (substituted mode);
3. semantic descriptions of values (what kind of thing a column holds, without the values);
4. structure plus statistics beyond counts (ranges, distributions);
5. structure plus simple counts (the v0 default: headers, types, blanks, distinct values, average
   words and length of text).

Where on the ladder each kind of question stops working is a research question. The benchmark
measures it per task and per model. The "minimum sufficient representation" for a question is the
lowest rung at which the answer stays right. Finding it automatically is future work.

## Surfaces

There are three possible surfaces:

1. **Excel itself**: the formula lands in the workbook.
2. **The sidebar** (the v0 task pane): the conversation, the columns and what gets sent.
3. **A future "shadow sheet"**: a working surface next to the data for results that aren't one
   formula.

Only the first two exist. A shadow sheet, and "quick actions" that read values to do something for
the user, would each need a deliberate redesign of invariant 9 (writes only through Insert). They
would also need new privacy classes for anything that reads values, with the auditor extended to
cover them. Neither is a small feature, and neither is scheduled.

## Principles

- **Invisible friction, not invisible behaviour.** Checks may run without asking, and should cost
  the user as little as possible. What leaves, what was blocked and why must always be visible on
  request (What gets sent, the Log).
- **Selection, not recolouring.** When Nymform points at cells, it selects them. It never colours,
  highlights or overlays them: that would be a write, it would outlive the pane, and Excel can't
  undo an add-in's writes.
- **Configuration is a business goal, not an engineering instruction.** "Most users never touch
  the settings" is a target for the defaults and the suggestions. It isn't a reason to hide
  settings, guess silently or add automation that changes what is sent.

## Type 1 and Type 2 decisions

Type 1 decisions are expensive to reverse, so they get care, tests and independent review. Type 2
decisions are cheap to change, so we try them and learn.

| Type 1: foundations | Type 2: try and learn |
| --- | --- |
| What data Nymform thinks it is looking at (selection and context resolution) | Sidebar layout and wording |
| What leaves: the builder, the auditor, the one sender, the fixed endpoint | Columns screen grouping, search and review order |
| Suggestions that prevent leaks: a column derived from a private one stays private (spec §7.3) | Wording and order of the other suggestions |
| How it writes: the formula gate, Insert as the only write | Which models are offered by default |
| The stand-in map in memory only, no telemetry | Which columns are suggested private "to be safe" |
| The public claim in the README | Onboarding, demo, the playground workbook |

A change to anything in the left column needs a spec change first.

## Not now

- the shadow sheet and quick actions;
- automatic choice of representation per question;
- saved keys;
- other hosts (Sheets, other Office apps);
- a Nymform server.
