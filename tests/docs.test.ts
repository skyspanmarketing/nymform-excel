// Wording checks for statements in the documents that describe observed Excel behaviour, so a
// document can't drift back to a description that a real-Excel run contradicted.
import { readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const security = readFileSync(join(ROOT, "SECURITY.md"), "utf8").replace(/\s+/g, " ");

describe("SECURITY.md", () => {
  it("describes a text constant over 255 characters as written and calculated to #VALUE!, not as a failed Insert", () => {
    // Real Excel for Mac 16.113 (QA pass, 2026-09-28): the formula is written, the cell shows #VALUE!,
    // and the Result screen reports the error outcome. Insert does not fail.
    expect(security).not.toMatch(/255 characters; Insert then fails/);
    expect(security).toMatch(/longer than 255 characters[^.]*written/);
    expect(security).toMatch(/#VALUE!/);
  });
});
