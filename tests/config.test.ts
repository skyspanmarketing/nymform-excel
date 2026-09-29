import { describe, expect, it } from "vitest";
import { buildLabel } from "../src/config";
import { COPY } from "../src/taskpane/copy";

describe("build label", () => {
  it("names the commit, or says local build when the bundle was built without git", () => {
    expect(buildLabel("0.1.0-alpha", "abc1234")).toBe("0.1.0-alpha · build abc1234");
    expect(buildLabel("0.1.0-alpha", "unknown")).toBe("0.1.0-alpha · local build");
    expect(buildLabel("0.1.0-alpha", " ")).toBe("0.1.0-alpha · local build");
  });

  it("the pane footer carries it and the publisher", () => {
    expect(COPY.footer("0.1.0-alpha", "unknown")).toBe("Nymform for Excel 0.1.0-alpha · local build · SkySpan");
    expect(COPY.footer("0.1.0-alpha", "abc1234")).toBe("Nymform for Excel 0.1.0-alpha · build abc1234 · SkySpan");
  });
});
