import { describe, expect, it } from "vitest";
import { check, tokenize, type GateContext } from "../src/core/formulaGate";
import { restoreFormula, restoreText, type TokenLookup } from "../src/core/restore";

const ctx: GateContext = {
  sheet: "Orders",
  sheets: ["Orders", "Regions"],
  tableNames: [],
  tables: [],
  allowedRanges: ["Orders!A1:F501"],
};

function lookup(entries: Record<string, string>, prefix?: "" | "NYMFORM_", reserved: readonly string[] = []): TokenLookup {
  const m = new Map(Object.entries(entries));
  return prefix === undefined ? { valueOf: (t) => m.get(t) } : { prefix, valueOf: (t) => m.get(t), isReserved: (t) => reserved.includes(t) };
}

const map = lookup({
  PERSON_001: 'Ana "AJ" Ruiz',
  PERSON_002: "Maria Lopez",
  EMAIL_001: "maria@example.test",
  ID_1000: "S-99812",
  TEXT_004: "a$&b$1",
});

/** Non-string tokens, in order, as [type, text]. */
function skeleton(formula: string): [string, string][] {
  return tokenize(formula)
    .filter((t) => t.type !== "string")
    .map((t) => [t.type, t.text]);
}

describe("restoreFormula", () => {
  it("T-T4: changes string literals only", () => {
    const before = '=IF(A2="PERSON_002",SUMIFS(E2:E501,B2:B501,"EMAIL_001")&PERSON_002,{"ID_1000",1})';
    const { formula, unknownTokens } = restoreFormula(before, map);
    expect(formula).toBe('=IF(A2="Maria Lopez",SUMIFS(E2:E501,B2:B501,"maria@example.test")&PERSON_002,{"S-99812",1})');
    expect(skeleton(formula)).toEqual(skeleton(before));
    expect(unknownTokens).toEqual([]);
  });

  it("T-T4: operators, references and parentheses stay byte-identical", () => {
    const before = "= SUMIFS( 'Orders'!$E$2:$E$501 , A2:A501 , \"PERSON_001\" ) * -1% & \"x\"\"PERSON_002\"\"y\"";
    const { formula } = restoreFormula(before, map);
    const a = tokenize(before);
    const b = tokenize(formula);
    expect(b.map((t) => t.type)).toEqual(a.map((t) => t.type));
    a.forEach((t, i) => {
      if (t.type !== "string") expect(b[i]!.text).toBe(t.text);
    });
    expect(b.filter((t) => t.type === "string").map((t) => t.value)).toEqual(['Ana "AJ" Ruiz', 'x"Maria Lopez"y']);
  });

  it("lists the stand-ins it put back, and only those inside string literals", () => {
    const r = restoreFormula('=IF(A2="PERSON_002",COUNTIFS(B2:B9,"EMAIL_001","PERSON_404"),PERSON_001)', map);
    expect(r.restoredTokens).toEqual(["PERSON_002", "EMAIL_001"]);
    expect(r.unknownTokens).toEqual(["PERSON_404"]);
  });

  it("T-T4: leaves tokens outside strings alone", () => {
    const before = "=PERSON_001&PERSON_002";
    expect(restoreFormula(before, map)).toEqual({ formula: before, unknownTokens: [], restoredTokens: [] });
  });

  it("T-T4: leaves unknown tokens in place and reports them once", () => {
    const { formula, unknownTokens } = restoreFormula('=IF(A2="PERSON_999","PERSON_001 and PERSON_999",EMAIL_404&"EMAIL_404")', map);
    expect(formula).toBe('=IF(A2="PERSON_999","Ana ""AJ"" Ruiz and PERSON_999",EMAIL_404&"EMAIL_404")');
    expect(unknownTokens).toEqual(["PERSON_999", "EMAIL_404"]);
  });

  it('T-T6: PERSON_001 -> Ana "AJ" Ruiz gives "Ana ""AJ"" Ruiz" and still passes the gate', () => {
    const before = '=COUNTIFS(A2:A501,"PERSON_001")';
    expect(check(before, ctx)).toEqual({ ok: true, reasons: [] });
    const { formula } = restoreFormula(before, map);
    expect(formula).toBe('=COUNTIFS(A2:A501,"Ana ""AJ"" Ruiz")');
    expect(formula).toContain('"Ana ""AJ"" Ruiz"');
    expect(check(formula, ctx)).toEqual({ ok: true, reasons: [] });
    expect(tokenize(formula).find((t) => t.type === "string")?.value).toBe('Ana "AJ" Ruiz');
  });

  it("restores tokens inside wildcards and longer text", () => {
    const { formula } = restoreFormula('=COUNTIF(A2:A501,"*PERSON_002*")+COUNTIF(B2:B501,"<>PERSON_002")', map);
    expect(formula).toBe('=COUNTIF(A2:A501,"*Maria Lopez*")+COUNTIF(B2:B501,"<>Maria Lopez")');
  });

  it("inserts values literally, even with $ patterns", () => {
    expect(restoreFormula('="TEXT_004"', map).formula).toBe('="a$&b$1"');
  });

  it("matches whole tokens only", () => {
    const { formula, unknownTokens } = restoreFormula('="XPERSON_001 PERSON_0010 PERSON_01 PERSON_001X person_001"', map);
    expect(formula).toBe('="XPERSON_001 PERSON_0010 PERSON_01 PERSON_001X person_001"');
    expect(unknownTokens).toEqual(["PERSON_0010"]);
  });

  it("handles NYMFORM_-prefixed tokens", () => {
    const prefixed = lookup({ NYMFORM_PERSON_001: "Ana Ruiz", NYMFORM_ID_002: "S-1" }, "NYMFORM_", ["PERSON_050"]);
    const { formula, unknownTokens } = restoreFormula('=IF(A2="NYMFORM_PERSON_001","NYMFORM_ID_002","PERSON_001","PERSON_050")', prefixed);
    // With the prefix in use, a bare PERSON_050 is text from the sheet, not a stand-in. A bare
    // PERSON_001 is left as written too, but flagged: it may be NYMFORM_PERSON_001 with the prefix dropped.
    expect(formula).toBe('=IF(A2="Ana Ruiz","S-1","PERSON_001","PERSON_050")');
    expect(unknownTokens).toEqual(["PERSON_001"]);
    expect(restoreFormula('="NYMFORM_PERSON_404"', prefixed).unknownTokens).toEqual(["NYMFORM_PERSON_404"]);
  });

  it("flags a stand-in whose NYMFORM_ prefix the model dropped (red-team prefix-drop proofs)", () => {
    const prefixed = lookup({ NYMFORM_PERSON_001: "Ana Ruiz", NYMFORM_PERSON_002: "Ben Okafor" }, "NYMFORM_", ["PERSON_050"]);
    const f = '=SUMIFS(C2:C4,A2:A4,"PERSON_001")';
    expect(restoreFormula(f, prefixed)).toEqual({ formula: f, unknownTokens: ["PERSON_001"], restoredTokens: [] });
    expect(restoreText("Total Amount for PERSON_001.", prefixed)).toEqual({
      text: "Total Amount for PERSON_001.",
      unknownTokens: ["PERSON_001"],
    });
    // Sheet text such as PERSON_050, whose prefixed form this session never made, stays quiet.
    expect(restoreFormula('=SUMIFS(C2:C4,B2:B4,"PERSON_050")', prefixed)).toEqual({
      formula: '=SUMIFS(C2:C4,B2:B4,"PERSON_050")',
      unknownTokens: [],
      restoredTokens: [],
    });
    // Without the prefix, nothing changes: a bare token is looked up as is.
    expect(restoreFormula('="PERSON_001"', lookup({ PERSON_001: "Ana Ruiz" }, "")).formula).toBe('="Ana Ruiz"');
  });

  it("never matches the tail of a prefixed token", () => {
    const { formula, unknownTokens } = restoreFormula('="NYMFORM_PERSON_001"', map);
    expect(formula).toBe('="NYMFORM_PERSON_001"');
    expect(unknownTokens).toEqual(["NYMFORM_PERSON_001"]);
  });

  it("returns a formula it can't read unchanged", () => {
    expect(restoreFormula('=SUM("PERSON_001"', map)).toEqual({ formula: '=SUM("PERSON_001"', unknownTokens: [], restoredTokens: [] });
  });

  it("gives a restored URL to the second gate check to reject", () => {
    const urlMap = lookup({ TEXT_001: "https://example.test/x" });
    const { formula } = restoreFormula('=IF(A2="TEXT_001",1,0)', urlMap);
    const r = check(formula, ctx);
    expect(r.ok).toBe(false);
    expect(r.reasons.map((x) => x.code)).toContain("G-URL");
    for (const reason of r.reasons) expect(reason.detail).not.toContain("example.test");
  });

  it("property: only string literal contents ever change", () => {
    let seed = 99;
    const next = () => {
      seed = (seed * 1103515245 + 12345) >>> 0;
      return seed / 4294967296;
    };
    const parts = ['"PERSON_001"', '"x PERSON_002 y"', '"ID_1000"', "PERSON_001", "A2", "&", "+", "(", ")", "SUM(", ",", '""', '"a""b"', " ", "{", "}"];
    for (let n = 0; n < 2000; n++) {
      let f = "=";
      const len = 1 + Math.floor(next() * 10);
      for (let i = 0; i < len; i++) f += parts[Math.floor(next() * parts.length)]!;
      let toks;
      try {
        toks = tokenize(f);
      } catch {
        expect(restoreFormula(f, map).formula).toBe(f);
        continue;
      }
      const out = restoreFormula(f, map).formula;
      const after = tokenize(out);
      expect(after.map((t) => t.type)).toEqual(toks.map((t) => t.type));
      toks.forEach((t, i) => {
        if (t.type !== "string") expect(after[i]!.text).toBe(t.text);
      });
    }
  });
});

describe("restoreText", () => {
  it("replaces tokens anywhere, without quoting", () => {
    const { text, unknownTokens } = restoreText('Counts rows where column A is PERSON_001 ("PERSON_002").', map);
    expect(text).toBe('Counts rows where column A is Ana "AJ" Ruiz ("Maria Lopez").');
    expect(unknownTokens).toEqual([]);
  });

  it("reports unknown tokens once and leaves them in place", () => {
    const { text, unknownTokens } = restoreText("PERSON_777 and PERSON_777 and EMAIL_001", map);
    expect(text).toBe("PERSON_777 and PERSON_777 and maria@example.test");
    expect(unknownTokens).toEqual(["PERSON_777"]);
  });

  it("handles NYMFORM_-prefixed tokens and values with $ patterns", () => {
    const prefixed = lookup({ NYMFORM_TEXT_001: "a$'b" }, "NYMFORM_");
    expect(restoreText("NYMFORM_TEXT_001 / TEXT_001", prefixed)).toEqual({ text: "a$'b / TEXT_001", unknownTokens: ["TEXT_001"] });
    expect(restoreText("NYMFORM_TEXT_001 / TEXT_002", prefixed)).toEqual({ text: "a$'b / TEXT_002", unknownTokens: ["TEXT_002"] });
    expect(restoreText("TEXT_004", map).text).toBe("a$&b$1");
  });

  it("leaves text without tokens unchanged", () => {
    expect(restoreText("", map)).toEqual({ text: "", unknownTokens: [] });
    expect(restoreText("Sum of column E.", map)).toEqual({ text: "Sum of column E.", unknownTokens: [] });
  });
});
