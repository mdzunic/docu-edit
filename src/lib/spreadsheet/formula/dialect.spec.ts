// SPEC-119 "xlsx dialect", "ODS dialect", "Crossing formats": both
// translators in both directions — every row of both tables, the `_xlpm.`
// lexical scope rule (nested LETs, a name shadowing a parameter), the three
// ODS dialects recognised by namespace URI, whitespace carried across, the
// corpus round trips (Excel-shaped, canonical and LibreOffice-shaped) and the
// listed non-identities asserted as such — and formulaForTarget, the one
// rule for a cell crossing formats.
import { describe, expect, it } from "vitest";
import { CANONICAL_CORPUS, EXCEL_CORPUS, LIBREOFFICE_CORPUS, ODS_NON_IDENTITIES } from "../fixtures/formulaCorpus";
import type { Cell } from "../model";
import {
  formulaForTarget,
  fromOdsFormula,
  fromXlsxFormula,
  MSOXL_NAMESPACE,
  OOOC_NAMESPACE,
  OPENFORMULA_NAMESPACE,
  readOdsFormula,
  readXlsxFormula,
  toOdsFormula,
  toXlsxFormula,
} from "./dialect";
import { FormulaSyntaxError } from "./parse";

/** The namespace declarations LibreOffice puts on content.xml's root. */
const LO_NAMESPACES = new Map([
  ["of", OPENFORMULA_NAMESPACE],
  ["oooc", OOOC_NAMESPACE],
  ["msoxl", MSOXL_NAMESPACE],
  ["table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"],
]);

describe("xlsx", () => {
  it.each([
    ["_xlfn.CONCAT(A1,B1)", "CONCAT(A1,B1)"],
    ["_xlfn.TEXTJOIN(\",\",TRUE,A1:A3)", "TEXTJOIN(\",\",TRUE,A1:A3)"],
    ["_xlfn._xlws.SORT(A1:A9)", "SORT(A1:A9)"],
    ["_xlfn._xlws.FILTER(A1:A9,B1:B9>0)", "FILTER(A1:A9,B1:B9>0)"],
    ["_xlfn.SINGLE(A1:A9)", "SINGLE(A1:A9)"],
    ["_xlfn.XOR(A1,B1)+_xlfn.IFNA(C1,0)+_xlfn.DAYS(D1,E1)", "XOR(A1,B1)+IFNA(C1,0)+DAYS(D1,E1)"],
    ["SUM(A1:A3)+VLOOKUP(A1,B:C,2,FALSE)", "SUM(A1:A3)+VLOOKUP(A1,B:C,2,FALSE)"],
    ["_xlfn.LET(_xlpm.x,5,_xlpm.x*2)", "LET(x,5,x*2)"],
  ])("%s ↔ %s", (wire, canonical) => {
    expect(fromXlsxFormula(wire)).toBe(canonical);
    expect(toXlsxFormula(canonical)).toBe(wire);
  });

  it("from: strips both prefixes (any case) and upper-cases function names, strips _xlpm., drops one stray leading =", () => {
    expect(fromXlsxFormula("_xlfn.concat(a1,_xlfn._xlws.sort(b1:b3))")).toBe("CONCAT(a1,SORT(b1:b3))");
    expect(fromXlsxFormula("myUdf(1)+sum(2)")).toBe("MYUDF(1)+SUM(2)");
    expect(fromXlsxFormula("=SUM(A1)")).toBe("SUM(A1)");
    expect(fromXlsxFormula("==A1")).toBe("=A1");
    expect(fromXlsxFormula("_xlpm.Rate*2")).toBe("Rate*2");
    expect(fromXlsxFormula("_xlfn.VALUETOTEXT(A1)")).toBe("VALUETOTEXT(A1)");
  });

  it("to: _xlpm. on LET parameters (arguments 1, 3, 5 … but the last) and LAMBDA's (all but the last), and on their uses inside the call", () => {
    expect(toXlsxFormula("LET(rate,0.1,n,12,PMT(rate/n,n,1000))")).toBe("_xlfn.LET(_xlpm.rate,0.1,_xlpm.n,12,PMT(_xlpm.rate/_xlpm.n,_xlpm.n,1000))");
    expect(toXlsxFormula("LAMBDA(a,b,a+b)")).toBe("_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)");
    expect(toXlsxFormula("LAMBDA(x)")).toBe("_xlfn.LAMBDA(x)");
    expect(toXlsxFormula("LET(x,1)")).toBe("_xlfn.LET(_xlpm.x,1)");
    expect(toXlsxFormula("LET(x)")).toBe("_xlfn.LET(x)");
  });

  it("to: lexical scope — nested LETs, and a name outside the call (or shadowed by a parameter's spelling) left alone", () => {
    expect(toXlsxFormula("LET(x,1,LET(y,x+1,x*y))")).toBe("_xlfn.LET(_xlpm.x,1,_xlfn.LET(_xlpm.y,_xlpm.x+1,_xlpm.x*_xlpm.y))");
    expect(toXlsxFormula("Rate*LET(rate,2,rate)+Rate")).toBe("Rate*_xlfn.LET(_xlpm.rate,2,_xlpm.rate)+Rate");
    expect(toXlsxFormula("LET(x,1,x)+x")).toBe("_xlfn.LET(_xlpm.x,1,_xlpm.x)+x");
    expect(toXlsxFormula("LET(X,1,x+Sheet1!x)")).toBe("_xlfn.LET(_xlpm.X,1,_xlpm.x+Sheet1!x)");
    expect(toXlsxFormula("SUM(x,LAMBDA(x,x*2)(3))")).toBe("SUM(x,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(3))");
  });

  it("to: restores each prefix from the table; names absent from it stay as written", () => {
    expect(toXlsxFormula("STDEV.S(A1:A9)+ISO.CEILING(B1)+F.DIST(1,2,3,TRUE)+FDIST(1,2,3)")).toBe(
      "_xlfn.STDEV.S(A1:A9)+ISO.CEILING(B1)+_xlfn.F.DIST(1,2,3,TRUE)+FDIST(1,2,3)",
    );
  });

  it("round-trips every Excel-shaped corpus formula, and every canonical one the other way", () => {
    for (const wire of EXCEL_CORPUS) {
      expect(toXlsxFormula(fromXlsxFormula(wire)), wire).toBe(wire);
    }
    for (const canonical of CANONICAL_CORPUS) {
      expect(fromXlsxFormula(toXlsxFormula(canonical)), canonical).toBe(canonical);
    }
  });

  it("keeps every token's whitespace", () => {
    expect(fromXlsxFormula(" _xlfn.CONCAT( A1 ,\r\n  B1 ) ")).toBe(" CONCAT( A1 ,\r\n  B1 ) ");
  });
});

describe("ODS", () => {
  it.each([
    ["of:=[.A1]+[.$A$1]", "A1+$A$1"],
    ["of:=SUM([.A1:.B2])", "SUM(A1:B2)"],
    ["of:=[$Sheet1.A1]", "Sheet1!A1"],
    ["of:=[$Sheet1.A1:.B2]", "Sheet1!A1:B2"],
    ["of:=[$'My Sheet'.A1]", "'My Sheet'!A1"],
    ["of:=[$'O''Brien'.A1]", "'O''Brien'!A1"],
    ["of:=SUM([$Sheet1.A1:$Sheet3.B2])", "SUM(Sheet1:Sheet3!A1:B2)"],
    ["of:=SUM([$Jan.B5:$Dec.B5])", "SUM(Jan:Dec!B5)"],
    ["of:=SUM([.A:.C])+SUM([.1:.3])+SUM([.$A:.$C])", "SUM(A:C)+SUM(1:3)+SUM($A:$C)"],
    ["of:=['file:///home/u/rates.ods'#$Sheet1.A1]", "['file:///home/u/rates.ods'#$Sheet1.A1]"],
    ["of:=IF([.A1]>0;1;2)", "IF(A1>0,1,2)"],
    ["of:=SUM(([.A1]~[.B1]))", "SUM((A1,B1))"],
    ["of:=SUM([.A1:.B2]![.B1:.C3])", "SUM(A1:B2 B1:C3)"],
    ["of:=SUM({1;2|3;4})", "SUM({1,2;3,4})"],
    ["of:=COM.MICROSOFT.CONCAT([.A1];[.B1])", "CONCAT(A1,B1)"],
    ["of:=LEGACY.CHIDIST([.A1];2)+FDIST([.A1];2;3;1)+FORMULA([.A1])", "CHIDIST(A1,2)+F.DIST(A1,2,3,1)+FORMULATEXT(A1)"],
    ["of:=ORG.OPENOFFICE.ERRORTYPE([.A1])", "ORG.OPENOFFICE.ERRORTYPE(A1)"],
  ])("%s ↔ %s", (wire, canonical) => {
    expect(fromOdsFormula(wire, LO_NAMESPACES)).toBe(canonical);
    expect(toOdsFormula(canonical)).toBe(wire);
  });

  it("from: a ~ union gains parentheses unless it already is a grouping paren's whole content", () => {
    expect(fromOdsFormula("of:=SUM([.A1]~[.B1])", LO_NAMESPACES)).toBe("SUM((A1,B1))");
    expect(fromOdsFormula("of:=SUM([.A1]~[.B1]~[.C1:.D2])", LO_NAMESPACES)).toBe("SUM((A1,B1,C1:D2))");
    expect(fromOdsFormula("of:=SUM(([.A1]~[.B1]))", LO_NAMESPACES)).toBe("SUM((A1,B1))");
    expect(fromOdsFormula("of:=-[.A1]~[.B1]", LO_NAMESPACES)).toBe("-(A1,B1)");
    expect(fromOdsFormula("of:=SUM([.A1:.B2]![.B1]~INDEX([.C1:.C9];2))", LO_NAMESPACES)).toBe("SUM((A1:B2 B1,INDEX(C1:C9,2)))");
    expect(fromOdsFormula("of:=SUM(([.A1]~[.B1])~[.C1])", LO_NAMESPACES)).toBe("SUM(((A1,B1),C1))");
    expect(fromOdsFormula("of:=[.A1]~[.B1]", LO_NAMESPACES)).toBe("(A1,B1)");
  });

  it("from: a relative sheet, an unanchored quoted sheet and a same-sheet second endpoint read as the plain sheet reference", () => {
    expect(fromOdsFormula("of:=[Sheet1.A1]", LO_NAMESPACES)).toBe("Sheet1!A1");
    expect(fromOdsFormula("of:=['My Sheet'.A1]", LO_NAMESPACES)).toBe("'My Sheet'!A1");
    expect(fromOdsFormula("of:=[$Sheet1.A1:$Sheet1.B2]", LO_NAMESPACES)).toBe("Sheet1!A1:B2");
    expect(fromOdsFormula("of:=[Sheet1.A1:Sheet3.B2]", LO_NAMESPACES)).toBe("Sheet1:Sheet3!A1:B2");
  });

  it("from: any bracket reference holding #REF! is the bare error", () => {
    expect(fromOdsFormula("of:=[.#REF!]+[$Sheet2.#REF!]+[$#REF!.A1]+[.A1:.#REF!]", LO_NAMESPACES)).toBe("#REF!+#REF!+#REF!+#REF!");
  });

  it("recognises the three dialects by namespace URI — whatever the prefix — and unprefixed text as OpenFormula", () => {
    const custom = new Map([
      ["x", OPENFORMULA_NAMESPACE],
      ["legacy", OOOC_NAMESPACE],
      ["excel", MSOXL_NAMESPACE],
    ]);
    expect(fromOdsFormula("x:=SUM([.A1];1)", custom)).toBe("SUM(A1,1)");
    expect(fromOdsFormula("legacy:=SUM([.A1];1)", custom)).toBe("SUM(A1,1)");
    expect(fromOdsFormula("excel:=SUM(A1,1)+_xlfn.CONCAT(B1)", custom)).toBe("SUM(A1,1)+CONCAT(B1)");
    expect(fromOdsFormula("=SUM([.A1];1)", custom)).toBe("SUM(A1,1)");
    expect(fromOdsFormula("oooc:=SUM([.B2:.B3])", LO_NAMESPACES)).toBe("SUM(B2:B3)");
    expect(fromOdsFormula("msoxl:=SUM(B2:B4)", LO_NAMESPACES)).toBe("SUM(B2:B4)");
    // A prefix bound to another URI is not OpenFormula, whatever it is called.
    expect(() => fromOdsFormula("of:=SUM([.A1])", new Map([["of", "urn:example:other"]]))).toThrow(FormulaSyntaxError);
    // An undeclared of:/oooc:/msoxl: falls back to its usual URI; anything else is unknown.
    expect(fromOdsFormula("of:=[.A1]", new Map())).toBe("A1");
    expect(() => fromOdsFormula("foo:=[.A1]", LO_NAMESPACES)).toThrow(FormulaSyntaxError);
  });

  it("refuses a bracket reference its grammar cannot read", () => {
    for (const wire of ["of:=[.A]", "of:=[A1]", "of:=[.A1:.B]", "of:=[.A1:$S.B2]", "of:=[.XFE1]", "of:=[$.A1]"]) {
      expect(() => fromOdsFormula(wire, LO_NAMESPACES), wire).toThrow(FormulaSyntaxError);
    }
  });

  it("to: of:= prefix, [.A1] same-sheet, [$Sheet.A1:.B2] another sheet, quoting by the shared predicate, separators, ~, ! and arrays", () => {
    expect(toOdsFormula("SUM('Q1 2026'!A1:B2,Data!$C$3,Jan:Mar!D4,A1:B2 C3,(A1,B1),{1,2;3,4})")).toBe(
      "of:=SUM([$'Q1 2026'.A1:.B2];[$Data.$C$3];[$Jan.D4:$Mar.D4];[.A1:.B2]![.C3];([.A1]~[.B1]);{1;2|3;4})",
    );
    expect(toOdsFormula("'A1'!B2+TRUE!C3")).toBe("of:=[$'A1'.B2]+[$'TRUE'.C3]");
    expect(toOdsFormula("Sheet2!#REF!+#REF!A1+#DIV/0!")).toBe("of:=#REF!+#REF!+#DIV/0!");
  });

  it("carries every token's lead, so whitespace survives both directions", () => {
    expect(fromOdsFormula("of:= SUM( [.A1] ;\n [.B1] ) ", LO_NAMESPACES)).toBe(" SUM( A1 ,\n B1 ) ");
    expect(toOdsFormula(" SUM( A1 ,\n B1 ) ")).toBe("of:= SUM( [.A1] ;\n [.B1] ) ");
  });

  it("round-trips every LibreOffice-shaped corpus formula exactly", () => {
    for (const wire of LIBREOFFICE_CORPUS) {
      expect(toOdsFormula(fromOdsFormula(wire, LO_NAMESPACES)), wire).toBe(wire);
    }
  });

  it("changes exactly the known non-identities — a relative sheet, a same-sheet endpoint, a ~ union, a #REF! reference", () => {
    for (const [wire, back] of ODS_NON_IDENTITIES) {
      const once = toOdsFormula(fromOdsFormula(wire, LO_NAMESPACES));
      expect(once, wire).not.toBe(wire);
      expect(once, wire).toBe(back);
      // …and the result is a fixed point.
      expect(toOdsFormula(fromOdsFormula(once, LO_NAMESPACES))).toBe(once);
    }
  });
});

describe("reading a formula — the parse check", () => {
  it("keeps the wire text, unparsed, when the translation does not parse: the xlsx <f> text, the whole ods value", () => {
    expect(readXlsxFormula("_xlfn.CONCAT(A1,B1)", { row: 0, col: 0 })).toEqual({ formula: "CONCAT(A1,B1)" });
    expect(readXlsxFormula("_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)", { row: 0, col: 0 })).toEqual({
      formula: "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)",
      formulaUnparsed: true,
    });
    expect(readOdsFormula("of:=SUM([.A1];1)", LO_NAMESPACES, { row: 0, col: 0 })).toEqual({ formula: "SUM(A1,1)" });
    expect(readOdsFormula("of:=SUM([.A1];", LO_NAMESPACES, { row: 0, col: 0 })).toEqual({ formula: "of:=SUM([.A1];", formulaUnparsed: true });
    expect(readOdsFormula("foo:=1", LO_NAMESPACES, { row: 0, col: 0 })).toEqual({ formula: "foo:=1", formulaUnparsed: true });
    expect(readOdsFormula("of:=[.A]", LO_NAMESPACES, { row: 0, col: 0 })).toEqual({ formula: "of:=[.A]", formulaUnparsed: true });
  });
});

describe("formulaForTarget — crossing formats", () => {
  it("lets a parsed formula cross unchanged (it is canonical)", () => {
    const cell: Cell = { col: 0, value: 60, formula: "SUM(B2:B4)", styleId: 3 };
    expect(formulaForTarget(cell, "ods")).toEqual({ cell });
    expect(formulaForTarget(cell, "ods").cell).toBe(cell);
    const array: Cell = { col: 0, value: 6, formula: "A2:A4*2", arrayRange: { top: 0, left: 0, bottom: 2, right: 0 } };
    expect(formulaForTarget(array, "xlsx").cell).toBe(array);
  });

  it("drops an unparsed formula and a data table to the cached value, reporting the drop", () => {
    const unparsed: Cell = { col: 0, value: 3, formula: "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)", formulaUnparsed: true, styleId: 2 };
    expect(formulaForTarget(unparsed, "ods")).toEqual({ cell: { col: 0, value: 3, styleId: 2 }, dropped: "unparsed" });
    const table: Cell = { col: 1, value: 10, dataTableXml: `<f t="dataTable" ref="B1:B3" r1="A1"/>` };
    expect(formulaForTarget(table, "ods")).toEqual({ cell: { col: 1, value: 10 }, dropped: "dataTable" });
    expect(unparsed.formula).toBeDefined();
  });

  it("lets a dynamic array cross to .ods as a matrix formula, without its cm", () => {
    const dynamic: Cell = { col: 7, value: 1, formula: "SORT(A2:A6)", arrayRange: { top: 1, left: 7, bottom: 5, right: 7 }, dynamicArray: true, cm: 1 };
    expect(formulaForTarget(dynamic, "ods")).toEqual({
      cell: { col: 7, value: 1, formula: "SORT(A2:A6)", arrayRange: { top: 1, left: 7, bottom: 5, right: 7 } },
      dropped: "dynamicArray",
    });
    expect(dynamic.cm).toBe(1);
  });
});
