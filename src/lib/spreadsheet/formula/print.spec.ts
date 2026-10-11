// SPEC-119 "Canonical text", "Sheet-name quoting", "The formula bar":
// printFormula is the oracle — parseFormula(printFormula(parseFormula(t)))
// deep-equals parseFormula(t), token spans aside, for every corpus formula —
// and prints canonical text: no whitespace but the intersection space, the
// minimal parentheses (a built AST gets exactly the ones precedence needs),
// references through formatReference with sheet names quoted per the rule.
// formulaBarText gives every row of the formula-bar table.
import { describe, expect, it } from "vitest";
import { CANONICAL_CORPUS } from "../fixtures/formulaCorpus";
import type { Cell, CellStyle, Sheet } from "../model";
import { SPREADSHEET_LOCALES } from "../numberFormat/locale";
import type { FormulaNode } from "./ast";
import { parseFormula } from "./parse";
import {
  dataTableText,
  formatReference,
  formatSheetPrefix,
  formulaBarContent,
  formulaBarText,
  printFormula,
  quoteSheetName,
  sheetNameNeedsQuotes,
  type FormulaBarContext,
} from "./print";
import type { ReferenceNode } from "./references";

/** A node with every token span removed, for span-blind equality. */
function withoutSpans(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(withoutSpans);
  }
  if (typeof value === "object" && value !== null) {
    return Object.fromEntries(Object.entries(value).filter(([key]) => key !== "from" && key !== "to").map(([key, child]) => [key, withoutSpans(child)]));
  }
  return value;
}

const at = { from: 0, to: 0 };
const num = (value: number): FormulaNode => ({ kind: "number", value, ...at });
const ref = (text: string): FormulaNode => {
  const node = parseFormula(text);
  return node;
};
const binary = (op: "+" | "-" | "*" | "^" | "," | " " | ":" | "&" | "=", left: FormulaNode, right: FormulaNode): FormulaNode => ({ kind: "binary", op, left, right, ...at });
const unary = (op: "+" | "-" | "%", operand: FormulaNode): FormulaNode => ({ kind: "unary", op, operand, ...at });
const call = (name: string, ...args: FormulaNode[]): FormulaNode => ({ kind: "function", name, args, ...at });

/** A reference part (0-based; absent row/col = whole column/row). */
function part(row: number | undefined, col: number | undefined, rowAbs = false, colAbs = false) {
  return { ...(row === undefined ? {} : { row }), ...(col === undefined ? {} : { col }), rowAbs, colAbs };
}

/** A grid sheet holding `cells` ([row, cell] pairs). */
function sheetOf(cells: Array<[number, Cell]>): Sheet {
  const rows: Sheet["rows"] = [];
  for (const [index, cell] of cells.toSorted((a, b) => a[0] - b[0] || a[1].col - b[1].col)) {
    const last = rows.at(-1);
    if (last?.index === index) {
      last.cells.push(cell);
    } else {
      rows.push({ index, cells: [cell] });
    }
  }
  return { conditionalFormats: [], id: "s", name: "S", visibility: "visible", kind: "grid", rows, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] };
}

describe("printFormula is the parser's oracle", () => {
  it("parse(print(parse(t))) deep-equals parse(t), spans aside, for every corpus formula", () => {
    for (const text of CANONICAL_CORPUS) {
      const parsed = parseFormula(text);
      const printed = printFormula(parsed);
      expect(withoutSpans(parseFormula(printed)), `${text} → ${printed}`).toEqual(withoutSpans(parsed));
    }
  });

  it("prints canonical text: no whitespace but the intersection space, upper-case references and functions, numbers as values", () => {
    expect(printFormula(parseFormula("  sum( a1 , $b$2 ) "))).toBe("SUM(A1,$B$2)");
    expect(printFormula(parseFormula("SUM(A1:C3   B2:D4)"))).toBe("SUM(A1:C3 B2:D4)");
    expect(printFormula(parseFormula("IF(A1>0,\r\n  1,\r\n  .5)"))).toBe("IF(A1>0,1,0.5)");
    expect(printFormula(parseFormula("1E3+1e-3+1.5E+10+1E21"))).toBe("1000+0.001+15000000000+1E+21");
    expect(printFormula(parseFormula(`"He said ""hi"""&A1`))).toBe(`"He said ""hi"""&A1`);
    expect(printFormula(parseFormula("{1,-2;\"a\",#N/A}"))).toBe(`{1,-2;"a",#N/A}`);
    expect(printFormula(parseFormula("IF(A1,,1)"))).toBe("IF(A1,,1)");
    expect(printFormula(parseFormula("'Sheet 1'!Rate+Sheet1!Rate"))).toBe("'Sheet 1'!Rate+Sheet1!Rate");
    expect(printFormula(parseFormula("Table1[[#This Row],[Amount]]+[1]Sheet1!A1+#REF!A1"))).toBe("Table1[[#This Row],[Amount]]+[1]Sheet1!A1+#REF!A1");
  });

  it("gives a built AST exactly the parentheses precedence, left associativity and argument commas need", () => {
    expect(printFormula(binary("*", binary("+", num(1), num(2)), num(3)))).toBe("(1+2)*3");
    expect(printFormula(binary("+", num(1), binary("*", num(2), num(3))))).toBe("1+2*3");
    expect(printFormula(binary("-", num(1), binary("-", num(2), num(3))))).toBe("1-(2-3)");
    expect(printFormula(binary("-", binary("-", num(1), num(2)), num(3)))).toBe("1-2-3");
    expect(printFormula(binary("^", num(2), binary("^", num(3), num(2))))).toBe("2^(3^2)");
    expect(printFormula(binary("^", unary("-", num(2)), num(2)))).toBe("-2^2");
    expect(printFormula(unary("-", binary("^", num(2), num(2))))).toBe("-(2^2)");
    expect(printFormula(unary("%", binary("^", num(2), num(3))))).toBe("(2^3)%");
    expect(printFormula(binary("^", num(2), unary("%", num(3))))).toBe("2^3%");
    expect(printFormula(binary("^", num(2), unary("-", num(2))))).toBe("2^-2");
    expect(printFormula(unary("-", ref("A1:A3")))).toBe("-A1:A3");
    expect(printFormula(binary("=", binary("=", ref("A1"), ref("B1")), ref("C1")))).toBe("A1=B1=C1");
    expect(printFormula(binary("=", ref("A1"), binary("=", ref("B1"), ref("C1"))))).toBe("A1=(B1=C1)");
    expect(printFormula(call("SUM", binary(",", ref("A1"), ref("B1"))))).toBe("SUM((A1,B1))");
    expect(printFormula(call("SUM", unary("-", binary(",", ref("A1"), ref("B1")))))).toBe("SUM(-(A1,B1))");
    expect(printFormula(binary(",", ref("A1"), ref("B1")))).toBe("A1,B1");
    expect(printFormula(binary(" ", binary(",", ref("A1"), ref("B1")), ref("C1")))).toBe("(A1,B1) C1");
    expect(printFormula(binary("&", binary("+", num(1), num(2)), num(3)))).toBe("1+2&3");
  });
});

describe("sheet-name quoting", () => {
  it.each([
    ["Sheet1", false],
    ["Data_2026", false],
    ["Übersicht", false],
    ["_hidden", false],
    ["Sheet.1", false],
    ["My Sheet", true],
    ["O'Brien", true],
    ["2026", true],
    ["Sheet-1", true],
    ["A1", true],
    ["xfd1048576", true],
    ["XFE1", false],
    ["R", true],
    ["c", true],
    ["R1", true],
    ["C1", true],
    ["RC", true],
    ["r1c1", true],
    ["RC1", true],
    ["Rate", false],
    ["TRUE", true],
    ["false", true],
    ["", true],
  ])("%s needs quotes: %s", (name, quoted) => {
    expect(sheetNameNeedsQuotes(name)).toBe(quoted);
  });

  it("quotes with every ' doubled, and a 3-D prefix whole", () => {
    expect(quoteSheetName("O'Brien")).toBe("'O''Brien'");
    expect(quoteSheetName("Sheet1")).toBe("Sheet1");
    expect(formatSheetPrefix("Jan", "Mar")).toBe("Jan:Mar!");
    expect(formatSheetPrefix("Q1 2026", "Q4 2026")).toBe("'Q1 2026:Q4 2026'!");
    expect(formatSheetPrefix("Jan", "Q4 2026")).toBe("'Jan:Q4 2026'!");
  });
});

describe("formatReference", () => {
  it.each<[ReferenceNode, string]>([
    [{ start: part(2, 1) }, "B3"],
    [{ start: part(0, 0, true, true) }, "$A$1"],
    [{ start: part(0, 0, true, false), end: part(9, 2, false, true) }, "A$1:$C10"],
    [{ start: part(undefined, 0), end: part(undefined, 2, false, true) }, "A:$C"],
    [{ start: part(undefined, 3) }, "D:D"],
    [{ start: part(0, undefined, true), end: part(4, undefined) }, "$1:5"],
    [{ sheet: "Sheet1", start: part(0, 0) }, "Sheet1!A1"],
    [{ sheet: "O'Brien", start: part(0, 0) }, "'O''Brien'!A1"],
    [{ sheet: "Jan", sheetEnd: "Mar", start: part(1, 1) }, "Jan:Mar!B2"],
    [{ sheet: "Q1 2026", sheetEnd: "Q4 2026", start: part(2, 2), end: part(3, 3) }, "'Q1 2026:Q4 2026'!C3:D4"],
    [{ start: part(1_048_575, 16_383) }, "XFD1048576"],
  ])("%j → %s", (reference, text) => {
    expect(formatReference(reference)).toBe(text);
  });
});

describe("formulaBarText", () => {
  const ctx: FormulaBarContext = { locale: SPREADSHEET_LOCALES["en-US"], dateSystem: "1900" };
  const sheet = sheetOf([
    [0, { col: 0, value: 60, formula: "SUM(B2:B4)" }],
    [0, { col: 1, value: 6, formula: "A2:A4*2", arrayRange: { top: 0, left: 1, bottom: 2, right: 1 } }],
    [1, { col: 1, value: 2 }],
    [0, { col: 2, value: 1, formula: "SORT(A2:A6)", arrayRange: { top: 0, left: 2, bottom: 4, right: 2 }, dynamicArray: true, cm: 1 }],
    [3, { col: 2, value: 4 }],
    [0, { col: 3, value: 10, dataTableXml: `<f t="dataTable" ref="D1:E3" dt2D="1" dtr="1" r1="H1" r2="H2"/>` }],
    [2, { col: 4, value: 30 }],
    [5, { col: 0, value: 3, formula: "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)", formulaUnparsed: true }],
    [5, { col: 1, value: 3, formula: "of:=FOO([.A1]", formulaUnparsed: true }],
    [6, { col: 0, value: 46303.5729166667 }],
    [6, { col: 1, value: "text" }],
    [7, { col: 0, value: null, formula: "B2+1" }],
  ]);
  const dateStyle: CellStyle = { numberFormat: "m/d/yyyy", builtinFormatId: 14 };

  it('shows "=" + formula for a formula cell (no cached value included)', () => {
    expect(formulaBarText(sheet, 0, 0, undefined, ctx)).toBe("=SUM(B2:B4)");
    expect(formulaBarText(sheet, 7, 0, undefined, ctx)).toBe("=B2+1");
  });

  it('shows "{=" + anchor formula + "}" for every cell of a legacy array, members included — even an empty member', () => {
    expect(formulaBarText(sheet, 0, 1, undefined, ctx)).toBe("{=A2:A4*2}");
    expect(formulaBarText(sheet, 1, 1, undefined, ctx)).toBe("{=A2:A4*2}");
    expect(formulaBarText(sheet, 2, 1, undefined, ctx)).toBe("{=A2:A4*2}");
  });

  it("shows a dynamic array's anchor as its formula, and its members the anchor's formula, muted", () => {
    expect(formulaBarContent(sheet, 0, 2, undefined, ctx)).toEqual({ text: "=SORT(A2:A6)", spillMember: false });
    expect(formulaBarContent(sheet, 3, 2, undefined, ctx)).toEqual({ text: "=SORT(A2:A6)", spillMember: true });
    expect(formulaBarContent(sheet, 4, 2, undefined, ctx)).toEqual({ text: "=SORT(A2:A6)", spillMember: true });
  });

  it("shows {=TABLE(r1,r2)} for a data table range; one-dimensional: {=TABLE(r1,)} with dtr=1, else {=TABLE(,r1)}", () => {
    expect(formulaBarText(sheet, 0, 3, undefined, ctx)).toBe("{=TABLE(H1,H2)}");
    expect(formulaBarText(sheet, 2, 4, undefined, ctx)).toBe("{=TABLE(H1,H2)}");
    expect(dataTableText(`<f t="dataTable" ref="B2:B4" dt2D="0" dtr="1" r1="A1"/>`)).toBe("{=TABLE(A1,)}");
    expect(dataTableText(`<f t="dataTable" ref="B2:B4" dt2D="0" dtr="0" r1="A1"/>`)).toBe("{=TABLE(,A1)}");
    expect(dataTableText(`<f t="dataTable" ref='B2:B4' r1="A1" ca="1"/>`)).toBe("{=TABLE(,A1)}");
    expect(dataTableText(`<f t="dataTable" ref="B2:D4" dt2D="1" dtr="0" r1="A1" r2="A2" del2="1"/>`)).toBe("{=TABLE(A1,#REF!)}");
  });

  it('shows an unparsed formula as stored: "=" + the xlsx <f> text; an ods value from its first "="', () => {
    expect(formulaBarText(sheet, 5, 0, undefined, ctx)).toBe("=_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)");
    expect(formulaBarText(sheet, 5, 1, undefined, { ...ctx, format: "ods" })).toBe("=FOO([.A1]");
  });

  it("shows SPEC-118's editText for every other cell, through the passed style", () => {
    expect(formulaBarText(sheet, 6, 0, dateStyle, ctx)).toBe("10/8/2026 1:45:00 PM");
    expect(formulaBarText(sheet, 6, 0, dateStyle, { ...ctx, locale: SPREADSHEET_LOCALES["de-DE"] })).toBe("08.10.2026 13:45:00");
    expect(formulaBarText(sheet, 6, 1, { quotePrefix: true }, ctx)).toBe("'text");
    expect(formulaBarText(sheet, 50, 50, undefined, ctx)).toBe("");
  });
});
