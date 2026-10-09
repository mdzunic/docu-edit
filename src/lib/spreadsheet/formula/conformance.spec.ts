// SPEC-120 "Conformance": every row of the spec's table, each labelled with
// its source per "Where expected values come from" — **Excel** (Microsoft's
// documented behaviour: "Calculation operators and precedence in Excel"'s
// conversion table, the function pages and their examples) or **pinned**
// (computed by hand from the rule the spec states, where Excel's behaviour
// is undocumented). No expected value here was produced by running this
// engine. Then the `.ods` settings variants (case sensitivity, plain and
// regular-expression criteria, partial-cell matching), and spec120-calc.xlsx,
// whose column G holds the self-contained rows with Excel's cached values:
// F9 over it must reproduce every one.
import { describe, expect, it } from "vitest";
import { getCell, parseA1 } from "../address";
import { buildSpec120CalcXlsxBytes, SPEC120_CONFORMANCE } from "../fixtures/build";
import { calcWorkbook, type CellSpec } from "../fixtures/calcWorkbook";
import type { CalcSettings, CellValue } from "../model";
import { SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "../numberFormat/locale";
import { readSpreadsheetBytes } from "../open";
import { evaluateCell, isUncomputable, storedValueContext } from "./evaluate";
import { recalculate } from "./recalc";

type Source = "Excel" | "pinned";

interface Row {
  formula: string;
  expected: CellValue;
  source: Source;
  cells?: Record<string, CellSpec>;
  host?: string;
  locale?: SpreadsheetLocaleTag;
  label?: string;
}

const VALUE = { error: "#VALUE!" } as const;

const ROWS: Row[] = [
  { formula: '"1"+"2"', expected: 3, source: "Excel" },
  { formula: '"1"+"$4.00"', expected: 5, source: "Excel", label: "en-US" },
  { formula: '"6/1/2001"-"5/1/2001"', expected: 31, source: "Excel", label: "en-US" },
  { formula: 'SQRT("8+1")', expected: VALUE, source: "Excel" },
  { formula: '"A"&TRUE', expected: "ATRUE", source: "Excel" },
  { formula: '1+""', expected: VALUE, source: "Excel" },
  { formula: '"TRUE"+0', expected: VALUE, source: "Excel" },
  { formula: '"a"*1', expected: VALUE, source: "Excel" },
  { formula: "TRUE+TRUE", expected: 2, source: "Excel" },
  { formula: "A1+1", expected: 6, source: "Excel", cells: { A1: "5" }, label: 'A1 the text "5"' },
  { formula: "A1", expected: 0, source: "Excel", label: "A1 empty" },
  { formula: "A1=0", expected: true, source: "Excel", label: "A1 empty" },
  { formula: 'A1=""', expected: true, source: "Excel", label: "A1 empty" },
  { formula: "A1=FALSE", expected: true, source: "Excel", label: "A1 empty" },
  { formula: '1="1"', expected: false, source: "Excel" },
  { formula: '1<"1"', expected: true, source: "Excel" },
  { formula: '"z"<TRUE', expected: true, source: "Excel" },
  { formula: "FALSE<TRUE", expected: true, source: "Excel" },
  { formula: '"A"="a"', expected: true, source: "Excel" },
  { formula: '"a"<"B"', expected: true, source: "Excel" },
  { formula: "0.1+0.2=0.3", expected: true, source: "Excel" },
  { formula: "0.1+0.2-0.3", expected: 0, source: "pinned" },
  { formula: "-2^2", expected: 4, source: "Excel" },
  { formula: "0-2^2", expected: -4, source: "Excel" },
  { formula: "2^3^2", expected: 64, source: "Excel" },
  { formula: "0^0", expected: { error: "#NUM!" }, source: "Excel" },
  { formula: '1/3&""', expected: "0.333333333333333", source: "Excel" },
  { formula: "#N/A+1/0", expected: { error: "#N/A" }, source: "pinned" },
  { formula: "1/0+#N/A", expected: { error: "#DIV/0!" }, source: "pinned" },
  { formula: "IF(TRUE,1,1/0)", expected: 1, source: "Excel" },
  { formula: 'SUM("3",TRUE)', expected: 4, source: "Excel" },
  { formula: "SUM(A1:A2)", expected: 0, source: "Excel", cells: { A1: "3", A2: true }, label: 'A1 "3" text, A2 TRUE' },
  { formula: 'MAX("a")', expected: VALUE, source: "Excel" },
  { formula: "MIN(A1:A3)", expected: 0, source: "Excel", label: "all empty" },
  { formula: "AVERAGE(A1:A3)", expected: { error: "#DIV/0!" }, source: "Excel", label: "all empty" },
  { formula: "AND(A1:A2)", expected: true, source: "pinned", cells: { A1: "x", A2: true }, label: 'A1 "x" text, A2 TRUE' },
  { formula: 'AND("x")', expected: VALUE, source: "pinned" },
  { formula: "ISBLANK(A1)", expected: false, source: "Excel", cells: { A1: { formula: '""', value: "" } }, label: 'A1 =""' },
  { formula: "COUNTBLANK(A1:A2)", expected: 2, source: "Excel", cells: { A1: { formula: '""', value: "" } }, label: 'A1 ="", A2 empty' },
  { formula: "COUNTA(A1:A2)", expected: 1, source: "Excel", cells: { A1: { formula: '""', value: "" } }, label: 'A1 ="", A2 empty' },
  { formula: '"1,5"+0', expected: VALUE, source: "pinned", locale: "en-US" },
  { formula: '"1,5"+0', expected: 1.5, source: "pinned", locale: "de-DE" },
  { formula: '"2026-10-08"+0', expected: 46303, source: "pinned" },
  { formula: '"12:00"+0', expected: 0.5, source: "pinned" },
  { formula: '"(5)"+0', expected: -5, source: "pinned" },
  { formula: '"10%"+0', expected: 0.1, source: "pinned" },
  { formula: "ROUND(2.5,0)", expected: 3, source: "Excel" },
  { formula: "ROUND(-2.5,0)", expected: -3, source: "Excel" },
  { formula: "ROUND(1.005,2)", expected: 1.01, source: "pinned" },
  { formula: "MOD(-3,2)", expected: 1, source: "Excel" },
  { formula: "MOD(3,-2)", expected: -1, source: "Excel" },
  { formula: "INT(-8.9)", expected: -9, source: "Excel" },
  { formula: "TRUNC(-8.9)", expected: -8, source: "Excel" },
  { formula: "CEILING(-2.5,2)", expected: -2, source: "Excel" },
  { formula: "FLOOR(-2.5,-2)", expected: -2, source: "Excel" },
  { formula: "MROUND(10,3)", expected: 9, source: "Excel" },
  { formula: "FLOOR(0.3,0.1)", expected: 0.3, source: "pinned" },
  { formula: "DATE(108,1,2)", expected: 39449, source: "Excel", label: "2008-01-02" },
  { formula: "DATE(2008,14,2)", expected: 39846, source: "Excel", label: "2009-02-02" },
  { formula: "DATE(1900,2,29)", expected: 60, source: "Excel" },
  { formula: "WEEKDAY(1)", expected: 1, source: "Excel" },
  { formula: "WEEKDAY(46303)", expected: 5, source: "pinned" },
  { formula: 'VLOOKUP("a",{"c",1;"a",2},2,)', expected: 2, source: "Excel" },
  { formula: 'MATCH("B*",{"apple","banana"},0)', expected: 2, source: "Excel" },
  { formula: "COUNTIF(A1:A3,1)", expected: 2, source: "pinned", cells: { A1: "1", A2: 1, A3: 2 }, label: 'A1 the text "1", A2 1, A3 2' },
  { formula: 'LEN("😀")', expected: 2, source: "pinned" },
  { formula: "PROPER(\"it's\")", expected: "It'S", source: "pinned" },
  { formula: 'TRIM("  a   b  ")', expected: "a b", source: "Excel" },
  { formula: 'TEXT(0.5,"0%")', expected: "50%", source: "Excel" },
  { formula: "A2:A4", expected: 30, source: "Excel", cells: { A2: 20, A3: 30, A4: 40 }, host: "B3", label: "in row 3, non-array: A3's value" },
  { formula: "A2:A4", expected: VALUE, source: "Excel", cells: { A2: 20, A3: 30, A4: 40 }, host: "B9", label: "in row 9" },
];

/** The value a formula at `host` (default Z100) of Sheet1 stores. */
function computed(formula: string, cells: Record<string, CellSpec> = {}, host = "Z100", locale: SpreadsheetLocaleTag = "en-US", calc?: CalcSettings, format: "xlsx" | "ods" = "xlsx") {
  const model = calcWorkbook({ Sheet1: { ...cells, [host]: { formula } } }, { calc, format });
  const at = parseA1(host)!;
  return evaluateCell(model, "sheet-Sheet1", at.row, at.col, storedValueContext(model, { locale: SPREADSHEET_LOCALES[locale] }));
}

describe("the conformance table", () => {
  for (const row of ROWS) {
    const where = [row.label, row.locale].filter(Boolean).join(", ");
    it(`=${row.formula}${where ? ` (${where})` : ""} is ${JSON.stringify(row.expected)} — ${row.source}`, () => {
      expect(computed(row.formula, row.cells, row.host, row.locale)).toEqual(row.expected);
    });
  }

  it("{=A1:A2*{10,100}} over B1:C3: the 2×2 products, row 3 #N/A — Excel", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, A2: 2, B1: { formula: "A1:A2*{10,100}", array: "B1:C3" } } });
    recalculate(model, [], { full: true, reason: "manual" });
    const value = (a1: string) => getCell(model.sheets[0]!, parseA1(a1)!.row, parseA1(a1)!.col)?.value;
    expect([value("B1"), value("C1"), value("B2"), value("C2")]).toEqual([10, 100, 20, 200]);
    expect([value("B3"), value("C3")]).toEqual([{ error: "#N/A" }, { error: "#N/A" }]);
  });
});

describe("the .ods settings variants (SPEC-120 \"Calculation settings\")", () => {
  const cells: Record<string, CellSpec> = { A1: "Apple", A2: "apple", A3: "a.c", A4: "abc" };

  it("text compares case-sensitively when the workbook says so — pinned", () => {
    expect(computed('"A"="a"', {}, "Z100", "en-US", { caseSensitive: true }, "ods")).toBe(false);
    expect(computed('"A"<"a"', {}, "Z100", "en-US", { caseSensitive: true }, "ods")).toBe(false);
    expect(computed('COUNTIF(A1:A2,"apple")', cells, "Z100", "en-US", { caseSensitive: true }, "ods")).toBe(1);
    expect(computed('COUNTIF(A1:A2,"apple")', cells)).toBe(2);
  });

  it("plain criteria match literally — pinned", () => {
    expect(computed('COUNTIF(A1:A4,"a*")', cells, "Z100", "en-US", { criteria: "plain" }, "ods")).toBe(0);
    // With wildcards (xlsx), case-insensitively: every one of the four.
    expect(computed('COUNTIF(A1:A4,"a*")', cells)).toBe(4);
  });

  it("a regular-expression criterion matches literally without a special character, and is uncomputable with one — pinned", () => {
    expect(computed('COUNTIF(A1:A4,"abc")', cells, "Z100", "en-US", { criteria: "regex" }, "ods")).toBe(1);
    const uncomputable = computed('COUNTIF(A1:A4,"a.c")', cells, "Z100", "en-US", { criteria: "regex" }, "ods");
    expect(isUncomputable(uncomputable) && uncomputable.uncomputable).toEqual({ kind: "regex" });
  });

  it("partial-cell matching makes = text criteria match substrings — pinned", () => {
    expect(computed('COUNTIF(A1:A4,"pp")', cells, "Z100", "en-US", { partialCellMatch: true, caseSensitive: true }, "ods")).toBe(2);
    expect(computed('COUNTIF(A1:A4,"pp")', cells)).toBe(0);
  });

  it("an .ods workbook's February 1900 has 28 days — pinned", () => {
    expect(computed("DATE(1900,2,29)", {}, "Z100", "en-US", undefined, "ods")).toBe(61);
  });
});

describe("spec120-calc.xlsx's conformance rows (cached with Excel's values)", () => {
  it("F9 reproduces every cached value in column G", async () => {
    const model = await readSpreadsheetBytes(await buildSpec120CalcXlsxBytes(), "xlsx", "spec120-calc");
    const calc = model.sheets[0]!;
    const cached = SPEC120_CONFORMANCE.map((row) => getCell(calc, parseA1(row.ref)!.row, parseA1(row.ref)!.col)?.value);
    expect(cached).toEqual(SPEC120_CONFORMANCE.map((row) => row.value));
    recalculate(model, [], { full: true, reason: "manual", now: () => new Date(2026, 9, 8, 9, 30) });
    const after = model.sheets[0]!;
    expect(SPEC120_CONFORMANCE.map((row) => getCell(after, parseA1(row.ref)!.row, parseA1(row.ref)!.col)?.value)).toEqual(
      SPEC120_CONFORMANCE.map((row) => row.value),
    );
  });
});
