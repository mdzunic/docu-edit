// SPEC-120 "The core functions" — functions/aggregate.ts (over common.ts and
// criteria.ts): AVERAGE AVERAGEIF AVERAGEIFS COUNT COUNTA COUNTBLANK COUNTIF
// COUNTIFS MIN MAX MINIFS MAXIFS, each evaluated in a cell of an in-memory
// workbook. Rows labelled "Excel" are Microsoft's function-reference examples
// or documented rules; the others are pinned from SPEC-120's rules: numbers
// only from references and arrays but direct arguments coerced (MAX("a") is
// #VALUE!), AVERAGE* of nothing #DIV/0! and MIN/MAX/MINIFS/MAXIFS of nothing
// 0, AVERAGEIF resizing its average range, the *IFS shape rule (#VALUE!),
// COUNT ignoring errors, COUNTA counting "" results and errors, COUNTBLANK
// counting empties and "" results, criteria that match empty cells ("", "<>x")
// counted from the range's size — whole columns included — and the .ods
// criteria settings end to end.
import { describe, expect, it } from "vitest";
import { calcWorkbook, type CalcWorkbookOptions, type CellSpec } from "../../fixtures/calcWorkbook";
import type { CellValue, ErrorCode } from "../../model";
import { SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "../../numberFormat/locale";
import { evaluateCell, storedValueContext } from "../evaluate";

type Cells = Record<string, CellSpec>;

/** `formula` evaluated in Sheet1!Z100 over `cells` — the value it stores. */
function value(formula: string, cells: Cells = {}, options: { locale?: SpreadsheetLocaleTag; workbook?: CalcWorkbookOptions } = {}) {
  const model = calcWorkbook({ Sheet1: { ...cells, Z100: `=${formula}` } }, options.workbook);
  return evaluateCell(model, "sheet-Sheet1", 99, 25, storedValueContext(model, { locale: SPREADSHEET_LOCALES[options.locale ?? "en-US"] }));
}

const err = (code: ErrorCode): CellValue => ({ error: code });
/** A formula cell whose (cached) result is "". */
const EMPTY_TEXT: CellSpec = { formula: '""', value: "" };

describe("AVERAGE", () => {
  // Excel's AVERAGE reference example: A2:A6 10, 7, 9, 27, 2; B2 15; C2 32.
  const sheet: Cells = { A2: 10, A3: 7, A4: 9, A5: 27, A6: 2, B2: 15, C2: 32 };

  it.each<[string, CellValue]>([
    ["AVERAGE(A2:A6)", 11],
    ["AVERAGE(A2:A6,5)", 10],
    ["AVERAGE(A2:C2)", 19],
  ])("Excel: %s is %j", (formula, expected) => {
    expect(value(formula, sheet)).toEqual(expected);
  });

  it.each<[string, CellValue, Cells]>([
    ["AVERAGE(A1:A5)", 0.5, { A1: 1, A2: "x", A3: true, A4: 0 }], // text, logicals, empties skipped; zeros count (Excel)
    ['AVERAGE(1,TRUE,"4")', 2, {}], // direct logicals and numeric text count (Excel)
    ['AVERAGE("x")', err("#VALUE!"), {}],
    ["AVERAGE(A1:A3)", err("#DIV/0!"), {}],
    ["AVERAGE(A1:A2)", err("#DIV/0!"), { A1: "x", A2: false }],
    ["AVERAGE(A1:A2)", err("#N/A"), { A1: 1, A2: { error: "#N/A" } }],
  ])("%s is %j", (formula, expected, cells) => {
    expect(value(formula, cells)).toEqual(expected);
  });
});

describe("AVERAGEIF and AVERAGEIFS", () => {
  // Excel's AVERAGEIF reference, example 1: property values A2:A5, commissions B2:B5.
  const property: Cells = { A2: 100_000, A3: 200_000, A4: 300_000, A5: 400_000, B2: 7_000, B3: 14_000, B4: 21_000, B5: 28_000 };
  // Example 2: regions A2:A6, profits B2:B6.
  const regions: Cells = { A2: "East", B2: 45_678, A3: "West", B3: 23_789, A4: "North", B4: -4_789, A5: "South (New Office)", B5: 0, A6: "MidWest", B6: 9_678 };
  // Excel's AVERAGEIFS reference, example 1: quiz grades B2:C5 and exam grades D2:D5.
  const grades: Cells = {
    B2: 75, C2: 85, D2: 87,
    B3: 94, C3: 80, D3: 88,
    B4: 86, C4: 93, D4: "Incomplete",
    B5: "Incomplete", C5: 75, D5: 75,
  };
  // Example 2: prices B2:B7, towns C2:C7, bedrooms D2:D7, garages E2:E7.
  const homes: Cells = {
    B2: 230_000, C2: "Issaquah", D2: 3, E2: "No",
    B3: 197_000, C3: "Bellevue", D3: 2, E3: "Yes",
    B4: 345_678, C4: "Bellevue", D4: 4, E4: "Yes",
    B5: 321_900, C5: "Issaquah", D5: 2, E5: "Yes",
    B6: 450_000, C6: "Bellevue", D6: 5, E6: "Yes",
    B7: 395_000, C7: "Bellevue", D7: 4, E7: "No",
  };

  it.each<[string, CellValue, Cells]>([
    ['AVERAGEIF(B2:B5,"<23000")', 14_000, property],
    ['AVERAGEIF(A2:A5,"<250000")', 150_000, property],
    ['AVERAGEIF(A2:A5,"<95000")', err("#DIV/0!"), property],
    ['AVERAGEIF(A2:A5,">250000",B2:B5)', 24_500, property],
    ['AVERAGEIF(A2:A6,"=*West",B2:B6)', 16_733.5, regions],
    ['AVERAGEIF(A2:A6,"<>*(New Office)",B2:B6)', 18_589, regions],
    ['AVERAGEIFS(B2:B5,B2:B5,">70",B2:B5,"<90")', 80.5, grades],
    ['AVERAGEIFS(C2:C5,C2:C5,">95")', err("#DIV/0!"), grades],
    ['AVERAGEIFS(D2:D5,D2:D5,"<>Incomplete",D2:D5,">80")', 87.5, grades],
    ['AVERAGEIFS(B2:B7,C2:C7,"Bellevue",D2:D7,">2",E2:E7,"Yes")', 397_839, homes],
    ['AVERAGEIFS(B2:B7,C2:C7,"Issaquah",D2:D7,"<=3",E2:E7,"No")', 230_000, homes],
  ])("Excel: %s is %j", (formula, expected, cells) => {
    expect(value(formula, cells)).toEqual(expected);
  });

  it.each<[string, CellValue, Cells]>([
    // The average range resizes from its top-left cell to the range's shape (Excel).
    ['AVERAGEIF(A1:A4,">0",B1)', 20, { A1: 1, A2: -1, A3: 2, A4: -2, B1: 10, B2: 20, B3: 30, B4: 40 }],
    // An empty criteria cell is the criterion 0 (Excel).
    ["AVERAGEIF(A1:A3,C1,B1:B3)", 25, { A1: 0, A2: 1, A3: 0, B1: 10, B2: 20, B3: 40 }],
    // Empty and text cells of the average range are ignored (Excel).
    ['AVERAGEIF(A1:A3,">0",B1:B3)', 10, { A1: 1, A2: 1, A3: 1, B1: 10, B3: "x" }],
    // TRUE and FALSE in the range are ignored (Excel).
    ['AVERAGEIF(A1:A3,">=0")', 3, { A1: true, A2: 2, A3: 4 }],
    ['AVERAGEIF(A1:A2,"x")', err("#DIV/0!"), { A1: "x", A2: "x" }], // the matched cells hold no numbers
    ['AVERAGEIFS(A1:A3,B1:B4,">0")', err("#VALUE!"), { A1: 1, B1: 1 }],
    ['AVERAGEIFS(A1:B2,C1:C4,">0")', err("#VALUE!"), { A1: 1, C1: 1 }], // as many cells, another shape
  ])("%s is %j", (formula, expected, cells) => {
    expect(value(formula, cells)).toEqual(expected);
  });
});

describe("COUNT, COUNTA, COUNTBLANK", () => {
  // After Excel's COUNT/COUNTA reference data: "Sales", 12/8/2008, a blank, 19, 22.24, TRUE, #DIV/0!.
  const sheet: Cells = { A2: "Sales", A3: 39_790, A5: 19, A6: 22.24, A7: true, A8: { error: "#DIV/0!" } };

  it.each<[string, CellValue, Cells]>([
    ["COUNT(A2:A7)", 3, sheet], // Excel's reference examples
    ["COUNT(A5:A7)", 2, sheet],
    ["COUNT(A2:A7,2)", 4, sheet],
    ["COUNT(A2:A8)", 3, sheet], // an error in a reference is not counted
    ['COUNT(1,"2",TRUE,"x")', 3, {}], // direct numbers, numeric text and logicals count; other text does not (Excel)
    ["COUNT(1/0,5)", 1, {}], // nor a direct error (Excel)
    ["COUNT({1,TRUE,\"3\"})", 1, {}], // an array counts its numbers only
    ["COUNT(A:A)", 3, { A1: 1, A500: "x", A1048576: 2, A9: 3 }],
    ["COUNTA(A2:A9)", 7, { ...sheet, A9: EMPTY_TEXT }], // all but the blank A4: "" results and errors count (Excel)
    ["COUNTA(A1:A2)", 1, { A1: EMPTY_TEXT }],
    ["COUNTA(A1:A2)", 0, { A1: null }], // a stored cell with no value is empty
    ['COUNTA(1,"x",TRUE)', 3, {}],
    ["COUNTA(A:A)", 3, { A1: 1, A500: "x", A1048576: true }],
    ["COUNTBLANK(A1:A5)", 3, { A1: 0, A2: EMPTY_TEXT, A4: "x" }], // "" results count, zeros do not (Excel)
    ["COUNTBLANK(A1:A2)", 2, { A1: EMPTY_TEXT }],
    ["COUNTBLANK(A1:C3)", 7, { A1: 1, C3: "x" }],
    ["COUNTBLANK(A1:A2)", 1, { A1: null, A2: 5 }],
    ["COUNTBLANK(A:A)", 1_048_573, { A1: 1, A500: "x", A1048576: true }],
  ])("%s is %j", (formula, expected, cells) => {
    expect(value(formula, cells)).toEqual(expected);
  });
});

describe("COUNTIF and COUNTIFS", () => {
  // Excel's COUNTIF reference example: fruits A2:A5, quantities B2:B5.
  const fruit: Cells = { A2: "apples", B2: 32, A3: "oranges", B3: 54, A4: "peaches", B4: 75, A5: "apples", B5: 86 };
  // Excel's COUNTIFS reference, example 1: salespeople's quotas B2:D5.
  const quotas: Cells = {
    B2: "Yes", C2: "No", D2: "No",
    B3: "Yes", C3: "Yes", D3: "No",
    B4: "Yes", C4: "Yes", D4: "Yes",
    B5: "No", C5: "Yes", D5: "Yes",
  };
  // Example 2: 1–6 in A2:A7 beside 5/1/2011–5/6/2011 (serials 40664–40669) in B2:B7.
  const dated: Cells = { A2: 1, B2: 40_664, A3: 2, B3: 40_665, A4: 3, B4: 40_666, A5: 4, B5: 40_667, A6: 5, B6: 40_668, A7: 6, B7: 40_669 };

  it.each<[string, CellValue, Cells]>([
    ['COUNTIF(A2:A5,"apples")', 2, fruit],
    ["COUNTIF(A2:A5,A4)", 1, fruit],
    ["COUNTIF(A2:A5,A2)+COUNTIF(A2:A5,A3)", 3, fruit],
    ['COUNTIF(B2:B5,">55")', 2, fruit],
    ['COUNTIF(B2:B5,"<>"&B4)', 3, fruit],
    ['COUNTIF(B2:B5,">=32")-COUNTIF(B2:B5,">85")', 3, fruit],
    ['COUNTIF(A2:A5,"*")', 4, fruit],
    ['COUNTIF(A2:A5,"?????es")', 2, fruit],
    ['COUNTIFS(B2:D2,"=Yes")', 1, quotas],
    ['COUNTIFS(B2:B5,"=Yes",C2:C5,"=Yes")', 2, quotas],
    ['COUNTIFS(B5:D5,"=Yes",B3:D3,"=Yes")', 1, quotas],
    ['COUNTIFS(A2:A7,"<6",A2:A7,">1")', 4, dated],
    ['COUNTIFS(A2:A7,"<5",B2:B7,"<5/3/2011")', 2, dated],
    ['COUNTIFS(A2:A7,"<"&A6,B2:B7,"<"&B4)', 2, dated],
  ])("Excel: %s is %j", (formula, expected, cells) => {
    expect(value(formula, cells)).toEqual(expected);
  });

  it.each<[string, CellValue, Cells]>([
    ['COUNTIF(A1:A10,"")', 7, { A1: "x", A2: 1, A3: true }], // empty positions come from the range's size
    ['COUNTIF(A1:A10,"<>x")', 9, { A1: "x", A2: "y", A3: 1 }], // empty cells are not "x"
    ['COUNTIF(A1:C3,"")', 7, { A1: 1, C3: "x" }],
    ['COUNTIF(A1:A3,"")', 2, { A1: EMPTY_TEXT, A3: "x" }], // empty cells and "" results
    ['COUNTIF(A1:A3,"=")', 1, { A1: EMPTY_TEXT, A3: "x" }], // empty cells only
    ['COUNTIF(A1:A3,"<>")', 2, { A1: EMPTY_TEXT, A3: "x" }], // every non-empty cell
    ["COUNTIF(A1:A3,1)", 2, { A1: "1", A2: 1, A3: 2 }], // numeric text matches a number
    ["COUNTIF(A1:A3,C1)", 1, { A1: 0, A3: 1 }], // an empty criteria cell is 0, which no empty cell matches
    ['COUNTIF(A1:A3,"APPLES")', 2, { A1: "apples", A2: "Apples", A3: "pears" }],
    ['COUNTIF(A1:A4,">0")', 2, { A1: 1, A2: { error: "#N/A" }, A3: "5", A4: 2 }], // errors and text are skipped
    ['COUNTIF(A1:A3,"#N/A")', 1, { A1: { error: "#N/A" }, A2: { error: "#DIV/0!" }, A3: "x" }],
    ['COUNTIF(A1:A3,">=10/8/2026")', 2, { A1: 46_302, A2: 46_303, A3: 46_304 }], // 10/8/2026 is serial 46303
    ['COUNTIF(A:A,"<>apples")', 1_048_574, { A1: "apples", A2: "pears", A1000: "apples" }],
    ['COUNTIF(A:A,">1")', 2, { A1: 1, A700: 2, A1048576: 3 }],
    ['COUNTIFS(A1:A5,"",B1:B5,"")', 2, { A1: "x", B2: "y", A3: "z" }], // positions empty in both
    ['COUNTIFS(A1:A5,"<>x",B1:B5,"")', 3, { A1: "x", B2: "y", A3: "z" }],
    ['COUNTIFS(A1:A5,"<>x",B1:B5,"y")', 1, { A1: "x", B2: "y", A3: "z" }],
    ['COUNTIFS(A1:A3,">0",B1:B4,">0")', err("#VALUE!"), { A1: 1, B1: 1 }],
    ['COUNTIFS(A1:B2,">0",C1:C4,">0")', err("#VALUE!"), { A1: 1, C1: 1 }], // as many cells, another shape
  ])("%s is %j", (formula, expected, cells) => {
    expect(value(formula, cells)).toEqual(expected);
  });

  it("reads a numeric criterion in the user's locale", () => {
    expect(value('COUNTIF(A1:A2,"1,5")', { A1: 1.5, A2: "1,5" }, { locale: "de-DE" })).toBe(2);
    expect(value('COUNTIF(A1:A2,"1,5")', { A1: 1.5, A2: "1,5" })).toBe(1); // en-US: the text "1,5" only
  });
});

/** The `value` options of an .ods workbook with these calculation settings. */
function ods(calc: CalcWorkbookOptions["calc"]): { workbook: CalcWorkbookOptions } {
  return { workbook: { format: "ods", calc } };
}

describe("COUNTIF in an .ods workbook", () => {
  it("matches case-sensitively when the file says so", () => {
    expect(value('COUNTIF(A1:A2,"apple")', { A1: "apple", A2: "Apple" }, ods({ caseSensitive: true }))).toBe(1);
  });

  it("matches substrings when the whole cell need not match", () => {
    expect(value('COUNTIF(A1:A3,"pp")', { A1: "apple", A2: "banana", A3: "pepper" }, ods({ partialCellMatch: true }))).toBe(2);
  });

  it("reads plain criteria literally, without wildcards", () => {
    expect(value('COUNTIF(A1:A3,"a*")', { A1: "a*", A2: "apple", A3: "A*" }, ods({ criteria: "plain" }))).toBe(2);
  });

  it("computes a regular-expression criterion without metacharacters, and leaves any other uncomputable", () => {
    const cells: Cells = { A1: "apple", A2: "Apple", A3: "apples" };
    expect(value('COUNTIF(A1:A3,"apple")', cells, ods({ criteria: "regex" }))).toBe(2);
    expect(value('COUNTIF(A1:A3,"app.*")', cells, ods({ criteria: "regex" }))).toEqual({ uncomputable: { kind: "regex" } });
  });
});

describe("MIN, MAX, MINIFS, MAXIFS", () => {
  // Excel's MIN/MAX reference example: A2:A6 10, 7, 9, 27, 2.
  const sheet: Cells = { A2: 10, A3: 7, A4: 9, A5: 27, A6: 2 };
  // Excel's MINIFS/MAXIFS reference, example 1: grades A2:A7, weights B2:B7.
  const weighted: Cells = { A2: 89, B2: 1, A3: 93, B3: 2, A4: 96, B4: 2, A5: 85, B5: 3, A6: 91, B6: 1, A7: 88, B7: 1 };

  it.each<[string, CellValue, Cells]>([
    ["MIN(A2:A6)", 2, sheet],
    ["MIN(A2:A6,0)", 0, sheet],
    ["MAX(A2:A6)", 27, sheet],
    ["MAX(A2:A6,30)", 30, sheet],
    ["MINIFS(A2:A7,B2:B7,1)", 88, weighted],
    ["MAXIFS(A2:A7,B2:B7,1)", 91, weighted],
  ])("Excel: %s is %j", (formula, expected, cells) => {
    expect(value(formula, cells)).toEqual(expected);
  });

  it.each<[string, CellValue, Cells]>([
    ['MAX("a")', err("#VALUE!"), {}],
    ["MIN(A1:A3)", 0, {}], // of nothing
    ["MAX(A1:A2)", 0, { A1: "x", A2: true }],
    ["MAX(A1:A3)", -5, { A1: -5, A2: "100", A3: true }], // a reference reads numbers only (Excel)
    ["MAX(-5,TRUE)", 1, {}], // a direct logical counts (Excel)
    ['MIN("3",5)', 3, {}], // as does direct numeric text (Excel)
    ["MIN({4,\"1\",FALSE,2})", 2, {}], // an array reads numbers only
    ["MAX(A1:A2)", err("#N/A"), { A1: 1, A2: { error: "#N/A" } }],
    ["MAX(A:A)", 3, { A1: 1, A1048576: 3, A9: -1 }],
    ["MINIFS(A2:A7,B2:B7,\"c\")", 0, weighted], // nothing matches
    ["MAXIFS(A2:A7,B2:B7,\"c\")", 0, weighted],
    ["MAXIFS(A2:A7,B2:B7,1,A2:A7,\"<90\")", 89, weighted],
    // Ranges of one shape need not be aligned: A2:A5 pairs with B3:B6 by position.
    ['MINIFS(A2:A5,B3:B6,"a")', 10, { A2: 10, A3: 11, A4: 100, A5: 111, B3: "a", B4: "a", B5: "b", B6: "a" }],
    ['MAXIFS(A2:A5,B3:B6,"a")', 111, { A2: 10, A3: 11, A4: 100, A5: 111, B3: "a", B4: "a", B5: "b", B6: "a" }],
    ['MINIFS(A1:A3,B1:B4,">0")', err("#VALUE!"), { A1: 1, B1: 1 }],
    ['MAXIFS(A1:B2,C1:C4,">0")', err("#VALUE!"), { A1: 1, C1: 1 }],
  ])("%s is %j", (formula, expected, cells) => {
    expect(value(formula, cells)).toEqual(expected);
  });
});
