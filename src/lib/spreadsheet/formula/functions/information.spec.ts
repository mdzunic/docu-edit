// SPEC-120 "The core functions" (information.ts): ISBLANK ISNUMBER ISTEXT
// ISERROR ISERR ISNA ISLOGICAL NA. The IS functions accept errors — an error
// argument is what they ask about, never a result they pass on, so
// ISERROR(1/0) is TRUE and ISNUMBER(1/0) FALSE; they coerce nothing, so the
// text "5" is no number; ISBLANK is TRUE only for an empty cell — a cell
// holding the formula result "" is not blank (the conformance row).
//
// Sources ("Where expected values come from"): "docs" rows are the IS
// functions' reference-page examples (the "Gold"/#REF!/330.92/#N/A table);
// the rest follow from the definitions there (ISERR: any error except #N/A)
// and SPEC-120's rules.
import { describe, expect, it } from "vitest";
import { parseA1 } from "../../address";
import { calcWorkbook, type CellSpec } from "../../fixtures/calcWorkbook";
import type { CellError, CellValue, ErrorCode } from "../../model";
import { evaluateCell, storedValueContext } from "../evaluate";

type Cells = Record<string, CellSpec>;

const err = (code: ErrorCode): CellError => ({ error: code });

/** What `formula`, in Z100 of Sheet1 beside `cells`, stores. */
function value(formula: string, cells: Cells = {}) {
  const host = "Z100";
  const model = calcWorkbook({ Sheet1: { ...cells, [host]: { formula } } });
  const at = parseA1(host)!;
  return evaluateCell(model, "sheet-Sheet1", at.row, at.col, storedValueContext(model));
}

describe("Microsoft's IS examples", () => {
  it("docs: literal arguments", () => {
    expect(value("ISLOGICAL(TRUE)")).toBe(true);
    expect(value('ISLOGICAL("TRUE")')).toBe(false);
    expect(value("ISNUMBER(4)")).toBe(true);
  });

  const cells: Cells = { A2: "Gold", A3: "Region1", A4: err("#REF!"), A5: 330.92, A6: err("#N/A") };
  it.each<[formula: string, expected: boolean]>([
    ["ISBLANK(A2)", false],
    ["ISERROR(A4)", true],
    ["ISNA(A4)", false],
    ["ISNA(A6)", true],
    ["ISERR(A6)", false],
    ["ISNUMBER(A5)", true],
    ["ISTEXT(A3)", true],
  ])("docs: =%s is %s", (formula, expected) => {
    expect(value(formula, cells)).toBe(expected);
  });
});

describe("ISBLANK", () => {
  it.each<[what: string, a1: CellSpec | undefined, expected: boolean]>([
    ["an empty cell", undefined, true],
    ['a cell holding the formula result ""', { formula: '""', value: "" }, false], // conformance row
    ["the text ' '", " ", false],
    ["0", 0, false],
    ["FALSE", false, false],
    ["an error", err("#DIV/0!"), false],
  ])("ISBLANK(A1) with A1 %s is %s", (_what, a1, expected) => {
    expect(value("ISBLANK(A1)", a1 === undefined ? {} : { A1: a1 })).toBe(expected);
  });

  it('is FALSE for a direct "" — text, not an empty cell', () => {
    expect(value('ISBLANK("")')).toBe(false);
  });
});

describe("ISNUMBER, ISTEXT and ISLOGICAL coerce nothing", () => {
  it.each<[formula: string, cells: Cells, expected: boolean]>([
    ['ISNUMBER("5")', {}, false],
    ["ISNUMBER(A1)", { A1: "5" }, false], // the text "5"
    ["ISNUMBER(A1)", { A1: 5 }, true],
    ["ISNUMBER(TRUE)", {}, false],
    ["ISNUMBER(A1)", {}, false], // an empty cell is no number here
    ["ISNUMBER(DATE(2026,10,8))", {}, true], // a date is a serial
    ['ISTEXT("")', {}, true],
    ["ISTEXT(A1)", { A1: { formula: '""', value: "" } }, true],
    ["ISTEXT(A1)", { A1: "5" }, true],
    ["ISTEXT(A1)", {}, false],
    ["ISTEXT(5)", {}, false],
    ["ISLOGICAL(A1)", { A1: false }, true],
    ["ISLOGICAL(1)", {}, false],
    ["ISLOGICAL(A1)", {}, false],
  ])("=%s over %j is %s", (formula, cells, expected) => {
    expect(value(formula, cells)).toBe(expected);
  });
});

describe("the IS functions accept errors", () => {
  it.each<[formula: string, expected: boolean]>([
    ["ISERROR(1/0)", true],
    ["ISNUMBER(1/0)", false],
    ["ISTEXT(NA())", false],
    ["ISLOGICAL(1/0)", false],
    ["ISBLANK(NA())", false],
  ])("=%s is %s, not the error", (formula, expected) => {
    expect(value(formula)).toBe(expected);
  });

  it.each<[argument: string, isError: boolean, isErr: boolean, isNa: boolean]>([
    ["1/0", true, true, false],
    ["NA()", true, false, true],
    ["SQRT(-1)", true, true, false],
    ["1", false, false, false],
    ['"#N/A"', false, false, false], // text that spells an error is text
  ])("ISERROR/ISERR/ISNA(%s) are %s/%s/%s", (argument, isError, isErr, isNa) => {
    expect(value(`ISERROR(${argument})`)).toBe(isError);
    expect(value(`ISERR(${argument})`)).toBe(isErr);
    expect(value(`ISNA(${argument})`)).toBe(isNa);
  });

  it.each<ErrorCode>(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"])(
    "a cell holding %s: ISERROR TRUE, ISERR unless #N/A, ISNA only for #N/A",
    (code) => {
      const cells: Cells = { A1: err(code) };
      expect(value("ISERROR(A1)", cells)).toBe(true);
      expect(value("ISERR(A1)", cells)).toBe(code !== "#N/A");
      expect(value("ISNA(A1)", cells)).toBe(code === "#N/A");
    },
  );
});

describe("in an array formula", () => {
  it("the IS functions test each element", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, A3: "x", B1: { formula: "ISBLANK(A1:A3)", array: "B1:B3" } } });
    expect(evaluateCell(model, "sheet-Sheet1", 0, 1, storedValueContext(model))).toEqual({ kind: "array", rows: 3, cols: 1, data: [false, true, false] });
  });
});

describe("NA", () => {
  it("returns #N/A, which propagates like any error", () => {
    expect(value("NA()")).toEqual(err("#N/A"));
    expect(value("NA()+1")).toEqual(err("#N/A"));
    expect(value("SUM(1,NA())")).toEqual(err("#N/A"));
    expect(value("ISNA(NA())")).toBe(true);
  });

  it.each<[formula: string, expected: CellValue]>([
    ['IF(ISNA(MATCH("z",{"a","b"},0)),"missing","found")', "missing"],
    ['IF(ISNA(MATCH("b",{"a","b"},0)),"missing","found")', "found"],
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});
