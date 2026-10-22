// SPEC-120 "The core functions" (logical.ts) with "Coercion" and "Errors"
// as they bind IF IFS IFERROR IFNA AND OR NOT XOR TRUE FALSE SWITCH CHOOSE:
// the six lazy functions evaluate only the branch they take — proved twice,
// by an untaken branch that would be an error and by an untaken RAND() that
// never draws from ctx.random — IF's omitted (FALSE) and empty (0) branches
// and a branch returning a reference, IFS's #N/A, SWITCH's default and #N/A,
// CHOOSE's truncated index (#VALUE! outside 1..n); AND/OR/XOR read logicals
// and numbers from references and arrays (text and empties skipped), coerce
// direct arguments logically, are #VALUE! with nothing to read, and pass on
// the first error. In an array formula IF, IFERROR and CHOOSE apply
// element-wise ({=SUM(IF(A1:A3>1,A1:A3))}).
//
// Sources ("Where expected values come from"): "docs" rows are Microsoft's
// function-reference examples, recomputed by hand from the formula and its
// data; "pinned" rows follow from the rule SPEC-120 states.
import { describe, expect, it } from "vitest";
import { parseA1 } from "../../address";
import { calcWorkbook, type CellSpec } from "../../fixtures/calcWorkbook";
import type { CellError, CellValue, ErrorCode } from "../../model";
import { evaluateCell, storedValueContext, type EvalContext } from "../evaluate";

type Cells = Record<string, CellSpec>;
type Context = Omit<EvalContext, "host" | "origin" | "arrayMode">;

const err = (code: ErrorCode): CellError => ({ error: code });

/** What `formula`, in `host` of Sheet1 beside `cells`, stores. */
function value(formula: string, cells: Cells = {}, host = "Z100", context: Partial<Context> = {}) {
  const model = calcWorkbook({ Sheet1: { ...cells, [host]: { formula } } });
  const at = parseA1(host)!;
  return evaluateCell(model, "sheet-Sheet1", at.row, at.col, { ...storedValueContext(model), ...context });
}

/** The result array of the array (CSE) formula `formula` anchored over
 * `range` of Sheet1 beside `cells`. */
function arrayResult(formula: string, cells: Cells, range: string) {
  const anchor = range.split(":")[0]!;
  const model = calcWorkbook({ Sheet1: { ...cells, [anchor]: { formula, array: range } } });
  const at = parseA1(anchor)!;
  return evaluateCell(model, "sheet-Sheet1", at.row, at.col, storedValueContext(model));
}

/** A formula's value and how often it drew from ctx.random: every RAND()
 * that is evaluated draws once, so an untaken RAND() branch draws nothing. */
function withDraws(formula: string) {
  let draws = 0;
  const result = value(formula, {}, "Z100", {
    random: () => {
      draws += 1;
      return 0.25;
    },
  });
  return { result, draws };
}

describe("IF", () => {
  it.each<[formula: string, expected: CellValue]>([
    ["IF(TRUE,1,1/0)", 1], // conformance row (Excel)
    ["IF(FALSE,1/0,2)", 2],
    ["IF(FALSE,1)", false], // docs: value_if_false omitted → FALSE
    ["IF(FALSE,1,)", 0], // docs: value_if_false an empty slot → 0
    ["IF(TRUE,)", 0], // docs: value_if_true an empty slot → 0
    ["IF(TRUE,,)", 0],
    ['IF(2,"yes","no")', "yes"], // a non-zero number is TRUE
    ['IF(0,"yes","no")', "no"],
    ['IF("true","yes","no")', "yes"], // Coercion: "TRUE"/"FALSE" in any case
    ['IF("x","yes","no")', err("#VALUE!")],
    ["IF(1/0,1,2)", err("#DIV/0!")],
    ["IF(NA(),1,2)", err("#N/A")],
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });

  it("reads an empty condition cell as FALSE", () => {
    expect(value('IF(A1,"yes","no")')).toBe("no");
  });

  it("evaluates only the branch it takes", () => {
    expect(withDraws("IF(TRUE,1,RAND())")).toEqual({ result: 1, draws: 0 });
    expect(withDraws("IF(FALSE,RAND(),2)")).toEqual({ result: 2, draws: 0 });
    expect(withDraws("IF(FALSE,1,RAND())")).toEqual({ result: 0.25, draws: 1 });
  });

  it("returns a branch's reference as a reference", () => {
    const cells: Cells = { A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30 };
    expect(value("SUM(IF(TRUE,A1:A3,B1:B3))", cells)).toBe(6);
    expect(value("SUM(IF(FALSE,A1:A3,B1:B3))", cells)).toBe(60);
    // A reference result is reduced like any top-level range: A2 for row 2.
    expect(value("IF(TRUE,A1:A3)", cells, "D2")).toBe(2);
  });
});

describe("IFS", () => {
  it.each<[score: number, grade: string]>([
    [93, "A"],
    [89, "B"],
    [70, "C"],
    [61, "D"],
    [59, "F"],
  ])("docs: the grade of %d is %s", (score, grade) => {
    expect(value('IFS(A1>89,"A",A1>79,"B",A1>69,"C",A1>59,"D",TRUE,"F")', { A1: score })).toBe(grade);
  });

  it.each<[formula: string, expected: CellValue]>([
    ["IFS(FALSE,1,FALSE,2)", err("#N/A")], // docs: no TRUE condition
    ['IFS("x",1)', err("#VALUE!")], // docs: a condition neither TRUE nor FALSE
    ["IFS(TRUE,1,1/0,2)", 1], // pinned: a later condition is never evaluated
    ['IFS(FALSE,1/0,TRUE,"b")', "b"], // pinned: an untaken value is never evaluated
    ["IFS(1/0,1,TRUE,2)", err("#DIV/0!")], // pinned: a condition's error is the result
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });

  it("evaluates only the value it returns", () => {
    expect(withDraws("IFS(TRUE,1,TRUE,RAND())")).toEqual({ result: 1, draws: 0 });
    expect(withDraws("IFS(FALSE,RAND(),TRUE,2)")).toEqual({ result: 2, draws: 0 });
  });
});

describe("IFERROR and IFNA", () => {
  it("docs: IFERROR(A2/B2,…) over quota and units sold", () => {
    const cells: Cells = { A2: 210, B2: 35, A3: 55, B3: 0, B4: 23 };
    expect(value('IFERROR(A2/B2,"Error in calculation")', cells)).toBe(6);
    expect(value('IFERROR(A3/B3,"Error in calculation")', cells)).toBe("Error in calculation");
    expect(value('IFERROR(A4/B4,"Error in calculation")', cells)).toBe(0); // A4 is empty
  });

  it.each<[formula: string, expected: CellValue]>([
    ["IFERROR(1,1/0)", 1],
    ['IFERROR(1/0,"x")', "x"],
    ['IFERROR(NA(),"x")', "x"],
    ['IFERROR(#REF!,"x")', "x"],
    ['IFNA(NA(),"x")', "x"],
    ['IFNA(1/0,"x")', err("#DIV/0!")], // only #N/A is caught
    ["IFNA(1,1/0)", 1],
    ['IFNA(VLOOKUP("Seattle",{"Atlanta",1;"Boston",2},2,FALSE),"Not found")', "Not found"],
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });

  it("evaluates the fallback only when it is returned", () => {
    expect(withDraws("IFERROR(1,RAND())")).toEqual({ result: 1, draws: 0 });
    expect(withDraws("IFERROR(1/0,RAND())")).toEqual({ result: 0.25, draws: 1 });
    expect(withDraws("IFNA(1/0,RAND())")).toEqual({ result: err("#DIV/0!"), draws: 0 });
  });
});

describe("AND, OR, XOR and NOT", () => {
  it("docs: the AND, OR, XOR and NOT examples", () => {
    const cells: Cells = { A2: 50, A3: 25 };
    expect(value("AND(A2>1,A2<100)", cells)).toBe(true);
    expect(value('IF(AND(A2<A3,A2<100),A2,"The value is out of range")', cells)).toBe("The value is out of range");
    expect(value('IF(AND(A3>1,A3<100),A3,"The value is out of range")', cells)).toBe(25);
    expect(value("OR(A2>1,A2<100)", cells)).toBe(true);
    expect(value('IF(OR(A2>1,A2<100),A3,"The value is out of range")', cells)).toBe(25);
    expect(value('IF(OR(A2<0,A2>50),A2,"The value is out of range")', cells)).toBe("The value is out of range");
    expect(value("XOR(3>0,2<9)")).toBe(false);
    expect(value("XOR(3>12,4>6)")).toBe(false);
    expect(value("NOT(A2>100)", cells)).toBe(true);
  });

  it.each<[formula: string, cells: Cells, expected: CellValue]>([
    ["AND(A1:A2)", { A1: "x", A2: true }, true], // conformance row (pinned)
    ["AND(A1:A3)", { A1: true, A2: 1, A3: 0 }, false],
    ["AND(A1:A3)", { A1: true, A2: 2 }, true], // A3 empty: skipped
    ["OR(A1:A3)", { A1: "TRUE", A3: false }, false], // text skipped, even "TRUE"
    ["OR(A1:A2)", { A1: 0, A2: -1 }, true],
    ["XOR(A1:A3)", { A1: true, A2: true, A3: true }, true], // three TRUE: odd
    ["XOR(A1:A2)", { A1: true, A2: 1 }, false], // two TRUE: even
    ['AND({TRUE,"x",1})', {}, true],
    ['OR({FALSE,"TRUE"})', {}, false],
  ])("=%s over %j reads logicals and numbers only: %j", (formula, cells, expected) => {
    expect(value(formula, cells)).toEqual(expected);
  });

  it.each<[formula: string, cells: Cells]>([
    ["AND(A1:A2)", { A1: "x" }],
    ["OR(A1)", {}],
    ["XOR(A1:A2)", { A1: "a", A2: "b" }],
    ['AND({"x"})', {}],
  ])("=%s over %j has nothing to read: #VALUE!", (formula, cells) => {
    // docs: "If the specified range contains no logical values, AND returns
    // the #VALUE! error value" (OR and XOR alike).
    expect(value(formula, cells)).toEqual(err("#VALUE!"));
  });

  it.each<[formula: string, expected: CellValue]>([
    ['AND("TRUE",1)', true],
    ['OR("false",0)', false],
    ['XOR("true",FALSE)', true],
    ["AND(2,-1)", true],
    ['AND("x")', err("#VALUE!")], // conformance row (pinned)
    ['OR(TRUE,"x")', err("#VALUE!")],
    ["NOT(0)", true],
    ["NOT(2)", false],
    ['NOT("false")', true],
    ['NOT("x")', err("#VALUE!")],
  ])("coerces a direct argument logically: =%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });

  it.each<[formula: string, cells: Cells, expected: CellError]>([
    ["AND(TRUE,1/0)", {}, err("#DIV/0!")],
    ["AND(FALSE,NA())", {}, err("#N/A")], // every argument is read: no short circuit
    ["OR(A1:A2)", { A1: true, A2: err("#N/A") }, err("#N/A")],
    ["XOR(A1:A2)", { A1: err("#REF!"), A2: true }, err("#REF!")],
    ["NOT(1/0)", {}, err("#DIV/0!")],
  ])("=%s over %j passes on the error", (formula, cells, expected) => {
    expect(value(formula, cells)).toEqual(expected);
  });
});

describe("in an array formula", () => {
  const cells: Cells = { A1: 1, A2: 2, A3: 3 };

  it("IF, IFERROR and CHOOSE apply element-wise", () => {
    expect(arrayResult('IF(A1:A3>1,"big","small")', cells, "B1:B3")).toEqual({ kind: "array", rows: 3, cols: 1, data: ["small", "big", "big"] });
    expect(arrayResult('IFERROR(1/A1:A3,"x")', { A1: 1, A2: 0, A3: 2 }, "B1:B3")).toEqual({ kind: "array", rows: 3, cols: 1, data: [1, "x", 0.5] });
    expect(arrayResult('CHOOSE({1,2},"a","b")', {}, "D1:E1")).toEqual({ kind: "array", rows: 1, cols: 2, data: ["a", "b"] });
  });

  it("SUM(IF(…)) — the classic CSE filter — and AND over a computed array", () => {
    expect(arrayResult("SUM(IF(A1:A3>1,A1:A3))", cells, "C1:C1")).toEqual({ kind: "array", rows: 1, cols: 1, data: [5] });
    expect(arrayResult("AND(A1:A3>0)", cells, "C1:C1")).toEqual({ kind: "array", rows: 1, cols: 1, data: [true] });
    expect(arrayResult("AND(A1:A3>1)", cells, "C1:C1")).toEqual({ kind: "array", rows: 1, cols: 1, data: [false] });
  });
});

describe("TRUE and FALSE", () => {
  it("return their logical value", () => {
    expect(value("TRUE()")).toBe(true);
    expect(value("FALSE()")).toBe(false);
    expect(value("TRUE()+TRUE()")).toBe(2);
  });
});

describe("SWITCH", () => {
  it.each<[formula: string, expected: CellValue]>([
    ['SWITCH(2,1,"Sunday",2,"Monday",3,"Tuesday")', "Monday"],
    ['SWITCH(99,1,"Sunday",2,"Monday",3,"Tuesday")', err("#N/A")], // no match, no default
    ['SWITCH(99,1,"Sunday",2,"Monday",3,"Tuesday","No match")', "No match"],
    ['SWITCH(2,1,"Sunday",7,"Saturday","weekday")', "weekday"],
    ['SWITCH(3,1,"Sunday",2,"Monday",3,"Tuesday","No match")', "Tuesday"],
    // 2018-02-22 was a Thursday: WEEKDAY is 5, which no value matches.
    ['SWITCH(WEEKDAY(DATE(2018,2,22)),1,"Sunday",2,"Monday",3,"Tuesday","No match")', "No match"],
    ['SWITCH(1/0,1,"a","b")', err("#DIV/0!")],
  ])("docs-style: =%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });

  it("evaluates only the result it returns and stops comparing at the match", () => {
    expect(value('SWITCH(1,1,"one",2,1/0)')).toBe("one");
    expect(value('SWITCH(2,1,1/0,2,"two")')).toBe("two");
    expect(value('SWITCH(1,1,"a",1/0,"b")')).toBe("a"); // pinned
    expect(withDraws('SWITCH(1,1,"one",RAND())')).toEqual({ result: "one", draws: 0 });
    expect(withDraws('SWITCH(5,1,"one",RAND())')).toEqual({ result: 0.25, draws: 1 });
  });
});

describe("CHOOSE", () => {
  it("docs: picks the index_num-th value", () => {
    const cells: Cells = { A2: "1st", A3: "2nd", A4: "3rd", A5: "Finished", B2: "Nails", B3: "Screws", B4: "Nuts", B5: "Bolts" };
    expect(value("CHOOSE(2,A2,A3,A4,A5)", cells)).toBe("2nd");
    expect(value("CHOOSE(4,B2,B3,B4,B5)", cells)).toBe("Bolts");
    expect(value('CHOOSE(3,"Wide",115,"world",8)')).toBe("world");
  });

  it("docs: returns a reference, so SUM(A2:CHOOSE(2,A3,A4,A5)) sums A2:A4", () => {
    expect(value("SUM(A2:CHOOSE(2,A3,A4,A5))", { A2: 23, A3: 45, A4: 12, A5: 10 })).toBe(80);
    expect(value("SUM(CHOOSE(2,A1:A3,B1:B3))", { A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30 })).toBe(60);
  });

  it.each<[formula: string, expected: CellValue]>([
    ['CHOOSE(2.9,"a","b","c")', "b"], // docs: a fraction is truncated
    ['CHOOSE(1.5,"a","b")', "a"],
    ['CHOOSE(0,"a","b")', err("#VALUE!")], // docs: less than 1
    ['CHOOSE(3,"a","b")', err("#VALUE!")], // docs: past the last value
    ['CHOOSE(0.9,"a")', err("#VALUE!")], // truncates to 0
    ['CHOOSE(-1,"a")', err("#VALUE!")],
    ['CHOOSE(1/0,"a")', err("#DIV/0!")],
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });

  it("evaluates only the chosen value", () => {
    expect(value('CHOOSE(1,"a",1/0)')).toBe("a");
    expect(value('CHOOSE(2,1/0,"b")')).toBe("b");
    expect(withDraws('CHOOSE(1,"a",RAND())')).toEqual({ result: "a", draws: 0 });
    expect(withDraws('CHOOSE(2,"a",RAND())')).toEqual({ result: 0.25, draws: 1 });
  });
});
