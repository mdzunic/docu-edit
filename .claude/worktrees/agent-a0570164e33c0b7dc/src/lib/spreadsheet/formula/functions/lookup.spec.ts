// SPEC-120 "The core functions" (lookup.ts): VLOOKUP HLOOKUP INDEX MATCH ROW
// COLUMN ROWS COLUMNS OFFSET INDIRECT SINGLE. Exact mode (MATCH 0,
// VLOOKUP/HLOOKUP FALSE or an empty slot) is the first match — wildcards for
// text, case-insensitive, empty cells never matching; approximate mode on
// sorted data is the largest value ≤ (MATCH -1: the smallest ≥ on descending
// data); col_index < 1 is #VALUE!, past the table #REF!, not found #N/A.
// INDEX and OFFSET return references (SUM(INDEX(A1:C3,0,2)) sums a column;
// out of bounds, off the grid or a size ≤ 0 is #REF!); MATCH on a 2-D range
// is #N/A; ROW()/COLUMN() are the host's, and of a multi-cell range in an
// array formula the array of indices; INDIRECT reads one A1 reference,
// sheet-qualified or quoted, and anything else — a name, garbage, a FALSE
// a1 argument — is #REF!; OFFSET and INDIRECT record the areas they return
// through ctx.recordDependency; SINGLE is implicit intersection in any mode.
//
// Sources ("Where expected values come from"): "docs" rows are Microsoft's
// function-reference examples (the density, Axles/Bearings/Bolts and fruit
// tables), recomputed by hand; "pinned" rows follow from SPEC-120's rules.
import { describe, expect, it } from "vitest";
import { parseA1 } from "../../address";
import { calcWorkbook, type CellSpec } from "../../fixtures/calcWorkbook";
import type { CellError, CellValue, ErrorCode } from "../../model";
import { evaluateCell, storedValueContext, type EvalContext } from "../evaluate";
import type { Area } from "../values";

type Cells = Record<string, CellSpec>;
type Context = Omit<EvalContext, "host" | "origin" | "arrayMode">;
type Row = [formula: string, expected: CellValue];

const err = (code: ErrorCode): CellError => ({ error: code });

interface Where {
  host?: string;
  /** The host is an array formula's anchor over this range. */
  array?: string;
  /** Sheets after Sheet1. */
  sheets?: Record<string, Cells>;
  context?: Partial<Context>;
}

/** What `formula`, in `host` (default Z100) of Sheet1 beside `cells`,
 * evaluates to (an ArrayValue for an array formula's anchor). */
function value(formula: string, cells: Cells = {}, where: Where = {}) {
  const host = where.host ?? "Z100";
  const self: CellSpec = where.array === undefined ? { formula } : { formula, array: where.array };
  const model = calcWorkbook({ Sheet1: { ...cells, [host]: self }, ...where.sheets });
  const at = parseA1(host)!;
  return evaluateCell(model, "sheet-Sheet1", at.row, at.col, { ...storedValueContext(model), ...where.context });
}

/** `formula`'s value and the areas it recorded as dynamic dependencies. */
function recording(formula: string, cells: Cells = {}, sheets: Record<string, Cells> = {}) {
  const recorded: Area[] = [];
  const result = value(formula, cells, { sheets, context: { recordDependency: (area) => void recorded.push(area) } });
  return { result, recorded };
}

/** Rows of values from `topLeft` down, as cells ("A2", [[1, 2], [3, 4]]). */
function grid(topLeft: string, rows: CellValue[][]): Cells {
  const at = parseA1(topLeft)!;
  const cells: Cells = {};
  rows.forEach((row, r) =>
    row.forEach((cell, c) => {
      if (cell !== null) {
        cells[`${String.fromCharCode(65 + at.col + c)}${at.row + r + 1}`] = cell;
      }
    }),
  );
  return cells;
}

describe("VLOOKUP", () => {
  const density = grid("A1", [
    ["Density", "Viscosity", "Temperature"],
    [0.457, 3.55, 500],
    [0.525, 3.25, 400],
    [0.606, 2.93, 300],
    [0.675, 2.75, 250],
    [0.746, 2.57, 200],
    [0.835, 2.38, 150],
    [0.946, 2.17, 100],
    [1.09, 1.95, 50],
    [1.29, 1.71, 0],
  ]);

  it.each<Row>([
    ["VLOOKUP(1,A2:C10,2)", 2.17], // docs: largest ≤ 1 is 0.946
    ["VLOOKUP(1,A2:C10,3,TRUE)", 100], // docs
    ["VLOOKUP(0.7,A2:C10,3,FALSE)", err("#N/A")], // docs: no exact match
    ["VLOOKUP(0.1,A2:C10,2,TRUE)", err("#N/A")], // docs: below the smallest
    ["VLOOKUP(2,A2:C10,2,TRUE)", 1.71], // docs: largest ≤ 2 is 1.29
    ["VLOOKUP(0.946,A2:C10,3,FALSE)", 100],
    ["VLOOKUP(1,A2:C10,0)", err("#VALUE!")], // docs: col_index_num < 1
    ["VLOOKUP(1,A2:C10,4)", err("#REF!")], // docs: past the table
  ])("docs table: =%s is %j", (formula, expected) => {
    expect(value(formula, density)).toEqual(expected);
  });

  it.each<Row>([
    ['VLOOKUP("a",{"c",1;"a",2},2,)', 2], // conformance row: an empty slot is exact
    ['VLOOKUP("A",{"c",1;"a",2},2,FALSE)', 2], // case-insensitive
    ['VLOOKUP("b*",{"apple",1;"banana",2},2,FALSE)', 2], // wildcards
    ['VLOOKUP("?pple",{"apple",1;"banana",2},2,FALSE)', 1],
    ['VLOOKUP("a",{"a",1;"a",2},2,FALSE)', 1], // the first match
    ['VLOOKUP("1",{1,"one";2,"two"},2,FALSE)', err("#N/A")], // text is no number
    ['VLOOKUP("z",{"a",1;"b",2},2,FALSE)', err("#N/A")],
    ['VLOOKUP("c",{"a",1;"b",2;"d",4},2)', 2], // approximate: largest ≤ "c"
    ['VLOOKUP(2.5,{1,"one";2,"two";3,"three"},2)', "two"],
    ['VLOOKUP(5,{1,"one";2,"two";3,"three"},2,TRUE)', "three"],
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });

  it("never matches an empty cell, even looking up 0", () => {
    expect(value("VLOOKUP(0,A1:B3,2,FALSE)", { B1: "blank", A2: 0, B2: "zero", A3: 1, B3: "one" })).toBe("zero");
  });
});

describe("HLOOKUP", () => {
  const parts = grid("A1", [["Axles", "Bearings", "Bolts"], [4, 4, 9], [5, 7, 10], [6, 8, 11]]);

  it.each<Row>([
    ['HLOOKUP("Axles",A1:C4,2,TRUE)', 4], // docs
    ['HLOOKUP("Bearings",A1:C4,3,FALSE)', 7], // docs
    ['HLOOKUP("B",A1:C4,3,TRUE)', 5], // docs: largest ≤ "B" is "Axles"
    ['HLOOKUP("Bolts",A1:C4,4)', 11], // docs
    ['HLOOKUP(3,{1,2,3;"a","b","c";"d","e","f"},2,TRUE)', "c"], // docs
    ['HLOOKUP("bolts",A1:C4,2,FALSE)', 9], // case-insensitive
    ['HLOOKUP("Axles",A1:C4,0)', err("#VALUE!")], // docs: row_index_num < 1
    ['HLOOKUP("Axles",A1:C4,5)', err("#REF!")], // docs: past the table
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula, parts)).toEqual(expected);
  });
});

describe("MATCH", () => {
  const counts = grid("A2", [["Bananas", 25], ["Oranges", 38], ["Apples", 40], ["Pears", 41]]);

  it.each<Row>([
    ["MATCH(39,B2:B5,1)", 2], // docs: the next lowest, 38
    ["MATCH(41,B2:B5,0)", 4], // docs
    ["MATCH(40,B2:B5,-1)", err("#N/A")], // docs: B2:B5 is not descending
    ["MATCH(39,B2:B5)", 2], // match_type defaults to 1
    ["MATCH(10,B2:B5,1)", err("#N/A")], // below the smallest
    ['MATCH("apples",A2:A5,0)', 3], // case-insensitive
  ])("docs table: =%s is %j", (formula, expected) => {
    expect(value(formula, counts)).toEqual(expected);
  });

  it.each<Row>([
    ['MATCH("b",{"a","b","c"},0)', 2], // docs
    ['MATCH("B*",{"apple","banana"},0)', 2], // conformance row (Excel)
    ['MATCH("?pple",{"banana","apple"},0)', 2],
    ['MATCH("~*",{"a","*"},0)', 2], // ~ escapes a literal *
    ["MATCH(42,{50,45,40,30},-1)", 2], // the smallest ≥ 42 on descending data
    ["MATCH(40,{50,45,40,30},-1)", 3],
    ["MATCH(60,{50,45,40,30},-1)", err("#N/A")],
    ['MATCH(1,{"1",1},0)', 2], // the text "1" is no match for 1
    ["MATCH(TRUE,{1,TRUE},0)", 2],
    ['MATCH("z",{"a","b"},0)', err("#N/A")],
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });

  it("searches one row or one column; a 2-D range is #N/A", () => {
    expect(value('MATCH("c",A1:C1,0)', { A1: "a", B1: "b", C1: "c" })).toBe(3);
    expect(value("MATCH(1,A1:B2,0)", { A1: 1, B1: 2, A2: 3, B2: 4 })).toEqual(err("#N/A"));
  });

  it("never matches an empty cell, even looking up 0", () => {
    expect(value("MATCH(0,A1:A3,0)", { A2: 0, A3: 0 })).toBe(2);
  });

  it("looks up each element of an array lookup value in an array formula (VLOOKUP too)", () => {
    expect(value('MATCH({"b","a"},{"a","b"},0)', {}, { host: "D1", array: "D1:E1" })).toEqual({ kind: "array", rows: 1, cols: 2, data: [2, 1] });
    expect(value('VLOOKUP({"a";"b"},{"a",1;"b",2},2,FALSE)', {}, { host: "C1", array: "C1:C2" })).toEqual({ kind: "array", rows: 2, cols: 1, data: [1, 2] });
  });
});

describe("INDEX", () => {
  // Powers of two, so every sum names the cells it took.
  const powers = grid("A1", [[1, 2, 4], [8, 16, 32], [64, 128, 256]]);

  it.each<Row>([
    ["INDEX(A1:C3,2,2)", 16], // B2
    ["SUM(INDEX(A1:C3,0,2))", 146], // column B: row 0 is the whole column
    ["SUM(INDEX(A1:C3,2,0))", 56], // row 2
    ["SUM(INDEX(A1:C3,2))", 56], // docs: only row_num on a 2-D source → the whole row
    ["SUM(INDEX(A1:C3,,2))", 146], // an empty row slot is 0
    ["SUM(INDEX(A1:C3,0,0))", 511], // docs: both 0 → the whole range
    ["INDEX(A1:C1,2)", 2], // a one-row source indexes columns
    ["INDEX(A1:A3,2)", 8],
    ["ROWS(INDEX(A1:C3,0,2))", 3],
    ["COLUMNS(INDEX(A1:C3,2,0))", 3],
    ["SUM(A1:INDEX(A1:C3,2,2))", 27], // a reference: A1:B2
    ["INDEX(A1:C3,4,1)", err("#REF!")],
    ["INDEX(A1:C3,1,4)", err("#REF!")],
    ["INDEX(A1:C1,4)", err("#REF!")],
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula, powers)).toEqual(expected);
  });

  it("a whole column at the top of a normal formula is intersected with the host's row", () => {
    expect(value("INDEX(A1:C3,0,2)", powers, { host: "E2" })).toBe(16);
  });

  it("docs: the array form", () => {
    const fruit = grid("A2", [["Apples", "Lemons"], ["Bananas", "Pears"]]);
    expect(value("INDEX(A2:B3,2,2)", fruit)).toBe("Pears");
    expect(value("INDEX(A2:B3,2,1)", fruit)).toBe("Bananas");
    expect(value("INDEX({1,2;3,4},0,2)", {}, { host: "E1", array: "E1:E2" })).toEqual({ kind: "array", rows: 2, cols: 1, data: [2, 4] });
    expect(value("INDEX({1,2;3,4},2,1)")).toBe(3);
    expect(value('INDEX({"a","b","c"},2)')).toBe("b");
  });

  it("docs: the reference form", () => {
    const stock = grid("A1", [
      ["Fruit", "Price", "Count"],
      ["Apples", 0.69, 40],
      ["Bananas", 0.34, 38],
      ["Lemons", 0.55, 15],
      ["Oranges", 0.25, 25],
      ["Pears", 0.59, 40],
      [null, null, null],
      ["Almonds", 2.8, 10],
      ["Cashews", 3.55, 16],
      ["Peanuts", 1.25, 20],
      ["Walnuts", 1.75, 12],
    ]);
    expect(value("INDEX(A2:C6,2,3)", stock)).toBe(38); // C3
    expect(value("INDEX((A1:C6,A8:C11),2,2,2)", stock)).toBe(3.55); // B9, in the second area
    expect(value("SUM(INDEX(A1:C11,0,3,1))", stock)).toBe(216); // C1:C11
    expect(value("SUM(B2:INDEX(A2:C6,5,2))", stock)).toBeCloseTo(2.42, 12); // B2:B6
  });
});

describe("ROW, COLUMN, ROWS and COLUMNS", () => {
  it.each<[formula: string, host: string, expected: CellValue]>([
    ["ROW()", "C7", 7], // docs: the formula's own row
    ["COLUMN()", "D7", 4],
    ["ROW(C10)", "Z100", 10], // docs
    ["COLUMN(D10)", "Z100", 4], // docs
    ["ROW(A5:A7)", "Z100", 5], // pinned: a normal formula takes the first row
    ["COLUMN(C:D)", "Z100", 3],
    ["ROWS(C1:E4)", "Z100", 4], // docs
    ["COLUMNS(C1:E4)", "Z100", 3], // docs
    ["ROWS({1,2,3;4,5,6})", "Z100", 2], // docs
    ["COLUMNS({1,2,3;4,5,6})", "Z100", 3], // docs
    ["ROWS(A:A)", "Z100", 1048576],
    ["COLUMNS(1:1)", "Z100", 16384],
  ])("=%s in %s is %j", (formula, host, expected) => {
    expect(value(formula, {}, { host })).toEqual(expected);
  });

  it("of a multi-cell range in an array formula return the array of indices", () => {
    expect(value("ROW(A1:A3)", {}, { host: "B1", array: "B1:B3" })).toEqual({ kind: "array", rows: 3, cols: 1, data: [1, 2, 3] });
    expect(value("COLUMN(B1:D1)", {}, { host: "A5", array: "A5:C5" })).toEqual({ kind: "array", rows: 1, cols: 3, data: [2, 3, 4] });
  });
});

describe("OFFSET", () => {
  const cells: Cells = { ...grid("B4", [[100, 1000], [200, 2000]]), ...grid("B6", [[1, 2, 3], [4, 5, 6], [7, 8, 9]]) };

  it.each<Row>([
    ["SUM(OFFSET(B3,1,0,2,1))", 300], // B4:B5
    ["OFFSET(D3,3,-2,1,1)", 1], // docs: B6
    ["SUM(OFFSET(D3:F5,3,-2,3,3))", 45], // docs: B6:D8
    ["SUM(OFFSET(B4:B5,0,1))", 3000], // the reference's own size: C4:C5
    ["OFFSET(B3,1,0)", 100],
    ["ROWS(OFFSET(A1,0,0,5,2))", 5],
    ["OFFSET(D3,-3,-3)", err("#REF!")], // docs: off the sheet
    ["OFFSET(A1,0,-1)", err("#REF!")],
    ["OFFSET(A1,1048576,0)", err("#REF!")],
    ["OFFSET(A1,0,0,1048577,1)", err("#REF!")],
    ["OFFSET(A1,0,0,0,1)", err("#REF!")], // pinned: height ≤ 0
    ["OFFSET(A1,0,0,1,-1)", err("#REF!")], // pinned: width ≤ 0
    ["OFFSET((A1,C1),0,0)", err("#VALUE!")], // docs: not one range of adjacent cells
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula, cells)).toEqual(expected);
  });

  it("records the area it returns as a dynamic dependency", () => {
    const b4b5: Area = { sheetId: "sheet-Sheet1", top: 3, left: 1, bottom: 4, right: 1 };
    expect(recording("SUM(OFFSET(B3,1,0,2,1))", cells)).toEqual({ result: 300, recorded: [b4b5] });
  });
});

describe("INDIRECT", () => {
  const cells: Cells = grid("A2", [["B2", 1.333], ["B3", 45], ["George", 10], [5, 62]]);
  const sheets: Record<string, Cells> = { Sheet2: { A1: 1, B1: 2, A2: 3, B2: 4 }, "My Sheet": { A1: "mine" } };

  it.each<Row>([
    ["INDIRECT(A2)", 1.333], // docs
    ["INDIRECT(A3)", 45], // docs
    ['INDIRECT("B"&A5)', 62], // docs
    ['INDIRECT("$B$2")', 1.333],
    ['INDIRECT("B2",TRUE)', 1.333],
    ['SUM(INDIRECT("Sheet2!A1:B2"))', 10],
    ['INDIRECT("sheet2!B2")', 4], // sheet names are case-insensitive
    ['INDIRECT("\'My Sheet\'!A1")', "mine"],
    ["INDIRECT(A4)", err("#REF!")], // "George" — no such name here (the docs define one)
    ['INDIRECT("garbage")', err("#REF!")],
    ['INDIRECT("A1+B1")', err("#REF!")],
    ['INDIRECT("Nope!A1")', err("#REF!")],
    ['INDIRECT("")', err("#REF!")],
    ['INDIRECT("B2",FALSE)', err("#REF!")], // pinned: R1C1 text is out of scope
  ])("=%s is %j", (formula, expected) => {
    expect(value(formula, cells, { sheets })).toEqual(expected);
  });

  it("records the area it returns as a dynamic dependency", () => {
    const sheet2: Area = { sheetId: "sheet-Sheet2", top: 0, left: 0, bottom: 1, right: 1 };
    const b2: Area = { sheetId: "sheet-Sheet1", top: 1, left: 1, bottom: 1, right: 1 };
    expect(recording('SUM(INDIRECT("Sheet2!A1:B2"))', cells, sheets)).toEqual({ result: 10, recorded: [sheet2] });
    expect(recording("INDIRECT(A2)", cells, sheets)).toEqual({ result: 1.333, recorded: [b2] });
  });
});

describe("SINGLE", () => {
  const cells: Cells = { A1: "r1", A2: "r2", A3: "r3", A5: "c1", B5: "c2", C5: "c3" };

  it.each<[formula: string, host: string, expected: CellValue]>([
    ["SINGLE(A1:A3)", "D2", "r2"], // one column: the host's row
    ["SINGLE(A1:A3)", "D9", err("#VALUE!")], // outside it
    ["SINGLE(A5:C5)", "B9", "c2"], // one row: the host's column
    ["SINGLE(A5:C5)", "E9", err("#VALUE!")],
    ["SINGLE(A1:B2)", "D2", err("#VALUE!")], // 2-D
    ["SINGLE(Sheet2!A1:A3)", "D2", "s2"], // the host's row on another sheet too
    ["SINGLE({1,2,3})", "D2", 1], // an array's top-left
    ["SINGLE(7)", "D2", 7],
  ])("=%s in %s is %j", (formula, host, expected) => {
    expect(value(formula, cells, { host, sheets: { Sheet2: { A1: "s1", A2: "s2", A3: "s3" } } })).toEqual(expected);
  });

  it("intersects in an array formula too", () => {
    expect(value('SINGLE(A1:A3)&"!"', cells, { host: "D2", array: "D2:D2" })).toEqual({ kind: "array", rows: 1, cols: 1, data: ["r2!"] });
    expect(value('A1:A3&"!"', cells, { host: "D2", array: "D2:D2" })).toEqual({ kind: "array", rows: 3, cols: 1, data: ["r1!", "r2!", "r3!"] });
  });
});
