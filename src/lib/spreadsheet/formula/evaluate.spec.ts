// SPEC-120 "Evaluation", "Ranges and areas", "Implicit intersection", "Array
// mode", "Errors": the interpreter over SPEC-119's FormulaNode — operators
// and their error order, implicit intersection (the host's row or column, on
// another sheet too), union, #NULL!, 3-D sums, unknown sheets and cross-sheet
// ranges as #REF!, array mode with broadcasting, the single-cell dynamic-array
// rule (computes; uncomputable once its result is larger than 1×1; a
// multi-cell one is never evaluated), lazy branches, MAX_ARRAY_CELLS, origin
// offsets on a shared AST, the final-operation rule at the root only,
// non-finite and overlong results, and the uncomputable reasons.
import { describe, expect, it } from "vitest";
import { parseA1 } from "../address";
import { calcWorkbook, type CellSpec } from "../fixtures/calcWorkbook";
import type { SpreadsheetModel } from "../model";
import { evalSettingsFor, evaluate, evaluateCell, MAX_EVAL_DEPTH, storedValueContext } from "./evaluate";
import { clearFormulaAstCache, formulaAstCacheSize, parseFormulaAt } from "./parse";
import { MAX_TEXT_LENGTH } from "./values";

function at(model: SpreadsheetModel, a1: string, sheet = "Sheet1", base = storedValueContext(model)) {
  const ref = parseA1(a1)!;
  return evaluateCell(model, `sheet-${sheet}`, ref.row, ref.col, base);
}

/** A formula's stored value at `host` of Sheet1, beside `cells`. */
function value(formula: string, cells: Record<string, CellSpec> = {}, host = "Z100") {
  return at(calcWorkbook({ Sheet1: { ...cells, [host]: { formula } } }), host);
}

describe("operators", () => {
  it.each([
    ["1+2*3", 7],
    ["(1+2)*3", 9],
    ["-2^2", 4],
    ["2^-1", 0.5],
    ["50%", 0.5],
    ["10%*10", 1],
    ['"a"&1&TRUE', "a1TRUE"],
    ['=+"a"', "a"],
    ["1/0", { error: "#DIV/0!" }],
    ["1E308*10", { error: "#NUM!" }],
    ["(-8)^(1/3)", { error: "#NUM!" }],
    ["0^-1", { error: "#DIV/0!" }],
    ['"a"+NA()', { error: "#VALUE!" }],
    ["NA()&1/0", { error: "#N/A" }],
  ] as const)("=%s is %j", (formula, expected) => {
    expect(value(formula.replace(/^=/, ""))).toEqual(expected);
  });

  it("text longer than MAX_TEXT_LENGTH is #VALUE!, never stored", () => {
    expect(value(`REPT("a",${MAX_TEXT_LENGTH})&"b"`)).toEqual({ error: "#VALUE!" });
    expect(value(`LEN(REPT("a",${MAX_TEXT_LENGTH}))`)).toBe(MAX_TEXT_LENGTH);
  });

  it("applies the final-operation rule at the root only (looking through parentheses)", () => {
    expect(value("0.1+0.2-0.3")).toBe(0);
    expect(value("(0.1+0.2-0.3)")).toBe(0);
    expect(value("0.3-(0.1+0.2)")).toBe(0);
    expect(value("1*(0.1+0.2-0.3)")).toBe(0.1 + 0.2 - 0.3);
    expect(value("-0.1+0.1")).toBe(0);
  });
});

describe("references, intersection and the reference operators", () => {
  const column = { A2: 20, A3: 30, A4: 40, B3: 7 };

  it("intersects a one-column range with the host's row and a one-row range with its column — on another sheet too", () => {
    expect(value("A2:A4*2", column, "C3")).toBe(60);
    expect(value("A2:A4", column, "C9")).toEqual({ error: "#VALUE!" });
    expect(value("A3:C3", column, "B10")).toBe(7);
    expect(value("A2:B4", column, "C3")).toEqual({ error: "#VALUE!" });
    const model = calcWorkbook({ Data: { A2: 1, A3: 2 }, Sheet1: { B3: "=Data!A2:A3" } });
    expect(at(model, "B3")).toBe(2);
  });

  it("unions, intersects (#NULL! when empty) and spans with `:` on one sheet (#REF! across sheets)", () => {
    const cells = { A1: 1, B1: 2, A2: 3, B2: 4 };
    expect(value("SUM((A1,B2))", cells)).toBe(5);
    expect(value("SUM(A1:B2 B1:B2)", cells)).toBe(6);
    expect(value("SUM(A1:A2 B1:B2)", cells)).toEqual({ error: "#NULL!" });
    expect(value("SUM(A1:INDEX(A1:B2,2,2))", cells)).toBe(10);
    const model = calcWorkbook({ Sheet1: { A1: 1, C1: "=SUM(Sheet1!A1:Sheet2!A1)" }, Sheet2: { A1: 2 } });
    expect(at(model, "C1")).toEqual({ error: "#REF!" });
  });

  it("sums a 3-D reference — one area per sheet between the names, either direction — and is #REF! for an unknown or foreign sheet", () => {
    const model = calcWorkbook(
      {
        Jan: { B2: 1 },
        Chart: {},
        Feb: { B2: 2 },
        Mar: { B2: 3 },
        Sheet1: { A1: "=SUM(Jan:Mar!B2)", A2: "=SUM(Mar:Feb!B2)", A3: "=Nowhere!A1", A4: "=Chart!A1" },
      },
      { foreign: ["Chart"] },
    );
    expect(at(model, "A1")).toBe(6);
    expect(at(model, "A2")).toBe(5);
    expect(at(model, "A3")).toEqual({ error: "#REF!" });
    expect(at(model, "A4")).toEqual({ error: "#REF!" });
  });

  it("resolves a shared AST's relative parts as part + (host − origin)", () => {
    clearFormulaAstCache();
    const model = calcWorkbook({ Sheet1: { A1: 1, A2: 2, A3: 3, B1: "=A1*10", B2: "=A2*10", B3: "=A3*10" } });
    expect(["B1", "B2", "B3"].map((a1) => at(model, a1))).toEqual([10, 20, 30]);
    expect(formulaAstCacheSize()).toBe(1);
    const shared = parseFormulaAt("A1*10", { row: 0, col: 1 });
    const ctx = { ...storedValueContext(model), host: { sheetId: "sheet-Sheet1", row: 2, col: 1 }, origin: { row: 0, col: 1 }, arrayMode: false };
    // The AST parsed for B1, evaluated at B3: A1 resolves to A3.
    expect("node" in shared ? evaluate(shared.node, ctx) : null).toBe(30);
  });

  it("reads an empty cell as EMPTY — 0 at the top level", () => {
    expect(value("A1")).toBe(0);
    expect(value('A1&"x"')).toBe("x");
  });
});

describe("array mode", () => {
  it("lifts operators and value arguments element-wise and broadcasts a row against a column", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, A2: 2, B1: { formula: "A1:A2*{10,100}", array: "B1:C3" }, D1: { formula: "ABS(-A1:A2)", array: "D1:D2" } } });
    expect(at(model, "B1")).toEqual({ kind: "array", rows: 2, cols: 2, data: [10, 100, 20, 200] });
    expect(at(model, "D1")).toEqual({ kind: "array", rows: 2, cols: 1, data: [1, 2] });
  });

  it("gives #N/A outside an operand when neither is a single row or column", () => {
    const model = calcWorkbook({ Sheet1: { A1: { formula: "{1,2;3,4}+{10;20;30}", array: "A1:B3" }, D1: { formula: "{1,2}+{1,2,3}", array: "D1:F1" } } });
    expect(at(model, "A1")).toEqual({ kind: "array", rows: 3, cols: 2, data: [11, 12, 23, 24, { error: "#N/A" }, { error: "#N/A" }] });
    expect(at(model, "D1")).toEqual({ kind: "array", rows: 1, cols: 3, data: [2, 4, { error: "#N/A" }] });
  });

  it("evaluates an \"array\" argument's subtree in array mode without CSE (SUMPRODUCT)", () => {
    expect(value("SUMPRODUCT((A1:A3>1)*B1:B3)", { A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30 })).toBe(50);
  });

  it("lifts a lazy function over an array condition, each element taking only its branch", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, A2: 5, A3: 9, B1: { formula: 'SUM(IF(A1:A3>2,A1:A3,0))', array: "B1" } } });
    expect(at(model, "B1")).toEqual({ kind: "array", rows: 1, cols: 1, data: [14] });
  });

  it("is #NUM! for a range materialising past MAX_ARRAY_CELLS (A:E is 5 × 1,048,576)", () => {
    expect(value("SUMPRODUCT(A:D)")).toBe(0);
    expect(value("SUMPRODUCT(A:E)")).toEqual({ error: "#NUM!" });
  });
});

/** A dynamic-array anchor over `array`, cached 99. */
const spill = (formula: string, array: string) => ({ formula, array, dynamicArray: true as const, value: 99 });

describe("dynamic arrays (SPEC-120 \"Array mode\")", () => {
  it("computes a single-cell dynamic-array anchor whose result fits its cell", () => {
    const model = calcWorkbook({ Sheet1: { A1: 2, A2: 3, B1: spill("A1+A2*10", "B1"), B2: spill("SUM(A1:A2)", "B2") } });
    expect(at(model, "B1")).toEqual({ kind: "array", rows: 1, cols: 1, data: [32] });
    expect(at(model, "B2")).toEqual({ kind: "array", rows: 1, cols: 1, data: [5] });
  });

  it("is uncomputable once its result is larger than 1×1, and a multi-cell anchor is never evaluated", () => {
    const model = calcWorkbook({ Sheet1: { A1: 2, A2: 3, B1: spill("A1:A2*2", "B1"), C1: spill("RAND()", "C1:C2") } });
    expect(at(model, "B1")).toEqual({ uncomputable: { kind: "dynamicArray" } });
    let draws = 0;
    const counting = { ...storedValueContext(model), random: () => (draws += 1) };
    expect(at(model, "C1", "Sheet1", counting)).toEqual({ uncomputable: { kind: "dynamicArray" } });
    expect(draws).toBe(0);
  });
});

describe("laziness and uncomputable formulas", () => {
  it("evaluates only the branch taken", () => {
    const model = calcWorkbook({ Sheet1: { A1: "=IF(TRUE,1,RAND())", A2: '=IFERROR(1,RAND())', A3: "=CHOOSE(1,5,RAND())" } });
    let draws = 0;
    const counting = { ...storedValueContext(model), random: () => (draws += 1) };
    expect(["A1", "A2", "A3"].map((a1) => at(model, a1, "Sheet1", counting))).toEqual([1, 1, 5]);
    expect(draws).toBe(0);
  });

  it.each([
    ["ZZZ.UNKNOWN(1)", { kind: "function", name: "ZZZ.UNKNOWN" }],
    ["Rate*2", { kind: "name", name: "Rate" }],
    ["SUM(Table1[Amount])", { kind: "structured" }],
    ["[1]Sheet1!A1+1", { kind: "external" }],
    ["IF(FALSE,NOPE(),1)", { kind: "function", name: "NOPE" }],
  ] as const)("=%s is uncomputable: %j", (formula, reason) => {
    expect(value(formula)).toEqual({ uncomputable: reason });
  });

  it("an unparsed formula, a data table and a non-formula cell", () => {
    const model = calcWorkbook({ Sheet1: { A1: { formula: "LAMBDA(x,x)(1)", formulaUnparsed: true, value: 1 }, A2: { dataTableXml: `<f t="dataTable" ref="A2"/>`, value: 3 }, A3: 7 } });
    expect(at(model, "A1")).toEqual({ uncomputable: { kind: "unparsed" } });
    expect(at(model, "A2")).toEqual({ uncomputable: { kind: "dataTable" } });
    expect(at(model, "A3")).toBe(7);
  });
});

describe("settings and bounds", () => {
  it("takes the workbook's settings: xlsx wildcards, whole cells, case-insensitive; ods per its calc and the proleptic calendar", () => {
    const xlsx = evalSettingsFor({ dateSystem: "1900", format: "xlsx" });
    expect(xlsx).toMatchObject({ caseSensitive: false, criteria: "wildcards", wholeCell: true, dateQuirks: "excel" });
    expect(xlsx.locale.tag).toBe("en-US");
    const ods = evalSettingsFor({ dateSystem: "1904", format: "ods", calc: { caseSensitive: true, criteria: "regex", partialCellMatch: true } });
    expect(ods).toMatchObject({ dateSystem: "1904", caseSensitive: true, criteria: "regex", wholeCell: false, dateQuirks: "proleptic" });
  });

  it("re-queues past MAX_EVAL_DEPTH = 512 nested evaluations (recalc.spec.ts drives the chain)", () => {
    expect(MAX_EVAL_DEPTH).toBe(512);
  });
});
