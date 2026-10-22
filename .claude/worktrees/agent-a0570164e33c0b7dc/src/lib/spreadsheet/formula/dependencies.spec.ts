// SPEC-120 "The dependency graph": cellKey/parseCellKey; single-cell
// precedents in a map, ranges in BAND_WIDTH-column bands and — past
// WIDE_RANGE_BANDS bands — the sheet's wide list; 3-D references as one area
// per sheet; a reader of an array range depending on its anchor; the
// volatile, dynamic and uncomputable sets; formulaCells in sheet order then
// row-major; a 100,000-row shared formula sharing one AST with
// dependentsOf(A50000) a single key; the graph as lib state (a WeakMap by
// model, never in the model); updateGraphCells and invalidateDependencyGraph;
// findCycles in a deterministic order. (Dynamic dependencies replaced on
// re-evaluation: recalc.spec.ts.)
import { describe, expect, it } from "vitest";
import { parseA1 } from "../address";
import { calcWorkbook, type CellSpec } from "../fixtures/calcWorkbook";
import type { SpreadsheetModel } from "../model";
import {
  BAND_WIDTH,
  cellKey,
  findCycles,
  graphFor,
  invalidateDependencyGraph,
  parseCellKey,
  updateGraphCells,
  WIDE_RANGE_BANDS,
} from "./dependencies";
import { clearFormulaAstCache, formulaAstCacheSize } from "./parse";

function keyOf(a1: string, sheet = "Sheet1"): string {
  const at = parseA1(a1)!;
  return cellKey(`sheet-${sheet}`, at.row, at.col);
}

function dependents(model: SpreadsheetModel, a1: string, sheet = "Sheet1"): string[] {
  const at = parseA1(a1)!;
  return graphFor(model).dependentsOf(`sheet-${sheet}`, at.row, at.col).toSorted();
}

describe("keys", () => {
  it("are `${sheetId}!${row},${col}`, split at the last \"!\"", () => {
    expect(cellKey("s1", 4, 2)).toBe("s1!4,2");
    expect(parseCellKey(cellKey("a!b", 1048575, 16383))).toEqual({ sheetId: "a!b", row: 1048575, col: 16383 });
    expect([BAND_WIDTH, WIDE_RANGE_BANDS]).toEqual([32, 8]);
  });
});

describe("precedents", () => {
  it("finds single-cell, band and wide-list precedents — and nothing outside them", () => {
    const model = calcWorkbook({
      Sheet1: { Z1: "=A1*2", Z2: "=SUM(A1:B9)", Z3: "=SUM(A1:JZ1)", Z4: "=SUM(C:C)", Z5: "=SUM(AH5:AI5)" },
    });
    expect(dependents(model, "A1")).toEqual([keyOf("Z1"), keyOf("Z2"), keyOf("Z3")].toSorted());
    expect(dependents(model, "B9")).toEqual([keyOf("Z2")]);
    expect(dependents(model, "B10")).toEqual([]);
    expect(dependents(model, "JZ1")).toEqual([keyOf("Z3")]); // 286 columns: the wide list
    expect(dependents(model, "C1048576")).toEqual([keyOf("Z4")]);
    expect(dependents(model, "AI5")).toEqual([keyOf("Z5")]); // the second band
    expect(graphFor(model).precedentsOf(keyOf("Z2"))).toEqual([{ sheetId: "sheet-Sheet1", top: 0, left: 0, bottom: 8, right: 1 }]);
  });

  it("expands a 3-D reference to one area per sheet", () => {
    const model = calcWorkbook({ Jan: { B2: 1 }, Feb: { B2: 2 }, Sum: { A1: "=SUM(Jan:Feb!B2)" } });
    expect(graphFor(model).precedentsOf(keyOf("A1", "Sum"))).toEqual([
      { sheetId: "sheet-Jan", top: 1, left: 1, bottom: 1, right: 1 },
      { sheetId: "sheet-Feb", top: 1, left: 1, bottom: 1, right: 1 },
    ]);
    expect(dependents(model, "B2", "Feb")).toEqual([keyOf("A1", "Sum")]);
  });

  it("makes a reader of an array range depend on its anchor, which stands for the whole range", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, A2: 2, B1: { formula: "A1:A2*2", array: "B1:B2", value: 2 }, B2: 4, C1: "=B2+1" } });
    expect(dependents(model, "B1")).toEqual([keyOf("C1")]);
    expect(dependents(model, "A2")).toEqual([keyOf("B1")]);
    expect([...graphFor(model).formulaCells()]).toEqual([keyOf("B1"), keyOf("C1")]);
  });

  it("keeps the volatile, dynamic and uncomputable cells, and the formula cells in sheet order then row-major", () => {
    const model = calcWorkbook({
      Sheet2: { A2: "=NOW()" },
      Sheet1: { B1: "=OFFSET(A1,1,0)", A2: "=INDIRECT(\"A1\")", A1: "=NOPE()", C3: { formula: "SORT(A1:A2)", array: "C3:C4", dynamicArray: true } },
    });
    const graph = graphFor(model);
    expect([...graph.formulaCells()]).toEqual([keyOf("A2", "Sheet2"), keyOf("A1"), keyOf("B1"), keyOf("A2"), keyOf("C3")]);
    expect([...graph.volatile].toSorted()).toEqual([keyOf("A2", "Sheet2"), keyOf("B1"), keyOf("A2")].toSorted());
    expect([...graph.dynamic].toSorted()).toEqual([keyOf("B1"), keyOf("A2")].toSorted());
    expect(graph.uncomputable.get(keyOf("A1"))).toEqual({ kind: "function", name: "NOPE" });
    expect(graph.uncomputable.get(keyOf("C3"))).toEqual({ kind: "function", name: "SORT" });
  });
});

describe("scale", () => {
  it("a 100,000-row shared formula shares one AST, and dependentsOf(A50000) is one key", () => {
    clearFormulaAstCache();
    const cells: Record<string, CellSpec> = {};
    for (let row = 1; row <= 100_000; row += 1) {
      cells[`B${row}`] = { formula: `A${row}*2`, value: 0 };
    }
    const model = calcWorkbook({ Sheet1: cells });
    const graph = graphFor(model);
    expect(formulaAstCacheSize()).toBe(1);
    expect(graph.dependentsOf("sheet-Sheet1", 49_999, 0)).toEqual([keyOf("B50000")]);
  });
});

describe("the graph as lib state", () => {
  it("is cached per model in a WeakMap, never stored in it, and rebuilt after invalidateDependencyGraph", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, B1: "=A1" } });
    const graph = graphFor(model);
    expect(graphFor(model)).toBe(graph);
    expect(JSON.parse(JSON.stringify(model))).toEqual(model);
    invalidateDependencyGraph(model);
    expect(graphFor(model)).not.toBe(graph);
  });

  it("updateGraphCells adds, rebuilds and drops the nodes of edited cells", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, B1: "=A1", C1: 5 } });
    const graph = graphFor(model);
    const sheet = model.sheets[0]!;
    sheet.rows = [{ index: 0, cells: [{ col: 0, value: 1 }, { col: 1, value: 5 }, { col: 2, value: null, formula: "B1+1" }] }];
    updateGraphCells(model, [keyOf("B1"), keyOf("C1")]);
    expect(graphFor(model)).toBe(graph);
    expect([...graph.formulaCells()]).toEqual([keyOf("C1")]);
    expect(dependents(model, "A1")).toEqual([]);
    expect(dependents(model, "B1")).toEqual([keyOf("C1")]);
  });
});

describe("findCycles", () => {
  it("names every cycle and self-reference, members and cycles in sheet order then row-major", () => {
    const model = calcWorkbook({
      Sheet2: { C1: "=C1+1" },
      Sheet1: { B2: "=A1", A1: "=B2", D1: "=A1", E1: "=SUM(E1:E2)" },
    });
    expect(findCycles(graphFor(model))).toEqual([[keyOf("C1", "Sheet2")], [keyOf("A1"), keyOf("B2")], [keyOf("E1")]]);
    expect(findCycles(graphFor(calcWorkbook({ Sheet1: { A1: "=B1", B1: 1 } })))).toEqual([]);
  });
});
