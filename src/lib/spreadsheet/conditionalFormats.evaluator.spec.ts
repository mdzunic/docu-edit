import { afterEach, describe, expect, it } from "vitest";
import { setClock } from "../core/clock";
import { DEFAULT_THEME } from "../core/theme";
import { parseRange, type GridRange } from "./address";
import { createCfEvaluator } from "./cfEvaluator";
import { addConditionalFormat, type CfRule, type ConditionalFormat, type Dxf } from "./conditionalFormats";
import { cellKey } from "./formula/dependencies";
import { storedValueContext } from "./formula/evaluate";
import { clearFormulaAstCache, formulaAstCacheSize } from "./formula/parse";
import type { Row, SpreadsheetModel } from "./model";
import { dateToSerial } from "./serial";

const RED: Dxf = { fill: { kind: "solid", color: "#FFC7CE" } };

function range(text: string): GridRange {
  return parseRange(text)!;
}

/** A one-sheet workbook whose column A holds `values` from row 0. */
function book(values: number[], formats: ConditionalFormat[], extra: Row[] = []): SpreadsheetModel {
  const rows: Row[] = values.map((value, index) => ({ index, cells: [{ col: 0, value }] }));
  rows.push(...extra);
  rows.sort((a, b) => a.index - b.index);
  return {
    kind: "spreadsheet",
    id: "book",
    title: "Book",
    format: "xlsx",
    readerVersion: 10,
    sheets: [{ id: "s1", name: "Sheet1", visibility: "visible", kind: "grid", rows, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], validations: [], conditionalFormats: formats }],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{}],
    defaultStyleId: 0,
    dxfs: [RED],
  };
}

const SCALE: CfRule = {
  id: "scale",
  priority: 1,
  kind: "colorScale",
  stops: [
    { at: { type: "min" }, color: "#000000" },
    { at: { type: "max" }, color: "#ffffff" },
  ],
};

/** An evaluator counting statistics computations per format id. */
function counted(model: SpreadsheetModel) {
  const statistics: string[] = [];
  const base = storedValueContext(model);
  const reads: string[] = [];
  const evaluator = createCfEvaluator(model, "s1", {
    base: { ...base, valueAt: (sheetId, row, col) => (reads.push(`${row}:${col}`), base.valueAt(sheetId, row, col)) },
    theme: DEFAULT_THEME,
    onStatistics: (id) => statistics.push(id),
  });
  return { evaluator, statistics, reads };
}

afterEach(() => setClock(null));

describe("createCfEvaluator", () => {
  it("computes a 100,000-cell colour scale's statistics exactly once over 1,000 resultFor calls", () => {
    const values = Array.from({ length: 100_000 }, (_, i) => i);
    const model = book(values, [{ id: "big", ranges: [range("A1:A100000")], rules: [SCALE] }]);
    const { evaluator, statistics } = counted(model);
    for (let row = 0; row < 1000; row += 1) {
      expect(evaluator.resultFor(row, 0)?.scaleColor).toMatch(/^#[0-9a-f]{6}$/);
    }
    expect(statistics).toEqual(["big"]);
    expect(evaluator.resultFor(99_999, 0)?.scaleColor).toBe("#ffffff");
    expect(statistics).toEqual(["big"]);
  });

  it("evaluates only the cells the grid asks for", () => {
    const model = book([1, 200, 3], [{ id: "f", ranges: [range("A1:A1000")], rules: [{ id: "r", priority: 1, kind: "cellIs", operator: "greaterThan", formulas: ["100"], dxfId: 0 }] }]);
    const { evaluator, reads } = counted(model);
    expect(evaluator.resultFor(1, 0)).toEqual({ dxf: RED });
    expect(evaluator.resultFor(1, 0)).toEqual({ dxf: RED });
    expect(reads).toEqual(["1:0"]);
    expect(evaluator.resultFor(5, 3)).toBeUndefined(); // outside every format: nothing read
    expect(reads).toEqual(["1:0"]);
  });

  it("keeps statistics until a changed cell lies in the format's ranges, or everything is invalidated", () => {
    const model = book([1, 2, 3], [{ id: "f", ranges: [range("A1:A3")], rules: [SCALE] }]);
    const { evaluator, statistics } = counted(model);
    expect(evaluator.resultFor(0, 0)?.scaleColor).toBe("#000000");
    evaluator.invalidate([cellKey("s1", 10, 0), cellKey("s2", 0, 0)]);
    expect(evaluator.resultFor(0, 0)?.scaleColor).toBe("#000000");
    expect(statistics).toEqual(["f"]);
    model.sheets = [{ ...model.sheets[0]!, rows: [{ index: 0, cells: [{ col: 0, value: 5 }] }, ...model.sheets[0]!.rows.slice(1)] }];
    evaluator.invalidate([cellKey("s1", 0, 0)]);
    expect(evaluator.resultFor(0, 0)?.scaleColor).toBe("#ffffff");
    expect(statistics).toEqual(["f", "f"]);
    evaluator.invalidate("all");
    evaluator.resultFor(1, 0);
    expect(statistics).toEqual(["f", "f", "f"]);
  });

  it("clears its per-cell cache on every invalidate, re-evaluating formula thresholds without recomputing statistics", () => {
    const model = book([], [{ id: "f", ranges: [range("A1:A2")], rules: [{ ...SCALE, stops: [{ at: { type: "num", value: "0" }, color: "#000000" }, { at: { type: "formula", value: "$C$1" }, color: "#ffffff" }] }] }], [
      { index: 0, cells: [{ col: 0, value: 10 }, { col: 2, value: 20 }] },
      { index: 1, cells: [{ col: 0, value: 20 }] },
    ]);
    const { evaluator, statistics } = counted(model);
    expect(evaluator.resultFor(0, 0)?.scaleColor).toBe("#808080");
    model.sheets = [{ ...model.sheets[0]!, rows: [{ index: 0, cells: [{ col: 0, value: 10 }, { col: 2, value: 10 }] }, model.sheets[0]!.rows[1]!] }];
    expect(evaluator.resultFor(0, 0)?.scaleColor).toBe("#808080"); // cached until invalidated
    evaluator.invalidate([cellKey("s1", 0, 2)]);
    expect(evaluator.resultFor(0, 0)?.scaleColor).toBe("#ffffff");
    expect(statistics).toEqual(["f"]);
  });

  it("follows a command's replaced formats without an explicit invalidate", () => {
    const model = book([150], []);
    const { evaluator } = counted(model);
    expect(evaluator.resultFor(0, 0)).toBeUndefined();
    addConditionalFormat(model, "s1", [range("A1")], { kind: "cellIs", operator: "greaterThan", formulas: ["100"], dxfId: 0 });
    expect(evaluator.resultFor(0, 0)).toEqual({ dxf: RED });
  });

  it("parses each rule formula once per origin, not once per cell", () => {
    clearFormulaAstCache();
    const model = book(
      Array.from({ length: 200 }, (_, i) => i),
      [{ id: "f", ranges: [range("A1:A200")], rules: [{ id: "r", priority: 1, kind: "expression", formula: "MOD(A1,2)=0", dxfId: 0 }] }],
    );
    const { evaluator } = counted(model);
    const formatted = Array.from({ length: 200 }, (_, row) => evaluator.resultFor(row, 0) !== undefined);
    expect(formatted.filter(Boolean)).toHaveLength(100);
    expect(formulaAstCacheSize()).toBe(1);
  });

  it("re-evaluates timePeriod rules and volatile formulas on every pass (setClock)", () => {
    setClock(new Date(2026, 9, 14, 9, 0));
    const today = dateToSerial(2026, 10, 14, "1900");
    const model = book([today, today + 1], [
      { id: "t", ranges: [range("A1:A2")], rules: [{ id: "r", priority: 1, kind: "timePeriod", period: "today", dxfId: 0 }] },
      { id: "v", ranges: [range("B1:B2")], rules: [{ id: "w", priority: 2, kind: "expression", formula: "$A1=TODAY()", dxfId: 0 }] },
    ]);
    const { evaluator } = counted(model);
    expect([evaluator.resultFor(0, 0), evaluator.resultFor(1, 0)].map(Boolean)).toEqual([true, false]);
    expect([evaluator.resultFor(0, 1), evaluator.resultFor(1, 1)].map(Boolean)).toEqual([true, false]);
    setClock(new Date(2026, 9, 15, 9, 0));
    evaluator.invalidate([]); // a recalculation pass that changed nothing
    expect([evaluator.resultFor(0, 0), evaluator.resultFor(1, 0)].map(Boolean)).toEqual([false, true]);
    expect([evaluator.resultFor(0, 1), evaluator.resultFor(1, 1)].map(Boolean)).toEqual([false, true]);
  });
});
