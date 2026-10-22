// SPEC-127 "SUBTOTAL and AGGREGATE" and the uncomputable rules: SUBTOTAL 1–11
// keep user-hidden rows (they skip only filtered ones, which SPEC-128 adds)
// and 101–111 skip them; both skip cells whose formula calls SUBTOTAL or
// AGGREGATE; hidden columns are never skipped; AGGREGATE's options;
// rowVisibility; visibilitySensitiveCells and the recalculation hiding rows
// triggers; an uncomputable formula with no cached value is #NAME? and
// stale; UnsupportedCaseError becomes { kind: "unsupportedCase", name }.
import { describe, expect, it } from "vitest";
import { getCell, parseA1 } from "../../address";
import { calcWorkbook, type CellSpec } from "../../fixtures/calcWorkbook";
import type { CellValue, SpreadsheetModel } from "../../model";
import { cellKey } from "../dependencies";
import { evaluateCell, isUncomputable, rowVisibilityIn, storedValueContext } from "../evaluate";
import { recalculate, rowVisibilityChanged, visibilitySensitiveCells } from "../recalc";

/** A workbook whose rows (1-based) listed in `hidden` are user-hidden. */
function book(cells: Record<string, CellSpec>, hidden: number[] = [], hiddenCols: number[] = []): SpreadsheetModel {
  const model = calcWorkbook({ Sheet1: cells });
  const sheet = model.sheets[0]!;
  const rows = sheet.rows.map((row) => (hidden.includes(row.index + 1) ? { ...row, hidden: true as const } : row));
  model.sheets = [{ ...sheet, rows, columns: hiddenCols.map((col) => ({ min: col, max: col, hidden: true as const })) as typeof sheet.columns }];
  return model;
}

function value(model: SpreadsheetModel, address: string): CellValue {
  recalculate(model, [], { full: true, reason: "manual" });
  const at = parseA1(address)!;
  return getCell(model.sheets[0]!, at.row, at.col)?.value ?? null;
}

const DATA: Record<string, CellSpec> = { A1: 10, A2: 20, A3: 30, A4: 40 };

describe("SUBTOTAL and hidden rows", () => {
  it("keeps user-hidden rows for 1–11 and skips them for 101–111", () => {
    expect(value(book({ ...DATA, C1: "=SUBTOTAL(9,A1:A4)" }, [2]), "C1")).toBe(100);
    expect(value(book({ ...DATA, C1: "=SUBTOTAL(109,A1:A4)" }, [2]), "C1")).toBe(80);
    expect(value(book({ ...DATA, C1: "=SUBTOTAL(102,A1:A4)" }, [2, 3]), "C1")).toBe(2);
    expect(value(book({ ...DATA, C1: "=SUBTOTAL(101,A1:A4)" }, [4]), "C1")).toBe(20);
  });

  it("skips cells whose formula calls SUBTOTAL or AGGREGATE", () => {
    const model = book({ ...DATA, A5: "=SUBTOTAL(9,A1:A4)", A6: "=AGGREGATE(9,4,A1:A2)", C1: "=SUBTOTAL(9,A1:A6)", C2: "=SUM(A1:A6)" });
    expect(value(model, "C1")).toBe(100);
    expect(value(model, "C2")).toBe(230);
  });

  it("never skips hidden columns", () => {
    expect(value(book({ A1: 1, B1: 2, C1: 3, E1: "=SUBTOTAL(109,A1:C1)" }, [], [2]), "E1")).toBe(6);
  });

  it("is #VALUE! for another function number", () => {
    expect(value(book({ ...DATA, C1: "=SUBTOTAL(12,A1:A4)" }), "C1")).toEqual({ error: "#VALUE!" });
  });
});

describe("AGGREGATE's options", () => {
  const withError: Record<string, CellSpec> = { ...DATA, A3: { error: "#DIV/0!" } };

  it("ignores errors with 6 and propagates them with 4", () => {
    expect(value(book({ ...withError, C1: "=AGGREGATE(9,6,A1:A4)" }), "C1")).toBe(70);
    expect(value(book({ ...withError, C1: "=AGGREGATE(9,4,A1:A4)" }), "C1")).toEqual({ error: "#DIV/0!" });
  });

  it("ignores hidden rows with 5 and 7, nested subtotals with 0–3", () => {
    expect(value(book({ ...DATA, C1: "=AGGREGATE(9,5,A1:A4)" }, [1]), "C1")).toBe(90);
    expect(value(book({ ...withError, C1: "=AGGREGATE(9,7,A1:A4)" }, [1]), "C1")).toBe(60);
    expect(value(book({ ...DATA, A5: "=SUM(A1:A4)", A6: "=SUBTOTAL(9,A1:A4)", C1: "=AGGREGATE(9,0,A1:A6)" }), "C1")).toBe(200);
  });

  it("computes the array form (14–19) over computed arrays", () => {
    expect(value(book({ ...DATA, C1: "=AGGREGATE(14,6,A1:A4,2)" }), "C1")).toBe(30);
    expect(value(book({ ...withError, C1: "=AGGREGATE(15,6,A1:A4/(A1:A4>15),1)" }), "C1")).toBe(20);
    expect(value(book({ ...DATA, C1: "=AGGREGATE(20,0,A1:A4)" }), "C1")).toEqual({ error: "#VALUE!" });
  });
});

describe("row visibility and recalculation", () => {
  it("answers hidden for Row.hidden and visible otherwise", () => {
    const model = book(DATA, [2]);
    expect(rowVisibilityIn(model, "sheet-Sheet1", 1)).toBe("hidden");
    expect(rowVisibilityIn(model, "sheet-Sheet1", 0)).toBe("visible");
    expect(rowVisibilityIn(model, "sheet-Sheet1", 99)).toBe("visible");
  });

  it("lists every formula cell calling SUBTOTAL or AGGREGATE", () => {
    const model = book({ ...DATA, B1: "=SUBTOTAL(9,A1:A4)", B2: "=SUM(A1:A4)", B3: "=1+AGGREGATE(9,0,A1:A4)" });
    expect(visibilitySensitiveCells(model)).toEqual([cellKey("sheet-Sheet1", 0, 1), cellKey("sheet-Sheet1", 2, 1)]);
  });

  it("tells a hide or unhide from other row edits, and recalculating those cells follows it", () => {
    const model = book({ ...DATA, C1: "=SUBTOTAL(109,A1:A4)" });
    expect(value(model, "C1")).toBe(100);
    const before = model.sheets[0]!.rows;
    const after = before.map((row) => (row.index === 0 ? { ...row, hidden: true as const } : row));
    expect(rowVisibilityChanged(before, after)).toBe(true);
    expect(rowVisibilityChanged(before, before.map((row) => ({ ...row })))).toBe(false);
    model.sheets = [{ ...model.sheets[0]!, rows: after }];
    recalculate(model, visibilitySensitiveCells(model));
    expect(getCell(model.sheets[0]!, 0, 2)?.value).toBe(90);
  });
});

describe("uncomputable cells (SPEC-127)", () => {
  it("gives a newly typed uncomputable formula #NAME?, stale, and keeps a cached value", () => {
    const model = calcWorkbook({ Sheet1: { A1: "=SEQUENCE(3)", A2: { formula: "SEQUENCE(3)", value: 7 } } });
    const result = recalculate(model, [cellKey("sheet-Sheet1", 0, 0), cellKey("sheet-Sheet1", 1, 0)]);
    expect(getCell(model.sheets[0]!, 0, 0)?.value).toEqual({ error: "#NAME?" });
    expect(getCell(model.sheets[0]!, 1, 0)?.value).toBe(7);
    expect(result.stale).toContain(cellKey("sheet-Sheet1", 0, 0));
  });

  it("turns an UnsupportedCaseError into { kind: \"unsupportedCase\", name }", () => {
    const model = calcWorkbook({ Sheet1: { A1: '=INFO("system")', A2: "=ROMAN(499,1)" } });
    for (const [row, name] of [
      [0, "INFO"],
      [1, "ROMAN"],
    ] as const) {
      const result = evaluateCell(model, "sheet-Sheet1", row, 0, storedValueContext(model));
      expect(isUncomputable(result) && result.uncomputable).toEqual({ kind: "unsupportedCase", name });
    }
  });
});
