// SPEC-127 "Engineering" — functions/erf.ts: ERF ERF.PRECISE ERFC
// ERFC.PRECISE in a cell of an in-memory workbook. "Excel" rows are
// Microsoft's reference examples (printed to the digits shown); the others
// are A&S Table 7.1 values (10 decimals) or the stated rules.
import { describe, expect, it } from "vitest";
import { calcWorkbook } from "../../fixtures/calcWorkbook";
import { evaluateCell, storedValueContext } from "../evaluate";

function value(formula: string): unknown {
  const model = calcWorkbook({ Sheet1: { Z100: `=${formula}` } });
  return evaluateCell(model, "sheet-Sheet1", 99, 25, storedValueContext(model));
}

function close(formula: string, expected: number, decimals: number): void {
  const result = value(formula);
  expect(typeof result).toBe("number");
  expect(Math.abs((result as number) - expected)).toBeLessThanOrEqual(0.5 * 10 ** -decimals * (1 + 1e-9));
}

describe("ERF, ERF.PRECISE, ERFC, ERFC.PRECISE", () => {
  it("Excel: ERF(0.745) is 0.70792892 and ERF(1,2) 0.15262147", () => {
    close("ERF(0.745)", 0.70792892, 8);
    close("ERF(1,2)", 0.15262147, 8);
  });

  it("Excel: ERF.PRECISE(0.745) 0.70792892, ERF.PRECISE(1) 0.84270079, ERFC(1) and ERFC.PRECISE(1) 0.15729921", () => {
    close("ERF.PRECISE(0.745)", 0.70792892, 8);
    close("ERF.PRECISE(1)", 0.84270079, 8);
    close("ERFC(1)", 0.15729921, 8);
    close("ERFC.PRECISE(1)", 0.15729921, 8);
  });

  it("A&S Table 7.1: erf 0.5 = 0.5204998778, negative arguments by symmetry", () => {
    close("ERF(0.5)", 0.5204998778, 10);
    close("ERF(-0.5)", -0.5204998778, 10);
    close("ERFC(-0.5)", 1.5204998778, 10); // 2 − erfc 0.5 = 1 + erf 0.5
    close("ERF(0.5,0)", -0.5204998778, 10);
  });

  it("is #VALUE! for a non-numeric argument", () => {
    expect(value('ERF("a")')).toEqual({ error: "#VALUE!" });
    expect(value('ERFC("a")')).toEqual({ error: "#VALUE!" });
    expect(value('ERF(1,"a")')).toEqual({ error: "#VALUE!" });
  });
});
