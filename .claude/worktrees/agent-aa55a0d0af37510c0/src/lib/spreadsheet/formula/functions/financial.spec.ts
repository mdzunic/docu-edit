// SPEC-127 Financial — functions/financial.ts, each function evaluated in a
// cell of an in-memory workbook. "Excel" rows are the examples on Microsoft's
// reference pages, compared at the decimals the page prints; the others are
// pinned — hand-derived from the annuity equation, OpenFormula's formulas or
// the rules the module header states, the arithmetic beside each — including
// RATE, IRR and XIRR returning #NUM! where Newton's method cannot converge
// within its cap.
import { describe, expect, it } from "vitest";
import { calcWorkbook, type CellSpec } from "../../fixtures/calcWorkbook";
import type { CellValue, ErrorCode } from "../../model";
import { evaluateCell, storedValueContext } from "../evaluate";

type Cells = Record<string, CellSpec>;

/** `formula` evaluated in Sheet1!Z100 over `cells` — the value it stores. */
function value(formula: string, cells: Cells = {}) {
  const model = calcWorkbook({ Sheet1: { ...cells, Z100: `=${formula}` } });
  return evaluateCell(model, "sheet-Sheet1", 99, 25, storedValueContext(model));
}

/** The formula's number, failing the test for anything else. */
function num(formula: string, cells: Cells = {}): number {
  const result = value(formula, cells);
  if (typeof result !== "number") {
    throw new Error(`${formula} gave ${JSON.stringify(result)}`);
  }
  return result;
}

/** Rounded to `places` decimals, as Microsoft's pages print. */
const at = (x: number, places: number): number => Math.round(x * 10 ** places) / 10 ** places;

const err = (code: ErrorCode): CellValue => ({ error: code });

describe("Microsoft's examples", () => {
  it.each<[string, number, number, Cells?]>([
    ["PMT(A2/12,A3,A4)", -1037.03, 2, { A2: 0.08, A3: 10, A4: 10000 }],
    ["PMT(A2/12,A3,A4,,1)", -1030.16, 2, { A2: 0.08, A3: 10, A4: 10000 }],
    ["PMT(0.06/12,18*12,0,50000)", -129.08, 2],
    ["IPMT(0.1/12,1,3*12,8000)", -66.67, 2],
    ["IPMT(0.1,3,3,8000)", -292.45, 2],
    ["PPMT(0.1/12,1,2*12,2000)", -75.62, 2],
    ["PPMT(0.08,10,10,200000)", -27598.05, 2],
    ["FV(0.06/12,10,-200,-500,1)", 2581.4, 2],
    ["FV(0.12/12,12,-1000)", 12682.5, 2],
    ["FV(0.11/12,35,-2000,,1)", 82846.25, 2],
    ["FV(0.06/12,12,-100,-1000,1)", 2301.4, 2],
    ["PV(0.08/12,12*20,500,,0)", -59777.15, 2],
    ["NPER(0.12/12,-100,-1000,10000,1)", 59.6738657, 7],
    ["NPER(0.12/12,-100,-1000,10000)", 60.0821229, 7],
    ["NPER(0.12/12,-100,-1000)", -9.57859404, 8],
    ["RATE(4*12,-200,8000)*12", 0.0924, 4],
    ["NPV(0.1,-10000,3000,4200,6800)", 1188.44, 2],
    ["NPV(0.08,A4:A8)+A3", 1922.06, 2, { A3: -40000, A4: 8000, A5: 9200, A6: 10000, A7: 12000, A8: 14500 }],
    ["NPV(0.08,A4:A8,-9000)+A3", -3749.47, 2, { A3: -40000, A4: 8000, A5: 9200, A6: 10000, A7: 12000, A8: 14500 }],
    ["IRR(A2:A6)", -0.021, 3, { A2: -70000, A3: 12000, A4: 15000, A5: 18000, A6: 21000, A7: 26000 }],
    ["IRR(A2:A7)", 0.087, 3, { A2: -70000, A3: 12000, A4: 15000, A5: 18000, A6: 21000, A7: 26000 }],
    ["IRR(A2:A4,-10%)", -0.444, 3, { A2: -70000, A3: 12000, A4: 15000 }],
    ["MIRR(A2:A7,A8,A9)", 0.13, 2, { A2: -120000, A3: 39000, A4: 30000, A5: 21000, A6: 37000, A7: 46000, A8: 0.1, A9: 0.12 }],
    ["MIRR(A2:A5,A8,A9)", -0.05, 2, { A2: -120000, A3: 39000, A4: 30000, A5: 21000, A8: 0.1, A9: 0.12 }],
    ["MIRR(A2:A7,A8,14%)", 0.13, 2, { A2: -120000, A3: 39000, A4: 30000, A5: 21000, A6: 37000, A7: 46000, A8: 0.1 }],
    ["SLN(30000,7500,10)", 2250, 2],
    ["SYD(30000,7500,10,1)", 4090.91, 2],
    ["SYD(30000,7500,10,10)", 409.09, 2],
    ["DDB(2400,300,10*365,1)", 1.32, 2],
    ["DDB(2400,300,10*12,1,2)", 40, 2],
    ["DDB(2400,300,10,1,2)", 480, 2],
    ["DDB(2400,300,10,2,1.5)", 306, 2],
    ["DDB(2400,300,10,10)", 22.12, 2],
    ["CUMIPMT(0.09/12,30*12,125000,13,24,0)", -11135.23213, 5],
    ["CUMIPMT(0.09/12,30*12,125000,1,1,0)", -937.5, 5],
    ["CUMPRINC(0.09/12,30*12,125000,13,24,0)", -934.1071234, 7],
    ["CUMPRINC(0.09/12,30*12,125000,1,1,0)", -68.27827118, 8],
    ["EFFECT(0.0525,4)", 0.0535427, 7],
    ["NOMINAL(0.053543,4)", 0.05250032, 8],
  ])("Excel: %s is %d", (formula, expected, places, cells) => {
    expect(at(num(formula, cells), places)).toBe(expected);
  });

  it("DB: the seven years of the 1,000,000 / 100,000 / 6-year example, month 7", () => {
    const expected = [186083.33, 259639.42, 176814.44, 120410.64, 81999.64, 55841.76, 15845.1];
    expected.forEach((amount, i) => {
      expect(at(num(`DB(1000000,100000,6,${i + 1},7)`), 2), `period ${i + 1}`).toBe(amount);
    });
  });

  it("XNPV and XIRR over the 2008–2009 schedule", () => {
    // 1/1/2008 39448, 3/1/2008 39508, 10/30/2008 39751, 2/15/2009 39859, 4/1/2009 39904.
    const cells: Cells = { A1: -10000, A2: 2750, A3: 4250, A4: 3250, A5: 2750, B1: 39448, B2: 39508, B3: 39751, B4: 39859, B5: 39904 };
    expect(at(num("XNPV(0.09,A1:A5,B1:B5)", cells), 2)).toBe(2086.65);
    // Printed 37.34% (0.373362535); XIRR itself stops at |Δrate| < 1e-8, so compare within 5e-9.
    expect(at(num("XIRR(A1:A5,B1:B5,0.1)", cells), 4)).toBe(0.3734);
    expect(num("XIRR(A1:A5,B1:B5,0.1)", cells)).toBeCloseTo(0.373362535, 8);
  });
});

describe("pinned annuity arithmetic", () => {
  it.each<[string, number]>([
    ["PMT(0,10,1000)", -100], // rate 0: −(pv + fv)/nper
    ["PMT(0,10,1000,-200)", -80],
    ["FV(0.1,2,-100)", 210], // 100·1.1 + 100
    ["FV(0.1,2,-100,0,1)", 231], // 100·1.21 + 100·1.1
    ["FV(0,5,-10,-50)", 100], // rate 0: −(pv + pmt·nper)
    ["PV(0.1,2,-100)", 21000 / 121], // 100/1.1 + 100/1.21
    ["PV(0,4,-25)", 100],
    ["NPER(0,-100,1000)", 10], // rate 0: −(pv + fv)/pmt
    ["NPER(0.1,0,-100,121)", 2], // 1.1^n = 1.21
    ["IPMT(0,2,10,1000)", 0], // no interest at rate 0
    ["PPMT(0,2,10,1000)", -100],
    ["IPMT(0.1,1,2,1000,0,1)", 0], // the first payment, due at the start, carries no interest
    ["NPV(0.1,110,121)", 200], // 110/1.1 + 121/1.21
    ['NPV(0.1,"110",TRUE,110)', 100], // only numbers count: text and logicals are skipped, even as arguments
    ["NPV(0.1,110,1/0,121)", 200], // so is an error value: 110/1.1 + 121/1.21
    ["NPV(0.1,A1:A3)", 200], // A2's text is skipped: 110/1.1 + 121/1.21
    ["SLN(1000,100,9)", 100],
    ["SYD(1000,100,2,1)", 600], // 900·2·2/(2·3)
    ["EFFECT(0.1,1.9)", 0.1], // npery truncated to 1
    ["NOMINAL(0.21,2)", 0.2], // 2·(1.21^(1/2) − 1)
    ["MIRR({-100,50,60},0.1,0.2)", Math.sqrt(1.2) - 1], // (50·1.2 + 60)/100 = 1.2 over two periods
    ["XNPV(0.1,{-100,110},{39448,39813})", 0], // 365 days apart: −100 + 110/1.1
    ["XNPV(0.1,{-100,110},{39448.9,39813.2})", 0], // dates truncated
    ["DDB(1000,0,4,1,4)", 1000], // rate factor/life = 1: everything in period 1
    ["DB(1000,0,4,2)", 0], // rate 1: all of it in period 1
  ])("%s is %d", (formula, expected) => {
    expect(num(formula, { A1: 110, A2: "121", A3: 121 })).toBeCloseTo(expected, 9);
  });

  it("solves RATE, IRR and XIRR to the root", () => {
    expect(num("RATE(2,0,-100,121)")).toBeCloseTo(0.1, 9); // 100·1.1² = 121
    expect(num("RATE(10,-100,1000)")).toBeCloseTo(0, 9); // pv + pmt·nper = 0
    expect(num("IRR({-100,110})")).toBeCloseTo(0.1, 9);
    expect(num("IRR({-100,0,121})")).toBeCloseTo(0.1, 9);
    expect(num('IRR({-100,"x",110})')).toBeCloseTo(0.1, 9); // text in the array is skipped
    expect(num("XIRR({-100,110},{39448,39813})")).toBeCloseTo(0.1, 9);
  });

  it("keeps RATE's root to the last digit where 1 + rate would round it", () => {
    // −100·(1 + r) + 110 = 0 at r = 0.1 exactly; with 1 + 0.1 formed first
    // the residual at 0.1 is −1.4e-14 and Newton steps off the root.
    expect(num("RATE(1,0,-100,110)")).toBe(0.1);
    expect(Number(num("RATE(2,0,-100,121)").toPrecision(15))).toBe(0.1);
  });

  it("DB: the rate rounded to three decimals, the period after life, and later periods", () => {
    // rate = 1 − 0.5^(1/2) = 0.29289… → 0.293; period 1 = 1000·0.293·6/12.
    expect(num("DB(1000,500,2,1,6)")).toBeCloseTo(146.5, 9);
    // period 2 = (1000 − 146.5)·0.293 = 250.0755; period 3 (after life) = (853.5 − 250.0755)·0.293·6/12.
    expect(num("DB(1000,500,2,2,6)")).toBeCloseTo(250.0755, 9);
    expect(num("DB(1000,500,2,3,6)")).toBeCloseTo(603.4245 * 0.293 * 0.5, 9);
    expect(num("DB(1000,500,2,4,6)")).toBe(0);
    expect(num("DB(1000,500,2,3)")).toBe(0); // month 12: nothing after life
  });
});

describe("error cases", () => {
  it.each<[string, CellValue]>([
    ["PMT(0.1,0,1000)", err("#NUM!")],
    ['PMT("a",10,1000)', err("#VALUE!")],
    ["IPMT(0.1,0,3,8000)", err("#NUM!")], // per outside 1…nper
    ["PPMT(0.1,4,3,8000)", err("#NUM!")],
    ["NPER(0,0,1000)", err("#NUM!")],
    ["NPER(0.1,-10,1000)", err("#NUM!")], // the payment never covers the interest
    ["RATE(0,-100,1000)", err("#NUM!")],
    ["RATE(10,100,1000)", err("#NUM!")], // every flow an inflow: no root
    ["RATE(2,0,-1,1E+40)", err("#NUM!")], // root 1E+20 − 1: Newton halves toward it, past 20 steps
    ["RATE(48,-200,8000,0,0,-1)", err("#NUM!")], // a guess at −1
    ["NPV(-1,100)", err("#DIV/0!")],
    ["IRR({100,200})", err("#NUM!")], // no negative flow
    ["IRR({-1,1E+10})", err("#NUM!")], // 1 + rate doubles each step: 20 are not enough
    ["MIRR({100,200},0.1,0.1)", err("#DIV/0!")],
    ["XNPV(0.1,{-100,110},{39448})", err("#NUM!")], // counts differ
    ['XNPV(0.1,{-100,"x"},{39448,39813})', err("#VALUE!")], // a non-number
    ["XNPV(0.1,{-100,110},{39448,39447})", err("#NUM!")], // a date before the first
    ["XNPV(0.1,{-100,110},{39448,-1})", err("#VALUE!")], // not a date
    ["XNPV(-1,{-100,110},{39448,39813})", err("#NUM!")],
    ["XIRR({100,110},{39448,39813})", err("#NUM!")], // no negative flow
    ["XIRR({-1,1E+40},{39448,39813})", err("#NUM!")], // 100 doublings fall short of 1E+40
    ["SLN(1000,100,0)", err("#DIV/0!")],
    ["SYD(1000,100,0,1)", err("#DIV/0!")],
    ["DB(0,0,5,1)", err("#NUM!")],
    ["DB(1000,100,5,1,13)", err("#NUM!")],
    ["DB(1000,100,5,0.5)", err("#NUM!")],
    ["DDB(1000,100,5,6)", err("#NUM!")], // period past life
    ["DDB(1000,1100,5,1)", err("#NUM!")], // salvage above cost
    ["DDB(1000,100,5,1,0)", err("#NUM!")],
    ["CUMIPMT(0,360,125000,1,1,0)", err("#NUM!")],
    ["CUMIPMT(0.0075,360,125000,0,1,0)", err("#NUM!")],
    ["CUMIPMT(0.0075,360,125000,2,1,0)", err("#NUM!")],
    ["CUMIPMT(0.0075,360,125000,1,361,0)", err("#NUM!")],
    ["CUMPRINC(0.0075,360,125000,1,1,2)", err("#NUM!")],
    ["CUMPRINC(0.0075,360,-125000,1,1,0)", err("#NUM!")],
    ["EFFECT(0,4)", err("#NUM!")],
    ["EFFECT(0.05,0.9)", err("#NUM!")],
    ['EFFECT("a",4)', err("#VALUE!")],
    ["NOMINAL(-0.1,4)", err("#NUM!")],
  ])("%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});
