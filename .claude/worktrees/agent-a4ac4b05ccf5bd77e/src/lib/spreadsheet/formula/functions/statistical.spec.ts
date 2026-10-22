// SPEC-127 "Statistical" — functions/statistical.ts in a cell of an
// in-memory workbook. "Excel" rows are Microsoft's reference examples at
// the digits the page prints; the others are pinned by hand from the rule
// SPEC-127 states (two-pass variance, PERCENTRANK truncation, MODE.SNGL's
// first-in-order tie, the -A reading, paired positions, legacy names).
import { describe, expect, it } from "vitest";
import { calcWorkbook, type CellSpec } from "../../fixtures/calcWorkbook";
import type { CellValue, ErrorCode } from "../../model";
import { evaluateCell, storedValueContext } from "../evaluate";

type Cells = Record<string, CellSpec>;

function value(formula: string, cells: Cells = {}): unknown {
  const model = calcWorkbook({ Sheet1: { ...cells, Z100: `=${formula}` } });
  return evaluateCell(model, "sheet-Sheet1", 99, 25, storedValueContext(model));
}

/** The result to `digits` significant digits. */
function sig(formula: string, cells: Cells, digits: number): number {
  const result = value(formula, cells);
  expect(typeof result, formula).toBe("number");
  return Number((result as number).toPrecision(digits));
}

const err = (code: ErrorCode): CellValue => ({ error: code });

describe("central tendency and spread", () => {
  const data: Cells = { A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, A6: 6 };
  it.each<[string, number]>([
    ["MEDIAN(A1:A5)", 3], // Excel's MEDIAN example
    ["MEDIAN(A1:A6)", 3.5],
    ["STDEV.P(A1:A5)", Math.sqrt(2)], // Σ(x−3)² = 10, /5
    ["VAR.S(A1:A5)", 2.5], // 10/4
    ["STDEV(A1:A5)", Math.sqrt(2.5)],
    ["VARP(A1:A5)", 2],
    ["DEVSQ(A1:A5)", 10],
    ["AVEDEV(A1:A5)", 1.2], // (2+1+0+1+2)/5
    ["LARGE(A1:A6,2)", 5],
    ["SMALL(A1:A6,2)", 2],
    ["TRIMMEAN(A1:A6,0.4)", 3.5], // drops ⌊6·0.4/2⌋ = 1 at each end
  ])("pinned: %s is %d", (formula, expected) => {
    expect(value(formula, data)).toBeCloseTo(expected, 13);
  });

  it("MODE.SNGL and MODE give the first value with the highest count, #N/A when none repeats", () => {
    expect(value("MODE.SNGL(2,1,1,2,3)")).toBe(2);
    expect(value("MODE(5.6,4,4,3,2,4)")).toBe(4); // Excel's MODE example data
    expect(value("MODE.SNGL(1,2,3)")).toEqual(err("#N/A"));
  });

  it("errors: no numbers, a data error, k out of range", () => {
    expect(value("MEDIAN(A1:A2)")).toEqual(err("#NUM!"));
    expect(value("STDEV.S(1)")).toEqual(err("#DIV/0!"));
    expect(value("VAR.P(A1:A2)", { A1: 1, A2: { error: "#N/A" } })).toEqual(err("#N/A"));
    expect(value("LARGE(A1:A6,7)", data)).toEqual(err("#NUM!"));
    expect(value("GEOMEAN(1,0)")).toEqual(err("#NUM!"));
  });
});

describe("ranks and percentiles", () => {
  // Excel's PERCENTRANK.INC example: 13, 12, 11, 8, 4, 3, 2, 1, 1, 1.
  const prank: Cells = { A2: 13, A3: 12, A4: 11, A5: 8, A6: 4, A7: 3, A8: 2, A9: 1, A10: 1, A11: 1 };
  it.each<[string, number]>([
    ["PERCENTRANK.INC(A2:A11,2)", 0.333],
    ["PERCENTRANK.INC(A2:A11,4)", 0.555],
    ["PERCENTRANK.INC(A2:A11,8)", 0.666],
    ["PERCENTRANK.INC(A2:A11,5)", 0.583],
    ["PERCENTRANK(A2:A11,5)", 0.583],
  ])("Excel: %s is %d", (formula, expected) => {
    expect(value(formula, prank)).toBe(expected);
  });

  // Excel's PERCENTRANK.EXC example: 1, 2, 3, 6, 6, 6, 7, 8, 9.
  const exc: Cells = { A2: 1, A3: 2, A4: 3, A5: 6, A6: 6, A7: 6, A8: 7, A9: 8, A10: 9 };
  it.each<[string, CellValue]>([
    ["PERCENTRANK.EXC(A2:A10,7)", 0.7],
    ["PERCENTRANK.EXC(A2:A10,5.43)", 0.381],
    ["PERCENTRANK.EXC(A2:A10,5.43,1)", 0.3],
    ["PERCENTRANK.EXC(A2:A10,10)", err("#N/A")],
    ["PERCENTRANK.EXC(A2:A10,5,0)", err("#NUM!")],
  ])("Excel/pinned: %s is %j", (formula, expected) => {
    expect(value(formula, exc)).toEqual(expected);
  });

  const ranks: Cells = { A2: 7, A3: 3.5, A4: 3.5, A5: 1, A6: 2 };
  it.each<[string, CellValue]>([
    ["RANK.EQ(A3,A2:A6,1)", 3], // Excel's RANK.EQ example
    ["RANK.EQ(A2,A2:A6,1)", 5],
    ["RANK(A3,A2:A6,1)", 3],
    ["RANK.AVG(A3,A2:A6)", 2.5], // descending: 7 first, the two 3.5s share 2 and 3
    ["RANK.EQ(4,A2:A6)", err("#N/A")],
  ])("%s is %j", (formula, expected) => {
    expect(value(formula, ranks)).toEqual(expected);
  });

  it("percentiles and quartiles through stats.ts", () => {
    const q: Cells = { A1: 1, A2: 2, A3: 4, A4: 7, A5: 8, A6: 9, A7: 10, A8: 12 };
    expect(value("QUARTILE.INC(A1:A8,1)", q)).toBeCloseTo(3.5, 13); // Excel's QUARTILE.INC example
    expect(value("QUARTILE(A1:A8,1)", q)).toBeCloseTo(3.5, 13);
    expect(value("PERCENTILE.INC({1,2,3,4},0.3)")).toBeCloseTo(1.9, 13); // Excel's PERCENTILE.INC example
    expect(value("PERCENTILE.EXC({1,2,3,4},0.1)")).toEqual(err("#NUM!"));
  });
});

describe("paired data", () => {
  // Excel's SLOPE/INTERCEPT example: known_ys 2,3,9,1,8,7,5; known_xs 6,5,11,7,5,4,4.
  const xy: Cells = { A2: 2, A3: 3, A4: 9, A5: 1, A6: 8, A7: 7, A8: 5, B2: 6, B3: 5, B4: 11, B5: 7, B6: 5, B7: 4, B8: 4 };
  it.each<[string, number, number]>([
    ["SLOPE(A2:A8,B2:B8)", 0.305556, 6],
    ["INTERCEPT(A2:A6,B2:B6)", 0.0483871, 6], // Excel's INTERCEPT example: the first five points
    ["STEYX(A2:A8,B2:B8)", 3.305719, 7],
    ["RSQ(A2:A8,B2:B8)", 0.05795, 4],
  ])("Excel: %s is %d", (formula, expected, digits) => {
    expect(sig(formula, xy, digits)).toBe(expected);
  });

  it("CORREL, PEARSON, COVARIANCE.*, COVAR, FORECAST.LINEAR and FORECAST (pinned)", () => {
    const line: Cells = { A1: 1, A2: 2, A3: 3, B1: 2, B2: 4, B3: 6 };
    expect(value("CORREL(A1:A3,B1:B3)", line)).toBeCloseTo(1, 14);
    expect(value("PEARSON(A1:A3,B1:B3)", line)).toBeCloseTo(1, 14);
    expect(value("COVARIANCE.P(A1:A3,B1:B3)", line)).toBeCloseTo(4 / 3, 14); // Σdxdy = 4
    expect(value("COVAR(A1:A3,B1:B3)", line)).toBeCloseTo(4 / 3, 14);
    expect(value("COVARIANCE.S(A1:A3,B1:B3)", line)).toBeCloseTo(2, 14);
    expect(value("FORECAST.LINEAR(4,B1:B3,A1:A3)", line)).toBeCloseTo(8, 13);
    expect(value("FORECAST(4,B1:B3,A1:A3)", line)).toBeCloseTo(8, 13);
  });

  it("uses only positions where both hold numbers; unequal sizes are #N/A; no variance #DIV/0!", () => {
    expect(value("CORREL(A1:A4,B1:B4)", { A1: 1, A2: 2, A3: 3, A4: "x", B1: 2, B2: 4, B3: 6, B4: 100 })).toBeCloseTo(1, 14);
    expect(value("SLOPE(A1:A3,B1:B4)", { A1: 1, B1: 1 })).toEqual(err("#N/A"));
    expect(value("CORREL(A1:A2,B1:B2)", { A1: 1, A2: 1, B1: 2, B2: 3 })).toEqual(err("#DIV/0!"));
    expect(value("COVARIANCE.S({1},{2})")).toEqual(err("#DIV/0!"));
  });
});

describe("distributions", () => {
  it.each<[string, number, number]>([
    ["NORM.DIST(42,40,1.5,TRUE)", 0.9087888, 7], // Excel's NORM.DIST example
    ["NORM.DIST(42,40,1.5,FALSE)", 0.10934, 5],
    ["NORMDIST(42,40,1.5,TRUE)", 0.9087888, 7],
    ["NORM.INV(0.908789,40,1.5)", 42.000002, 8],
    ["NORM.S.DIST(1.333333,TRUE)", 0.908788726, 9],
    ["NORMSDIST(1.333333)", 0.908788726, 9],
    ["NORM.S.INV(0.908789)", 1.3333347, 8],
    ["STANDARDIZE(42,40,1.5)", 1.33333333, 9],
    ["T.DIST(60,1,TRUE)", 0.99469533, 8],
    ["T.DIST(8,3,FALSE)", 0.00073691, 5],
    ["T.DIST.2T(1.959999998,60)", 0.054645, 5],
    ["T.DIST.RT(1.959999998,60)", 0.027322, 5],
    ["TDIST(1.959999998,60,2)", 0.054644930, 8],
    ["T.INV(0.75,2)", 0.8164966, 7],
    ["T.INV.2T(0.546449,60)", 0.606533, 6],
    ["TINV(0.546449,60)", 0.606533, 6],
    ["BINOM.DIST(6,10,0.5,FALSE)", 0.2050781, 7],
    ["BINOM.DIST(6,10,0.5,TRUE)", 0.828125, 6],
    ["POISSON.DIST(2,5,TRUE)", 0.124652, 6],
    ["POISSON.DIST(2,5,FALSE)", 0.084224, 5],
    ["EXPON.DIST(0.2,10,TRUE)", 0.86466472, 8],
    ["EXPON.DIST(0.2,10,FALSE)", 1.35335283, 9],
    ["CONFIDENCE.NORM(0.05,2.5,50)", 0.692952, 6],
    ["GAMMA(2.5)", 1.329340388, 10],
    ["GAMMA(-3.75)", 0.267866129, 9],
    ["GAMMALN(4)", 1.791759469, 10],
    ["GAMMALN.PRECISE(4)", 1.7917595, 8],
  ])("Excel: %s is %d", (formula, expected, digits) => {
    expect(sig(formula, {}, digits)).toBe(expected);
  });

  it.each<[string, string]>([
    ["NORM.DIST(1,0,0,TRUE)", "#NUM!"],
    ["NORM.S.INV(0)", "#NUM!"],
    ["T.DIST.2T(-1,2)", "#NUM!"],
    ["TDIST(1,2,3)", "#NUM!"],
    ["TDIST(-1,2,1)", "#NUM!"],
    ["T.INV(0.5,0)", "#NUM!"],
    ["BINOM.DIST(11,10,0.5,TRUE)", "#NUM!"],
    ["POISSON.DIST(-1,5,TRUE)", "#NUM!"],
    ["EXPON.DIST(0.2,0,TRUE)", "#NUM!"],
    ["CONFIDENCE.NORM(1,2.5,50)", "#NUM!"],
    ["GAMMA(0)", "#NUM!"],
    ["GAMMA(-1)", "#NUM!"],
    ["GAMMALN(0)", "#NUM!"],
  ])("%s is %s", (formula, code) => {
    expect(value(formula)).toEqual({ error: code });
  });
});

describe("the -A variants", () => {
  // A1 10, A2 TRUE, A3 the text "x", A4 empty, A5 FALSE: 10, 1, 0, 0 counted.
  const mixed: Cells = { A1: 10, A2: true, A3: "x", A5: false };
  it.each<[string, number]>([
    ["AVERAGEA(A1:A5)", 2.75],
    ["MAXA(A1:A5)", 10],
    ["MINA(A1:A5)", 0],
    ["VARPA(A1:A5)", 17.6875], // mean 2.75: (52.5625 + 3.0625 + 7.5625 + 7.5625)/4 = 70.75/4
    ["VARA(A1:A5)", 70.75 / 3],
    ["STDEVPA(A1:A5)", Math.sqrt(17.6875)],
    ["STDEVA(A1:A5)", Math.sqrt(70.75 / 3)],
    ["MAXA(A4)", 0],
  ])("pinned: %s is %d", (formula, expected) => {
    expect(value(formula, mixed)).toBeCloseTo(expected, 12);
  });

  it("coerces direct arguments and rejects non-numeric direct text", () => {
    expect(value('AVERAGEA(1,TRUE,"3")')).toBeCloseTo(5 / 3, 14);
    expect(value('AVERAGEA(1,"x")')).toEqual(err("#VALUE!"));
  });
});

describe("shape statistics and FREQUENCY", () => {
  it("KURT, SKEW and SKEW.P (Excel's examples)", () => {
    const k: Cells = { A1: 3, A2: 4, A3: 5, A4: 2, A5: 3, A6: 4, A7: 5, A8: 6, A9: 4, A10: 7 };
    expect(sig("KURT(A1:A10)", k, 9)).toBe(-0.151799637);
    expect(sig("SKEW(A1:A10)", k, 6)).toBe(0.359543);
    expect(sig("SKEW.P(A1:A10)", k, 6)).toBe(0.303193);
    expect(value("KURT(1,2,3)")).toEqual(err("#DIV/0!"));
  });

  it("GEOMEAN and HARMEAN", () => {
    expect(value("GEOMEAN(4,9)")).toBeCloseTo(6, 14);
    expect(value("HARMEAN(1,2,4)")).toBeCloseTo(12 / 7, 14);
  });

  it("FREQUENCY counts per bin as a vertical array over an array formula's range", () => {
    // Excel's FREQUENCY example: scores 79, 85, 78, 85, 50, 81, 95, 88, 97; bins 70, 79, 89.
    const scores: Cells = { A2: 79, A3: 85, A4: 78, A5: 85, A6: 50, A7: 81, A8: 95, A9: 88, A10: 97, B2: 70, B3: 79, B4: 89 };
    const model = calcWorkbook({ Sheet1: { ...scores, C2: { formula: "=FREQUENCY(A2:A10,B2:B4)", array: "C2:C5" } } });
    expect(evaluateCell(model, "sheet-Sheet1", 1, 2, storedValueContext(model))).toEqual({ kind: "array", rows: 4, cols: 1, data: [1, 2, 4, 2] });
    expect(value("FREQUENCY(A2:A10,B2:B4)", scores)).toBe(1); // a plain cell shows the first count
    expect(value("FREQUENCY({1,2,3},{2,1})", {})).toBe(1); // unsorted bins: 2 counts (1,2]
  });
});
