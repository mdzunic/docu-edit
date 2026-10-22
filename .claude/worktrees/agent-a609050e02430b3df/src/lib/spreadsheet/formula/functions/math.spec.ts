// SPEC-120 "The core functions" — functions/math.ts: SUM SUMIF SUMIFS
// SUMPRODUCT PRODUCT ROUND ROUNDUP ROUNDDOWN INT TRUNC MOD ABS SIGN SQRT POWER
// EXP LN LOG LOG10 PI CEILING FLOOR MROUND RAND RANDBETWEEN, each evaluated in
// a cell of an in-memory workbook. Rows labelled "Excel" are Microsoft's
// function-reference examples or documented rules; the others are pinned —
// computed by hand from the rule SPEC-120 states: rounding on the
// 15-significant-digit form (ROUND(1.005,2) is 1.01, FLOOR(0.3,0.1) is 0.3),
// CEILING/FLOOR/MROUND's sign rules and zero significance, MOD taking the
// divisor's sign, PRODUCT of nothing 0, SUMIF resizing its sum range, SUMIFS
// needing one shape, SUMPRODUCT's equal dimensions and non-numbers read as 0,
// numbers only from references but direct arguments coerced, errors in a
// range propagating, whole columns over sparse data, and RAND/RANDBETWEEN
// drawing from an injected ctx.random().
import { describe, expect, it } from "vitest";
import { calcWorkbook, type CellSpec } from "../../fixtures/calcWorkbook";
import type { CellValue, ErrorCode } from "../../model";
import { SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "../../numberFormat/locale";
import { evaluateCell, storedValueContext } from "../evaluate";

type Cells = Record<string, CellSpec>;

/** `formula` evaluated in Sheet1!Z100 over `cells` — the value it stores. */
function value(formula: string, cells: Cells = {}, options: { locale?: SpreadsheetLocaleTag; random?: () => number } = {}) {
  const model = calcWorkbook({ Sheet1: { ...cells, Z100: `=${formula}` } });
  const base = storedValueContext(model, { locale: SPREADSHEET_LOCALES[options.locale ?? "en-US"] });
  return evaluateCell(model, "sheet-Sheet1", 99, 25, options.random ? { ...base, random: options.random } : base);
}

const err = (code: ErrorCode): CellValue => ({ error: code });

/** mulberry32 — a small seeded PRNG, so the draws are the same every run. */
function seeded(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

describe("SUM", () => {
  // Excel's SUM reference example: A2 −5, A3 15, A4 30, A5 the text "5", A6 TRUE.
  const sheet: Cells = { A2: -5, A3: 15, A4: 30, A5: "5", A6: true };

  it.each<[string, CellValue]>([
    ["SUM(3,2)", 5],
    ['SUM("5",15,TRUE)', 21],
    ["SUM(A2:A4)", 40],
    ["SUM(A2:A4,15)", 55],
    ["SUM(A5,A6,2)", 2], // text and logicals in references are ignored
  ])("Excel: %s is %j", (formula, expected) => {
    expect(value(formula, sheet)).toEqual(expected);
  });

  it.each<[string, CellValue, Cells]>([
    ['SUM("3",TRUE)', 4, {}], // direct arguments coerce (Excel)
    ["SUM(A1:A2)", 0, { A1: "3", A2: true }], // a reference reads numbers only (Excel)
    ['SUM("a")', err("#VALUE!"), {}],
    ["SUM({1,2;3,4})", 10, {}],
    ['SUM({1,"2",TRUE})', 1, {}], // an array reads numbers only
    ["SUM(1,,2)", 3, {}], // an empty slot reads as 0
    ["SUM(A1:A3)", err("#DIV/0!"), { A1: 1, A2: { error: "#DIV/0!" }, A3: 3 }],
    ["SUM(A1:B2)", err("#N/A"), { A1: 1, B1: { error: "#N/A" }, A2: { error: "#DIV/0!" } }], // the first error, row-major
    ["SUM(1,#N/A)", err("#N/A"), {}],
  ])("%s is %j", (formula, expected, cells) => {
    expect(value(formula, cells)).toEqual(expected);
  });

  it("sums a whole column over sparse data, reading only its stored cells", () => {
    const model = calcWorkbook({ Sheet1: { A1: 1, A1000: 2, A1048576: 3, B5: 100, Z100: "=SUM(A:A)" } });
    const base = storedValueContext(model);
    let reads = 0;
    const counting = {
      ...base,
      valueAt(sheetId: string, row: number, col: number) {
        reads += 1;
        return base.valueAt(sheetId, row, col);
      },
    };
    expect(evaluateCell(model, "sheet-Sheet1", 99, 25, counting)).toBe(6);
    expect(reads).toBe(3);
  });

  it("coerces a direct text argument in the user's locale", () => {
    expect(value('SUM("1,5",1)', {}, { locale: "de-DE" })).toBe(2.5);
    expect(value('SUM("1,5",1)')).toEqual(err("#VALUE!"));
  });
});

describe("SUMIF", () => {
  // Excel's SUMIF reference, example 1: property values A2:A5, commissions B2:B5, C2 250,000.
  const property: Cells = { A2: 100_000, A3: 200_000, A4: 300_000, A5: 400_000, B2: 7_000, B3: 14_000, B4: 21_000, B5: 28_000, C2: 250_000 };
  // Example 2: categories A2:A7 (A5 blank), foods B2:B7, sales C2:C7.
  const groceries: Cells = {
    A2: "Vegetables", B2: "Tomatoes", C2: 2_300, A3: "Vegetables", B3: "Celery", C3: 5_500,
    A4: "Fruits", B4: "Oranges", C4: 800, B5: "Butter", C5: 400,
    A6: "Vegetables", B6: "Carrots", C6: 4_200, A7: "Fruits", B7: "Apples", C7: 1_200,
  };

  it.each<[string, CellValue, Cells]>([
    ['SUMIF(A2:A5,">160000",B2:B5)', 63_000, property],
    ['SUMIF(A2:A5,">160000")', 900_000, property],
    ["SUMIF(A2:A5,300000,B2:B5)", 21_000, property],
    ['SUMIF(A2:A5,">"&C2,B2:B5)', 49_000, property],
    ['SUMIF(A2:A7,"Fruits",C2:C7)', 2_000, groceries],
    ['SUMIF(A2:A7,"Vegetables",C2:C7)', 12_000, groceries],
    ['SUMIF(B2:B7,"*es",C2:C7)', 4_300, groceries],
    ['SUMIF(A2:A7,"",C2:C7)', 400, groceries],
  ])("Excel: %s is %j", (formula, expected, cells) => {
    expect(value(formula, cells)).toEqual(expected);
  });

  // A1:A4 1, −1, 2, −2 beside B1:B4 10, 20, 30, 40.
  const signs: Cells = { A1: 1, A2: -1, A3: 2, A4: -2, B1: 10, B2: 20, B3: 30, B4: 40 };

  it.each([['SUMIF(A1:A4,">0",B1:B4)'], ['SUMIF(A1:A4,">0",B1)'], ['SUMIF(A1:A4,">0",B1:B2)']])(
    "%s resizes the sum range from its top-left cell to the criteria range's shape (Excel): 40",
    (formula) => {
      expect(value(formula, signs)).toBe(40);
    },
  );

  it.each<[string, CellValue, Cells]>([
    // A 2-D criteria range A1:B2 resizes C1 to C1:D2 (Excel's documented table).
    ['SUMIF(A1:B2,">0",C1)', 50, { A1: 1, B1: -1, A2: -2, B2: 3, C1: 10, D1: 20, C2: 30, D2: 40 }],
    // An empty criteria cell is not "x", so its row is summed.
    ['SUMIF(A1:A3,"<>x",B1:B3)', 6, { A1: "x", A2: "y", B1: 1, B2: 2, B3: 4 }],
    // The sum range reads numbers only.
    ['SUMIF(A1:A3,"x",B1:B3)', 1, { A1: "x", A2: "x", A3: "x", B1: 1, B2: "5", B3: true }],
    // An error the criterion leaves out is ignored; one it selects propagates.
    ['SUMIF(A1:A3,"<>#N/A")', 4, { A1: 1, A2: { error: "#N/A" }, A3: 3 }],
    ['SUMIF(A1:A2,"x",B1:B2)', err("#DIV/0!"), { A1: "x", A2: "x", B1: 1, B2: { error: "#DIV/0!" } }],
  ])("%s is %j", (formula, expected, cells) => {
    expect(value(formula, cells)).toEqual(expected);
  });

  it("matches whole columns over sparse data, empty criteria cells included", () => {
    const cells: Cells = { A1: "x", B1: 1, A5: "y", B5: 100, A1000: "x", B1000: 2, B2000: 1_000 };
    expect(value('SUMIF(A:A,"x",B:B)', cells)).toBe(3);
    expect(value('SUMIF(A:A,"<>x",B:B)', cells)).toBe(1_100); // B5 beside "y", B2000 beside an empty cell
  });
});

describe("SUMIFS", () => {
  // Excel's SUMIFS reference example: quantities A2:A9, products B2:B9, salespeople C2:C9.
  const sales: Cells = {
    A2: 5, B2: "Apples", C2: "Tom", A3: 4, B3: "Apples", C3: "Sarah",
    A4: 15, B4: "Artichokes", C4: "Tom", A5: 3, B5: "Artichokes", C5: "Sarah",
    A6: 22, B6: "Bananas", C6: "Tom", A7: 12, B7: "Bananas", C7: "Sarah",
    A8: 10, B8: "Carrots", C8: "Tom", A9: 33, B9: "Carrots", C9: "Sarah",
  };

  it.each<[string, CellValue]>([
    ['SUMIFS(A2:A9,B2:B9,"=A*",C2:C9,"Tom")', 20],
    ['SUMIFS(A2:A9,B2:B9,"<>Bananas",C2:C9,"Tom")', 30],
  ])("Excel: %s is %j", (formula, expected) => {
    expect(value(formula, sales)).toEqual(expected);
  });

  it.each([
    ['SUMIFS(A2:A9,B2:B8,"=A*")'],
    ['SUMIFS(A2:A9,B2:B9,"=A*",C2:C8,"Tom")'],
    ['SUMIFS(A2:B3,C2:C5,">0")'], // as many cells, another shape
  ])("%s is #VALUE!: every range must have the sum range's shape", (formula) => {
    expect(value(formula, sales)).toEqual(err("#VALUE!"));
  });
});

describe("SUMPRODUCT", () => {
  it("Excel: SUMPRODUCT(A2:B4,C2:D4) is 156 (the reference example)", () => {
    const arrays: Cells = { A2: 3, B2: 4, C2: 2, D2: 7, A3: 8, B3: 6, C3: 6, D3: 7, A4: 1, B4: 9, C4: 5, D4: 3 };
    expect(value("SUMPRODUCT(A2:B4,C2:D4)", arrays)).toBe(156);
    expect(value("SUMPRODUCT(A2:B4,C2:C4)", arrays)).toEqual(err("#VALUE!")); // unequal dimensions
  });

  const ramp: Cells = { A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30 };

  it.each<[string, CellValue, Cells]>([
    ["SUMPRODUCT({1,2,3},{4,5,6})", 32, {}],
    ["SUMPRODUCT({1,2,3},{1;2;3})", err("#VALUE!"), {}], // a row and a column differ in dimensions
    ["SUMPRODUCT(A1:A3,B1:B3)", 10, { A1: 1, A2: "2", A3: true, B1: 10, B2: 20, B3: 30 }], // non-numbers are 0 (Excel)
    ["SUMPRODUCT((A1:A3>1)*B1:B3)", 50, ramp], // an array argument needs no CSE
    ["SUMPRODUCT(A1:A3>1)", 0, ramp], // logicals are non-numbers
    ["SUMPRODUCT(--(A1:A3>1))", 2, ramp],
    ["SUMPRODUCT(A1:A2,B1:B2)", err("#N/A"), { A1: 1, A2: 2, B1: 3, B2: { error: "#N/A" } }],
    ["SUMPRODUCT(A:A,B:B)", 26, { A1: 2, B1: 3, A5: 4, B5: 5, A9: 7 }], // whole columns: 2·3 + 4·5 + 7·0
  ])("%s is %j", (formula, expected, cells) => {
    expect(value(formula, cells)).toEqual(expected);
  });
});

describe("PRODUCT", () => {
  it.each<[string, CellValue, Cells]>([
    ["PRODUCT(A2:A4)", 2_250, { A2: 5, A3: 15, A4: 30 }], // Excel's reference example
    ["PRODUCT(A2:A4,2)", 4_500, { A2: 5, A3: 15, A4: 30 }], // Excel's reference example
    ["PRODUCT(A1:A3)", 0, {}], // of no numbers
    ["PRODUCT(A1:A2)", 0, { A1: "x", A2: true }],
    ["PRODUCT(A1:A3)", 2, { A1: 2, A2: "3", A3: true }], // a reference reads numbers only
    ['PRODUCT("2",TRUE,4)', 8, {}], // direct arguments coerce
    ['PRODUCT("x")', err("#VALUE!"), {}],
    ["PRODUCT(A1:A2,3)", err("#NUM!"), { A1: 2, A2: { error: "#NUM!" } }],
  ])("%s is %j", (formula, expected, cells) => {
    expect(value(formula, cells)).toEqual(expected);
  });
});

describe("ROUND, ROUNDUP, ROUNDDOWN", () => {
  it.each<[string, number]>([
    // Excel's ROUND reference examples.
    ["ROUND(2.15,1)", 2.2], ["ROUND(2.149,1)", 2.1],
    ["ROUND(-1.475,2)", -1.48], ["ROUND(21.5,-1)", 20],
    ["ROUND(626.3,-3)", 1_000], ["ROUND(1.98,-1)", 0],
    ["ROUND(-50.55,-2)", -100],
    // Excel's ROUNDUP reference examples.
    ["ROUNDUP(3.2,0)", 4], ["ROUNDUP(76.9,0)", 77],
    // eslint-disable-next-line approx-constant -- Excel's documented result, not an approximation of π
    ["ROUNDUP(3.14159,3)", 3.142], ["ROUNDUP(-3.14159,1)", -3.2],
    ["ROUNDUP(31415.92654,-2)", 31_500],
    // Excel's ROUNDDOWN reference examples.
    ["ROUNDDOWN(3.2,0)", 3], ["ROUNDDOWN(76.9,0)", 76],
    // eslint-disable-next-line approx-constant -- Excel's documented result, not an approximation of π
    ["ROUNDDOWN(3.14159,3)", 3.141], ["ROUNDDOWN(-3.14159,1)", -3.1],
    ["ROUNDDOWN(31415.92654,-2)", 31_400],
    // Halves away from zero (Excel).
    ["ROUND(2.5,0)", 3], ["ROUND(-2.5,0)", -3],
  ])("Excel: %s is %d", (formula, expected) => {
    expect(value(formula)).toBe(expected);
  });

  // Pinned: the 15-significant-digit decimal form is what rounds, never the binary value.
  it.each<[string, number]>([
    ["ROUND(1.005,2)", 1.01], // 1.00499999999999989… in binary
    ["ROUND(2.675,2)", 2.68], // 2.67499999999999982…
    ["ROUNDUP(0.1+0.2,1)", 0.3], // 0.30000000000000004 is 0.3 at 15 digits
    ["ROUNDDOWN(1-0.9,1)", 0.1], // 0.09999999999999998 is 0.1 at 15 digits
  ])("pinned: %s is %d", (formula, expected) => {
    expect(value(formula)).toBe(expected);
  });
});

describe("INT, TRUNC, MOD, ABS, SIGN, SQRT, POWER", () => {
  it.each<[string, CellValue, Cells?]>([
    // Excel's reference examples.
    ["INT(8.9)", 8], ["INT(-8.9)", -9], ["A2-INT(A2)", 0.5, { A2: 19.5 }],
    ["TRUNC(8.9)", 8], ["TRUNC(-8.9)", -8], ["TRUNC(0.45)", 0],
    ["TRUNC(-4.3)", -4], ["INT(-4.3)", -5], // they differ only for negatives (Excel's remark)
    ["MOD(3,2)", 1], ["MOD(-3,2)", 1], ["MOD(3,-2)", -1], ["MOD(-3,-2)", -1],
    ["MOD(3,0)", err("#DIV/0!")],
    ["ABS(2)", 2], ["ABS(-2)", 2], ["ABS(A2)", 4, { A2: -4 }],
    ["SIGN(10)", 1], ["SIGN(4-4)", 0], ["SIGN(-0.00001)", -1],
    ["SQRT(16)", 4], ["SQRT(A2)", err("#NUM!"), { A2: -16 }], ["SQRT(ABS(A2))", 4, { A2: -16 }],
    ["POWER(5,2)", 25],
  ])("Excel: %s is %j", (formula, expected, cells = {}) => {
    expect(value(formula, cells)).toEqual(expected);
  });

  it.each<[string, CellValue]>([
    ["TRUNC(3.14159,2)", 3.14],
    ["TRUNC(1234.5,-2)", 1_200],
    ["MOD(5.5,2)", 1.5], // n − d·INT(n/d)
    ['SQRT("8+1")', err("#VALUE!")], // Excel (the operators article)
    ["POWER(-2,3)", -8],
    ["POWER(0,0)", err("#NUM!")],
    ["POWER(-8,1/3)", err("#NUM!")], // a negative base, a non-integer exponent
    ["POWER(0,-1)", err("#DIV/0!")],
    ["POWER(10,400)", err("#NUM!")], // a non-finite result
  ])("%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });

  it("POWER's fractional reference examples", () => {
    expect(value("POWER(98.6,3.2)")).toBeCloseTo(2_401_077.222, 2);
    expect(value("POWER(4,5/4)")).toBeCloseTo(5.656854249, 8);
  });
});

describe("EXP, LN, LOG, LOG10, PI", () => {
  it.each<[string, number, number, Cells?]>([
    // Excel's reference examples, to the digits they display (EXP(1) is e, PI() π).
    ["EXP(1)", Math.E, 8], ["EXP(2)", 7.3890561, 7],
    ["LN(86)", 4.4543473, 6], ["LN(2.7182818)", 1, 6], ["LN(EXP(3))", 3, 12],
    ["LOG(10)", 1, 12], ["LOG(8,2)", 3, 12], ["LOG(86,2.7182818)", 4.4543473, 6],
    ["LOG10(86)", 1.934498451, 9], ["LOG10(10)", 1, 12], ["LOG10(1E5)", 5, 12], ["LOG10(10^5)", 5, 12],
    ["PI()", Math.PI, 14], ["PI()/2", 1.570796327, 9], ["PI()*(A3^2)", 28.27433388, 8, { A3: 3 }],
  ])("Excel: %s ≈ %d", (formula, expected, digits, cells = {}) => {
    expect(value(formula, cells)).toBeCloseTo(expected, digits);
  });

  it.each<[string, CellValue]>([
    ["LN(0)", err("#NUM!")],
    ["LN(-1)", err("#NUM!")],
    ["LOG(0)", err("#NUM!")],
    ["LOG(-10)", err("#NUM!")],
    ["LOG(8,1)", err("#DIV/0!")], // the logarithm of the base is 0
    ["LOG10(0)", err("#NUM!")],
    ["LOG10(-1)", err("#NUM!")],
    ["EXP(1000)", err("#NUM!")], // a non-finite result
  ])("%s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});

describe("CEILING, FLOOR, MROUND", () => {
  it.each<[string, CellValue]>([
    // Excel's CEILING reference examples.
    ["CEILING(2.5,1)", 3], ["CEILING(-2.5,-2)", -4], ["CEILING(-2.5,2)", -2],
    ["CEILING(1.5,0.1)", 1.5], ["CEILING(0.234,0.01)", 0.24],
    // Excel's FLOOR reference examples.
    ["FLOOR(3.7,2)", 2], ["FLOOR(-2.5,-2)", -2], ["FLOOR(2.5,-2)", err("#NUM!")],
    ["FLOOR(1.58,0.1)", 1.5], ["FLOOR(0.234,0.01)", 0.23],
    // Excel's MROUND reference examples.
    ["MROUND(10,3)", 9], ["MROUND(-10,-3)", -9], ["MROUND(1.3,0.2)", 1.4], ["MROUND(5,-2)", err("#NUM!")],
    // Excel's documented sign rules.
    ["CEILING(2.5,-2)", err("#NUM!")], // a positive number, a negative significance
    ["CEILING(-0.5,1)", 0], // negative with a positive significance: up toward zero (and never −0)
    ["FLOOR(-2.5,2)", -4], // negative with a positive significance: down, away from zero
    ["MROUND(5,2)", 6], ["MROUND(-5,-2)", -6], // a remainder of half the multiple rounds away from zero
  ])("Excel: %s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });

  it.each<[string, CellValue]>([
    ["CEILING(2.5,0)", 0],
    ["FLOOR(2.5,0)", err("#DIV/0!")],
    ["FLOOR(0.3,0.1)", 0.3], // the quotient 2.9999999999999996 is 3 at 15 digits
    ["CEILING(1.1,0.1)", 1.1], // the quotient 11.000000000000002 is 11 at 15 digits
  ])("pinned: %s is %j", (formula, expected) => {
    expect(value(formula)).toEqual(expected);
  });
});

describe("RAND and RANDBETWEEN", () => {
  it("RAND() is ctx.random(), never Math.random", () => {
    expect(value("RAND()", {}, { random: () => 0.25 })).toBe(0.25);
    expect(value("RAND()*100", {}, { random: () => 0.25 })).toBe(25); // Excel's example: from 0 up to 100
  });

  // ceil(a) + floor(random()·(floor(b) − ceil(a) + 1)); #NUM! when ceil(a) > floor(b).
  it.each<[string, number, CellValue]>([
    ["RANDBETWEEN(1,100)", 0, 1],
    ["RANDBETWEEN(1,100)", 0.5, 51],
    ["RANDBETWEEN(1,100)", 0.999_999, 100],
    ["RANDBETWEEN(-1,1)", 0.5, 0],
    ["RANDBETWEEN(1.5,3.5)", 0.25, 2],
    ["RANDBETWEEN(1.5,3.5)", 0.75, 3],
    ["RANDBETWEEN(3,3)", 0.9, 3],
    ["RANDBETWEEN(5,1)", 0.5, err("#NUM!")],
    ["RANDBETWEEN(1.2,1.8)", 0.5, err("#NUM!")],
  ])("%s with random() = %d is %j", (formula, draw, expected) => {
    expect(value(formula, {}, { random: () => draw })).toEqual(expected);
  });

  it("draws once per call from a seeded ctx.random(), reproducibly", () => {
    const random = seeded(120);
    const twin = seeded(120);
    for (let i = 0; i < 100; i += 1) {
      expect(value("RAND()", {}, { random })).toBe(twin());
      expect(value("RANDBETWEEN(1,6)", {}, { random })).toBe(1 + Math.floor(twin() * 6));
    }
  });
});
