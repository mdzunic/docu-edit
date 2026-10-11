// SPEC-127 "Numerical methods" — functions/numeric.ts alone, against
// published tables, each test citing its source: Abramowitz & Stegun
// (A&S, 10th printing) Tables 7.1, 7.3, 26.1, 26.6 and 26.10 and formulas
// 26.7.1/26.7.4; NIST DLMF 5.4.6–5.4.10 and Table 5.4.1 (Γ) and 8.17.5
// (I_x for integer parameters). A value is compared to the table's printed
// precision: half a unit in its last place (1e-13 relative or better where
// the table prints that many digits). And each iteration cap is tripped by
// a constructed non-convergent input.
import { describe, expect, it } from "vitest";
import { erf, erfc, gamma, gammaLn, incompleteBeta, newton, normSDist, normSInv, studentTCdf, studentTInv } from "./numeric";

/** |actual − printed| within half a unit of the printed last decimal. */
function matchesTable(actual: number, printed: number, decimals: number): void {
  expect(Math.abs(actual - printed)).toBeLessThanOrEqual(0.5 * 10 ** -decimals * (1 + 1e-9));
}

/** |actual − expected| ≤ tolerance·|expected|. */
function relativelyClose(actual: number, expected: number, tolerance: number): void {
  expect(Math.abs(actual - expected)).toBeLessThanOrEqual(tolerance * Math.abs(expected));
}

describe("erf and erfc (Cody's CALERF)", () => {
  it.each<[number, number]>([
    [0.1, 0.1124629160],
    [0.3, 0.3286267595],
    [0.46, 0.4846553900],
    [0.5, 0.5204998778],
    [0.8, 0.7421009647],
    [1.0, 0.8427007929],
    [1.5, 0.9661051465],
    [2.0, 0.9953222650],
  ])("A&S Table 7.1: erf %d to 10 decimals", (x, printed) => {
    matchesTable(erf(x), printed, 10);
    matchesTable(erf(-x), -printed, 10);
    matchesTable(erfc(x), 1 - printed, 10);
  });

  // The table prints 15 decimals; its last one is compared to within a few
  // units (14 decimals), the leading value to 1e-13 relative below.
  it.each<[number, number, number]>([
    [1, 0.012188882184803, 14],
    [2, 0.000392750588282, 14],
    [3, 0.000014144402689, 14],
    [5, 0.000000020826552, 14],
    [6, 0.000000000825422, 14],
  ])("A&S Table 7.3 (Emersleben): erfc √(%iπ) to its printed decimals", (n, printed, decimals) => {
    matchesTable(erfc(Math.sqrt(n * Math.PI)), printed, decimals);
  });

  it("keeps erfc(√π) to better than 1e-13 relative (A&S Table 7.3)", () => {
    relativelyClose(erfc(Math.sqrt(Math.PI)), 0.012188882184803, 1e-13);
  });

  it.each<[number, number]>([
    [0.25, 0.5107914],
    [0.1, 0.5394141],
    [0.01, 0.5614099],
    [0.005, 0.5627896],
  ])("A&S Table 7.3: x·e^{x²}·erfc x at x⁻² = %d to 7 decimals (the far tail)", (inverseSquare, printed) => {
    const x = 1 / Math.sqrt(inverseSquare);
    matchesTable(x * Math.exp(x * x) * erfc(x), printed, 7);
  });

  it("is odd, with erfc(−x) = 2 − erfc(x), 0 and 1 at the ends", () => {
    expect(erf(0)).toBe(0);
    expect(erf(30)).toBe(1);
    expect(erfc(30)).toBe(0);
    expect(erfc(-30)).toBe(2);
  });
});

describe("gammaLn and gamma (Lanczos, g = 7)", () => {
  it.each<[string, number, number]>([
    ["5.4.6 Γ(1/2)", 1 / 2, Number("1.7724538509055160273")],
    ["5.4.7 Γ(1/3)", 1 / 3, Number("2.6789385347077476337")],
    ["5.4.8 Γ(2/3)", 2 / 3, Number("1.3541179394264004169")],
    ["5.4.9 Γ(1/4)", 1 / 4, Number("3.6256099082219083119")],
    ["5.4.10 Γ(3/4)", 3 / 4, Number("1.2254167024651776451")],
    ["Table 5.4.1 Γ(x₀)", Number("1.4616321449683623413"), Number("0.88560319441088870028")],
  ])("DLMF %s", (_label, x, value) => {
    relativelyClose(gamma(x), value, 1e-13);
    relativelyClose(gammaLn(x), Math.log(value), 1e-13);
  });

  it("gives Γ(n+1) = n! (DLMF 5.4.1) and its logarithm", () => {
    relativelyClose(gamma(6), 120, 1e-14);
    relativelyClose(gammaLn(101), 363.73937555556347, 1e-13); // ln 100! (DLMF 5.4.1)
  });

  it("reflects below ½ and has poles at 0 and the negative integers", () => {
    relativelyClose(gamma(-0.5), -2 * Number("1.7724538509055160273"), 1e-13); // Γ(−½) = −2√π (DLMF 5.5.1 with 5.4.6)
    expect(gamma(0)).toBeNaN();
    expect(gamma(-2)).toBeNaN();
    expect(gammaLn(-1)).toBe(Number.POSITIVE_INFINITY);
  });
});

describe("the regularised incomplete beta (DLMF 8.17.22, modified Lentz)", () => {
  it.each<[number, number, number, number]>([
    // DLMF 8.17.5: I_x(m, n−m+1) = Σ_{j=m}^{n} C(n,j) x^j (1−x)^{n−j}, exact decimals.
    [0.3, 2, 3, 0.3483], // n = 4: 6·0.09·0.49 + 4·0.027·0.7 + 0.0081
    [0.5, 3, 3, 0.5], // n = 5: (10 + 5 + 1)/32
    [0.2, 1, 4, 0.5904], // 1 − 0.8⁴
    [0.9, 5, 1, 0.59049], // 0.9⁵
    [0.6, 2, 2, 0.648], // n = 3: 3·0.36·0.4 + 0.216
  ])("DLMF 8.17.5: I_%d(%i, %i) = %d", (x, a, b, expected) => {
    relativelyClose(incompleteBeta(x, a, b), expected, 1e-13);
  });

  it("gives the t-distribution's closed form for ν = 2 (A&S 26.7.1 with 26.7.4)", () => {
    // A(t|2) = sin θ = t/√(t²+2), so P(T ≤ 1) = ½ + 1/(2√3).
    relativelyClose(studentTCdf(1, 2), 0.5 + 1 / (2 * Math.sqrt(3)), 1e-13);
  });

  it("is NaN once its 300-iteration cap is exhausted", () => {
    // a = b = 10⁸ at x = ½ needs thousands of terms.
    expect(incompleteBeta(0.5, 1e8, 1e8)).toBeNaN();
    expect(incompleteBeta(0.5, 10, 10, 3)).toBeNaN();
    expect(incompleteBeta(0.5, 10, 10)).toBeCloseTo(0.5, 14); // I_{1/2}(a,a) = ½ by DLMF 8.17.4
  });
});

describe("the standard normal and its inverse", () => {
  it.each<[number, number]>([
    [0.1, 0.539827837277029],
    [0.5, 0.691462461274013],
    [0.8, 0.788144601416604],
    [1.0, 0.841344746068543],
  ])("A&S Table 26.1: P(%d) (15 decimals) to 1e-13 relative, and back through the inverse", (x, printed) => {
    relativelyClose(normSDist(x), printed, 1e-13);
    relativelyClose(normSInv(printed), x, 1e-13);
    relativelyClose(normSInv(1 - printed), -x, 1e-13);
  });

  it.each<[number, number]>([
    [1e-4, 3.71902],
    [1e-9, 5.99781],
    [1e-14, 7.65063],
    [1e-23, 9.97305],
  ])("A&S Table 26.6: Q(x) = %d at x = %d (5 decimals)", (q, printed) => {
    matchesTable(-normSInv(q), printed, 5);
  });
});

describe("Student's t inverse", () => {
  it.each<[number, number, number]>([
    [1, 0.95, 12.706],
    [10, 0.95, 2.228],
    [30, 0.99, 2.750],
    [1, 0.999999, 636619.772],
  ])("A&S Table 26.10: ν = %i, A = %d gives t = %d (3 decimals)", (nu, a, printed) => {
    matchesTable(studentTInv(1 - (1 - a) / 2, nu), printed, 3);
  });

  it("is NaN once its 100-iteration cap is exhausted", () => {
    expect(studentTInv(0.975, 10, 3)).toBeNaN();
    expect(studentTInv(0.975, 10)).not.toBeNaN();
  });
});

// f(x) = x² + 1 has no real root: Newton never converges.
const f = (x: number): number => x * x + 1;
const df = (x: number): number => 2 * x;

describe("newton's caps", () => {
  it("fails after RATE/IRR's 20 iterations at 1e-7", () => {
    expect(newton(f, df, { start: 0.1, maxIterations: 20, tolerance: 1e-7 })).toBeNull();
  });

  it("fails after XIRR's 100 iterations at 1e-8", () => {
    expect(newton(f, df, { start: 0.1, maxIterations: 100, tolerance: 1e-8 })).toBeNull();
  });

  it("converges on a root inside the cap", () => {
    expect(newton((x) => x * x - 2, (x) => 2 * x, { start: 1, maxIterations: 20, tolerance: 1e-7 })).toBeCloseTo(Math.SQRT2, 12);
  });
});
