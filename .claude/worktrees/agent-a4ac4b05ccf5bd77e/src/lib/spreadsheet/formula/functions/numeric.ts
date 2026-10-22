// Numerical methods (SPEC-127 "Numerical methods") — framework-free, and
// tested alone against published tables (numeric.spec.ts).
//
// - newton: Newton's method with the stopping rules RATE (20 iterations,
//   |Δrate| < 1e-7), IRR (20, 1e-7) and XIRR (100, 1e-8) take from
//   Microsoft's pages, starting at `guess` (0.1). Exhausting the cap, a zero
//   or non-finite derivative, or a non-finite iterate or one at or below the
//   floor (a rate ≤ −1) fails — null, which the caller returns as #NUM! —
//   never a half-converged number.
// - erf / erfc: W. J. Cody's rational Chebyshev approximations (Math. Comp.
//   23, 1969), transcribed from netlib specfun's CALERF: |x| ≤ 0.46875 a
//   rational function of x² for erf; up to 4 one of |x| for erfc; beyond, one
//   of 1/x² for x·e^{x²}·erfc x. e^{−x²} is split as e^{−y²}·e^{−(x−y)(x+y)}
//   with y = x truncated to sixteenths, so erfc keeps its full relative
//   accuracy far into the tail (erfc 26 ≈ 5.7e−296).
// - gammaLn: Lanczos's approximation with g = 7 and the nine standard
//   coefficients, ln|Γ(x)| by the reflection formula below x = 0.5; gamma
//   evaluates the same series directly (Excel's GAMMA: a pole at 0 and each
//   negative integer, negative non-integers allowed).
// - incompleteBeta: the regularised I_x(a, b) by the continued fraction of
//   DLMF 8.17.22 (on the side of x = (a+1)/(a+b+2) where it converges, the
//   other through I_x(a,b) = 1 − I_{1−x}(b,a)), evaluated by the modified
//   Lentz method: at most 300 iterations, converged once a step changes the
//   value by a relative 1e-15; non-convergence is NaN (the caller's #NUM!).
// - Student's t: the CDF from the incomplete beta (A&S 26.7.1); the inverse
//   brackets the root (doubling, then squaring, the upper end, and
//   geometric bisection while the bracket spans more than a factor of two)
//   and then runs Newton on the CDF with the density as its derivative,
//   falling back to bisection whenever a step leaves the bracket — at most
//   100 iterations in all, converged once |Δx| < 1e-12 (relative to |x|
//   past 1); failure is NaN.
// - normal: normSDist through erfc; normSInv by Acklam's rational
//   approximation (relative error 1.15e−9) and one Halley step on erfc, the
//   upper half through the lower by symmetry (1 − p is exact there).
// - binomial and Poisson point probabilities in log space through gammaLn;
//   the cumulative forms add point probabilities — from the argument down
//   to where the terms vanish below the mode, else 1 minus the upper tail
//   summed the same way.

export interface NewtonOptions {
  /** The first iterate (the function's `guess`). */
  start: number;
  /** The most steps taken before giving up. */
  maxIterations: number;
  /** Converged once a step moves the iterate by less than this. */
  tolerance: number;
  /** An iterate at or below this fails (−1 for a rate). */
  floor?: number;
}

/** A root of `f` by Newton's method with its derivative `df` (see the module
 * header); null when the iteration fails. */
export function newton(f: (x: number) => number, df: (x: number) => number, options: NewtonOptions): number | null {
  let x = options.start;
  for (let i = 0; i < options.maxIterations; i += 1) {
    const value = f(x);
    const slope = df(x);
    if (!Number.isFinite(value) || !Number.isFinite(slope) || slope === 0) {
      return null;
    }
    const next = x - value / slope;
    if (!Number.isFinite(next) || (options.floor !== undefined && next <= options.floor)) {
      return null;
    }
    if (Math.abs(next - x) < options.tolerance) {
      return next;
    }
    x = next;
  }
  return null;
}

// --- erf and erfc (Cody, netlib specfun CALERF) --------------------------------------

// CALERF's DATA values (IEEE double precision), each written as the double
// nearest the published 18-digit constant.
/** 1/√π. */
const SQRPI = 0.5641895835477563;
/** Where the first interval ends. */
const THRESH = 0.46875;
/** Below this, x² is negligible beside 1. */
const XSMALL = 1.11e-16;
/** From here erfc underflows: it is 0. */
const XBIG = 26.543;
/** erf on |x| ≤ 0.46875. */
const ERF_A = [3.1611237438705655, 113.86415415105016, 377.485237685302, 3209.3775891384694, 0.18577770618460315];
const ERF_B = [23.601290952344122, 244.02463793444417, 1282.6165260773723, 2844.236833439171];
/** erfc on 0.46875 < |x| ≤ 4. */
const ERF_C = [
  0.5641884969886701, 8.883149794388377, 66.11919063714163, 298.6351381974001, 881.952221241769, 1712.0476126340707, 2051.0783778260716,
  1230.3393547979972, 2.1531153547440383e-8,
];
const ERF_D = [
  15.744926110709835, 117.6939508913125, 537.1811018620099, 1621.3895745666903, 3290.7992357334597, 4362.619090143247, 3439.3676741437216,
  1230.3393548037495,
];
/** x·e^{x²}·erfc x on |x| > 4, in 1/x². */
const ERF_P = [0.30532663496123236, 0.36034489994980445, 0.12578172611122926, 0.016083785148742275, 0.0006587491615298378, 0.016315387137302097];
const ERF_Q = [2.568520192289822, 1.8729528499234604, 0.5279051029514285, 0.06051834131244132, 0.0023352049762686918];

/** erf x for |x| ≤ THRESH (CALERF's first interval). */
function erfNearZero(x: number): number {
  const y = Math.abs(x);
  const ysq = y > XSMALL ? y * y : 0;
  let num = ERF_A[4]! * ysq;
  let den = ysq;
  for (let i = 0; i < 3; i += 1) {
    num = (num + ERF_A[i]!) * ysq;
    den = (den + ERF_B[i]!) * ysq;
  }
  return (x * (num + ERF_A[3]!)) / (den + ERF_B[3]!);
}

/** erfc y for y > THRESH (CALERF's second and third intervals). */
function erfcAbove(y: number): number {
  let ratio: number;
  if (y <= 4) {
    let num = ERF_C[8]! * y;
    let den = y;
    for (let i = 0; i < 7; i += 1) {
      num = (num + ERF_C[i]!) * y;
      den = (den + ERF_D[i]!) * y;
    }
    ratio = (num + ERF_C[7]!) / (den + ERF_D[7]!);
  } else {
    if (y >= XBIG) {
      return 0;
    }
    const z = 1 / (y * y);
    let num = ERF_P[5]! * z;
    let den = z;
    for (let i = 0; i < 4; i += 1) {
      num = (num + ERF_P[i]!) * z;
      den = (den + ERF_Q[i]!) * z;
    }
    ratio = (SQRPI - (z * (num + ERF_P[4]!)) / (den + ERF_Q[4]!)) / y;
  }
  // e^{−y²} as e^{−s²}·e^{−(y−s)(y+s)}, s = y truncated to sixteenths.
  const s = Math.trunc(y * 16) / 16;
  return Math.exp(-s * s) * Math.exp(-(y - s) * (y + s)) * ratio;
}

/** The error function, erf x = (2/√π)∫₀ˣ e^{−t²} dt. */
export function erf(x: number): number {
  if (Math.abs(x) <= THRESH) {
    return erfNearZero(x);
  }
  const result = 0.5 - erfcAbove(Math.abs(x)) + 0.5;
  return x < 0 ? -result : result;
}

/** The complementary error function, erfc x = 1 − erf x, to full relative
 * accuracy in the upper tail. */
export function erfc(x: number): number {
  if (Math.abs(x) <= THRESH) {
    return 1 - erfNearZero(x);
  }
  const tail = erfcAbove(Math.abs(x));
  return x < 0 ? 2 - tail : tail;
}

// --- the gamma function (Lanczos, g = 7) ---------------------------------------------

const LANCZOS_G = 7;
const LANCZOS = [
  0.9999999999998099, 676.5203681218851, -1259.1392167224028, 771.3234287776531, -176.6150291621406, 12.507343278686905,
  -0.13857109526572012, 9.984369578019572e-6, 1.5056327351493116e-7,
];
/** ln √(2π). */
const LN_SQRT_2PI = 0.9189385332046728;
const SQRT_2PI = 2.5066282746310007;

/** Lanczos's series A_g(x − 1) for x ≥ 0.5. */
function lanczosSum(x: number): number {
  const z = x - 1;
  let sum = LANCZOS[0]!;
  for (let i = 1; i < LANCZOS.length; i += 1) {
    sum += LANCZOS[i]! / (z + i);
  }
  return sum;
}

/** sin(πx), its argument reduced first so a large |x| keeps its accuracy. */
function sinPi(x: number): number {
  const r = x - 2 * Math.round(x / 2);
  return Math.sin(Math.PI * r);
}

/** Whether x is 0 or a negative integer — a pole of Γ. */
function isPole(x: number): boolean {
  return x <= 0 && Number.isInteger(x);
}

/** ln |Γ(x)|: Lanczos for x ≥ 0.5, the reflection formula
 * Γ(x)Γ(1−x) = π / sin(πx) below; +∞ at a pole, NaN for NaN. */
export function gammaLn(x: number): number {
  if (Number.isNaN(x)) {
    return Number.NaN;
  }
  if (isPole(x)) {
    return Number.POSITIVE_INFINITY;
  }
  if (x < 0.5) {
    return Math.log(Math.PI / Math.abs(sinPi(x))) - gammaLn(1 - x);
  }
  const t = x - 1 + LANCZOS_G + 0.5;
  return LN_SQRT_2PI + (x - 0.5) * Math.log(t) - t + Math.log(lanczosSum(x));
}

/** Γ(x) (Excel's GAMMA): NaN at a pole (0, −1, −2, …), ±∞ past the double
 * range, negative non-integers through the reflection formula. */
export function gamma(x: number): number {
  if (Number.isNaN(x) || isPole(x)) {
    return Number.NaN;
  }
  if (x < 0.5) {
    return Math.PI / (sinPi(x) * gamma(1 - x));
  }
  const t = x - 1 + LANCZOS_G + 0.5;
  // t^(x−½)·e^{−t}, the power split in two so it overflows only with Γ.
  const half = Math.pow(t, (x - 0.5) / 2);
  return SQRT_2PI * half * (half * Math.exp(-t)) * lanczosSum(x);
}

// --- the regularised incomplete beta (DLMF 8.17.22, modified Lentz) ---------------------

/** Lentz's guard against a zero denominator. */
const LENTZ_TINY = 1e-300;
/** The incomplete beta's iteration cap and convergence threshold. */
export const BETA_MAX_ITERATIONS = 300;
const BETA_EPSILON = 1e-15;

/** The continued fraction 1/(1+ d₁/(1+ d₂/(1+ …))) of DLMF 8.17.22–23 by the
 * modified Lentz method; NaN when it has not converged in `maxIterations`
 * (each m taking the even and the odd term). */
function betaContinuedFraction(x: number, a: number, b: number, maxIterations: number): number {
  let c = 1;
  let d = 1 - ((a + b) * x) / (a + 1); // 1 + d₁
  if (Math.abs(d) < LENTZ_TINY) {
    d = LENTZ_TINY;
  }
  d = 1 / d;
  let h = d;
  for (let m = 1; m <= maxIterations; m += 1) {
    const m2 = 2 * m;
    // d_{2m}
    let term = (m * (b - m) * x) / ((a + m2 - 1) * (a + m2));
    d = 1 + term * d;
    if (Math.abs(d) < LENTZ_TINY) {
      d = LENTZ_TINY;
    }
    c = 1 + term / c;
    if (Math.abs(c) < LENTZ_TINY) {
      c = LENTZ_TINY;
    }
    d = 1 / d;
    h *= d * c;
    // d_{2m+1}
    term = (-(a + m) * (a + b + m) * x) / ((a + m2) * (a + m2 + 1));
    d = 1 + term * d;
    if (Math.abs(d) < LENTZ_TINY) {
      d = LENTZ_TINY;
    }
    c = 1 + term / c;
    if (Math.abs(c) < LENTZ_TINY) {
      c = LENTZ_TINY;
    }
    d = 1 / d;
    const delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1) < BETA_EPSILON) {
      return h;
    }
  }
  return Number.NaN;
}

/** I_x(a, b) given x and y = 1 − x separately (each accurate on its own,
 * as Student's t supplies them); see the module header. */
function betaRegularized(x: number, y: number, a: number, b: number, maxIterations: number): number {
  if (!(a > 0 && b > 0 && x >= 0 && y >= 0)) {
    return Number.NaN;
  }
  if (x === 0) {
    return 0;
  }
  if (y === 0) {
    return 1;
  }
  const lnFront = a * Math.log(x) + b * Math.log(y) - (gammaLn(a) + gammaLn(b) - gammaLn(a + b));
  if (x < (a + 1) / (a + b + 2)) {
    return (Math.exp(lnFront) * betaContinuedFraction(x, a, b, maxIterations)) / a;
  }
  return 1 - (Math.exp(lnFront) * betaContinuedFraction(y, b, a, maxIterations)) / b;
}

/** The regularised incomplete beta function I_x(a, b) = B_x(a,b)/B(a,b)
 * (DLMF 8.17.2) for 0 ≤ x ≤ 1 and a, b > 0; NaN outside that or when the
 * continued fraction has not converged in `maxIterations` (300). */
export function incompleteBeta(x: number, a: number, b: number, maxIterations = BETA_MAX_ITERATIONS): number {
  return x <= 1 ? betaRegularized(x, 1 - x, a, b, maxIterations) : Number.NaN;
}

// --- Student's t -----------------------------------------------------------------------

/** P(T > t) for Student's t with `df` degrees of freedom: half of
 * I_{ν/(ν+t²)}(ν/2, ½) (A&S 26.7.1), reflected for t < 0; NaN when df ≤ 0
 * or the incomplete beta does not converge. */
export function studentTUpper(t: number, df: number, maxIterations = BETA_MAX_ITERATIONS): number {
  if (!(df > 0) || Number.isNaN(t)) {
    return Number.NaN;
  }
  // x = ν/(ν+t²) and 1 − x = t²/(ν+t²), each computed directly from the
  // smaller of u² and 1/u² (u = |t|/√ν), so neither cancels nor overflows.
  const u = Math.abs(t) / Math.sqrt(df);
  let x: number;
  let y: number;
  if (u > 1) {
    const s = 1 / (u * u);
    x = s / (1 + s);
    y = 1 / (1 + s);
  } else {
    const r = u * u;
    x = 1 / (1 + r);
    y = r / (1 + r);
  }
  const half = betaRegularized(x, y, df / 2, 0.5, maxIterations) / 2;
  return t >= 0 ? half : 1 - half;
}

/** P(T ≤ t), the left-tailed CDF. */
export function studentTCdf(t: number, df: number, maxIterations = BETA_MAX_ITERATIONS): number {
  return studentTUpper(-t, df, maxIterations);
}

/** Student's t density, Γ((ν+1)/2) / (√(νπ)Γ(ν/2)) · (1 + t²/ν)^(−(ν+1)/2). */
export function studentTPdf(t: number, df: number): number {
  if (!(df > 0)) {
    return Number.NaN;
  }
  const lnScale = gammaLn((df + 1) / 2) - gammaLn(df / 2) - 0.5 * Math.log(df * Math.PI);
  return Math.exp(lnScale - ((df + 1) / 2) * Math.log1p((t * t) / df));
}

/** The t inverse's iteration cap and step tolerance. */
export const T_INV_MAX_ITERATIONS = 100;
const T_INV_TOLERANCE = 1e-12;

/** The t ≥ 0 with P(T > t) = q, for 0 < q ≤ ½ (see the module header);
 * NaN when the iteration fails. */
function studentTUpperInverse(q: number, df: number, maxIterations: number): number {
  let iterations = 0;
  let failed = false;
  /** P(T > t) − q, one iteration each. */
  const g = (t: number): number => {
    iterations += 1;
    const value = studentTUpper(t, df) - q;
    if (Number.isNaN(value)) {
      failed = true;
    }
    return value;
  };
  // Bracket: g(lo) > 0 ≥ g(hi).
  let lo = 0;
  let hi = 1;
  while (g(hi) > 0) {
    if (failed || iterations >= maxIterations || hi === Number.MAX_VALUE) {
      return Number.NaN;
    }
    lo = hi;
    hi = hi < 16 ? hi * 2 : Math.min(hi * hi, Number.MAX_VALUE);
  }
  // Geometric bisection while the bracket spans more than a factor of two.
  while (lo > 0 && hi > 2 * lo) {
    if (failed || iterations >= maxIterations) {
      return Number.NaN;
    }
    const mid = Math.sqrt(lo) * Math.sqrt(hi);
    if (g(mid) > 0) {
      lo = mid;
    } else {
      hi = mid;
    }
  }
  // Newton on g (g′ = −density), bisecting whenever a step leaves the bracket.
  let x = (lo + hi) / 2;
  for (;;) {
    if (iterations >= maxIterations) {
      return Number.NaN;
    }
    const value = g(x);
    if (failed) {
      return Number.NaN;
    }
    if (value > 0) {
      lo = x;
    } else {
      hi = x;
    }
    const slope = -studentTPdf(x, df);
    let next = slope < 0 && Number.isFinite(slope) ? x - value / slope : Number.NaN;
    if (!(next > lo && next < hi)) {
      next = (lo + hi) / 2;
    }
    if (Math.abs(next - x) < T_INV_TOLERANCE * Math.max(1, Math.abs(x))) {
      // A bracket collapsed where the tail no longer resolves q (|t| past
      // the double range) is a failure, not a root.
      return Math.abs(value) <= q * 1e-3 ? next : Number.NaN;
    }
    x = next;
  }
}

/** The left-tailed inverse: the t with P(T ≤ t) = p, for 0 < p < 1;
 * NaN otherwise or when the iteration fails (at most `maxIterations`). */
export function studentTInv(p: number, df: number, maxIterations = T_INV_MAX_ITERATIONS): number {
  if (!(p > 0 && p < 1) || !(df > 0)) {
    return Number.NaN;
  }
  if (p === 0.5) {
    return 0;
  }
  // 1 − p is exact for p ≥ ½.
  return p < 0.5 ? -studentTUpperInverse(p, df, maxIterations) : studentTUpperInverse(1 - p, df, maxIterations);
}

// --- the normal distribution --------------------------------------------------------------

/** The standard normal CDF Φ(z) = ½·erfc(−z/√2). */
export function normSDist(z: number): number {
  return 0.5 * erfc(-z / Math.SQRT2);
}

/** The standard normal density e^{−z²/2}/√(2π). */
export function normSPdf(z: number): number {
  return Math.exp(-0.5 * z * z) / SQRT_2PI;
}

const ACKLAM_A = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239];
const ACKLAM_B = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
const ACKLAM_C = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
const ACKLAM_D = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
const ACKLAM_LOW = 0.02425;

/** Acklam's rational approximation to Φ⁻¹(p) for 0 < p ≤ ½. */
function acklamLowerHalf(p: number): number {
  if (p < ACKLAM_LOW) {
    const q = Math.sqrt(-2 * Math.log(p));
    const [c1, c2, c3, c4, c5, c6] = ACKLAM_C as [number, number, number, number, number, number];
    const [d1, d2, d3, d4] = ACKLAM_D as [number, number, number, number];
    return (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / ((((d1 * q + d2) * q + d3) * q + d4) * q + 1);
  }
  const q = p - 0.5;
  const r = q * q;
  const [a1, a2, a3, a4, a5, a6] = ACKLAM_A as [number, number, number, number, number, number];
  const [b1, b2, b3, b4, b5] = ACKLAM_B as [number, number, number, number, number];
  return ((((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q) / (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1);
}

/** Φ⁻¹(p) for 0 < p ≤ ½: Acklam's approximation, then one Halley step. */
function normSInvLowerHalf(p: number): number {
  const x = acklamLowerHalf(p);
  const e = 0.5 * erfc(-x / Math.SQRT2) - p;
  const u = e * SQRT_2PI * Math.exp((x * x) / 2);
  const refined = x - u / (1 + (x * u) / 2);
  return Number.isFinite(refined) ? refined : x;
}

/** The standard normal inverse Φ⁻¹(p) for 0 < p < 1 (NaN otherwise). */
export function normSInv(p: number): number {
  if (!(p > 0 && p < 1)) {
    return Number.NaN;
  }
  return p <= 0.5 ? normSInvLowerHalf(p) : -normSInvLowerHalf(1 - p);
}

// --- the binomial and Poisson distributions ------------------------------------------------

/** A negligible term beside the running sum. */
const SUM_EPSILON = 1e-17;

/** Σ term(i) for i from `from` stepping by `step` while `inRange`, the terms
 * shrinking monotonically: stops once a term is negligible. */
function tailSum(from: number, step: 1 | -1, inRange: (i: number) => boolean, term: (i: number) => number): number {
  let total = 0;
  for (let i = from; inRange(i); i += step) {
    const value = term(i);
    total += value;
    if (value <= total * SUM_EPSILON) {
      break;
    }
  }
  return total;
}

/** P(X = k) for X ~ Binomial(n, p): C(n,k)·p^k·(1−p)^(n−k) in log space
 * (k, n integers, 0 ≤ k ≤ n, 0 ≤ p ≤ 1). */
export function binomialPmf(k: number, n: number, p: number): number {
  if (p === 0 || p === 1) {
    return (p === 0 ? k === 0 : k === n) ? 1 : 0;
  }
  const lnChoose = gammaLn(n + 1) - gammaLn(k + 1) - gammaLn(n - k + 1);
  return Math.exp(lnChoose + k * Math.log(p) + (n - k) * Math.log1p(-p));
}

/** P(X ≤ k) for X ~ Binomial(n, p), as a sum of point probabilities. */
export function binomialCdf(k: number, n: number, p: number): number {
  if (k >= n) {
    return 1;
  }
  const mode = Math.floor((n + 1) * p);
  if (k < mode) {
    return Math.min(1, tailSum(k, -1, (i) => i >= 0, (i) => binomialPmf(i, n, p)));
  }
  return Math.max(0, 1 - tailSum(k + 1, 1, (i) => i <= n, (i) => binomialPmf(i, n, p)));
}

/** P(X = k) for X ~ Poisson(mean): e^{−λ}λ^k/k! in log space (k a
 * non-negative integer, mean ≥ 0). */
export function poissonPmf(k: number, mean: number): number {
  if (mean === 0) {
    return k === 0 ? 1 : 0;
  }
  return Math.exp(k * Math.log(mean) - mean - gammaLn(k + 1));
}

/** P(X ≤ k) for X ~ Poisson(mean), as a sum of point probabilities. */
export function poissonCdf(k: number, mean: number): number {
  const mode = Math.floor(mean);
  if (k < mode) {
    return Math.min(1, tailSum(k, -1, (i) => i >= 0, (i) => poissonPmf(i, mean)));
  }
  return Math.max(0, 1 - tailSum(k + 1, 1, () => true, (i) => poissonPmf(i, mean)));
}
