// Statistical (SPEC-127 "Statistical", "Function pins"): MEDIAN MODE.SNGL
// STDEV.S STDEV.P VAR.S VAR.P LARGE SMALL RANK.EQ RANK.AVG PERCENTILE.INC
// PERCENTILE.EXC QUARTILE.INC QUARTILE.EXC PERCENTRANK.INC PERCENTRANK.EXC
// CORREL PEARSON RSQ COVARIANCE.P COVARIANCE.S SLOPE INTERCEPT STEYX
// FORECAST.LINEAR NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV STANDARDIZE
// T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T BINOM.DIST POISSON.DIST
// EXPON.DIST CONFIDENCE.NORM AVERAGEA MAXA MINA STDEVA STDEVPA VARA VARPA
// GEOMEAN HARMEAN AVEDEV DEVSQ KURT SKEW SKEW.P TRIMMEAN FREQUENCY GAMMA
// GAMMALN GAMMALN.PRECISE, and under "Compatibility" the legacy names, each
// sharing its modern function's `impl` (TDIST picks the T.DIST.RT or
// T.DIST.2T kernel by `tails`; NORMSDIST, without `cumulative`, is the CDF).
//
// - Data sets read as SPEC-120's aggregates do (common.ts eachNumber):
//   numbers only from references and arrays, direct arguments coerced, the
//   first error propagating. The -A variants (AVERAGEA … VARPA) also count
//   TRUE as 1 and FALSE as 0 everywhere, and text in references and arrays
//   as 0 (empty cells skipped; a direct text that is no number is #VALUE!).
// - Variance and standard deviation are two-pass (stats.ts), as are the
//   paired sums of squares and products.
// - The paired functions (CORREL, PEARSON, RSQ, COVARIANCE.*, SLOPE,
//   INTERCEPT, STEYX, FORECAST.LINEAR and their legacy names) take their
//   data as Excel's array arguments (evaluated in array mode), need the two
//   arrays to hold the same number of positions (#N/A otherwise) and use
//   only the positions where both hold numbers. With no such position
//   CORREL and the covariances are #DIV/0! while the rest are #N/A; too few
//   points or zero variance are #DIV/0!, each as Microsoft's page states.
// - MODE.SNGL is the first value, in argument order, with the highest
//   count (#N/A when none repeats). RANK.EQ/.AVG need a reference, compare
//   at 15 significant digits and are #N/A when the number is absent.
// - PERCENTRANK.INC ranks a value v of the data as (values below v)/(n−1)
//   and .EXC as (values below v + 1)/(n+1); between two data values the
//   rank is interpolated linearly between theirs; outside the data #N/A;
//   the result truncated (never rounded) to `significance` digits, default
//   3 (< 1 is #NUM!).
// - TRIMMEAN drops ⌊n·percent/2⌋ values at each end; FREQUENCY counts the
//   data at or below each bin and above the one before it (bins in sorted
//   order, results in the order given), plus those above the last bin, as a
//   vertical array — the evaluator distributes it over an array formula's
//   range and a plain cell shows its first count.
// - The distributions come from numeric.ts (erf/erfc, Acklam + Halley,
//   Lentz's incomplete beta, Lanczos's gamma, log-space binomial and
//   Poisson terms); degrees of freedom, BINOM.DIST's number_s and trials,
//   POISSON.DIST's x and CONFIDENCE.NORM's size are truncated to integers.
//   A computation that fails to converge or leaves the double range is
//   #NUM!.
import type { CellError } from "../../model";
import { areaCols, areaSize, forEachScalar, forEachStoredCell, makeArray, toArray } from "../areas";
import type { EvalContext } from "../evaluate";
import type { ArgKind, ArgList, FunctionCategory, FunctionSpec } from "../registry";
import { DIV0_ERROR, isError, isRange, NA_ERROR, NUM_ERROR, roundTo15, toNumber, VALUE_ERROR, type EvalValue } from "../values";
import { arg, eachNumber, logicalArg, numberArg, numberOr, optional, repeating } from "./common";
import {
  binomialCdf,
  binomialPmf,
  gamma,
  gammaLn,
  normSDist,
  normSInv,
  normSPdf,
  poissonCdf,
  poissonPmf,
  studentTInv,
  studentTPdf,
  studentTUpper,
  studentTCdf,
} from "./numeric";
import {
  ascending,
  kth,
  mean,
  median,
  modeSingle,
  percentileExc,
  percentileInc,
  quartileExc,
  quartileInc,
  squaredDeviations,
  stdev,
  sum,
  variance,
} from "./stats";

type Impl = FunctionSpec["impl"];
type ArgMeta = FunctionSpec["args"][number];

// --- reading arguments -------------------------------------------------------------

/** Arguments `from` up to `to` of a call, as a list of their own. */
function slice(args: ArgList, from: number, to: number = args.length): ArgList {
  return {
    length: Math.max(0, Math.min(to, args.length) - from),
    get: (index) => args.get(from + index),
    isMissing: (index) => args.isMissing(from + index),
  };
}

/** Every number the arguments hold, read as the aggregates read them
 * (eachNumber), or the first error. */
function numbersOf(args: ArgList, ctx: EvalContext): number[] | CellError {
  const xs: number[] = [];
  const failure = eachNumber(args, ctx, (n) => {
    xs.push(n);
  });
  return failure ?? xs;
}

/** The -A variants' reading (see the module header), or the first error. */
function valuesA(args: ArgList, ctx: EvalContext): number[] | CellError {
  const xs: number[] = [];
  let failure: CellError | undefined;
  for (let i = 0; i < args.length; i += 1) {
    if (failure) {
      break;
    }
    forEachScalar(ctx, args.get(i), (value, direct) => {
      if (isError(value)) {
        failure = value;
        return false;
      }
      if (typeof value === "number") {
        xs.push(value);
      } else if (typeof value === "boolean") {
        xs.push(value ? 1 : 0);
      } else if (direct) {
        // Typed directly: numeric text coerces, other text is #VALUE!, and
        // an empty slot reads as 0.
        const n = toNumber(value, ctx.settings);
        if (isError(n)) {
          failure = n;
          return false;
        }
        xs.push(n);
      } else if (typeof value === "string") {
        xs.push(0);
      }
      return true;
    });
  }
  return failure ?? xs;
}

/** The first `count` arguments as numbers (value arguments), or the first
 * error. */
function numberArgs(args: ArgList, ctx: EvalContext, count: number): number[] | CellError {
  const values: number[] = [];
  for (let i = 0; i < count; i += 1) {
    const n = numberArg(args, i, ctx);
    if (isError(n)) {
      return n;
    }
    values.push(n);
  }
  return values;
}

/** `f` over a list read without error; the error otherwise. */
function over(xs: number[] | CellError, f: (xs: number[]) => number | CellError): number | CellError {
  return isError(xs) ? xs : f(xs);
}

/** A number, or #NUM! for NaN and ±∞ (a failed or overflowing method). */
function finite(x: number): number | CellError {
  return Number.isFinite(x) ? x : NUM_ERROR;
}

// --- spec builders -----------------------------------------------------------------

const NUMBER1 = arg("number1", "A number, or a range or array of numbers.");
const NUMBER2 = repeating("number2", "More numbers, ranges or arrays.");
const VALUE1 = arg("value1", "A value, range or array; text counts as 0 and TRUE as 1.");
const VALUE2 = repeating("value2", "More values, ranges or arrays.");

/** F(number1, [number2], …) over every argument. */
function variadic(name: string, category: FunctionCategory, summary: string, impl: Impl, first = NUMBER1, rest = NUMBER2): FunctionSpec {
  return {
    name,
    category,
    minArgs: 1,
    maxArgs: 255,
    argKinds: ["any", "any"],
    repeat: { from: 1, count: 1 },
    summary,
    args: [first, rest],
    impl,
  };
}

/** A fixed-arity function: `params` in order, one kind each. */
function fixed(name: string, category: FunctionCategory, summary: string, params: ArgMeta[], argKinds: ArgKind[], impl: Impl): FunctionSpec {
  return {
    name,
    category,
    minArgs: params.filter((param) => param.optional !== true).length,
    maxArgs: params.length,
    argKinds,
    summary,
    args: params,
    impl,
  };
}

// --- data-set statistics -------------------------------------------------------------

const medianImpl: Impl = (args, ctx) => over(numbersOf(args, ctx), median);
const modeImpl: Impl = (args, ctx) => over(numbersOf(args, ctx), modeSingle);
const stdevSampleImpl: Impl = (args, ctx) => over(numbersOf(args, ctx), (xs) => stdev(xs, true));
const stdevPopulationImpl: Impl = (args, ctx) => over(numbersOf(args, ctx), (xs) => stdev(xs, false));
const varSampleImpl: Impl = (args, ctx) => over(numbersOf(args, ctx), (xs) => variance(xs, true));
const varPopulationImpl: Impl = (args, ctx) => over(numbersOf(args, ctx), (xs) => variance(xs, false));

/** F(array, k): the numbers of the first argument and a number. */
function arrayAndNumber(kernel: (xs: number[], k: number) => number | CellError): Impl {
  return (args, ctx) => {
    const xs = numbersOf(slice(args, 0, 1), ctx);
    if (isError(xs)) {
      return xs;
    }
    const k = numberArg(args, 1, ctx);
    return isError(k) ? k : kernel(xs, k);
  };
}

const percentileIncImpl = arrayAndNumber(percentileInc);
const quartileIncImpl = arrayAndNumber(quartileInc);

/** RANK.EQ (`average` false) and RANK.AVG: the number's place among the
 * reference's numbers, descending unless `order` is non-zero. */
function rankImpl(average: boolean): Impl {
  return (args, ctx) => {
    const number = numberArg(args, 0, ctx);
    if (isError(number)) {
      return number;
    }
    const xs = numbersOf(slice(args, 1, 2), ctx);
    if (isError(xs)) {
      return xs;
    }
    const order = numberOr(args, 2, ctx, 0);
    if (isError(order)) {
      return order;
    }
    const target = roundTo15(number);
    let before = 0;
    let equal = 0;
    for (const x of xs) {
      const value = roundTo15(x);
      if (value === target) {
        equal += 1;
      } else if (order === 0 ? value > target : value < target) {
        before += 1;
      }
    }
    if (equal === 0) {
      return NA_ERROR;
    }
    return average ? before + (equal + 1) / 2 : before + 1;
  };
}

const rankEqImpl = rankImpl(false);

/** x truncated (toward zero) to `digits` decimals, through its 15-digit
 * form so 0.7 stays 0.7. */
function truncateDigits(x: number, digits: number): number {
  const scale = 10 ** digits;
  return Math.floor(roundTo15(x * scale)) / scale;
}

/** PERCENTRANK.INC (`exclusive` false) and .EXC (see the module header). */
function percentRankImpl(exclusive: boolean): Impl {
  return (args, ctx) => {
    const xs = numbersOf(slice(args, 0, 1), ctx);
    if (isError(xs)) {
      return xs;
    }
    const x = numberArg(args, 1, ctx);
    if (isError(x)) {
      return x;
    }
    const significance = numberOr(args, 2, ctx, 3);
    if (isError(significance)) {
      return significance;
    }
    const digits = Math.trunc(significance);
    if (xs.length === 0 || digits < 1) {
      return NUM_ERROR;
    }
    const sorted = ascending(xs);
    const n = sorted.length;
    if (x < sorted[0]! || x > sorted[n - 1]!) {
      return NA_ERROR;
    }
    /** How many values lie below v. */
    const below = (v: number): number => {
      let lo = 0;
      let hi = n;
      while (lo < hi) {
        const mid = (lo + hi) >>> 1;
        if (sorted[mid]! < v) {
          lo = mid + 1;
        } else {
          hi = mid;
        }
      }
      return lo;
    };
    const rank = (count: number): number => (exclusive ? (count + 1) / (n + 1) : n === 1 ? 1 : count / (n - 1));
    const at = below(x);
    if (sorted[at] === x) {
      return truncateDigits(rank(at), digits);
    }
    const low = sorted[at - 1]!;
    const high = sorted[at]!;
    const lowRank = rank(below(low));
    return truncateDigits(lowRank + ((x - low) / (high - low)) * (rank(at) - lowRank), digits);
  };
}

const percentRankIncImpl = percentRankImpl(false);

// --- paired data -----------------------------------------------------------------------

interface PositionNumbers {
  /** The positions the argument holds. */
  size: number;
  /** Its numbers by row-major position. */
  numbers: Map<number, number>;
}

/** A data array's size and numbers by position (stored cells only for a
 * reference), or the first error; several areas are #VALUE!. */
function positionNumbers(ctx: EvalContext, value: EvalValue): PositionNumbers | CellError {
  const numbers = new Map<number, number>();
  let failure: CellError | undefined;
  if (isRange(value)) {
    if (value.areas.length !== 1) {
      return VALUE_ERROR;
    }
    const area = value.areas[0]!;
    const cols = areaCols(area);
    forEachStoredCell(ctx, area, (cell, row, col) => {
      if (isError(cell)) {
        failure = cell;
        return false;
      }
      if (typeof cell === "number") {
        numbers.set((row - area.top) * cols + (col - area.left), cell);
      }
      return true;
    });
    return failure ?? { size: areaSize(area), numbers };
  }
  const array = toArray(ctx, value);
  if (isError(array)) {
    return array;
  }
  array.data.forEach((element, position) => {
    if (failure === undefined && isError(element)) {
      failure = element;
    } else if (typeof element === "number") {
      numbers.set(position, element);
    }
  });
  return failure ?? { size: array.data.length, numbers };
}

interface PairedMoments {
  n: number;
  /** The means of the first and second argument's numbers. */
  mean1: number;
  mean2: number;
  /** Two-pass sums of squared deviations and of their products. */
  s11: number;
  s22: number;
  s12: number;
}

/** The two arguments' paired positions (see the module header) summed. */
function pairedMoments(args: ArgList, ctx: EvalContext, from: number): PairedMoments | CellError {
  const first = positionNumbers(ctx, args.get(from));
  if (isError(first)) {
    return first;
  }
  const second = positionNumbers(ctx, args.get(from + 1));
  if (isError(second)) {
    return second;
  }
  if (first.size !== second.size) {
    return NA_ERROR;
  }
  const xs: number[] = [];
  const ys: number[] = [];
  for (const [position, x] of first.numbers) {
    const y = second.numbers.get(position);
    if (y !== undefined) {
      xs.push(x);
      ys.push(y);
    }
  }
  const n = xs.length;
  const mean1 = n === 0 ? 0 : sum(xs) / n;
  const mean2 = n === 0 ? 0 : sum(ys) / n;
  let s11 = 0;
  let s22 = 0;
  let s12 = 0;
  for (let i = 0; i < n; i += 1) {
    const dx = xs[i]! - mean1;
    const dy = ys[i]! - mean2;
    s11 += dx * dx;
    s22 += dy * dy;
    s12 += dx * dy;
  }
  return { n, mean1, mean2, s11, s22, s12 };
}

/** F(array1, array2) over the paired moments. */
function pairedImpl(compute: (m: PairedMoments) => number | CellError): Impl {
  return (args, ctx) => {
    const m = pairedMoments(args, ctx, 0);
    return isError(m) ? m : compute(m);
  };
}

/** Pearson's r; `empty` is the error for no pairs. */
function correlation(m: PairedMoments, empty: CellError): number | CellError {
  if (m.n === 0) {
    return empty;
  }
  return m.s11 === 0 || m.s22 === 0 ? DIV0_ERROR : m.s12 / Math.sqrt(m.s11 * m.s22);
}

/** The regression of the first argument (the y's) on the second (the
 * x's): #N/A for no pairs, #DIV/0! when the x's do not vary. */
function regression(m: PairedMoments): { slope: number; intercept: number } | CellError {
  if (m.n === 0) {
    return NA_ERROR;
  }
  if (m.s22 === 0) {
    return DIV0_ERROR;
  }
  const slope = m.s12 / m.s22;
  return { slope, intercept: m.mean1 - slope * m.mean2 };
}

const covariancePopulationImpl = pairedImpl((m) => (m.n === 0 ? DIV0_ERROR : m.s12 / m.n));

const forecastImpl: Impl = (args, ctx) => {
  const x = numberArg(args, 0, ctx);
  if (isError(x)) {
    return x;
  }
  const m = pairedMoments(args, ctx, 1);
  if (isError(m)) {
    return m;
  }
  const fit = regression(m);
  return isError(fit) ? fit : fit.intercept + fit.slope * x;
};

// --- distributions -----------------------------------------------------------------

const normDistImpl: Impl = (args, ctx) => {
  const values = numberArgs(args, ctx, 3);
  if (isError(values)) {
    return values;
  }
  const cumulative = logicalArg(args, 3);
  if (isError(cumulative)) {
    return cumulative;
  }
  const [x, mu, sigma] = values as [number, number, number];
  if (sigma <= 0) {
    return NUM_ERROR;
  }
  const z = (x - mu) / sigma;
  return finite(cumulative ? normSDist(z) : normSPdf(z) / sigma);
};

const normInvImpl: Impl = (args, ctx) => {
  const values = numberArgs(args, ctx, 3);
  if (isError(values)) {
    return values;
  }
  const [p, mu, sigma] = values as [number, number, number];
  if (!(p > 0 && p < 1) || sigma <= 0) {
    return NUM_ERROR;
  }
  return finite(mu + sigma * normSInv(p));
};

/** NORM.S.DIST(z, cumulative); NORMSDIST(z), without the argument, is the
 * CDF. */
const normSDistImpl: Impl = (args, ctx) => {
  const z = numberArg(args, 0, ctx);
  if (isError(z)) {
    return z;
  }
  const cumulative = args.length > 1 ? logicalArg(args, 1) : true;
  if (isError(cumulative)) {
    return cumulative;
  }
  return finite(cumulative ? normSDist(z) : normSPdf(z));
};

const normSInvImpl: Impl = (args, ctx) => {
  const p = numberArg(args, 0, ctx);
  if (isError(p)) {
    return p;
  }
  return p > 0 && p < 1 ? finite(normSInv(p)) : NUM_ERROR;
};

/** Degrees of freedom truncated, #NUM! below 1. */
function degreesOfFreedom(df: number): number | CellError {
  const whole = Math.trunc(df);
  return whole < 1 ? NUM_ERROR : whole;
}

/** The right tail P(T > x) — T.DIST.RT's kernel. */
function tDistRight(x: number, df: number): number | CellError {
  const nu = degreesOfFreedom(df);
  return isError(nu) ? nu : finite(studentTUpper(x, nu));
}

/** Both tails P(|T| > x), x ≥ 0 — T.DIST.2T's kernel. */
function tDistTwoTailed(x: number, df: number): number | CellError {
  const nu = degreesOfFreedom(df);
  if (isError(nu)) {
    return nu;
  }
  return x < 0 ? NUM_ERROR : finite(2 * studentTUpper(x, nu));
}

/** F(x, deg_freedom) over a kernel. */
function tKernelImpl(kernel: (x: number, df: number) => number | CellError): Impl {
  return (args, ctx) => {
    const values = numberArgs(args, ctx, 2);
    return isError(values) ? values : kernel(values[0]!, values[1]!);
  };
}

/** The two-tailed inverse: the t ≥ 0 with P(|T| > t) = p, 0 < p ≤ 1. */
const tInvTwoTailedImpl: Impl = (args, ctx) => {
  const values = numberArgs(args, ctx, 2);
  if (isError(values)) {
    return values;
  }
  const [p, df] = values as [number, number];
  const nu = degreesOfFreedom(df);
  if (isError(nu)) {
    return nu;
  }
  if (!(p > 0 && p <= 1)) {
    return NUM_ERROR;
  }
  return p === 1 ? 0 : finite(-studentTInv(p / 2, nu));
};

const binomDistImpl: Impl = (args, ctx) => {
  const values = numberArgs(args, ctx, 3);
  if (isError(values)) {
    return values;
  }
  const cumulative = logicalArg(args, 3);
  if (isError(cumulative)) {
    return cumulative;
  }
  const successes = Math.trunc(values[0]!);
  const trials = Math.trunc(values[1]!);
  const p = values[2]!;
  if (successes < 0 || successes > trials || !(p >= 0 && p <= 1)) {
    return NUM_ERROR;
  }
  return finite(cumulative ? binomialCdf(successes, trials, p) : binomialPmf(successes, trials, p));
};

const poissonDistImpl: Impl = (args, ctx) => {
  const values = numberArgs(args, ctx, 2);
  if (isError(values)) {
    return values;
  }
  const cumulative = logicalArg(args, 2);
  if (isError(cumulative)) {
    return cumulative;
  }
  const k = Math.trunc(values[0]!);
  const lambda = values[1]!;
  if (k < 0 || lambda < 0) {
    return NUM_ERROR;
  }
  return finite(cumulative ? poissonCdf(k, lambda) : poissonPmf(k, lambda));
};

const exponDistImpl: Impl = (args, ctx) => {
  const values = numberArgs(args, ctx, 2);
  if (isError(values)) {
    return values;
  }
  const cumulative = logicalArg(args, 2);
  if (isError(cumulative)) {
    return cumulative;
  }
  const [x, lambda] = values as [number, number];
  if (x < 0 || lambda <= 0) {
    return NUM_ERROR;
  }
  return finite(cumulative ? -Math.expm1(-lambda * x) : lambda * Math.exp(-lambda * x));
};

const confidenceNormImpl: Impl = (args, ctx) => {
  const values = numberArgs(args, ctx, 3);
  if (isError(values)) {
    return values;
  }
  const [alpha, sigma, rawSize] = values as [number, number, number];
  const size = Math.trunc(rawSize);
  if (!(alpha > 0 && alpha < 1) || sigma <= 0 || size < 1) {
    return NUM_ERROR;
  }
  return finite((-normSInv(alpha / 2) * sigma) / Math.sqrt(size));
};

// --- the -A variants and the remaining data-set statistics ----------------------------------

/** MAXA and MINA: 0 when there is nothing to compare. */
function extremeA(pick: (a: number, b: number) => number): Impl {
  return (args, ctx) => over(valuesA(args, ctx), (xs) => (xs.length === 0 ? 0 : xs.reduce((a, b) => pick(a, b))));
}

/** The mean and the sample standard deviation's ingredients of a list:
 * Σ((x − mean)/s)^power with s the sample (`sample`) or population
 * standard deviation, or #DIV/0! with fewer than `minimum` values or s = 0. */
function standardizedPowerSum(xs: readonly number[], power: number, sample: boolean, minimum: number): number | CellError {
  if (xs.length < minimum) {
    return DIV0_ERROR;
  }
  const s = stdev(xs, sample);
  if (isError(s)) {
    return s;
  }
  if (s === 0) {
    return DIV0_ERROR;
  }
  const m = sum(xs) / xs.length;
  let total = 0;
  for (const x of xs) {
    total += ((x - m) / s) ** power;
  }
  return total;
}

const kurtImpl: Impl = (args, ctx) =>
  over(numbersOf(args, ctx), (xs) => {
    const total = standardizedPowerSum(xs, 4, true, 4);
    if (isError(total)) {
      return total;
    }
    const n = xs.length;
    return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * total - (3 * (n - 1) * (n - 1)) / ((n - 2) * (n - 3));
  });

const skewImpl: Impl = (args, ctx) =>
  over(numbersOf(args, ctx), (xs) => {
    const total = standardizedPowerSum(xs, 3, true, 3);
    const n = xs.length;
    return isError(total) ? total : (n / ((n - 1) * (n - 2))) * total;
  });

const skewPopulationImpl: Impl = (args, ctx) =>
  over(numbersOf(args, ctx), (xs) => {
    const total = standardizedPowerSum(xs, 3, false, 1);
    return isError(total) ? total : total / xs.length;
  });

const geomeanImpl: Impl = (args, ctx) =>
  over(numbersOf(args, ctx), (xs) => {
    if (xs.length === 0 || xs.some((x) => x <= 0)) {
      return NUM_ERROR;
    }
    let logs = 0;
    for (const x of xs) {
      logs += Math.log(x);
    }
    return finite(Math.exp(logs / xs.length));
  });

const harmeanImpl: Impl = (args, ctx) =>
  over(numbersOf(args, ctx), (xs) => {
    if (xs.length === 0 || xs.some((x) => x <= 0)) {
      return NUM_ERROR;
    }
    let reciprocals = 0;
    for (const x of xs) {
      reciprocals += 1 / x;
    }
    return xs.length / reciprocals;
  });

const avedevImpl: Impl = (args, ctx) =>
  over(numbersOf(args, ctx), (xs) => {
    if (xs.length === 0) {
      return NUM_ERROR;
    }
    const m = sum(xs) / xs.length;
    let total = 0;
    for (const x of xs) {
      total += Math.abs(x - m);
    }
    return total / xs.length;
  });

const devsqImpl: Impl = (args, ctx) => over(numbersOf(args, ctx), (xs) => (xs.length === 0 ? NUM_ERROR : squaredDeviations(xs)));

const trimmeanImpl: Impl = (args, ctx) => {
  const xs = numbersOf(slice(args, 0, 1), ctx);
  if (isError(xs)) {
    return xs;
  }
  const percent = numberArg(args, 1, ctx);
  if (isError(percent)) {
    return percent;
  }
  if (!(percent >= 0 && percent <= 1)) {
    return NUM_ERROR;
  }
  const trim = Math.floor(roundTo15((xs.length * percent) / 2));
  const kept = ascending(xs).slice(trim, xs.length - trim);
  return kept.length === 0 ? NUM_ERROR : sum(kept) / kept.length;
};

const frequencyImpl: Impl = (args, ctx) => {
  const data = numbersOf(slice(args, 0, 1), ctx);
  if (isError(data)) {
    return data;
  }
  const bins = numbersOf(slice(args, 1, 2), ctx);
  if (isError(bins)) {
    return bins;
  }
  const sortedData = ascending(data);
  const order = bins.map((_, index) => index).toSorted((i, j) => bins[i]! - bins[j]! || i - j);
  const counts = Array.from<number>({ length: bins.length + 1 }).fill(0);
  let cursor = 0;
  for (const index of order) {
    const start = cursor;
    while (cursor < sortedData.length && sortedData[cursor]! <= bins[index]!) {
      cursor += 1;
    }
    counts[index] = cursor - start;
  }
  counts[bins.length] = sortedData.length - cursor;
  return makeArray(bins.length + 1, 1, counts);
};

/** GAMMALN and GAMMALN.PRECISE: ln Γ(x) for x > 0. */
const gammaLnImpl: Impl = (args, ctx) => {
  const x = numberArg(args, 0, ctx);
  if (isError(x)) {
    return x;
  }
  return x <= 0 ? NUM_ERROR : finite(gammaLn(x));
};

// --- metadata shared by a modern function and its legacy name ------------------------------

const ARRAY = arg("array", "The range or array of numbers.");
const PAIRED_ARGS = [arg("array1", "The first range or array of values."), arg("array2", "The second range or array, as many values.")];
const REGRESSION_ARGS = [
  arg("known_ys", "The dependent values, as many as known_xs."),
  arg("known_xs", "The independent values, as many as known_ys."),
];
const NORM_DIST_ARGS = [
  arg("x", "The value to evaluate the distribution at."),
  arg("mean", "The mean of the distribution."),
  arg("standard_dev", "The standard deviation; it must be positive."),
  arg("cumulative", "TRUE for the cumulative probability, FALSE for the density."),
];
const NORM_INV_ARGS = [
  arg("probability", "A probability strictly between 0 and 1."),
  arg("mean", "The mean of the distribution."),
  arg("standard_dev", "The standard deviation; it must be positive."),
];
const BINOM_ARGS = [
  arg("number_s", "The number of successes, truncated to an integer."),
  arg("trials", "The number of independent trials, truncated to an integer."),
  arg("probability_s", "The probability of success on each trial."),
  arg("cumulative", "TRUE for at most number_s successes, FALSE for exactly that many."),
];
const POISSON_ARGS = [
  arg("x", "The number of events, truncated to an integer."),
  arg("mean", "The expected number of events; not negative."),
  arg("cumulative", "TRUE for at most x events, FALSE for exactly x."),
];
const EXPON_ARGS = [
  arg("x", "The value to evaluate at; not negative."),
  arg("lambda", "The rate parameter; it must be positive."),
  arg("cumulative", "TRUE for the cumulative probability, FALSE for the density."),
];
const CONFIDENCE_ARGS = [
  arg("alpha", "The significance level, strictly between 0 and 1."),
  arg("standard_dev", "The population standard deviation; it must be positive."),
  arg("size", "The sample size, truncated to an integer of at least 1."),
];
const PERCENTILE_ARGS = [ARRAY, arg("k", "The percentile, from 0 to 1.")];
const QUARTILE_ARGS = [ARRAY, arg("quart", "Which quartile: 0 minimum, 1, 2 median, 3 or 4 maximum.")];
const RANK_ARGS = [
  arg("number", "The number whose rank is wanted."),
  arg("ref", "A reference to the list of numbers."),
  optional("order", "0 or omitted ranks descending; any other number ascending."),
];
const PERCENTRANK_ARGS = [
  ARRAY,
  arg("x", "The value whose rank is wanted."),
  optional("significance", "The significant digits kept, truncated (default 3)."),
];
const DF = arg("deg_freedom", "The degrees of freedom, truncated to an integer of at least 1.");

/** RANK.EQ's specs (modern and legacy) differ only in name and category. */
function rankSpec(name: string, category: FunctionCategory, summary: string, impl: Impl): FunctionSpec {
  return fixed(name, category, summary, RANK_ARGS, ["value", "range", "value"], impl);
}

const stat = "Statistical";
const legacy = "Compatibility";

export const STATISTICAL_FUNCTIONS: FunctionSpec[] = [
  variadic("MEDIAN", stat, "Returns the median, the middle value of the numbers.", medianImpl),
  variadic("MODE.SNGL", stat, "Returns the most frequent number, the first one met on a tie.", modeImpl),
  variadic("STDEV.S", stat, "Estimates the standard deviation from a sample.", stdevSampleImpl),
  variadic("STDEV.P", stat, "Returns the standard deviation of an entire population.", stdevPopulationImpl),
  variadic("VAR.S", stat, "Estimates the variance from a sample.", varSampleImpl),
  variadic("VAR.P", stat, "Returns the variance of an entire population.", varPopulationImpl),
  fixed(
    "LARGE",
    stat,
    "Returns the k-th largest number of a data set.",
    [ARRAY, arg("k", "The position from the largest, 1 for the largest.")],
    ["any", "value"],
    arrayAndNumber((xs, k) => kth(xs, k, true)),
  ),
  fixed(
    "SMALL",
    stat,
    "Returns the k-th smallest number of a data set.",
    [ARRAY, arg("k", "The position from the smallest, 1 for the smallest.")],
    ["any", "value"],
    arrayAndNumber((xs, k) => kth(xs, k, false)),
  ),
  rankSpec("RANK.EQ", stat, "Ranks a number in a list; equal numbers share the top rank.", rankEqImpl),
  rankSpec("RANK.AVG", stat, "Ranks a number in a list; equal numbers share their average rank.", rankImpl(true)),
  fixed("PERCENTILE.INC", stat, "Returns the k-th percentile, k from 0 to 1 inclusive.", PERCENTILE_ARGS, ["any", "value"], percentileIncImpl),
  fixed(
    "PERCENTILE.EXC",
    stat,
    "Returns the k-th percentile, k strictly between 0 and 1.",
    [ARRAY, arg("k", "The percentile, strictly between 0 and 1.")],
    ["any", "value"],
    arrayAndNumber(percentileExc),
  ),
  fixed("QUARTILE.INC", stat, "Returns a quartile of a data set, from the minimum to the maximum.", QUARTILE_ARGS, ["any", "value"], quartileIncImpl),
  fixed(
    "QUARTILE.EXC",
    stat,
    "Returns the first, second or third quartile, excluding the ends.",
    [ARRAY, arg("quart", "Which quartile: 1, 2 (the median) or 3.")],
    ["any", "value"],
    arrayAndNumber(quartileExc),
  ),
  fixed("PERCENTRANK.INC", stat, "Returns the rank of a value in a data set as a fraction from 0 to 1.", PERCENTRANK_ARGS, ["any", "value", "value"], percentRankIncImpl),
  fixed(
    "PERCENTRANK.EXC",
    stat,
    "Returns the rank of a value in a data set as a fraction strictly inside 0 to 1.",
    PERCENTRANK_ARGS,
    ["any", "value", "value"],
    percentRankImpl(true),
  ),
  fixed("CORREL", stat, "Returns the correlation coefficient of two data sets.", PAIRED_ARGS, ["array", "array"], pairedImpl((m) => correlation(m, DIV0_ERROR))),
  fixed("PEARSON", stat, "Returns Pearson's product-moment correlation coefficient.", PAIRED_ARGS, ["array", "array"], pairedImpl((m) => correlation(m, NA_ERROR))),
  fixed(
    "RSQ",
    stat,
    "Returns the square of Pearson's correlation coefficient.",
    REGRESSION_ARGS,
    ["array", "array"],
    pairedImpl((m) => {
      if (m.n === 1) {
        return DIV0_ERROR;
      }
      const r = correlation(m, NA_ERROR);
      return isError(r) ? r : r * r;
    }),
  ),
  fixed("COVARIANCE.P", stat, "Returns the population covariance of two data sets.", PAIRED_ARGS, ["array", "array"], covariancePopulationImpl),
  fixed(
    "COVARIANCE.S",
    stat,
    "Returns the sample covariance of two data sets.",
    PAIRED_ARGS,
    ["array", "array"],
    pairedImpl((m) => (m.n < 2 ? DIV0_ERROR : m.s12 / (m.n - 1))),
  ),
  fixed(
    "SLOPE",
    stat,
    "Returns the slope of the least-squares regression line.",
    REGRESSION_ARGS,
    ["array", "array"],
    pairedImpl((m) => {
      const fit = regression(m);
      return isError(fit) ? fit : fit.slope;
    }),
  ),
  fixed(
    "INTERCEPT",
    stat,
    "Returns where the least-squares regression line crosses the y-axis.",
    REGRESSION_ARGS,
    ["array", "array"],
    pairedImpl((m) => {
      const fit = regression(m);
      return isError(fit) ? fit : fit.intercept;
    }),
  ),
  fixed(
    "STEYX",
    stat,
    "Returns the standard error of the y values predicted by the regression.",
    REGRESSION_ARGS,
    ["array", "array"],
    pairedImpl((m) => {
      if (m.n === 0) {
        return NA_ERROR;
      }
      if (m.n < 3 || m.s22 === 0) {
        return DIV0_ERROR;
      }
      return Math.sqrt(Math.max(0, m.s11 - (m.s12 * m.s12) / m.s22) / (m.n - 2));
    }),
  ),
  fixed(
    "FORECAST.LINEAR",
    stat,
    "Predicts a y value for an x from a least-squares line through known points.",
    [arg("x", "The x value to predict a y value for."), ...REGRESSION_ARGS],
    ["value", "array", "array"],
    forecastImpl,
  ),
  fixed("NORM.DIST", stat, "Returns the normal distribution's density or cumulative probability.", NORM_DIST_ARGS, ["value", "value", "value", "value"], normDistImpl),
  fixed("NORM.INV", stat, "Returns the value at which a normal distribution reaches a probability.", NORM_INV_ARGS, ["value", "value", "value"], normInvImpl),
  fixed(
    "NORM.S.DIST",
    stat,
    "Returns the standard normal distribution's density or cumulative probability.",
    [arg("z", "The value to evaluate the distribution at."), arg("cumulative", "TRUE for the cumulative probability, FALSE for the density.")],
    ["value", "value"],
    normSDistImpl,
  ),
  fixed("NORM.S.INV", stat, "Returns the inverse of the standard normal cumulative distribution.", [NORM_INV_ARGS[0]!], ["value"], normSInvImpl),
  fixed(
    "STANDARDIZE",
    stat,
    "Returns how many standard deviations a value lies from the mean.",
    [arg("x", "The value to standardise."), arg("mean", "The mean of the distribution."), arg("standard_dev", "The standard deviation; it must be positive.")],
    ["value", "value", "value"],
    (args, ctx) => {
      const values = numberArgs(args, ctx, 3);
      if (isError(values)) {
        return values;
      }
      const [x, mu, sigma] = values as [number, number, number];
      return sigma <= 0 ? NUM_ERROR : (x - mu) / sigma;
    },
  ),
  fixed(
    "T.DIST",
    stat,
    "Returns Student's left-tailed t-distribution or its density.",
    [arg("x", "The value to evaluate the distribution at."), DF, arg("cumulative", "TRUE for the cumulative probability, FALSE for the density.")],
    ["value", "value", "value"],
    (args, ctx) => {
      const values = numberArgs(args, ctx, 2);
      if (isError(values)) {
        return values;
      }
      const cumulative = logicalArg(args, 2);
      if (isError(cumulative)) {
        return cumulative;
      }
      const nu = degreesOfFreedom(values[1]!);
      if (isError(nu)) {
        return nu;
      }
      const x = values[0]!;
      return finite(cumulative ? studentTCdf(x, nu) : studentTPdf(x, nu));
    },
  ),
  fixed(
    "T.DIST.2T",
    stat,
    "Returns Student's two-tailed t-distribution.",
    [arg("x", "The value to evaluate at; not negative."), DF],
    ["value", "value"],
    tKernelImpl(tDistTwoTailed),
  ),
  fixed(
    "T.DIST.RT",
    stat,
    "Returns Student's right-tailed t-distribution.",
    [arg("x", "The value to evaluate the distribution at."), DF],
    ["value", "value"],
    tKernelImpl(tDistRight),
  ),
  fixed(
    "T.INV",
    stat,
    "Returns the left-tailed inverse of Student's t-distribution.",
    [arg("probability", "The left-tail probability, strictly between 0 and 1."), DF],
    ["value", "value"],
    (args, ctx) => {
      const values = numberArgs(args, ctx, 2);
      if (isError(values)) {
        return values;
      }
      const nu = degreesOfFreedom(values[1]!);
      if (isError(nu)) {
        return nu;
      }
      const p = values[0]!;
      return p > 0 && p < 1 ? finite(studentTInv(p, nu)) : NUM_ERROR;
    },
  ),
  fixed(
    "T.INV.2T",
    stat,
    "Returns the two-tailed inverse of Student's t-distribution.",
    [arg("probability", "The two-tailed probability, above 0 and at most 1."), DF],
    ["value", "value"],
    tInvTwoTailedImpl,
  ),
  fixed("BINOM.DIST", stat, "Returns the binomial distribution's probability of a number of successes.", BINOM_ARGS, ["value", "value", "value", "value"], binomDistImpl),
  fixed("POISSON.DIST", stat, "Returns the Poisson distribution's probability of a number of events.", POISSON_ARGS, ["value", "value", "value"], poissonDistImpl),
  fixed("EXPON.DIST", stat, "Returns the exponential distribution's density or cumulative probability.", EXPON_ARGS, ["value", "value", "value"], exponDistImpl),
  fixed("CONFIDENCE.NORM", stat, "Returns the half-width of a confidence interval for a mean, normally.", CONFIDENCE_ARGS, ["value", "value", "value"], confidenceNormImpl),
  variadic(
    "AVERAGEA",
    stat,
    "Returns the mean of values, counting text as 0 and TRUE as 1.",
    (args, ctx) => over(valuesA(args, ctx), mean),
    VALUE1,
    VALUE2,
  ),
  variadic("MAXA", stat, "Returns the largest value, counting text as 0 and TRUE as 1.", extremeA(Math.max), VALUE1, VALUE2),
  variadic("MINA", stat, "Returns the smallest value, counting text as 0 and TRUE as 1.", extremeA(Math.min), VALUE1, VALUE2),
  variadic(
    "STDEVA",
    stat,
    "Estimates a sample's standard deviation, counting text as 0 and TRUE as 1.",
    (args, ctx) => over(valuesA(args, ctx), (xs) => stdev(xs, true)),
    VALUE1,
    VALUE2,
  ),
  variadic(
    "STDEVPA",
    stat,
    "Returns a population's standard deviation, counting text as 0 and TRUE as 1.",
    (args, ctx) => over(valuesA(args, ctx), (xs) => stdev(xs, false)),
    VALUE1,
    VALUE2,
  ),
  variadic(
    "VARA",
    stat,
    "Estimates a sample's variance, counting text as 0 and TRUE as 1.",
    (args, ctx) => over(valuesA(args, ctx), (xs) => variance(xs, true)),
    VALUE1,
    VALUE2,
  ),
  variadic(
    "VARPA",
    stat,
    "Returns a population's variance, counting text as 0 and TRUE as 1.",
    (args, ctx) => over(valuesA(args, ctx), (xs) => variance(xs, false)),
    VALUE1,
    VALUE2,
  ),
  variadic("GEOMEAN", stat, "Returns the geometric mean of positive numbers.", geomeanImpl),
  variadic("HARMEAN", stat, "Returns the harmonic mean of positive numbers.", harmeanImpl),
  variadic("AVEDEV", stat, "Returns the mean absolute deviation of numbers from their mean.", avedevImpl),
  variadic("DEVSQ", stat, "Returns the sum of squared deviations of numbers from their mean.", devsqImpl),
  variadic("KURT", stat, "Returns the sample excess kurtosis of a data set.", kurtImpl),
  variadic("SKEW", stat, "Returns the sample skewness of a distribution.", skewImpl),
  variadic("SKEW.P", stat, "Returns the skewness of a population.", skewPopulationImpl),
  fixed(
    "TRIMMEAN",
    stat,
    "Returns the mean of a data set with a fraction of its extremes left out.",
    [ARRAY, arg("percent", "The fraction of values to leave out, split between both ends.")],
    ["any", "value"],
    trimmeanImpl,
  ),
  fixed(
    "FREQUENCY",
    stat,
    "Counts how many values fall into each bin, as a vertical array.",
    [arg("data_array", "The values to count."), arg("bins_array", "The upper bounds of the bins.")],
    ["any", "any"],
    frequencyImpl,
  ),
  fixed(
    "GAMMA",
    stat,
    "Returns the gamma function of a number.",
    [arg("number", "The number; 0 and negative integers are #NUM!.")],
    ["value"],
    (args, ctx) => {
      const x = numberArg(args, 0, ctx);
      return isError(x) ? x : finite(gamma(x));
    },
  ),
  fixed("GAMMALN", stat, "Returns the natural logarithm of the gamma function.", [arg("x", "A positive number.")], ["value"], gammaLnImpl),
  fixed("GAMMALN.PRECISE", stat, "Returns the natural logarithm of the gamma function, precisely.", [arg("x", "A positive number.")], ["value"], gammaLnImpl),

  // Compatibility: Excel's legacy names, each sharing its modern function's impl.
  variadic("STDEV", legacy, "Estimates the standard deviation from a sample (legacy STDEV.S).", stdevSampleImpl),
  variadic("VAR", legacy, "Estimates the variance from a sample (legacy VAR.S).", varSampleImpl),
  fixed("PERCENTILE", legacy, "Returns the k-th percentile, k from 0 to 1 (legacy PERCENTILE.INC).", PERCENTILE_ARGS, ["any", "value"], percentileIncImpl),
  fixed("QUARTILE", legacy, "Returns a quartile of a data set (legacy QUARTILE.INC).", QUARTILE_ARGS, ["any", "value"], quartileIncImpl),
  rankSpec("RANK", legacy, "Ranks a number in a list of numbers (legacy RANK.EQ).", rankEqImpl),
  variadic("MODE", legacy, "Returns the most frequent number (legacy MODE.SNGL).", modeImpl),
  fixed("NORMDIST", legacy, "Returns the normal distribution (legacy NORM.DIST).", NORM_DIST_ARGS, ["value", "value", "value", "value"], normDistImpl),
  fixed("NORMINV", legacy, "Returns the inverse of the normal cumulative distribution (legacy NORM.INV).", NORM_INV_ARGS, ["value", "value", "value"], normInvImpl),
  fixed(
    "NORMSDIST",
    legacy,
    "Returns the standard normal cumulative distribution (legacy NORM.S.DIST).",
    [arg("z", "The value to evaluate the distribution at.")],
    ["value"],
    normSDistImpl,
  ),
  fixed(
    "NORMSINV",
    legacy,
    "Returns the inverse of the standard normal distribution (legacy NORM.S.INV).",
    [NORM_INV_ARGS[0]!],
    ["value"],
    normSInvImpl,
  ),
  fixed(
    "TDIST",
    legacy,
    "Returns Student's t-distribution, one- or two-tailed (legacy).",
    [arg("x", "The value to evaluate at; not negative."), DF, arg("tails", "1 for one tail, 2 for both; truncated to an integer.")],
    ["value", "value", "value"],
    (args, ctx) => {
      const values = numberArgs(args, ctx, 3);
      if (isError(values)) {
        return values;
      }
      const [x, df, rawTails] = values as [number, number, number];
      const nu = degreesOfFreedom(df);
      if (isError(nu)) {
        return nu;
      }
      const tails = Math.trunc(rawTails);
      if (x < 0 || (tails !== 1 && tails !== 2)) {
        return NUM_ERROR;
      }
      return tails === 1 ? tDistRight(x, nu) : tDistTwoTailed(x, nu);
    },
  ),
  fixed(
    "TINV",
    legacy,
    "Returns the two-tailed inverse of Student's t-distribution (legacy T.INV.2T).",
    [arg("probability", "The two-tailed probability, above 0 and at most 1."), DF],
    ["value", "value"],
    tInvTwoTailedImpl,
  ),
  fixed("BINOMDIST", legacy, "Returns the binomial distribution probability (legacy BINOM.DIST).", BINOM_ARGS, ["value", "value", "value", "value"], binomDistImpl),
  fixed("POISSON", legacy, "Returns the Poisson distribution probability (legacy POISSON.DIST).", POISSON_ARGS, ["value", "value", "value"], poissonDistImpl),
  fixed("EXPONDIST", legacy, "Returns the exponential distribution (legacy EXPON.DIST).", EXPON_ARGS, ["value", "value", "value"], exponDistImpl),
  fixed("CONFIDENCE", legacy, "Returns a confidence interval's half-width (legacy CONFIDENCE.NORM).", CONFIDENCE_ARGS, ["value", "value", "value"], confidenceNormImpl),
  fixed("COVAR", legacy, "Returns the population covariance (legacy COVARIANCE.P).", PAIRED_ARGS, ["array", "array"], covariancePopulationImpl),
  fixed(
    "FORECAST",
    legacy,
    "Predicts a y value along a least-squares line (legacy FORECAST.LINEAR).",
    [arg("x", "The x value to predict a y value for."), ...REGRESSION_ARGS],
    ["value", "array", "array"],
    forecastImpl,
  ),
  fixed("PERCENTRANK", legacy, "Returns the rank of a value as a fraction of a data set (legacy).", PERCENTRANK_ARGS, ["any", "value", "value"], percentRankIncImpl),
  variadic("STDEVP", legacy, "Returns the standard deviation of a population (legacy STDEV.P).", stdevPopulationImpl),
  variadic("VARP", legacy, "Returns the variance of a population (legacy VAR.P).", varPopulationImpl),
];
