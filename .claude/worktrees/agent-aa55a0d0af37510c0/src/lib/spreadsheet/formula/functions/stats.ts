// Statistical kernels (SPEC-127 "Function pins", "SUBTOTAL and AGGREGATE")
// over plain lists of numbers — what the statistical functions and SUBTOTAL /
// AGGREGATE compute once their arguments are read:
//
// - Variance and standard deviation use the two-pass algorithm (the mean,
//   then the squared deviations from it), for stability. The sample forms
//   need two numbers and the population forms one (#DIV/0! otherwise).
// - MEDIAN of nothing is #NUM!. MODE.SNGL is the first value, in argument
//   order, with the highest count; with none repeated it is #N/A.
// - LARGE/SMALL take the k-th value, k rounded up as Excel does
//   (#NUM! for k < 1 or past the count).
// - PERCENTILE.INC interpolates at rank k·(n−1) from 0 (k in 0…1);
//   PERCENTILE.EXC at rank k·(n+1) from 1 (k strictly inside 0…1 and the
//   rank within 1…n); QUARTILE.INC/.EXC are those at quart/4, the quart
//   truncated (0…4, and 1…3 for .EXC). Anything else is #NUM!.
import type { CellError } from "../../model";
import { DIV0_ERROR, NA_ERROR, NUM_ERROR, roundTo15 } from "../values";

/** The sum of a list. */
export function sum(xs: readonly number[]): number {
  let total = 0;
  for (const x of xs) {
    total += x;
  }
  return total;
}

/** The arithmetic mean; #DIV/0! for no numbers. */
export function mean(xs: readonly number[]): number | CellError {
  return xs.length === 0 ? DIV0_ERROR : sum(xs) / xs.length;
}

/** Σ(x − mean)², two-pass (DEVSQ); 0 for no numbers. */
export function squaredDeviations(xs: readonly number[]): number {
  if (xs.length === 0) {
    return 0;
  }
  const m = sum(xs) / xs.length;
  let total = 0;
  for (const x of xs) {
    total += (x - m) * (x - m);
  }
  return total;
}

/** The variance — `sample` divides by n − 1 (two numbers needed), else by
 * n (one needed); #DIV/0! with too few. */
export function variance(xs: readonly number[], sample: boolean): number | CellError {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) {
    return DIV0_ERROR;
  }
  return squaredDeviations(xs) / (sample ? n - 1 : n);
}

/** The standard deviation (see variance). */
export function stdev(xs: readonly number[], sample: boolean): number | CellError {
  const v = variance(xs, sample);
  return typeof v === "number" ? Math.sqrt(v) : v;
}

/** A sorted copy, ascending. */
export function ascending(xs: readonly number[]): number[] {
  return xs.toSorted((a, b) => a - b);
}

/** The median; #NUM! for no numbers. */
export function median(xs: readonly number[]): number | CellError {
  const n = xs.length;
  if (n === 0) {
    return NUM_ERROR;
  }
  const sorted = ascending(xs);
  const mid = n >> 1;
  return n % 2 === 1 ? sorted[mid]! : (sorted[mid - 1]! + sorted[mid]!) / 2;
}

/** The most frequent value — the first, in order, of those with the highest
 * count; #N/A when no value repeats. Values are equal at 15 digits. */
export function modeSingle(xs: readonly number[]): number | CellError {
  const counts = new Map<number, number>();
  let best: number | undefined;
  let bestCount = 1;
  for (const x of xs) {
    const key = roundTo15(x);
    const count = (counts.get(key) ?? 0) + 1;
    counts.set(key, count);
    if (count > bestCount) {
      bestCount = count;
      best = undefined;
    }
  }
  if (bestCount < 2) {
    return NA_ERROR;
  }
  // The first value, in argument order, reaching the highest count.
  for (const x of xs) {
    if (counts.get(roundTo15(x)) === bestCount) {
      best = x;
      break;
    }
  }
  return best ?? NA_ERROR;
}

/** The k-th largest (`largest`) or smallest value, k rounded up; #NUM! for
 * k < 1 or past the count. */
export function kth(xs: readonly number[], k: number, largest: boolean): number | CellError {
  const rank = Math.ceil(roundTo15(k));
  if (!Number.isFinite(rank) || rank < 1 || rank > xs.length) {
    return NUM_ERROR;
  }
  const sorted = ascending(xs);
  return largest ? sorted[sorted.length - rank]! : sorted[rank - 1]!;
}

/** The value at a 0-based fractional rank of sorted numbers. */
function interpolate(sorted: readonly number[], rank: number): number {
  const lower = Math.floor(rank);
  const fraction = rank - lower;
  const low = sorted[lower]!;
  return fraction === 0 || lower + 1 >= sorted.length ? low : low + fraction * (sorted[lower + 1]! - low);
}

/** PERCENTILE.INC: k in 0…1, interpolated at k·(n−1); #NUM! otherwise or
 * for no numbers. */
export function percentileInc(xs: readonly number[], k: number): number | CellError {
  if (xs.length === 0 || !(k >= 0 && k <= 1)) {
    return NUM_ERROR;
  }
  return interpolate(ascending(xs), roundTo15(k * (xs.length - 1)));
}

/** PERCENTILE.EXC: k strictly inside 0…1 and its rank k·(n+1) within
 * 1…n, interpolated there; #NUM! otherwise. */
export function percentileExc(xs: readonly number[], k: number): number | CellError {
  const n = xs.length;
  if (n === 0 || !(k > 0 && k < 1)) {
    return NUM_ERROR;
  }
  const rank = roundTo15(k * (n + 1));
  if (rank < 1 || rank > n) {
    return NUM_ERROR;
  }
  return interpolate(ascending(xs), rank - 1);
}

/** QUARTILE.INC: quart truncated, 0…4. */
export function quartileInc(xs: readonly number[], quart: number): number | CellError {
  const q = Math.trunc(quart);
  return q < 0 || q > 4 ? NUM_ERROR : percentileInc(xs, q / 4);
}

/** QUARTILE.EXC: quart truncated, 1…3. */
export function quartileExc(xs: readonly number[], quart: number): number | CellError {
  const q = Math.trunc(quart);
  return q <= 0 || q >= 4 ? NUM_ERROR : percentileExc(xs, q / 4);
}
