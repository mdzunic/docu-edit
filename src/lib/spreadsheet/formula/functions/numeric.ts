// Numerical methods (SPEC-127 "Numerical methods") — framework-free, and
// tested alone against published tables (numeric.spec.ts).
//
// - newton: Newton's method with the stopping rules RATE (20 iterations,
//   |Δrate| < 1e-7), IRR (20, 1e-7) and XIRR (100, 1e-8) take from
//   Microsoft's pages, starting at `guess` (0.1). Exhausting the cap, a zero
//   or non-finite derivative, or a non-finite iterate or one at or below the
//   floor (a rate ≤ −1) fails — null, which the caller returns as #NUM! —
//   never a half-converged number.

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
