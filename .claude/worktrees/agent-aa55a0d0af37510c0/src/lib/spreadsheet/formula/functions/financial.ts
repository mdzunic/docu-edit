// Financial (SPEC-127 "Function pins", "Numerical methods"): PMT IPMT PPMT FV
// PV NPER RATE NPV XNPV IRR XIRR MIRR SLN SYD DB DDB CUMIPMT CUMPRINC EFFECT
// NOMINAL — Microsoft's reference pages first, OpenFormula (ODF 1.2 part 2,
// 6.12) where they are silent.
//
// - The annuity functions solve pv·(1+rate)^nper + pmt·(1+rate·type)·
//   ((1+rate)^nper − 1)/rate + fv = 0, which at rate 0 is pv + pmt·nper +
//   fv = 0; (1+rate)^n is exp(n·log1p(rate)), so a small rate keeps its
//   digits. A non-zero `type` means payments at the start of each period.
//   PMT needs nper > 0 and IPMT/PPMT 1 ≤ per ≤ nper (#NUM! otherwise); IPMT
//   is the rate times the balance before the payment (0 for a first payment
//   due at the start) and PPMT the payment less it.
// - RATE, IRR and XIRR iterate by Newton's method (numeric.ts) exactly as
//   "Numerical methods" states — from `guess` (0.1), with analytic
//   derivatives, 20 steps and |Δrate| < 1e-7 for RATE and IRR, 100 steps and
//   1e-8 for XIRR. Exhausting the cap, a zero or non-finite derivative, a
//   non-finite iterate or a rate ≤ −1 (a guess too) is #NUM!, never a
//   half-converged number. RATE needs nper > 0; IRR and XIRR a positive and
//   a negative cash flow.
// - NPV counts numbers only, as its page's Remarks state: empty cells,
//   logicals, text (a number's text too) and error values are skipped,
//   whether arguments or inside references and arrays; the first number is
//   period 1, and a rate of −1 is #DIV/0!. IRR and MIRR read numbers only from
//   their range or array (text, logicals and empty cells skipped, zeros
//   kept); MIRR without both signs, or with a rate of −1, is #DIV/0!.
// - XNPV and XIRR pair values and dates element by element, row-major, empty
//   cells included: different counts are #NUM!, a non-number #VALUE!, a date
//   outside 0…9999-12-31 #VALUE!; dates are truncated, one before the first
//   date is #NUM!, and days count over a 365-day year. XNPV needs rate > −1.
// - DB rounds its rate to three decimals (half away from zero, on the
//   15-digit form — Microsoft's rule) and takes OpenFormula's limits: cost >
//   0, salvage ≥ 0, life > 0, 0 < month < 13, the period truncated and ≥ 1;
//   the period after life gets the first year's missing months and any later
//   one 0. DDB is OpenFormula's algorithm (fractional periods allowed) under
//   0 ≤ salvage ≤ cost, 1 ≤ period ≤ life and factor > 0. SLN and SYD are
//   their formulas alone (a zero divisor is #DIV/0!).
// - CUMIPMT/CUMPRINC: rate, nper and pv > 0, start and end truncated with
//   1 ≤ start ≤ end ≤ nper, type exactly 0 or 1 (#NUM! otherwise); the sum
//   of IPMT/PPMT over those periods, telescoped through the balance so a long
//   span costs no loop. EFFECT and NOMINAL truncate npery and need a rate > 0
//   and npery ≥ 1.
import type { CellError } from "../../model";
import { forEachScalar, toArray } from "../areas";
import type { EvalContext } from "../evaluate";
import type { ArgList, FunctionSpec } from "../registry";
import { DIV0_ERROR, isError, NUM_ERROR, VALUE_ERROR, type EvalValue } from "../values";
import { arg, numberArg, numberOr, optional, repeating } from "./common";
import { maxSerial } from "./dateTime";
import { roundDecimal } from "./math";
import { newton } from "./numeric";

/** RATE's and IRR's iteration (SPEC-127 "Numerical methods"). */
const RATE_ITERATIONS = 20;
const RATE_TOLERANCE = 1e-7;
/** XIRR's iteration. */
const XIRR_ITERATIONS = 100;
const XIRR_TOLERANCE = 1e-8;
/** The rate every iterate must stay above. */
const RATE_FLOOR = -1;

// --- the annuity equation -------------------------------------------------------------

/** (1 + rate)^n — through exp(n·log1p(rate)) above −1, so a small rate keeps
 * its digits; directly at or below it. */
function compound(rate: number, n: number): number {
  return rate > -1 ? Math.exp(n * Math.log1p(rate)) : (1 + rate) ** n;
}

/** (1 + rate)^n − 1, the growth over n periods, without rounding 1 + rate. */
function growth(rate: number, n: number): number {
  return rate > -1 ? Math.expm1(n * Math.log1p(rate)) : (1 + rate) ** n - 1;
}

/** ((1 + rate)^n − 1) / rate, the annuity factor; n at rate 0. */
function annuityFactor(rate: number, n: number): number {
  return rate === 0 ? n : growth(rate, n) / rate;
}

/** The future value the annuity equation gives. */
function futureValue(rate: number, nper: number, pmt: number, pv: number, type: number): number {
  return -(pv * compound(rate, nper) + pmt * (1 + rate * type) * annuityFactor(rate, nper));
}

/** The present value the annuity equation gives. */
function presentValue(rate: number, nper: number, pmt: number, fv: number, type: number): number {
  return -(fv + pmt * (1 + rate * type) * annuityFactor(rate, nper)) / compound(rate, nper);
}

/** The payment per period the annuity equation gives. */
function payment(rate: number, nper: number, pv: number, fv: number, type: number): number {
  return -(fv + pv * compound(rate, nper)) / ((1 + rate * type) * annuityFactor(rate, nper));
}

/** The signed balance after `k` payments of `pmt` on a loan of `pv` (pv
 * itself before any): a payment due at the start of a period is made before
 * that period's interest accrues. */
function balance(rate: number, k: number, pmt: number, pv: number, type: number): number {
  const principal = pv * compound(rate, type === 1 && k > 0 ? k - 1 : k);
  return principal + pmt * annuityFactor(rate, k);
}

/** The interest part of payment `per` — the rate times the balance before
 * it, nothing for a first payment due at the start. */
function interestPart(rate: number, per: number, pmt: number, pv: number, type: number): number {
  return type === 1 && per === 1 ? 0 : -rate * balance(rate, per - 1, pmt, pv, type);
}

/** Arguments read as numbers, or the first error. */
type Numbers = number[] | CellError;

/** The first `count` value arguments as numbers, or the first error. */
function numbers(args: ArgList, ctx: EvalContext, count: number): Numbers {
  const out: number[] = [];
  for (let i = 0; i < count; i += 1) {
    const n = numberArg(args, i, ctx);
    if (isError(n)) {
      return n;
    }
    out.push(n);
  }
  return out;
}

/** Optional trailing arguments with their defaults (an empty slot reads as
 * 0), appended to `given`, or the first error. */
function withOptional(given: Numbers, args: ArgList, ctx: EvalContext, defaults: readonly number[]): Numbers {
  if (isError(given)) {
    return given;
  }
  const out = [...given];
  for (const fallback of defaults) {
    const n = numberOr(args, out.length, ctx, fallback);
    if (isError(n)) {
      return n;
    }
    out.push(n);
  }
  return out;
}

/** `type` as 0 (end of period) or 1 (start): anything non-zero is 1. */
function due(type: number): number {
  return type === 0 ? 0 : 1;
}

/** Argument metadata the annuity functions share. */
const RATE_ARG = arg("rate", "The interest rate per period.");
const NPER_ARG = arg("nper", "The total number of payment periods.");
const PV_ARG = arg("pv", "The present value: what the series of payments is worth now.");
const PMT_ARG = arg("pmt", "The payment made each period.");
const FV_ARG = optional("fv", "The balance wanted after the last payment (default 0).");
const TYPE_ARG = optional("type", "When payments fall due: 0 at the end of each period (default), 1 at the start.");
const GUESS_ARG = optional("guess", "A starting estimate of the rate (default 0.1).");

/** IPMT or PPMT: one period's interest or principal part. */
function periodPart(name: "IPMT" | "PPMT", summary: string): FunctionSpec {
  return {
    name,
    category: "Financial",
    minArgs: 4,
    maxArgs: 6,
    argKinds: ["value", "value", "value", "value", "value", "value"],
    summary,
    args: [RATE_ARG, arg("per", "The period to compute, from 1 to nper."), NPER_ARG, PV_ARG, FV_ARG, TYPE_ARG],
    impl(args, ctx) {
      const read = withOptional(numbers(args, ctx, 4), args, ctx, [0, 0]);
      if (isError(read)) {
        return read;
      }
      const [rate, per, nper, pv, fv, type] = read as [number, number, number, number, number, number];
      if (per < 1 || per > nper) {
        return NUM_ERROR;
      }
      const t = due(type);
      const pmt = payment(rate, nper, pv, fv, t);
      const interest = interestPart(rate, per, pmt, pv, t);
      return name === "IPMT" ? interest : pmt - interest;
    },
  };
}

/** CUMIPMT or CUMPRINC: interest or principal summed over start…end. */
function cumulative(name: "CUMIPMT" | "CUMPRINC", summary: string): FunctionSpec {
  return {
    name,
    category: "Financial",
    minArgs: 6,
    maxArgs: 6,
    argKinds: ["value", "value", "value", "value", "value", "value"],
    summary,
    args: [
      RATE_ARG,
      NPER_ARG,
      PV_ARG,
      arg("start_period", "The first period to include, counting from 1."),
      arg("end_period", "The last period to include."),
      arg("type", "When payments fall due: 0 at the end of each period, 1 at the start."),
    ],
    impl(args, ctx) {
      const read = numbers(args, ctx, 6);
      if (isError(read)) {
        return read;
      }
      const [rate, nper, pv, startPeriod, endPeriod, type] = read as [number, number, number, number, number, number];
      const start = Math.trunc(startPeriod);
      const end = Math.trunc(endPeriod);
      if (rate <= 0 || nper <= 0 || pv <= 0 || start < 1 || end < start || end > nper || (type !== 0 && type !== 1)) {
        return NUM_ERROR;
      }
      const pmt = payment(rate, nper, pv, 0, type);
      // Each period's principal is the balance's change, so the sum telescopes.
      const principal = balance(rate, end, pmt, pv, type) - balance(rate, start - 1, pmt, pv, type);
      return name === "CUMPRINC" ? principal : (end - start + 1) * pmt - principal;
    },
  };
}

/** A view of `args` from `from` on (NPV's values after its rate). */
function argsFrom(args: ArgList, from: number): ArgList {
  return {
    length: Math.max(0, args.length - from),
    get: (index) => args.get(index + from),
    isMissing: (index) => args.isMissing(index + from),
  };
}

/** The numbers of a range or array argument (IRR, MIRR): text, logicals and
 * empty cells skipped, zeros kept; the first error met instead. */
function cashFlows(ctx: EvalContext, value: EvalValue): number[] | CellError {
  const flows: number[] = [];
  let failure: CellError | undefined;
  forEachScalar(ctx, value, (scalar) => {
    if (isError(scalar)) {
      failure = scalar;
      return false;
    }
    if (typeof scalar === "number") {
      flows.push(scalar);
    }
    return true;
  });
  return failure ?? flows;
}

/** Whether a list holds a positive and a negative number. */
function hasBothSigns(flows: readonly number[]): boolean {
  return flows.some((flow) => flow > 0) && flows.some((flow) => flow < 0);
}

/** XNPV's and XIRR's cash flows and their times in 365-day years from the
 * first date (see the module header), or the error the pair gives. */
function datedFlows(ctx: EvalContext, valuesArg: EvalValue, datesArg: EvalValue): { flows: number[]; years: number[] } | CellError {
  const values = toArray(ctx, valuesArg);
  if (isError(values)) {
    return values;
  }
  const dates = toArray(ctx, datesArg);
  if (isError(dates)) {
    return dates;
  }
  const failure = values.data.find((value) => isError(value)) ?? dates.data.find((value) => isError(value));
  if (failure !== undefined) {
    return failure as CellError;
  }
  if (values.data.length !== dates.data.length) {
    return NUM_ERROR;
  }
  const last = maxSerial(ctx.settings);
  const flows: number[] = [];
  const days: number[] = [];
  for (let i = 0; i < values.data.length; i += 1) {
    const value = values.data[i];
    const date = dates.data[i];
    if (typeof value !== "number" || typeof date !== "number" || date < 0 || date >= last + 1) {
      return VALUE_ERROR;
    }
    flows.push(value);
    days.push(Math.trunc(date));
  }
  const first = days[0]!;
  if (days.some((day) => day < first)) {
    return NUM_ERROR;
  }
  return { flows, years: days.map((day) => (day - first) / 365) };
}

/** Σ flow·(1 + rate)^−time and its derivative in the rate — the net present
 * value IRR (times 0, 1, 2…) and XIRR (times in years) drive to zero. */
function discounted(flows: readonly number[], times: readonly number[], rate: number): { value: number; slope: number } {
  const log = Math.log1p(rate);
  let value = 0;
  let slope = 0;
  for (let i = 0; i < flows.length; i += 1) {
    const time = times[i]!;
    const term = flows[i]! * Math.exp(-time * log);
    value += term;
    slope -= (time * term) / (1 + rate);
  }
  return { value, slope };
}

/** A rate where `flows` at `times` are worth nothing now, by Newton's method
 * from `guess` (see the module header); #NUM! when it fails. */
function solveRate(flows: readonly number[], times: readonly number[], guess: number, maxIterations: number, tolerance: number): number | CellError {
  if (guess <= RATE_FLOOR) {
    return NUM_ERROR;
  }
  const root = newton(
    (rate) => discounted(flows, times, rate).value,
    (rate) => discounted(flows, times, rate).slope,
    { start: guess, maxIterations, tolerance, floor: RATE_FLOOR },
  );
  return root ?? NUM_ERROR;
}

/** The depreciation `period` (an integer ≥ 1) of DB, by OpenFormula's
 * recurrence on the rate rounded to three decimals. */
function decliningBalance(cost: number, salvage: number, life: number, period: number, month: number): number {
  const rate = roundDecimal(1 - (salvage / cost) ** (1 / life), 3, "half");
  let depreciation = (cost * rate * month) / 12;
  let total = depreciation;
  for (let p = 2; p <= period; p += 1) {
    if (p <= life) {
      depreciation = (cost - total) * rate;
    } else if (p - 1 <= life && month < 12) {
      // The period after life: the first year's missing months.
      depreciation = ((cost - total) * rate * (12 - month)) / 12;
    } else {
      return 0;
    }
    total += depreciation;
  }
  return depreciation;
}

/** DDB by OpenFormula's algorithm, which allows a fractional period. */
function doubleDeclining(cost: number, salvage: number, life: number, period: number, factor: number): number {
  let rate = factor / life;
  let before: number;
  if (rate >= 1) {
    rate = 1;
    before = period === 1 ? cost : 0;
  } else {
    before = cost * (1 - rate) ** (period - 1);
  }
  const after = cost * (1 - rate) ** period;
  const depreciation = after < salvage ? before - salvage : before - after;
  return Math.max(depreciation, 0);
}

export const FINANCIAL_FUNCTIONS: FunctionSpec[] = [
  {
    name: "PMT",
    category: "Financial",
    minArgs: 3,
    maxArgs: 5,
    argKinds: ["value", "value", "value", "value", "value"],
    summary: "Returns the payment per period for a loan or annuity at a constant rate.",
    args: [RATE_ARG, NPER_ARG, PV_ARG, FV_ARG, TYPE_ARG],
    impl(args, ctx) {
      const read = withOptional(numbers(args, ctx, 3), args, ctx, [0, 0]);
      if (isError(read)) {
        return read;
      }
      const [rate, nper, pv, fv, type] = read as [number, number, number, number, number];
      return nper <= 0 ? NUM_ERROR : payment(rate, nper, pv, fv, due(type));
    },
  },
  periodPart("IPMT", "Returns the interest part of one period's payment on a loan or annuity."),
  periodPart("PPMT", "Returns the principal part of one period's payment on a loan or annuity."),
  {
    name: "FV",
    category: "Financial",
    minArgs: 3,
    maxArgs: 5,
    argKinds: ["value", "value", "value", "value", "value"],
    summary: "Returns the future value of an investment with constant payments and rate.",
    args: [RATE_ARG, NPER_ARG, PMT_ARG, optional("pv", "The present value (default 0)."), TYPE_ARG],
    impl(args, ctx) {
      const read = withOptional(numbers(args, ctx, 3), args, ctx, [0, 0]);
      if (isError(read)) {
        return read;
      }
      const [rate, nper, pmt, pv, type] = read as [number, number, number, number, number];
      return futureValue(rate, nper, pmt, pv, due(type));
    },
  },
  {
    name: "PV",
    category: "Financial",
    minArgs: 3,
    maxArgs: 5,
    argKinds: ["value", "value", "value", "value", "value"],
    summary: "Returns the present value of a series of equal future payments.",
    args: [RATE_ARG, NPER_ARG, PMT_ARG, FV_ARG, TYPE_ARG],
    impl(args, ctx) {
      const read = withOptional(numbers(args, ctx, 3), args, ctx, [0, 0]);
      if (isError(read)) {
        return read;
      }
      const [rate, nper, pmt, fv, type] = read as [number, number, number, number, number];
      return presentValue(rate, nper, pmt, fv, due(type));
    },
  },
  {
    name: "NPER",
    category: "Financial",
    minArgs: 3,
    maxArgs: 5,
    argKinds: ["value", "value", "value", "value", "value"],
    summary: "Returns the number of periods an investment or loan needs.",
    args: [RATE_ARG, PMT_ARG, PV_ARG, FV_ARG, TYPE_ARG],
    impl(args, ctx) {
      const read = withOptional(numbers(args, ctx, 3), args, ctx, [0, 0]);
      if (isError(read)) {
        return read;
      }
      const [rate, pmt, pv, fv, type] = read as [number, number, number, number, number];
      if (rate === 0) {
        return pmt === 0 ? NUM_ERROR : -(pv + fv) / pmt;
      }
      if (rate <= -1) {
        return NUM_ERROR;
      }
      // (1 + rate)^nper · (pv + k) = k − fv, with k the payments' level.
      const k = (pmt * (1 + rate * due(type))) / rate;
      const ratio = (k - fv) / (k + pv);
      return k + pv === 0 || !(ratio > 0) ? NUM_ERROR : Math.log(ratio) / Math.log1p(rate);
    },
  },
  {
    name: "RATE",
    category: "Financial",
    minArgs: 3,
    maxArgs: 6,
    argKinds: ["value", "value", "value", "value", "value", "value"],
    summary: "Returns the interest rate per period of an annuity, found by iteration.",
    args: [NPER_ARG, PMT_ARG, PV_ARG, FV_ARG, TYPE_ARG, GUESS_ARG],
    impl(args, ctx) {
      const read = withOptional(numbers(args, ctx, 3), args, ctx, [0, 0, 0.1]);
      if (isError(read)) {
        return read;
      }
      const [nper, pmt, pv, fv, type, guess] = read as [number, number, number, number, number, number];
      if (nper <= 0 || guess <= RATE_FLOOR) {
        return NUM_ERROR;
      }
      const t = due(type);
      // f(r) = pv·(1+r)^n + pmt·(1+r·t)·A(r) + fv, A(r) = ((1+r)^n − 1)/r,
      // summed as (pv + fv) + pv·((1+r)^n − 1) + …: forming 1 + r first
      // would round away r's last digits, and the root with them.
      const f = (rate: number): number => pv + fv + pv * growth(rate, nper) + pmt * (1 + rate * t) * annuityFactor(rate, nper);
      const df = (rate: number): number => {
        const factor = annuityFactor(rate, nper);
        // A'(r) = (n·(1+r)^(n−1)·r − ((1+r)^n − 1)) / r², n(n−1)/2 at 0.
        const slope = rate === 0 ? (nper * (nper - 1)) / 2 : (nper * compound(rate, nper - 1) - factor) / rate;
        return pv * nper * compound(rate, nper - 1) + pmt * (t * factor + (1 + rate * t) * slope);
      };
      return newton(f, df, { start: guess, maxIterations: RATE_ITERATIONS, tolerance: RATE_TOLERANCE, floor: RATE_FLOOR }) ?? NUM_ERROR;
    },
  },
  {
    name: "NPV",
    category: "Financial",
    minArgs: 2,
    maxArgs: 255,
    argKinds: ["value", "any", "any"],
    repeat: { from: 2, count: 1 },
    summary: "Returns the net present value of periodic cash flows at a discount rate.",
    args: [
      arg("rate", "The discount rate per period."),
      arg("value1", "A cash flow, or a range or array of them, the first at the end of period 1."),
      repeating("value2", "More cash flows, ranges or arrays, in period order."),
    ],
    impl(args, ctx) {
      const rate = numberArg(args, 0, ctx);
      if (isError(rate)) {
        return rate;
      }
      let total = 0;
      let period = 0;
      for (let i = 1; i < args.length; i += 1) {
        forEachScalar(ctx, args.get(i), (flow) => {
          if (typeof flow === "number") {
            period += 1;
            total += flow / compound(rate, period);
          }
        });
      }
      return rate === -1 ? DIV0_ERROR : total;
    },
  },
  {
    name: "XNPV",
    category: "Financial",
    minArgs: 3,
    maxArgs: 3,
    argKinds: ["value", "array", "array"],
    summary: "Returns the net present value of cash flows on irregular dates.",
    args: [
      arg("rate", "The annual discount rate."),
      arg("values", "The cash flows, one for each date."),
      arg("dates", "The payment dates; the first starts the schedule and none may precede it."),
    ],
    impl(args, ctx) {
      const rate = numberArg(args, 0, ctx);
      if (isError(rate)) {
        return rate;
      }
      const dated = datedFlows(ctx, args.get(1), args.get(2));
      if (isError(dated)) {
        return dated;
      }
      return rate <= RATE_FLOOR ? NUM_ERROR : discounted(dated.flows, dated.years, rate).value;
    },
  },
  {
    name: "IRR",
    category: "Financial",
    minArgs: 1,
    maxArgs: 2,
    argKinds: ["array", "value"],
    summary: "Returns the internal rate of return of periodic cash flows.",
    args: [arg("values", "A range or array of cash flows with at least one positive and one negative."), GUESS_ARG],
    impl(args, ctx) {
      const flows = cashFlows(ctx, args.get(0));
      if (isError(flows)) {
        return flows;
      }
      const guess = numberOr(args, 1, ctx, 0.1);
      if (isError(guess)) {
        return guess;
      }
      if (!hasBothSigns(flows)) {
        return NUM_ERROR;
      }
      return solveRate(flows, flows.map((_, i) => i), guess, RATE_ITERATIONS, RATE_TOLERANCE);
    },
  },
  {
    name: "XIRR",
    category: "Financial",
    minArgs: 2,
    maxArgs: 3,
    argKinds: ["array", "array", "value"],
    summary: "Returns the internal rate of return of cash flows on irregular dates.",
    args: [
      arg("values", "The cash flows, one for each date, with at least one positive and one negative."),
      arg("dates", "The payment dates; the first starts the schedule and none may precede it."),
      GUESS_ARG,
    ],
    impl(args, ctx) {
      const dated = datedFlows(ctx, args.get(0), args.get(1));
      if (isError(dated)) {
        return dated;
      }
      const guess = numberOr(args, 2, ctx, 0.1);
      if (isError(guess)) {
        return guess;
      }
      if (!hasBothSigns(dated.flows)) {
        return NUM_ERROR;
      }
      return solveRate(dated.flows, dated.years, guess, XIRR_ITERATIONS, XIRR_TOLERANCE);
    },
  },
  {
    name: "MIRR",
    category: "Financial",
    minArgs: 3,
    maxArgs: 3,
    argKinds: ["array", "value", "value"],
    summary: "Returns the modified internal rate of return of periodic cash flows.",
    args: [
      arg("values", "A range or array of cash flows with at least one positive and one negative."),
      arg("finance_rate", "The interest rate paid on the money the cash flows borrow."),
      arg("reinvest_rate", "The interest rate earned on the cash flows reinvested."),
    ],
    impl(args, ctx) {
      const flows = cashFlows(ctx, args.get(0));
      if (isError(flows)) {
        return flows;
      }
      const rates = numbers(argsFrom(args, 1), ctx, 2);
      if (isError(rates)) {
        return rates;
      }
      const [financeRate, reinvestRate] = rates as [number, number];
      if (!hasBothSigns(flows) || financeRate === -1 || reinvestRate === -1) {
        return DIV0_ERROR;
      }
      // The inflows carried to the last period against the outflows brought
      // back to the first.
      const n = flows.length;
      let future = 0;
      let present = 0;
      flows.forEach((flow, i) => {
        if (flow > 0) {
          future += flow * compound(reinvestRate, n - 1 - i);
        } else {
          present += flow / compound(financeRate, i);
        }
      });
      const ratio = future / -present;
      return ratio > 0 ? Math.expm1(Math.log(ratio) / (n - 1)) : NUM_ERROR;
    },
  },
  {
    name: "SLN",
    category: "Financial",
    minArgs: 3,
    maxArgs: 3,
    argKinds: ["value", "value", "value"],
    summary: "Returns the straight-line depreciation of an asset for one period.",
    args: [
      arg("cost", "The initial cost of the asset."),
      arg("salvage", "The value of the asset at the end of its life."),
      arg("life", "The number of periods the asset is depreciated over."),
    ],
    impl(args, ctx) {
      const read = numbers(args, ctx, 3);
      if (isError(read)) {
        return read;
      }
      const [cost, salvage, life] = read as [number, number, number];
      return life === 0 ? DIV0_ERROR : (cost - salvage) / life;
    },
  },
  {
    name: "SYD",
    category: "Financial",
    minArgs: 4,
    maxArgs: 4,
    argKinds: ["value", "value", "value", "value"],
    summary: "Returns the sum-of-years'-digits depreciation of an asset for a period.",
    args: [
      arg("cost", "The initial cost of the asset."),
      arg("salvage", "The value of the asset at the end of its life."),
      arg("life", "The number of periods the asset is depreciated over."),
      arg("per", "The period to compute, in the same units as life."),
    ],
    impl(args, ctx) {
      const read = numbers(args, ctx, 4);
      if (isError(read)) {
        return read;
      }
      const [cost, salvage, life, per] = read as [number, number, number, number];
      const digits = life * (life + 1);
      return digits === 0 ? DIV0_ERROR : ((cost - salvage) * (life - per + 1) * 2) / digits;
    },
  },
  {
    name: "DB",
    category: "Financial",
    minArgs: 4,
    maxArgs: 5,
    argKinds: ["value", "value", "value", "value", "value"],
    summary: "Returns the fixed-declining-balance depreciation of an asset for a period.",
    args: [
      arg("cost", "The initial cost of the asset."),
      arg("salvage", "The value of the asset at the end of its life."),
      arg("life", "The number of periods the asset is depreciated over."),
      arg("period", "The period to compute, in the same units as life."),
      optional("month", "The number of months in the first year (default 12)."),
    ],
    impl(args, ctx) {
      const read = withOptional(numbers(args, ctx, 4), args, ctx, [12]);
      if (isError(read)) {
        return read;
      }
      const [cost, salvage, life, periodGiven, month] = read as [number, number, number, number, number];
      const period = Math.trunc(periodGiven);
      if (cost <= 0 || salvage < 0 || life <= 0 || period < 1 || month <= 0 || month >= 13) {
        return NUM_ERROR;
      }
      return decliningBalance(cost, salvage, life, period, month);
    },
  },
  {
    name: "DDB",
    category: "Financial",
    minArgs: 4,
    maxArgs: 5,
    argKinds: ["value", "value", "value", "value", "value"],
    summary: "Returns the declining-balance depreciation of an asset at a chosen rate.",
    args: [
      arg("cost", "The initial cost of the asset."),
      arg("salvage", "The value of the asset at the end of its life; may be 0."),
      arg("life", "The number of periods the asset is depreciated over."),
      arg("period", "The period to compute, in the same units as life."),
      optional("factor", "How fast the balance declines (default 2, double-declining)."),
    ],
    impl(args, ctx) {
      const read = withOptional(numbers(args, ctx, 4), args, ctx, [2]);
      if (isError(read)) {
        return read;
      }
      const [cost, salvage, life, period, factor] = read as [number, number, number, number, number];
      if (salvage < 0 || salvage > cost || period < 1 || period > life || factor <= 0) {
        return NUM_ERROR;
      }
      return doubleDeclining(cost, salvage, life, period, factor);
    },
  },
  cumulative("CUMIPMT", "Returns the interest paid on a loan between two periods."),
  cumulative("CUMPRINC", "Returns the principal paid on a loan between two periods."),
  {
    name: "EFFECT",
    category: "Financial",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns the effective annual interest rate of a nominal rate.",
    args: [arg("nominal_rate", "The nominal annual interest rate."), arg("npery", "The number of compounding periods per year.")],
    impl(args, ctx) {
      const read = numbers(args, ctx, 2);
      if (isError(read)) {
        return read;
      }
      const rate = read[0]!;
      const periods = Math.trunc(read[1]!);
      return rate <= 0 || periods < 1 ? NUM_ERROR : Math.expm1(periods * Math.log1p(rate / periods));
    },
  },
  {
    name: "NOMINAL",
    category: "Financial",
    minArgs: 2,
    maxArgs: 2,
    argKinds: ["value", "value"],
    summary: "Returns the nominal annual interest rate of an effective rate.",
    args: [arg("effect_rate", "The effective annual interest rate."), arg("npery", "The number of compounding periods per year.")],
    impl(args, ctx) {
      const read = numbers(args, ctx, 2);
      if (isError(read)) {
        return read;
      }
      const rate = read[0]!;
      const periods = Math.trunc(read[1]!);
      return rate <= 0 || periods < 1 ? NUM_ERROR : periods * Math.expm1(Math.log1p(rate) / periods);
    },
  },
];
