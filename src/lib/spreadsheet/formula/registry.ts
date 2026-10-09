// The function registry (SPEC-120 "The registry") — the one table every
// worksheet function is looked up in: this spec registers its core functions
// (formula/functions/*.ts, loaded through functions/index.ts) and
// SPEC-127 registers the rest of the library through the same
// registerFunction.
//
// A FunctionSpec declares its arity and one ArgKind per argument; the
// EVALUATOR applies the kinds, never the function: a "value" argument
// arrives as a scalar (implicit intersection outside array mode, lifted
// element-wise inside it), a "range" argument as a RangeValue (anything else
// is #VALUE!), an "any" argument as evaluated and an "array" argument as its
// subtree evaluated in array mode. An argument past `argKinds` takes its kind
// from `repeat`. An empty slot (`F(1,,3)`) reads as EMPTY; `isMissing` is
// true for it and for an omitted trailing argument, so a function can give
// the latter its default. Value-kind errors are returned before `impl` runs
// unless the function declares `acceptsErrors`; a `lazy` function evaluates
// an argument only when it asks for it (IF, IFS, SWITCH, CHOOSE, IFERROR,
// IFNA evaluate only the branch they take).
//
// `summary` and `args` follow SPEC-127's metadata invariants as far as they
// are stated there — a summary ending in ".", snake_case argument names,
// required arguments before optional ones; SPEC-127 owns and rewrites them.
import type { EvalContext, UncomputableReason } from "./evaluate";
import type { EvalValue } from "./values";

export type ArgKind = "value" | "range" | "any" | "array";

/** SPEC-127's display categories, adopted here so it narrows nothing. */
export type FunctionCategory =
  | "Financial"
  | "Date & Time"
  | "Math & Trig"
  | "Statistical"
  | "Lookup & Reference"
  | "Text"
  | "Logical"
  | "Information"
  | "Engineering"
  | "Compatibility";

export interface ArgList {
  readonly length: number;
  /** Lazy functions: evaluated on first get, memoised. Past `length`: EMPTY. */
  get(index: number): EvalValue;
  /** An empty slot `F(1,,3)` or an omitted trailing argument. */
  isMissing(index: number): boolean;
}

export interface FunctionSpec {
  /** Canonical, upper-case. */
  name: string;
  category: FunctionCategory;
  minArgs: number;
  /** 255 for variadic. */
  maxArgs: number;
  /** One per entry of `args` — a variadic function's repeating group once
   * in each (marked `repeating` in `args`). */
  argKinds: ArgKind[];
  /** args[from … from + count − 1] repeat as a group up to maxArgs (SPEC-127's
   *  field, defined here because argument kinds cycle the same way). */
  repeat?: { from: number; count: number };
  volatile?: true;
  lazy?: true;
  acceptsErrors?: true;
  impl(args: ArgList, ctx: EvalContext): EvalValue;
  summary: string;
  args: { name: string; description: string; optional?: true; repeating?: true }[];
}

/** Thrown by a function for a case this engine cannot compute (an ODS
 * regular-expression criterion): the evaluator turns it into the cell's
 * uncomputable reason, so the cell keeps its cached value. */
export class UncomputableError extends Error {
  readonly reason: UncomputableReason;

  constructor(reason: UncomputableReason) {
    super(`Uncomputable: ${reason.kind}`);
    this.name = "UncomputableError";
    this.reason = reason;
  }
}

const registry = new Map<string, FunctionSpec>();

/** Registers a function under its canonical name. Throws on a duplicate. */
export function registerFunction(spec: FunctionSpec): void {
  const name = spec.name.toUpperCase();
  if (registry.has(name)) {
    throw new Error(`The function ${name} is already registered`);
  }
  registry.set(name, spec);
}

/** The function a canonical name (any case) names, if registered. */
export function getFunction(name: string): FunctionSpec | undefined {
  return registry.get(name.toUpperCase());
}

/** Every registered name, sorted. */
export function registeredFunctionNames(): string[] {
  return [...registry.keys()].toSorted();
}

/** The kind of argument `index` (see the module header). */
export function argKindAt(spec: Pick<FunctionSpec, "argKinds" | "repeat">, index: number): ArgKind {
  const kinds = spec.argKinds;
  if (index < kinds.length) {
    return kinds[index]!;
  }
  if (spec.repeat) {
    const { from, count } = spec.repeat;
    return kinds[from + ((index - from) % count)] ?? "value";
  }
  return kinds.at(-1) ?? "value";
}
