// What the core function files share (SPEC-120 "The core functions"):
// typed argument reads, the aggregates' number walk ("Coercion": numbers only
// from references and arrays, direct arguments coerced) and the *IF/*IFS
// machinery (criteria pairs of one shape, matches found over stored cells,
// empty positions counted from an area's size).
import type { CellError } from "../../model";
import { areaCols, areaRows, forEachScalar, forEachStoredCell } from "../areas";
import type { EvalContext } from "../evaluate";
import type { ArgList, FunctionSpec } from "../registry";
import {
  isArray,
  isError,
  isRange,
  toLogical,
  toNumber,
  toText,
  VALUE_ERROR,
  type Area,
  type EvalValue,
  type ScalarValue,
} from "../values";
import { parseCriterion, type Criterion } from "./criteria";

type ArgMeta = FunctionSpec["args"][number];

/** A required argument's metadata. */
export function arg(name: string, description: string): ArgMeta {
  return { name, description };
}

/** An optional argument's metadata. */
export function optional(name: string, description: string): ArgMeta {
  return { name, description, optional: true };
}

/** A repeating (optional) argument's metadata. */
export function repeating(name: string, description: string): ArgMeta {
  return { name, description, optional: true, repeating: true };
}

/** A value-kind argument as the scalar the evaluator made it. */
export function scalarArg(args: ArgList, index: number): ScalarValue {
  const value = args.get(index);
  return isRange(value) || isArray(value) ? VALUE_ERROR : value;
}

/** A value argument as a number. */
export function numberArg(args: ArgList, index: number, ctx: EvalContext): number | CellError {
  return toNumber(scalarArg(args, index), ctx.settings);
}

/** A value argument as a number truncated toward zero. */
export function integerArg(args: ArgList, index: number, ctx: EvalContext): number | CellError {
  const n = numberArg(args, index, ctx);
  return isError(n) ? n : Math.trunc(n);
}

/** A value argument as text. */
export function textArg(args: ArgList, index: number): string | CellError {
  return toText(scalarArg(args, index));
}

/** A value argument as a logical. */
export function logicalArg(args: ArgList, index: number): boolean | CellError {
  return toLogical(scalarArg(args, index));
}

/** An argument omitted at the end of the call takes `fallback`; an empty
 * slot reads as EMPTY (0) like any other value. */
export function numberOr(args: ArgList, index: number, ctx: EvalContext, fallback: number): number | CellError {
  return index >= args.length ? fallback : numberArg(args, index, ctx);
}

/** Feeds every number an aggregate reads (see "Coercion"): from references
 * and arrays only numbers — text, logicals and empties skipped — and the
 * first error; a direct argument coerced (an empty slot is 0). Returns the
 * first error met. */
export function eachNumber(args: ArgList, ctx: EvalContext, visit: (n: number) => void): CellError | undefined {
  let failure: CellError | undefined;
  for (let i = 0; i < args.length; i += 1) {
    forEachScalar(ctx, args.get(i), (value, direct) => {
      if (isError(value)) {
        failure = value;
        return false;
      }
      if (typeof value === "number") {
        visit(value);
      } else if (direct) {
        const n = toNumber(value, ctx.settings);
        if (isError(n)) {
          failure = n;
          return false;
        }
        visit(n);
      }
      return true;
    });
    if (failure) {
      return failure;
    }
  }
  return undefined;
}

/** A single-area range argument's area, or #VALUE!. */
export function singleArea(value: EvalValue): Area | CellError {
  if (isError(value)) {
    return value;
  }
  return isRange(value) && value.areas.length === 1 ? value.areas[0]! : VALUE_ERROR;
}

/** `area` resized to `shape`'s rows × columns from its top-left cell (SUMIF
 * and AVERAGEIF's sum range), clamped to the grid. */
export function resizedArea(area: Area, shape: Area): Area {
  return {
    sheetId: area.sheetId,
    top: area.top,
    left: area.left,
    bottom: Math.min(area.top + areaRows(shape) - 1, 1_048_575),
    right: Math.min(area.left + areaCols(shape) - 1, 16_383),
  };
}

function sameShape(a: Area, b: Area): boolean {
  return areaRows(a) === areaRows(b) && areaCols(a) === areaCols(b);
}

export interface CriteriaPair {
  area: Area;
  criterion: Criterion;
}

/** The (range, criterion) pairs from argument `from` on, every range the
 * shape of `shape` (or of the first range) — #VALUE! otherwise. */
export function criteriaPairs(args: ArgList, from: number, ctx: EvalContext, shape?: Area): CriteriaPair[] | CellError {
  if ((args.length - from) % 2 !== 0 || args.length - from < 2) {
    return VALUE_ERROR;
  }
  const pairs: CriteriaPair[] = [];
  for (let i = from; i < args.length; i += 2) {
    const area = singleArea(args.get(i));
    if (isError(area)) {
      return area;
    }
    const reference = shape ?? pairs[0]?.area ?? area;
    if (!sameShape(area, reference)) {
      return VALUE_ERROR;
    }
    pairs.push({ area, criterion: parseCriterion(scalarArg(args, i + 1), ctx.settings) });
  }
  return pairs;
}

/** Whether every criterion holds at offset (dr, dc) of its range. */
function holdsAt(pairs: readonly CriteriaPair[], dr: number, dc: number, ctx: EvalContext): boolean {
  return pairs.every((pair) => pair.criterion.matches(ctx.valueAt(pair.area.sheetId, pair.area.top + dr, pair.area.left + dc)));
}

/** Visits, row-major, the stored cells of `values` at whose offset every
 * criterion holds; `visit` returns false to stop. */
export function forEachMatch(
  ctx: EvalContext,
  values: Area,
  pairs: readonly CriteriaPair[],
  visit: (value: ScalarValue) => boolean | void,
): void {
  forEachStoredCell(ctx, values, (value, row, col) => (holdsAt(pairs, row - values.top, col - values.left, ctx) ? visit(value) : true));
}

/** How many positions satisfy every criterion — COUNTIF/COUNTIFS. Only
 * stored positions are visited; the positions empty in every range count
 * from the area's size when every criterion matches an empty cell. */
export function countMatches(ctx: EvalContext, pairs: readonly CriteriaPair[]): number {
  const strict = pairs.find((pair) => !pair.criterion.matchesEmpty);
  let count = 0;
  if (strict) {
    // A match needs a stored cell in this range.
    forEachStoredCell(ctx, strict.area, (_value, row, col) => {
      if (holdsAt(pairs, row - strict.area.top, col - strict.area.left, ctx)) {
        count += 1;
      }
    });
    return count;
  }
  const first = pairs[0]!.area;
  const cols = areaCols(first);
  const stored = new Set<number>();
  for (const pair of pairs) {
    forEachStoredCell(ctx, pair.area, (_value, row, col) => {
      stored.add((row - pair.area.top) * cols + (col - pair.area.left));
    });
  }
  for (const offset of stored) {
    if (holdsAt(pairs, Math.floor(offset / cols), offset % cols, ctx)) {
      count += 1;
    }
  }
  return count + areaRows(first) * cols - stored.size;
}
