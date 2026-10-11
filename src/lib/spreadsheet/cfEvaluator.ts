// Conditional-format evaluation (SPEC-129 "Evaluation semantics", "Value
// objects and statistics", "Priority and layering", "The evaluator") — THIS
// FILE is the normative home in code for when a rule is true, the statistics
// and thresholds statistics-based kinds resolve, the colour-scale, data-bar
// and icon geometry, how the rules covering a cell combine, and the
// viewport-driven evaluator the grid asks.
//
// The evaluator never walks the sheet on its own: `resultFor(row, col)`
// evaluates only the cells the grid mounts. Its caches are plain Maps it
// owns (never the model's):
// - a per-cell result cache, cleared by every `invalidate` (so a `timePeriod`
//   rule, or a formula calling a volatile function, re-evaluates on every
//   recalculation pass — the view invalidates on each RecalcResult);
// - per format, keyed by the format OBJECT (a command replaces what it
//   touches, so a changed format starts fresh): its origin and bounding box
//   (formats that cannot cover a cell are skipped without a range test), its
//   parsed formulas (parseFormulaAt, whose LRU shares an AST between equal
//   relative formulas), its statistics — computed lazily once, over the
//   stored cells of all its ranges, and kept until `invalidate` reports a
//   changed cell inside its ranges (or "all") — and each rule's resolved
//   thresholds, re-evaluated whenever the per-cell cache is cleared (a
//   formula threshold may read any cell).
// A replaced `conditionalFormats` array or `dxfs` table clears the per-cell
// cache by itself.
//
// Formulas: a `cellIs`/`expression` formula is stored relative to the
// format's origin, parsed once at it and evaluated at each cell with
// `host` = that cell and `origin` = the parse's origin — SPEC-120's own
// `part + (host − origin)` rule, so no text is re-shifted per cell. A
// value-object formula is absolute and evaluates with `host` = the origin.
// A formula that does not parse, or whose result is uncomputable, makes its
// rule false (its threshold unresolved).
import { resolveColor, type DocumentTheme, type ModelColor } from "../core/theme";
import type { CellRef, GridRange } from "./address";
import {
  boundingBox,
  DXF_BORDER_EDGES,
  formatOrigin,
  isFormattingRule,
  rangesContain,
  rulesByPriority,
  sheetFormats,
  type CfCellResult,
  type CfDataBar,
  type CfIconSet,
  type CfRule,
  type CfTimePeriod,
  type CfValue,
  type ConditionalFormat,
  type Dxf,
} from "./conditionalFormats";
import type { FormulaNode } from "./formula/ast";
import { toScalar } from "./formula/areas";
import { parseCellKey, type CellKey } from "./formula/dependencies";
import { evaluate, formulaFacts, type EvalContext } from "./formula/evaluate";
import { percentileInc } from "./formula/functions/stats";
import { FormulaSyntaxError, parseFormulaAt } from "./formula/parse";
import { UncomputableError } from "./formula/registry";
import { compareScalars, isEmpty, isError, roundTo15, toText, type ScalarValue } from "./formula/values";
import { MAX_COLS, type Cell, type CellValue, type Sheet, type SpreadsheetModel } from "./model";
import { daysFromCivil, serialToParts } from "./serial";

export interface CfEvaluator {
  resultFor(row: number, col: number): CfCellResult | undefined;
  invalidate(changed: CellKey[] | "all"): void;
}

export interface CfEvaluatorDeps {
  /** valueAt reads stored values; recordDependency a no-op (storedValueContext). */
  base: Omit<EvalContext, "host" | "origin" | "arrayMode">;
  /** model.theme ?? DEFAULT_THEME. */
  theme: DocumentTheme;
  /** Test hook: a format's statistics were computed. */
  onStatistics?(formatId: string): void;
}

// --- statistics ---------------------------------------------------------------------------------

/** What statistics-based kinds read: the numeric values (not booleans, not
 * numeric text) of all of a format's ranges, and every non-blank value's
 * duplicate key. */
export interface CfStatistics {
  count: number;
  min: number;
  max: number;
  mean: number;
  /** Population standard deviation. */
  stdDev: number;
  /** Ascending. */
  sorted: number[];
  /** Duplicate key → occurrences. */
  keys: Map<string, number>;
}

/** A value's duplicate key: the number, lower-cased text, booleans and
 * errors by their text — typed, so 1 and "1" differ; undefined for a blank. */
export function duplicateKey(value: CellValue | ScalarValue): string | undefined {
  if (value === null || value === "" || (typeof value === "object" && isEmpty(value as ScalarValue))) {
    return undefined;
  }
  if (typeof value === "number") {
    return `n${roundTo15(value)}`;
  }
  if (typeof value === "string") {
    return `s${value.toLowerCase()}`;
  }
  if (typeof value === "boolean") {
    return `b${value ? "TRUE" : "FALSE"}`;
  }
  return isError(value) ? `e${value.error}` : undefined;
}

/** The first stored row at or after `top` (binary search). */
function firstRowIndex(sheet: Sheet, top: number): number {
  let lo = 0;
  let hi = sheet.rows.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (sheet.rows[mid]!.index < top) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return lo;
}

/** Calls `visit` for every stored cell inside `range` — sparse: rows by
 * binary search, never a scan of the grid. */
function forEachStoredCell(sheet: Sheet, range: GridRange, visit: (row: number, cell: Cell) => void): void {
  for (let at = firstRowIndex(sheet, range.top); at < sheet.rows.length; at += 1) {
    const row = sheet.rows[at]!;
    if (row.index > range.bottom) {
      return;
    }
    for (const cell of row.cells) {
      if (cell.col > range.right) {
        break;
      }
      if (cell.col >= range.left) {
        visit(row.index, cell);
      }
    }
  }
}

function overlapping(ranges: readonly GridRange[]): boolean {
  for (let i = 0; i < ranges.length; i += 1) {
    for (let j = i + 1; j < ranges.length; j += 1) {
      const a = ranges[i]!;
      const b = ranges[j]!;
      if (a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right) {
        return true;
      }
    }
  }
  return false;
}

/** The statistics of `ranges` on `sheet` (a cell two ranges share counts
 * once). */
export function cfStatistics(sheet: Sheet, ranges: readonly GridRange[]): CfStatistics {
  const numbers: number[] = [];
  const keys = new Map<string, number>();
  const seen = overlapping(ranges) ? new Set<number>() : null;
  for (const range of ranges) {
    forEachStoredCell(sheet, range, (row, cell) => {
      if (seen) {
        const at = row * MAX_COLS + cell.col;
        if (seen.has(at)) {
          return;
        }
        seen.add(at);
      }
      if (typeof cell.value === "number") {
        numbers.push(cell.value);
      }
      const key = duplicateKey(cell.value);
      if (key !== undefined) {
        keys.set(key, (keys.get(key) ?? 0) + 1);
      }
    });
  }
  const sorted = numbers.toSorted((a, b) => a - b);
  const count = sorted.length;
  let total = 0;
  for (const value of sorted) {
    total += value;
  }
  const mean = count === 0 ? 0 : total / count;
  let squares = 0;
  for (const value of sorted) {
    squares += (value - mean) ** 2;
  }
  return {
    count,
    min: count === 0 ? 0 : sorted[0]!,
    max: count === 0 ? 0 : sorted[count - 1]!,
    mean,
    stdDev: count === 0 ? 0 : Math.sqrt(squares / count),
    sorted,
    keys,
  };
}

// --- value objects ------------------------------------------------------------------------------

const PLAIN_NUMBER = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?\s*$/i;

/** A value object's operand text as a number, or — when it is not a plain
 * number — the formula's result through `formula` (null: unresolved). */
function operand(text: string | undefined, formula: (text: string) => number | null): number | null {
  if (text === undefined || text.trim() === "") {
    return null;
  }
  return PLAIN_NUMBER.test(text) ? Number(text) : formula(text.replace(/^\s*=/, ""));
}

/** A value object resolved against `stats` (SPEC-129 "Value objects and
 * statistics"); null when it cannot be — no numeric cell for min/max/percent/
 * percentile, or a formula that yields no number. */
export function resolveCfValue(value: CfValue, stats: CfStatistics, formula: (text: string) => number | null): number | null {
  const needsData = value.type !== "num" && value.type !== "formula";
  if (needsData && stats.count === 0) {
    return null;
  }
  switch (value.type) {
    case "min":
      return stats.min;
    case "max":
      return stats.max;
    case "autoMin":
      return Math.min(0, stats.min);
    case "autoMax":
      return Math.max(0, stats.max);
    case "num":
      return operand(value.value, formula);
    case "formula":
      return value.value === undefined ? null : formula(value.value.replace(/^\s*=/, ""));
    case "percent": {
      const p = operand(value.value, formula);
      return p === null ? null : stats.min + ((stats.max - stats.min) * p) / 100;
    }
    case "percentile": {
      const p = operand(value.value, formula);
      if (p === null) {
        return null;
      }
      const result = percentileInc(stats.sorted, p / 100);
      return typeof result === "number" ? result : null;
    }
  }
}

// --- colour scales, data bars, icons ----------------------------------------------------------

/** A `#rrggbb` colour's sRGB channels, 0..255. */
function channels(hex: string): [number, number, number] {
  const n = hex.replace("#", "");
  return [parseInt(n.slice(0, 2), 16), parseInt(n.slice(2, 4), 16), parseInt(n.slice(4, 6), 16)];
}

/** Channels 0..255 as a lowercase `#rrggbb`. */
function hexOf(rgb: readonly number[]): string {
  return `#${rgb.map((channel) => Math.min(255, Math.max(0, Math.round(channel))).toString(16).padStart(2, "0")).join("")}`;
}

/** `hex` mixed `amount` (0..1) of the way toward white, per sRGB channel —
 * a gradient bar's far end, computed here rather than by a CSS colour
 * function. */
export function mixTowardWhite(hex: string, amount: number): string {
  return hexOf(channels(hex).map((channel) => channel + (255 - channel) * amount));
}

/** The colour a colour scale gives `value` (SPEC-129 "Colour scale"): the
 * first stop's at or below it, the last's at or above it, else each sRGB
 * channel interpolated linearly between the bracketing stops and rounded;
 * coincident stops resolve to the lower stop's colour. */
export function colorScaleColor(value: number, stops: readonly { value: number; color: string }[]): string {
  const first = stops[0]!;
  const last = stops.at(-1)!;
  if (value <= first.value) {
    return first.color;
  }
  if (value >= last.value) {
    return last.color;
  }
  for (let i = 0; i < stops.length - 1; i += 1) {
    const lo = stops[i]!;
    const hi = stops[i + 1]!;
    if (value < lo.value || value > hi.value) {
      continue;
    }
    if (hi.value === lo.value) {
      return lo.color;
    }
    const t = (value - lo.value) / (hi.value - lo.value);
    const a = channels(lo.color);
    const b = channels(hi.color);
    return hexOf(a.map((channel, i) => channel + (b[i]! - channel) * t));
  }
  return last.color;
}

export interface DataBarGeometry {
  /** Percent of the bar area from the leading (left) edge. */
  startPct: number;
  widthPct: number;
  negative: boolean;
  /** Where the axis sits, percent from the left; absent with no axis. */
  axisPct?: number;
}

function clamp(value: number, lo: number, hi: number): number {
  return Math.min(hi, Math.max(lo, value));
}

function pct(fraction: number): number {
  return Math.round(fraction * 10_000) / 100;
}

/** A data bar's geometry for `value` between the resolved `lo` and `hi`
 * (SPEC-129 "Data bar"): with lo ≥ 0 or no axis, a width between
 * minLengthPct and maxLengthPct from the leading edge (hi = lo → the
 * maximum); otherwise an axis at |lo|/(|lo| + hi) of the bar area (middle:
 * half; every value negative: the far edge), a positive value drawing
 * (v/hi)·(1 − a) right of it and a negative one (|v|/|lo|)·a left of it.
 * `rightToLeft` mirrors; `context` reads left to right. */
export function dataBarGeometry(value: number, lo: number, hi: number, bar: Pick<CfDataBar, "axis" | "direction" | "minLengthPct" | "maxLengthPct">): DataBarGeometry {
  const minLen = clamp(bar.minLengthPct, 0, 100) / 100;
  const maxLen = clamp(bar.maxLengthPct, 0, 100) / 100;
  let geometry: DataBarGeometry;
  if (lo >= 0 || bar.axis === "none") {
    const width = hi <= lo ? maxLen : minLen + ((maxLen - minLen) * clamp(value - lo, 0, hi - lo)) / (hi - lo);
    geometry = { startPct: 0, widthPct: pct(width), negative: value < 0 };
  } else {
    const axis = bar.axis === "middle" ? 0.5 : hi <= 0 ? 1 : Math.abs(lo) / (Math.abs(lo) + hi);
    if (value >= 0) {
      const width = hi <= 0 ? 0 : (clamp(value, 0, hi) / hi) * (1 - axis);
      geometry = { startPct: pct(axis), widthPct: pct(width), negative: false, axisPct: pct(axis) };
    } else {
      const width = (clamp(Math.abs(value), 0, Math.abs(lo)) / Math.abs(lo)) * axis;
      geometry = { startPct: pct(axis - width), widthPct: pct(width), negative: true, axisPct: pct(axis) };
    }
  }
  if (bar.direction === "rightToLeft") {
    geometry = { ...geometry, startPct: pct(1 - (geometry.startPct + geometry.widthPct) / 100) };
    if (geometry.axisPct !== undefined) {
      geometry.axisPct = pct(1 - geometry.axisPct / 100);
    }
  }
  return geometry;
}

/** The icon `value` earns (SPEC-129 "Icon set"): the highest i whose
 * threshold it meets (≥ when gte, else >), the first threshold being the
 * floor; `reverse` maps i to n − 1 − i. */
export function iconIndex(value: number, thresholds: readonly { value: number; gte: boolean }[], reverse: boolean): number {
  const v = roundTo15(value);
  let index = 0;
  for (let i = 1; i < thresholds.length; i += 1) {
    const threshold = thresholds[i]!;
    const at = roundTo15(threshold.value);
    if (threshold.gte ? v >= at : v > at) {
      index = i;
    }
  }
  return reverse ? thresholds.length - 1 - index : index;
}

// --- time periods -----------------------------------------------------------------------------

/** Days since 1970-01-01 of a serial's date. */
function civilDay(serial: number, system: "1900" | "1904"): number {
  const parts = serialToParts(serial, system);
  return daysFromCivil(parts.year, parts.month, parts.day);
}

/** Sunday = 0 … Saturday = 6 (1970-01-01 was a Thursday). */
function weekday(day: number): number {
  return (((day + 4) % 7) + 7) % 7;
}

/** Whether `serial`'s date lies in `period` around `today` (local time).
 * Weeks start on Sunday; `last7Days` is today − 6 … today. */
export function inTimePeriod(serial: number, period: CfTimePeriod, today: Date, system: "1900" | "1904"): boolean {
  const day = civilDay(Math.floor(serial), system);
  const t = daysFromCivil(today.getFullYear(), today.getMonth() + 1, today.getDate());
  const week = (offset: number) => {
    const sunday = t - weekday(t) + offset * 7;
    return day >= sunday && day <= sunday + 6;
  };
  const parts = serialToParts(Math.floor(serial), system);
  const months = parts.year * 12 + parts.month - (today.getFullYear() * 12 + today.getMonth() + 1);
  const years = parts.year - today.getFullYear();
  switch (period) {
    case "today":
      return day === t;
    case "yesterday":
      return day === t - 1;
    case "tomorrow":
      return day === t + 1;
    case "last7Days":
      return day >= t - 6 && day <= t;
    case "thisWeek":
      return week(0);
    case "lastWeek":
      return week(-1);
    case "nextWeek":
      return week(1);
    case "thisMonth":
      return months === 0;
    case "lastMonth":
      return months === -1;
    case "nextMonth":
      return months === 1;
    case "thisYear":
      return years === 0;
    case "lastYear":
      return years === -1;
    case "nextYear":
      return years === 1;
  }
}

// --- layering ------------------------------------------------------------------------------------

/** Merges `dxf` into `into`, taking only the properties still unset (the
 * fill only while `fillTaken` is false). Returns whether the fill is taken
 * after the merge. */
function mergeUnset(into: Dxf, dxf: Dxf, fillTaken: boolean): boolean {
  if (dxf.font) {
    const font = (into.font ??= {});
    for (const key of ["color", "bold", "italic", "underline", "strike"] as const) {
      if (font[key] === undefined && dxf.font[key] !== undefined) {
        (font as Record<string, unknown>)[key] = dxf.font[key];
      }
    }
    if (Object.keys(font).length === 0) {
      delete into.font;
    }
  }
  let taken = fillTaken;
  if (!taken && dxf.fill !== undefined) {
    into.fill = dxf.fill;
    taken = true;
  }
  if (dxf.borders) {
    const borders = (into.borders ??= {});
    for (const edge of DXF_BORDER_EDGES) {
      if (borders[edge] === undefined && dxf.borders[edge] !== undefined) {
        borders[edge] = dxf.borders[edge];
      }
    }
    if (Object.keys(borders).length === 0) {
      delete into.borders;
    }
  }
  if (into.numberFormat === undefined && dxf.numberFormat !== undefined) {
    into.numberFormat = dxf.numberFormat;
  }
  return taken;
}

// --- the evaluator --------------------------------------------------------------------------------

interface ParsedFormula {
  node: FormulaNode;
  origin: CellRef;
}

interface FormatRuntime {
  format: ConditionalFormat;
  origin: CellRef;
  bbox: GridRange | null;
  stats?: CfStatistics;
  /** Rule id → its formulas parsed at the origin (null: one does not parse
   * or cannot be computed here). */
  parsed: Map<string, (ParsedFormula | null)[]>;
  /** Rule id → its resolved thresholds (null: unresolved). */
  thresholds: Map<string, number[] | null>;
}

/** One rule in the sheet's priority order, with its format's runtime. */
interface OrderedRule {
  rule: CfRule;
  runtime: FormatRuntime;
}

/** A formula parsed at `at` (parseFormulaAt's LRU), or null when it does
 * not parse or cannot be computed here. */
function parse(text: string, at: CellRef): ParsedFormula | null {
  const parsed = parseFormulaAt(text, at);
  if (parsed instanceof FormulaSyntaxError || formulaFacts(parsed.node).reason !== undefined) {
    return null;
  }
  return parsed;
}

function isTrueResult(value: ScalarValue | null): boolean {
  return value === true || (typeof value === "number" && value !== 0);
}

/** The viewport-driven evaluator of one sheet's conditional formats (see
 * the module header). */
export function createCfEvaluator(model: SpreadsheetModel, sheetId: string, deps: CfEvaluatorDeps): CfEvaluator {
  const runtimes = new WeakMap<ConditionalFormat, FormatRuntime>();
  let cells = new Map<number, CfCellResult | null>();
  let seenFormats: readonly ConditionalFormat[] | null = null;
  let seenDxfs: readonly Dxf[] | null = null;
  let order: OrderedRule[] = [];
  let active: FormatRuntime[] = [];

  function runtimeOf(format: ConditionalFormat): FormatRuntime {
    let runtime = runtimes.get(format);
    if (!runtime) {
      runtime = { format, origin: formatOrigin(format), bbox: boundingBox(format.ranges), parsed: new Map(), thresholds: new Map() };
      runtimes.set(format, runtime);
    }
    return runtime;
  }

  /** Re-reads the sheet's formats and the dxf table when either was
   * replaced, clearing the per-cell cache. */
  function sync(sheet: Sheet): void {
    const formats = sheetFormats(sheet);
    if (formats === seenFormats && model.dxfs === seenDxfs) {
      return;
    }
    seenFormats = formats;
    seenDxfs = model.dxfs;
    cells = new Map();
    active = formats.map(runtimeOf);
    const byFormat = new Map(active.map((runtime) => [runtime.format, runtime] as const));
    order = rulesByPriority(formats).map(({ format, rule }) => ({ rule, runtime: byFormat.get(format)! }));
  }

  function context(row: number, col: number, origin: CellRef): EvalContext {
    return { ...deps.base, host: { sheetId, row, col }, origin, arrayMode: false };
  }

  /** A parsed formula's scalar at (row, col); null when uncomputable. */
  function scalarAt(parsed: ParsedFormula, row: number, col: number): ScalarValue | null {
    const ctx = context(row, col, parsed.origin);
    try {
      return toScalar(ctx, evaluate(parsed.node, ctx));
    } catch (error) {
      if (error instanceof UncomputableError) {
        return null;
      }
      throw error;
    }
  }

  /** A rule's `cellIs`/`expression` formulas, parsed once at the origin. */
  function formulasOf(runtime: FormatRuntime, rule: CfRule & { kind: "cellIs" | "expression" }): (ParsedFormula | null)[] {
    let parsed = runtime.parsed.get(rule.id);
    if (!parsed) {
      const texts = rule.kind === "cellIs" ? rule.formulas : [rule.formula];
      parsed = texts.map((text) => parse(text, runtime.origin));
      runtime.parsed.set(rule.id, parsed);
    }
    return parsed;
  }

  function statsOf(runtime: FormatRuntime, sheet: Sheet): CfStatistics {
    if (!runtime.stats) {
      runtime.stats = cfStatistics(sheet, runtime.format.ranges);
      deps.onStatistics?.(runtime.format.id);
    }
    return runtime.stats;
  }

  /** A value-object formula evaluated at the origin, as a number. */
  function numberAtOrigin(runtime: FormatRuntime, text: string): number | null {
    const parsed = parse(text, runtime.origin);
    if (!parsed) {
      return null;
    }
    const value = scalarAt(parsed, runtime.origin.row, runtime.origin.col);
    return typeof value === "number" ? value : null;
  }

  /** A statistics-based rule's thresholds, resolved once per cache epoch. */
  function thresholdsOf(runtime: FormatRuntime, rule: CfRule, values: readonly CfValue[], sheet: Sheet): number[] | null {
    if (runtime.thresholds.has(rule.id)) {
      return runtime.thresholds.get(rule.id)!;
    }
    const stats = statsOf(runtime, sheet);
    const resolved: number[] = [];
    for (const value of values) {
      const number = resolveCfValue(value, stats, (text) => numberAtOrigin(runtime, text));
      if (number === null) {
        runtime.thresholds.set(rule.id, null);
        return null;
      }
      resolved.push(number);
    }
    runtime.thresholds.set(rule.id, resolved);
    return resolved;
  }

  /** A colour as a lowercase `#rrggbb`. */
  function color(value: ModelColor): string {
    return resolveColor(value, deps.theme).toLowerCase();
  }

  /** Whether a formatting rule is true for `value` at (row, col). */
  function formattingRuleIsTrue(rule: CfRule, runtime: FormatRuntime, value: ScalarValue, row: number, col: number, sheet: Sheet): boolean {
    const settings = deps.base.settings;
    switch (rule.kind) {
      case "cellIs": {
        const results = formulasOf(runtime, rule).map((parsed) => (parsed ? scalarAt(parsed, row, col) : null));
        if (results.some((result) => result === null || isError(result))) {
          return false;
        }
        const compare = (a: ScalarValue, b: ScalarValue): number | null => {
          const sign = compareScalars(a, b, settings);
          return typeof sign === "number" ? sign : null;
        };
        const first = compare(value, results[0]!);
        if (first === null) {
          return false;
        }
        switch (rule.operator) {
          case "equal":
            return first === 0;
          case "notEqual":
            return first !== 0;
          case "greaterThan":
            return first > 0;
          case "greaterThanOrEqual":
            return first >= 0;
          case "lessThan":
            return first < 0;
          case "lessThanOrEqual":
            return first <= 0;
          case "between":
          case "notBetween": {
            if (results.length < 2) {
              return false;
            }
            const swapped = compare(results[0]!, results[1]!);
            const [lo, hi] = swapped !== null && swapped > 0 ? [results[1]!, results[0]!] : [results[0]!, results[1]!];
            const aboveLo = compare(value, lo);
            const belowHi = compare(value, hi);
            if (aboveLo === null || belowHi === null) {
              return false;
            }
            const inside = aboveLo >= 0 && belowHi <= 0;
            return rule.operator === "between" ? inside : !inside;
          }
        }
        return false;
      }
      case "expression": {
        const parsed = formulasOf(runtime, rule)[0];
        return parsed ? isTrueResult(scalarAt(parsed, row, col)) : false;
      }
      case "top10": {
        if (typeof value !== "number") {
          return false;
        }
        const stats = statsOf(runtime, sheet);
        if (stats.count === 0 || rule.rank < 1) {
          return false;
        }
        const rank = Math.min(stats.count, rule.percent ? Math.max(1, Math.floor((stats.count * rule.rank) / 100)) : Math.floor(rule.rank));
        const v = roundTo15(value);
        return rule.bottom ? v <= roundTo15(stats.sorted[rank - 1]!) : v >= roundTo15(stats.sorted[stats.count - rank]!);
      }
      case "aboveAverage": {
        if (typeof value !== "number") {
          return false;
        }
        const stats = statsOf(runtime, sheet);
        if (stats.count === 0) {
          return false;
        }
        const v = roundTo15(value);
        if (rule.stdDev !== undefined) {
          const band = rule.stdDev * stats.stdDev;
          return rule.below ? v < roundTo15(stats.mean - band) : v > roundTo15(stats.mean + band);
        }
        const mean = roundTo15(stats.mean);
        if (rule.below) {
          return rule.equalAverage ? v <= mean : v < mean;
        }
        return rule.equalAverage ? v >= mean : v > mean;
      }
      case "duplicates": {
        const key = duplicateKey(value);
        if (key === undefined) {
          return false;
        }
        const count = statsOf(runtime, sheet).keys.get(key) ?? 0;
        return rule.unique ? count === 1 : count > 1;
      }
      case "text": {
        if (isError(value)) {
          return rule.operator === "notContains";
        }
        const text = (toText(value) as string).toLowerCase();
        const needle = rule.text.toLowerCase();
        switch (rule.operator) {
          case "contains":
            return text.includes(needle);
          case "notContains":
            return !text.includes(needle);
          case "beginsWith":
            return text.startsWith(needle);
          case "endsWith":
            return text.endsWith(needle);
        }
        return false;
      }
      case "blanks": {
        if (isError(value)) {
          return false;
        }
        const blank = (toText(value) as string).trim() === "";
        return rule.not ? !blank : blank;
      }
      case "errors":
        return rule.not ? !isError(value) : isError(value);
      case "timePeriod":
        return typeof value === "number" && inTimePeriod(value, rule.period, deps.base.now(), deps.base.settings.dateSystem);
      default:
        return false;
    }
  }

  function scaleColorFor(rule: CfRule & { kind: "colorScale" }, runtime: FormatRuntime, value: number, sheet: Sheet): string | undefined {
    if (rule.stops.length < 2) {
      return undefined;
    }
    const at = thresholdsOf(runtime, rule, rule.stops.map((stop) => stop.at), sheet);
    if (!at) {
      return undefined;
    }
    return colorScaleColor(
      value,
      rule.stops.map((stop, i) => ({ value: at[i]!, color: color(stop.color) })),
    );
  }

  function dataBarFor(rule: CfRule & { kind: "dataBar" }, runtime: FormatRuntime, value: number, sheet: Sheet): CfCellResult["dataBar"] {
    const bar = rule.bar;
    const at = thresholdsOf(runtime, rule, [bar.min, bar.max], sheet);
    if (!at) {
      return undefined;
    }
    const geometry = dataBarGeometry(value, at[0]!, at[1]!, bar);
    const fill = color(geometry.negative && bar.negativeColor !== undefined ? bar.negativeColor : bar.color);
    const border = geometry.negative && bar.negativeBorderColor !== undefined ? bar.negativeBorderColor : bar.border;
    const result: NonNullable<CfCellResult["dataBar"]> = { ...geometry, color: fill };
    if (bar.gradient) {
      result.endColor = mixTowardWhite(fill, 0.9);
    }
    if (border !== undefined) {
      result.borderColor = color(border);
    }
    if (geometry.axisPct !== undefined) {
      result.axisColor = color(bar.axisColor ?? "#000000");
    }
    return result;
  }

  function iconFor(rule: CfRule & { kind: "iconSet" }, runtime: FormatRuntime, value: number, sheet: Sheet): { set: CfIconSet["set"]; index: number } | undefined {
    const icons = rule.icons;
    const at = thresholdsOf(
      runtime,
      rule,
      icons.thresholds.map((threshold) => threshold.at),
      sheet,
    );
    if (!at || at.length === 0) {
      return undefined;
    }
    const index = iconIndex(
      value,
      at.map((threshold, i) => ({ value: threshold, gte: icons.thresholds[i]!.gte })),
      icons.reverse,
    );
    return { set: icons.set, index };
  }

  function compute(sheet: Sheet, row: number, col: number): CfCellResult | undefined {
    let value: ScalarValue | undefined;
    const merged: Dxf = {};
    const result: CfCellResult = {};
    let fillTaken = false;
    let barTaken = false;
    let iconTaken = false;
    let hide = false;
    for (const { rule, runtime } of order) {
      const box = runtime.bbox;
      if (!box || row < box.top || row > box.bottom || col < box.left || col > box.right || !rangesContain(runtime.format.ranges, row, col)) {
        continue;
      }
      value ??= deps.base.valueAt(sheetId, row, col);
      if (rule.kind === "preserved") {
        continue;
      }
      if (rule.kind === "colorScale") {
        if (!fillTaken && typeof value === "number") {
          const scale = scaleColorFor(rule, runtime, value, sheet);
          if (scale !== undefined) {
            result.scaleColor = scale;
            fillTaken = true;
          }
        }
        continue;
      }
      if (rule.kind === "dataBar") {
        if (!barTaken && typeof value === "number") {
          const bar = dataBarFor(rule, runtime, value, sheet);
          if (bar) {
            result.dataBar = bar;
            barTaken = true;
            hide ||= !rule.bar.showValue;
          }
        }
        continue;
      }
      if (rule.kind === "iconSet") {
        if (!iconTaken && typeof value === "number") {
          const icon = iconFor(rule, runtime, value, sheet);
          if (icon) {
            iconTaken = true;
            hide ||= !rule.icons.showValue;
            if (icon.set !== "NoIcons") {
              result.icon = icon;
            }
          }
        }
        continue;
      }
      if (!isFormattingRule(rule) || !formattingRuleIsTrue(rule, runtime, value, row, col, sheet)) {
        continue;
      }
      const dxf = rule.dxfId === undefined ? undefined : model.dxfs[rule.dxfId];
      if (dxf) {
        fillTaken = mergeUnset(merged, dxf, fillTaken);
      }
      if (rule.stopIfTrue) {
        break;
      }
    }
    if (merged.font !== undefined || merged.fill !== undefined || merged.borders !== undefined || merged.numberFormat !== undefined) {
      result.dxf = merged;
    }
    if (hide) {
      result.hideValue = true;
    }
    return Object.keys(result).length === 0 ? undefined : result;
  }

  return {
    resultFor(row, col) {
      const sheet = model.sheets.find((each) => each.id === sheetId);
      if (!sheet) {
        return undefined;
      }
      sync(sheet);
      if (order.length === 0) {
        return undefined;
      }
      const key = row * MAX_COLS + col;
      const cached = cells.get(key);
      if (cached !== undefined) {
        return cached ?? undefined;
      }
      const result = compute(sheet, row, col);
      cells.set(key, result ?? null);
      return result;
    },
    invalidate(changed) {
      cells = new Map();
      for (const runtime of active) {
        runtime.thresholds.clear();
      }
      if (changed === "all") {
        for (const runtime of active) {
          delete runtime.stats;
        }
        seenFormats = null;
        return;
      }
      for (const key of changed) {
        const { sheetId: keySheet, row, col } = parseCellKey(key);
        if (keySheet !== sheetId) {
          continue;
        }
        for (const runtime of active) {
          const box = runtime.bbox;
          if (runtime.stats && box && row >= box.top && row <= box.bottom && col >= box.left && col <= box.right && rangesContain(runtime.format.ranges, row, col)) {
            delete runtime.stats;
          }
        }
      }
    },
  };
}
