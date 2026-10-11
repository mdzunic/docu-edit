// The formula interpreter (SPEC-120 "Evaluation", "Errors", "Implicit
// intersection", "Array mode", "Comparison") — a recursive interpreter over
// SPEC-119's FormulaNode. There is no code generation: formula text is data,
// and the Content-Security-Policy has no 'unsafe-eval' (purity.spec.ts holds
// this file to that).
//
// - An AST comes from parseFormulaAt, shared by every cell of a shared
//   formula; a relative RefPart resolves as `part + (host − origin)`.
// - Operators coerce every operand (values.ts); a binary operator returns
//   its left operand's error first. A non-finite number is #NUM!, a text
//   longer than MAX_TEXT_LENGTH #VALUE! — wherever a result is made.
// - Outside array mode a range or array in a value position (an operand, a
//   "value" argument, the result) is implicitly intersected (areas.ts). In
//   array mode — an `arrayRange` anchor's whole formula, an "array"
//   argument's subtree — a range there is materialised, operators apply
//   element-wise with broadcasting, and a function whose "value" argument
//   receives an array is applied per element (lifted), returning an array.
// - The evaluator applies the registry's argument kinds and arity; a lazy
//   function evaluates an argument only when it asks for it, and is lifted
//   by re-running it per element once an argument turns out to be an array.
// - The final operation: a root `+`/`−` (looking through parentheses) whose
//   non-zero operands cancel at 15 significant digits is exactly 0 —
//   `=0.1+0.2-0.3` is 0 while `=1*(0.1+0.2-0.3)` keeps its residue.
// - Per-AST facts — volatile, dynamic, the uncomputable reason, the
//   references, array constants — are memoised in WeakMaps keyed by the
//   shared node.
// - SPEC-127: a function's UnsupportedCaseError (a valid call the library
//   does not compute) makes the cell uncomputable with the reason
//   { kind: "unsupportedCase", name }; `rowVisibility` answers "hidden" for a
//   user-hidden row (SPEC-117's Row.hidden) and "visible" otherwise —
//   SPEC-128 widens it to answer "filtered" — for SUBTOTAL and AGGREGATE.
import { now as clockNow } from "../../core/clock";
import { findRow, getCell, type CellRef } from "../address";
import type { CalcSettings, CellError, CellValue, SpreadsheetModel } from "../model";
import { SPREADSHEET_LOCALES, type SpreadsheetLocale } from "../numberFormat/locale";
import type { BinaryOp, FormulaNode, NodeOf } from "./ast";
import {
  areaSize,
  broadcastShape,
  elementAt,
  intersect,
  makeArray,
  materialize,
  resolveReference,
  shapeOf,
  sheetById,
  toArray,
  toScalar,
  topLeft,
  type Shape,
} from "./areas";
// Loading the core library registers its functions.
import "./functions";
import { FormulaSyntaxError, parseFormulaAt } from "./parse";
import { argKindAt, getFunction, UncomputableError, UnsupportedCaseError, type ArgKind, type ArgList, type FunctionSpec } from "./registry";
import {
  compareScalars,
  DIV0_ERROR,
  EMPTY,
  errorOf,
  fromCellValue,
  isArray,
  isError,
  isRange,
  MAX_ARRAY_CELLS,
  MAX_TEXT_LENGTH,
  NAME_ERROR,
  NULL_ERROR,
  NUM_ERROR,
  power,
  REF_ERROR,
  roundTo15,
  toCellValue,
  toNumber,
  toText,
  VALUE_ERROR,
  type Area,
  type ArrayValue,
  type EvalValue,
  type ScalarValue,
} from "./values";
import type { ReferenceNode } from "./references";

export interface EvalSettings {
  locale: SpreadsheetLocale;
  dateSystem: "1900" | "1904";
  /** ods setting; xlsx false. */
  caseSensitive: boolean;
  /** ods setting; xlsx "wildcards". */
  criteria: "wildcards" | "regex" | "plain";
  /** ods setting; xlsx true. */
  wholeCell: boolean;
  /** SPEC-118's FormatContext value: "proleptic" for ods. */
  dateQuirks: "excel" | "proleptic";
}

/** How a row shows (SPEC-127): "hidden" by the user (Row.hidden),
 * "filtered" out by a filter (SPEC-128), else "visible". */
export type RowVisibility = "visible" | "filtered" | "hidden";

export interface EvalContext {
  model: SpreadsheetModel;
  host: { sheetId: string; row: number; col: number };
  /** From parseFormulaAt. */
  origin: CellRef;
  arrayMode: boolean;
  settings: EvalSettings;
  /** core/clock.ts now by default. */
  now(): Date;
  /** Math.random by default. */
  random(): number;
  /** Evaluates a dirty precedent first. */
  valueAt(sheetId: string, row: number, col: number): ScalarValue;
  /** INDIRECT/OFFSET results. */
  recordDependency(area: Area): void;
  /** Whether a row of a sheet shows (SUBTOTAL and AGGREGATE skip hidden
   * ones): "hidden" for Row.hidden, else "visible" (SPEC-128 adds
   * "filtered"). */
  rowVisibility(sheetId: string, row: number): RowVisibility;
  /** SPEC-133 adds `resolveName(name, sheetQualifier, host)` and
   *  `resolveStructured(node, host)` and owns their signatures. While a hook
   *  is absent, a name or structured node makes the cell uncomputable. */
}

export type UncomputableReason =
  | { kind: "function"; name: string }
  | { kind: "name"; name: string }
  | { kind: "structured" }
  | { kind: "external" }
  | { kind: "dynamicArray" }
  | { kind: "unparsed" }
  | { kind: "iterative" }
  | { kind: "regex" }
  | { kind: "dataTable" }
  | { kind: "unsupportedCase"; name: string };

export interface Uncomputable {
  uncomputable: UncomputableReason;
}

/** Nested on-demand evaluations past this re-queue the outer cell instead
 * of recursing (recalc.ts). */
export const MAX_EVAL_DEPTH = 512;

/** OFFSET and INDIRECT: their precedents are known only by evaluating them. */
const DYNAMIC_FUNCTIONS: ReadonlySet<string> = new Set(["OFFSET", "INDIRECT"]);

/** The settings a workbook evaluates under: its date system, its calculation
 * settings (`.ods`), the proleptic calendar for an `.ods`, en-US unless the
 * view passes the user's locale. */
export function evalSettingsFor(
  model: Pick<SpreadsheetModel, "dateSystem" | "format"> & { calc?: CalcSettings },
  overrides: Partial<EvalSettings> = {},
): EvalSettings {
  return {
    locale: SPREADSHEET_LOCALES["en-US"],
    dateSystem: model.dateSystem,
    caseSensitive: model.calc?.caseSensitive === true,
    criteria: model.calc?.criteria ?? "wildcards",
    wholeCell: model.calc?.partialCellMatch !== true,
    dateQuirks: model.format === "ods" ? "proleptic" : "excel",
    ...overrides,
  };
}

/** A context that reads the model as stored — for evaluating outside a
 * recalculation pass (tests, previews). */
export function storedValueContext(
  model: SpreadsheetModel,
  settings: Partial<EvalSettings> = {},
): Omit<EvalContext, "host" | "origin" | "arrayMode"> {
  return {
    model,
    settings: evalSettingsFor(model, settings),
    now: clockNow,
    random: Math.random,
    valueAt(sheetId, row, col) {
      const sheet = sheetById(model, sheetId);
      return fromCellValue(sheet ? (getCell(sheet, row, col)?.value ?? null) : null);
    },
    recordDependency() {},
    rowVisibility: (sheetId, row) => rowVisibilityIn(model, sheetId, row),
  };
}

/** A row's visibility as the model stores it (see EvalContext). */
export function rowVisibilityIn(model: SpreadsheetModel, sheetId: string, row: number): RowVisibility {
  const sheet = sheetById(model, sheetId);
  return sheet && findRow(sheet, row)?.hidden ? "hidden" : "visible";
}

export function isUncomputable(value: unknown): value is Uncomputable {
  return typeof value === "object" && value !== null && "uncomputable" in value;
}

// --- per-AST facts -----------------------------------------------------------------

export interface FormulaFacts {
  /** Calls a volatile function (NOW, TODAY, RAND, RANDBETWEEN, OFFSET, INDIRECT). */
  volatile: boolean;
  /** Calls OFFSET or INDIRECT. */
  dynamic: boolean;
  /** Why the formula cannot be computed here, when it cannot. */
  reason?: UncomputableReason;
  /** Every reference node, in source order. */
  refs: readonly ReferenceNode[];
}

const factsCache = new WeakMap<FormulaNode, FormulaFacts>();

/** The facts of a (shared) AST, memoised by node. */
export function formulaFacts(root: FormulaNode): FormulaFacts {
  let facts = factsCache.get(root);
  if (facts) {
    return facts;
  }
  let volatile = false;
  let dynamic = false;
  let reason: UncomputableReason | undefined;
  const refs: ReferenceNode[] = [];
  const visit = (node: FormulaNode): void => {
    switch (node.kind) {
      case "ref":
        refs.push(node.ref);
        return;
      case "name":
        reason ??= { kind: "name", name: node.name };
        return;
      case "structured":
        reason ??= { kind: "structured" };
        return;
      case "external":
        reason ??= { kind: "external" };
        return;
      case "function": {
        const spec = getFunction(node.name);
        if (!spec) {
          reason ??= { kind: "function", name: node.name };
        } else if (spec.volatile) {
          volatile = true;
        }
        if (DYNAMIC_FUNCTIONS.has(node.name)) {
          dynamic = true;
        }
        node.args.forEach(visit);
        return;
      }
      case "unary":
        visit(node.operand);
        return;
      case "binary":
        visit(node.left);
        visit(node.right);
        return;
      case "paren":
        visit(node.inner);
        return;
      default:
        return;
    }
  };
  visit(root);
  facts = reason ? { volatile, dynamic, reason, refs } : { volatile, dynamic, refs };
  factsCache.set(root, facts);
  return facts;
}

/** The reference operators: `:` (the bounding area of ranges on one sheet),
 * `" "` (the intersection, #NULL! when empty) and `,` (the union). */
function referenceOperator(op: ":" | " " | ",", left: EvalValue, right: EvalValue): EvalValue {
  if (isError(left)) {
    return left;
  }
  if (isError(right)) {
    return right;
  }
  if (!isRange(left) || !isRange(right)) {
    return VALUE_ERROR;
  }
  if (op === ",") {
    return { kind: "range", areas: [...left.areas, ...right.areas] };
  }
  if (op === ":") {
    const areas = [...left.areas, ...right.areas];
    const sheetId = areas[0]!.sheetId;
    if (areas.some((area) => area.sheetId !== sheetId)) {
      return REF_ERROR;
    }
    return {
      kind: "range",
      areas: [
        {
          sheetId,
          top: Math.min(...areas.map((area) => area.top)),
          left: Math.min(...areas.map((area) => area.left)),
          bottom: Math.max(...areas.map((area) => area.bottom)),
          right: Math.max(...areas.map((area) => area.right)),
        },
      ],
    };
  }
  const areas: Area[] = [];
  for (const a of left.areas) {
    for (const b of right.areas) {
      const top = Math.max(a.top, b.top);
      const bottom = Math.min(a.bottom, b.bottom);
      const leftCol = Math.max(a.left, b.left);
      const rightCol = Math.min(a.right, b.right);
      if (a.sheetId === b.sheetId && top <= bottom && leftCol <= rightCol) {
        areas.push({ sheetId: a.sheetId, top, left: leftCol, bottom, right: rightCol });
      }
    }
  }
  return areas.length > 0 ? { kind: "range", areas } : NULL_ERROR;
}

// --- values in value positions -------------------------------------------------------

/** A value where one value is expected: implicitly intersected outside
 * array mode, materialised (ranges) inside it. */
function inValuePosition(value: EvalValue, ctx: EvalContext): ScalarValue | ArrayValue {
  if (isRange(value)) {
    return ctx.arrayMode ? materialize(ctx, value) : intersect(ctx, value);
  }
  if (isArray(value)) {
    return ctx.arrayMode ? value : topLeft(value);
  }
  return value;
}

/** A number or text result made safe to store: non-finite → #NUM!, overlong
 * text → #VALUE!, elementwise for an array. */
function checked(value: EvalValue): EvalValue {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : NUM_ERROR;
  }
  if (typeof value === "string") {
    return value.length > MAX_TEXT_LENGTH ? VALUE_ERROR : value;
  }
  if (isArray(value)) {
    let clean = true;
    for (const element of value.data) {
      if (checked(element) !== element) {
        clean = false;
        break;
      }
    }
    return clean ? value : makeArray(value.rows, value.cols, value.data.map((element) => checked(element) as ScalarValue));
  }
  return value;
}

/** f over one operand, element-wise for an array. */
function map1(value: ScalarValue | ArrayValue, f: (x: ScalarValue) => ScalarValue): ScalarValue | ArrayValue {
  return isArray(value) ? makeArray(value.rows, value.cols, value.data.map(f)) : f(value);
}

/** f over two operands, broadcast element-wise when either is an array. */
function map2(
  a: ScalarValue | ArrayValue,
  b: ScalarValue | ArrayValue,
  f: (x: ScalarValue, y: ScalarValue) => ScalarValue,
): ScalarValue | ArrayValue {
  if (!isArray(a) && !isArray(b)) {
    return f(a, b);
  }
  const shape = broadcastShape(shapeOf(a), shapeOf(b));
  if (shape.rows * shape.cols > MAX_ARRAY_CELLS) {
    return NUM_ERROR;
  }
  const data: ScalarValue[] = [];
  for (let r = 0; r < shape.rows; r += 1) {
    for (let c = 0; c < shape.cols; c += 1) {
      data.push(f(elementAt(a, r, c), elementAt(b, r, c)));
    }
  }
  return makeArray(shape.rows, shape.cols, data);
}

// --- operators ---------------------------------------------------------------------

function arithmetic(op: BinaryOp, a: ScalarValue, b: ScalarValue, settings: EvalSettings, final: boolean): ScalarValue {
  const x = toNumber(a, settings);
  if (isError(x)) {
    return x;
  }
  const y = toNumber(b, settings);
  if (isError(y)) {
    return y;
  }
  let result: number;
  switch (op) {
    case "+":
      if (final && x !== 0 && y !== 0 && roundTo15(x) === roundTo15(-y)) {
        return 0;
      }
      result = x + y;
      break;
    case "-":
      if (final && x !== 0 && y !== 0 && roundTo15(x) === roundTo15(y)) {
        return 0;
      }
      result = x - y;
      break;
    case "*":
      result = x * y;
      break;
    case "/":
      if (y === 0) {
        return DIV0_ERROR;
      }
      result = x / y;
      break;
    default:
      return power(x, y);
  }
  return Number.isFinite(result) ? result : NUM_ERROR;
}

function concatenate(a: ScalarValue, b: ScalarValue): ScalarValue {
  const x = toText(a);
  if (isError(x)) {
    return x;
  }
  const y = toText(b);
  if (isError(y)) {
    return y;
  }
  return x.length + y.length > MAX_TEXT_LENGTH ? VALUE_ERROR : x + y;
}

function comparison(op: BinaryOp, a: ScalarValue, b: ScalarValue, settings: EvalSettings): ScalarValue {
  const order = compareScalars(a, b, settings);
  if (isError(order)) {
    return order;
  }
  switch (op) {
    case "=":
      return order === 0;
    case "<>":
      return order !== 0;
    case "<":
      return order < 0;
    case "<=":
      return order <= 0;
    case ">":
      return order > 0;
    default:
      return order >= 0;
  }
}

function evalUnary(node: NodeOf<"unary">, ctx: EvalContext): EvalValue {
  const operand = inValuePosition(evalNode(node.operand, ctx), ctx);
  if (node.op === "+") {
    return operand; // unary plus does not coerce (`=+"a"` is "a")
  }
  return map1(operand, (value) => {
    const n = toNumber(value, ctx.settings);
    if (isError(n)) {
      return n;
    }
    return node.op === "-" ? -n : n / 100;
  });
}

function evalBinary(node: NodeOf<"binary">, ctx: EvalContext, final: boolean): EvalValue {
  const op = node.op;
  if (op === ":" || op === " " || op === ",") {
    return referenceOperator(op, evalNode(node.left, ctx), evalNode(node.right, ctx));
  }
  const left = inValuePosition(evalNode(node.left, ctx), ctx);
  const right = inValuePosition(evalNode(node.right, ctx), ctx);
  switch (op) {
    case "&":
      return map2(left, right, concatenate);
    case "=":
    case "<>":
    case "<":
    case "<=":
    case ">":
    case ">=":
      return map2(left, right, (a, b) => comparison(op, a, b, ctx.settings));
    default:
      return map2(left, right, (a, b) => arithmetic(op, a, b, ctx.settings, final));
  }
}

// --- functions ---------------------------------------------------------------------

const arrayConstants = new WeakMap<FormulaNode, ArrayValue>();

function arrayConstant(node: NodeOf<"array">): ArrayValue {
  let value = arrayConstants.get(node);
  if (!value) {
    value = makeArray(node.rows.length, node.rows[0]?.length ?? 0, node.rows.flat());
    arrayConstants.set(node, value);
  }
  return value;
}

/** An argument evaluated per its kind (see registry.ts). */
function evalArgument(arg: FormulaNode, kind: ArgKind, ctx: EvalContext): EvalValue {
  if (arg.kind === "missing") {
    return EMPTY;
  }
  switch (kind) {
    case "value":
      return inValuePosition(evalNode(arg, ctx), ctx);
    case "range": {
      const value = evalNode(arg, ctx);
      return isRange(value) || isError(value) ? value : VALUE_ERROR;
    }
    case "any":
      return evalNode(arg, ctx);
    default:
      return evalNode(arg, ctx.arrayMode ? ctx : { ...ctx, arrayMode: true });
  }
}

/** One element of a lifted call reduced to a scalar. */
function liftedElement(value: EvalValue, ctx: EvalContext): ScalarValue {
  if (isRange(value)) {
    const array = materialize(ctx, value);
    return isError(array) ? array : topLeft(array);
  }
  return isArray(value) ? topLeft(value) : value;
}

/** Runs a non-lazy function once over evaluated arguments. */
function applyOnce(spec: FunctionSpec, values: EvalValue[], kinds: ArgKind[], isMissing: (index: number) => boolean, ctx: EvalContext): EvalValue {
  if (!spec.acceptsErrors) {
    for (let i = 0; i < values.length; i += 1) {
      const value = values[i]!;
      if ((kinds[i] === "value" || kinds[i] === "range") && isError(value)) {
        return value;
      }
    }
  }
  const args: ArgList = { length: values.length, get: (index) => values[index] ?? EMPTY, isMissing };
  return checked(spec.impl(args, ctx));
}

function callEager(spec: FunctionSpec, node: NodeOf<"function">, kinds: ArgKind[], ctx: EvalContext): EvalValue {
  const values = node.args.map((arg, i) => evalArgument(arg, kinds[i]!, ctx));
  const isMissing = (index: number): boolean => index >= values.length || node.args[index]!.kind === "missing";
  let shape: Shape | null = null;
  values.forEach((value, i) => {
    if (kinds[i] === "value" && isArray(value)) {
      shape = shape === null ? shapeOf(value) : broadcastShape(shape, shapeOf(value));
    }
  });
  if (shape === null) {
    return applyOnce(spec, values, kinds, isMissing, ctx);
  }
  const { rows, cols } = shape as Shape;
  if (rows * cols > MAX_ARRAY_CELLS) {
    return NUM_ERROR;
  }
  const data: ScalarValue[] = [];
  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      const scalars = values.map((value, i) => (kinds[i] === "value" ? elementAt(value as ScalarValue | ArrayValue, r, c) : value));
      data.push(liftedElement(applyOnce(spec, scalars, kinds, isMissing, ctx), ctx));
    }
  }
  return makeArray(rows, cols, data);
}

/** Thrown inside a lazy function's run when an argument is an array the
 * current lifting shape does not cover. */
class LiftSignal {
  constructor(readonly shape: Shape) {}
}

function callLazy(spec: FunctionSpec, node: NodeOf<"function">, kinds: ArgKind[], ctx: EvalContext): EvalValue {
  const count = node.args.length;
  const memo = new Map<number, EvalValue>();
  const evaluated = (index: number): EvalValue => {
    if (index >= count) {
      return EMPTY;
    }
    let value = memo.get(index);
    if (value === undefined) {
      value = evalArgument(node.args[index]!, kinds[index]!, ctx);
      memo.set(index, value);
    }
    return value;
  };
  const isMissing = (index: number): boolean => index >= count || node.args[index]!.kind === "missing";
  if (!ctx.arrayMode) {
    return checked(spec.impl({ length: count, get: evaluated, isMissing }, ctx));
  }
  const arrays = new Map<number, ArrayValue | CellError>();
  const asArray = (index: number, value: EvalValue): ScalarValue | ArrayValue => {
    if (!isRange(value)) {
      return value;
    }
    let array = arrays.get(index);
    if (array === undefined) {
      array = materialize(ctx, value);
      arrays.set(index, array);
    }
    return array;
  };
  let shape: Shape | null = null;
  for (;;) {
    try {
      if (shape === null) {
        // A first run: lifted only if a value argument it reads is an array.
        const probe: ArgList = {
          length: count,
          isMissing,
          get(index) {
            const value = evaluated(index);
            if (kinds[index] === "value" && isArray(value)) {
              if (value.rows * value.cols === 1) {
                return value.data[0]!;
              }
              throw new LiftSignal(shapeOf(value));
            }
            return value;
          },
        };
        return checked(spec.impl(probe, ctx));
      }
      const current: Shape = shape;
      if (current.rows * current.cols > MAX_ARRAY_CELLS) {
        return NUM_ERROR;
      }
      const data: ScalarValue[] = [];
      for (let r = 0; r < current.rows; r += 1) {
        for (let c = 0; c < current.cols; c += 1) {
          const element: ArgList = {
            length: count,
            isMissing,
            get(index) {
              const value = asArray(index, evaluated(index));
              if (!isArray(value)) {
                return value;
              }
              const wider = broadcastShape(current, shapeOf(value));
              if (wider.rows !== current.rows || wider.cols !== current.cols) {
                throw new LiftSignal(wider);
              }
              return elementAt(value, r, c);
            },
          };
          data.push(liftedElement(spec.impl(element, ctx), ctx));
        }
      }
      return checked(makeArray(current.rows, current.cols, data));
    } catch (signal) {
      if (!(signal instanceof LiftSignal)) {
        throw signal;
      }
      shape = shape === null ? signal.shape : broadcastShape(shape, signal.shape);
    }
  }
}

function callFunction(node: NodeOf<"function">, ctx: EvalContext): EvalValue {
  const spec = getFunction(node.name);
  if (!spec) {
    return NAME_ERROR;
  }
  const count = node.args.length;
  if (count < spec.minArgs || count > spec.maxArgs) {
    return VALUE_ERROR;
  }
  const kinds = node.args.map((_, i) => argKindAt(spec, i));
  try {
    return spec.lazy ? callLazy(spec, node, kinds, ctx) : callEager(spec, node, kinds, ctx);
  } catch (error) {
    if (error instanceof UnsupportedCaseError) {
      throw new UncomputableError({ kind: "unsupportedCase", name: spec.name });
    }
    throw error;
  }
}

// --- the interpreter -----------------------------------------------------------------

function evalNode(node: FormulaNode, ctx: EvalContext): EvalValue {
  switch (node.kind) {
    case "number":
    case "string":
    case "boolean":
      return node.value;
    case "error":
      return errorOf(node.code);
    case "array":
      return arrayConstant(node);
    case "ref":
      return resolveReference(node.ref, ctx);
    case "function":
      return callFunction(node, ctx);
    case "missing":
      return EMPTY;
    case "unary":
      return evalUnary(node, ctx);
    case "binary":
      return evalBinary(node, ctx, false);
    case "paren":
      return evalNode(node.inner, ctx);
    case "name":
      return NAME_ERROR;
    default:
      return REF_ERROR; // structured and external references (uncomputable)
  }
}

/** The formula root looking through parentheses. */
function throughParens(node: FormulaNode): FormulaNode {
  let root = node;
  while (root.kind === "paren") {
    root = root.inner;
  }
  return root;
}

/** Evaluates a formula's root node — ranges and arrays as they are (the
 * caller reduces a result), the final-operation rule at the root. */
export function evaluate(node: FormulaNode, ctx: EvalContext): EvalValue {
  const root = throughParens(node);
  if (root.kind === "binary" && (root.op === "+" || root.op === "-")) {
    return evalBinary(root, ctx, true);
  }
  return evalNode(node, ctx);
}

const uncomputable = (reason: UncomputableReason): Uncomputable => ({ uncomputable: reason });

/** Evaluates one cell's formula: its stored value for a cell without one;
 * the scalar to store for a normal formula (a top-level range or array
 * reduced as a value position, EMPTY stored as 0); the result ARRAY for an
 * `arrayRange` anchor (the caller distributes it); or why it cannot be
 * computed here (see "Stale cells"). */
export function evaluateCell(
  model: SpreadsheetModel,
  sheetId: string,
  row: number,
  col: number,
  base: Omit<EvalContext, "host" | "origin" | "arrayMode">,
): CellValue | ArrayValue | Uncomputable {
  const sheet = sheetById(model, sheetId);
  const cell = sheet ? getCell(sheet, row, col) : undefined;
  if (cell?.dataTableXml !== undefined) {
    return uncomputable({ kind: "dataTable" });
  }
  if (!cell || cell.formula === undefined) {
    return cell?.value ?? null;
  }
  if (cell.formulaUnparsed) {
    return uncomputable({ kind: "unparsed" });
  }
  const range = cell.arrayRange;
  const singleCell = range !== undefined && range.top === range.bottom && range.left === range.right;
  if (cell.dynamicArray && range !== undefined && !singleCell) {
    return uncomputable({ kind: "dynamicArray" });
  }
  const parsed = parseFormulaAt(cell.formula, { row, col });
  if (parsed instanceof FormulaSyntaxError) {
    return uncomputable({ kind: "unparsed" });
  }
  const facts = formulaFacts(parsed.node);
  if (facts.reason) {
    return uncomputable(facts.reason);
  }
  const ctx: EvalContext = { ...base, host: { sheetId, row, col }, origin: parsed.origin, arrayMode: range !== undefined };
  let value: EvalValue;
  try {
    value = evaluate(parsed.node, ctx);
  } catch (error) {
    if (error instanceof UncomputableError) {
      return uncomputable(error.reason);
    }
    throw error;
  }
  if (range === undefined) {
    return toCellValue(toScalar(ctx, value));
  }
  // A single-cell dynamic array computes only while its result fits its cell.
  if (cell.dynamicArray && ((isRange(value) && (value.areas.length !== 1 || areaSize(value.areas[0]!) > 1)) || (isArray(value) && value.rows * value.cols > 1))) {
    return uncomputable({ kind: "dynamicArray" });
  }
  const array = toArray(ctx, value);
  return isError(array) ? makeArray(1, 1, [array]) : array;
}
