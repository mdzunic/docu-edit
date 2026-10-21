// Data validation (SPEC-130) — THIS FILE is the normative source for the
// rule model (`DataValidation`, its types, operators, list sources and
// messages), the validation semantics (`validateValue`), the list a dropdown
// offers (`listItems`) and the rule commands.
//
// The model:
// - `Sheet.validations` holds a sheet's rules in source order (new ones
//   appended). A rule is a list of ranges and its settings; its ORIGIN is
//   `ranges[0]`'s top-left, which `formula1`/`formula2` are stored relative
//   to — parsed once with parseFormulaAt at the origin and evaluated at a
//   cell with `host` = that cell (SPEC-120's `part + (host − origin)` rule),
//   so no text is re-shifted per cell. Overlapping rules (files can hold
//   them; the dialog never makes them): the first in model order wins, for
//   enforcement and for the ODS cell attribute.
// - `showDropdown` is the MODEL's sense: true shows the arrow (xlsx
//   `showDropDown="1"` means hide it — the wire mappings invert it).
// - `sourceXml` is the verbatim source element while the rule is unchanged;
//   every command drops it from the rules it touches.
//
// What validates (Excel's rule): ONLY typed entry — the cell editor's
// commit, the formula bar's, a dropdown pick and Ctrl/Cmd+Enter's fill of
// the selection — checked through the view's entry gate
// (views/spreadsheet/useDataValidation.ts) before the set-cell command runs.
// Nothing else validates: paste (SPEC-125), fill, drag-move, Delete/Backspace
// clearing, undo/redo, sort, CSV import and recalculated formula results all
// write without a check. A copy-paste does not change which rules cover the
// destination cells; a cut-paste or drag-move carries the moved cells' rules
// (move.ts's `validations` adjuster).
//
// Commands (record → command → commit through SPEC-121's undo stack, one
// entry labelled DATA_VALIDATION_LABEL): setValidation and clearValidation.
// Each subtracts the target ranges from every other rule (subtractRanges; a
// rule that lost its origin is re-anchored, one left with no range is
// removed), replaces what it touches (copy-on-write: `validations` is a
// SHEET_BY_REFERENCE_FIELDS entry, shared by reference with undo snapshots),
// drops the touched rules' `sourceXml` and sets `Sheet.dirty`.
//
// Layering rule (boundary.spec.ts): framework-free, imports only spreadsheet
// siblings and src/lib/core.
import { getCell, subtractRanges, type CellRef, type GridRange } from "./address";
import { forEachStoredCell, toScalar } from "./formula/areas";
import type { FormulaNode } from "./formula/ast";
import { evaluate, formulaFacts, type EvalContext } from "./formula/evaluate";
import { FormulaSyntaxError, parseFormula, parseFormulaAt } from "./formula/parse";
import { shiftFormula } from "./formula/references";
import { UncomputableError } from "./formula/registry";
import { tokenize } from "./formula/tokens";
import { fromCellValue, isEmpty, isError, isRange, isArray, numberToText, roundTo15, toCellValue, type ScalarValue } from "./formula/values";
import type { CellValue, Sheet, SpreadsheetModel } from "./model";
import { formatCodeOf } from "./numberFormat/builtin";
import { parseInput } from "./numberFormat/input";
import { formatValue, plainText, type FormatContext } from "./numberFormat/render";
import { cellStyle, effectiveStyleId } from "./styles";

// --- the rule model ----------------------------------------------------------------------------

export type ValidationType = "any" | "whole" | "decimal" | "list" | "date" | "time" | "textLength" | "custom";
export type ValidationOperator =
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "lessThan"
  | "greaterThanOrEqual"
  | "lessThanOrEqual";
export type ValidationErrorStyle = "stop" | "warning" | "information";

export interface DataValidation {
  id: string; // minted at read; never written
  ranges: GridRange[]; // the origin is ranges[0]'s top-left
  type: ValidationType;
  operator: ValidationOperator; // used by whole/decimal/date/time/textLength
  formula1?: string; // canonical, relative to the origin
  formula2?: string;
  allowBlank: boolean;
  showDropdown: boolean; // true = show the arrow (xlsx showDropDown="1" means HIDE)
  input: { show: boolean; title: string; text: string };
  error: { show: boolean; style: ValidationErrorStyle; title: string; text: string };
  x14?: true; // xlsx: written into the worksheet's x14 ext
  sourceXml?: string; // verbatim source element while unchanged
  odsName?: string; // table:content-validation@table:name
  odsListOrder?: "sort-ascending"; // ods table:display-list="sort-ascending"
}

/** A rule as the dialog builds it: everything but the minted `id`, the
 * ranges the command is given and the wire identities. */
export type ValidationSettings = Omit<DataValidation, "id" | "ranges" | "sourceXml" | "odsName">;

export const VALIDATION_TYPES: readonly ValidationType[] = ["any", "whole", "decimal", "list", "date", "time", "textLength", "custom"];
export const VALIDATION_OPERATORS: readonly ValidationOperator[] = [
  "between",
  "notBetween",
  "equal",
  "notEqual",
  "greaterThan",
  "lessThan",
  "greaterThanOrEqual",
  "lessThanOrEqual",
];

/** The types an operator applies to. */
export function usesOperator(type: ValidationType): boolean {
  return type === "whole" || type === "decimal" || type === "date" || type === "time" || type === "textLength";
}

/** Whether an operator takes two bounds. */
export function isRangeOperator(operator: ValidationOperator): boolean {
  return operator === "between" || operator === "notBetween";
}

/** The most items a dropdown lists. */
export const VALIDATION_LIST_MAX = 32_767;
/** The longest literal list Source (Excel's limit on the quoted string). */
export const LITERAL_LIST_MAX_CHARS = 255;

/** The undo entry every validation command records. */
export const DATA_VALIDATION_LABEL = "Data validation";

/** The rule the dialog starts a cell without one from (Excel's defaults). */
export function defaultValidationSettings(): ValidationSettings {
  return {
    type: "any",
    operator: "between",
    allowBlank: true,
    showDropdown: true,
    input: { show: true, title: "", text: "" },
    error: { show: true, style: "stop", title: "", text: "" },
  };
}

/** A fresh rule id (never written to a file). */
export function mintValidationId(): string {
  return crypto.randomUUID();
}

/** A rule's origin: its first range's top-left. */
export function validationOrigin(dv: Pick<DataValidation, "ranges">): CellRef {
  const first = dv.ranges[0];
  return first ? { row: first.top, col: first.left } : { row: 0, col: 0 };
}

/** A sheet's rules, `[]` for a sheet stored before SPEC-130 that no reader
 * upgrade has reached. */
export function sheetValidations(sheet: Pick<Sheet, "validations">): DataValidation[] {
  return sheet.validations ?? [];
}

function covers(range: GridRange, row: number, col: number): boolean {
  return row >= range.top && row <= range.bottom && col >= range.left && col <= range.right;
}

/** The rule covering `ref` — the first in model order (see the header). */
export function validationAt(sheet: Pick<Sheet, "validations">, ref: CellRef): DataValidation | undefined {
  return sheetValidations(sheet).find((dv) => dv.ranges.some((range) => covers(range, ref.row, ref.col)));
}

// --- list sources -------------------------------------------------------------------------------

export type ListSource = { kind: "literal"; items: string[] } | { kind: "reference" | "name" | "formula"; formula: string };

/** A literal list's items — `"\"a,b,c\""` → `["a", "b", "c"]`: the string
 * literal's text split on commas, each item trimmed (an empty item kept) —
 * or null when `formula1` is not one string literal. */
export function parseListLiteral(formula1: string): string[] | null {
  const text = formula1.trim();
  if (text.length < 2 || !text.startsWith('"') || !text.endsWith('"')) {
    return null;
  }
  const body = text.slice(1, -1);
  // A lone quote inside would end the literal early: not one string.
  if (/(^|[^"])"(?!")/.test(body.replace(/""/g, ""))) {
    return null;
  }
  return body
    .replace(/""/g, '"')
    .split(",")
    .map((item) => item.trim());
}

/** The literal `formula1` of a list of items: comma-joined, quotes doubled. */
export function formatListLiteral(items: readonly string[]): string {
  return `"${items.join(",").replace(/"/g, '""')}"`;
}

function rootOf(node: FormulaNode): FormulaNode {
  let at = node;
  while (at.kind === "paren") {
    at = at.inner;
  }
  return at;
}

/** A `list` rule's source (see "List sources"), or null for a rule that is
 * not a list or has no Source. */
export function listSourceOf(dv: Pick<DataValidation, "type" | "formula1">): ListSource | null {
  if (dv.type !== "list" || dv.formula1 === undefined || dv.formula1.trim() === "") {
    return null;
  }
  const items = parseListLiteral(dv.formula1);
  if (items !== null) {
    return { kind: "literal", items };
  }
  let node: FormulaNode;
  try {
    node = rootOf(parseFormula(dv.formula1));
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      return { kind: "formula", formula: dv.formula1 };
    }
    throw error;
  }
  if (node.kind === "ref" || (node.kind === "binary" && node.op === ":" && rootOf(node.left).kind === "ref" && rootOf(node.right).kind === "ref")) {
    return { kind: "reference", formula: dv.formula1 };
  }
  if (node.kind === "name") {
    return { kind: "name", formula: dv.formula1 };
  }
  return { kind: "formula", formula: dv.formula1 };
}

// --- evaluation ---------------------------------------------------------------------------------

export interface ValidationDeps {
  /** valueAt reads stored values (storedValueContext). */
  base: Omit<EvalContext, "host" | "origin" | "arrayMode">;
  /** The format context item text is formatted in (the user's locale). */
  ctx: FormatContext;
}

export interface ValidationListItem {
  value: CellValue;
  text: string;
}

export interface ValidationList {
  items: ValidationListItem[];
  unavailable?: true;
}

/** A rule formula's result at `at`, or null when it does not parse or
 * cannot be computed here (an unknown name until SPEC-133, an unsupported
 * function …). `overlay` answers one cell's value instead of the model's. */
function evaluateAt(
  sheet: Sheet,
  dv: DataValidation,
  formula: string,
  at: CellRef,
  deps: ValidationDeps,
  overlay?: { at: CellRef; value: CellValue },
): { value: ReturnType<typeof evaluate>; ctx: EvalContext } | null {
  const parsed = parseFormulaAt(formula, validationOrigin(dv));
  if (parsed instanceof FormulaSyntaxError || formulaFacts(parsed.node).reason !== undefined) {
    return null;
  }
  const base = deps.base;
  let ctx: EvalContext = { ...base, host: { sheetId: sheet.id, row: at.row, col: at.col }, origin: parsed.origin, arrayMode: false };
  if (overlay) {
    // The candidate answers its cell both to a direct read and to a range
    // walk over the stored cells (which consults the model's rows).
    const overlaid = overlaidSheet(sheet, overlay.at, overlay.value);
    const model: SpreadsheetModel = { ...base.model, sheets: base.model.sheets.map((each) => (each.id === sheet.id ? overlaid : each)) };
    const valueAt: EvalContext["valueAt"] = (sheetId, row, col) =>
      sheetId === sheet.id && row === overlay.at.row && col === overlay.at.col ? fromCellValue(overlay.value) : base.valueAt(sheetId, row, col);
    ctx = { ...ctx, model, valueAt };
  }
  try {
    return { value: evaluate(parsed.node, ctx), ctx };
  } catch (error) {
    if (error instanceof UncomputableError) {
      return null;
    }
    throw error;
  }
}

/** `sheet` with `value` stored at `at` (rows copied on write; the original
 * untouched) — `custom`'s overlay. */
function overlaidSheet(sheet: Sheet, at: CellRef, value: CellValue): Sheet {
  const rows = sheet.rows.slice();
  let index = rows.findIndex((row) => row.index >= at.row);
  if (index < 0) {
    index = rows.length;
  }
  const row = rows[index];
  if (row && row.index === at.row) {
    const cells = row.cells.filter((cell) => cell.col !== at.col);
    cells.push({ col: at.col, value });
    rows[index] = { ...row, cells: cells.toSorted((a, b) => a.col - b.col) };
  } else {
    rows.splice(index, 0, { index: at.row, cells: [{ col: at.col, value }] });
  }
  return { ...sheet, rows };
}

/** A bound's number at `at`: blank = 0; null when the bound cannot be
 * evaluated, is an error or is not a number (the entry is then accepted). */
function boundAt(sheet: Sheet, dv: DataValidation, formula: string | undefined, at: CellRef, deps: ValidationDeps): number | null {
  if (formula === undefined || formula.trim() === "") {
    return null;
  }
  const result = evaluateAt(sheet, dv, formula, at, deps);
  if (!result) {
    return null;
  }
  const scalar = toScalar(result.ctx, result.value);
  if (isEmpty(scalar)) {
    return 0;
  }
  return typeof scalar === "number" ? scalar : null;
}

/** The text of a source cell's value through its own number format. */
function itemText(model: SpreadsheetModel, sheetId: string, row: number, col: number, value: CellValue, deps: ValidationDeps): string {
  if (value === null) {
    return "";
  }
  const sheet = model.sheets.find((each) => each.id === sheetId);
  const style = sheet ? cellStyle(model, effectiveStyleId(sheet, row, col)) : undefined;
  const code = style ? formatCodeOf(style, deps.ctx.locale) : undefined;
  return plainText(formatValue(value, code, deps.ctx));
}

/** A scalar's text as an item (an array element: General, no cell format). */
function scalarText(value: CellValue, deps: ValidationDeps): string {
  return value === null ? "" : plainText(formatValue(value, undefined, deps.ctx));
}

/** The items a `list` rule offers at `at`, in source order (see "List
 * sources"): a literal's items — each item's value what typing its text
 * would enter (SPEC-122's parseInput in `deps.ctx`'s locale, so `10` is the
 * number 10) — or a one-row or one-column range or array's cells (a range
 * cell's value as stored, its text through its own number format; blank
 * cells as empty items; duplicates kept), capped at VALIDATION_LIST_MAX; or
 * `unavailable` for a 2-D result, an error or an uncomputable source. A
 * `sort-ascending` ODS list is sorted (sortItemsAscending). */
export function listItems(
  model: SpreadsheetModel,
  sheet: Sheet,
  dv: DataValidation,
  at: CellRef,
  deps: ValidationDeps,
): ValidationList {
  const source = listSourceOf(dv);
  if (!source) {
    return { items: [], unavailable: true };
  }
  let items: ValidationListItem[];
  if (source.kind === "literal") {
    items = source.items.slice(0, VALIDATION_LIST_MAX).map((text) => ({ value: literalItemValue(text, deps.ctx), text }));
  } else {
    const result = evaluateAt(sheet, dv, source.formula, at, deps);
    if (!result) {
      return { items: [], unavailable: true };
    }
    const value = result.value;
    if (isRange(value)) {
      if (value.areas.length !== 1) {
        return { items: [], unavailable: true };
      }
      const area = value.areas[0]!;
      if (area.top !== area.bottom && area.left !== area.right) {
        return { items: [], unavailable: true };
      }
      const count = Math.min(VALIDATION_LIST_MAX, (area.bottom - area.top + 1) * (area.right - area.left + 1));
      const values: CellValue[] = Array.from<CellValue>({ length: count }).fill(null);
      const vertical = area.left === area.right;
      const capped = vertical
        ? { ...area, bottom: Math.min(area.bottom, area.top + count - 1) }
        : { ...area, right: Math.min(area.right, area.left + count - 1) };
      forEachStoredCell(result.ctx, capped, (scalar, row, col) => {
        values[vertical ? row - area.top : col - area.left] = isEmpty(scalar) ? null : toCellValue(scalar);
      });
      items = values.map((cellValue, index) => {
        const row = vertical ? area.top + index : area.top;
        const col = vertical ? area.left : area.left + index;
        return { value: cellValue, text: itemText(model, area.sheetId, row, col, cellValue, deps) };
      });
    } else if (isArray(value)) {
      if (value.rows !== 1 && value.cols !== 1) {
        return { items: [], unavailable: true };
      }
      items = value.data.slice(0, VALIDATION_LIST_MAX).map((scalar) => {
        const cellValue = isEmpty(scalar) ? null : toCellValue(scalar);
        return { value: cellValue, text: scalarText(cellValue, deps) };
      });
    } else if (isError(value)) {
      return { items: [], unavailable: true };
    } else {
      const cellValue = isEmpty(value) ? null : toCellValue(value);
      items = [{ value: cellValue, text: scalarText(cellValue, deps) }];
    }
  }
  if (dv.odsListOrder === "sort-ascending") {
    items = sortItemsAscending(items, deps.ctx.locale.tag);
  }
  return { items };
}

/** What typing a literal item's text enters (SPEC-122's parseInput). */
export function literalItemValue(text: string, ctx: Pick<FormatContext, "locale" | "dateSystem">): CellValue {
  return text === "" ? null : parseInput(text, ctx.locale, ctx.dateSystem).value;
}

/** LibreOffice's sort-ascending list: numbers numerically, before text
 * ordered by the locale's collator; blanks last. */
export function sortItemsAscending(items: readonly ValidationListItem[], localeTag: string): ValidationListItem[] {
  const collator = new Intl.Collator(localeTag);
  const rank = (item: ValidationListItem): number => (typeof item.value === "number" ? 0 : item.value === null ? 2 : 1);
  return items
    .map((item, index) => ({ item, index }))
    .sort((a, b) => {
      const ra = rank(a.item);
      const rb = rank(b.item);
      if (ra !== rb) {
        return ra - rb;
      }
      if (ra === 0) {
        return (a.item.value as number) - (b.item.value as number) || a.index - b.index;
      }
      return collator.compare(a.item.text, b.item.text) || a.index - b.index;
    })
    .map(({ item }) => item);
}

/** A value's text as validation reads it: a string as is, a number in
 * General's invariant text, a boolean TRUE/FALSE, an error its code. */
export function validationText(value: CellValue): string {
  if (value === null) {
    return "";
  }
  if (typeof value === "string") {
    return value;
  }
  if (typeof value === "number") {
    return numberToText(value);
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  return value.error;
}

/** Whether a candidate matches a list item: both numbers and numerically
 * equal (15 significant digits), or texts equal case-insensitively. */
export function matchesItem(candidate: CellValue, item: ValidationListItem): boolean {
  if (typeof candidate === "number" && typeof item.value === "number") {
    return roundTo15(candidate) === roundTo15(item.value);
  }
  const text = validationText(candidate).toLowerCase();
  return text === item.text.toLowerCase() || (item.value !== null && typeof item.value !== "number" && text === validationText(item.value).toLowerCase());
}

/** The item a cell's value selects in its dropdown (-1: none). */
export function itemIndexOf(items: readonly ValidationListItem[], value: CellValue): number {
  if (value === null || value === "") {
    return -1;
  }
  return items.findIndex((item) => item.value !== null && matchesItem(value, item));
}

function satisfies(value: number, operator: ValidationOperator, a: number, b: number | null): boolean {
  const v = roundTo15(value);
  const x = roundTo15(a);
  switch (operator) {
    case "between":
    case "notBetween": {
      if (b === null) {
        return true;
      }
      const y = roundTo15(b);
      const inside = v >= Math.min(x, y) && v <= Math.max(x, y);
      return operator === "between" ? inside : !inside;
    }
    case "equal":
      return v === x;
    case "notEqual":
      return v !== x;
    case "greaterThan":
      return v > x;
    case "lessThan":
      return v < x;
    case "greaterThanOrEqual":
      return v >= x;
    case "lessThanOrEqual":
      return v <= x;
  }
}

export type ValidationVerdict = { ok: true } | { ok: false; rule: DataValidation };

/** Whether `candidate` may be entered at `at` (see "Validation semantics"):
 * `{ ok: true }`, or the verdict naming the rule it breaks. A cell no rule
 * covers is always ok. */
export function validateValue(model: SpreadsheetModel, sheet: Sheet, at: CellRef, candidate: CellValue, deps: ValidationDeps): ValidationVerdict {
  const rule = validationAt(sheet, at);
  if (!rule || rule.type === "any" || ruleAccepts(model, sheet, rule, at, candidate, deps)) {
    return { ok: true };
  }
  return { ok: false, rule };
}

function ruleAccepts(model: SpreadsheetModel, sheet: Sheet, dv: DataValidation, at: CellRef, v: CellValue, deps: ValidationDeps): boolean {
  if (v === null || v === "") {
    return dv.allowBlank;
  }
  switch (dv.type) {
    case "any":
      return true;
    case "whole":
    case "decimal":
    case "date":
    case "time": {
      if (typeof v !== "number") {
        return false;
      }
      if (dv.type === "whole" && v !== Math.trunc(v)) {
        return false;
      }
      return boundsAccept(sheet, dv, at, v, deps);
    }
    case "textLength":
      return boundsAccept(sheet, dv, at, validationText(v).length, deps);
    case "list": {
      const list = listItems(model, sheet, dv, at, deps);
      return list.unavailable === true || list.items.some((item) => matchesItem(v, item));
    }
    case "custom": {
      if (dv.formula1 === undefined) {
        return true;
      }
      const result = evaluateAt(sheet, dv, dv.formula1, at, deps, { at, value: v });
      if (!result) {
        return true;
      }
      const scalar: ScalarValue = toScalar(result.ctx, result.value);
      return scalar === true || (typeof scalar === "number" && scalar !== 0);
    }
  }
}

function boundsAccept(sheet: Sheet, dv: DataValidation, at: CellRef, v: number, deps: ValidationDeps): boolean {
  const a = boundAt(sheet, dv, dv.formula1, at, deps);
  if (a === null) {
    return true;
  }
  if (isRangeOperator(dv.operator)) {
    const b = boundAt(sheet, dv, dv.formula2, at, deps);
    if (b === null) {
      return true;
    }
    return satisfies(v, dv.operator, a, b);
  }
  return satisfies(v, dv.operator, a, null);
}

// --- placement and identity --------------------------------------------------------------------

/** Whether a formula references a sheet other than `hostSheetName` (an
 * unqualified reference is the host's). */
export function referencesOtherSheet(formula: string, hostSheetName: string): boolean {
  const host = hostSheetName.toUpperCase();
  return tokenize(formula).some((token) => {
    const sheet = token.ref?.sheet;
    return sheet !== undefined && (sheet.toUpperCase() !== host || token.ref?.sheetEnd !== undefined);
  });
}

/** Whether an xlsx writer puts a rule built now into the x14 ext (SPEC-130
 * "x14"): one whose formulas reference another sheet, as Excel does. */
export function writesToX14(dv: Pick<DataValidation, "formula1" | "formula2">, hostSheetName: string): boolean {
  return [dv.formula1, dv.formula2].some((formula) => formula !== undefined && referencesOtherSheet(formula, hostSheetName));
}

/** `dv` with its `x14` flag set as `format`'s writer will place it (only
 * xlsx workbooks have an x14 ext), so a save → reopen reads the same flag. */
export function withX14Placement<T extends Pick<DataValidation, "formula1" | "formula2" | "x14">>(dv: T, format: SpreadsheetModel["format"], hostSheetName: string): T {
  const x14 = format === "xlsx" && writesToX14(dv, hostSheetName);
  if (x14 === (dv.x14 === true)) {
    return dv;
  }
  if (x14) {
    return { ...dv, x14: true };
  }
  const { x14: _x14, ...rest } = dv;
  return rest as T;
}

/** `dv` re-anchored from origin `from` to origin `to`: its relative formulas
 * shifted by (to − from), so every covered cell keeps its meaning. */
export function reanchorValidation<T extends Pick<DataValidation, "formula1" | "formula2">>(dv: T, from: CellRef, to: CellRef): T {
  const dRow = to.row - from.row;
  const dCol = to.col - from.col;
  if (dRow === 0 && dCol === 0) {
    return dv;
  }
  const next = { ...dv };
  if (dv.formula1 !== undefined) {
    next.formula1 = shiftFormula(dv.formula1, dRow, dCol);
  }
  if (dv.formula2 !== undefined) {
    next.formula2 = shiftFormula(dv.formula2, dRow, dCol);
  }
  return next;
}

/** A rule without its source text. */
export function withoutValidationSource<T extends Pick<DataValidation, "sourceXml">>(dv: T): T {
  if (dv.sourceXml === undefined) {
    return dv;
  }
  const { sourceXml: _source, ...rest } = dv;
  return rest as T;
}

/** A rule's settings — what "identical" compares (everything but `id`,
 * `ranges`, `sourceXml` and `odsName`), in a fixed key order. */
export function validationSettingsOf(dv: ValidationSettings | DataValidation): ValidationSettings {
  const settings: ValidationSettings = {
    type: dv.type,
    operator: dv.operator,
    allowBlank: dv.allowBlank,
    showDropdown: dv.showDropdown,
    input: { show: dv.input.show, title: dv.input.title, text: dv.input.text },
    error: { show: dv.error.show, style: dv.error.style, title: dv.error.title, text: dv.error.text },
  };
  if (dv.formula1 !== undefined) {
    settings.formula1 = dv.formula1;
  }
  if (dv.formula2 !== undefined) {
    settings.formula2 = dv.formula2;
  }
  if (dv.x14) {
    settings.x14 = true;
  }
  if (dv.odsListOrder !== undefined) {
    settings.odsListOrder = dv.odsListOrder;
  }
  return settings;
}

/** Whether two rules' settings are identical (see validationSettingsOf). */
export function sameValidationSettings(a: ValidationSettings | DataValidation, b: ValidationSettings | DataValidation): boolean {
  return JSON.stringify(validationSettingsOf(a)) === JSON.stringify(validationSettingsOf(b));
}

/** Whether a rule says anything — "Any value" with no input message is no
 * rule at all (the dialog's OK then removes the rules). */
export function isEmptyValidation(settings: ValidationSettings): boolean {
  return settings.type === "any" && (!settings.input.show || (settings.input.title === "" && settings.input.text === ""));
}

// --- commands ------------------------------------------------------------------------------------

function sheetById(model: SpreadsheetModel, sheetId: string): Sheet {
  const sheet = model.sheets.find((each) => each.id === sheetId);
  if (!sheet) {
    throw new Error(`No sheet ${sheetId}`);
  }
  return sheet;
}

function sameRanges(a: readonly GridRange[], b: readonly GridRange[]): boolean {
  return (
    a.length === b.length &&
    a.every((range, i) => range.top === b[i]!.top && range.left === b[i]!.left && range.bottom === b[i]!.bottom && range.right === b[i]!.right)
  );
}

/** `rules` with `cut` removed from every rule's ranges: a rule that lost its
 * origin re-anchored onto its new one, one left with no range removed, a
 * touched one without its source. The same array when `cut` misses all. */
export function subtractFromValidations(rules: readonly DataValidation[], cut: readonly GridRange[]): DataValidation[] {
  let changed = false;
  const next: DataValidation[] = [];
  for (const dv of rules) {
    const ranges = subtractRanges(dv.ranges, cut);
    if (sameRanges(ranges, dv.ranges)) {
      next.push(dv);
      continue;
    }
    changed = true;
    if (ranges.length === 0) {
      continue;
    }
    next.push(withoutValidationSource(reanchorValidation({ ...dv, ranges }, validationOrigin(dv), validationOrigin({ ranges }))));
  }
  return changed ? next : (rules as DataValidation[]);
}

function commitValidations(model: SpreadsheetModel, sheetId: string, validations: DataValidation[]): void {
  model.sheets = model.sheets.map((sheet) => (sheet.id === sheetId ? { ...sheet, validations, dirty: true } : sheet));
}

/** Sets `rule` (formulas relative to `ranges[0]`'s top-left) on `ranges`
 * (see the module header): the target is subtracted from every other rule,
 * then added to an existing rule whose settings are identical (its formulas
 * compared re-anchored onto that rule's origin), or else a new rule over the
 * target is appended. Returns the id of the rule now covering the target. */
export function setValidation(model: SpreadsheetModel, sheetId: string, ranges: readonly GridRange[], rule: ValidationSettings): string {
  if (ranges.length === 0) {
    throw new Error("A data validation rule needs a range");
  }
  const sheet = sheetById(model, sheetId);
  const target = ranges.map((range) => ({ ...range }));
  const origin = validationOrigin({ ranges: target });
  const settings = withX14Placement(validationSettingsOf(rule), model.format, sheet.name);
  const rest = subtractFromValidations(sheetValidations(sheet), target);
  const at = rest.findIndex((dv) => sameValidationSettings(dv, reanchorValidation(settings, origin, validationOrigin(dv))));
  if (at >= 0) {
    const next = rest.slice();
    const merged = withoutValidationSource({ ...rest[at]!, ranges: [...rest[at]!.ranges, ...target] });
    next[at] = merged;
    commitValidations(model, sheetId, next);
    return merged.id;
  }
  const created: DataValidation = { id: mintValidationId(), ranges: target, ...structuredClone(settings) };
  commitValidations(model, sheetId, [...rest, created]);
  return created.id;
}

/** Removes every rule from `ranges` (the dialog's OK after "Clear All").
 * Returns whether anything changed (nothing is marked dirty if not). */
export function clearValidation(model: SpreadsheetModel, sheetId: string, ranges: readonly GridRange[]): boolean {
  const sheet = sheetById(model, sheetId);
  const before = sheetValidations(sheet);
  const next = subtractFromValidations(before, ranges);
  if (next === before) {
    return false;
  }
  commitValidations(model, sheetId, next);
  return true;
}

/** The distinct rules covering any cell of `ranges` (in model order) and
 * whether some cell of them has none — the dialog's mixed-selection test.
 * Sparse: only the rules' own rectangles are intersected. */
export function validationsInRanges(sheet: Pick<Sheet, "validations">, ranges: readonly GridRange[]): { rules: DataValidation[]; uncovered: boolean } {
  const rules = sheetValidations(sheet);
  const hit: DataValidation[] = [];
  for (const dv of rules) {
    if (dv.ranges.some((range) => ranges.some((target) => target.top <= range.bottom && range.top <= target.bottom && target.left <= range.right && range.left <= target.right))) {
      hit.push(dv);
    }
  }
  // Uncovered: what is left of the targets once every rule's ranges are cut.
  const uncovered = subtractRanges(ranges, rules.flatMap((dv) => dv.ranges)).length > 0;
  return { rules: hit, uncovered };
}

/** Every formula a rule holds, mapped (copy-on-write: the same rule when
 * nothing changes; a changed one without its source). */
export function mapValidationFormulas(dv: DataValidation, map: (formula: string) => string): DataValidation {
  const formula1 = dv.formula1 === undefined ? undefined : map(dv.formula1);
  const formula2 = dv.formula2 === undefined ? undefined : map(dv.formula2);
  if (formula1 === dv.formula1 && formula2 === dv.formula2) {
    return dv;
  }
  const next: DataValidation = { ...dv };
  if (formula1 !== undefined) {
    next.formula1 = formula1;
  }
  if (formula2 !== undefined) {
    next.formula2 = formula2;
  }
  return withoutValidationSource(next);
}

/** The value a candidate cell holds — what the gate checks when the entry is
 * a formula (its evaluated result at the target) or a value. Uncomputable
 * formulas yield undefined (the entry is accepted). */
export function formulaCandidateValue(sheet: Sheet, at: CellRef, formula: string, deps: Pick<ValidationDeps, "base">): CellValue | undefined {
  const parsed = parseFormulaAt(formula, at);
  if (parsed instanceof FormulaSyntaxError || formulaFacts(parsed.node).reason !== undefined) {
    return undefined;
  }
  const ctx: EvalContext = { ...deps.base, host: { sheetId: sheet.id, row: at.row, col: at.col }, origin: parsed.origin, arrayMode: false };
  try {
    const scalar = toScalar(ctx, evaluate(parsed.node, ctx));
    return isEmpty(scalar) ? 0 : toCellValue(scalar);
  } catch (error) {
    if (error instanceof UncomputableError) {
      return undefined;
    }
    throw error;
  }
}

/** The cell's stored value (the dropdown selects its item). */
export function storedValueAt(sheet: Sheet, at: CellRef): CellValue {
  return getCell(sheet, at.row, at.col)?.value ?? null;
}
