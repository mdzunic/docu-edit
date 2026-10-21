// The Data validation dialog's fields (SPEC-130 "The dialog") — THIS FILE is
// the normative source for how ValidationDialog's controls map to and from a
// rule's settings: the choices `validation-allow`, `validation-operator` and
// `validation-error-style` offer, the label each value field carries per
// type and operator, the message limits, the dialog's five errors (verbatim)
// and the mixed-selection notice. Framework-free and pure, so it unit-tests
// without a component.
//
// Values, both ways:
// - A rule's formulas are stored relative to its origin and SHOWN relative to
//   the ACTIVE cell (reanchorValidation), as the user would type them: a
//   literal list as its comma-joined items; a whole / decimal / text length
//   number as a number in the user's locale; a date or time bound that is a
//   plain number as a date or time in the user's locale (SPEC-118's
//   editText); anything else — a reference, a name, a formula — with a
//   leading "=".
// - On OK a value starting with "=" is a formula: parsed with parseFormula
//   (one that does not parse is "The formula contains an error"), stored
//   canonical (SPEC-119's printFormula) and moved from the active cell onto
//   the rule's origin — the selection's first range's top-left (shiftFormula
//   by origin − active). Anything else goes through SPEC-122's parseInput in
//   the user's locale and date system: a number, date or time is stored as
//   invariant number text (numberToText); a bound that is empty or reads as
//   anything but a number is "The formula contains an error" (the dialog's
//   five messages are its only ones). A custom Formula is always a formula,
//   its "=" optional. A list Source without "=" becomes a literal: split on
//   commas (never a locale separator), each item trimmed, formatListLiteral —
//   over LITERAL_LIST_MAX_CHARS characters is "The list source can't be
//   longer than 255 characters"; an empty Source, or one that is a reference
//   spanning more than one row and more than one column (or several sheets),
//   is "The list source must be a delimited list, or a reference to a single
//   row or column". A text length bound that is a constant must be a whole
//   number ≥ 0 ("Enter a whole number of 0 or more"). Between / not between
//   with two constant bounds refuses a Maximum below the Minimum ("The
//   Maximum must be greater than or equal to the Minimum").
// - Operator-less types store `operator: "between"` and only a list keeps an
//   unticked In-cell dropdown (the model's defaults otherwise), so a rule's
//   settings stay canonical for setValidation's identical-rule merge. A list
//   loaded with LibreOffice's sort-ascending order keeps it (the dialog has
//   no control for it) until Clear All or another type.
import { subtractRanges, type CellRef, type GridRange } from "../../lib/spreadsheet/address";
import { FormulaSyntaxError, parseFormula } from "../../lib/spreadsheet/formula/parse";
import type { FormulaNode } from "../../lib/spreadsheet/formula/ast";
import { printFormula } from "../../lib/spreadsheet/formula/print";
import type { RefPart } from "../../lib/spreadsheet/formula/references";
import { shiftFormula } from "../../lib/spreadsheet/formula/references";
import { numberToText } from "../../lib/spreadsheet/formula/values";
import type { Sheet } from "../../lib/spreadsheet/model";
import { parseInput } from "../../lib/spreadsheet/numberFormat/input";
import { editText, type FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import {
  defaultValidationSettings,
  formatListLiteral,
  isRangeOperator,
  listSourceOf,
  LITERAL_LIST_MAX_CHARS,
  reanchorValidation,
  sameValidationSettings,
  usesOperator,
  validationOrigin,
  validationsInRanges,
  type DataValidation,
  type ValidationErrorStyle,
  type ValidationOperator,
  type ValidationSettings,
  type ValidationType,
} from "../../lib/spreadsheet/validation";

/** `validation-allow`'s options, in Excel's order. */
export const VALIDATION_ALLOW_OPTIONS: readonly { id: ValidationType; label: string }[] = [
  { id: "any", label: "Any value" },
  { id: "whole", label: "Whole number" },
  { id: "decimal", label: "Decimal" },
  { id: "list", label: "List" },
  { id: "date", label: "Date" },
  { id: "time", label: "Time" },
  { id: "textLength", label: "Text length" },
  { id: "custom", label: "Custom" },
];

/** `validation-operator`'s options (whole, decimal, date, time, text length). */
export const VALIDATION_OPERATOR_OPTIONS: readonly { id: ValidationOperator; label: string }[] = [
  { id: "between", label: "between" },
  { id: "notBetween", label: "not between" },
  { id: "equal", label: "equal to" },
  { id: "notEqual", label: "not equal to" },
  { id: "greaterThan", label: "greater than" },
  { id: "lessThan", label: "less than" },
  { id: "greaterThanOrEqual", label: "greater than or equal to" },
  { id: "lessThanOrEqual", label: "less than or equal to" },
];

/** `validation-error-style`'s options. */
export const VALIDATION_ERROR_STYLE_OPTIONS: readonly { id: ValidationErrorStyle; label: string }[] = [
  { id: "stop", label: "Stop" },
  { id: "warning", label: "Warning" },
  { id: "information", label: "Information" },
];

/** The message limits (Excel's): the input message's title and text, the
 * error alert's title and text. */
export const INPUT_TITLE_MAX = 32;
export const INPUT_TEXT_MAX = 255;
export const ERROR_TITLE_MAX = 32;
export const ERROR_TEXT_MAX = 225;

// The dialog's errors, verbatim (SPEC-130 "The dialog").
export const FORMULA_ERROR = "The formula contains an error";
export const BOUNDS_ORDER_ERROR = "The Maximum must be greater than or equal to the Minimum";
export const TEXT_LENGTH_ERROR = "Enter a whole number of 0 or more";
export const LIST_TOO_LONG_ERROR = "The list source can't be longer than 255 characters";
export const LIST_SOURCE_ERROR = "The list source must be a delimited list, or a reference to a single row or column";

/** `validation-mixed-notice`. */
export const MIXED_SELECTION_NOTICE = "The selection contains cells with different validation settings — OK applies these settings to all of them";

/** What the dialog's controls hold. */
export interface ValidationFields {
  type: ValidationType;
  operator: ValidationOperator;
  /** `validation-value-1` / `validation-value-2` as typed. */
  value1: string;
  value2: string;
  allowBlank: boolean;
  showDropdown: boolean;
  inputShow: boolean;
  inputTitle: string;
  inputText: string;
  errorShow: boolean;
  errorStyle: ValidationErrorStyle;
  errorTitle: string;
  errorText: string;
  /** No control: a loaded list's LibreOffice sort order, kept on OK. */
  odsListOrder?: "sort-ascending";
}

/** Which value field an error belongs to. */
export type ValidationFieldKey = "value1" | "value2";

/** Where the values are typed and stored: the active cell the user types
 * relative to, the origin the rule is stored relative to, and the format
 * context (locale, date system) values are read and shown in. */
export interface ValidationFieldContext {
  active: CellRef;
  origin: CellRef;
  format: FormatContext;
}

/** Excel's defaults ("Clear All", or a cell without a rule). */
export function defaultValidationFields(): ValidationFields {
  return fieldsOfSettings(defaultValidationSettings(), "", "");
}

function fieldsOfSettings(settings: ValidationSettings, value1: string, value2: string): ValidationFields {
  const fields: ValidationFields = {
    type: settings.type,
    operator: settings.operator,
    value1,
    value2,
    allowBlank: settings.allowBlank,
    showDropdown: settings.showDropdown,
    inputShow: settings.input.show,
    inputTitle: settings.input.title,
    inputText: settings.input.text,
    errorShow: settings.error.show,
    errorStyle: settings.error.style,
    errorTitle: settings.error.title,
    errorText: settings.error.text,
  };
  if (settings.type === "list" && settings.odsListOrder !== undefined) {
    fields.odsListOrder = settings.odsListOrder;
  }
  return fields;
}

/** Whether two field sets are the same (the dialog's OK is then a no-op on
 * a selection that agrees). */
export function sameValidationFields(a: ValidationFields, b: ValidationFields): boolean {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]) as Set<keyof ValidationFields>;
  return [...keys].every((key) => a[key] === b[key]);
}

/** The labels of the value fields the type and operator show (null: the
 * field is hidden). */
export function valueLabels(type: ValidationType, operator: ValidationOperator): { value1: string | null; value2: string | null } {
  if (type === "list") {
    return { value1: "Source", value2: null };
  }
  if (type === "custom") {
    return { value1: "Formula", value2: null };
  }
  if (!usesOperator(type)) {
    return { value1: null, value2: null };
  }
  const [min, max] = type === "date" ? ["Start date", "End date"] : type === "time" ? ["Start time", "End time"] : ["Minimum", "Maximum"];
  switch (operator) {
    case "between":
    case "notBetween":
      return { value1: min, value2: max };
    case "equal":
    case "notEqual":
      return { value1: type === "date" ? "Date" : type === "time" ? "Time" : type === "textLength" ? "Length" : "Value", value2: null };
    // The spec's labels: "Minimum" / "Maximum" for the one-sided
    // comparisons of every type (Start/End name only a between's bounds).
    case "greaterThan":
    case "greaterThanOrEqual":
      return { value1: "Minimum", value2: null };
    case "lessThan":
    case "lessThanOrEqual":
      return { value1: "Maximum", value2: null };
  }
}

const NUMBER_TEXT = /^-?\d+(\.\d+)?([eE][+-]?\d+)?$/;

/** A stored bound as the user reads it (see the module header). */
function boundText(stored: string | undefined, type: ValidationType, format: FormatContext): string {
  if (stored === undefined || stored === "") {
    return "";
  }
  if (!NUMBER_TEXT.test(stored)) {
    return `=${stored}`;
  }
  const value = Number(stored);
  if (type === "date") {
    return editText({ value }, { ...format, style: { numberFormat: format.locale.shortDate } });
  }
  if (type === "time") {
    return editText({ value }, { ...format, style: { numberFormat: format.locale.longTime } });
  }
  return editText({ value }, format);
}

/** The dialog's fields for the active cell's rule (Excel's defaults when it
 * has none), its formulas shown relative to the active cell. */
export function fieldsFromRule(rule: DataValidation | undefined, active: CellRef, format: FormatContext): ValidationFields {
  if (!rule) {
    return defaultValidationFields();
  }
  const shown = reanchorValidation(rule, validationOrigin(rule), active);
  switch (rule.type) {
    case "any":
      return fieldsOfSettings(rule, "", "");
    case "list": {
      const source = listSourceOf(shown);
      const text = source === null ? "" : source.kind === "literal" ? source.items.join(",") : `=${source.formula}`;
      return fieldsOfSettings(rule, text, "");
    }
    case "custom":
      return fieldsOfSettings(rule, shown.formula1 ? `=${shown.formula1}` : "", "");
    default:
      return fieldsOfSettings(rule, boundText(shown.formula1, rule.type, format), boundText(shown.formula2, rule.type, format));
  }
}

/** A field's error, or the settings the fields describe. */
export type SettingsResult = { settings: ValidationSettings } | { error: string; field: ValidationFieldKey };

class FieldError extends Error {
  constructor(
    message: string,
    readonly field: ValidationFieldKey,
  ) {
    super(message);
  }
}

/** A formula body (no "=") typed at the active cell: canonical and moved
 * onto the origin. */
function formulaFrom(body: string, ctx: ValidationFieldContext, field: ValidationFieldKey): string {
  if (body.trim() === "") {
    throw new FieldError(FORMULA_ERROR, field);
  }
  let canonical: string;
  try {
    canonical = printFormula(parseFormula(body));
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      throw new FieldError(FORMULA_ERROR, field);
    }
    throw error;
  }
  return shiftFormula(canonical, ctx.origin.row - ctx.active.row, ctx.origin.col - ctx.active.col);
}

/** A bound: an `=formula`, or a constant that must read as a number. */
function boundFrom(text: string, type: ValidationType, ctx: ValidationFieldContext, field: ValidationFieldKey): { formula: string; constant?: number } {
  const value = text.trim();
  if (value.startsWith("=")) {
    return { formula: formulaFrom(value.slice(1), ctx, field) };
  }
  const parsed = value === "" ? null : parseInput(value, ctx.format.locale, ctx.format.dateSystem).value;
  if (type === "textLength") {
    if (typeof parsed !== "number" || !Number.isInteger(parsed) || parsed < 0) {
      throw new FieldError(TEXT_LENGTH_ERROR, field);
    }
  } else if (typeof parsed !== "number") {
    throw new FieldError(FORMULA_ERROR, field);
  }
  return { formula: numberToText(parsed), constant: parsed };
}

function rootOf(node: FormulaNode): FormulaNode {
  let at = node;
  while (at.kind === "paren") {
    at = at.inner;
  }
  return at;
}

/** Whether a reference Source spans one row or one column of one sheet. */
function singleRowOrColumn(formula: string): boolean {
  const parts: RefPart[] = [];
  const collect = (node: FormulaNode): boolean => {
    const at = rootOf(node);
    if (at.kind === "ref") {
      if (at.ref.sheetEnd !== undefined) {
        return false;
      }
      parts.push(at.ref.start, at.ref.end ?? at.ref.start);
      return true;
    }
    return at.kind === "binary" && at.op === ":" && collect(at.left) && collect(at.right);
  };
  if (!collect(parseFormula(formula))) {
    return false;
  }
  const same = (key: "row" | "col") => parts.every((part) => part[key] !== undefined && part[key] === parts[0]![key]);
  return same("row") || same("col");
}

/** A list Source (see the module header). */
function listSourceFrom(text: string, ctx: ValidationFieldContext): string {
  const value = text.trim();
  if (value === "" || value === "=") {
    throw new FieldError(value === "" ? LIST_SOURCE_ERROR : FORMULA_ERROR, "value1");
  }
  const literalOf = (items: string[]): string => {
    if (items.join(",").length > LITERAL_LIST_MAX_CHARS) {
      throw new FieldError(LIST_TOO_LONG_ERROR, "value1");
    }
    return formatListLiteral(items);
  };
  if (!value.startsWith("=")) {
    return literalOf(value.split(",").map((item) => item.trim()));
  }
  const formula = formulaFrom(value.slice(1), ctx, "value1");
  const source = listSourceOf({ type: "list", formula1: formula });
  if (source?.kind === "literal") {
    return literalOf(source.items);
  }
  if (source?.kind === "reference" && !singleRowOrColumn(formula)) {
    throw new FieldError(LIST_SOURCE_ERROR, "value1");
  }
  return formula;
}

/** The settings the fields describe, relative to `ctx.origin` — or the
 * first value in error. */
export function settingsFromFields(fields: ValidationFields, ctx: ValidationFieldContext): SettingsResult {
  const settings: ValidationSettings = {
    type: fields.type,
    operator: usesOperator(fields.type) ? fields.operator : "between",
    allowBlank: fields.allowBlank,
    showDropdown: fields.type === "list" ? fields.showDropdown : true,
    input: { show: fields.inputShow, title: fields.inputTitle, text: fields.inputText },
    error: { show: fields.errorShow, style: fields.errorStyle, title: fields.errorTitle, text: fields.errorText },
  };
  try {
    switch (fields.type) {
      case "any":
        break;
      case "list":
        settings.formula1 = listSourceFrom(fields.value1, ctx);
        if (fields.odsListOrder !== undefined) {
          settings.odsListOrder = fields.odsListOrder;
        }
        break;
      case "custom": {
        const text = fields.value1.trim();
        settings.formula1 = formulaFrom(text.startsWith("=") ? text.slice(1) : text, ctx, "value1");
        break;
      }
      default: {
        const first = boundFrom(fields.value1, fields.type, ctx, "value1");
        settings.formula1 = first.formula;
        if (isRangeOperator(fields.operator)) {
          const second = boundFrom(fields.value2, fields.type, ctx, "value2");
          if (first.constant !== undefined && second.constant !== undefined && second.constant < first.constant) {
            throw new FieldError(BOUNDS_ORDER_ERROR, "value2");
          }
          settings.formula2 = second.formula;
        }
      }
    }
  } catch (error) {
    if (error instanceof FieldError) {
      return { error: error.message, field: error.field };
    }
    throw error;
  }
  return { settings };
}

/** Whether the selection's cells disagree — some covered by a rule and some
 * not, or covered by rules whose settings differ (compared re-anchored onto
 * one origin, so the same rule written twice agrees). */
export function selectionDisagrees(sheet: Pick<Sheet, "validations">, ranges: readonly GridRange[]): boolean {
  const { rules, uncovered } = validationsInRanges(sheet, ranges);
  const [first, ...rest] = rules;
  if (!first) {
    return false;
  }
  if (uncovered) {
    return true;
  }
  const origin = validationOrigin(first);
  return rest.some((rule) => !sameValidationSettings(first, reanchorValidation(rule, validationOrigin(rule), origin)));
}

/** `ranges` made disjoint (each range less the ones before it), so the rule
 * a selection's OK writes never covers a cell twice; the first range's
 * top-left — the rule's origin — stays first. */
export function disjointRanges(ranges: readonly GridRange[]): GridRange[] {
  const out: GridRange[] = [];
  for (const range of ranges) {
    out.push(...subtractRanges([{ ...range }], out));
  }
  return out;
}
