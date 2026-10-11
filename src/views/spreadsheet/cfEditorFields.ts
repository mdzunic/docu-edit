// The conditional-format rule editor's fields (SPEC-129 "The rule editor") —
// THIS FILE is the normative source for how ConditionalFormatEditor's inputs
// map to and from a CfRuleDraft: Excel's six rule types (`cf-rule-type`),
// each type's fields, and the value rules — `cf-value-1`/`cf-value-2` take a
// locale value through parseInput or an `=` formula, which is parsed (an
// unparseable one is an inline error and nothing is saved). A cell-value or
// "Use a formula…" formula is typed relative to the ACTIVE cell and stored
// relative to the format's origin (shiftFormula by origin − active);
// colour-scale, data-bar and icon thresholds are absolute and evaluated at the
// origin. Framework-free and pure, so it unit-tests without a component.
import type { ModelColor } from "../../lib/core/theme";
import type { CellRef } from "../../lib/spreadsheet/address";
import {
  defaultIconThresholds,
  iconSetSize,
  type CfDataBar,
  type CfOperator,
  type CfRuleDraft,
  type CfTimePeriod,
  type CfValue,
  type IconSetName,
} from "../../lib/spreadsheet/conditionalFormats";
import { parseFormulaAt } from "../../lib/spreadsheet/formula/parse";
import { shiftFormula } from "../../lib/spreadsheet/formula/references";
import { isCellError } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { parseInput } from "../../lib/spreadsheet/numberFormat/input";
import { galleryDataBar } from "./cfPresets";

export type CfRuleType = "values" | "contain" | "rank" | "average" | "unique" | "formula";

/** `cf-rule-type`'s options, in Excel's order. */
export const CF_RULE_TYPES: readonly { id: CfRuleType; label: string }[] = [
  { id: "values", label: "Format all cells based on their values" },
  { id: "contain", label: "Format only cells that contain" },
  { id: "rank", label: "Format only top or bottom ranked values" },
  { id: "average", label: "Format only values that are above or below average" },
  { id: "unique", label: "Format only unique or duplicate values" },
  { id: "formula", label: "Use a formula to determine which cells to format" },
];

export type CfContainKind = "cellValue" | "text" | "dates" | "blanks" | "noBlanks" | "errors" | "noErrors";

/** `cf-contain-kind`'s options. */
export const CF_CONTAIN_KINDS: readonly { id: CfContainKind; label: string }[] = [
  { id: "cellValue", label: "Cell value" },
  { id: "text", label: "Specific text" },
  { id: "dates", label: "Dates occurring" },
  { id: "blanks", label: "Blanks" },
  { id: "noBlanks", label: "No blanks" },
  { id: "errors", label: "Errors" },
  { id: "noErrors", label: "No errors" },
];

/** `cf-operator`'s options for a cell value. */
export const CF_OPERATORS: readonly { id: CfOperator; label: string }[] = [
  { id: "between", label: "between" },
  { id: "notBetween", label: "not between" },
  { id: "equal", label: "equal to" },
  { id: "notEqual", label: "not equal to" },
  { id: "greaterThan", label: "greater than" },
  { id: "lessThan", label: "less than" },
  { id: "greaterThanOrEqual", label: "greater than or equal to" },
  { id: "lessThanOrEqual", label: "less than or equal to" },
];

export type CfTextOperator = "contains" | "notContains" | "beginsWith" | "endsWith";

/** `cf-operator`'s options for specific text. */
export const CF_TEXT_OPERATORS: readonly { id: CfTextOperator; label: string }[] = [
  { id: "contains", label: "containing" },
  { id: "notContains", label: "not containing" },
  { id: "beginsWith", label: "beginning with" },
  { id: "endsWith", label: "ending with" },
];

export type CfValuesStyle = "2-color" | "3-color" | "dataBar" | "iconSet";

/** `cf-style`'s options for "Format all cells based on their values". */
export const CF_VALUES_STYLES: readonly { id: CfValuesStyle; label: string }[] = [
  { id: "2-color", label: "2-Color Scale" },
  { id: "3-color", label: "3-Color Scale" },
  { id: "dataBar", label: "Data Bar" },
  { id: "iconSet", label: "Icon Sets" },
];

/** A threshold's type, as `cf-stop-type-<i>` offers it. */
export const CF_VALUE_TYPES: readonly { id: CfValue["type"]; label: string }[] = [
  { id: "min", label: "Lowest value" },
  { id: "max", label: "Highest value" },
  { id: "autoMin", label: "Automatic (lowest)" },
  { id: "autoMax", label: "Automatic (highest)" },
  { id: "num", label: "Number" },
  { id: "percent", label: "Percent" },
  { id: "percentile", label: "Percentile" },
  { id: "formula", label: "Formula" },
];

export interface CfThresholdField {
  type: CfValue["type"];
  value: string;
}

export interface CfEditorFields {
  type: CfRuleType;
  style: CfValuesStyle;
  /** Lowest, midpoint, highest — a 2-colour scale uses the first and last;
   * a data bar takes its min and max from them. */
  stops: (CfThresholdField & { color: ModelColor })[];
  bar: CfDataBar;
  iconSet: IconSetName;
  iconThresholds: (CfThresholdField & { gte: boolean })[];
  iconReverse: boolean;
  showValue: boolean;
  containKind: CfContainKind;
  operator: CfOperator;
  textOperator: CfTextOperator;
  value1: string;
  value2: string;
  period: CfTimePeriod;
  bottom: boolean;
  percent: boolean;
  rank: string;
  below: boolean;
  equalAverage: boolean;
  stdDev: 0 | 1 | 2 | 3;
  unique: boolean;
}

/** Where the values are typed and stored: the format's origin, the active
 * cell the user types relative to, and the locale/date system parseInput
 * reads a value in. */
export interface CfEditorContext {
  origin: CellRef;
  active: CellRef;
  locale: SpreadsheetLocale;
  dateSystem: "1900" | "1904";
}

function defaultBar(): CfDataBar {
  return galleryDataBar("#638EC6", true);
}

function iconThresholdFields(set: IconSetName): CfEditorFields["iconThresholds"] {
  return defaultIconThresholds(iconSetSize(set)).map(({ at, gte }) => ({ type: at.type, value: at.value ?? "", gte }));
}

/** A blank editor ("New rule…": a cell value greater than …). */
export function defaultFields(): CfEditorFields {
  return {
    type: "contain",
    style: "3-color",
    stops: [
      { type: "min", value: "", color: "#F8696B" },
      { type: "percentile", value: "50", color: "#FFEB84" },
      { type: "max", value: "", color: "#63BE7B" },
    ],
    bar: defaultBar(),
    iconSet: "3TrafficLights1",
    iconThresholds: iconThresholdFields("3TrafficLights1"),
    iconReverse: false,
    showValue: true,
    containKind: "cellValue",
    operator: "greaterThan",
    textOperator: "contains",
    value1: "",
    value2: "",
    period: "today",
    bottom: false,
    percent: false,
    rank: "10",
    below: false,
    equalAverage: false,
    stdDev: 0,
    unique: false,
  };
}

/** The icon set's thresholds reset to its defaults (a new set may hold a
 * different number of icons). */
export function withIconSet(fields: CfEditorFields, set: IconSetName): CfEditorFields {
  return { ...fields, iconSet: set, iconThresholds: iconThresholdFields(set) };
}

const NUMBER_TEXT = /^-?\d+(\.\d+)?([eE][+-]?\d+)?$/;
const STRING_LITERAL = /^"((?:[^"]|"")*)"$/;

/** A stored cellIs/threshold value as the user reads it: a number in the
 * locale's decimal separator, a string literal unquoted, anything else as
 * `=formula` (moved by `shift` — active − origin for a relative value). */
function valueText(stored: string, locale: SpreadsheetLocale, shift: { dRow: number; dCol: number }): string {
  if (NUMBER_TEXT.test(stored)) {
    return stored.replace(".", locale.decimal);
  }
  const literal = STRING_LITERAL.exec(stored);
  if (literal) {
    return literal[1]!.replace(/""/g, '"');
  }
  return `=${shiftFormula(stored, shift.dRow, shift.dCol)}`;
}

const NO_SHIFT = { dRow: 0, dCol: 0 };

function threshold(at: CfValue, locale: SpreadsheetLocale): CfThresholdField {
  return { type: at.type, value: at.value === undefined ? "" : valueText(at.value, locale, NO_SHIFT) };
}

/** The editor's fields for an existing rule (or a preset's draft). */
export function fieldsFromDraft(draft: CfRuleDraft, ctx: CfEditorContext): CfEditorFields {
  const fields = defaultFields();
  const toActive = { dRow: ctx.active.row - ctx.origin.row, dCol: ctx.active.col - ctx.origin.col };
  switch (draft.kind) {
    case "cellIs":
      return {
        ...fields,
        type: "contain",
        containKind: "cellValue",
        operator: draft.operator,
        value1: draft.formulas[0] ? valueText(draft.formulas[0], ctx.locale, toActive) : "",
        value2: draft.formulas[1] ? valueText(draft.formulas[1], ctx.locale, toActive) : "",
      };
    case "expression":
      return { ...fields, type: "formula", value1: draft.formula ? `=${shiftFormula(draft.formula, toActive.dRow, toActive.dCol)}` : "" };
    case "colorScale": {
      const [low, mid, high] = draft.stops;
      const stops = fields.stops.slice();
      const asStop = (stop: { at: CfValue; color: ModelColor } | undefined, fallback: CfEditorFields["stops"][number]) =>
        stop ? { ...threshold(stop.at, ctx.locale), color: stop.color } : fallback;
      stops[0] = asStop(low, stops[0]!);
      if (draft.stops.length >= 3) {
        stops[1] = asStop(mid, stops[1]!);
        stops[2] = asStop(high, stops[2]!);
      } else {
        stops[2] = asStop(mid, stops[2]!);
      }
      return { ...fields, type: "values", style: draft.stops.length >= 3 ? "3-color" : "2-color", stops };
    }
    case "dataBar": {
      const stops = fields.stops.slice();
      stops[0] = { ...stops[0]!, ...threshold(draft.bar.min, ctx.locale) };
      stops[2] = { ...stops[2]!, ...threshold(draft.bar.max, ctx.locale) };
      return { ...fields, type: "values", style: "dataBar", stops, bar: structuredClone(draft.bar), showValue: draft.bar.showValue };
    }
    case "iconSet":
      return {
        ...fields,
        type: "values",
        style: "iconSet",
        iconSet: draft.icons.set,
        iconThresholds: draft.icons.thresholds.map(({ at, gte }) => ({ ...threshold(at, ctx.locale), gte })),
        iconReverse: draft.icons.reverse,
        showValue: draft.icons.showValue,
      };
    case "top10":
      return { ...fields, type: "rank", bottom: draft.bottom, percent: draft.percent, rank: String(draft.rank) };
    case "aboveAverage":
      return { ...fields, type: "average", below: draft.below, equalAverage: draft.equalAverage, stdDev: draft.stdDev ?? 0 };
    case "duplicates":
      return { ...fields, type: "unique", unique: draft.unique };
    case "text":
      return { ...fields, type: "contain", containKind: "text", textOperator: draft.operator, value1: draft.text };
    case "blanks":
      return { ...fields, type: "contain", containKind: draft.not ? "noBlanks" : "blanks" };
    case "errors":
      return { ...fields, type: "contain", containKind: draft.not ? "noErrors" : "errors" };
    case "timePeriod":
      return { ...fields, type: "contain", containKind: "dates", period: draft.period };
    case "preserved":
      return fields;
  }
}

export const VALUE_REQUIRED_ERROR = "Enter a value";
export const FORMULA_ERROR = "This formula isn't valid";
export const NUMBER_ERROR = "Enter a number";
export const RANK_ERROR = "Enter a whole number from 1 to 1000";

/** Which input an error belongs to. */
export type CfFieldKey = "value1" | "value2" | "rank" | `stop-${number}` | `icon-${number}`;

export type DraftResult = { draft: CfRuleDraft } | { error: string; field: CfFieldKey };

class FieldError extends Error {
  constructor(
    message: string,
    readonly field: CfFieldKey,
  ) {
    super(message);
  }
}

/** A formula typed relative to `typedAt`, validated, then stored relative to
 * `storedAt`. */
function formulaFrom(text: string, typedAt: CellRef, storedAt: CellRef, field: CfFieldKey): string {
  const formula = text.startsWith("=") ? text.slice(1) : text;
  if (formula.trim() === "" || !("node" in parseFormulaAt(formula, typedAt))) {
    throw new FieldError(FORMULA_ERROR, field);
  }
  return shiftFormula(formula, storedAt.row - typedAt.row, storedAt.col - typedAt.col);
}

/** A cell-value comparand: a locale value through parseInput, or an
 * `=formula` typed relative to the active cell. */
function comparand(text: string, ctx: CfEditorContext, field: CfFieldKey): string {
  if (text.trim() === "") {
    throw new FieldError(VALUE_REQUIRED_ERROR, field);
  }
  const parsed = parseInput(text, ctx.locale, ctx.dateSystem);
  if (parsed.formula !== undefined) {
    return formulaFrom(parsed.formula, ctx.active, ctx.origin, field);
  }
  const value = parsed.value;
  if (typeof value === "number") {
    return String(value);
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  if (value !== null && isCellError(value)) {
    return value.error;
  }
  return `"${String(value ?? "").replace(/"/g, '""')}"`;
}

/** A threshold value object: min/max/auto carry no value; a number,
 * percent or percentile must read as a number; a formula is absolute,
 * evaluated at the origin. */
function thresholdValue(field: CfThresholdField, ctx: CfEditorContext, key: CfFieldKey): CfValue {
  if (field.type === "min" || field.type === "max" || field.type === "autoMin" || field.type === "autoMax") {
    return { type: field.type };
  }
  if (field.type === "formula") {
    return { type: "formula", value: formulaFrom(field.value, ctx.origin, ctx.origin, key) };
  }
  const parsed = parseInput(field.value.trim(), ctx.locale, ctx.dateSystem);
  if (typeof parsed.value !== "number") {
    throw new FieldError(NUMBER_ERROR, key);
  }
  return { type: field.type, value: String(parsed.value) };
}

function valuesDraft(fields: CfEditorFields, ctx: CfEditorContext): CfRuleDraft {
  const stop = (index: number) => ({ at: thresholdValue(fields.stops[index]!, ctx, `stop-${index}`), color: fields.stops[index]!.color });
  switch (fields.style) {
    case "2-color":
      return { kind: "colorScale", stops: [stop(0), stop(2)] };
    case "3-color":
      return { kind: "colorScale", stops: [stop(0), stop(1), stop(2)] };
    case "dataBar":
      return {
        kind: "dataBar",
        bar: { ...fields.bar, min: stop(0).at, max: stop(2).at, showValue: fields.showValue },
      };
    case "iconSet":
      return {
        kind: "iconSet",
        icons: {
          set: fields.iconSet,
          thresholds: fields.iconThresholds.map((each, index) => ({ at: thresholdValue(each, ctx, `icon-${index}`), gte: each.gte })),
          reverse: fields.iconReverse,
          showValue: fields.showValue,
        },
      };
  }
}

function containDraft(fields: CfEditorFields, ctx: CfEditorContext): CfRuleDraft {
  switch (fields.containKind) {
    case "cellValue": {
      const two = fields.operator === "between" || fields.operator === "notBetween";
      const formulas = [comparand(fields.value1, ctx, "value1")];
      if (two) {
        formulas.push(comparand(fields.value2, ctx, "value2"));
      }
      return { kind: "cellIs", operator: fields.operator, formulas };
    }
    case "text":
      if (fields.value1 === "") {
        throw new FieldError(VALUE_REQUIRED_ERROR, "value1");
      }
      return { kind: "text", operator: fields.textOperator, text: fields.value1 };
    case "dates":
      return { kind: "timePeriod", period: fields.period };
    case "blanks":
    case "noBlanks":
      return { kind: "blanks", not: fields.containKind === "noBlanks" };
    case "errors":
    case "noErrors":
      return { kind: "errors", not: fields.containKind === "noErrors" };
  }
}

/** The draft the fields describe (formatting kinds without their `dxfId`,
 * which the manager interns on save), or the first field in error. */
export function draftFromFields(fields: CfEditorFields, ctx: CfEditorContext): DraftResult {
  try {
    switch (fields.type) {
      case "values":
        return { draft: valuesDraft(fields, ctx) };
      case "contain":
        return { draft: containDraft(fields, ctx) };
      case "rank": {
        const rank = Number(fields.rank.trim());
        if (!Number.isInteger(rank) || rank < 1 || rank > 1000) {
          throw new FieldError(RANK_ERROR, "rank");
        }
        return { draft: { kind: "top10", bottom: fields.bottom, percent: fields.percent, rank } };
      }
      case "average":
        return {
          draft: { kind: "aboveAverage", below: fields.below, equalAverage: fields.equalAverage, ...(fields.stdDev ? { stdDev: fields.stdDev } : {}) },
        };
      case "unique":
        return { draft: { kind: "duplicates", unique: fields.unique } };
      case "formula":
        return { draft: { kind: "expression", formula: formulaFrom(fields.value1, ctx.active, ctx.origin, "value1") } };
    }
  } catch (error) {
    if (error instanceof FieldError) {
      return { error: error.message, field: error.field };
    }
    throw error;
  }
}

/** Whether the fields describe a formatting rule (one that takes a `cf-format`
 * dxf) rather than a colour scale, data bar or icon set. */
export function takesFormat(fields: Pick<CfEditorFields, "type">): boolean {
  return fields.type !== "values";
}
