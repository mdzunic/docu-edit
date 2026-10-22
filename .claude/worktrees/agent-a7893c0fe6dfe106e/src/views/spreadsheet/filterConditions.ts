// The filter popover's conditions (SPEC-128 "UI"): the submenu — "Text
// Filters", "Number Filters" or "Date Filters" by the column's majority type
// (filter.ts columnValueKind) with Excel's items — and the two dialogs it
// opens. Framework-free, so the popover (useFilterPopover) and the dialogs
// (FilterCustomDialog, FilterTop10Dialog) only bind what this returns.
//
// - Each "…" item opens the custom dialog pre-set (Between… as ≥ and ≤ joined
//   by And); Top 10… opens the Top 10 dialog; Above/Below Average and the
//   date items apply a `dynamic` criterion at once.
// - The custom dialog's operators are Excel's twelve. A value typed into it
//   goes through SPEC-122's parseInput in the user's locale: under a
//   comparison a number or date is stored as invariant number text
//   (String(n)), anything else as typed; the text operators store text with
//   their wildcards — Begins With x → equal "x*", Ends With → "*x",
//   Contains → "*x*", and the negations as notEqual — and `*` and `?` typed
//   by the user stay wildcards (Excel).
import type { DynamicFilterType, FilterCondition, FilterCriteria } from "../../lib/spreadsheet/filter";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { parseInput } from "../../lib/spreadsheet/numberFormat/input";

export type ColumnKind = "text" | "number" | "date";

export type CustomOperator =
  | "equals"
  | "notEquals"
  | "greaterThan"
  | "greaterThanOrEqual"
  | "lessThan"
  | "lessThanOrEqual"
  | "beginsWith"
  | "notBeginsWith"
  | "endsWith"
  | "notEndsWith"
  | "contains"
  | "notContains";

/** What the custom dialog holds: two conditions, the second optional (""). */
export interface CustomDraft {
  op1: CustomOperator;
  value1: string;
  join: "and" | "or";
  op2: CustomOperator | "";
  value2: string;
}

export type ConditionAction = { kind: "custom"; draft: CustomDraft } | { kind: "top10" } | { kind: "dynamic"; type: DynamicFilterType };

export interface ConditionItem {
  id: string;
  label: string;
  action: ConditionAction;
}

export const CONDITIONS_LABELS: Readonly<Record<ColumnKind, string>> = {
  text: "Text Filters",
  number: "Number Filters",
  date: "Date Filters",
};

export const CUSTOM_OPERATORS: readonly CustomOperator[] = [
  "equals",
  "notEquals",
  "greaterThan",
  "greaterThanOrEqual",
  "lessThan",
  "lessThanOrEqual",
  "beginsWith",
  "notBeginsWith",
  "endsWith",
  "notEndsWith",
  "contains",
  "notContains",
];

const OPERATOR_LABELS: Readonly<Record<CustomOperator, string>> = {
  equals: "equals",
  notEquals: "does not equal",
  greaterThan: "is greater than",
  greaterThanOrEqual: "is greater than or equal to",
  lessThan: "is less than",
  lessThanOrEqual: "is less than or equal to",
  beginsWith: "begins with",
  notBeginsWith: "does not begin with",
  endsWith: "ends with",
  notEndsWith: "does not end with",
  contains: "contains",
  notContains: "does not contain",
};

const DATE_OPERATOR_LABELS: Partial<Record<CustomOperator, string>> = {
  greaterThan: "is after",
  greaterThanOrEqual: "is after or equal to",
  lessThan: "is before",
  lessThanOrEqual: "is before or equal to",
};

/** An operator's label in the custom dialog ("is before" on a date column). */
export function customOperatorLabel(op: CustomOperator, kind: ColumnKind): string {
  return (kind === "date" ? DATE_OPERATOR_LABELS[op] : undefined) ?? OPERATOR_LABELS[op];
}

const draft = (op1: CustomOperator, op2: CustomOperator | "" = ""): CustomDraft => ({ op1, value1: "", join: "and", op2, value2: "" });
const custom = (id: string, label: string, op1: CustomOperator, op2: CustomOperator | "" = ""): ConditionItem => ({
  id,
  label,
  action: { kind: "custom", draft: draft(op1, op2) },
});
const dynamic = (label: string, type: DynamicFilterType): ConditionItem => ({ id: type, label, action: { kind: "dynamic", type } });
const CUSTOM_FILTER = custom("custom", "Custom Filter…", "equals");

const ITEMS: Readonly<Record<ColumnKind, readonly ConditionItem[]>> = {
  text: [
    custom("equals", "Equals…", "equals"),
    custom("notEquals", "Does Not Equal…", "notEquals"),
    custom("beginsWith", "Begins With…", "beginsWith"),
    custom("endsWith", "Ends With…", "endsWith"),
    custom("contains", "Contains…", "contains"),
    custom("notContains", "Does Not Contain…", "notContains"),
    CUSTOM_FILTER,
  ],
  number: [
    custom("equals", "Equals…", "equals"),
    custom("notEquals", "Does Not Equal…", "notEquals"),
    custom("greaterThan", "Greater Than…", "greaterThan"),
    custom("greaterThanOrEqual", "Greater Than Or Equal To…", "greaterThanOrEqual"),
    custom("lessThan", "Less Than…", "lessThan"),
    custom("lessThanOrEqual", "Less Than Or Equal To…", "lessThanOrEqual"),
    custom("between", "Between…", "greaterThanOrEqual", "lessThanOrEqual"),
    { id: "top10", label: "Top 10…", action: { kind: "top10" } },
    dynamic("Above Average", "aboveAverage"),
    dynamic("Below Average", "belowAverage"),
    CUSTOM_FILTER,
  ],
  date: [
    custom("equals", "Equals…", "equals"),
    custom("before", "Before…", "lessThan"),
    custom("after", "After…", "greaterThan"),
    custom("between", "Between…", "greaterThanOrEqual", "lessThanOrEqual"),
    dynamic("Tomorrow", "tomorrow"),
    dynamic("Today", "today"),
    dynamic("Yesterday", "yesterday"),
    dynamic("Next Week", "nextWeek"),
    dynamic("This Week", "thisWeek"),
    dynamic("Last Week", "lastWeek"),
    dynamic("Next Month", "nextMonth"),
    dynamic("This Month", "thisMonth"),
    dynamic("Last Month", "lastMonth"),
    dynamic("Next Year", "nextYear"),
    dynamic("This Year", "thisYear"),
    dynamic("Last Year", "lastYear"),
    dynamic("Year to Date", "yearToDate"),
    CUSTOM_FILTER,
  ],
};

/** The conditions submenu's items for a column of `kind`, in Excel's order. */
export function conditionItems(kind: ColumnKind): readonly ConditionItem[] {
  return ITEMS[kind];
}

/** The text operators: the wildcard pattern around the typed text. */
const PATTERNS: Partial<Record<CustomOperator, { operator: "equal" | "notEqual"; wrap: (x: string) => string }>> = {
  beginsWith: { operator: "equal", wrap: (x) => `${x}*` },
  notBeginsWith: { operator: "notEqual", wrap: (x) => `${x}*` },
  endsWith: { operator: "equal", wrap: (x) => `*${x}` },
  notEndsWith: { operator: "notEqual", wrap: (x) => `*${x}` },
  contains: { operator: "equal", wrap: (x) => `*${x}*` },
  notContains: { operator: "notEqual", wrap: (x) => `*${x}*` },
};

const COMPARISONS: Partial<Record<CustomOperator, FilterCondition["operator"]>> = {
  equals: "equal",
  notEquals: "notEqual",
  greaterThan: "greaterThan",
  greaterThanOrEqual: "greaterThanOrEqual",
  lessThan: "lessThan",
  lessThanOrEqual: "lessThanOrEqual",
};

/** One typed condition as its wire form (see the module header). */
export function conditionFor(op: CustomOperator, text: string, locale: SpreadsheetLocale, dateSystem: "1900" | "1904"): FilterCondition {
  const pattern = PATTERNS[op];
  if (pattern) {
    return { operator: pattern.operator, value: pattern.wrap(text) };
  }
  const parsed = parseInput(text.trim(), locale, dateSystem);
  const value = typeof parsed.value === "number" && parsed.formula === undefined ? String(parsed.value) : text;
  return { operator: COMPARISONS[op] ?? "equal", value };
}

/** The custom dialog's OK: a `custom` criterion (null while the first value
 * is empty). A second condition counts only with an operator and a value. */
export function customCriteria(d: CustomDraft, locale: SpreadsheetLocale, dateSystem: "1900" | "1904"): FilterCriteria | null {
  if (d.value1 === "") {
    return null;
  }
  const first = conditionFor(d.op1, d.value1, locale, dateSystem);
  if (d.op2 === "" || d.value2 === "") {
    return { kind: "custom", join: "and", conditions: [first] };
  }
  return { kind: "custom", join: d.join, conditions: [first, conditionFor(d.op2, d.value2, locale, dateSystem)] };
}

/** A wire value as the dialog shows it: an invariant number in the
 * locale's decimal separator, anything else as stored. */
function shownValue(value: string, locale: SpreadsheetLocale): string {
  return /^-?\d+(\.\d+)?$/.test(value) ? value.replace(".", locale.decimal) : value;
}

/** Text with no wildcard or escape in it. */
const plain = (text: string): boolean => text.length > 0 && !/[*?~]/.test(text);

/** The operator and text a stored condition reads back as in the dialog. */
function conditionDraft(condition: FilterCondition, locale: SpreadsheetLocale): { op: CustomOperator; value: string } {
  const { operator, value } = condition;
  if (operator === "equal" || operator === "notEqual") {
    const negated = operator === "notEqual";
    const inner = (from: number, to: number): string => value.slice(from, value.length - to);
    if (value.length > 2 && value.startsWith("*") && value.endsWith("*") && plain(inner(1, 1))) {
      return { op: negated ? "notContains" : "contains", value: inner(1, 1) };
    }
    if (value.endsWith("*") && plain(inner(0, 1))) {
      return { op: negated ? "notBeginsWith" : "beginsWith", value: inner(0, 1) };
    }
    if (value.startsWith("*") && plain(inner(1, 0))) {
      return { op: negated ? "notEndsWith" : "endsWith", value: inner(1, 0) };
    }
  }
  const op = (Object.keys(COMPARISONS) as CustomOperator[]).find((key) => COMPARISONS[key] === operator) ?? "equals";
  return { op, value: shownValue(value, locale) };
}

/** Custom Filter… opens on the column's current `custom` criterion, if any. */
export function customDraftOf(criteria: FilterCriteria | undefined, locale: SpreadsheetLocale): CustomDraft | null {
  if (criteria?.kind !== "custom") {
    return null;
  }
  const [first, second] = criteria.conditions;
  const one = conditionDraft(first, locale);
  const two = second ? conditionDraft(second, locale) : null;
  return { op1: one.op, value1: one.value, join: criteria.join, op2: two?.op ?? "", value2: two?.value ?? "" };
}

/** The Top 10 dialog's bounds: 1–500 items or 1–100 percent. */
export const TOP10_MAX = { items: 500, percent: 100 } as const;

/** The Top 10 dialog's OK: a `top10` criterion, or null while the value is
 * not a whole number in range. */
export function top10Criteria(top: boolean, percent: boolean, text: string): FilterCriteria | null {
  const value = Number(text.trim());
  const max = percent ? TOP10_MAX.percent : TOP10_MAX.items;
  if (text.trim() === "" || !Number.isInteger(value) || value < 1 || value > max) {
    return null;
  }
  return { kind: "top10", top, percent, value };
}

/** Narrows a component's runtime-typed `draft` prop. */
export function asCustomDraft(value: unknown): CustomDraft {
  const given = (value ?? {}) as Partial<CustomDraft>;
  const op = (candidate: unknown): CustomOperator | null =>
    typeof candidate === "string" && (CUSTOM_OPERATORS as readonly string[]).includes(candidate) ? (candidate as CustomOperator) : null;
  return {
    op1: op(given.op1) ?? "equals",
    value1: typeof given.value1 === "string" ? given.value1 : "",
    join: given.join === "or" ? "or" : "and",
    op2: op(given.op2) ?? "",
    value2: typeof given.value2 === "string" ? given.value2 : "",
  };
}

/** Narrows a component's runtime-typed `kind` prop. */
export function asColumnKind(value: unknown): ColumnKind {
  return value === "number" || value === "date" ? value : "text";
}
