// Conditional-format rule descriptions (SPEC-129 "Rule descriptions") — THIS
// FILE is the normative source for the one-line summary ConditionalFormatsPane
// shows beside each rule (`cf-rule-description`). Framework-free and pure so
// it unit-tests without mounting a component; every string here is the exact
// text the spec names.
import { ICON_SET_LABELS } from "../../lib/spreadsheet/cfIcons";
import type { CfOperator, CfRule, CfTimePeriod } from "../../lib/spreadsheet/conditionalFormats";

/** `cellIs`'s comparison operators, as the description shows them (SPEC-129:
 * "=", "≠", ">", "≥", "<", "≤" — `between`/`notBetween` read as prose
 * instead and never reach this table). */
const OPERATOR_SYMBOLS: Readonly<Record<Exclude<CfOperator, "between" | "notBetween">, string>> = {
  equal: "=",
  notEqual: "≠",
  greaterThan: ">",
  greaterThanOrEqual: "≥",
  lessThan: "<",
  lessThanOrEqual: "≤",
};

/** `timePeriod`'s periods, as "Date occurring: <label>". */
const TIME_PERIOD_LABELS: Readonly<Record<CfTimePeriod, string>> = {
  today: "Today",
  yesterday: "Yesterday",
  tomorrow: "Tomorrow",
  last7Days: "Last 7 days",
  thisWeek: "This week",
  lastWeek: "Last week",
  nextWeek: "Next week",
  thisMonth: "This month",
  lastMonth: "Last month",
  nextMonth: "Next month",
  thisYear: "This year",
  lastYear: "Last year",
  nextYear: "Next year",
};

/** `text`'s operators, as "Cell value <phrase> "…"". */
function textPhrase(operator: "contains" | "notContains" | "beginsWith" | "endsWith"): string {
  switch (operator) {
    case "contains":
      return "contains";
    case "notContains":
      return "does not contain";
    case "beginsWith":
      return "begins with";
    case "endsWith":
      return "ends with";
  }
}

/** `cellIs`'s value text, as stored (a canonical formula, relative to the
 * format's origin) — shown verbatim, the way a plain number or an absolute
 * reference reads unchanged regardless of origin. */
function describeCellIs(rule: Extract<CfRule, { kind: "cellIs" }>): string {
  const [first = "", second = ""] = rule.formulas;
  if (rule.operator === "between") {
    return `Cell value between ${first} and ${second}`;
  }
  if (rule.operator === "notBetween") {
    return `Cell value not between ${first} and ${second}`;
  }
  return `Cell value ${OPERATOR_SYMBOLS[rule.operator]} ${first}`;
}

/** The one-line summary ConditionalFormatsPane's `cf-rule-description`
 * shows for `rule` (SPEC-129 "Rule descriptions", exact strings). */
export function describeRule(rule: CfRule): string {
  switch (rule.kind) {
    case "cellIs":
      return describeCellIs(rule);
    case "expression":
      return `Formula: =${rule.formula}`;
    case "colorScale":
      return "Graded colour scale";
    case "dataBar":
      return "Data bar";
    case "iconSet":
      return `Icon set (${ICON_SET_LABELS[rule.icons.set]})`;
    case "top10":
      return `${rule.bottom ? "Bottom" : "Top"} ${rule.rank}${rule.percent ? "%" : ""}`;
    case "aboveAverage": {
      const verb = rule.below ? "Below average" : "Above average";
      return rule.stdDev ? `${verb} (${rule.stdDev} std dev)` : verb;
    }
    case "duplicates":
      return rule.unique ? "Unique values" : "Duplicate values";
    case "text":
      return `Cell value ${textPhrase(rule.operator)} "${rule.text}"`;
    case "blanks":
      return rule.not ? "No blanks" : "Blanks";
    case "errors":
      return rule.not ? "No errors" : "Errors";
    case "timePeriod":
      return `Date occurring: ${TIME_PERIOD_LABELS[rule.period]}`;
    case "preserved":
      return "Rule kept from the file";
  }
}
