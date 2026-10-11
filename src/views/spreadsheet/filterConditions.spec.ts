// SPEC-128's filter conditions: the submenus by column type, the custom
// dialog's values through parseInput (numbers and dates as invariant number
// text, text with its operator's wildcards, typed wildcards kept), the
// dialog read back from a stored criterion, and the Top 10 bounds.
import { describe, expect, it } from "vitest";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { conditionFor, conditionItems, CONDITIONS_LABELS, customCriteria, customDraftOf, customOperatorLabel, top10Criteria } from "./filterConditions";

const US = SPREADSHEET_LOCALES["en-US"];
const DE = SPREADSHEET_LOCALES["de-DE"];

describe("SPEC-128 filter conditions", () => {
  it("lists Excel's items per column type under Text, Number or Date Filters", () => {
    expect(CONDITIONS_LABELS).toEqual({ text: "Text Filters", number: "Number Filters", date: "Date Filters" });
    expect(conditionItems("text").map((item) => item.label)).toEqual([
      "Equals…", "Does Not Equal…", "Begins With…", "Ends With…", "Contains…", "Does Not Contain…", "Custom Filter…",
    ]);
    expect(conditionItems("number").map((item) => item.label)).toEqual([
      "Equals…", "Does Not Equal…", "Greater Than…", "Greater Than Or Equal To…", "Less Than…", "Less Than Or Equal To…",
      "Between…", "Top 10…", "Above Average", "Below Average", "Custom Filter…",
    ]);
    expect(conditionItems("date").map((item) => item.label)).toEqual([
      "Equals…", "Before…", "After…", "Between…", "Tomorrow", "Today", "Yesterday", "Next Week", "This Week", "Last Week",
      "Next Month", "This Month", "Last Month", "Next Year", "This Year", "Last Year", "Year to Date", "Custom Filter…",
    ]);
    expect(conditionItems("number").find((item) => item.id === "between")?.action).toEqual({
      kind: "custom",
      draft: { op1: "greaterThanOrEqual", value1: "", join: "and", op2: "lessThanOrEqual", value2: "" },
    });
    expect(conditionItems("date").find((item) => item.id === "today")?.action).toEqual({ kind: "dynamic", type: "today" });
    expect(customOperatorLabel("lessThan", "date")).toBe("is before");
    expect(customOperatorLabel("lessThan", "number")).toBe("is less than");
  });

  it("stores a typed number or date as invariant number text and anything else as typed, with the text operators' wildcards", () => {
    expect(conditionFor("greaterThan", "1,5", DE, "1900")).toEqual({ operator: "greaterThan", value: "1.5" });
    expect(conditionFor("equals", "50%", US, "1900")).toEqual({ operator: "equal", value: "0.5" });
    expect(conditionFor("lessThan", "1/2/2024", US, "1900")).toEqual({ operator: "lessThan", value: "45293" });
    expect(conditionFor("equals", "West", US, "1900")).toEqual({ operator: "equal", value: "West" });
    expect(conditionFor("beginsWith", "We", US, "1900")).toEqual({ operator: "equal", value: "We*" });
    expect(conditionFor("notEndsWith", "st", US, "1900")).toEqual({ operator: "notEqual", value: "*st" });
    expect(conditionFor("contains", "a?c", US, "1900")).toEqual({ operator: "equal", value: "*a?c*" });
    expect(conditionFor("notContains", "x", US, "1900")).toEqual({ operator: "notEqual", value: "*x*" });
    expect(conditionFor("equals", "W*t", US, "1900")).toEqual({ operator: "equal", value: "W*t" });
  });

  it("builds a custom criterion of one or two conditions, and reads one back into the dialog", () => {
    expect(customCriteria({ op1: "equals", value1: "", join: "and", op2: "", value2: "" }, US, "1900")).toBeNull();
    const two = customCriteria({ op1: "greaterThanOrEqual", value1: "10", join: "or", op2: "beginsWith", value2: "N" }, US, "1900");
    expect(two).toEqual({ kind: "custom", join: "or", conditions: [{ operator: "greaterThanOrEqual", value: "10" }, { operator: "equal", value: "N*" }] });
    expect(customDraftOf(two!, US)).toEqual({ op1: "greaterThanOrEqual", value1: "10", join: "or", op2: "beginsWith", value2: "N" });
    expect(customDraftOf({ kind: "custom", join: "and", conditions: [{ operator: "lessThan", value: "2.5" }] }, DE)).toMatchObject({ op1: "lessThan", value1: "2,5", op2: "" });
    expect(customDraftOf({ kind: "top10", top: true, percent: false, value: 10 }, US)).toBeNull();
  });

  it("accepts Top 10 values from 1 to 500 items or 1 to 100 percent", () => {
    expect(top10Criteria(true, false, "500")).toEqual({ kind: "top10", top: true, percent: false, value: 500 });
    expect(top10Criteria(false, true, "100")).toEqual({ kind: "top10", top: false, percent: true, value: 100 });
    expect(top10Criteria(true, true, "101")).toBeNull();
    expect(top10Criteria(true, false, "0")).toBeNull();
    expect(top10Criteria(true, false, "2.5")).toBeNull();
    expect(top10Criteria(true, false, "")).toBeNull();
  });
});
