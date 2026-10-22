// SPEC-129 "Rule descriptions" — describeRule's exact strings.
import { describe, expect, it } from "vitest";
import { describeRule } from "./cfDescriptions";
import type { CfRule, CfRuleDraft } from "../../lib/spreadsheet/conditionalFormats";

function rule(body: CfRuleDraft): CfRule {
  return { id: "r1", priority: 1, ...body } as CfRule;
}

describe("describeRule (SPEC-129 Rule descriptions)", () => {
  it("describes cellIs operators", () => {
    expect(describeRule(rule({ kind: "cellIs", operator: "greaterThan", formulas: ["100"] }))).toBe("Cell value > 100");
    expect(describeRule(rule({ kind: "cellIs", operator: "equal", formulas: ["5"] }))).toBe("Cell value = 5");
    expect(describeRule(rule({ kind: "cellIs", operator: "notEqual", formulas: ["5"] }))).toBe("Cell value ≠ 5");
    expect(describeRule(rule({ kind: "cellIs", operator: "greaterThanOrEqual", formulas: ["5"] }))).toBe("Cell value ≥ 5");
    expect(describeRule(rule({ kind: "cellIs", operator: "lessThan", formulas: ["5"] }))).toBe("Cell value < 5");
    expect(describeRule(rule({ kind: "cellIs", operator: "lessThanOrEqual", formulas: ["5"] }))).toBe("Cell value ≤ 5");
    expect(describeRule(rule({ kind: "cellIs", operator: "between", formulas: ["1", "10"] }))).toBe("Cell value between 1 and 10");
    expect(describeRule(rule({ kind: "cellIs", operator: "notBetween", formulas: ["1", "10"] }))).toBe("Cell value not between 1 and 10");
  });

  it("describes a formula rule with its stored (origin-relative) text", () => {
    expect(describeRule(rule({ kind: "expression", formula: "$C2>AVERAGE($C$2:$C$21)" }))).toBe("Formula: =$C2>AVERAGE($C$2:$C$21)");
  });

  it("describes the statistics-based kinds", () => {
    expect(describeRule(rule({ kind: "colorScale", stops: [] }))).toBe("Graded colour scale");
    expect(
      describeRule(
        rule({
          kind: "dataBar",
          bar: {
            min: { type: "min" },
            max: { type: "max" },
            color: "#000000",
            gradient: true,
            axis: "automatic",
            direction: "context",
            minLengthPct: 10,
            maxLengthPct: 90,
            showValue: true,
          },
        }),
      ),
    ).toBe("Data bar");
    expect(describeRule(rule({ kind: "iconSet", icons: { set: "3Arrows", thresholds: [], reverse: false, showValue: true } }))).toBe(
      "Icon set (3 Arrows)",
    );
  });

  it("describes top10, aboveAverage and duplicates", () => {
    expect(describeRule(rule({ kind: "top10", bottom: false, percent: false, rank: 10 }))).toBe("Top 10");
    expect(describeRule(rule({ kind: "top10", bottom: true, percent: true, rank: 10 }))).toBe("Bottom 10%");
    expect(describeRule(rule({ kind: "aboveAverage", below: false, equalAverage: false }))).toBe("Above average");
    expect(describeRule(rule({ kind: "aboveAverage", below: false, equalAverage: false, stdDev: 1 }))).toBe("Above average (1 std dev)");
    expect(describeRule(rule({ kind: "aboveAverage", below: true, equalAverage: false }))).toBe("Below average");
    expect(describeRule(rule({ kind: "duplicates", unique: false }))).toBe("Duplicate values");
    expect(describeRule(rule({ kind: "duplicates", unique: true }))).toBe("Unique values");
  });

  it("describes text, blanks, errors and timePeriod", () => {
    expect(describeRule(rule({ kind: "text", operator: "contains", text: "abc" }))).toBe('Cell value contains "abc"');
    expect(describeRule(rule({ kind: "text", operator: "notContains", text: "abc" }))).toBe('Cell value does not contain "abc"');
    expect(describeRule(rule({ kind: "text", operator: "beginsWith", text: "abc" }))).toBe('Cell value begins with "abc"');
    expect(describeRule(rule({ kind: "text", operator: "endsWith", text: "abc" }))).toBe('Cell value ends with "abc"');
    expect(describeRule(rule({ kind: "blanks", not: false }))).toBe("Blanks");
    expect(describeRule(rule({ kind: "blanks", not: true }))).toBe("No blanks");
    expect(describeRule(rule({ kind: "errors", not: false }))).toBe("Errors");
    expect(describeRule(rule({ kind: "errors", not: true }))).toBe("No errors");
    expect(describeRule(rule({ kind: "timePeriod", period: "last7Days" }))).toBe("Date occurring: Last 7 days");
  });

  it("describes a preserved rule", () => {
    expect(describeRule(rule({ kind: "preserved", xml: "<cfRule/>", where: "base" }))).toBe("Rule kept from the file");
  });
});
