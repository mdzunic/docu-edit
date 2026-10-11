// SPEC-129 "The rule editor" — the editor's fields to and from a draft:
// locale values through parseInput, `=` formulas parsed and moved from the
// active cell onto the origin, threshold value objects, and the inline
// errors that save nothing.
import { describe, expect, it } from "vitest";
import { defaultIconThresholds, type CfRuleDraft } from "../../lib/spreadsheet/conditionalFormats";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { defaultFields, draftFromFields, fieldsFromDraft, FORMULA_ERROR, NUMBER_ERROR, RANK_ERROR, takesFormat, VALUE_REQUIRED_ERROR, withIconSet, type CfEditorContext } from "./cfEditorFields";

const ctx = (over: Partial<CfEditorContext> = {}): CfEditorContext => ({
  origin: { row: 1, col: 1 },
  active: { row: 1, col: 1 },
  locale: SPREADSHEET_LOCALES["en-US"],
  dateSystem: "1900",
  ...over,
});

function roundTrip(draft: CfRuleDraft, context = ctx()): CfRuleDraft {
  const result = draftFromFields(fieldsFromDraft(draft, context), context);
  if (!("draft" in result)) {
    throw new Error(result.error);
  }
  return result.draft;
}

describe("cfEditorFields (SPEC-129)", () => {
  it("round-trips every editable kind", () => {
    const drafts: CfRuleDraft[] = [
      { kind: "cellIs", operator: "between", formulas: ["1", "10"] },
      { kind: "cellIs", operator: "equal", formulas: ['"late"'] },
      { kind: "cellIs", operator: "lessThan", formulas: ["$Z$1"] },
      { kind: "expression", formula: "$H2=\"late\"" },
      { kind: "colorScale", stops: [{ at: { type: "min" }, color: "#63BE7B" }, { at: { type: "percentile", value: "50" }, color: "#FFEB84" }, { at: { type: "max" }, color: "#F8696B" }] },
      { kind: "colorScale", stops: [{ at: { type: "num", value: "0" }, color: "#FFFFFF" }, { at: { type: "formula", value: "$A$1" }, color: "#F8696B" }] },
      { kind: "iconSet", icons: { set: "4Arrows", thresholds: defaultIconThresholds(4), reverse: true, showValue: false } },
      { kind: "top10", bottom: true, percent: true, rank: 5 },
      { kind: "aboveAverage", below: false, equalAverage: true, stdDev: 2 },
      { kind: "duplicates", unique: true },
      { kind: "text", operator: "beginsWith", text: "ab\"c" },
      { kind: "blanks", not: true },
      { kind: "errors", not: false },
      { kind: "timePeriod", period: "last7Days" },
    ];
    for (const draft of drafts) {
      expect(roundTrip(draft)).toEqual(draft);
    }
    const bar: CfRuleDraft = {
      kind: "dataBar",
      bar: { min: { type: "autoMin" }, max: { type: "num", value: "100" }, color: "#FF555A", gradient: false, axis: "middle", direction: "context", minLengthPct: 0, maxLengthPct: 100, showValue: false },
    };
    expect(roundTrip(bar)).toEqual(bar);
  });

  it("reads values in the locale and shows stored numbers in it", () => {
    const de = ctx({ locale: SPREADSHEET_LOCALES["de-DE"] });
    const fields = fieldsFromDraft({ kind: "cellIs", operator: "greaterThan", formulas: ["1.5"] }, de);
    expect(fields.value1).toBe("1,5");
    expect(draftFromFields({ ...fields, value1: "2,25" }, de)).toEqual({ draft: { kind: "cellIs", operator: "greaterThan", formulas: ["2.25"] } });
    expect(draftFromFields({ ...fields, value1: "TRUE" }, de)).toEqual({ draft: { kind: "cellIs", operator: "greaterThan", formulas: ["TRUE"] } });
  });

  it("moves a formula typed at the active cell onto the origin, and shows it back at the active cell", () => {
    const context = ctx({ active: { row: 4, col: 2 } });
    const fields = { ...defaultFields(), type: "formula" as const, value1: "=C5>AVERAGE($C$2:$C$21)" };
    expect(draftFromFields(fields, context)).toEqual({ draft: { kind: "expression", formula: "B2>AVERAGE($C$2:$C$21)" } });
    expect(fieldsFromDraft({ kind: "expression", formula: "B2>AVERAGE($C$2:$C$21)" }, context).value1).toBe("=C5>AVERAGE($C$2:$C$21)");
    const cell = { ...defaultFields(), value1: "=C5" };
    expect(draftFromFields(cell, context)).toEqual({ draft: { kind: "cellIs", operator: "greaterThan", formulas: ["B2"] } });
  });

  it("refuses an unparseable formula, a missing value, a non-numeric threshold and a bad rank", () => {
    expect(draftFromFields({ ...defaultFields(), type: "formula", value1: "=SUM(" }, ctx())).toEqual({ error: FORMULA_ERROR, field: "value1" });
    expect(draftFromFields({ ...defaultFields(), type: "formula", value1: "" }, ctx())).toEqual({ error: FORMULA_ERROR, field: "value1" });
    expect(draftFromFields({ ...defaultFields(), value1: "" }, ctx())).toEqual({ error: VALUE_REQUIRED_ERROR, field: "value1" });
    expect(draftFromFields({ ...defaultFields(), operator: "between", value1: "1", value2: "=)" }, ctx())).toEqual({ error: FORMULA_ERROR, field: "value2" });
    const scale = { ...defaultFields(), type: "values" as const, style: "3-color" as const };
    scale.stops = scale.stops.map((stop, index) => (index === 1 ? { ...stop, value: "abc" } : stop));
    expect(draftFromFields(scale, ctx())).toEqual({ error: NUMBER_ERROR, field: "stop-1" });
    expect(draftFromFields({ ...defaultFields(), type: "rank", rank: "0" }, ctx())).toEqual({ error: RANK_ERROR, field: "rank" });
    expect(draftFromFields({ ...defaultFields(), type: "rank", rank: "2.5" }, ctx())).toEqual({ error: RANK_ERROR, field: "rank" });
  });

  it("resets the thresholds when the icon set changes size, and knows which types take a format", () => {
    expect(withIconSet(defaultFields(), "5Quarters").iconThresholds).toHaveLength(5);
    expect(takesFormat({ type: "values" })).toBe(false);
    for (const type of ["contain", "rank", "average", "unique", "formula"] as const) {
      expect(takesFormat({ type })).toBe(true);
    }
  });
});
