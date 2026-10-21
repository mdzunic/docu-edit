import { afterEach, describe, expect, it } from "vitest";
import { setClock } from "../core/clock";
import { DEFAULT_THEME } from "../core/theme";
import { parseA1, parseRange, subtractRanges, type GridRange } from "./address";
import { colorScaleColor, createCfEvaluator, dataBarGeometry, iconIndex, mixTowardWhite, resolveCfValue, cfStatistics } from "./cfEvaluator";
import {
  addConditionalFormat,
  clearRulesInRanges,
  defaultIconThresholds,
  deleteRule,
  layerDxf,
  moveRule,
  renumberPriorities,
  rulesByPriority,
  setFormatRanges,
  setStopIfTrue,
  updateRule,
  type CfDataBar,
  type CfRule,
  type CfRuleBody,
  type ConditionalFormat,
  type Dxf,
} from "./conditionalFormats";
import { storedValueContext } from "./formula/evaluate";
import type { Cell, CellValue, Row, Sheet, SpreadsheetModel } from "./model";
import { dateToSerial } from "./serial";

// --- builders ------------------------------------------------------------------------------------

const RED: Dxf = { font: { color: "#9C0006" }, fill: { kind: "solid", color: "#FFC7CE" } };
const GREEN: Dxf = { font: { color: "#006100", bold: true }, fill: { kind: "solid", color: "#C6EFCE" } };
const ITALIC: Dxf = { font: { italic: true, color: "#0000FF" }, borders: { top: { style: "thin" } } };

function range(text: string): GridRange {
  return parseRange(text)!;
}

/** A sheet from A1 → value. */
function sheetOf(values: Record<string, CellValue>, formats: ConditionalFormat[] = [], id = "s1", name = "Sheet1"): Sheet {
  const rows = new Map<number, Cell[]>();
  for (const [address, value] of Object.entries(values)) {
    const ref = parseA1(address)!;
    rows.set(ref.row, [...(rows.get(ref.row) ?? []), { col: ref.col, value }]);
  }
  const sorted: Row[] = [...rows.entries()].toSorted(([a], [b]) => a - b).map(([index, cells]) => ({ index, cells: cells.toSorted((a, b) => a.col - b.col) }));
  return { id, name, visibility: "visible", kind: "grid", rows: sorted, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], validations: [], conditionalFormats: formats };
}

function book(values: Record<string, CellValue>, formats: ConditionalFormat[] = [], dxfs: Dxf[] = [RED, GREEN, ITALIC]): SpreadsheetModel {
  return {
    kind: "spreadsheet",
    id: "book",
    title: "Book",
    format: "xlsx",
    readerVersion: 10,
    sheets: [sheetOf(values, formats)],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{}],
    defaultStyleId: 0,
    dxfs,
  };
}

let ids = 0;
function rule(body: CfRuleBody, extra: Partial<CfRule> = {}): CfRule {
  ids += 1;
  return { id: `r${ids}`, priority: ids, ...body, ...extra } as CfRule;
}

function format(ranges: string, ...rules: CfRule[]): ConditionalFormat {
  return { id: `f${rules[0]?.id ?? ""}`, ranges: ranges.split(" ").map(range), rules };
}

function evaluator(model: SpreadsheetModel) {
  return createCfEvaluator(model, model.sheets[0]!.id, { base: storedValueContext(model), theme: DEFAULT_THEME });
}

/** The result at an A1 address. */
function at(model: SpreadsheetModel, address: string) {
  const ref = parseA1(address)!;
  return evaluator(model).resultFor(ref.row, ref.col);
}

/** Whether the rule's dxf applies at an address. */
function formatted(model: SpreadsheetModel, address: string): boolean {
  return at(model, address)?.dxf !== undefined;
}

const BAR: CfDataBar = {
  min: { type: "min" },
  max: { type: "max" },
  color: "#638EC6",
  gradient: false,
  axis: "automatic",
  direction: "context",
  minLengthPct: 0,
  maxLengthPct: 100,
  showValue: true,
};

/** A 1900-system date serial. */
function day(y: number, m: number, d: number): number {
  return dateToSerial(y, m, d, "1900");
}

/** The value-object formula stub: $B$1 is 100, anything else unresolved. */
function formula(text: string): number | null {
  return text === "$B$1" ? 100 : null;
}

/** A 3-icon icon-set rule body with the default thresholds. */
function icons(set: "3Arrows" | "NoIcons", reverse = false): CfRuleBody {
  return { kind: "iconSet", icons: { set, thresholds: defaultIconThresholds(3), reverse, showValue: true } };
}

afterEach(() => setClock(null));

// --- semantics ------------------------------------------------------------------------------------

describe("Evaluation semantics", () => {
  it("cellIs compares with SPEC-120's rules: a blank is 0 against a number, text case-insensitively", () => {
    const greater = book({ B2: 120, B3: 140, B4: 90, B6: "abc" }, [format("B2:B8", rule({ kind: "cellIs", operator: "greaterThan", formulas: ["100"] }, { dxfId: 0 }))]);
    expect(formatted(greater, "B3")).toBe(true);
    expect(formatted(greater, "B4")).toBe(false);
    expect(formatted(greater, "B5")).toBe(false); // blank = 0
    expect(formatted(greater, "B6")).toBe(true); // text sorts after numbers
    const less = book({ A1: 3, A2: 9 }, [format("A1:A3", rule({ kind: "cellIs", operator: "lessThan", formulas: ["5"] }, { dxfId: 0 }))]);
    expect(formatted(less, "A1")).toBe(true);
    expect(formatted(less, "A2")).toBe(false);
    expect(formatted(less, "A3")).toBe(true); // Excel formats empty cells for "less than 5"
    const text = book({ A1: "ABC", A2: "abd" }, [format("A1:A3", rule({ kind: "cellIs", operator: "equal", formulas: ['"abc"'] }, { dxfId: 0 }))]);
    expect(formatted(text, "A1")).toBe(true);
    expect(formatted(text, "A2")).toBe(false);
    expect(formatted(text, "A3")).toBe(false);
  });

  it("between is inclusive with its bounds in either order; notBetween is its complement", () => {
    const values = { A1: 1, A2: 5, A3: 10, A4: 11, A5: 0 };
    const between = book(values, [format("A1:A5", rule({ kind: "cellIs", operator: "between", formulas: ["10", "1"] }, { dxfId: 0 }))]);
    expect(["A1", "A2", "A3", "A4", "A5"].map((a) => formatted(between, a))).toEqual([true, true, true, false, false]);
    const outside = book(values, [format("A1:A5", rule({ kind: "cellIs", operator: "notBetween", formulas: ["1", "10"] }, { dxfId: 0 }))]);
    expect(["A1", "A2", "A3", "A4", "A5"].map((a) => formatted(outside, a))).toEqual([false, false, false, true, true]);
  });

  it("evaluates a formula stored relative to the origin at each cell, never re-shifting its text", () => {
    const model = book({ A2: 1, A3: 2, A4: 3, H2: "late", H3: "on time", H4: "LATE", Z1: 2 }, [
      format("A2:H4", rule({ kind: "expression", formula: '$H2="late"' }, { dxfId: 0 })),
      format("C2:C4", rule({ kind: "cellIs", operator: "greaterThan", formulas: ["$Z$1"] }, { dxfId: 1 })),
    ]);
    expect(formatted(model, "A2")).toBe(true);
    expect(formatted(model, "D3")).toBe(false);
    expect(formatted(model, "G4")).toBe(true);
    const relative = book({ A1: 1, A2: 5, B1: 2, B2: 4 }, [format("A1:A2", rule({ kind: "expression", formula: "A1<B1" }, { dxfId: 0 }))]);
    expect(formatted(relative, "A1")).toBe(true);
    expect(formatted(relative, "A2")).toBe(false);
  });

  it("an expression is true only for TRUE or a non-zero number; an error, text or an unparseable formula is false", () => {
    const cases: [string, boolean][] = [["TRUE", true], ["2", true], ["-0.5", true], ["0", false], ["FALSE", false], ['"TRUE"', false], ["1/0", false], ["NA()", false], ["1+", false]];
    for (const [text, expected] of cases) {
      expect(formatted(book({ A1: 1 }, [format("A1", rule({ kind: "expression", formula: text }, { dxfId: 0 }))]), "A1"), text).toBe(expected);
    }
  });

  it("top10 takes the rank-th largest or smallest numeric value, ties included, percent ranks at least 1", () => {
    const values = { A1: 10, A2: 9, A3: 9, A4: 7, A5: 1, A6: "x", A7: true };
    const top2 = book(values, [format("A1:A7", rule({ kind: "top10", bottom: false, percent: false, rank: 2 }, { dxfId: 0 }))]);
    expect(["A1", "A2", "A3", "A4", "A6", "A7"].map((a) => formatted(top2, a))).toEqual([true, true, true, false, false, false]);
    const bottom1 = book(values, [format("A1:A7", rule({ kind: "top10", bottom: true, percent: false, rank: 1 }, { dxfId: 0 }))]);
    expect(["A4", "A5"].map((a) => formatted(bottom1, a))).toEqual([false, true]);
    const percent = book(values, [format("A1:A7", rule({ kind: "top10", bottom: false, percent: true, rank: 10 }, { dxfId: 0 }))]);
    expect(["A1", "A2"].map((a) => formatted(percent, a))).toEqual([true, false]); // max(1, floor(5 × 10 / 100)) = 1
  });

  it("aboveAverage uses the mean, equalAverage and the population standard deviation", () => {
    const values = { A1: 2, A2: 4, A3: 4, A4: 4, A5: 5, A6: 5, A7: 7, A8: 9 }; // mean 5, σ 2
    const above = book(values, [format("A1:A8", rule({ kind: "aboveAverage", below: false, equalAverage: false }, { dxfId: 0 }))]);
    expect(["A5", "A7"].map((a) => formatted(above, a))).toEqual([false, true]);
    const atOrAbove = book(values, [format("A1:A8", rule({ kind: "aboveAverage", below: false, equalAverage: true }, { dxfId: 0 }))]);
    expect(formatted(atOrAbove, "A5")).toBe(true);
    const below = book(values, [format("A1:A8", rule({ kind: "aboveAverage", below: true, equalAverage: false }, { dxfId: 0 }))]);
    expect(["A1", "A5"].map((a) => formatted(below, a))).toEqual([true, false]);
    const sigma = book(values, [format("A1:A8", rule({ kind: "aboveAverage", below: false, equalAverage: false, stdDev: 1 }, { dxfId: 0 }))]);
    expect(["A7", "A8"].map((a) => formatted(sigma, a))).toEqual([false, true]); // > 5 + 2
    const sigmaBelow = book(values, [format("A1:A8", rule({ kind: "aboveAverage", below: true, equalAverage: false, stdDev: 1 }, { dxfId: 0 }))]);
    expect(["A1", "A2"].map((a) => formatted(sigmaBelow, a))).toEqual([true, false]); // < 5 − 2
  });

  it("duplicates compare case-insensitively across all ranges; blanks never match; unique is exactly once", () => {
    const values = { A1: "Apple", A2: "apple", A3: 1, C1: "1", C2: 1, C3: "pear" };
    const dupes = book(values, [format("A1:A4 C1:C3", rule({ kind: "duplicates", unique: false }, { dxfId: 0 }))]);
    expect(["A1", "A2", "A3", "C1", "C2", "C3", "A4"].map((a) => formatted(dupes, a))).toEqual([true, true, true, false, true, false, false]);
    const unique = book(values, [format("A1:A4 C1:C3", rule({ kind: "duplicates", unique: true }, { dxfId: 0 }))]);
    expect(["A1", "C1", "C3", "A4"].map((a) => formatted(unique, a))).toEqual([false, true, true, false]);
  });

  it("text rules match the value's text case-insensitively; notContains is true for a blank", () => {
    const values = { A1: "Total Sales", A2: "subtotal", A3: 1234, A4: { error: "#N/A" as const } };
    const kinds: [CfRuleBody, boolean[]][] = [
      [{ kind: "text", operator: "contains", text: "TOTAL" }, [true, true, false, false, false]],
      [{ kind: "text", operator: "notContains", text: "total" }, [false, false, true, true, true]],
      [{ kind: "text", operator: "beginsWith", text: "tot" }, [true, false, false, false, false]],
      [{ kind: "text", operator: "endsWith", text: "34" }, [false, false, true, false, false]],
    ];
    for (const [body, expected] of kinds) {
      const model = book(values, [format("A1:A5", rule(body, { dxfId: 0 }))]);
      expect(["A1", "A2", "A3", "A4", "A5"].map((a) => formatted(model, a)), JSON.stringify(body)).toEqual(expected);
    }
  });

  it("blanks test the trimmed text; errors test for a CellError", () => {
    const values = { A1: "  ", A2: "x", A3: 0, A4: { error: "#DIV/0!" as const } };
    const blanks = book(values, [format("A1:A5", rule({ kind: "blanks", not: false }, { dxfId: 0 }))]);
    expect(["A1", "A2", "A3", "A5"].map((a) => formatted(blanks, a))).toEqual([true, false, false, true]);
    const noBlanks = book(values, [format("A1:A5", rule({ kind: "blanks", not: true }, { dxfId: 0 }))]);
    expect(["A1", "A2", "A3", "A5"].map((a) => formatted(noBlanks, a))).toEqual([false, true, true, false]);
    const errors = book(values, [format("A1:A5", rule({ kind: "errors", not: false }, { dxfId: 0 }))]);
    expect(["A2", "A4", "A5"].map((a) => formatted(errors, a))).toEqual([false, true, false]);
    const noErrors = book(values, [format("A1:A5", rule({ kind: "errors", not: true }, { dxfId: 0 }))]);
    expect(["A2", "A4", "A5"].map((a) => formatted(noErrors, a))).toEqual([true, false, true]);
  });

  it("timePeriod tests floor(v) against today from core/clock, weeks starting on Sunday", () => {
    setClock(new Date(2026, 9, 14, 15, 30)); // Wednesday 14 October 2026, local time
    const values = {
      A1: day(2026, 10, 14) + 0.75,
      A2: day(2026, 10, 13),
      A3: day(2026, 10, 15),
      A4: day(2026, 10, 8),
      A5: day(2026, 10, 7),
      A6: day(2026, 10, 11), // Sunday — this week's first day
      A7: day(2026, 10, 10), // Saturday — last week
      A8: day(2026, 9, 30),
      A9: day(2027, 1, 1),
      A10: "today",
    };
    const matches = (period: Parameters<typeof rule>[0] & { kind: "timePeriod" }) => {
      const model = book(values, [format("A1:A10", rule(period, { dxfId: 0 }))]);
      return Object.keys(values).filter((a) => formatted(model, a));
    };
    expect(matches({ kind: "timePeriod", period: "today" })).toEqual(["A1"]);
    expect(matches({ kind: "timePeriod", period: "yesterday" })).toEqual(["A2"]);
    expect(matches({ kind: "timePeriod", period: "tomorrow" })).toEqual(["A3"]);
    expect(matches({ kind: "timePeriod", period: "last7Days" })).toEqual(["A1", "A2", "A4", "A6", "A7"]);
    expect(matches({ kind: "timePeriod", period: "thisWeek" })).toEqual(["A1", "A2", "A3", "A6"]);
    expect(matches({ kind: "timePeriod", period: "lastWeek" })).toEqual(["A4", "A5", "A7"]);
    expect(matches({ kind: "timePeriod", period: "lastMonth" })).toEqual(["A8"]);
    expect(matches({ kind: "timePeriod", period: "nextYear" })).toEqual(["A9"]);
    expect(matches({ kind: "timePeriod", period: "thisMonth" })).toEqual(["A1", "A2", "A3", "A4", "A5", "A6", "A7"]);
  });

  it("colour scales, data bars and icon sets only take numbers; a preserved rule never evaluates", () => {
    const values = { A1: 1, A2: "2", A3: true, A5: 10 };
    const model = book(values, [
      format(
        "A1:A5",
        rule({ kind: "colorScale", stops: [{ at: { type: "min" }, color: "#000000" }, { at: { type: "max" }, color: "#FFFFFF" }] }),
        rule({ kind: "dataBar", bar: BAR }),
        rule({ kind: "iconSet", icons: { set: "3Arrows", thresholds: defaultIconThresholds(3), reverse: false, showValue: true } }),
        rule({ kind: "preserved", xml: "<cfRule/>", where: "base" }, { dxfId: 0 }),
      ),
    ]);
    expect(at(model, "A1")).toMatchObject({ scaleColor: "#000000", icon: { set: "3Arrows", index: 0 } });
    expect(at(model, "A2")).toBeUndefined();
    expect(at(model, "A3")).toBeUndefined();
    expect(at(model, "A4")).toBeUndefined();
  });
});

// --- statistics and value objects -----------------------------------------------------------------

describe("Value objects and statistics", () => {
  const sheet = sheetOf({ A1: -4, A2: 2, A3: 6, A4: 10, A5: "12", A6: true, B1: 100 });
  const stats = cfStatistics(sheet, [range("A1:A6")]);

  it("takes count, min, max, mean, σ and the sorted values over the numeric cells only", () => {
    expect(stats).toMatchObject({ count: 4, min: -4, max: 10, mean: 3.5, sorted: [-4, 2, 6, 10] });
    expect(stats.stdDev).toBeCloseTo(Math.sqrt((7.5 ** 2 + 1.5 ** 2 + 2.5 ** 2 + 6.5 ** 2) / 4), 12);
  });

  it("resolves every value-object type", () => {
    const resolve = (type: Parameters<typeof resolveCfValue>[0]["type"], value?: string) => resolveCfValue({ type, value }, stats, formula);
    expect(resolve("min")).toBe(-4);
    expect(resolve("max")).toBe(10);
    expect(resolve("autoMin")).toBe(-4);
    expect(resolve("autoMax")).toBe(10);
    expect(resolveCfValue({ type: "autoMin" }, { ...stats, min: 3 }, formula)).toBe(0);
    expect(resolveCfValue({ type: "autoMax" }, { ...stats, max: -1 }, formula)).toBe(0);
    expect(resolve("num", "7.5")).toBe(7.5);
    expect(resolve("num", "$B$1")).toBe(100);
    expect(resolve("percent", "50")).toBe(3); // -4 + 14 × 0.5
    expect(resolve("percentile", "50")).toBe(4); // PERCENTILE.INC({-4,2,6,10}, 0.5)
    expect(resolve("percentile", "25")).toBe(0.5);
    expect(resolve("formula", "$B$1")).toBe(100);
    expect(resolve("formula", "BAD(")).toBeNull();
    expect(resolveCfValue({ type: "min" }, cfStatistics(sheet, [range("C1:C9")]), formula)).toBeNull();
  });
});

describe("Colour scale", () => {
  const stops = [
    { value: 0, color: "#f8696b" },
    { value: 50, color: "#ffeb84" },
    { value: 100, color: "#63be7b" },
  ];

  it("interpolates each sRGB channel linearly between the bracketing stops, rounded", () => {
    expect(colorScaleColor(25, stops)).toBe("#fcaa78");
    expect(colorScaleColor(50, stops)).toBe("#ffeb84");
    expect(colorScaleColor(75, stops)).toBe("#b1d580");
    expect(colorScaleColor(1, [stops[0]!, stops[2]!])).toBe("#f76a6b"); // 248−1.49, 105+0.85, 107+0.16
  });

  it("clamps outside the stops and resolves coincident stops to the lower stop's colour", () => {
    expect(colorScaleColor(-5, stops)).toBe("#f8696b");
    expect(colorScaleColor(500, stops)).toBe("#63be7b");
    expect(colorScaleColor(5, [{ value: 5, color: "#000000" }, { value: 5, color: "#ffffff" }, { value: 9, color: "#ff0000" }])).toBe("#000000");
  });

  it("evaluates a three-colour scale at the 50th percentile with theme colours", () => {
    const model = book({ A1: 0, A2: 10, A3: 20, A4: 30, A5: 40 }, [
      format(
        "A1:A5",
        rule({
          kind: "colorScale",
          stops: [
            { at: { type: "min" }, color: "#000000" },
            { at: { type: "percentile", value: "50" }, color: { theme: "lt1" } },
            { at: { type: "max" }, color: "#000000" },
          ],
        }),
      ),
    ]);
    expect(at(model, "A3")?.scaleColor).toBe("#ffffff");
    expect(at(model, "A2")?.scaleColor).toBe("#808080");
    expect(at(model, "A5")?.scaleColor).toBe("#000000");
  });
});

describe("Data bar", () => {
  it("spans minLength..maxLength from the leading edge when every value is non-negative", () => {
    const bar = { ...BAR, minLengthPct: 10, maxLengthPct: 90 };
    expect(dataBarGeometry(0, 0, 100, bar)).toEqual({ startPct: 0, widthPct: 10, negative: false });
    expect(dataBarGeometry(50, 0, 100, bar)).toEqual({ startPct: 0, widthPct: 50, negative: false });
    expect(dataBarGeometry(100, 0, 100, bar)).toEqual({ startPct: 0, widthPct: 90, negative: false });
    expect(dataBarGeometry(150, 0, 100, bar).widthPct).toBe(90);
    expect(dataBarGeometry(7, 7, 7, bar).widthPct).toBe(90); // hi = lo → the maximum
  });

  it("draws negative values left of an automatic axis at |min| / (|min| + max)", () => {
    expect(dataBarGeometry(-20, -20, 60, BAR)).toEqual({ startPct: 0, widthPct: 25, negative: true, axisPct: 25 });
    expect(dataBarGeometry(-10, -20, 60, BAR)).toEqual({ startPct: 12.5, widthPct: 12.5, negative: true, axisPct: 25 });
    expect(dataBarGeometry(30, -20, 60, BAR)).toEqual({ startPct: 25, widthPct: 37.5, negative: false, axisPct: 25 });
    expect(dataBarGeometry(30, -20, 60, { ...BAR, axis: "middle" })).toEqual({ startPct: 50, widthPct: 25, negative: false, axisPct: 50 });
    expect(dataBarGeometry(-10, -20, 60, { ...BAR, axis: "none" })).toEqual({ startPct: 0, widthPct: 12.5, negative: true });
    expect(dataBarGeometry(30, -20, 60, { ...BAR, direction: "rightToLeft" })).toEqual({ startPct: 37.5, widthPct: 37.5, negative: false, axisPct: 75 });
  });

  it("colours a negative bar in the negative colour, mixes the gradient's end 90% toward white, and hides the value on showValue false", () => {
    const model = book({ A1: -5, A2: 5, A3: 10 }, [
      format("A1:A3", rule({ kind: "dataBar", bar: { ...BAR, gradient: true, negativeColor: "#FF0000", border: "#112233", axisColor: "#000000", showValue: false } })),
    ]);
    expect(at(model, "A2")).toEqual({
      dataBar: { startPct: 33.33, widthPct: 33.33, negative: false, axisPct: 33.33, color: "#638ec6", endColor: mixTowardWhite("#638EC6", 0.9), borderColor: "#112233", axisColor: "#000000" },
      hideValue: true,
    });
    expect(at(model, "A1")?.dataBar).toMatchObject({ negative: true, color: "#ff0000", startPct: 0, widthPct: 33.33 });
    expect(mixTowardWhite("#638EC6", 0.9)).toBe("#eff4f9");
  });
});

describe("Icon set", () => {
  const thresholds = [
    { value: 0, gte: true },
    { value: 33, gte: true },
    { value: 67, gte: false },
  ];

  it("picks the highest threshold met (≥ when gte, else >), the first being the floor; reverse flips", () => {
    expect(iconIndex(-10, thresholds, false)).toBe(0);
    expect(iconIndex(33, thresholds, false)).toBe(1);
    expect(iconIndex(67, thresholds, false)).toBe(1);
    expect(iconIndex(67.5, thresholds, false)).toBe(2);
    expect(iconIndex(67.5, thresholds, true)).toBe(0);
  });

  it("resolves percent thresholds over the format's numbers; NoIcons draws nothing but still wins the icon", () => {
    const model = book({ D2: 100, D3: 50, D4: 0 }, [format("D2:D4", rule(icons("3Arrows")))]);
    expect(at(model, "D2")?.icon).toEqual({ set: "3Arrows", index: 2 });
    expect(at(model, "D3")?.icon).toEqual({ set: "3Arrows", index: 1 });
    expect(at(model, "D4")?.icon).toEqual({ set: "3Arrows", index: 0 });
    const none = book({ D2: 100 }, [format("D2:D4", rule(icons("NoIcons"))), format("D2:D4", rule(icons("3Arrows")))]);
    expect(at(none, "D2")).toBeUndefined();
  });
});

// --- priority and layering ------------------------------------------------------------------------

describe("Priority and layering", () => {
  it("walks rules by ascending priority, each true rule adding only the properties still unset", () => {
    const model = book({ A1: 200 }, [
      format("A1", rule({ kind: "cellIs", operator: "greaterThan", formulas: ["100"] }, { dxfId: 1, priority: 2 })),
      format("A1", rule({ kind: "cellIs", operator: "greaterThan", formulas: ["150"] }, { dxfId: 0, priority: 1 })),
      format("A1", rule({ kind: "expression", formula: "TRUE" }, { dxfId: 2, priority: 3 })),
    ]);
    expect(at(model, "A1")?.dxf).toEqual({
      font: { color: "#9C0006", bold: true, italic: true },
      fill: { kind: "solid", color: "#FFC7CE" },
      borders: { top: { style: "thin" } },
    });
  });

  it("ends the walk at a true rule with stopIfTrue — but a colour scale, data bar or icon set's stopIfTrue is ignored", () => {
    const stopped = book({ A1: 200 }, [
      format("A1", rule({ kind: "cellIs", operator: "greaterThan", formulas: ["100"] }, { dxfId: 0, priority: 1, stopIfTrue: true })),
      format("A1", rule({ kind: "expression", formula: "TRUE" }, { dxfId: 2, priority: 2 })),
      format("A1", rule({ kind: "dataBar", bar: BAR }, { priority: 3 })),
    ]);
    expect(at(stopped, "A1")).toEqual({ dxf: RED });
    const ignored = book({ A1: 200, A2: 100 }, [
      format("A1:A2", rule({ kind: "dataBar", bar: BAR }, { priority: 1, stopIfTrue: true })),
      format("A1:A2", rule({ kind: "expression", formula: "TRUE" }, { dxfId: 0, priority: 2 })),
    ]);
    expect(at(ignored, "A1")).toMatchObject({ dxf: RED, dataBar: { widthPct: 100 } });
  });

  it("keeps one fill — a dxf fill or a colour-scale colour, whichever ranks higher — and the first data bar and icon", () => {
    const scaleFirst = book({ A1: 1, A2: 3 }, [
      format("A1:A2", rule({ kind: "colorScale", stops: [{ at: { type: "min" }, color: "#000000" }, { at: { type: "max" }, color: "#FFFFFF" }] }, { priority: 1 })),
      format("A1:A2", rule({ kind: "expression", formula: "TRUE" }, { dxfId: 0, priority: 2 })),
      format("A1:A2", rule({ kind: "dataBar", bar: { ...BAR, color: "#00FF00" } }, { priority: 3 })),
      format("A1:A2", rule({ kind: "dataBar", bar: BAR }, { priority: 4 })),
      format("A1:A2", rule({ kind: "iconSet", icons: { set: "3Flags", thresholds: defaultIconThresholds(3), reverse: false, showValue: true } }, { priority: 5 })),
      format("A1:A2", rule({ kind: "iconSet", icons: { set: "3Arrows", thresholds: defaultIconThresholds(3), reverse: false, showValue: true } }, { priority: 6 })),
    ]);
    expect(at(scaleFirst, "A2")).toMatchObject({ scaleColor: "#ffffff", dxf: { font: { color: "#9C0006" } }, dataBar: { color: "#00ff00" }, icon: { set: "3Flags", index: 2 } });
    expect(at(scaleFirst, "A2")?.dxf?.fill).toBeUndefined();
    const fillFirst = book({ A1: 1, A2: 3 }, [
      format("A1:A2", rule({ kind: "expression", formula: "TRUE" }, { dxfId: 0, priority: 1 })),
      format("A1:A2", rule({ kind: "colorScale", stops: [{ at: { type: "min" }, color: "#000000" }, { at: { type: "max" }, color: "#FFFFFF" }] }, { priority: 2 })),
    ]);
    expect(at(fillFirst, "A2")).toEqual({ dxf: RED });
  });
});

describe("layerDxf", () => {
  it("replaces exactly the properties the dxf sets", () => {
    const base = { font: { family: "Arial", sizePt: 12, bold: true as const, underline: "double" as const, color: "#111111" }, fill: { kind: "solid" as const, color: "#EEEEEE" }, numberFormat: "0.00", builtinFormatId: 2 };
    expect(layerDxf(base, { font: { bold: false, underline: "none", strike: true } })).toEqual({
      font: { family: "Arial", sizePt: 12, color: "#111111", strike: true },
      fill: { kind: "solid", color: "#EEEEEE" },
      numberFormat: "0.00",
      builtinFormatId: 2,
    });
    expect(layerDxf(base, { fill: { kind: "solid", color: "#FFC7CE" }, borders: { left: { style: "thin", color: "#9C0006" } }, numberFormat: "0%" })).toEqual({
      font: base.font,
      fill: { kind: "solid", color: "#FFC7CE" },
      borders: { left: { style: "thin", color: "#9C0006" } },
      numberFormat: "0%",
    });
    expect(base.font.bold).toBe(true);
  });
});

// --- commands ---------------------------------------------------------------------------------------

describe("commands", () => {
  function commandBook(): SpreadsheetModel {
    const model = book({ B5: 1 }, [
      { ...format("B5:B20", rule({ kind: "expression", formula: "B5>$Z$1" }, { id: "x", priority: 3, dxfId: 0 })), id: "fx", sourceXml: "<x/>" },
      { ...format("C1:C9", rule({ kind: "dataBar", bar: BAR }, { id: "y", priority: 7 })), id: "fy", sourceXml: "<y/>" },
    ]);
    return model;
  }

  it("renumbers priorities densely, keeping a reader's sourceXml only on request", () => {
    const formats = commandBook().sheets[0]!.conditionalFormats;
    const renumbered = renumberPriorities(formats);
    expect(rulesByPriority(renumbered).map((entry) => [entry.rule.id, entry.rule.priority])).toEqual([["x", 1], ["y", 2]]);
    expect(renumbered.every((each) => each.sourceXml === undefined)).toBe(true);
    expect(renumberPriorities(formats, { keepSource: true }).map((each) => each.sourceXml)).toEqual(["<x/>", "<y/>"]);
    expect(renumberPriorities(renumbered)).toBe(renumbered);
  });

  it("adds a format at priority 1, copy-on-write, marking the sheet dirty", () => {
    const model = commandBook();
    const before = model.sheets[0]!;
    const { formatId, ruleId } = addConditionalFormat(model, "s1", [range("B2:B13")], { kind: "cellIs", operator: "greaterThan", formulas: ["120"], dxfId: 1 });
    const sheet = model.sheets[0]!;
    expect(sheet).not.toBe(before);
    expect(before.conditionalFormats).toHaveLength(2);
    expect(sheet.dirty).toBe(true);
    expect(rulesByPriority(sheet.conditionalFormats).map((entry) => [entry.rule.id, entry.rule.priority])).toEqual([[ruleId, 1], ["x", 2], ["y", 3]]);
    expect(sheet.conditionalFormats.at(-1)).toMatchObject({ id: formatId, ranges: [range("B2:B13")] });
    expect(sheet.conditionalFormats.some((each) => each.sourceXml !== undefined)).toBe(false);
  });

  it("updates, re-flags stop-if-true, moves and deletes rules, removing a format left empty", () => {
    const model = commandBook();
    updateRule(model, "s1", "x", { kind: "cellIs", operator: "lessThan", formulas: ["5"], dxfId: 2 });
    expect(model.sheets[0]!.conditionalFormats[0]!.rules[0]).toMatchObject({ id: "x", priority: 1, kind: "cellIs", formulas: ["5"] });
    setStopIfTrue(model, "s1", "x", true);
    expect(model.sheets[0]!.conditionalFormats[0]!.rules[0]!.stopIfTrue).toBe(true);
    moveRule(model, "s1", "x", "down");
    expect(rulesByPriority(model.sheets[0]!.conditionalFormats).map((entry) => entry.rule.id)).toEqual(["y", "x"]);
    const unchanged = model.sheets;
    moveRule(model, "s1", "y", "up");
    expect(model.sheets).toBe(unchanged);
    deleteRule(model, "s1", "y");
    expect(model.sheets[0]!.conditionalFormats.map((each) => each.id)).toEqual(["fx"]);
    expect(model.sheets[0]!.conditionalFormats[0]!.rules[0]!.priority).toBe(1);
  });

  it("sets a format's ranges and clears rules from ranges, re-anchoring a format that lost its origin", () => {
    const model = commandBook();
    setFormatRanges(model, "s1", "fy", [range("D1:D9")]);
    expect(model.sheets[0]!.conditionalFormats[1]!.ranges).toEqual([range("D1:D9")]);
    expect(clearRulesInRanges(model, "s1", [range("B5:B6")])).toBe(true);
    expect(model.sheets[0]!.conditionalFormats[0]).toMatchObject({ ranges: [range("B7:B20")], rules: [{ formula: "B7>$Z$1" }] });
    expect(clearRulesInRanges(model, "s1", [range("A1:A2")])).toBe(false);
    expect(clearRulesInRanges(model, "s1", [range("A1:Z100")])).toBe(true);
    expect(model.sheets[0]!.conditionalFormats).toEqual([]);
  });

  it("subtractRanges leaves the uncovered rectangles in source order", () => {
    expect(subtractRanges([range("B2:D4")], range("C3"))).toEqual([range("B2:D2"), range("B3"), range("D3"), range("B4:D4")]);
    expect(subtractRanges([range("A1:A10"), range("C1:C2")], [range("A1:A3"), range("A9:C9")])).toEqual([range("A4:A8"), range("A10"), range("C1:C2")]);
    const untouched = [range("F1:F2")];
    expect(subtractRanges(untouched, range("A1"))[0]).toBe(untouched[0]);
    expect(subtractRanges([range("A1:B2")], range("A1:C3"))).toEqual([]);
  });
});
