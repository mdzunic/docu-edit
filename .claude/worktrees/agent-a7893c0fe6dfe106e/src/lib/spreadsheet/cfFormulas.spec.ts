// cfFormulas.ts: which texts of a conditional-format rule are formulas —
// the value-object decision exactly as the evaluator makes it, a preserved
// xlsx rule's <formula>/<xm:f> texts — the relative/absolute split,
// x14 re-placement, and copy-on-write identity.
import { describe, expect, it } from "vitest";
import { isCfValueFormula, mapFormatFormulas, mapFormatsFormulas, mapRuleFormulas, mapXlsxPreservedFormulas, type RuleHost } from "./cfFormulas";
import type { CfRule, CfValue, ConditionalFormat } from "./conditionalFormats";

const HOST: RuleHost = { format: "xlsx", sheetName: "Sheet1" };
/** Marks every formula it maps. */
const mark = (formula: string): string => `(${formula})`;
const same = (formula: string): string => formula;

describe("isCfValueFormula — the evaluator's decision", () => {
  it("a formula value always is; min/max/autoMin/autoMax never are", () => {
    expect(isCfValueFormula({ type: "formula", value: "10" })).toBe(true);
    for (const type of ["min", "max", "autoMin", "autoMax"] as const) {
      expect(isCfValueFormula({ type, value: "$A$1" })).toBe(false);
    }
    expect(isCfValueFormula({ type: "formula" })).toBe(false);
  });

  it("num/percent/percentile: a formula unless blank or a plain number (signs, decimals, exponents, spaces)", () => {
    const cases: [CfValue["type"], string, boolean][] = [
      ["num", "10", false],
      ["num", " -1.5e3 ", false],
      ["percent", ".5", false],
      ["percentile", "", false],
      ["num", "$A$1", true],
      ["percent", "=$A$1*2", true],
      ["percentile", "MAX(0,$B$1)", true],
    ];
    for (const [type, value, formula] of cases) {
      expect(isCfValueFormula({ type, value }), `${type} ${value}`).toBe(formula);
    }
  });
});

describe("mapRuleFormulas", () => {
  const maps = { relative: (formula: string): string => `R${formula}`, absolute: (formula: string): string => `A${formula}` };

  it("maps cellIs/expression formulas as relative, value objects as absolute (keeping an editor's leading =)", () => {
    const cellIs: CfRule = { id: "c", priority: 1, kind: "cellIs", operator: "between", formulas: ["1", "B2"] };
    expect(mapRuleFormulas(cellIs, maps, HOST)).toMatchObject({ formulas: ["R1", "RB2"] });
    const scale: CfRule = {
      id: "s",
      priority: 2,
      kind: "colorScale",
      stops: [
        { at: { type: "num", value: "=$A$1" }, color: "#ff0000" },
        { at: { type: "percentile", value: "50" }, color: "#ffff00" },
        { at: { type: "max" }, color: "#00ff00" },
      ],
    };
    const mapped = mapRuleFormulas(scale, maps, HOST);
    expect(mapped).toMatchObject({ stops: [{ at: { type: "num", value: "=A$A$1" } }, { at: { type: "percentile", value: "50" } }, { at: { type: "max" } }] });
    if (mapped.kind === "colorScale" && scale.kind === "colorScale") {
      expect(mapped.stops[1]).toBe(scale.stops[1]);
    }
  });

  it("maps a data bar's min/max and an icon set's thresholds", () => {
    const bar: CfRule = {
      id: "b",
      priority: 1,
      kind: "dataBar",
      bar: { min: { type: "formula", value: "$C$1" }, max: { type: "autoMax" }, color: "#638EC6", gradient: true, axis: "automatic", direction: "context", minLengthPct: 10, maxLengthPct: 90, showValue: true },
    };
    expect(mapRuleFormulas(bar, maps, HOST)).toMatchObject({ bar: { min: { value: "A$C$1" }, max: { type: "autoMax" } } });
    const icons: CfRule = {
      id: "i",
      priority: 1,
      kind: "iconSet",
      icons: { set: "3Arrows", reverse: false, showValue: true, thresholds: [{ at: { type: "percent", value: "0" }, gte: true }, { at: { type: "num", value: "$D$1" }, gte: true }, { at: { type: "percent", value: "67" }, gte: false }] },
    };
    expect(mapRuleFormulas(icons, maps, HOST)).toMatchObject({ icons: { thresholds: [{ at: { value: "0" } }, { at: { value: "A$D$1" } }, { at: { value: "67" } }] } });
  });

  it("maps a preserved xlsx rule's <formula> and <xm:f> texts through the xlsx dialect; an ods one, and formula-less kinds, stay", () => {
    const base: CfRule = { id: "p", priority: 1, kind: "preserved", where: "base", xml: `<cfRule type="expression" priority="1"><formula>$A$1&gt;0</formula></cfRule>` };
    expect(mapRuleFormulas(base, maps, HOST)).toMatchObject({ xml: `<cfRule type="expression" priority="1"><formula>A$A$1&gt;0</formula></cfRule>` });
    const x14: CfRule = { id: "x", priority: 2, kind: "preserved", where: "x14", xml: `<x14:cfRule type="iconSet" priority="2" id="{G}"><x14:iconSet custom="1"><x14:cfvo type="num"><xm:f>Sheet2!$B$1</xm:f></x14:cfvo></x14:iconSet></x14:cfRule>` };
    expect(mapRuleFormulas(x14, maps, HOST)).toMatchObject({ xml: x14.xml.replace("<xm:f>Sheet2!$B$1</xm:f>", "<xm:f>ASheet2!$B$1</xm:f>") });
    const ods: CfRule = { id: "o", priority: 3, kind: "preserved", where: "ods", xml: `<calcext:condition calcext:value="formula-is([.A1])"/>` };
    expect(mapRuleFormulas(ods, maps, HOST)).toBe(ods);
    const top: CfRule = { id: "t", priority: 4, kind: "top10", bottom: false, percent: false, rank: 10 };
    expect(mapRuleFormulas(top, maps, HOST)).toBe(top);
  });

  it("re-places a changed cellIs/expression rule in x14 as its formulas now do or do not reference another sheet", () => {
    const rule: CfRule = { id: "e", priority: 1, kind: "expression", formula: "A1>$Z$1" };
    const qualified = mapRuleFormulas(rule, { relative: (formula) => formula.replace("$Z$1", "Other!$Z$1"), absolute: same }, HOST);
    expect(qualified).toMatchObject({ formula: "A1>Other!$Z$1", x14: true });
    const back = mapRuleFormulas(qualified, { relative: () => "A1>#REF!", absolute: same }, HOST);
    expect(back.x14).toBeUndefined();
    expect(mapRuleFormulas(qualified, { relative: () => "A1>Other!$Z$2", absolute: same }, { format: "ods", sheetName: "Sheet1" }).x14).toBeUndefined();
  });
});

describe("mapFormatFormulas / mapFormatsFormulas — copy-on-write", () => {
  const format: ConditionalFormat = { id: "f", ranges: [{ top: 0, left: 0, bottom: 9, right: 0 }], rules: [{ id: "r", priority: 1, kind: "expression", formula: "A1>0" }], sourceXml: "<conditionalFormatting/>" };

  it("returns the same format and array when nothing changes", () => {
    expect(mapFormatFormulas(format, { relative: same, absolute: same }, HOST)).toBe(format);
    const formats = [format];
    expect(mapFormatsFormulas(formats, same, HOST)).toBe(formats);
  });

  it("replaces a changed format without its sourceXml, leaving the original untouched", () => {
    const snapshot = structuredClone(format);
    const [next] = mapFormatsFormulas([format], mark, HOST);
    expect(next).toMatchObject({ id: "f", rules: [{ formula: "(A1>0)" }] });
    expect(next!.sourceXml).toBeUndefined();
    expect(format).toEqual(snapshot);
  });

  it("leaves XML that does not scan as it is", () => {
    expect(mapXlsxPreservedFormulas(`<cfRule><formula>A1</cfRule>`, mark)).toBe(`<cfRule><formula>A1</cfRule>`);
  });
});
