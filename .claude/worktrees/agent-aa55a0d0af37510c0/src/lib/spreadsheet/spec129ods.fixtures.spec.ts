// SPEC-129's ODS conditional-formatting fixtures (spec129-cf.ods /
// spec129-legacy.ods, built by fixtures/build.ts): the committed e2e files
// are the builders' bytes and whole packages, the producer shapes are
// really in the parts (Calc's `calcext:conditional-formats` after the rows
// over named styles in styles.xml, "Accent 3" encoded `Accent_20_3`; an
// older producer's `style:map`s on automatic styles over a repeated run and
// a column default), and both read as SPEC129_CF_ODS / SPEC129_LEGACY_ODS
// record — down to the e2e's claim that B3 shows Bad's fill, evaluated
// through createCfEvaluator.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { DEFAULT_THEME } from "../core/theme";
import { unzip } from "../core/zip";
import { parseA1, parseRange, type GridRange } from "./address";
import { createCfEvaluator } from "./cfEvaluator";
import type { CfRule, ConditionalFormat } from "./conditionalFormats";
import { buildSpec129CfOdsBytes, buildSpec129LegacyOdsBytes, SPEC129_CF_ODS, SPEC129_LEGACY_ODS } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import { storedValueContext } from "./formula/evaluate";
import type { SpreadsheetModel } from "./model";
import { readSpreadsheetBytes } from "./open";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const decoder = new TextDecoder();
const range = (a1: string): GridRange => parseRange(a1)!;

/** A rule as the facts state it: no id, priority or dxf index. */
function shape(rule: CfRule): Record<string, unknown> {
  const { id: _id, priority: _priority, dxfId: _dxf, ...rest } = rule;
  return rest;
}

/** The style name each rule of `formats` applies (undefined: none). */
function appliedStyles(model: SpreadsheetModel, formats: readonly ConditionalFormat[]): (string | undefined)[][] {
  return formats.map((format) => format.rules.map((rule) => (rule.dxfId === undefined ? undefined : model.dxfs[rule.dxfId]?.sourceStyleName)));
}

/** The conditional-format result at an A1 address of a sheet. */
function resultAt(model: SpreadsheetModel, sheetIndex: number, a1: string) {
  const evaluator = createCfEvaluator(model, model.sheets[sheetIndex]!.id, { base: storedValueContext(model), theme: DEFAULT_THEME });
  const ref = parseA1(a1)!;
  return evaluator.resultFor(ref.row, ref.col);
}

describe("spec129-cf.ods", () => {
  const bytes = buildSpec129CfOdsBytes();

  it("is the builder's bytes, a whole package, with Calc's rules after each table's rows over styles.xml's named styles", async () => {
    expect(equalBytes(new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec129-cf.ods"))), await bytes)).toBe(true);
    expect(await packageIntegrity(await bytes)).toEqual([]);
    const parts = await unzip(await bytes);
    const content = decoder.decode(parts.get("content.xml"));
    expect(content).toContain(`<calcext:conditional-formats>${SPEC129_CF_ODS.scoresFormats.join("")}</calcext:conditional-formats></table:table>`);
    expect(content).toContain(`<calcext:conditional-formats>${SPEC129_CF_ODS.moreFormats.join("")}</calcext:conditional-formats></table:table>`);
    expect(content).not.toContain("<style:map style:condition=\"cell-content");
    const styles = decoder.decode(parts.get("styles.xml"));
    expect(styles).toContain(`<style:style style:name="Accent_20_3" style:display-name="Accent 3" style:family="table-cell" style:parent-style-name="Accent">`);
    for (const name of ["Bad", "Good", "Status", "Accent"]) {
      expect(styles).toContain(`<style:style style:name="${name}" style:family="table-cell"`);
    }
  });

  it("reads every Scores and More format: ranges, kinds, the applied named styles as dxfs, each format's source", async () => {
    const model = await readSpreadsheetBytes(await bytes, "ods", "fixture");
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(SPEC129_CF_ODS.tableNames);
    expect(model.dxfs).toEqual(SPEC129_CF_ODS.dxfs);
    const [scores, more] = model.sheets.map((sheet) => sheet.conditionalFormats);
    expect(scores!.map((format) => format.sourceXml)).toEqual(SPEC129_CF_ODS.scoresFormats);
    expect(more!.map((format) => format.sourceXml)).toEqual(SPEC129_CF_ODS.moreFormats);
    expect(scores!.map((format) => format.ranges)).toEqual(["B2:B13", "A2:H13", "C2:C13", "D2:D13", "E2:E13", "F2:F13", "G2:G13"].map((a1) => [range(a1)]));
    expect(scores!.map((format) => format.rules.map(shape))).toEqual([
      [{ kind: "cellIs", operator: "greaterThan", formulas: ["100"] }],
      [{ kind: "expression", formula: `$H2="late"` }],
      [
        {
          kind: "dataBar",
          bar: {
            min: { type: "autoMin" },
            max: { type: "autoMax" },
            color: "#638ec6",
            negativeColor: "#ff0000",
            axisColor: "#000000",
            gradient: false,
            axis: "automatic",
            direction: "context",
            minLengthPct: 0,
            maxLengthPct: 100,
            showValue: true,
          },
        },
      ],
      [
        {
          kind: "iconSet",
          icons: {
            set: "3Arrows",
            thresholds: [
              { at: { type: "percent", value: "0" }, gte: true },
              { at: { type: "percent", value: "33" }, gte: false },
              { at: { type: "percent", value: "67" }, gte: false },
            ],
            reverse: false,
            showValue: true,
          },
        },
      ],
      [
        {
          kind: "colorScale",
          stops: [
            { at: { type: "min" }, color: "#f8696b" },
            { at: { type: "percentile", value: "50" }, color: "#ffeb84" },
            { at: { type: "max" }, color: "#63be7b" },
          ],
        },
      ],
      [
        {
          kind: "iconSet",
          icons: {
            set: "3Stars",
            thresholds: [0, 33, 67].map((value) => ({ at: { type: "percent", value: String(value) }, gte: true })),
            reverse: false,
            showValue: true,
          },
        },
      ],
      // The custom icon set: carried verbatim, never evaluated.
      [{ kind: "preserved", where: "ods", xml: /<calcext:icon-set [\s\S]*<\/calcext:icon-set>/.exec(SPEC129_CF_ODS.scoresFormats[6])![0] }],
    ]);
    expect(more!.map((format) => format.ranges)).toEqual(["A2:A7", "B2:B7", "C2:C7"].map((a1) => [range(a1)]));
    expect(more!.map((format) => format.rules.map(shape))).toEqual([
      [
        { kind: "top10", bottom: false, percent: true, rank: 10 },
        { kind: "aboveAverage", below: false, equalAverage: false },
        { kind: "duplicates", unique: false },
      ],
      [
        { kind: "text", operator: "contains", text: "late" },
        { kind: "expression", formula: "LEN(TRIM(B2))=0" },
      ],
      [{ kind: "timePeriod", period: "last7Days" }],
    ]);
    expect(appliedStyles(model, scores!)).toEqual([["Bad"], ["Accent_20_3"], [undefined], [undefined], [undefined], [undefined], [undefined]]);
    expect(appliedStyles(model, more!)).toEqual([["Good", "Accent_20_3", "Bad"], ["Bad", "Accent_20_3"], ["Good"]]);
    // Priorities run 1..n in document order across each table.
    expect(scores!.flatMap((format) => format.rules.map((rule) => rule.priority))).toEqual([1, 2, 3, 4, 5, 6, 7]);
    expect(more!.flatMap((format) => format.rules.map((rule) => rule.priority))).toEqual([1, 2, 3, 4, 5, 6]);
  });

  it("evaluates as the e2e expects: B3 (140 > 100) shows Bad's fill and text, B4 (90) none; D2 the top arrow; C3's bar negative", async () => {
    const model = await readSpreadsheetBytes(await bytes, "ods", "fixture");
    expect(SPEC129_CF_ODS.records[1]![1]).toBe(140);
    expect(SPEC129_CF_ODS.records[2]![1]).toBe(90);
    const b3 = resultAt(model, 0, "B3");
    expect(b3?.dxf?.fill).toEqual({ kind: "solid", color: SPEC129_CF_ODS.styles.bad.fill });
    expect(b3?.dxf?.font?.color).toBe(SPEC129_CF_ODS.styles.bad.text);
    expect(resultAt(model, 0, "B4")?.dxf).toBeUndefined();
    // A late row takes Accent 3 (its fill and Accent's weight) where Bad does not apply.
    expect(resultAt(model, 0, "A2")?.dxf).toEqual({ font: { bold: true }, fill: { kind: "solid", color: "#dddddd" } });
    expect(resultAt(model, 0, "D2")?.icon).toEqual({ set: "3Arrows", index: 2 });
    expect(resultAt(model, 0, "C3")?.dataBar?.negative).toBe(true);
    expect(resultAt(model, 0, "G2")?.icon).toBeUndefined();
  });
});

describe("spec129-legacy.ods", () => {
  const bytes = buildSpec129LegacyOdsBytes();

  it("is the builder's bytes, a whole package, with style:map conditions only — on automatic styles over a repeated run and a column default", async () => {
    expect(equalBytes(new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec129-legacy.ods"))), await bytes)).toBe(true);
    expect(await packageIntegrity(await bytes)).toEqual([]);
    const content = decoder.decode((await unzip(await bytes)).get("content.xml"));
    expect(content).not.toContain("conditional-format");
    expect(content.match(/<style:map style:condition="(?:cell-content|is-true-formula)/g)).toHaveLength(3);
    expect(content).toContain(`table:number-rows-repeated="${SPEC129_LEGACY_ODS.emptyRun}"><table:table-cell/><table:table-cell table:style-name="ce6" table:number-columns-repeated="3"/>`);
    expect(content).toContain(`<table:table-column table:style-name="co1" table:default-cell-style-name="ce7"/>`);
  });

  it("reads one format per mapped style over rectangles, never cell by cell, and evaluates them", async () => {
    const model = await readSpreadsheetBytes(await bytes, "ods", "fixture");
    const formats = model.sheets[0]!.conditionalFormats;
    expect(formats.map((format) => format.ranges)).toEqual(SPEC129_LEGACY_ODS.formats.map((format) => format.ranges));
    expect(formats.map((format) => format.rules.map(shape))).toEqual(SPEC129_LEGACY_ODS.formats.map((format) => format.rules.map(({ style: _style, ...rule }) => rule)));
    expect(appliedStyles(model, formats)).toEqual(SPEC129_LEGACY_ODS.formats.map((format) => format.rules.map((rule) => rule.style)));
    expect(formats.every((format) => format.sourceXml === undefined)).toBe(true);

    expect(resultAt(model, 0, "B2")?.dxf?.fill).toEqual({ kind: "solid", color: "#ffc7ce" }); // 120 > 100
    expect(resultAt(model, 0, "C2")?.dxf?.fill).toEqual({ kind: "solid", color: "#c6efce" }); // 80 in 50..100
    expect(resultAt(model, 0, "D2")?.dxf).toBeUndefined(); // 40
    expect(resultAt(model, 0, "E2")?.dxf?.fill).toEqual({ kind: "solid", color: "#ffc7ce" }); // "late"
    expect(resultAt(model, 0, "E3")?.dxf).toBeUndefined(); // "ok"
  });
});
