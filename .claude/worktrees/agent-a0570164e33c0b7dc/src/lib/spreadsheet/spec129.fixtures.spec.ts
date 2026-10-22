// SPEC-129's xlsx conditional-formatting fixture (spec129-cf.xlsx, built by
// fixtures/build.ts): the committed e2e file is the builder's bytes, the
// package is whole, the producer shapes the reader and writer are proven on
// are really in the parts — `mc:Ignorable="x14ac xr xr2 xr3"`, `xr:uid`s on
// the blocks, three bgColor-only dxfs (dxf 0 Excel's light red), the x14
// data-bar twin sharing its GUID, the x14-only 3Stars and the custom icon
// set, the sparkline ext after the conditional-formatting one — and the
// facts the e2e flow leans on hold (only B3 above 100, D2 the largest, rows
// 3 and 4 on time, C2 clearly negative). The model the reader makes of the
// file is asserted once here too, as the e2e opens it, and so is what
// createCfEvaluator makes of it where the e2e looks: B3 red, the C column's
// bars by value with C2's left of the axis, D2 the green up arrow.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { DEFAULT_THEME, resolveColor } from "../core/theme";
import { unzip } from "../core/zip";
import { formatRange, parseA1 } from "./address";
import { createCfEvaluator } from "./cfEvaluator";
import type { CfRule } from "./conditionalFormats";
import { buildSpec129CfXlsxBytes, SPEC129_CF_XLSX } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import { storedValueContext } from "./formula/evaluate";
import { readSpreadsheetBytes } from "./open";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const decoder = new TextDecoder();
const F = SPEC129_CF_XLSX;

async function parts(): Promise<Map<string, Uint8Array>> {
  return unzip(await buildSpec129CfXlsxBytes());
}

const text = (all: Map<string, Uint8Array>, name: string): string => decoder.decode(all.get(name));

/** A rule without its minted id, sorted by priority with its format's sqref. */
function summary(rule: CfRule, sqref: string): Record<string, unknown> {
  const { id: _id, ...rest } = rule;
  return { sqref, ...rest };
}

describe("spec129-cf.xlsx", () => {
  it("is the builder's bytes, whole, with Excel's parts in order", async () => {
    const committed = new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec129-cf.xlsx")));
    expect(equalBytes(committed, await buildSpec129CfXlsxBytes())).toBe(true);
    expect(await packageIntegrity(committed)).toEqual([]);
    expect([...(await parts()).keys()]).toEqual([...F.parts]);
  });

  it("holds Excel's shapes: the ignorable namespaces, xr:uid blocks, bgColor-only dxfs, the x14 twin, the x14-only and custom rules, the sparkline ext after the CF ext", async () => {
    const all = await parts();
    const scores = text(all, F.scores.partName);
    expect(scores).toContain(`mc:Ignorable="x14ac xr xr2 xr3"`);
    const blocks = [...scores.matchAll(/<conditionalFormatting [^>]*>/g)].map((match) => match[0]);
    expect(blocks).toHaveLength(F.scores.blocks.length);
    expect(blocks.every((block) => /xr:uid="\{[0-9A-F-]{36}\}"/.test(block))).toBe(true);
    expect(scores).toContain(F.scores.blocks.join("") + `<pageMargins`);
    // The CF ext first, then the sparkline ext, verbatim, in one extLst.
    expect(scores).toContain(`<extLst>${F.scores.cfExt}${F.scores.sparklineExt}</extLst></worksheet>`);
    expect(scores.split(F.scores.dataBar.guid)).toHaveLength(3);
    expect(F.scores.cfExt).toContain(`custom="1"`);
    expect(F.scores.cfExt).toContain(`<x14:cfIcon iconSet="3Flags" iconId="0"/>`);
    const styles = text(all, "xl/styles.xml");
    expect(styles).toContain(`<dxfs count="3">${F.dxfXml.join("")}</dxfs><tableStyles`);
    expect(F.dxfXml.every((dxf) => /<patternFill><bgColor [^>]*\/><\/patternFill>/.test(dxf) && !dxf.includes("patternType"))).toBe(true);
    expect(F.dxfXml[0]).toBe(`<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>`);
    const more = text(all, F.more.partName);
    for (const block of F.more.blocks) {
      expect(more).toContain(block);
    }
  });

  it("keeps the facts the e2e flow leans on", () => {
    const S = F.scores;
    expect(S.scores.filter((score) => score > 100)).toEqual([S.cellIs.hotValue]);
    expect(S.scores[1]).toBe(140);
    expect(S.scores[2]).toBe(90);
    expect(Math.max(...S.trends)).toBe(S.trends[0]);
    expect([S.statuses[1], S.statuses[2]]).toEqual(["on time", "on time"]);
    expect(new Set(S.changes).size).toBe(S.changes.length);
    expect(S.changes.flatMap((change, i) => (change < 0 ? [`C${i + 2}`] : []))).toEqual([...S.negativeCells]);
    expect(Math.min(...S.changes)).toBe(S.changes[0]);
  });

  it("reads as intended: dxfs, every rule in one priority order, the twin overlaid, the x14-only and preserved rules", async () => {
    const model = await readSpreadsheetBytes(await buildSpec129CfXlsxBytes(), "xlsx", "book");
    expect(model.dxfs.map(({ sourceXml: _xml, ...dxf }) => dxf)).toEqual(F.dxfs);
    expect(model.dxfs.map((dxf) => dxf.sourceXml)).toEqual(F.dxfXml);
    const scores = model.sheets[0]!;
    const rules = scores.conditionalFormats
      .flatMap((format) => format.rules.map((rule) => summary(rule, format.ranges.map(formatRange).join(" "))))
      .toSorted((a, b) => (a.priority as number) - (b.priority as number));
    expect(rules).toEqual([
      { sqref: "B2:B13", priority: 1, dxfId: 0, kind: "cellIs", operator: "greaterThan", formulas: ["100"] },
      {
        sqref: "C2:C13",
        priority: 2,
        kind: "dataBar",
        x14Id: F.scores.dataBar.guid,
        bar: {
          min: { type: "autoMin" },
          max: { type: "autoMax" },
          color: "#638ec6",
          gradient: false,
          axis: "automatic",
          direction: "context",
          minLengthPct: 0,
          maxLengthPct: 100,
          showValue: true,
          negativeColor: "#ff0000",
          axisColor: "#000000",
        },
      },
      {
        sqref: "D2:D13",
        priority: 3,
        kind: "iconSet",
        icons: {
          set: "3Arrows",
          reverse: false,
          showValue: true,
          thresholds: [
            { at: { type: "percent", value: "0" }, gte: true },
            { at: { type: "percent", value: "33" }, gte: false },
            { at: { type: "percent", value: "67" }, gte: false },
          ],
        },
      },
      {
        sqref: "E2:E13",
        priority: 4,
        kind: "colorScale",
        stops: [
          { at: { type: "min" }, color: { theme: "accent2", lumMod: 1 - 0.39997558519241921, lumOff: 0.39997558519241921, xlsxTint: "0.39997558519241921" } },
          { at: { type: "percentile", value: "50" }, color: { theme: "lt1", lumMod: 1 - 0.249977111117893, xlsxTint: "-0.249977111117893" } },
          { at: { type: "max" }, color: { theme: "accent6" } },
        ],
      },
      {
        sqref: "F2:F13",
        priority: 5,
        kind: "iconSet",
        x14: true,
        x14Id: F.scores.stars.guid,
        icons: {
          set: "3Stars",
          reverse: false,
          showValue: true,
          thresholds: [
            { at: { type: "percent", value: "0" }, gte: true },
            { at: { type: "percent", value: "33" }, gte: true },
            { at: { type: "percent", value: "67" }, gte: true },
          ],
        },
      },
      { sqref: "G2:G13", priority: 6, kind: "preserved", where: "x14", xml: F.scores.custom.xml },
      { sqref: "A2:H13", priority: 7, dxfId: 1, stopIfTrue: true, kind: "expression", formula: `$H2="late"` },
    ]);
    // The base blocks keep their source text; the x14-only ones theirs.
    expect(scores.conditionalFormats.slice(0, 5).map((format) => format.sourceXml)).toEqual([...F.scores.blocks]);
    expect(scores.conditionalFormats[5]!.sourceXml).toMatch(/^<x14:conditionalFormatting xmlns:xm="[^"]+"><x14:cfRule type="iconSet" priority="5"/);

    const more = model.sheets[1]!.conditionalFormats.flatMap((format) => format.rules.map((rule) => summary(rule, format.ranges.map(formatRange).join(" "))));
    expect(more).toEqual([
      { sqref: "A2:A11", priority: 1, dxfId: 2, kind: "top10", bottom: false, percent: true, rank: 20 },
      { sqref: "B2:B11", priority: 2, dxfId: 1, kind: "aboveAverage", below: false, equalAverage: false, stdDev: 1 },
      { sqref: "C2:C11", priority: 3, dxfId: 0, kind: "duplicates", unique: false },
      { sqref: "C2:C11", priority: 4, dxfId: 2, kind: "text", operator: "contains", text: "an" },
      { sqref: "D2:D11", priority: 5, dxfId: 1, kind: "blanks", not: false },
      { sqref: "E2:E11", priority: 6, dxfId: 0, kind: "timePeriod", period: "last7Days" },
    ]);
  });

  it("evaluates where the e2e looks: B3 light red with dark red text, B4 plain, the C bars growing with their values and C2's left of the axis, D2 the green up arrow", async () => {
    const model = await readSpreadsheetBytes(await buildSpec129CfXlsxBytes(), "xlsx", "book");
    const theme = model.theme ?? DEFAULT_THEME;
    const sheet = model.sheets[0]!;
    const evaluator = createCfEvaluator(model, sheet.id, { base: storedValueContext(model), theme });
    const at = (a1: string) => {
      const ref = parseA1(a1)!;
      return evaluator.resultFor(ref.row, ref.col);
    };
    const hot = at(F.scores.cellIs.hot)?.dxf;
    expect(hot?.fill?.kind === "solid" ? resolveColor(hot.fill.color, theme) : undefined).toBe("#ffc7ce");
    expect(resolveColor(hot!.font!.color!, theme)).toBe("#9c0006");
    expect(at(F.scores.cellIs.coolCell)?.dxf?.fill).toBeUndefined();

    const bars = F.scores.changes.map((value, i) => ({ value, bar: at(`C${i + 2}`)?.dataBar }));
    expect(bars.every(({ bar }) => bar !== undefined)).toBe(true);
    const positive = bars.filter(({ value }) => value > 0).toSorted((a, b) => a.value - b.value);
    for (let i = 1; i < positive.length; i += 1) {
      expect(positive[i]!.bar!.widthPct).toBeGreaterThan(positive[i - 1]!.bar!.widthPct);
    }
    const c2 = bars[0]!.bar!;
    expect(c2.negative).toBe(true);
    expect(c2.axisPct).toBeGreaterThan(0);
    expect(c2.startPct + c2.widthPct).toBeLessThanOrEqual(c2.axisPct!);
    expect(positive.every(({ bar }) => bar!.startPct >= c2.axisPct!)).toBe(true);

    expect(at(F.scores.arrows.topCell)?.icon).toEqual({ set: "3Arrows", index: 2 });
  });
});
