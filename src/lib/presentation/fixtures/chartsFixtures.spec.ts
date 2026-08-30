// SPEC-092 chart fixtures (AC-28): each deck opens to a clustered column, a
// line, a pie and one out-of-catalogue chart (preserved, not parsed).
import { describe, expect, it } from "vitest";
import { readPresentationBytes } from "../open";
import { buildOdpChartsFixtureBytes, buildPptxChartsFixtureBytes, EXPECTED_FIXTURE_CHARTS } from "./build";
import type { ChartFrame, PresentationModel } from "../model";

function charts(model: PresentationModel): ChartFrame[] {
  return model.slides[0].frames.filter((f): f is ChartFrame => f.kind === "chart");
}

function assertFixture(model: PresentationModel, unsupportedMarker: string): void {
  const all = charts(model);
  expect(all).toHaveLength(4);
  const supported = all.filter((c) => c.unsupportedChartXml === undefined);
  const unsupported = all.filter((c) => c.unsupportedChartXml !== undefined);
  expect(supported).toHaveLength(3);
  expect(unsupported).toHaveLength(1);

  const byType = new Map(supported.map((c) => [c.chartType, c]));
  const column = byType.get("columnClustered");
  const line = byType.get("line");
  const pie = byType.get("pie");
  expect(column).toBeDefined();
  expect(line).toBeDefined();
  expect(pie).toBeDefined();

  expect(column!.categories).toEqual([...EXPECTED_FIXTURE_CHARTS.categories]);
  expect(column!.series.map((s) => s.name)).toEqual(["Revenue", "Cost"]);
  expect(column!.series[0].values).toEqual(EXPECTED_FIXTURE_CHARTS.column.series[0].values);
  expect(line!.series[0].values).toEqual(EXPECTED_FIXTURE_CHARTS.line.series[0].values);
  expect(pie!.series[0].values).toEqual(EXPECTED_FIXTURE_CHARTS.pie.series[0].values);

  // The out-of-catalogue chart is preserved verbatim, not parsed (AC-8/9b).
  expect(unsupported[0].unsupportedChartXml).toContain(unsupportedMarker);
}

describe("SPEC-092 chart fixtures (AC-28)", () => {
  it("pptx: opens to three modelled charts + a preserved 3-D chart", async () => {
    const model = await readPresentationBytes(await buildPptxChartsFixtureBytes(), "pptx", "Charts");
    assertFixture(model, EXPECTED_FIXTURE_CHARTS.unsupported.pptxMarker);
  });

  it("odp: opens to three modelled charts + a preserved surface chart", async () => {
    const model = await readPresentationBytes(await buildOdpChartsFixtureBytes(), "odp", "Charts");
    assertFixture(model, EXPECTED_FIXTURE_CHARTS.unsupported.odpMarker);
  });
});
