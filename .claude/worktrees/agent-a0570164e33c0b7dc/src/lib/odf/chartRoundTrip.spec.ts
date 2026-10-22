// SPEC-052: writer-level round trips through the real writeOdt/parseOdt
// pipeline for a chart's embedded object — mirrors
// docx/chartRoundTrip.spec.ts's own conventions.
import { describe, it, expect } from "vitest";
import { parseOdt } from "./reader";
import { writeOdt } from "./writer";
import { unzip } from "../core/zip";
import { equalStructural } from "../equality";
import { equalBytes } from "../core/equality";
import { base64ToBytes, bytesToBase64 } from "../opaque";
import { asChart, createDocument, isChart, type Chart, type DocumentModel } from "../documentModel";

function baseChart(overrides: Partial<Chart> = {}): Chart {
  return {
    type: "chart",
    chartType: "columnClustered",
    categories: ["Q1", "Q2", "Q3"],
    series: [
      { name: "Revenue", values: [10, 20, 30], color: "#4e79a7" },
      { name: "Cost", values: [5, 8, null], color: "#f28e2b" },
    ],
    options: { legend: "right", dataLabels: false, gridlines: true },
    widthCm: 12,
    heightCm: 7.5,
    wrap: "square",
    ...overrides,
  };
}

function modelWithChart(chart: Chart): DocumentModel {
  const model = createDocument("odt-chart-test-doc");
  model.sections[0].blocks = [model.sections[0].blocks[0], chart];
  return model;
}

function rewrap(parsed: Awaited<ReturnType<typeof parseOdt>>, title = "Untitled document"): DocumentModel {
  return {
    id: "odt-chart-preserved-doc",
    title,
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "odt",
    opaque: parsed.opaque,
    sourceXml: parsed.sourceXml,
  };
}

async function openedModelWithChart(
  chart: Chart,
  patchChartXml?: (xml: string) => string,
): Promise<{ model: DocumentModel; chartPartName: string }> {
  const { bytes } = await writeOdt(modelWithChart(chart));
  const parsed = await parseOdt(bytes);
  const model = rewrap(parsed);
  const chartBlock = model.sections[0].blocks[1] as Chart;
  const chartPartName = chartBlock.sourcePart!;
  if (patchChartXml) {
    const original = model.opaque!.find((p) => p.name === chartPartName)!;
    const originalXml = new TextDecoder().decode(base64ToBytes(original.data));
    original.data = bytesToBase64(new TextEncoder().encode(patchChartXml(originalXml)));
  }
  return { model, chartPartName };
}

describe("writeOdt / parseOdt — a chart created fresh", () => {
  it("saves and re-opens with the same chart data (AC-37), with manifest entries for the embedded object", async () => {
    const chart = baseChart();
    const { bytes, dropped } = await writeOdt(modelWithChart(chart));
    expect(dropped).toEqual([]);
    const reopened = await parseOdt(bytes);
    const reopenedChart = asChart(reopened.sections[0].blocks[1]);
    expect(reopenedChart.chartType).toBe("columnClustered");
    expect(reopenedChart.categories).toEqual(chart.categories);
    expect(reopenedChart.series.map((s) => s.name)).toEqual(chart.series.map((s) => s.name));

    const parts = await unzip(bytes);
    expect(parts.has("Object 1/content.xml")).toBe(true);
    const manifest = new TextDecoder().decode(parts.get("META-INF/manifest.xml")!);
    expect(manifest).toContain("Object 1/");
    expect(manifest).toContain("application/vnd.oasis.opendocument.chart");
    const contentXml = new TextDecoder().decode(parts.get("content.xml")!);
    expect(contentXml).toContain("draw:object");
  });
});

describe("writeOdt — an untouched preserved chart (AC-25, AC-26)", () => {
  it("is written back byte-identical on a same-format save", async () => {
    const { model, chartPartName } = await openedModelWithChart(baseChart());
    const { bytes: firstBytes } = await writeOdt(model);
    const firstParts = await unzip(firstBytes);

    const reopenedModel = rewrap(await parseOdt(firstBytes));
    const firstParagraph = reopenedModel.sections[0].blocks[0];
    if (firstParagraph.type === "paragraph") {
      firstParagraph.runs = [{ text: "Edited.", formatting: {} }];
    }
    const { bytes: secondBytes, dropped } = await writeOdt(reopenedModel);
    expect(dropped).toEqual([]);
    const secondParts = await unzip(secondBytes);

    expect(equalBytes(secondParts.get(chartPartName)!, firstParts.get(chartPartName)!)).toBe(true);

    const reopenedAgain = await parseOdt(secondBytes);
    const chartBlock = asChart(reopenedAgain.sections[0].blocks[1]);
    expect(chartBlock.edited).toBeFalsy();
  });
});

describe("writeOdt — an edited chart (AC-27, AC-28)", () => {
  it("regenerates the embedded object from the model and reports what the rewrite lost", async () => {
    const { model, chartPartName } = await openedModelWithChart(baseChart(), (xml) =>
      xml.replace(/(<chart:series[^>]*)\/>/, "$1><chart:regression-curve/></chart:series>"),
    );
    const sourceXml = new TextDecoder().decode(base64ToBytes(model.opaque!.find((p) => p.name === chartPartName)!.data));
    const chart = model.sections[0].blocks[1] as Chart;
    chart.edited = true;
    chart.categories = ["Q1", "Q2", "Q3", "Q4"];
    chart.series = [{ name: "Revenue", values: [10, 20, 30, 40], color: "#4e79a7" }];

    const { bytes, dropped } = await writeOdt(model);
    const parts = await unzip(bytes);
    const writtenXml = new TextDecoder().decode(parts.get(chartPartName)!);
    expect(writtenXml).not.toEqual(sourceXml);
    expect(writtenXml).toContain("Q4");
    expect(writtenXml).not.toContain("chart:regression-curve");

    const chartFeatureDrop = dropped.find((d) => d.kind === "chartFeature");
    expect(chartFeatureDrop, "a chartFeature drop for the lost trendline").toBeDefined();
    expect(chartFeatureDrop!.reason.toLowerCase()).toContain("trendline");
  });
});

describe("writeOdt — deleting a chart (AC-32, AC-33)", () => {
  it("removes the embedded object and its manifest entries on save", async () => {
    const { model, chartPartName } = await openedModelWithChart(baseChart());
    const { bytes: firstBytes } = await writeOdt(model);

    const reopenedModel = rewrap(await parseOdt(firstBytes));
    reopenedModel.sections[0].blocks = reopenedModel.sections[0].blocks.filter((b) => !isChart(b));

    const { bytes: secondBytes } = await writeOdt(reopenedModel);
    const parts = await unzip(secondBytes);
    expect(parts.has(chartPartName)).toBe(false);
    const manifest = new TextDecoder().decode(parts.get("META-INF/manifest.xml")!);
    expect(manifest).not.toContain("Object 1/");

    const reopenedAgain = await parseOdt(secondBytes);
    expect(reopenedAgain.sections[0].blocks.some((b) => isChart(b))).toBe(false);
  });
});

describe("writeOdt — an unsupported chart type (AC-29)", () => {
  it("stays a placeholder, byte-identical, and is never rewritten even if the block claims edited", async () => {
    const { model, chartPartName } = await openedModelWithChart(baseChart(), (xml) =>
      xml.replace('chart:class="chart:bar"', 'chart:class="chart:stock"'),
    );
    const chart = model.sections[0].blocks[1] as Chart;
    const unsupportedXml = new TextDecoder().decode(base64ToBytes(model.opaque!.find((p) => p.name === chartPartName)!.data));
    chart.unsupportedChartXml = unsupportedXml;
    chart.edited = true;

    const { bytes, dropped } = await writeOdt(model);
    expect(dropped.filter((d) => d.kind === "chartFeature")).toEqual([]);
    const parts = await unzip(bytes);
    const written = new TextDecoder().decode(parts.get(chartPartName)!);
    expect(written).toBe(unsupportedXml);
  });
});

describe("writeOdt — structural equality after an unedited fresh-then-reopen cycle", () => {
  it("equalStructural holds across two consecutive same-format saves", async () => {
    const chart = baseChart();
    const { bytes: firstBytes } = await writeOdt(modelWithChart(chart));
    const opened = await parseOdt(firstBytes);
    // writeOdt stamps dc:date with the current time when no `modified` is
    // given (odf/writer.ts:3153), so two saves either side of a second tick
    // compare unequal. Pin it to the first save's stamp, as
    // odf/roundTrip.spec.ts:141 does — the docx twin of this test had the same
    // race and was fixed in b79da32.
    const { bytes: secondBytes } = await writeOdt(rewrap(opened), {
      modified: new Date(opened.properties!.modified!),
    });
    const reopenedAgain = await parseOdt(secondBytes);
    expect(equalStructural(opened, reopenedAgain)).toBe(true);
  });
});
