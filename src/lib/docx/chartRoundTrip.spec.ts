// SPEC-052: writer-level round trips through the real writeDocx/parseDocx
// pipeline — the byte-identical-when-untouched / regenerate-when-edited
// contract "Rewriting a chart" describes, plus deletion and the unsupported-
// type placeholder path. Mirrors docx/roundTrip.spec.ts's own conventions.
import { describe, it, expect } from "vitest";
import { parseDocx } from "./reader";
import { writeDocx } from "./writer";
import { writeOdt } from "../odf/writer";
import { parseOdt } from "../odf/reader";
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
    relativeToH: "column",
    relativeToV: "paragraph",
    offsetXCm: 0,
    offsetYCm: 0,
    ...overrides,
  };
}

function modelWithChart(chart: Chart, format?: "docx"): DocumentModel {
  const model = createDocument("chart-test-doc");
  model.sections[0].blocks = [model.sections[0].blocks[0], chart];
  if (format) {
    model.format = format;
  }
  return model;
}

/** Wraps a parsed document (docx) back into a writable model — the
 * canonical, defaults-resolved model an opened document actually holds
 * (mirrors docx/roundTrip.spec.ts's own modelFrom). */
function rewrap(parsed: Awaited<ReturnType<typeof parseDocx>>, title = "Untitled document"): DocumentModel {
  return {
    id: "chart-preserved-doc",
    title,
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "docx",
    opaque: parsed.opaque,
    sourceXml: parsed.sourceXml,
  };
}

/** Builds a genuinely realistic "opened from a real .docx" model: writes a
 * fresh document containing `chart` and reads it straight back, so
 * `model.opaque` holds a complete, real package (rels, content types, the
 * chart part itself) rather than a hand-picked subset — the same "one
 * write, one read" shortcut docx/roundTrip.spec.ts's own fixtures replace,
 * without needing a prebuilt binary file. `patchChartXml`, if given,
 * rewrites the chart part's bytes afterward (both in model.opaque and the
 * Chart's own placement-parsed fields stay whatever the original model
 * produced) — standing in for a chart authored by real Word with a property
 * this editor's model cannot carry. */
async function openedModelWithChart(
  chart: Chart,
  patchChartXml?: (xml: string) => string,
): Promise<{ model: DocumentModel; chartPartName: string }> {
  const { bytes } = await writeDocx(modelWithChart(chart));
  const parsed = await parseDocx(bytes);
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

describe("writeDocx / parseDocx — a chart created fresh", () => {
  it("saves and re-opens with the same chart data (AC-36)", async () => {
    const chart = baseChart();
    const { bytes, dropped } = await writeDocx(modelWithChart(chart));
    expect(dropped).toEqual([]);
    const reopened = await parseDocx(bytes);
    const reopenedChart = asChart(reopened.sections[0].blocks[1]);
    expect(reopenedChart.chartType).toBe("columnClustered");
    expect(reopenedChart.categories).toEqual(chart.categories);
    expect(reopenedChart.series.map((s) => s.name)).toEqual(chart.series.map((s) => s.name));
    expect(reopenedChart.series.map((s) => s.values)).toEqual(chart.series.map((s) => s.values));
    expect(reopenedChart.sourcePart).toBe("word/charts/chart1.xml");

    const parts = await unzip(bytes);
    expect(parts.has("word/charts/chart1.xml")).toBe(true);
    const docXml = new TextDecoder().decode(parts.get("word/document.xml")!);
    expect(docXml).toContain("<w:drawing");
    const relsXml = new TextDecoder().decode(parts.get("word/_rels/document.xml.rels")!);
    expect(relsXml).toContain("relationships/chart");
  });
});

describe("writeDocx — an untouched preserved chart (AC-25, AC-26)", () => {
  it("is written back byte-identical on a same-format save, and the whole rels file stays byte-identical too", async () => {
    const { model, chartPartName } = await openedModelWithChart(baseChart());
    const { bytes: firstBytes } = await writeDocx(model);
    const firstParts = await unzip(firstBytes);

    // Re-open, edit an unrelated paragraph, save again.
    const reopenedModel = rewrap(await parseDocx(firstBytes));
    const firstParagraph = reopenedModel.sections[0].blocks[0];
    if (firstParagraph.type === "paragraph") {
      firstParagraph.runs = [{ text: "Edited.", formatting: {} }];
    }
    const { bytes: secondBytes, dropped } = await writeDocx(reopenedModel);
    expect(dropped).toEqual([]);
    const secondParts = await unzip(secondBytes);

    expect(equalBytes(secondParts.get(chartPartName)!, firstParts.get(chartPartName)!)).toBe(true);
    expect(
      equalBytes(secondParts.get("word/_rels/document.xml.rels")!, firstParts.get("word/_rels/document.xml.rels")!),
    ).toBe(true);

    const reopenedAgain = await parseDocx(secondBytes);
    const chartBlock = asChart(reopenedAgain.sections[0].blocks[1]);
    expect(chartBlock.edited).toBeFalsy();
    expect(chartBlock.chartType).toBe("columnClustered");
  });
});

describe("writeDocx — an edited chart (AC-27, AC-28)", () => {
  it("regenerates the part from the model and reports what the rewrite lost", async () => {
    // Patches the opened chart part with a feature this editor's model has
    // no field for (a trendline) — standing in for a chart authored by real
    // Word with styling/analysis this editor never round-trips.
    const { model, chartPartName } = await openedModelWithChart(baseChart(), (xml) =>
      xml.replace("</c:ser>", '<c:trendline><c:trendlineType val="linear"/></c:trendline></c:ser>'),
    );
    const sourceXml = new TextDecoder().decode(base64ToBytes(model.opaque!.find((p) => p.name === chartPartName)!.data));
    const chart = model.sections[0].blocks[1] as Chart;
    chart.edited = true;
    chart.categories = ["Q1", "Q2", "Q3", "Q4"];
    chart.series = [{ name: "Revenue", values: [10, 20, 30, 40], color: "#4e79a7" }];

    const { bytes, dropped } = await writeDocx(model);
    const parts = await unzip(bytes);
    const writtenXml = new TextDecoder().decode(parts.get(chartPartName)!);
    expect(writtenXml).not.toEqual(sourceXml);
    expect(writtenXml).toContain("Q4");
    expect(writtenXml).not.toContain("c:trendline");

    const chartFeatureDrop = dropped.find((d) => d.kind === "chartFeature");
    expect(chartFeatureDrop, "a chartFeature drop for the lost trendline").toBeDefined();
    expect(chartFeatureDrop!.reason.toLowerCase()).toContain("trendline");

    const reopened = await parseDocx(bytes);
    const reopenedChart = asChart(reopened.sections[0].blocks[1]);
    expect(reopenedChart.categories).toEqual(["Q1", "Q2", "Q3", "Q4"]);
  });

  it("reports nothing extra for an edited chart whose source had nothing exotic to lose", async () => {
    const { model } = await openedModelWithChart(baseChart());
    const chart = model.sections[0].blocks[1] as Chart;
    chart.edited = true;
    chart.options = { ...chart.options, title: "New title" };

    const { dropped } = await writeDocx(model);
    expect(dropped.filter((d) => d.kind === "chartFeature")).toEqual([]);
  });
});

describe("writeDocx — duplicating an opened chart", () => {
  it("editing the duplicate never touches the original's part (regression: duplicates must not share sourcePart)", async () => {
    const { model, chartPartName } = await openedModelWithChart(baseChart());
    const originalXml = new TextDecoder().decode(base64ToBytes(model.opaque!.find((p) => p.name === chartPartName)!.data));
    const original = model.sections[0].blocks[1] as Chart;
    // Mirrors useCharts.ts's handleChartDuplicate: a catalogue-type
    // duplicate carries no `sourcePart` (and no `edited`), so the write
    // plan always regenerates it into its own fresh part rather than
    // reusing — and overwriting — the original's.
    const duplicate: Chart = { ...original, offsetXCm: (original.offsetXCm ?? 0) + 0.5, offsetYCm: (original.offsetYCm ?? 0) + 0.5 };
    delete duplicate.sourcePart;
    delete duplicate.edited;
    duplicate.categories = ["Q1", "Q2", "Q3", "Q4"];
    duplicate.series = [{ name: "Revenue", values: [10, 20, 30, 40], color: "#4e79a7" }];
    model.sections[0].blocks.push(duplicate);

    const { bytes, dropped } = await writeDocx(model);
    expect(dropped).toEqual([]);
    const parts = await unzip(bytes);

    const writtenOriginal = new TextDecoder().decode(parts.get(chartPartName)!);
    expect(writtenOriginal).toBe(originalXml);

    const newChartParts = [...parts.keys()].filter((name) => name !== chartPartName && /^word\/charts\/chart\d+\.xml$/.test(name));
    expect(newChartParts).toHaveLength(1);
    expect(new TextDecoder().decode(parts.get(newChartParts[0])!)).toContain("Q4");

    const reopened = await parseDocx(bytes);
    const reopenedOriginal = asChart(reopened.sections[0].blocks[1]);
    const reopenedDuplicate = asChart(reopened.sections[0].blocks[2]);
    expect(reopenedOriginal.categories).toEqual(["Q1", "Q2", "Q3"]);
    expect(reopenedDuplicate.categories).toEqual(["Q1", "Q2", "Q3", "Q4"]);
  });
});

describe("writeDocx — deleting a chart (AC-32, AC-33)", () => {
  it("removes the part and relationship on save, and SPEC-008 preservation does not resurrect it", async () => {
    const { model, chartPartName } = await openedModelWithChart(baseChart());
    const { bytes: firstBytes } = await writeDocx(model);

    const reopenedModel = rewrap(await parseDocx(firstBytes));
    // Delete the chart block.
    reopenedModel.sections[0].blocks = reopenedModel.sections[0].blocks.filter((b) => !isChart(b));

    const { bytes: secondBytes } = await writeDocx(reopenedModel);
    const parts = await unzip(secondBytes);
    expect(parts.has(chartPartName)).toBe(false);
    const relsXml = new TextDecoder().decode(parts.get("word/_rels/document.xml.rels")!);
    expect(relsXml).not.toContain(chartPartName.replace("word/", ""));

    const reopenedAgain = await parseDocx(secondBytes);
    expect(reopenedAgain.sections[0].blocks.some((b) => isChart(b))).toBe(false);
  });
});

describe("writeDocx — an unsupported chart type (AC-29)", () => {
  it("stays a placeholder, byte-identical, and is never rewritten even if the block claims edited", async () => {
    // Corrupts the opened chart into a combo chart (two chart-type elements)
    // — genuinely outside the catalogue, so the bytes on disk really are
    // unsupported, not just flagged as such on the model.
    const { model, chartPartName } = await openedModelWithChart(baseChart(), (xml) =>
      xml.replace(
        "</c:barChart>",
        '</c:barChart><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:marker val="0"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>',
      ),
    );
    const chart = model.sections[0].blocks[1] as Chart;
    const unsupportedXml = new TextDecoder().decode(base64ToBytes(model.opaque!.find((p) => p.name === chartPartName)!.data));
    chart.unsupportedChartXml = unsupportedXml;
    // Defensive: even if something set `edited`, an unsupported chart's
    // part must never be regenerated from its placeholder model fields.
    chart.edited = true;

    const { bytes, dropped } = await writeDocx(model);
    expect(dropped.filter((d) => d.kind === "chartFeature")).toEqual([]);
    const parts = await unzip(bytes);
    const written = new TextDecoder().decode(parts.get(chartPartName)!);
    expect(written).toBe(unsupportedXml);

    const reopened = await parseDocx(bytes);
    const reopenedChart = asChart(reopened.sections[0].blocks[1]);
    expect(reopenedChart.unsupportedChartXml).toBe(unsupportedXml);
  });
});

describe("cross-format — docx chart survives a save to .odt and back (AC-37)", () => {
  it("round-trips chartType/categories/series/options structurally", async () => {
    const chart = baseChart({ options: { title: "Revenue", legend: "top", dataLabels: true, gridlines: true } });
    const { model: openedModel } = await openedModelWithChart(chart);
    const { bytes: odtBytes, dropped } = await writeOdt(openedModel);
    // The catalogue chart converts losslessly — nothing reported for it.
    expect(dropped.some((d) => d.kind === "chart")).toBe(false);
    const openedOdt = await parseOdt(odtBytes);
    const odtChart = asChart(openedOdt.sections[0].blocks[1]);
    expect(odtChart.chartType).toBe("columnClustered");
    expect(odtChart.categories).toEqual(chart.categories);
    expect(odtChart.series.map((s) => s.name)).toEqual(chart.series.map((s) => s.name));
    expect(odtChart.options.title).toBe("Revenue");
    expect(odtChart.options.legend).toBe("top");
  });

  it("drops an unsupported chart on a cross-format save (AC 'chart' DropRecord)", async () => {
    const { model } = await openedModelWithChart(baseChart());
    const chart = model.sections[0].blocks[1] as Chart;
    chart.unsupportedChartXml = "<c:chartSpace/>";
    const { bytes, dropped } = await writeOdt(model);
    expect(dropped.some((d) => d.kind === "chart")).toBe(true);
    const opened = await parseOdt(bytes);
    expect(opened.sections[0].blocks.some((b) => isChart(b))).toBe(false);
  });
});

describe("writeDocx — structural equality after an unedited fresh-then-reopen cycle", () => {
  it("equalStructural holds across two consecutive same-format saves", async () => {
    const chart = baseChart();
    const { bytes: firstBytes } = await writeDocx(modelWithChart(chart));
    const opened = await parseDocx(firstBytes);
    const openedModel: DocumentModel = {
      id: "chart-test-doc",
      title: "Untitled document",
      sections: opened.sections,
      defaultStyles: opened.defaultStyles,
      format: "docx",
      opaque: opened.opaque,
      sourceXml: opened.sourceXml,
    };
    const { bytes: secondBytes } = await writeDocx(openedModel);
    const reopenedAgain = await parseDocx(secondBytes);
    expect(equalStructural(opened, reopenedAgain)).toBe(true);
  });
});
