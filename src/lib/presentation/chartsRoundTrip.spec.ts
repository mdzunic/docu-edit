// SPEC-092 AC-18, AC-19, AC-24, AC-25, AC-26, AC-27: the chart write pipeline,
// end-to-end. A chart part is injected into the shared pptx fixture (reusing
// its master/layout/theme/presentation boilerplate), then read → edited/
// inserted → written and re-read, asserting byte-identity where required and
// structural reproduction where the model changed.
import { describe, it, expect } from "vitest";
import { unzip } from "../core/zip";
import { zip } from "../core/zipWriter";
import { equalBytes } from "../core/equality";
import { equalStructural } from "../equality";
import { buildPptxFixtureBytes } from "./fixtures/build";
import { readPresentationBytes } from "./open";
import { writePptx } from "./pptx/writer";
import { chartPartXml } from "../core/chartPartOoxml";
import { insertChartFrame, setSlideChartData } from "./charts";
import type { ChartFrame, PresentationModel } from "./model";
import type { ChartVisual } from "../core/charts";

const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const CHART_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const enc = (s: string) => new TextEncoder().encode(s);
const dec = (b: Uint8Array) => new TextDecoder().decode(b);

const VISUAL: ChartVisual = {
  chartType: "columnClustered",
  categories: ["Q1", "Q2", "Q3"],
  series: [{ name: "Revenue", values: [10, 20, 30], color: "#4e79a7" }],
  options: { legend: "none", dataLabels: false, gridlines: false },
};

const GRAPHIC_FRAME =
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Chart 9"/>` +
  `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
  `<p:xfrm><a:off x="914400" y="457200"/><a:ext cx="4320000" cy="2700000"/></p:xfrm>` +
  `<a:graphic><a:graphicData uri="${CHART_URI}">` +
  `<c:chart xmlns:c="${CHART_URI}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId9"/>` +
  `</a:graphicData></a:graphic></p:graphicFrame>`;

/** The shared pptx fixture with a chart added to slide 1: a graphicFrame in the
 * shape tree, a chart part, its slide relationship, and a Content-Types
 * Override. */
async function chartDeckBytes(): Promise<Uint8Array> {
  const parts = await unzip(await buildPptxFixtureBytes());
  const slide1 = dec(parts.get("ppt/slides/slide1.xml")!).replace("</p:spTree>", `${GRAPHIC_FRAME}</p:spTree>`);
  parts.set("ppt/slides/slide1.xml", enc(slide1));
  parts.set("ppt/charts/chart1.xml", enc(chartPartXml(VISUAL)));
  const relsPath = "ppt/slides/_rels/slide1.xml.rels";
  const rels = dec(parts.get(relsPath)!).replace(
    "</Relationships>",
    `<Relationship Id="rId9" Type="${CHART_REL}" Target="../charts/chart1.xml"/></Relationships>`,
  );
  parts.set(relsPath, enc(rels));
  const ct = dec(parts.get("[Content_Types].xml")!).replace(
    "</Types>",
    `<Override PartName="/ppt/charts/chart1.xml" ContentType="${CHART_CT}"/></Types>`,
  );
  parts.set("[Content_Types].xml", enc(ct));
  return zip([...parts].map(([name, data]) => ({ name, data })));
}

async function openChartDeck(): Promise<{ model: PresentationModel; source: Map<string, Uint8Array> }> {
  const bytes = await chartDeckBytes();
  return { model: await readPresentationBytes(bytes, "pptx", "Deck"), source: await unzip(bytes) };
}

function chartFrameOf(model: PresentationModel): ChartFrame {
  const frame = model.slides[0].frames.find((f) => f.kind === "chart");
  if (!frame || frame.kind !== "chart") {
    throw new Error("no chart frame");
  }
  return frame;
}

describe("pptx chart read → write round trip", () => {
  it("reads the injected chart into a ChartFrame (AC-7 end-to-end)", async () => {
    const { model } = await openChartDeck();
    const chart = chartFrameOf(model);
    expect(chart.chartType).toBe("columnClustered");
    expect(chart.series[0].values).toEqual([10, 20, 30]);
    expect(chart.sourcePart).toBe("ppt/charts/chart1.xml");
  });

  it("a no-edit download stays byte-identical for every part incl. the chart (AC-24)", async () => {
    const { model, source } = await openChartDeck();
    const written = await unzip((await writePptx(model)).bytes);
    for (const [name, data] of source) {
      expect(equalBytes(written.get(name)!, data), `part ${name} must round-trip byte-identical`).toBe(true);
    }
  });

  it("editing one chart regenerates its part and leaves every other part identical (AC-19, AC-25)", async () => {
    const { model, source } = await openChartDeck();
    const chart = chartFrameOf(model);
    setSlideChartData(model, 0, chart.id, ["Q1", "Q2", "Q3"], [{ name: "Revenue", values: [99, 20, 30], color: "#4e79a7" }]);
    const written = await unzip((await writePptx(model)).bytes);
    // The chart part changed to carry the new value.
    expect(equalBytes(written.get("ppt/charts/chart1.xml")!, source.get("ppt/charts/chart1.xml")!)).toBe(false);
    expect(dec(written.get("ppt/charts/chart1.xml")!)).toContain("<c:v>99</c:v>");
    // Every other part is byte-identical (the slide part rebuilt is expected).
    for (const [name, data] of source) {
      if (name === "ppt/charts/chart1.xml" || name === "ppt/slides/slide1.xml") {
        continue;
      }
      expect(equalBytes(written.get(name)!, data), `part ${name} must stay byte-identical`).toBe(true);
    }
  });

  it("an edited chart, saved and reopened, reproduces its ChartFrame (AC-26)", async () => {
    const { model } = await openChartDeck();
    const chart = chartFrameOf(model);
    setSlideChartData(model, 0, chart.id, ["Q1", "Q2", "Q3"], [{ name: "Revenue", values: [99, 20, 30], color: "#4e79a7" }]);
    const reopened = await readPresentationBytes((await writePptx(model)).bytes, "pptx", "Deck");
    const reChart = chartFrameOf(reopened);
    expect(reChart.series[0].values).toEqual([99, 20, 30]);
    expect(reChart.categories).toEqual(["Q1", "Q2", "Q3"]);
  });

  it("writing twice produces identical bytes (AC-27)", async () => {
    const { model } = await openChartDeck();
    setSlideChartData(model, 0, chartFrameOf(model).id, ["Q1", "Q2", "Q3"], [{ name: "Revenue", values: [5, 6, 7], color: "#4e79a7" }]);
    const first = await unzip((await writePptx(model)).bytes);
    const second = await unzip((await writePptx(model)).bytes);
    for (const [name, data] of first) {
      expect(equalBytes(second.get(name)!, data), `part ${name}`).toBe(true);
    }
  });

  it("an inserted chart mints a new part, rel and content-type Override, and reopens (AC-18, AC-26)", async () => {
    const { model, source } = await openChartDeck();
    const id = insertChartFrame(model, 0, "pie", { xCm: 14, yCm: 2, widthCm: 12, heightCm: 7.5 })!;
    const written = await unzip((await writePptx(model)).bytes);
    // A fresh chart part appears past the existing chart1.
    const newChartParts = [...written.keys()].filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n) && !source.has(n));
    expect(newChartParts.length).toBe(1);
    // Its content-type Override was spliced.
    expect(dec(written.get("[Content_Types].xml")!)).toContain(`PartName="/${newChartParts[0]}"`);
    // A chart relationship was minted on slide 1.
    expect(dec(written.get("ppt/slides/_rels/slide1.xml.rels")!)).toContain(CHART_REL);
    // Reopen: both charts present.
    const reopened = await readPresentationBytes((await writePptx(model)).bytes, "pptx", "Deck");
    const charts = reopened.slides[0].frames.filter((f) => f.kind === "chart");
    expect(charts.length).toBe(2);
    const inserted = reopened.slides[0].frames.find((f) => f.kind === "chart" && f.chartType === "pie") as ChartFrame | undefined;
    expect(inserted).toBeDefined();
    expect(id).toBeTruthy();
    void equalStructural;
  });
});
