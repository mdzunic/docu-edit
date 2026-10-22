// SPEC-092 AC-21, AC-24, AC-25, AC-26, AC-27 for `.odp`: the embedded-chart
// write pipeline, end-to-end. A chart is injected into the shared odp fixture
// (a draw:frame in the first page, an Object 1/content.xml object, manifest
// entries), then read → edited/inserted → written and re-read.
import { describe, it, expect } from "vitest";
import { unzip } from "../core/zip";
import { zip } from "../core/zipWriter";
import { equalBytes } from "../core/equality";
import { buildOdpFixtureBytes } from "./fixtures/build";
import { readPresentationBytes } from "./open";
import { writeOdp } from "./odp/writer";
import { odfChartContentXml } from "../core/chartPartOdf";
import { insertChartFrame, setSlideChartData } from "./charts";
import { deleteFrame, duplicateFrame } from "./frames";
import type { ChartFrame, PresentationModel } from "./model";
import type { ChartVisual } from "../core/charts";

const enc = (s: string) => new TextEncoder().encode(s);
const dec = (b: Uint8Array) => new TextDecoder().decode(b);

const VISUAL: ChartVisual = {
  chartType: "columnClustered",
  categories: ["Q1", "Q2", "Q3"],
  series: [{ name: "Revenue", values: [10, 20, 30], color: "#4e79a7" }],
  options: { legend: "none", dataLabels: false, gridlines: false },
};

const DRAW_FRAME =
  `<draw:frame presentation:class="chart" svg:x="2cm" svg:y="1cm" svg:width="12cm" svg:height="7.5cm">` +
  `<draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
  `</draw:frame>`;

async function chartDeckBytes(): Promise<Uint8Array> {
  const parts = await unzip(await buildOdpFixtureBytes());
  const content = dec(parts.get("content.xml")!).replace(/(<draw:page\b[^>]*>)/, `$1${DRAW_FRAME}`);
  parts.set("content.xml", enc(content));
  parts.set("Object 1/content.xml", enc(odfChartContentXml(VISUAL, 12, 7.5)));
  const manifest = dec(parts.get("META-INF/manifest.xml")!).replace(
    "</manifest:manifest>",
    `<manifest:file-entry manifest:full-path="Object 1/" manifest:media-type="application/vnd.oasis.opendocument.chart"/>` +
      `<manifest:file-entry manifest:full-path="Object 1/content.xml" manifest:media-type="text/xml"/></manifest:manifest>`,
  );
  parts.set("META-INF/manifest.xml", enc(manifest));
  // mimetype must stay first + stored.
  const ordered = [...parts].toSorted(([a], [b]) => (a === "mimetype" ? -1 : b === "mimetype" ? 1 : 0));
  return zip(ordered.map(([name, data]) => ({ name, data, store: name === "mimetype" })));
}

async function openChartDeck(): Promise<{ model: PresentationModel; source: Map<string, Uint8Array> }> {
  const bytes = await chartDeckBytes();
  return { model: await readPresentationBytes(bytes, "odp", "Deck"), source: await unzip(bytes) };
}

function chartFrameOf(model: PresentationModel): ChartFrame {
  for (const slide of model.slides) {
    const frame = slide.frames.find((f) => f.kind === "chart");
    if (frame && frame.kind === "chart") {
      return frame;
    }
  }
  throw new Error("no chart frame");
}

function slideIndexOfChart(model: PresentationModel): number {
  return model.slides.findIndex((s) => s.frames.some((f) => f.kind === "chart"));
}

describe("odp chart read → write round trip", () => {
  it("reads the injected embedded chart into a ChartFrame (AC-9 end-to-end)", async () => {
    const { model } = await openChartDeck();
    const chart = chartFrameOf(model);
    expect(chart.chartType).toBe("columnClustered");
    expect(chart.series[0].values).toEqual([10, 20, 30]);
    expect(chart.sourcePart).toBe("Object 1/content.xml");
  });

  it("a no-edit download stays byte-identical for every part incl. the object (AC-24)", async () => {
    const { model, source } = await openChartDeck();
    const written = await unzip((await writeOdp(model)).bytes);
    for (const [name, data] of source) {
      expect(equalBytes(written.get(name)!, data), `part ${name} must round-trip byte-identical`).toBe(true);
    }
  });

  it("editing the chart regenerates its content.xml; other parts stay identical (AC-21, AC-25)", async () => {
    const { model, source } = await openChartDeck();
    const chart = chartFrameOf(model);
    const idx = slideIndexOfChart(model);
    setSlideChartData(model, idx, chart.id, ["Q1", "Q2", "Q3"], [{ name: "Revenue", values: [99, 20, 30], color: "#4e79a7" }]);
    const written = await unzip((await writeOdp(model)).bytes);
    expect(equalBytes(written.get("Object 1/content.xml")!, source.get("Object 1/content.xml")!)).toBe(false);
    expect(dec(written.get("Object 1/content.xml")!)).toContain("99");
    // content.xml (the page) rebuilds; everything else stays identical.
    for (const [name, data] of source) {
      if (name === "Object 1/content.xml" || name === "content.xml") {
        continue;
      }
      expect(equalBytes(written.get(name)!, data), `part ${name} must stay byte-identical`).toBe(true);
    }
  });

  it("an edited chart reopens with its new values (AC-26)", async () => {
    const { model } = await openChartDeck();
    const idx = slideIndexOfChart(model);
    setSlideChartData(model, idx, chartFrameOf(model).id, ["Q1", "Q2", "Q3"], [{ name: "Revenue", values: [99, 20, 30], color: "#4e79a7" }]);
    const reopened = await readPresentationBytes((await writeOdp(model)).bytes, "odp", "Deck");
    expect(chartFrameOf(reopened).series[0].values).toEqual([99, 20, 30]);
  });

  it("an inserted chart mints a fresh Object with a manifest entry and reopens (AC-21, AC-26)", async () => {
    const { model, source } = await openChartDeck();
    const idx = slideIndexOfChart(model);
    insertChartFrame(model, idx, "pie", { xCm: 14, yCm: 2, widthCm: 12, heightCm: 7.5 });
    const written = await unzip((await writeOdp(model)).bytes);
    const newObjects = [...written.keys()].filter((n) => /^Object \d+\/content\.xml$/.test(n) && !source.has(n));
    expect(newObjects.length).toBe(1);
    expect(dec(written.get("META-INF/manifest.xml")!)).toContain(newObjects[0]);
    const reopened = await readPresentationBytes((await writeOdp(model)).bytes, "odp", "Deck");
    const charts = reopened.slides[idx].frames.filter((f) => f.kind === "chart");
    expect(charts.length).toBe(2);
  });

  it("deleting a chart frame excludes its object, preview and manifest entries (AC-20)", async () => {
    const parts = await unzip(await chartDeckBytes());
    // Add an ObjectReplacements preview like a real Impress deck carries.
    parts.set("ObjectReplacements/Object 1", enc("preview-bytes"));
    const manifest = dec(parts.get("META-INF/manifest.xml")!).replace(
      "</manifest:manifest>",
      `<manifest:file-entry manifest:full-path="ObjectReplacements/Object 1" manifest:media-type="application/vnd.oasis.opendocument.chart"/></manifest:manifest>`,
    );
    parts.set("META-INF/manifest.xml", enc(manifest));
    const ordered = [...parts].toSorted(([a], [b]) => (a === "mimetype" ? -1 : b === "mimetype" ? 1 : 0));
    const bytes = await zip(ordered.map(([name, data]) => ({ name, data, store: name === "mimetype" })));
    const model = await readPresentationBytes(bytes, "odp", "Deck");

    const idx = slideIndexOfChart(model);
    const chart = chartFrameOf(model);
    expect(deleteFrame(model, idx, chart.id)).not.toBeNull();

    const written = await unzip((await writeOdp(model)).bytes);
    expect(written.has("Object 1/content.xml")).toBe(false);
    expect(written.has("ObjectReplacements/Object 1")).toBe(false);
    const manifestXml = dec(written.get("META-INF/manifest.xml")!);
    expect(manifestXml).not.toContain('manifest:full-path="Object 1/"');
    expect(manifestXml).not.toContain('manifest:full-path="Object 1/content.xml"');
    expect(manifestXml).not.toContain('manifest:full-path="ObjectReplacements/Object 1"');
  });

  it("duplicating a chart keeps the original object byte-identical and mints a fresh one (AC-18a)", async () => {
    const { model, source } = await openChartDeck();
    const chart = chartFrameOf(model);
    expect(duplicateFrame(model, slideIndexOfChart(model), chart.id)).not.toBeNull();
    const written = await unzip((await writeOdp(model)).bytes);
    expect(equalBytes(written.get("Object 1/content.xml")!, source.get("Object 1/content.xml")!)).toBe(true);
    const newObjects = [...written.keys()].filter((n) => /^Object \d+\/content\.xml$/.test(n) && !source.has(n));
    expect(newObjects.length).toBe(1);
  });

  it("editing one of two frames sharing an object never corrupts the other (AC-18a, AC-25)", async () => {
    const parts = await unzip(await chartDeckBytes());
    const content = dec(parts.get("content.xml")!).replace(DRAW_FRAME, `${DRAW_FRAME}${DRAW_FRAME}`);
    parts.set("content.xml", enc(content));
    const ordered = [...parts].toSorted(([a], [b]) => (a === "mimetype" ? -1 : b === "mimetype" ? 1 : 0));
    const bytes = await zip(ordered.map(([name, data]) => ({ name, data, store: name === "mimetype" })));
    const source = await unzip(bytes);
    const model = await readPresentationBytes(bytes, "odp", "Deck");
    const idx = slideIndexOfChart(model);
    const chartFrames = model.slides[idx].frames.filter((f): f is ChartFrame => f.kind === "chart");
    expect(chartFrames.length).toBe(2);
    expect(chartFrames.every((f) => f.sourcePart === "Object 1/content.xml")).toBe(true);

    setSlideChartData(model, idx, chartFrames[0].id, ["Q1", "Q2", "Q3"], [{ name: "Revenue", values: [42, 20, 30], color: "#4e79a7" }]);
    const written = await unzip((await writeOdp(model)).bytes);
    // The object the untouched frame still points at survives byte-identical.
    expect(equalBytes(written.get("Object 1/content.xml")!, source.get("Object 1/content.xml")!)).toBe(true);
    const newObjects = [...written.keys()].filter((n) => /^Object \d+\/content\.xml$/.test(n) && !source.has(n));
    expect(newObjects.length).toBe(1);
    expect(dec(written.get(newObjects[0])!)).toContain("42");
    // Reopen: one chart edited, the other keeps its original data.
    const reopened = await readPresentationBytes((await writeOdp(model)).bytes, "odp", "Deck");
    const values = reopened.slides[idx].frames
      .filter((f): f is ChartFrame => f.kind === "chart")
      .map((c) => c.series[0].values[0] as number)
      .toSorted((a, b) => a - b);
    expect(values).toEqual([10, 42]);
  });
});
