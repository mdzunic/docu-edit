// The chart commands (SPEC-092), framework-free like every other
// presentation/ module and importing only presentation siblings and
// src/lib/core — never the text product or a Vue module (boundary.spec.ts).
// Insert a chart, edit its data, edit its options, recolour a series. Each
// data/options command sets `edited = true`, drops the frame's verbatim source
// container `xml` (the serialiser then regenerates the part from the model),
// and marks the slide edited. A geometry change (move/resize/reorder) is
// SPEC-082's job and must NOT set `edited` — see frames.ts's
// dropChartContainerXml.
import {
  defaultChartOptions,
  sampleChartCategories,
  sampleChartSeries,
  type ChartOptions,
  type ChartSeries,
  type ChartType,
} from "../core/charts";
import { lostChartFeatures } from "../core/chartFeatures";
import { base64ToBytes } from "../core/packageParts";
import { markSlideEdited } from "./text";
import type { FrameRect } from "./frames";
import type { ChartFrame, PresentationModel, Slide } from "./model";

function r3(value: number): number {
  return Math.round(value * 1000) / 1000;
}

function slideAt(model: PresentationModel, slideIndex: number): Slide | null {
  return model.slides[slideIndex] ?? null;
}

/** The chart frame `frameId` on slide `slideIndex`, or null (wrong slide,
 * missing frame, or a non-chart frame). */
function chartFrameAt(model: PresentationModel, slideIndex: number, frameId: string): { slide: Slide; frame: ChartFrame } | null {
  const slide = slideAt(model, slideIndex);
  const frame = slide?.frames.find((f) => f.id === frameId);
  if (!slide || !frame || frame.kind !== "chart") {
    return null;
  }
  return { slide, frame };
}

/** Marks a chart's own content edited: regenerate the part from the model on
 * save (drop `xml`), then mark the slide dirty. Never called by a geometry
 * command (D11). */
function markChartEdited(model: PresentationModel, slide: Slide, frame: ChartFrame): void {
  frame.edited = true;
  delete frame.xml;
  markSlideEdited(slide, model.format);
}

/** Inserts a `chartType` chart at `rect` (the view centres it at
 * `NEW_CHART_WIDTH_CM` × `NEW_CHART_HEIGHT_CM`), seeded with the sample data
 * and default options, appended on top in z-order. Returns its id (AC-13a,
 * AC-16). A chart created here carries no `sourcePart` — the writer always
 * mints a fresh part. */
export function insertChartFrame(
  model: PresentationModel,
  slideIndex: number,
  chartType: ChartType,
  rect: FrameRect,
): string | null {
  const slide = slideAt(model, slideIndex);
  if (!slide) {
    return null;
  }
  const frame: ChartFrame = {
    id: crypto.randomUUID(),
    xCm: r3(rect.xCm),
    yCm: r3(rect.yCm),
    widthCm: r3(Math.max(rect.widthCm, 0)),
    heightCm: r3(Math.max(rect.heightCm, 0)),
    kind: "chart",
    chartType,
    categories: sampleChartCategories(),
    series: sampleChartSeries(),
    options: defaultChartOptions(),
    edited: true,
  };
  slide.frames.push(frame);
  markSlideEdited(slide, model.format);
  return frame.id;
}

/** Replaces a chart's categories and series from the data sheet (SPEC-052's
 * data-sheet Apply). No-op for an unsupported chart. Returns the frame id, or
 * null when the target is not an editable chart (AC-16). */
export function setSlideChartData(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  categories: string[],
  series: ChartSeries[],
): string | null {
  const found = chartFrameAt(model, slideIndex, frameId);
  if (!found || found.frame.unsupportedChartXml !== undefined) {
    return null;
  }
  found.frame.categories = categories.map((c) => c);
  found.frame.series = series.map((s) => ({ name: s.name, color: s.color, values: s.values.slice() }));
  markChartEdited(model, found.slide, found.frame);
  return frameId;
}

/** Replaces a chart's options from the options panel (SPEC-052's options
 * Apply). No-op for an unsupported chart. Returns the frame id (AC-16). */
export function setSlideChartOptions(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  options: ChartOptions,
): string | null {
  const found = chartFrameAt(model, slideIndex, frameId);
  if (!found || found.frame.unsupportedChartXml !== undefined) {
    return null;
  }
  found.frame.options = { ...options };
  markChartEdited(model, found.slide, found.frame);
  return frameId;
}

/** Recolours series `seriesIndex` (the options panel's live colour input). A
 * no-op for an out-of-range index or an unsupported chart. Returns the frame
 * id (AC-16). */
export function setSlideChartSeriesColor(
  model: PresentationModel,
  slideIndex: number,
  frameId: string,
  seriesIndex: number,
  color: string,
): string | null {
  const found = chartFrameAt(model, slideIndex, frameId);
  if (!found || found.frame.unsupportedChartXml !== undefined) {
    return null;
  }
  const series = found.frame.series[seriesIndex];
  if (!series) {
    return null;
  }
  series.color = color;
  markChartEdited(model, found.slide, found.frame);
  return frameId;
}

/** What the first edit of an imported chart costs, as human-readable phrases —
 * the raw material for the `chart-rewrite-notice` (AC-23). Built from
 * `lostChartFeatures` over the source part's preserved bytes, plus a workbook
 * line when a `.pptx` chart carried an embedded `c:externalData` workbook that
 * the regenerated part drops (D5). Empty for a chart created here (no
 * `sourcePart`), for an unsupported chart (never rewritten), or when the
 * source had nothing exotic to lose. */
export function lostChartFeaturesOnRewrite(model: PresentationModel, frame: ChartFrame): string[] {
  if (!frame.sourcePart || frame.unsupportedChartXml !== undefined) {
    return [];
  }
  const part = (model.opaque ?? []).find((p) => p.name === frame.sourcePart);
  if (!part) {
    return [];
  }
  const xml = new TextDecoder("utf-8").decode(base64ToBytes(part.data));
  const format = model.format === "odp" ? "odf" : "ooxml";
  const lost = lostChartFeatures(xml, format);
  if (format === "ooxml" && /<c:externalData[\s/>]/.test(xml)) {
    lost.push("its embedded data workbook (Edit Data now seeds from the chart's cached values)");
  }
  return lost;
}
