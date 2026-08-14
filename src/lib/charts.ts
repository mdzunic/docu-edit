// Chart structure commands (SPEC-052), framework-free like every lib
// module. A Chart is a block-level DrawingPlacement, addressed the same way
// shapes.ts addresses a Shape/images.ts addresses an Image — mirrors
// shapes.ts's own command set (insert, move, resize, wrap, alt text,
// delete) rather than sharing it, the same "near-identical, independent
// module per block kind" precedent images.ts/shapes.ts already established.
// Z-order and grouping are NOT duplicated here: Chart already participates
// in shapes.ts's generic allDrawings-based z-order/grouping machinery
// (SPEC-052 AC-24 "for free") — see that module's own comment.
import {
  isChart,
  type Chart,
  type ChartOptions,
  type ChartSeries,
  type ChartType,
  type DocumentModel,
  type ImageWrapMode,
} from "./documentModel";
import { DEFAULT_CHART_COLORS } from "./documentModel";
import { blockLength, clampCaret, type CaretPosition } from "./textEditing";
import { caretSelection, type CommandResult, type ModelSelection } from "./commands";
import { DEFAULT_DRAWING_POSITION } from "./shapes";
import type { Block, Paragraph } from "./documentModel";

export interface ChartLocation {
  section: number;
  block: number;
}

/** Typed `null` seed for a chart-location ref in `.vue` scripts (espree bars
 * the annotation — see shapes.ts's noShapeLocation). */
export function noChartLocation(): ChartLocation | null {
  return null;
}

export type ChartCommand = (model: DocumentModel, selection: ModelSelection, location: ChartLocation) => CommandResult;

export function noChartCommand(): ChartCommand | null {
  return null;
}

function cloneContainers(model: DocumentModel): DocumentModel {
  return {
    ...model,
    sections: model.sections.map((section) => ({
      ...section,
      blocks: [...section.blocks],
    })),
  };
}

function emptyParagraph(): Paragraph {
  return { type: "paragraph", runs: [] };
}

function splitParagraphAt(paragraph: Paragraph, offset: number): [Paragraph, Paragraph] {
  let consumed = 0;
  const before: Paragraph = { ...paragraph, runs: [] };
  const after: Paragraph = {
    ...paragraph,
    runs: [],
    formatting: paragraph.formatting ? { ...paragraph.formatting } : undefined,
  };
  const beforeRuns = [];
  const afterRuns = [];
  for (const run of paragraph.runs) {
    const end = consumed + run.text.length;
    if (end <= offset) {
      beforeRuns.push(run);
    } else if (consumed >= offset) {
      afterRuns.push(run);
    } else {
      const cut = offset - consumed;
      beforeRuns.push({ text: run.text.slice(0, cut), formatting: { ...run.formatting } });
      afterRuns.push({ text: run.text.slice(cut), formatting: { ...run.formatting } });
    }
    consumed = end;
  }
  before.runs = beforeRuns;
  after.runs = afterRuns;
  return [before, after];
}

/** The chart a location addresses, or null. */
export function chartAt(model: DocumentModel, location: ChartLocation): Chart | null {
  const target = model.sections[location.section]?.blocks[location.block];
  return isChart(target) ? target : null;
}

function replaceChart(model: DocumentModel, location: ChartLocation, next: Chart): void {
  model.sections[location.section].blocks.splice(location.block, 1, next);
}

/** "Insert → Chart" size (SPEC-052 AC-3): 12cm x 7.5cm. */
export const NEW_CHART_WIDTH_CM = 12;
export const NEW_CHART_HEIGHT_CM = 7.5;

/** The sample data a freshly inserted chart carries (AC-3): three
 * categories, two series — enough to show every catalogue type's shape
 * (grouped bars/lines/slices) without the data sheet opening to an
 * intimidating wall of numbers. */
export function sampleChartCategories(): string[] {
  return ["Category 1", "Category 2", "Category 3"];
}
export function sampleChartSeries(): ChartSeries[] {
  return [
    { name: "Series 1", values: [4.3, 2.5, 3.5], color: DEFAULT_CHART_COLORS[0] },
    { name: "Series 2", values: [2.4, 4.4, 1.8], color: DEFAULT_CHART_COLORS[1] },
  ];
}

export function defaultChartOptions(): ChartOptions {
  return { legend: "right", dataLabels: false, gridlines: true };
}

/** Inserts a chart at the caret (AC-2, AC-3), splitting the caret's
 * paragraph around the insertion point exactly like shapes.ts's
 * insertShape/images.ts's insertImage. A freshly inserted chart floats
 * (`wrap: "square"`) at the shared anchored default, sized 12cm x 7.5cm,
 * with sample data — created here, so it carries no `sourcePart` and is
 * always written fresh (never has anything to preserve byte-identically). */
export function insertChart(model: DocumentModel, selection: ModelSelection, chartType: ChartType): CommandResult {
  const next = cloneContainers(model);
  const at = clampCaret(next, selection.focus);
  if (at.cell) {
    return { model: next, selection: caretSelection(at) };
  }
  const blocks = next.sections[at.section].blocks;
  const target = blocks[at.block];
  const chart: Chart = {
    type: "chart",
    chartType,
    categories: sampleChartCategories(),
    series: sampleChartSeries(),
    options: defaultChartOptions(),
    widthCm: NEW_CHART_WIDTH_CM,
    heightCm: NEW_CHART_HEIGHT_CM,
    wrap: "square",
    relativeToH: DEFAULT_DRAWING_POSITION.relativeToH,
    relativeToV: DEFAULT_DRAWING_POSITION.relativeToV,
    offsetXCm: DEFAULT_DRAWING_POSITION.offsetXCm,
    offsetYCm: DEFAULT_DRAWING_POSITION.offsetYCm,
  };

  const replacement: Block[] = [];
  let chartBlockOffset: number;
  if (target.type !== "paragraph") {
    if (at.offset === 0) {
      replacement.push(chart, target);
      chartBlockOffset = 0;
    } else {
      replacement.push(target, chart, emptyParagraph());
      chartBlockOffset = 1;
    }
  } else if (at.offset === 0) {
    replacement.push(chart, target);
    chartBlockOffset = 0;
  } else if (at.offset >= blockLength(target)) {
    replacement.push(target, chart, emptyParagraph());
    chartBlockOffset = 1;
  } else {
    const [before, after] = splitParagraphAt(target, at.offset);
    replacement.push(before, chart, after);
    chartBlockOffset = 1;
  }
  blocks.splice(at.block, 1, ...replacement);
  const chartIndex = at.block + chartBlockOffset;
  const landing: CaretPosition = { section: at.section, block: chartIndex + 1, offset: 0 };
  return { model: next, selection: caretSelection(landing) };
}

/** "Chart from table" (AC-21, AC-22): builds a chart from a table's cells.
 * `unparsable` collects the text of every cell (beyond the header row/
 * category column) that could not be read as a number — reported by the
 * caller rather than silently dropped or coerced to 0. */
export interface ChartFromTableResult {
  categories: string[];
  series: ChartSeries[];
  unparsable: string[];
}

/** Reads a table's cells (SPEC-022's `Table.rows[].cells[].blocks: Paragraph[]`)
 * into chart data: the first row (from its second cell on) gives series
 * names, the first column (from its second row on) gives categories, every
 * other cell is a value. A one-time copy — this module keeps no reference to
 * the table afterwards (SPEC-052 "Out of scope": no live link). */
export function chartDataFromTableCells(cellText: string[][]): ChartFromTableResult {
  const unparsable: string[] = [];
  if (cellText.length < 2 || (cellText[0]?.length ?? 0) < 2) {
    return { categories: [], series: [], unparsable };
  }
  const header = cellText[0];
  const seriesCount = header.length - 1;
  const series: ChartSeries[] = Array.from({ length: seriesCount }, (_, i) => ({
    name: header[i + 1]?.trim() || `Series ${i + 1}`,
    values: [],
    color: DEFAULT_CHART_COLORS[i % DEFAULT_CHART_COLORS.length],
  }));
  const categories: string[] = [];
  for (let r = 1; r < cellText.length; r++) {
    const row = cellText[r];
    categories.push(row[0]?.trim() ?? "");
    for (let s = 0; s < seriesCount; s++) {
      const raw = (row[s + 1] ?? "").trim();
      if (raw === "") {
        series[s].values.push(null);
        continue;
      }
      const n = Number(raw);
      if (Number.isFinite(n)) {
        series[s].values.push(n);
      } else {
        series[s].values.push(null);
        unparsable.push(raw);
      }
    }
  }
  return { categories, series, unparsable };
}

/** "Insert → Chart from table" (AC-21): places a new chart immediately after
 * the table, built from its cells. */
export function insertChartFromTable(
  model: DocumentModel,
  selection: ModelSelection,
  tableLocation: { section: number; block: number },
  chartType: ChartType,
  data: ChartFromTableResult,
): CommandResult {
  const next = cloneContainers(model);
  const chart: Chart = {
    type: "chart",
    chartType,
    categories: data.categories,
    series: data.series,
    options: defaultChartOptions(),
    widthCm: NEW_CHART_WIDTH_CM,
    heightCm: NEW_CHART_HEIGHT_CM,
    wrap: "square",
    relativeToH: DEFAULT_DRAWING_POSITION.relativeToH,
    relativeToV: DEFAULT_DRAWING_POSITION.relativeToV,
    offsetXCm: DEFAULT_DRAWING_POSITION.offsetXCm,
    offsetYCm: DEFAULT_DRAWING_POSITION.offsetYCm,
  };
  const insertAt = tableLocation.block + 1;
  next.sections[tableLocation.section].blocks.splice(insertAt, 0, chart);
  return {
    model: next,
    selection: caretSelection({ section: tableLocation.section, block: insertAt + 1, offset: 0 }),
  };
}

/** Commits the data sheet's edits — cell edits, row/column insert/delete,
 * pasted TSV/CSV (AC-7..AC-10) — replacing categories/series wholesale and
 * marking the chart edited (SPEC-052: "only the data sheet and the options
 * panel set it"). */
export function setChartData(
  model: DocumentModel,
  selection: ModelSelection,
  location: ChartLocation,
  categories: string[],
  series: ChartSeries[],
): CommandResult {
  const next = cloneContainers(model);
  const chart = chartAt(next, location);
  if (!chart) {
    return { model: next, selection };
  }
  replaceChart(next, location, { ...chart, categories, series, edited: true });
  return { model: next, selection };
}

/** Applies the options panel's settings (AC-14..AC-20), marking the chart
 * edited. */
export function setChartOptions(
  model: DocumentModel,
  selection: ModelSelection,
  location: ChartLocation,
  options: ChartOptions,
): CommandResult {
  const next = cloneContainers(model);
  const chart = chartAt(next, location);
  if (!chart) {
    return { model: next, selection };
  }
  replaceChart(next, location, { ...chart, options, edited: true });
  return { model: next, selection };
}

/** Sets a per-series colour (AC-19) — a thin wrapper over setChartOptions'
 * sibling for the series list, since colour lives on `Chart.series`, not
 * `Chart.options`. */
export function setChartSeriesColor(
  model: DocumentModel,
  selection: ModelSelection,
  location: ChartLocation,
  seriesIndex: number,
  color: string,
): CommandResult {
  const next = cloneContainers(model);
  const chart = chartAt(next, location);
  if (!chart || !chart.series[seriesIndex]) {
    return { model: next, selection };
  }
  const series = chart.series.map((s, i) => (i === seriesIndex ? { ...s, color } : s));
  replaceChart(next, location, { ...chart, series, edited: true });
  return { model: next, selection };
}

/** Sets a chart's display size in cm (a resize-handle drag on release) —
 * mirrors shapes.ts's setShapeSize/MIN_SHAPE_SIZE_CM. Never sets `edited`:
 * resizing does not touch the chart part. */
export const MIN_CHART_SIZE_CM = 1;

export function setChartSize(
  model: DocumentModel,
  selection: ModelSelection,
  location: ChartLocation,
  widthCm: number,
  heightCm: number,
): CommandResult {
  const next = cloneContainers(model);
  const chart = chartAt(next, location);
  if (!chart) {
    return { model: next, selection };
  }
  replaceChart(next, location, {
    ...chart,
    widthCm: Math.max(MIN_CHART_SIZE_CM, widthCm),
    heightCm: Math.max(MIN_CHART_SIZE_CM, heightCm),
  });
  return { model: next, selection };
}

/** Moves a chart by a page-space delta (a drag, or a keyboard nudge) —
 * mirrors shapes.ts's moveShape. Never sets `edited`. */
export function moveChart(
  model: DocumentModel,
  selection: ModelSelection,
  location: ChartLocation,
  dxCm: number,
  dyCm: number,
): CommandResult {
  const next = cloneContainers(model);
  const chart = chartAt(next, location);
  if (!chart) {
    return { model: next, selection };
  }
  replaceChart(next, location, {
    ...chart,
    offsetXCm: (chart.offsetXCm ?? DEFAULT_DRAWING_POSITION.offsetXCm) + dxCm,
    offsetYCm: (chart.offsetYCm ?? DEFAULT_DRAWING_POSITION.offsetYCm) + dyCm,
  });
  return { model: next, selection };
}

export function setChartWrap(
  model: DocumentModel,
  selection: ModelSelection,
  location: ChartLocation,
  wrap: ImageWrapMode,
): CommandResult {
  const next = cloneContainers(model);
  const chart = chartAt(next, location);
  if (!chart) {
    return { model: next, selection };
  }
  const updated: Chart = { ...chart, wrap };
  if (wrap === "inline") {
    delete updated.offsetXCm;
    delete updated.offsetYCm;
    delete updated.relativeToH;
    delete updated.relativeToV;
  } else if (chart.wrap === "inline") {
    updated.relativeToH = DEFAULT_DRAWING_POSITION.relativeToH;
    updated.relativeToV = DEFAULT_DRAWING_POSITION.relativeToV;
    updated.offsetXCm = DEFAULT_DRAWING_POSITION.offsetXCm;
    updated.offsetYCm = DEFAULT_DRAWING_POSITION.offsetYCm;
  }
  replaceChart(next, location, updated);
  return { model: next, selection };
}

export function setChartAltText(
  model: DocumentModel,
  selection: ModelSelection,
  location: ChartLocation,
  altText: string,
): CommandResult {
  const next = cloneContainers(model);
  const chart = chartAt(next, location);
  if (!chart) {
    return { model: next, selection };
  }
  const updated: Chart = { ...chart };
  if (altText) {
    updated.altText = altText;
  } else {
    delete updated.altText;
  }
  replaceChart(next, location, updated);
  return { model: next, selection };
}

/** Deletes a chart (AC-32): removes the block, which is all a same-format
 * save needs to also drop its part/relationship/data on the next save —
 * see docx/chartWritePlan.ts's/odf/chartWritePlan.ts's own orphan
 * detection, keyed off which `sourcePart`s are still referenced by a live
 * Chart block. */
export function deleteChart(model: DocumentModel, selection: ModelSelection, location: ChartLocation): CommandResult {
  const next = cloneContainers(model);
  const chart = chartAt(next, location);
  if (!chart) {
    return { model: next, selection };
  }
  next.sections[location.section].blocks.splice(location.block, 1);
  return { model: next, selection: caretSelection({ section: location.section, block: Math.max(0, location.block - 1), offset: 0 }) };
}
