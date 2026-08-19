// Reads/writes the ODF chart markup SPEC-052 models: a `draw:frame` holding
// a `draw:object` reference to an embedded document subpackage (`Object N/`)
// whose own `content.xml` carries `office:document-content` ->
// `office:body` -> `office:chart` -> `chart:chart`, with a `table:table`
// holding the category/series data. Mirrors drawing.ts/shape.ts's own split
// (parse.../...Xml) so the three modules read side by side.
//
// A handful of real ODF chart properties are normally carried on a
// `style:chart-properties` element inside a referenced automatic style
// (bar direction, stacking, symbol type, axis min/max) rather than directly
// on the element they describe. This module writes them as direct
// attributes on the owning element instead (still in the `chart:`
// namespace, still self-consistent for our own reader) — enough for this
// editor's own round trip and read leniently by LibreOffice in practice, but
// not the schema-strict placement a hand-authored ODF file would use. This
// mirrors shape.ts's own `de:rotation`/`de:group` simplification for
// concepts ODF has no direct per-element slot for.
import type { Chart } from "../documentModel";
import type { ChartOptions, ChartSeries, ChartType } from "../core/charts";
import { DEFAULT_CHART_COLORS, SINGLE_SERIES_CHART_TYPES } from "../core/charts";
import { escapeXmlAttribute, escapeXmlText, parseXml, serializeElement, XmlParseError } from "../core/xml";
import { cmToOdfLength, odfLengthToCm } from "./units";
import { DRAW_NAMESPACE, OFFICE_NAMESPACE, SVG_NAMESPACE, TABLE_NAMESPACE, TEXT_NAMESPACE, XLINK_NAMESPACE, svgAttr, textAttr, xlinkAttr } from "./xml";

export const CHART_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";

export const CHART_OBJECT_CONTENT_TYPE = "application/vnd.oasis.opendocument.chart";

function firstChildNS(parent: Element, ns: string, localName: string): Element | undefined {
  for (const child of Array.from(parent.children)) {
    if (child.namespaceURI === ns && child.localName === localName) {
      return child;
    }
  }
  return undefined;
}
function childrenNS(parent: Element, ns: string, localName: string): Element[] {
  return Array.from(parent.children).filter((child) => child.namespaceURI === ns && child.localName === localName);
}
function chartAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(CHART_NAMESPACE, localName);
}

const LEGEND_POS_TO_ODF: Record<Exclude<ChartOptions["legend"], "none">, string> = {
  right: "end",
  left: "start",
  top: "top",
  bottom: "bottom",
};
const ODF_TO_LEGEND_POS: Record<string, ChartOptions["legend"]> = {
  end: "right",
  start: "left",
  top: "top",
  bottom: "bottom",
};

// --- reading -----------------------------------------------------------------

function cellText(cellEl: Element | undefined): string {
  if (!cellEl) {
    return "";
  }
  const p = firstChildNS(cellEl, TEXT_NAMESPACE, "p");
  if (p) {
    return p.textContent ?? "";
  }
  return cellEl.textContent ?? "";
}

function cellValue(cellEl: Element | undefined): number | null {
  if (!cellEl) {
    return null;
  }
  const valueType = cellEl.getAttributeNS(OFFICE_NAMESPACE, "value-type");
  const value = cellEl.getAttributeNS(OFFICE_NAMESPACE, "value");
  if (valueType !== "float" || value === null || value === "") {
    return null;
  }
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

interface ParsedTableData {
  categories: string[];
  seriesNames: string[];
  values: (number | null)[][];
}

/** Reads the `table:table` under `chart:chart` — the shared data every
 * `chart:series` references by cell range. First header row is series
 * names (first cell is the corner, skipped); every following row is one
 * category plus its per-series values. */
function parseTableData(chartEl: Element): ParsedTableData | null {
  const table = firstChildNS(chartEl, TABLE_NAMESPACE, "table");
  if (!table) {
    return null;
  }
  const headerRows = firstChildNS(table, TABLE_NAMESPACE, "table-header-rows");
  const headerRow = headerRows ? firstChildNS(headerRows, TABLE_NAMESPACE, "table-row") : undefined;
  const headerCells = headerRow ? childrenNS(headerRow, TABLE_NAMESPACE, "table-cell") : [];
  const seriesNames = headerCells.slice(1).map((cell) => cellText(cell));

  const bodyRows = firstChildNS(table, TABLE_NAMESPACE, "table-rows");
  const rowEls = bodyRows ? childrenNS(bodyRows, TABLE_NAMESPACE, "table-row") : [];
  const categories: string[] = [];
  const values: (number | null)[][] = seriesNames.map(() => []);
  for (const row of rowEls) {
    const cells = childrenNS(row, TABLE_NAMESPACE, "table-cell");
    categories.push(cellText(cells[0]));
    seriesNames.forEach((_, i) => {
      values[i].push(cellValue(cells[i + 1]));
    });
  }
  return { categories, seriesNames, values };
}

function textOf(container: Element | undefined): string | undefined {
  if (!container) {
    return undefined;
  }
  const p = firstChildNS(container, TEXT_NAMESPACE, "p");
  const text = p?.textContent ?? "";
  return text.length > 0 ? text : undefined;
}

export interface ParsedOdfChartPart {
  chartType: ChartType;
  categories: string[];
  series: ChartSeries[];
  options: ChartOptions;
}

/** Parses an embedded chart object's `content.xml` text, or null when it is
 * outside "Chart catalogue" — the caller then preserves the whole embedded
 * object verbatim via `unsupportedChartXml` (SPEC-008). */
export function parseOdfChartXml(xmlText: string): ParsedOdfChartPart | null {
  let doc: Document;
  try {
    doc = parseXml(xmlText);
  } catch (err) {
    if (err instanceof XmlParseError) {
      return null;
    }
    throw err;
  }
  const root = doc.documentElement;
  const body = firstChildNS(root, OFFICE_NAMESPACE, "body");
  const officeChart = body && firstChildNS(body, OFFICE_NAMESPACE, "chart");
  const chartEl = officeChart && firstChildNS(officeChart, CHART_NAMESPACE, "chart");
  if (!chartEl) {
    return null;
  }
  const chartClass = chartAttr(chartEl, "class");
  const vertical = chartAttr(chartEl, "vertical") !== "false";
  const stacked = chartAttr(chartEl, "stacked") === "true";
  const symbolType = chartAttr(chartEl, "symbol-type");

  let chartType: ChartType | null = null;
  if (chartClass === "chart:bar") {
    chartType = !vertical ? "barClustered" : stacked ? "columnStacked" : "columnClustered";
  } else if (chartClass === "chart:line") {
    chartType = symbolType === "named-symbol" ? "lineMarkers" : "line";
  } else if (chartClass === "chart:area") {
    chartType = "area";
  } else if (chartClass === "chart:circle") {
    chartType = "pie";
  } else if (chartClass === "chart:ring") {
    chartType = "doughnut";
  } else if (chartClass === "chart:scatter") {
    chartType = "scatter";
  }
  if (!chartType) {
    return null;
  }

  const table = parseTableData(chartEl);
  if (!table || table.seriesNames.length === 0) {
    return null;
  }
  const series: ChartSeries[] = table.seriesNames.map((name, index) => ({
    name: name || `Series ${index + 1}`,
    color: seriesColorOf(chartEl, index) ?? DEFAULT_CHART_COLORS[index % DEFAULT_CHART_COLORS.length],
    values: table.values[index],
  }));

  const plotArea = firstChildNS(chartEl, CHART_NAMESPACE, "plot-area");
  const axes = plotArea ? childrenNS(plotArea, CHART_NAMESPACE, "axis") : [];
  const catAx = axes.find((ax) => chartAttr(ax, "dimension") === "x");
  const valAx = axes.find((ax) => chartAttr(ax, "dimension") === "y");
  const gridlines = !!valAx && !!firstChildNS(valAx, CHART_NAMESPACE, "grid");
  const legendEl = firstChildNS(chartEl, CHART_NAMESPACE, "legend");
  const legendPos = legendEl ? chartAttr(legendEl, "legend-position") : null;
  const dataLabels = plotArea ? chartAttr(plotArea, "data-labels") === "true" : false;

  const options: ChartOptions = {
    legend: legendPos ? ODF_TO_LEGEND_POS[legendPos] ?? "right" : "none",
    dataLabels,
    gridlines,
  };
  const title = textOf(firstChildNS(chartEl, CHART_NAMESPACE, "title"));
  if (title) {
    options.title = title;
  }
  const catTitle = catAx && textOf(firstChildNS(catAx, CHART_NAMESPACE, "title"));
  if (catTitle) {
    options.categoryAxisTitle = catTitle;
  }
  const valTitle = valAx && textOf(firstChildNS(valAx, CHART_NAMESPACE, "title"));
  if (valTitle) {
    options.valueAxisTitle = valTitle;
  }
  const min = valAx && chartAttr(valAx, "minimum");
  const max = valAx && chartAttr(valAx, "maximum");
  if (min) {
    options.valueAxisMin = Number(min);
  }
  if (max) {
    options.valueAxisMax = Number(max);
  }

  return { chartType, categories: table.categories, series, options };
}

function seriesColorOf(chartEl: Element, index: number): string | undefined {
  const plotArea = firstChildNS(chartEl, CHART_NAMESPACE, "plot-area");
  const seriesEls = plotArea ? childrenNS(plotArea, CHART_NAMESPACE, "series") : [];
  const serEl = seriesEls[index];
  const color = serEl?.getAttributeNS(CHART_NAMESPACE, "color");
  return color ? color.toLowerCase() : undefined;
}

export interface ParsedOdfChartDrawing {
  chartType: ChartType;
  categories: string[];
  series: ChartSeries[];
  options: ChartOptions;
  unsupportedChartXml?: string;
  sourcePart: string;
  widthCm: number;
  heightCm: number;
  wrap: Chart["wrap"];
  altText?: string;
}

/** Parses a `draw:frame` into a ParsedOdfChartDrawing, or null when it does
 * not hold a `draw:object` reference this spec follows (a picture, a shape,
 * a href the package carries no embedded object for) — the caller's
 * existing opaque-inline fallback then preserves it verbatim (SPEC-008). */
export function parseChartFrame(frameEl: Element, parts: Map<string, Uint8Array>): ParsedOdfChartDrawing | null {
  const objectEl = firstChildNS(frameEl, DRAW_NAMESPACE, "object");
  if (!objectEl) {
    return null;
  }
  const href = xlinkAttr(objectEl, "href");
  if (!href) {
    return null;
  }
  const objectDir = href.replace(/^\.\//, "").replace(/\/$/, "");
  const partPath = `${objectDir}/content.xml`;
  const bytes = parts.get(partPath);
  if (!bytes) {
    return null;
  }
  const xmlText = new TextDecoder("utf-8").decode(bytes);
  const parsed = parseOdfChartXml(xmlText);

  const widthCm = odfLengthToCm(svgAttr(frameEl, "width"));
  const heightCm = odfLengthToCm(svgAttr(frameEl, "height"));
  if (widthCm === undefined || heightCm === undefined || widthCm <= 0 || heightCm <= 0) {
    return null;
  }
  const anchorType = textAttr(frameEl, "anchor-type");
  const wrap: Chart["wrap"] = anchorType === null || anchorType === "as-char" ? "inline" : "square";
  const descEl = firstChildNS(frameEl, SVG_NAMESPACE, "desc");
  const altText = descEl?.textContent || undefined;

  const result: ParsedOdfChartDrawing = {
    chartType: parsed?.chartType ?? "columnClustered",
    categories: parsed?.categories ?? [],
    series: parsed?.series ?? [],
    options: parsed?.options ?? { legend: "none", dataLabels: false, gridlines: false },
    sourcePart: partPath,
    widthCm,
    heightCm,
    wrap,
  };
  if (!parsed) {
    result.unsupportedChartXml = xmlText;
  }
  if (altText) {
    result.altText = altText;
  }
  return result;
}

// --- writing -------------------------------------------------------------

const FRAME_XMLNS = ` xmlns:draw="${DRAW_NAMESPACE}" xmlns:xlink="${XLINK_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}"`;

/** Serialises a Chart into a `draw:frame` referencing `objectHref` (the
 * embedded object subpackage's package-relative directory, e.g.
 * "./Object 1") — the chart-flavoured sibling of drawing.ts's drawFrameXml/
 * shape.ts's shapeElementXml. */
export function chartFrameXml(chart: Chart, objectHref: string): string {
  const anchorType = chart.wrap === "inline" ? "as-char" : "paragraph";
  const descXml = chart.altText ? `<svg:desc>${escapeXmlAttribute(chart.altText)}</svg:desc>` : "";
  return (
    `<draw:frame${FRAME_XMLNS} text:anchor-type="${anchorType}"` +
    ` svg:width="${cmToOdfLength(chart.widthCm)}" svg:height="${cmToOdfLength(Math.max(0, chart.heightCm))}">` +
    `<draw:object xlink:href="${escapeXmlAttribute(objectHref)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
    descXml +
    "</draw:frame>"
  );
}

/** A Chart member of a `draw:g` group, as a `draw:frame`/`draw:object`
 * positioned at absolute `xCm`/`yCm` in the anchor frame (SPEC-057
 * AC-62/AC-63) — the chart sibling of drawing.ts's `drawFrameGroupChildXml`.
 * `draw:g` defines no geometry, so members carry their own svg:x/svg:y. */
export function chartFrameGroupChildXml(chart: Chart, objectHref: string, xCm: number, yCm: number): string {
  const descXml = chart.altText ? `<svg:desc>${escapeXmlAttribute(chart.altText)}</svg:desc>` : "";
  return (
    `<draw:frame${FRAME_XMLNS} text:anchor-type="paragraph"` +
    ` svg:x="${cmToOdfLength(xCm)}" svg:y="${cmToOdfLength(yCm)}"` +
    ` svg:width="${cmToOdfLength(chart.widthCm)}" svg:height="${cmToOdfLength(Math.max(0, chart.heightCm))}">` +
    `<draw:object xlink:href="${escapeXmlAttribute(objectHref)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
    descXml +
    "</draw:frame>"
  );
}

function tableColumnLetter(index: number): string {
  return String.fromCharCode(66 + (index % 24));
}

function tableXml(chart: Chart, series: ChartSeries[]): string {
  const seriesCells = series
    .map((s) => `<table:table-cell office:value-type="string"><text:p>${escapeXmlText(s.name)}</text:p></table:table-cell>`)
    .join("");
  const headerRow = `<table:table-row><table:table-cell/>${seriesCells}</table:table-row>`;
  const bodyRows = chart.categories
    .map((cat, rowIndex) => {
      const catCell = `<table:table-cell office:value-type="string"><text:p>${escapeXmlText(cat)}</text:p></table:table-cell>`;
      const valueCells = series
        .map((s) => {
          const v = s.values[rowIndex];
          return v === null || v === undefined
            ? "<table:table-cell/>"
            : `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
        })
        .join("");
      return `<table:table-row>${catCell}${valueCells}</table:table-row>`;
    })
    .join("");
  const columns = series.map(() => '<table:table-column/>').join("");
  return (
    `<table:table table:name="local-table">` +
    `<table:table-header-columns><table:table-column/></table:table-header-columns>` +
    `<table:table-columns>${columns}</table:table-columns>` +
    `<table:table-header-rows>${headerRow}</table:table-header-rows>` +
    `<table:table-rows>${bodyRows}</table:table-rows>` +
    "</table:table>"
  );
}

function seriesXml(chart: Chart, series: ChartSeries, index: number): string {
  const col = tableColumnLetter(index);
  const rows = chart.categories.length;
  const classFor: Record<ChartType, string> = {
    columnClustered: "chart:bar",
    columnStacked: "chart:bar",
    barClustered: "chart:bar",
    line: "chart:line",
    lineMarkers: "chart:line",
    area: "chart:area",
    pie: "chart:circle",
    doughnut: "chart:ring",
    scatter: "chart:scatter",
  };
  return (
    `<chart:series chart:class="${classFor[chart.chartType]}" chart:color="${escapeXmlAttribute(series.color)}"` +
    ` chart:values-cell-range-address="local-table.$${col}$2:.$${col}$${rows + 1}"` +
    ` chart:label-cell-address="local-table.$${col}$1"/>`
  );
}

function axisTitleXml(text: string): string {
  return `<chart:title><text:p>${escapeXmlText(text)}</text:p></chart:title>`;
}

function axesXml(chart: Chart): string {
  if (chart.chartType === "pie" || chart.chartType === "doughnut") {
    return "";
  }
  const catTitle = chart.options.categoryAxisTitle ? axisTitleXml(chart.options.categoryAxisTitle) : "";
  const valTitle = chart.options.valueAxisTitle ? axisTitleXml(chart.options.valueAxisTitle) : "";
  const gridXml = chart.options.gridlines ? '<chart:grid chart:class="major"/>' : "";
  const minMaxAttrs =
    (chart.options.valueAxisMin !== undefined ? ` chart:minimum="${chart.options.valueAxisMin}"` : "") +
    (chart.options.valueAxisMax !== undefined ? ` chart:maximum="${chart.options.valueAxisMax}"` : "");
  return (
    `<chart:axis chart:dimension="x" chart:name="primary-x">${catTitle}</chart:axis>` +
    `<chart:axis chart:dimension="y" chart:name="primary-y"${minMaxAttrs}>${valTitle}${gridXml}</chart:axis>`
  );
}

/** Regenerates an embedded chart object's full `content.xml` from the model
 * — used when `chart.edited` is true (SPEC-052 "Rewriting a chart"). An
 * untouched chart never calls this: its embedded object's original bytes
 * are re-emitted verbatim from `DocumentModel.opaque` instead. */
export function odfChartContentXml(chart: Chart): string {
  const vertical = chart.chartType !== "barClustered";
  const stacked = chart.chartType === "columnStacked";
  const classFor: Record<ChartType, string> = {
    columnClustered: "chart:bar",
    columnStacked: "chart:bar",
    barClustered: "chart:bar",
    line: "chart:line",
    lineMarkers: "chart:line",
    area: "chart:area",
    pie: "chart:circle",
    doughnut: "chart:ring",
    scatter: "chart:scatter",
  };
  const extraAttrs =
    classFor[chart.chartType] === "chart:bar"
      ? ` chart:vertical="${vertical}" chart:stacked="${stacked}"`
      : chart.chartType === "lineMarkers"
        ? ' chart:symbol-type="named-symbol"'
        : chart.chartType === "line"
          ? ' chart:symbol-type="none"'
          : "";
  const titleXml = chart.options.title ? axisTitleXml(chart.options.title) : "";
  const legendXml =
    chart.options.legend === "none"
      ? ""
      : `<chart:legend chart:legend-position="${LEGEND_POS_TO_ODF[chart.options.legend]}"/>`;
  // Pie/doughnut render (and so persist) only the first series — a real
  // chart:circle/chart:ring holds exactly one chart:series; the model keeps
  // the rest for the data sheet to still show (SPEC-052 "Chart catalogue").
  const seriesToWrite = SINGLE_SERIES_CHART_TYPES.has(chart.chartType) ? chart.series.slice(0, 1) : chart.series;
  const seriesXmlStr = seriesToWrite.map((s, i) => seriesXml(chart, s, i)).join("");
  const dataLabelsAttr = chart.options.dataLabels ? ' chart:data-labels="true"' : "";
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    `<office:document-content xmlns:office="${OFFICE_NAMESPACE}" xmlns:chart="${CHART_NAMESPACE}"` +
    ` xmlns:table="${TABLE_NAMESPACE}" xmlns:text="${TEXT_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}" office:version="1.2">` +
    "<office:body><office:chart>" +
    `<chart:chart chart:class="${classFor[chart.chartType]}"${extraAttrs}` +
    ` svg:width="${cmToOdfLength(chart.widthCm)}" svg:height="${cmToOdfLength(Math.max(0, chart.heightCm))}">` +
    titleXml +
    legendXml +
    `<chart:plot-area${dataLabelsAttr}>${axesXml(chart)}${seriesXmlStr}</chart:plot-area>` +
    tableXml(chart, seriesToWrite) +
    "</chart:chart>" +
    "</office:chart></office:body>" +
    "</office:document-content>"
  );
}

// Re-exported purely so callers already holding an Element (rather than
// text) can round-trip an unsupported chart's XML the same way shape.ts's
// unsupportedGeometryXml capture does.
export { serializeElement };
