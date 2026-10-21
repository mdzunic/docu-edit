// The ODF *chart-object* codec: the `office:document-content` ->
// `office:body` -> `office:chart` -> `chart:chart` document that an embedded
// chart sub-package (`Object N/content.xml`) holds, identical whether the
// host is an `.odt` text document or an `.odp` deck (SPEC-052 modelled it for
// the text product; SPEC-092 hoists it here so the presentation product
// shares the identical schema). Product-neutral: it reads/writes the object's
// own `content.xml` bytes against the `ChartVisual` model and knows nothing of
// the `draw:frame`/`draw:object` container that references it — that container
// codec stays in each product's own layer (`odf/chart.ts`,
// `presentation/odp`).
//
// A handful of real ODF chart properties are normally carried on a
// `style:chart-properties` element inside a referenced automatic style (bar
// direction, stacking, symbol type, axis min/max) rather than directly on the
// element they describe. This module writes them as direct attributes on the
// owning element instead (still in the `chart:` namespace, still
// self-consistent for our own reader) — enough for this editor's own round
// trip and read leniently by LibreOffice in practice.
import type { ChartOptions, ChartSeries, ChartType, ChartVisual } from "./charts";
import { DEFAULT_CHART_COLORS, SINGLE_SERIES_CHART_TYPES } from "./charts";
import { escapeXmlAttribute, escapeXmlText, parseXml, XmlParseError } from "./xml";
import { DEFAULT_THEME, resolveColor, type DocumentTheme } from "./theme";
import { cmToOdfLength, OFFICE_NAMESPACE, SVG_NAMESPACE, TABLE_NAMESPACE, TEXT_NAMESPACE } from "./odfXml";

export const CHART_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";

export const CHART_OBJECT_CONTENT_TYPE = "application/vnd.oasis.opendocument.chart";

/** A host's live source references for a regenerated object (SPEC-132): WIRE
 * cell-range-address texts (`Sheet1.B2:Sheet1.B5`, the codec never
 * interprets them). `plotArea` → `chart:plot-area@table:cell-range-address`,
 * `categories` → the x axis's `chart:categories@table:cell-range-address`;
 * per series (index-aligned with the series written) `values` →
 * `@chart:values-cell-range-address`, `label` → `@chart:label-cell-address`,
 * `domain` → a `chart:domain@table:cell-range-address` child (scatter x). An
 * absent entry keeps the `local-table` address the object writes without
 * refs. */
export interface OdfChartRefs {
  plotArea?: string;
  categories?: string;
  series: { values?: string; label?: string; domain?: string }[];
}

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

// --- writing -------------------------------------------------------------

function tableColumnLetter(index: number): string {
  return String.fromCharCode(66 + (index % 24));
}

/** The `local-table` data cache an object written from `chart` holds (the
 * series pie/doughnut persist only) — SPEC-132's patch path replaces a
 * source object's table with it. */
export function odfChartTableXml(chart: ChartVisual): string {
  const seriesToWrite = SINGLE_SERIES_CHART_TYPES.has(chart.chartType) ? chart.series.slice(0, 1) : chart.series;
  return tableXml(chart, seriesToWrite);
}

function tableXml(chart: ChartVisual, series: ChartSeries[]): string {
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

function seriesXml(chart: ChartVisual, series: ChartSeries, index: number, theme: DocumentTheme): string {
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
    // ODF's chart:color has no theme vocabulary — a reference writes its
    // resolved literal (SPEC-077 D8/AC-39 semantics for the ODF family).
    `<chart:series chart:class="${classFor[chart.chartType]}" chart:color="${escapeXmlAttribute(resolveColor(series.color, theme))}"` +
    ` chart:values-cell-range-address="local-table.$${col}$2:.$${col}$${rows + 1}"` +
    ` chart:label-cell-address="local-table.$${col}$1"/>`
  );
}

function axisTitleXml(text: string): string {
  return `<chart:title><text:p>${escapeXmlText(text)}</text:p></chart:title>`;
}

function axesXml(chart: ChartVisual): string {
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

/** Regenerates an embedded chart object's full `content.xml` from the model —
 * used when a chart is `edited` (SPEC-052 "Rewriting a chart"). Geometry
 * (`widthCm`/`heightCm`) is passed explicitly because `ChartVisual` — unlike
 * the text product's `Chart` — carries no size; its container (the
 * `draw:frame`/`p:graphicFrame`) owns placement. An untouched chart never
 * calls this: its object's original bytes are re-emitted verbatim instead. */
export function odfChartContentXml(
  chart: ChartVisual,
  widthCm: number,
  heightCm: number,
  theme: DocumentTheme = DEFAULT_THEME,
  refs?: OdfChartRefs,
): string {
  void refs; // SPEC-132 skeleton: implemented by the core codec change
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
  const seriesXmlStr = seriesToWrite.map((s, i) => seriesXml(chart, s, i, theme)).join("");
  const dataLabelsAttr = chart.options.dataLabels ? ' chart:data-labels="true"' : "";
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    `<office:document-content xmlns:office="${OFFICE_NAMESPACE}" xmlns:chart="${CHART_NAMESPACE}"` +
    ` xmlns:table="${TABLE_NAMESPACE}" xmlns:text="${TEXT_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}" office:version="1.2">` +
    "<office:body><office:chart>" +
    `<chart:chart chart:class="${classFor[chart.chartType]}"${extraAttrs}` +
    ` svg:width="${cmToOdfLength(widthCm)}" svg:height="${cmToOdfLength(Math.max(0, heightCm))}">` +
    titleXml +
    legendXml +
    `<chart:plot-area${dataLabelsAttr}>${axesXml(chart)}${seriesXmlStr}</chart:plot-area>` +
    tableXml(chart, seriesToWrite) +
    "</chart:chart>" +
    "</office:chart></office:body>" +
    "</office:document-content>"
  );
}
