// The DrawingML *chart part* codec: the `c:chartSpace`/`c:chart`/`c:plotArea`
// document that both a `.docx` (`word/charts/chartN.xml`) and a `.pptx`
// (`ppt/charts/chartN.xml`) package hold, referenced from the host part
// through a `c:chart r:id` (SPEC-052 modelled it for the text product;
// SPEC-092 hoists it here so the presentation product shares the identical
// schema). Product-neutral: it reads/writes the part's own bytes against the
// `ChartVisual` model (`core/charts.ts`) and knows nothing of the `w:drawing`
// or `p:graphicFrame` container that carries the reference — those container
// codecs stay in each product's own layer (`docx/chart.ts`,
// `presentation/pptx`).
//
// A chart's *content* lives in one of two states: `sourcePart`'s bytes
// verbatim (an untouched chart, SPEC-008's byte-identical guarantee) or
// freshly regenerated from the model (an edited chart). This module only ever
// produces the "regenerate from model" half — `chartPartXml`; each writer
// decides which one to emit for a given chart.
import type { ChartOptions, ChartSeries, ChartType, ChartVisual } from "./charts";
import { DEFAULT_CHART_COLORS, SINGLE_SERIES_CHART_TYPES } from "./charts";
import { childElements, escapeXmlText, firstChildElement, parseXml, XmlParseError } from "./xml";
import { cssColorToOoxml } from "./colors";
import { isThemeColor, type ModelColor } from "./theme";
import { schemeClrRefFrom, schemeClrXmlFor } from "./themeDrawingMl";

const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const C_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const CHART_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
export const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

const CHART_TYPE_ELEMENT_TAGS = [
  "barChart",
  "bar3DChart",
  "lineChart",
  "line3DChart",
  "areaChart",
  "area3DChart",
  "pieChart",
  "pie3DChart",
  "ofPieChart",
  "doughnutChart",
  "scatterChart",
  "radarChart",
  "stockChart",
  "bubbleChart",
  "surfaceChart",
  "surface3DChart",
] as const;

const LEGEND_POS_TO_OOXML: Record<Exclude<ChartOptions["legend"], "none">, string> = {
  top: "t",
  bottom: "b",
  left: "l",
  right: "r",
};
const OOXML_TO_LEGEND_POS: Record<string, ChartOptions["legend"]> = {
  t: "top",
  b: "bottom",
  l: "left",
  r: "right",
  tr: "right",
};

// --- reading -----------------------------------------------------------------

interface CachedPoint {
  idx: number;
  v: string;
}

function firstOfAny(el: Element, ...names: string[]): Element | undefined {
  for (const name of names) {
    const found = firstChildElement(el, C_NAMESPACE, name);
    if (found) {
      return found;
    }
  }
  return undefined;
}

function cacheOf(refEl: Element | undefined): Element | undefined {
  return refEl ? firstOfAny(refEl, "strCache", "numCache") : undefined;
}

function ptCountOf(refEl: Element | undefined): number {
  const cache = cacheOf(refEl);
  const ptCount = cache && firstChildElement(cache, C_NAMESPACE, "ptCount");
  const val = ptCount?.getAttribute("val");
  return val ? Number(val) : 0;
}

function ptsOf(refEl: Element | undefined): CachedPoint[] {
  const cache = cacheOf(refEl);
  if (!cache) {
    return [];
  }
  return childElements(cache, C_NAMESPACE, "pt").map((pt) => ({
    idx: Number(pt.getAttribute("idx") ?? "0"),
    v: firstChildElement(pt, C_NAMESPACE, "v")?.textContent ?? "",
  }));
}

function refOf(container: Element | undefined, tag: string): Element | undefined {
  const wrapper = container && firstChildElement(container, C_NAMESPACE, tag);
  return wrapper ? firstOfAny(wrapper, "strRef", "numRef") : undefined;
}

function categoriesFrom(ref: Element | undefined): string[] {
  const count = ptCountOf(ref);
  const cats = Array.from<string>({ length: count }).fill("");
  for (const pt of ptsOf(ref)) {
    if (pt.idx >= 0 && pt.idx < count) {
      cats[pt.idx] = pt.v;
    }
  }
  return cats;
}

function valuesFrom(ref: Element | undefined): (number | null)[] {
  const count = ptCountOf(ref);
  const values = Array.from<number | null>({ length: count }).fill(null);
  for (const pt of ptsOf(ref)) {
    if (pt.idx >= 0 && pt.idx < count && pt.v !== "") {
      const n = Number(pt.v);
      values[pt.idx] = Number.isFinite(n) ? n : null;
    }
  }
  return values;
}

function textOf(container: Element | undefined): string | undefined {
  if (!container) {
    return undefined;
  }
  const runs = Array.from(container.getElementsByTagNameNS(A_NAMESPACE, "t"));
  if (runs.length === 0) {
    return undefined;
  }
  const text = runs.map((t) => t.textContent ?? "").join("");
  return text.length > 0 ? text : undefined;
}

function seriesNameOf(serEl: Element, fallback: string): string {
  const tx = firstChildElement(serEl, C_NAMESPACE, "tx");
  const strRef = tx && firstChildElement(tx, C_NAMESPACE, "strRef");
  const cached = strRef && firstChildElement(strRef, C_NAMESPACE, "strCache");
  const pt = cached && firstChildElement(cached, C_NAMESPACE, "pt");
  const v = pt && firstChildElement(pt, C_NAMESPACE, "v")?.textContent;
  if (v) {
    return v;
  }
  const literal = tx && firstChildElement(tx, C_NAMESPACE, "v")?.textContent;
  return literal || fallback;
}

function seriesColorOf(serEl: Element, fallback: string): ModelColor {
  const spPr = firstChildElement(serEl, C_NAMESPACE, "spPr");
  const solidFill = spPr && firstChildElement(spPr, A_NAMESPACE, "solidFill");
  const srgbClr = solidFill && firstChildElement(solidFill, A_NAMESPACE, "srgbClr");
  const val = srgbClr?.getAttribute("val");
  if (val) {
    return `#${val.toLowerCase()}`;
  }
  // An `a:schemeClr` the model can carry reads back as a ThemeColorRef
  // (SPEC-077 AC-29) — the form serFillXml writes for a themed series.
  const schemeClr = solidFill && firstChildElement(solidFill, A_NAMESPACE, "schemeClr");
  const ref = schemeClr ? schemeClrRefFrom(schemeClr) : undefined;
  return ref ?? fallback;
}

function legendPositionOf(chartEl: Element): ChartOptions["legend"] {
  const legendEl = firstChildElement(chartEl, C_NAMESPACE, "legend");
  if (!legendEl) {
    return "none";
  }
  const pos = firstChildElement(legendEl, C_NAMESPACE, "legendPos")?.getAttribute("val") ?? "r";
  return OOXML_TO_LEGEND_POS[pos] ?? "right";
}

function dataLabelsOn(typeEl: Element, serEls: Element[]): boolean {
  const dLbls = firstChildElement(typeEl, C_NAMESPACE, "dLbls") ?? (serEls[0] && firstChildElement(serEls[0], C_NAMESPACE, "dLbls"));
  return !!dLbls && firstChildElement(dLbls, C_NAMESPACE, "showVal")?.getAttribute("val") === "1";
}

function axisMinMax(axEl: Element | undefined): { min?: number; max?: number } {
  const scaling = axEl && firstChildElement(axEl, C_NAMESPACE, "scaling");
  const minEl = scaling && firstChildElement(scaling, C_NAMESPACE, "min");
  const maxEl = scaling && firstChildElement(scaling, C_NAMESPACE, "max");
  const min = minEl?.getAttribute("val");
  const max = maxEl?.getAttribute("val");
  return { min: min ? Number(min) : undefined, max: max ? Number(max) : undefined };
}

export interface ParsedChartPart {
  chartType: ChartType;
  categories: string[];
  series: ChartSeries[];
  options: ChartOptions;
}

/** Parses a `c:chartSpace` part's text into the model, or null when the
 * chart is outside "Chart catalogue" (a combo chart, a 3-D variant, radar/
 * stock/bubble/surface, or malformed XML) — the caller then preserves the
 * whole part verbatim via `unsupportedChartXml` (SPEC-008). */
export function parseChartXml(xmlText: string): ParsedChartPart | null {
  let doc: Document;
  try {
    doc = parseXml(xmlText);
  } catch (err) {
    if (err instanceof XmlParseError) {
      return null;
    }
    throw err;
  }
  const chartSpace = doc.documentElement;
  const chartEl = firstChildElement(chartSpace, C_NAMESPACE, "chart");
  const plotArea = chartEl && firstChildElement(chartEl, C_NAMESPACE, "plotArea");
  if (!chartEl || !plotArea) {
    return null;
  }

  const present = CHART_TYPE_ELEMENT_TAGS.map((tag) => ({ tag, el: firstChildElement(plotArea, C_NAMESPACE, tag) })).filter(
    (x): x is { tag: (typeof CHART_TYPE_ELEMENT_TAGS)[number]; el: Element } => !!x.el,
  );
  if (present.length !== 1) {
    // Either no recognised chart-type element, or more than one (a combo
    // chart) — both outside the catalogue.
    return null;
  }
  const { tag, el: typeEl } = present[0];

  let chartType: ChartType | null = null;
  if (tag === "barChart") {
    const barDir = firstChildElement(typeEl, C_NAMESPACE, "barDir")?.getAttribute("val");
    const grouping = firstChildElement(typeEl, C_NAMESPACE, "grouping")?.getAttribute("val") ?? "clustered";
    if (barDir === "col" && grouping === "clustered") {
      chartType = "columnClustered";
    } else if (barDir === "col" && grouping === "stacked") {
      chartType = "columnStacked";
    } else if (barDir === "bar" && grouping === "clustered") {
      chartType = "barClustered";
    }
  } else if (tag === "lineChart") {
    const marker = firstChildElement(typeEl, C_NAMESPACE, "marker")?.getAttribute("val");
    chartType = marker === "1" ? "lineMarkers" : "line";
  } else if (tag === "areaChart") {
    chartType = "area";
  } else if (tag === "pieChart") {
    chartType = "pie";
  } else if (tag === "doughnutChart") {
    chartType = "doughnut";
  } else if (tag === "scatterChart") {
    chartType = "scatter";
  }
  if (!chartType) {
    return null;
  }

  const serEls = childElements(typeEl, C_NAMESPACE, "ser");
  if (serEls.length === 0) {
    return null;
  }
  const isScatter = chartType === "scatter";
  let categories: string[] = [];
  const series: ChartSeries[] = serEls.map((serEl, index) => {
    const catRef = isScatter ? refOf(serEl, "xVal") : refOf(serEl, "cat");
    const valRef = isScatter ? refOf(serEl, "yVal") : refOf(serEl, "val");
    const cats = categoriesFrom(catRef);
    if (index === 0) {
      categories = cats;
    }
    return {
      name: seriesNameOf(serEl, `Series ${index + 1}`),
      color: seriesColorOf(serEl, DEFAULT_CHART_COLORS[index % DEFAULT_CHART_COLORS.length]),
      values: valuesFrom(valRef),
    };
  });

  const isScatterOrPie = chartType === "pie" || chartType === "doughnut";
  const axes = isScatterOrPie ? undefined : childElements(plotArea, C_NAMESPACE, isScatter ? "valAx" : "catAx");
  const catAx = axes?.[0];
  const valAxCandidates = isScatterOrPie ? [] : childElements(plotArea, C_NAMESPACE, "valAx");
  const valAx = isScatter ? valAxCandidates[1] ?? valAxCandidates[0] : valAxCandidates[0];
  const { min, max } = axisMinMax(valAx);

  const options: ChartOptions = {
    legend: legendPositionOf(chartEl),
    dataLabels: dataLabelsOn(typeEl, serEls),
    gridlines: !!valAx && !!firstChildElement(valAx, C_NAMESPACE, "majorGridlines"),
  };
  const title = textOf(firstChildElement(chartEl, C_NAMESPACE, "title"));
  if (title) {
    options.title = title;
  }
  const catTitle = catAx && textOf(firstChildElement(catAx, C_NAMESPACE, "title"));
  if (catTitle) {
    options.categoryAxisTitle = catTitle;
  }
  const valTitle = valAx && textOf(firstChildElement(valAx, C_NAMESPACE, "title"));
  if (valTitle) {
    options.valueAxisTitle = valTitle;
  }
  if (min !== undefined) {
    options.valueAxisMin = min;
  }
  if (max !== undefined) {
    options.valueAxisMax = max;
  }

  return { chartType, categories, series, options };
}

// --- writing -------------------------------------------------------------

function columnLetterFor(seriesIndex: number): string {
  // 'B'.. — column A is reserved for categories, as every emitted c:f cached
  // formula string below assumes. These formulas are never resolved by a
  // live workbook (no xlsx is embedded, only the cache) so a wrap past 'Z'
  // is harmless — the cache, not the formula text, is what any reader (ours
  // included) actually reads.
  return String.fromCharCode(66 + (seriesIndex % 24));
}

function catXml(categories: string[]): string {
  const pts = categories
    .map((c, i) => (c === "" ? "" : `<c:pt idx="${i}"><c:v>${escapeXmlText(c)}</c:v></c:pt>`))
    .join("");
  return (
    `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${categories.length + 1}</c:f>` +
    `<c:strCache><c:ptCount val="${categories.length}"/>${pts}</c:strCache></c:strRef></c:cat>`
  );
}

function xValXml(categories: string[]): string {
  const pts = categories
    .map((c, i) => {
      const n = Number(c);
      return `<c:pt idx="${i}"><c:v>${Number.isFinite(n) ? n : i + 1}</c:v></c:pt>`;
    })
    .join("");
  return (
    `<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$${categories.length + 1}</c:f>` +
    `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${categories.length}"/>${pts}</c:numCache></c:numRef></c:xVal>`
  );
}

function valuesXml(tag: "val" | "yVal", values: (number | null)[], colLetter: string): string {
  const pts = values.map((v, i) => (v === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("");
  return (
    `<c:${tag}><c:numRef><c:f>Sheet1!$${colLetter}$2:$${colLetter}$${values.length + 1}</c:f>` +
    `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numCache></c:numRef></c:${tag}>`
  );
}

function txXml(name: string, colLetter: string): string {
  return (
    `<c:tx><c:strRef><c:f>Sheet1!$${colLetter}$1</c:f>` +
    `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXmlText(name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`
  );
}

function serFillXml(color: ModelColor): string {
  // A theme reference writes `a:schemeClr` (SPEC-077 AC-31); a literal keeps
  // the `a:srgbClr` form byte-for-byte.
  if (isThemeColor(color)) {
    return `<c:spPr><a:solidFill>${schemeClrXmlFor(color)}</a:solidFill></c:spPr>`;
  }
  const hex = cssColorToOoxml(color);
  return hex ? `<c:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></c:spPr>` : "";
}

function serXml(chart: ChartVisual, series: ChartSeries, index: number): string {
  const colLetter = columnLetterFor(index);
  let marker = "";
  if (chart.chartType === "lineMarkers") {
    marker = '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>';
  } else if (chart.chartType === "line") {
    marker = '<c:marker><c:symbol val="none"/></c:marker>';
  }
  const isScatter = chart.chartType === "scatter";
  const catOrX = isScatter ? xValXml(chart.categories) : catXml(chart.categories);
  const valTag = isScatter ? "yVal" : "val";
  return (
    `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>` +
    txXml(series.name, colLetter) +
    serFillXml(series.color) +
    marker +
    catOrX +
    valuesXml(valTag, series.values, colLetter) +
    (isScatter ? '<c:smooth val="0"/>' : "") +
    "</c:ser>"
  );
}

function dLblsXml(on: boolean): string {
  return (
    "<c:dLbls>" +
    '<c:showLegendKey val="0"/>' +
    `<c:showVal val="${on ? 1 : 0}"/>` +
    '<c:showCatName val="0"/>' +
    '<c:showSerName val="0"/>' +
    '<c:showPercent val="0"/>' +
    '<c:showBubbleSize val="0"/>' +
    "</c:dLbls>"
  );
}

const CAT_AX_ID = 111111111;
const VAL_AX_ID = 222222222;

function chartTypeXml(chart: ChartVisual, seriesXml: string): string {
  const dLbls = dLblsXml(chart.options.dataLabels);
  switch (chart.chartType) {
    case "columnClustered":
    case "columnStacked":
    case "barClustered": {
      const barDir = chart.chartType === "barClustered" ? "bar" : "col";
      const grouping = chart.chartType === "columnStacked" ? "stacked" : "clustered";
      return (
        "<c:barChart>" +
        `<c:barDir val="${barDir}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>` +
        seriesXml +
        dLbls +
        '<c:gapWidth val="150"/>' +
        (grouping === "stacked" ? '<c:overlap val="100"/>' : "") +
        `<c:axId val="${CAT_AX_ID}"/><c:axId val="${VAL_AX_ID}"/>` +
        "</c:barChart>"
      );
    }
    case "line":
    case "lineMarkers":
      return (
        "<c:lineChart>" +
        '<c:grouping val="standard"/><c:varyColors val="0"/>' +
        seriesXml +
        dLbls +
        `<c:marker val="${chart.chartType === "lineMarkers" ? 1 : 0}"/>` +
        `<c:axId val="${CAT_AX_ID}"/><c:axId val="${VAL_AX_ID}"/>` +
        "</c:lineChart>"
      );
    case "area":
      return (
        "<c:areaChart>" +
        '<c:grouping val="standard"/><c:varyColors val="0"/>' +
        seriesXml +
        dLbls +
        `<c:axId val="${CAT_AX_ID}"/><c:axId val="${VAL_AX_ID}"/>` +
        "</c:areaChart>"
      );
    case "pie":
      return `<c:pieChart><c:varyColors val="1"/>${seriesXml}${dLbls}<c:firstSliceAng val="0"/></c:pieChart>`;
    case "doughnut":
      return (
        `<c:doughnutChart><c:varyColors val="1"/>${seriesXml}${dLbls}` +
        '<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>'
      );
    case "scatter":
      return (
        "<c:scatterChart>" +
        '<c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>' +
        seriesXml +
        `<c:axId val="${CAT_AX_ID}"/><c:axId val="${VAL_AX_ID}"/>` +
        "</c:scatterChart>"
      );
  }
}

function axTitleXml(text: string): string {
  return (
    "<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r>" +
    `<a:t>${escapeXmlText(text)}</a:t>` +
    "</a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>"
  );
}

function axesXml(chart: ChartVisual): string {
  if (chart.chartType === "pie" || chart.chartType === "doughnut") {
    return "";
  }
  const isScatter = chart.chartType === "scatter";
  const catTag = isScatter ? "valAx" : "catAx";
  const catTitle = chart.options.categoryAxisTitle ? axTitleXml(chart.options.categoryAxisTitle) : "";
  const valTitle = chart.options.valueAxisTitle ? axTitleXml(chart.options.valueAxisTitle) : "";
  const gridlinesXml = chart.options.gridlines ? "<c:majorGridlines/>" : "";
  const minMaxXml =
    (chart.options.valueAxisMin !== undefined ? `<c:min val="${chart.options.valueAxisMin}"/>` : "") +
    (chart.options.valueAxisMax !== undefined ? `<c:max val="${chart.options.valueAxisMax}"/>` : "");
  const catAxXml =
    `<c:${catTag}><c:axId val="${CAT_AX_ID}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
    `<c:delete val="0"/><c:axPos val="b"/>${catTitle}<c:crossAx val="${VAL_AX_ID}"/></c:${catTag}>`;
  const valAxXml =
    `<c:valAx><c:axId val="${VAL_AX_ID}"/><c:scaling><c:orientation val="minMax"/>${minMaxXml}</c:scaling>` +
    `<c:delete val="0"/><c:axPos val="l"/>${gridlinesXml}${valTitle}<c:crossAx val="${CAT_AX_ID}"/></c:valAx>`;
  return catAxXml + valAxXml;
}

/** Regenerates a chart part's full `c:chartSpace` XML from the model — used
 * when a chart is `edited` (SPEC-052 "Rewriting a chart"). An untouched chart
 * never calls this: its part's original bytes are re-emitted verbatim from
 * the package's preserved parts instead. */
export function chartPartXml(chart: ChartVisual): string {
  // Pie/doughnut render (and so persist) only the first series — a real
  // c:pieChart/c:doughnutChart holds exactly one c:ser; the model keeps the
  // rest for the data sheet to still show (SPEC-052 "Chart catalogue").
  const seriesToWrite = SINGLE_SERIES_CHART_TYPES.has(chart.chartType) ? chart.series.slice(0, 1) : chart.series;
  const seriesXml = seriesToWrite.map((s, i) => serXml(chart, s, i)).join("");
  const typeXml = chartTypeXml(chart, seriesXml);
  const titleXml = chart.options.title ? axTitleXml(chart.options.title) : '<c:autoTitleDeleted val="1"/>';
  const legendXml =
    chart.options.legend === "none"
      ? ""
      : `<c:legend><c:legendPos val="${LEGEND_POS_TO_OOXML[chart.options.legend]}"/><c:overlay val="0"/></c:legend>`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<c:chartSpace xmlns:c="${C_NAMESPACE}" xmlns:a="${A_NAMESPACE}" xmlns:r="${R_NAMESPACE}">` +
    "<c:chart>" +
    titleXml +
    `<c:plotArea><c:layout/>${typeXml}${axesXml(chart)}</c:plotArea>` +
    legendXml +
    '<c:plotVisOnly val="1"/>' +
    "</c:chart>" +
    "</c:chartSpace>"
  );
}
