// Reads/writes the DrawingML chart markup SPEC-052 models: a `c:chart`
// reference (inside `a:graphicData`, inside the same `wp:inline`/`wp:anchor`
// container drawing.ts/shape.ts use) pointing at a second package part,
// `word/charts/chartN.xml`, holding `c:chartSpace`/`c:chart`/`c:plotArea`.
// Mirrors drawing.ts/shape.ts's own split (parse.../...Xml) so the three
// modules read side by side.
//
// Unlike an image or a shape, a chart's *content* can live in one of two
// states: `sourcePart`'s bytes verbatim (an untouched chart, SPEC-008's
// byte-identical guarantee) or freshly regenerated from the model (an edited
// chart). This module only ever produces the "regenerate from model" half —
// `chartPartXml`; the writer decides which one to emit for a given chart.
import type { Chart } from "../documentModel";
import type { ChartOptions, ChartSeries, ChartType } from "../core/charts";
import { DEFAULT_CHART_COLORS, SINGLE_SERIES_CHART_TYPES } from "../core/charts";
import {
  childElements,
  escapeXmlAttribute,
  escapeXmlText,
  firstChildElement,
  parseXml,
  resolvePartPath,
  serializeElement,
  type Relationship,
  XmlParseError,
} from "../core/xml";
import { cssColorToOoxml } from "../core/colors";
import { cmToEmu } from "../core/units";

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const C_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
// Same editor-only extension namespace shape.ts's DE_NAMESPACE uses, for the
// same reason (round-tripping groupOrdinal, a purely editor-side concept no
// OOXML element carries) — an unrelated attribute in an unknown namespace,
// ignored by Word/LibreOffice.
const DE_NAMESPACE = "urn:docu-edit:shape";

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

function seriesColorOf(serEl: Element, fallback: string): string {
  const spPr = firstChildElement(serEl, C_NAMESPACE, "spPr");
  const solidFill = spPr && firstChildElement(spPr, A_NAMESPACE, "solidFill");
  const srgbClr = solidFill && firstChildElement(solidFill, A_NAMESPACE, "srgbClr");
  const val = srgbClr?.getAttribute("val");
  return val ? `#${val.toLowerCase()}` : fallback;
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

/** One parsed `w:drawing` holding a `c:chart` reference: the modelled Chart
 * (placement fields filled in by the caller, drawing.ts-style) plus the
 * referenced part's own name and raw text, so the caller can both build the
 * model and let the part's bytes flow into `DocumentModel.opaque` through
 * the reader's existing generic "everything unconsumed" sweep (SPEC-008) —
 * this module never removes the part from that sweep, since `sourcePart` is
 * a name into that exact list, not a second copy of the bytes. */
export interface ParsedChartDrawing {
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
  offsetXCm?: number;
  offsetYCm?: number;
  relativeToH?: Chart["relativeToH"];
  relativeToV?: Chart["relativeToV"];
  zOrder?: number;
  groupOrdinal?: number;
}

/** Parses a `w:drawing` element into a ParsedChartDrawing, or null when it is
 * not a single, simply-wrapped chart reference (the caller's existing
 * opaque-inline fallback then preserves it verbatim, SPEC-008) — mirrors
 * drawing.ts's parseDrawing/shape.ts's parseShapeDrawing. */
export function parseChartDrawing(
  drawingEl: Element,
  mainRels: Map<string, Relationship>,
  mainPartPath: string,
  parts: Map<string, Uint8Array>,
): ParsedChartDrawing | null {
  const children = Array.from(drawingEl.children).filter((el) => el.namespaceURI === WP_NAMESPACE);
  if (children.length !== 1) {
    return null;
  }
  const container = children[0];
  const isAnchor = container.localName === "anchor";
  if (!isAnchor && container.localName !== "inline") {
    return null;
  }

  const extent = firstChildElement(container, WP_NAMESPACE, "extent");
  const cx = extent ? Number(extent.getAttribute("cx")) : NaN;
  const cy = extent ? Number(extent.getAttribute("cy")) : NaN;
  if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx < 0 || cy < 0) {
    return null;
  }

  const graphic = firstChildElement(container, A_NAMESPACE, "graphic");
  const graphicData = graphic ? firstChildElement(graphic, A_NAMESPACE, "graphicData") : undefined;
  const chartRef = graphicData ? firstChildElement(graphicData, C_NAMESPACE, "chart") : undefined;
  const relId = chartRef ? chartRef.getAttributeNS(R_NAMESPACE, "id") : null;
  if (!relId) {
    return null;
  }
  const rel = mainRels.get(relId);
  if (!rel || rel.type !== CHART_RELATIONSHIP) {
    return null;
  }
  const partPath = resolvePartPath(mainPartPath, rel.target);
  const bytes = parts.get(partPath);
  if (!bytes) {
    return null;
  }
  const xmlText = new TextDecoder("utf-8").decode(bytes);
  const parsed = parseChartXml(xmlText);

  let wrap: Chart["wrap"] = "inline";
  let offsetXCm: number | undefined;
  let offsetYCm: number | undefined;
  let zOrder: number | undefined;
  if (isAnchor) {
    if (firstChildElement(container, WP_NAMESPACE, "wrapSquare")) {
      wrap = "square";
    } else if (firstChildElement(container, WP_NAMESPACE, "wrapTopAndBottom")) {
      wrap = "topAndBottom";
    } else if (firstChildElement(container, WP_NAMESPACE, "wrapNone")) {
      wrap = container.getAttribute("behindDoc") === "1" ? "behind" : "infront";
    } else {
      return null;
    }
    const positionH = firstChildElement(container, WP_NAMESPACE, "positionH");
    const posOffsetH = positionH ? firstChildElement(positionH, WP_NAMESPACE, "posOffset") : undefined;
    if (posOffsetH?.textContent) {
      offsetXCm = Number(posOffsetH.textContent) / (914400 / 2.54);
    }
    const positionV = firstChildElement(container, WP_NAMESPACE, "positionV");
    const posOffsetV = positionV ? firstChildElement(positionV, WP_NAMESPACE, "posOffset") : undefined;
    if (posOffsetV?.textContent) {
      offsetYCm = Number(posOffsetV.textContent) / (914400 / 2.54);
    }
    const relativeHeight = container.getAttribute("relativeHeight");
    if (relativeHeight) {
      zOrder = Number(relativeHeight) - 251658240;
    }
  }

  const docPr = firstChildElement(container, WP_NAMESPACE, "docPr");
  const altText = docPr?.getAttribute("descr") || undefined;
  const groupOrdinalAttr = docPr?.getAttributeNS(DE_NAMESPACE, "group");
  const groupOrdinal = groupOrdinalAttr ? Number(groupOrdinalAttr) : undefined;

  const result: ParsedChartDrawing = {
    chartType: parsed?.chartType ?? "columnClustered",
    categories: parsed?.categories ?? [],
    series: parsed?.series ?? [],
    options: parsed?.options ?? { legend: "none", dataLabels: false, gridlines: false },
    sourcePart: partPath,
    widthCm: (cx / 914400) * 2.54,
    heightCm: (cy / 914400) * 2.54,
    wrap,
  };
  if (!parsed) {
    result.unsupportedChartXml = xmlText;
  }
  if (altText) {
    result.altText = altText;
  }
  if (groupOrdinal !== undefined && Number.isFinite(groupOrdinal)) {
    result.groupOrdinal = groupOrdinal;
  }
  if (offsetXCm !== undefined) {
    result.offsetXCm = offsetXCm;
  }
  if (offsetYCm !== undefined) {
    result.offsetYCm = offsetYCm;
  }
  if (wrap !== "inline") {
    result.relativeToH = "column";
    result.relativeToV = "paragraph";
  }
  if (zOrder !== undefined) {
    result.zOrder = zOrder;
  }
  return result;
}

// --- writing -------------------------------------------------------------

/** Serialises a Chart + the relationship id its part was written under into
 * a `w:drawing` element — the chart-flavoured sibling of drawing.ts's
 * drawingXml/shape.ts's shapeDrawingXml. Always regenerated fresh from the
 * model's placement fields, whether or not the chart's own content
 * (`edited`) was regenerated — moving/resizing a chart never touches its
 * part. */
export function chartDrawingXml(chart: Chart, relId: string, docPrId: number): string {
  const cx = cmToEmu(chart.widthCm);
  const cy = cmToEmu(Math.max(0, chart.heightCm));
  const id = docPrId;
  const name = `Chart ${id}`;
  const descrAttr = chart.altText ? ` descr="${escapeXmlAttribute(chart.altText)}"` : "";
  const groupAttr = chart.groupOrdinal !== undefined ? ` de:group="${chart.groupOrdinal}"` : "";
  const docPr =
    `<wp:docPr id="${id}" name="${escapeXmlAttribute(name)}"${descrAttr}` +
    (groupAttr ? ` xmlns:de="${DE_NAMESPACE}"${groupAttr}` : "") +
    "/>";
  const graphic =
    `<a:graphic xmlns:a="${A_NAMESPACE}">` +
    `<a:graphicData uri="${C_NAMESPACE}">` +
    `<c:chart xmlns:c="${C_NAMESPACE}" xmlns:r="${R_NAMESPACE}" r:id="${relId}"/>` +
    "</a:graphicData>" +
    "</a:graphic>";

  if (chart.wrap === "inline") {
    return (
      `<w:drawing xmlns:wp="${WP_NAMESPACE}">` +
      '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      docPr +
      "<wp:cNvGraphicFramePr/>" +
      graphic +
      "</wp:inline>" +
      "</w:drawing>"
    );
  }

  const behindDoc = chart.wrap === "behind" ? "1" : "0";
  const wrapXml =
    chart.wrap === "square"
      ? '<wp:wrapSquare wrapText="bothSides"/>'
      : chart.wrap === "topAndBottom"
        ? "<wp:wrapTopAndBottom/>"
        : "<wp:wrapNone/>";
  const offsetXEmu = cmToEmu(chart.offsetXCm ?? 0);
  const offsetYEmu = cmToEmu(chart.offsetYCm ?? 0);
  const relativeHeight = 251658240 + (chart.zOrder ?? 0);
  return (
    `<w:drawing xmlns:wp="${WP_NAMESPACE}">` +
    `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"` +
    ` relativeHeight="${relativeHeight}" behindDoc="${behindDoc}" locked="0"` +
    ' layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/>' +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${offsetXEmu}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${offsetYEmu}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    wrapXml +
    docPr +
    "<wp:cNvGraphicFramePr/>" +
    graphic +
    "</wp:anchor>" +
    "</w:drawing>"
  );
}

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

function serFillXml(color: string): string {
  const hex = cssColorToOoxml(color);
  return hex ? `<c:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></c:spPr>` : "";
}

function serXml(chart: Chart, series: ChartSeries, index: number): string {
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

function chartTypeXml(chart: Chart, seriesXml: string): string {
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

function axesXml(chart: Chart): string {
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
 * when `chart.edited` is true (SPEC-052 "Rewriting a chart"). An untouched
 * chart never calls this: its part's original bytes are re-emitted verbatim
 * from `DocumentModel.opaque` instead. */
export function chartPartXml(chart: Chart): string {
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

// Re-exported purely so callers that already have an Element (rather than
// text) can round-trip an unsupported chart's XML the same way shape.ts's
// unsupportedGeometryXml capture does.
export { serializeElement };
