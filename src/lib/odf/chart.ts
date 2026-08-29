// Reads/writes the ODF *container* that carries a chart on an `.odt` page: a
// `draw:frame` holding a `draw:object` reference to an embedded document
// subpackage (`Object N/`). Mirrors drawing.ts/shape.ts's own split
// (parse.../...Xml) so the three modules read side by side.
//
// The embedded object's own `content.xml` codec — `office:chart`/`chart:chart`
// with its `table:table` data cache, identical schema in an `.odp` deck — was
// hoisted into `src/lib/core/chartPartOdf.ts` (SPEC-092) so the presentation
// product shares it; `parseOdfChartXml`/`odfChartContentXml`/
// `ParsedOdfChartPart` and the two chart constants are re-exported here so
// every existing `odf/*` importer keeps its path. The one signature change:
// `odfChartContentXml(chart, widthCm, heightCm)` now takes geometry
// explicitly, since the neutral `ChartVisual` it accepts carries no size.
import type { Chart } from "../documentModel";
import type { ChartOptions, ChartSeries, ChartType } from "../core/charts";
import { escapeXmlAttribute, serializeElement } from "../core/xml";
import { cmToOdfLength, odfLengthToCm } from "./units";
import { DRAW_NAMESPACE, SVG_NAMESPACE, XLINK_NAMESPACE, svgAttr, textAttr, xlinkAttr } from "../core/odfXml";
import {
  CHART_NAMESPACE,
  CHART_OBJECT_CONTENT_TYPE,
  odfChartContentXml,
  parseOdfChartXml,
  type ParsedOdfChartPart,
} from "../core/chartPartOdf";

// The hoisted chart-object codec, re-exported from its old home so
// `odf/chart.spec.ts` and `odf/chartRoundTrip.spec.ts` keep their import paths
// (SPEC-092).
export { CHART_NAMESPACE, CHART_OBJECT_CONTENT_TYPE, odfChartContentXml, parseOdfChartXml, serializeElement };
export type { ParsedOdfChartPart };

function firstChildNS(parent: Element, ns: string, localName: string): Element | undefined {
  for (const child of Array.from(parent.children)) {
    if (child.namespaceURI === ns && child.localName === localName) {
      return child;
    }
  }
  return undefined;
}

// --- reading -----------------------------------------------------------------

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
