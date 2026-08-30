// Reads/writes the DrawingML *container* that carries a chart on a `.docx`
// page: a `c:chart` reference (inside `a:graphicData`, inside the same
// `wp:inline`/`wp:anchor` container drawing.ts/shape.ts use) pointing at a
// second package part, `word/charts/chartN.xml`. Mirrors drawing.ts/shape.ts's
// own split (parse.../...Xml) so the three modules read side by side.
//
// The chart *part* codec itself — `c:chartSpace`/`c:chart`/`c:plotArea`,
// identical schema in a `.pptx` deck — was hoisted into
// `src/lib/core/chartPartOoxml.ts` (SPEC-092) so the presentation product
// shares it; `parseChartXml`/`chartPartXml`/`ParsedChartPart` and the three
// chart constants are re-exported here unchanged so every existing `docx/*`
// importer keeps its path.
import type { Chart } from "../documentModel";
import type { ChartOptions, ChartSeries, ChartType } from "../core/charts";
import {
  escapeXmlAttribute,
  firstChildElement,
  resolvePartPath,
  serializeElement,
  type Relationship,
} from "../core/xml";
import { cmToEmu } from "../core/units";
import {
  CHART_CONTENT_TYPE,
  CHART_RELATIONSHIP,
  C_NAMESPACE,
  chartPartXml,
  parseChartXml,
  type ParsedChartPart,
} from "../core/chartPartOoxml";

// The hoisted chart-part codec, re-exported from its old home so
// `docx/chart.spec.ts`, `docx/chartRoundTrip.spec.ts` and the writer keep
// their import paths (SPEC-092).
export { CHART_CONTENT_TYPE, CHART_RELATIONSHIP, C_NAMESPACE, chartPartXml, parseChartXml, serializeElement };
export type { ParsedChartPart };

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
// Same editor-only extension namespace shape.ts's DE_NAMESPACE uses, for the
// same reason (round-tripping groupOrdinal, a purely editor-side concept no
// OOXML element carries) — an unrelated attribute in an unknown namespace,
// ignored by Word/LibreOffice.
const DE_NAMESPACE = "urn:docu-edit:shape";

// --- reading -----------------------------------------------------------------

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
