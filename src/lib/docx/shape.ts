// Reads/writes the DrawingML shape markup SPEC-051 models: a `wps:wsp`
// (`prstGeom`/`solidFill`/`ln`/`txbx`) inside a `wp:inline`/`wp:anchor`, the
// same container `drawing.ts` uses for a picture. Mirrors drawing.ts's own
// split (parse.../...Xml, ParsedDrawing-shaped result) so the two modules
// read side by side; this one never touches Image and drawing.ts never
// touches Shape, per SPEC-024's "no shipped call site changes".
import type { Block, Paragraph, Shape } from "../documentModel";
import type { ShapeGeometry, ShapeOutline } from "../core/shapes";
import { LINE_GEOMETRIES } from "../core/shapes";
import { SHAPE_DEFINITIONS } from "../core/shapeCatalogue";
import { escapeXmlAttribute, firstChildElement, serializeElement } from "../core/xml";
import { cssColorToOoxml, ooxmlColorToCss } from "../core/colors";
import { cmToEmu, emuToCm } from "../core/units";

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
// Our own extension namespace, used only to round-trip groupOrdinal (a
// purely editor-side concept — a flat groupOrdinal tag, not a real wpg:wgp
// group container, see "Specs this changes" / the grouping design note in
// the implement summary). Unknown-namespaced attributes are ignored by
// Word/LibreOffice, so this never corrupts a file another app opens.
const DE_NAMESPACE = "urn:docu-edit:shape";

const EMU_PER_POINT = 12700;
function ptToEmu(pt: number): number {
  return Math.round(pt * EMU_PER_POINT);
}
function emuToPt(emu: number): number {
  return emu / EMU_PER_POINT;
}
// OOXML rotation (a:xfrm/@rot) is 60000ths of a degree.
function degToOoxmlRot(deg: number): number {
  return Math.round(deg * 60000);
}
function ooxmlRotToDeg(rot: number): number {
  return Math.round(rot / 60000);
}

// The OOXML preset-name maps, both derived from the single SHAPE_DEFINITIONS
// table (SPEC-069 AC-15a) rather than maintained here beside it.
const GEOMETRY_TO_PRST: Record<string, string> = Object.fromEntries(
  SHAPE_DEFINITIONS.map((def) => [def.geometry, def.prst]),
);
const PRST_TO_GEOMETRY: Record<string, ShapeGeometry> = Object.fromEntries(
  SHAPE_DEFINITIONS.map((def) => [def.prst, def.geometry]),
);

/** Everything needed to parse a `wps:txbx`'s `w:txbxContent` into
 * `Shape.text.blocks` — injected by the caller (reader.ts) so this module
 * stays independent of the paragraph parser's own dependency graph
 * (styleCatalog/numberingCatalog/links live in reader.ts). */
export interface ShapeTextContext {
  parseParagraph: (pEl: Element) => Paragraph;
}

function vAlignOf(bodyPr: Element | undefined): "top" | "middle" | "bottom" {
  const anchor = bodyPr?.getAttribute("anchor");
  if (anchor === "ctr") {
    return "middle";
  }
  if (anchor === "b") {
    return "bottom";
  }
  return "top";
}

function parseShapeText(wsp: Element, ctx: ShapeTextContext): Shape["text"] | undefined {
  const txbx = firstChildElement(wsp, WPS_NAMESPACE, "txbx");
  if (!txbx) {
    return undefined;
  }
  const content = firstChildElement(txbx, WORD_NAMESPACE, "txbxContent");
  if (!content) {
    return undefined;
  }
  const blocks: Block[] = [];
  for (const child of Array.from(content.children)) {
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "p") {
      blocks.push(ctx.parseParagraph(child));
    }
  }
  if (blocks.length === 0) {
    blocks.push({ type: "paragraph", runs: [] });
  }
  const bodyPr = firstChildElement(wsp, WPS_NAMESPACE, "bodyPr");
  return { blocks, vAlign: vAlignOf(bodyPr) };
}

function parseOutline(spPr: Element): ShapeOutline | undefined {
  const ln = firstChildElement(spPr, A_NAMESPACE, "ln");
  if (!ln) {
    return undefined;
  }
  if (firstChildElement(ln, A_NAMESPACE, "noFill")) {
    return undefined;
  }
  const solidFill = firstChildElement(ln, A_NAMESPACE, "solidFill");
  const srgbClr = solidFill ? firstChildElement(solidFill, A_NAMESPACE, "srgbClr") : undefined;
  const color = srgbClr ? ooxmlColorToCss(srgbClr.getAttribute("val")) : undefined;
  if (!color) {
    return undefined;
  }
  const wAttr = ln.getAttribute("w");
  const widthPt = wAttr ? emuToPt(Number(wAttr)) : 0.75;
  let dash: ShapeOutline["dash"] = "solid";
  const prstDash = firstChildElement(ln, A_NAMESPACE, "prstDash");
  const val = prstDash?.getAttribute("val");
  if (val === "dash" || val === "dashDot" || val === "lgDash") {
    dash = "dash";
  } else if (val === "sysDot" || val === "sysDash") {
    dash = "dot";
  }
  return { color, widthPt, dash };
}

function parseFill(spPr: Element): string | undefined {
  const solidFill = firstChildElement(spPr, A_NAMESPACE, "solidFill");
  if (!solidFill) {
    return undefined;
  }
  const srgbClr = firstChildElement(solidFill, A_NAMESPACE, "srgbClr");
  return srgbClr ? ooxmlColorToCss(srgbClr.getAttribute("val")) : undefined;
}

/** One parsed `w:drawing` holding a shape (`wps:wsp`), the shape-flavoured
 * analogue of drawing.ts's ParsedDrawing. */
export interface ParsedShapeDrawing {
  shape: Shape;
}

/** Parses a `w:drawing` element into a Shape, or null when it is not a
 * single, recognisable `wps:wsp` this spec models — the caller's existing
 * SPEC-008 opaque-inline fallback then preserves it verbatim. */
export function parseShapeDrawing(
  drawingEl: Element,
  ctx: ShapeTextContext,
): ParsedShapeDrawing | null {
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
  const wsp = graphicData ? firstChildElement(graphicData, WPS_NAMESPACE, "wsp") : undefined;
  if (!wsp) {
    return null;
  }
  const spPr = firstChildElement(wsp, WPS_NAMESPACE, "spPr");
  if (!spPr) {
    return null;
  }
  const prstGeom = firstChildElement(spPr, A_NAMESPACE, "prstGeom");
  const prst = prstGeom?.getAttribute("prst") ?? "";
  const geometry = PRST_TO_GEOMETRY[prst];

  const xfrm = firstChildElement(spPr, A_NAMESPACE, "xfrm");
  const rotAttr = xfrm?.getAttribute("rot");
  const rotationDeg = rotAttr ? ooxmlRotToDeg(Number(rotAttr)) : 0;

  const docPr = firstChildElement(container, WP_NAMESPACE, "docPr");
  const altText = docPr?.getAttribute("descr") || undefined;
  const groupOrdinalAttr = docPr?.getAttributeNS(DE_NAMESPACE, "group");
  const groupOrdinal = groupOrdinalAttr ? Number(groupOrdinalAttr) : undefined;

  let wrap: Shape["wrap"] = "inline";
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
      offsetXCm = emuToCm(Number(posOffsetH.textContent));
    }
    const positionV = firstChildElement(container, WP_NAMESPACE, "positionV");
    const posOffsetV = positionV ? firstChildElement(positionV, WP_NAMESPACE, "posOffset") : undefined;
    if (posOffsetV?.textContent) {
      offsetYCm = emuToCm(Number(posOffsetV.textContent));
    }
    const relativeHeight = container.getAttribute("relativeHeight");
    if (relativeHeight) {
      zOrder = Number(relativeHeight) - 251658240;
    }
  }

  const shape: Shape = {
    type: "shape",
    geometry: geometry ?? "rect",
    widthCm: emuToCm(cx),
    heightCm: emuToCm(cy),
    wrap,
  };
  if (!geometry) {
    shape.unsupportedGeometryXml = serializeElement(prstGeom ?? spPr);
  }
  if (rotationDeg) {
    shape.rotationDeg = rotationDeg;
  }
  if (altText) {
    shape.altText = altText;
  }
  if (groupOrdinal !== undefined && Number.isFinite(groupOrdinal)) {
    shape.groupOrdinal = groupOrdinal;
  }
  if (offsetXCm !== undefined) {
    shape.offsetXCm = offsetXCm;
  }
  if (offsetYCm !== undefined) {
    shape.offsetYCm = offsetYCm;
  }
  if (wrap !== "inline") {
    shape.relativeToH = "column";
    shape.relativeToV = "paragraph";
  }
  if (zOrder !== undefined) {
    shape.zOrder = zOrder;
  }
  const fill = parseFill(spPr);
  if (fill) {
    shape.fill = fill;
  }
  const outline = parseOutline(spPr);
  if (outline) {
    shape.outline = outline;
  }
  const text = parseShapeText(wsp, ctx);
  if (text) {
    shape.text = text;
  }
  return { shape };
}

function dashXml(dash: ShapeOutline["dash"]): string {
  if (dash === "dash") {
    return '<a:prstDash val="dash"/>';
  }
  if (dash === "dot") {
    return '<a:prstDash val="sysDot"/>';
  }
  return "";
}

function outlineXml(outline: ShapeOutline | undefined): string {
  if (!outline) {
    return "<a:ln><a:noFill/></a:ln>";
  }
  const color = cssColorToOoxml(outline.color);
  if (!color) {
    return "<a:ln><a:noFill/></a:ln>";
  }
  return (
    `<a:ln w="${ptToEmu(outline.widthPt)}">` +
    `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
    dashXml(outline.dash) +
    "</a:ln>"
  );
}

function fillXml(fill: string | undefined): string {
  const color = fill ? cssColorToOoxml(fill) : undefined;
  return color ? `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` : "<a:noFill/>";
}

/** Serialises a Shape's text frame's blocks into `wps:txbx/w:txbxContent`
 * paragraphs — injected by the caller (writer.ts) so this module stays
 * independent of paragraphXml's own dependency graph. */
export interface ShapeTextWriteContext {
  paragraphXml: (paragraph: Paragraph) => string;
}

function anchorAttrOf(vAlign: "top" | "middle" | "bottom" | undefined): string {
  if (vAlign === "middle") {
    return "ctr";
  }
  if (vAlign === "bottom") {
    return "b";
  }
  return "t";
}

function txbxXml(shape: Shape, ctx: ShapeTextWriteContext): string {
  if (!shape.text) {
    return "";
  }
  const paragraphs = shape.text.blocks
    .map((block) => (block.type === "paragraph" ? ctx.paragraphXml(block) : ""))
    .join("");
  return (
    `<wps:txbx><w:txbxContent xmlns:w="${WORD_NAMESPACE}">${paragraphs || "<w:p/>"}</w:txbxContent></wps:txbx>` +
    `<wps:bodyPr wrap="square" anchor="${anchorAttrOf(shape.text.vAlign)}"/>`
  );
}

/** Serialises a Shape + the assigned docPrId into a `w:drawing` element
 * (wrapped by the caller in `<w:r>`, in `<w:p>` — mirrors drawing.ts's
 * drawingXml). Writes explicit `wp:positionH`/`wp:positionV` offsets (unlike
 * Image, which always writes 0) since Shape actually models a position. */
export function shapeDrawingXml(shape: Shape, docPrId: number, ctx: ShapeTextWriteContext): string {
  const cx = cmToEmu(shape.widthCm);
  const cy = cmToEmu(Math.max(0, shape.heightCm));
  const id = docPrId;
  const name = `Shape ${id}`;
  const descrAttr = shape.altText ? ` descr="${escapeXmlAttribute(shape.altText)}"` : "";
  const groupAttr = shape.groupOrdinal !== undefined ? ` de:group="${shape.groupOrdinal}"` : "";
  const docPr =
    `<wp:docPr id="${id}" name="${escapeXmlAttribute(name)}"${descrAttr}` +
    (groupAttr ? ` xmlns:de="${DE_NAMESPACE}"${groupAttr}` : "") +
    "/>";

  const prst = shape.unsupportedGeometryXml ? undefined : GEOMETRY_TO_PRST[shape.geometry];
  const geomXml = shape.unsupportedGeometryXml ?? `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>`;
  const rotAttr = shape.rotationDeg ? ` rot="${degToOoxmlRot(shape.rotationDeg)}"` : "";
  const isLine = LINE_GEOMETRIES.has(shape.geometry) && !shape.unsupportedGeometryXml;

  const spPr =
    "<wps:spPr>" +
    `<a:xfrm${rotAttr}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    geomXml +
    (isLine ? "" : fillXml(shape.fill)) +
    outlineXml(shape.outline) +
    "</wps:spPr>";

  const wsp =
    `<wps:wsp xmlns:wps="${WPS_NAMESPACE}">` +
    `<wps:cNvSpPr/>` +
    spPr +
    txbxXml(shape, ctx) +
    "</wps:wsp>";

  const graphic =
    `<a:graphic xmlns:a="${A_NAMESPACE}">` +
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    wsp +
    "</a:graphicData>" +
    "</a:graphic>";

  if (shape.wrap === "inline") {
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

  const behindDoc = shape.wrap === "behind" ? "1" : "0";
  let wrapXml: string;
  if (shape.wrap === "square") {
    wrapXml = '<wp:wrapSquare wrapText="bothSides"/>';
  } else if (shape.wrap === "topAndBottom") {
    wrapXml = "<wp:wrapTopAndBottom/>";
  } else {
    wrapXml = "<wp:wrapNone/>";
  }
  const offsetXEmu = cmToEmu(shape.offsetXCm ?? 0);
  const offsetYEmu = cmToEmu(shape.offsetYCm ?? 0);
  const relativeHeight = 251658240 + (shape.zOrder ?? 0);
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
