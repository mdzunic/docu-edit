// Reads/writes the DrawingML shape markup SPEC-051 models: a `wps:wsp`
// (`prstGeom`/`solidFill`/`ln`/`txbx`) inside a `wp:inline`/`wp:anchor`, the
// same container `drawing.ts` uses for a picture. Mirrors drawing.ts's own
// split (parse.../...Xml, ParsedDrawing-shaped result) so the two modules
// read side by side; this one never touches Image and drawing.ts never
// touches Shape, per SPEC-024's "no shipped call site changes".
import type { Block, Paragraph, Shape } from "../documentModel";
import type { ShapeOutline } from "../core/shapes";
import { LINE_GEOMETRIES } from "../core/shapes";
import { adjustmentGuideName, adjustmentToGuideValue } from "../core/shapeCatalogue";
import { GEOMETRY_TO_PRST, PRST_TO_GEOMETRY, parseAvLstAdjustments } from "../core/drawingMl";
import { arcToCubics, PATH_GRID, quadraticToCubic, type PathStep, type ShapePath, type ShapeSubpath } from "../core/shapePath";
import { childElements, escapeXmlAttribute, firstChildElement, serializeElement } from "../core/xml";
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

// The OOXML preset-name maps and the `a:avLst` adjustment parser now live in
// core/drawingMl.ts (SPEC-079's two-phase move) so the `.pptx` reader shares
// them; this module re-imports them, its behaviour unchanged.

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

/** The `a:avLst` a shape's adjustments serialise to (SPEC-069 AC-6a) —
 * empty when the shape carries none, so a plain shape is byte-unchanged. */
function avLstXml(shape: Shape): string {
  const adjustments = shape.adjustments;
  if (!adjustments || adjustments.length === 0) {
    return "<a:avLst/>";
  }
  const gds = adjustments
    .map((fraction, i) => `<a:gd name="${adjustmentGuideName(i, adjustments.length)}" fmla="val ${adjustmentToGuideValue(fraction)}"/>`)
    .join("");
  return `<a:avLst>${gds}</a:avLst>`;
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

// --- custom geometry (SPEC-071 `a:custGeom`) --------------------------------

/** A single `a:pt` as a fraction of the path's own `w`/`h` extent, or null
 * when the coordinate is a guide name rather than a number (a formula-driven
 * custGeom the model cannot carry — the caller then preserves it verbatim). */
function custPoint(el: Element | undefined, w: number, h: number): { x: number; y: number } | null {
  if (!el) {
    return null;
  }
  const x = Number(el.getAttribute("x"));
  const y = Number(el.getAttribute("y"));
  if (!Number.isFinite(x) || !Number.isFinite(y)) {
    return null;
  }
  return { x: x / w, y: y / h };
}

/** Parses `a:custGeom/a:pathLst` into a `ShapePath` in 0..1 space, converting
 * `a:arcTo`/`a:quadBezTo` to cubics (SPEC-071 "Primitive reduction"). Returns
 * null — so the caller keeps the SPEC-008 preservation path — when the
 * geometry references guides (formula-driven) or uses a command we do not
 * model. `converted` is true when any arc/quadratic was reduced, so the
 * caller can preserve the source XML for a lossless same-format save. */
function parseCustGeom(custGeom: Element): { path: ShapePath; converted: boolean } | null {
  const pathLst = firstChildElement(custGeom, A_NAMESPACE, "pathLst");
  if (!pathLst) {
    return null;
  }
  const subpaths: ShapeSubpath[] = [];
  let converted = false;
  for (const pathEl of childElements(pathLst, A_NAMESPACE, "path")) {
    const w = Number(pathEl.getAttribute("w")) || PATH_GRID;
    const h = Number(pathEl.getAttribute("h")) || PATH_GRID;
    const steps: PathStep[] = [];
    let cur = { x: 0, y: 0 };
    let closed = false;
    const cmds = Array.from(pathEl.children).filter((c) => c.namespaceURI === A_NAMESPACE);
    for (const cmd of cmds) {
      const pts = Array.from(cmd.children).filter((c) => c.namespaceURI === A_NAMESPACE && c.localName === "pt");
      if (cmd.localName === "moveTo") {
        const p = custPoint(pts[0], w, h);
        if (!p) {
          return null;
        }
        steps.push({ op: "move", x: p.x, y: p.y });
        cur = p;
      } else if (cmd.localName === "lnTo") {
        const p = custPoint(pts[0], w, h);
        if (!p) {
          return null;
        }
        steps.push({ op: "line", x: p.x, y: p.y });
        cur = p;
      } else if (cmd.localName === "cubicBezTo") {
        const c1 = custPoint(pts[0], w, h);
        const c2 = custPoint(pts[1], w, h);
        const end = custPoint(pts[2], w, h);
        if (!c1 || !c2 || !end) {
          return null;
        }
        steps.push({ op: "cubic", x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y, x: end.x, y: end.y });
        cur = end;
      } else if (cmd.localName === "quadBezTo") {
        const c = custPoint(pts[0], w, h);
        const end = custPoint(pts[1], w, h);
        if (!c || !end) {
          return null;
        }
        steps.push({ op: "cubic", ...quadraticToCubic(cur, c, end) });
        cur = end;
        converted = true;
      } else if (cmd.localName === "arcTo") {
        const wR = Number(cmd.getAttribute("wR"));
        const hR = Number(cmd.getAttribute("hR"));
        const stAng = Number(cmd.getAttribute("stAng"));
        const swAng = Number(cmd.getAttribute("swAng"));
        if (![wR, hR, stAng, swAng].every(Number.isFinite)) {
          return null;
        }
        const rx = wR / w;
        const ry = hR / h;
        const start = (stAng / 60000) * (Math.PI / 180);
        const sweep = (swAng / 60000) * (Math.PI / 180);
        const cx = cur.x - rx * Math.cos(start);
        const cy = cur.y - ry * Math.sin(start);
        for (const seg of arcToCubics(cx, cy, rx, ry, start, sweep)) {
          steps.push({ op: "cubic", ...seg });
          cur = { x: seg.x, y: seg.y };
        }
        converted = true;
      } else if (cmd.localName === "close") {
        steps.push({ op: "close" });
        closed = true;
      } else {
        return null;
      }
    }
    if (steps.length > 0) {
      subpaths.push({ steps, closed });
    }
  }
  if (subpaths.length === 0) {
    return null;
  }
  const path: ShapePath = { subpaths };
  // Per-vertex smoothness is our own concept (SPEC-071) with no OOXML home,
  // round-tripped through an ignored de: attribute so point-edit mode
  // restores the right handles on reopen rather than inferring them.
  const smoothAttr = custGeom.getAttributeNS(DE_NAMESPACE, "smooth");
  if (smoothAttr) {
    path.smooth = smoothAttr.split("").map((c) => c === "1");
  }
  return { path, converted };
}

/** Serializes a `ShapePath` as `a:custGeom` on the shared 21600 grid
 * (AC-26). `a:gdLst`/`a:ahLst`/`a:cxnLst` are deliberately not written. */
export function custGeomXml(path: ShapePath): string {
  const g = PATH_GRID;
  const smoothAttr =
    path.smooth && path.smooth.length > 0
      ? ` xmlns:de="${DE_NAMESPACE}" de:smooth="${path.smooth.map((b) => (b ? "1" : "0")).join("")}"`
      : "";
  const pt = (x: number, y: number): string => `<a:pt x="${Math.round(x * g)}" y="${Math.round(y * g)}"/>`;
  let steps = "";
  for (const subpath of path.subpaths) {
    for (const step of subpath.steps) {
      if (step.op === "move") {
        steps += `<a:moveTo>${pt(step.x, step.y)}</a:moveTo>`;
      } else if (step.op === "line") {
        steps += `<a:lnTo>${pt(step.x, step.y)}</a:lnTo>`;
      } else if (step.op === "cubic") {
        steps += `<a:cubicBezTo>${pt(step.x1, step.y1)}${pt(step.x2, step.y2)}${pt(step.x, step.y)}</a:cubicBezTo>`;
      } else {
        steps += "<a:close/>";
      }
    }
  }
  return `<a:custGeom${smoothAttr}><a:pathLst><a:path w="${g}" h="${g}">${steps}</a:path></a:pathLst></a:custGeom>`;
}

/** The size + wrap a `wps:wsp` cannot supply for itself: for a top-level
 * shape they come from its `wp:inline`/`wp:anchor` container, for a group
 * member from the group's `chOff`/`chExt` mapping applied to the member's own
 * `a:xfrm` (see group.ts). */
export interface WspBox {
  widthCm: number;
  heightCm: number;
  wrap: Shape["wrap"];
}

/** Builds a Shape from a bare `wps:wsp` and an externally-resolved box —
 * everything the shape markup itself carries (geometry, custom path, fill,
 * outline, rotation, text), and nothing that lives on a container. Returns
 * null when the `wsp` has no `wps:spPr` to read.
 *
 * Split out of parseShapeDrawing so a `wpg:wgp` group member — which has no
 * container of its own — reads through exactly the same code as a top-level
 * shape, rather than a second parser drifting from this one. The caller adds
 * the container-derived fields: alt text, offsets, `relativeTo*`, z-order and
 * `groupOrdinal`. */
export function shapeFromWsp(wsp: Element, box: WspBox, ctx: ShapeTextContext): Shape | null {
  const spPr = firstChildElement(wsp, WPS_NAMESPACE, "spPr");
  if (!spPr) {
    return null;
  }
  const prstGeom = firstChildElement(spPr, A_NAMESPACE, "prstGeom");
  let geometry = PRST_TO_GEOMETRY[prstGeom?.getAttribute("prst") ?? ""];
  // SPEC-070 "Untangling the arrow preset": a `prst="line"` carrying an
  // `a:tailEnd` is the model's `arrow` (a line with an end arrowhead). The
  // legacy `straightConnector1` → `arrow` mapping stays in PRST_TO_GEOMETRY
  // for documents saved before this change.
  if (geometry === "line") {
    const lnEl = firstChildElement(spPr, A_NAMESPACE, "ln");
    const tailEnd = lnEl ? firstChildElement(lnEl, A_NAMESPACE, "tailEnd") : undefined;
    if (tailEnd && tailEnd.getAttribute("type") !== "none") {
      geometry = "arrow";
    }
  }

  const xfrm = firstChildElement(spPr, A_NAMESPACE, "xfrm");
  const rotAttr = xfrm?.getAttribute("rot");
  const rotationDeg = rotAttr ? ooxmlRotToDeg(Number(rotAttr)) : 0;

  const shape: Shape = {
    type: "shape",
    geometry: geometry ?? "rect",
    widthCm: box.widthCm,
    heightCm: box.heightCm,
    wrap: box.wrap,
  };
  const custGeom = !geometry ? firstChildElement(spPr, A_NAMESPACE, "custGeom") : undefined;
  const parsedPath = custGeom ? parseCustGeom(custGeom) : null;
  if (parsedPath) {
    // A modelled path renders its real outline (AC-4); geometry stays the
    // "rect" bounding-box stand-in and nothing is preserved opaquely.
    shape.path = parsedPath.path;
    if (parsedPath.converted && custGeom) {
      shape.pathSource = { format: "docx", xml: serializeElement(custGeom) };
    }
  } else if (!geometry) {
    shape.unsupportedGeometryXml = serializeElement(prstGeom ?? custGeom ?? spPr);
  } else if (prstGeom) {
    const adjustments = parseAvLstAdjustments(prstGeom);
    if (adjustments.length > 0) {
      shape.adjustments = adjustments;
    }
  }
  if (rotationDeg) {
    shape.rotationDeg = rotationDeg;
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
  return shape;
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

  const shape = shapeFromWsp(wsp, { widthCm: emuToCm(cx), heightCm: emuToCm(cy), wrap }, ctx);
  if (!shape) {
    return null;
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

/** The outline `a:ln`. `tailEnd` (SPEC-070 "Untangling the arrow preset") adds
 * an `<a:tailEnd type="triangle"/>` so the model's `arrow` geometry — written
 * as `prst="line"` — reads back as an arrow rather than a plain line. */
function outlineXml(outline: ShapeOutline | undefined, tailEnd?: string): string {
  const tail = tailEnd ? `<a:tailEnd type="${tailEnd}"/>` : "";
  if (!outline) {
    return `<a:ln>${tail || "<a:noFill/>"}</a:ln>`;
  }
  const color = cssColorToOoxml(outline.color);
  if (!color) {
    return `<a:ln>${tail || "<a:noFill/>"}</a:ln>`;
  }
  return (
    `<a:ln w="${ptToEmu(outline.widthPt)}">` +
    `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
    dashXml(outline.dash) +
    tail +
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

  // SPEC-070: the model's `arrow` is "a line with an end arrowhead", not a
  // connector preset — it writes as `prst="line"` plus an `a:tailEnd`, the
  // same representation `.odt` already uses. `straightConnector1` is no longer
  // written (its read-only compatibility mapping stays in PRST_TO_GEOMETRY so
  // documents this app already saved still open as arrows).
  const isArrow = shape.geometry === "arrow" && !shape.unsupportedGeometryXml && !shape.path;
  // A path shape writes as `a:custGeom` (AC-26). When its path was reduced
  // from arcs/quadratics on read (`pathSource`) a same-format save re-emits
  // the verbatim source unchanged, losing nothing; otherwise the model's
  // move/line/cubic/close is written on the 21600 grid.
  const prst = shape.unsupportedGeometryXml || shape.path ? undefined : isArrow ? "line" : GEOMETRY_TO_PRST[shape.geometry];
  const geomXml = shape.path
    ? shape.pathSource?.format === "docx"
      ? shape.pathSource.xml
      : custGeomXml(shape.path)
    : (shape.unsupportedGeometryXml ?? `<a:prstGeom prst="${prst}">${avLstXml(shape)}</a:prstGeom>`);
  const rotAttr = shape.rotationDeg ? ` rot="${degToOoxmlRot(shape.rotationDeg)}"` : "";
  const isLine = LINE_GEOMETRIES.has(shape.geometry) && !shape.unsupportedGeometryXml && !shape.path;

  const spPr =
    "<wps:spPr>" +
    `<a:xfrm${rotAttr}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    geomXml +
    (isLine ? "" : fillXml(shape.fill)) +
    outlineXml(shape.outline, isArrow ? "triangle" : undefined) +
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
