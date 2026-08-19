// Reads a native `draw:g` group (SPEC-057) from a content.xml element into a
// DrawingGroup container. `draw:g` defines no geometry of its own — its
// children are positioned in the anchor's coordinate space directly — so the
// container's box is the union of the children's unrotated rectangles, its
// offsets the union origin, and each member's offset is its child position
// minus that origin. Anchor type, style (wrap), `draw:z-index` and
// `svg:title`/`svg:desc` sit on the `draw:g` and read by the same rules a
// single shape uses. Shape and nested-group members are modelled; any other
// member kind (a `draw:frame` image/chart) keeps the drawing on SPEC-008
// opaque preservation for now.
import { firstChildElement } from "../core/xml";
import { svgAttr } from "./xml";
import { odfLengthToCm } from "./units";
import { parseShapeElement, type OdfShapeTextContext } from "./shape";
import { parseDrawFrame } from "./drawing";
import type { MediaPart } from "../core/media";
import type { DrawingGroup, DrawingMember, ImageWrapMode } from "../documentModel";

const DRAW_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const TEXT_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

export interface OdfGroupReadContext extends OdfShapeTextContext {
  contentDoc: Document;
  stylesDoc: Document | undefined;
  graphicWrapStyles: Map<string, ImageWrapMode>;
  /** Package parts + a media sink, so a `draw:frame` picture member resolves
   * to an Image and its bytes reach `DocumentModel.media` (SPEC-024/-057). */
  parts?: Map<string, Uint8Array>;
  collectMedia?: (part: MediaPart) => void;
}

function memberOf(el: Element, ctx: OdfGroupReadContext): DrawingMember | null {
  if (el.namespaceURI === DRAW_NAMESPACE && (el.localName === "custom-shape" || el.localName === "line")) {
    const parsed = parseShapeElement(el, ctx.contentDoc, ctx.stylesDoc, ctx);
    return parsed?.shape ?? null;
  }
  if (el.namespaceURI === DRAW_NAMESPACE && el.localName === "frame") {
    if (!ctx.parts) {
      return null;
    }
    const parsed = parseDrawFrame(el, ctx.parts, ctx.graphicWrapStyles);
    if (!parsed) {
      return null;
    }
    ctx.collectMedia?.(parsed.media);
    // `draw:frame` positions its members by svg:x/svg:y in the anchor frame,
    // which parseDrawFrame (a standalone-image reader) does not read.
    const image = { ...parsed.image };
    image.offsetXCm = odfLengthToCm(svgAttr(el, "x")) ?? 0;
    image.offsetYCm = odfLengthToCm(svgAttr(el, "y")) ?? 0;
    return image;
  }
  if (el.namespaceURI === DRAW_NAMESPACE && el.localName === "g") {
    return parseGroupElement(el, ctx);
  }
  return null;
}

/** Parses a `draw:g` element into a DrawingGroup, computing its box from the
 * members' absolute positions. Returns null when any child is a member kind
 * this reader does not model (so the caller keeps the SPEC-008 opaque
 * fallback), or when there are fewer than two modelled members. */
export function parseGroupElement(gEl: Element, ctx: OdfGroupReadContext): DrawingGroup | null {
  const childElements = Array.from(gEl.children).filter(
    (c) => !(c.namespaceURI === SVG_NAMESPACE && (c.localName === "title" || c.localName === "desc")),
  );
  const members: DrawingMember[] = [];
  for (const child of childElements) {
    const member = memberOf(child, ctx);
    if (!member) {
      return null;
    }
    members.push(member);
  }
  if (members.length < 2) {
    return null;
  }
  // Members carry absolute anchor-frame offsets; fold the union origin out.
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const member of members) {
    const x = member.offsetXCm ?? 0;
    const y = member.offsetYCm ?? 0;
    minX = Math.min(minX, x);
    minY = Math.min(minY, y);
    maxX = Math.max(maxX, x + member.widthCm);
    maxY = Math.max(maxY, y + member.heightCm);
  }
  const normalizedMembers = members.map((member) => {
    const next: DrawingMember = {
      ...member,
      wrap: "square",
      offsetXCm: (member.offsetXCm ?? 0) - minX,
      offsetYCm: (member.offsetYCm ?? 0) - minY,
    };
    delete next.relativeToH;
    delete next.relativeToV;
    delete next.zOrder;
    return next;
  });

  const anchorType = gEl.getAttributeNS(TEXT_NAMESPACE, "anchor-type");
  const styleName = gEl.getAttributeNS(DRAW_NAMESPACE, "style-name");
  const wrap: ImageWrapMode = anchorType === "as-char" ? "inline" : (styleName ? ctx.graphicWrapStyles.get(styleName) ?? "square" : "square");
  const titleEl = firstChildElement(gEl, SVG_NAMESPACE, "title");
  const descEl = firstChildElement(gEl, SVG_NAMESPACE, "desc");
  const altText = (descEl?.textContent || titleEl?.textContent) || undefined;
  const zIndexAttr = gEl.getAttributeNS(DRAW_NAMESPACE, "z-index");

  const group: DrawingGroup = {
    type: "drawingGroup",
    widthCm: maxX - minX,
    heightCm: maxY - minY,
    wrap,
    members: normalizedMembers,
  };
  if (wrap !== "inline") {
    group.relativeToH = "column";
    group.relativeToV = "paragraph";
    group.offsetXCm = minX;
    group.offsetYCm = minY;
  }
  if (altText) {
    group.altText = altText;
  }
  if (zIndexAttr) {
    const z = Number(zIndexAttr);
    if (Number.isFinite(z) && z !== 0) {
      group.zOrder = z;
    }
  }
  return group;
}
