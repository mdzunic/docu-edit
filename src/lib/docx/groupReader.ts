// Reads a native `wpg:wgp` drawing group (SPEC-057) from a `w:drawing`
// element into a DrawingGroup container. The anchor supplies the container's
// wrap/position/z-order/alt text (same rules as a single shape); the group's
// own `a:xfrm` supplies its box (`a:ext`), rotation (`@rot`) and child
// coordinate space (`a:chOff`/`a:chExt`) — a member in child space `off/ext`
// has frame rectangle `(off − chOff)·s` at size `ext·s`, `s = ext/chExt`.
// Shape and nested-group members are modelled; a member kind this reader does
// not yet model (a picture, a chart frame) seals the container — it is
// preserved verbatim so it still round-trips, moves and ungroups.
import { firstChildElement, resolvePartPath, type Relationship } from "../core/xml";
import { emuToCm } from "../core/units";
import type { MediaPart } from "../core/media";
import { parseGroupMemberWsp } from "./shape";
import { parseGroupMemberImage } from "./drawing";
import { parseChartXml, C_NAMESPACE, CHART_RELATIONSHIP } from "./chart";
import type { Chart, DrawingGroup, DrawingMember, Paragraph } from "../documentModel";

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const WPS_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WPG_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const GROUP_URI = "wordprocessingGroup";

/** Parses a `graphicFrame` chart member of a group (SPEC-057 AC-55) into a
 * Chart, or null when it is not a resolvable chart frame (the caller then
 * seals the container). The frame's own `xfrm` positions it in the parent's
 * child space (the same transform shapes/images apply); its
 * `a:graphic/a:graphicData[@uri chart]/c:chart@r:id` resolves to the chart
 * part, parsed by the shipped standalone reader so a chart member and a
 * standalone chart model identically. */
function parseGroupMemberChart(frame: Element, transform: ChildTransform, ctx: GroupShapeTextContext): Chart | null {
  if (!ctx.mainRels || ctx.mainPartPath === undefined || !ctx.parts) {
    return null;
  }
  const xfrm = Array.from(frame.children).find((el) => el.localName === "xfrm");
  const off = xfrm ? firstChildElement(xfrm, A_NAMESPACE, "off") : undefined;
  const ext = xfrm ? firstChildElement(xfrm, A_NAMESPACE, "ext") : undefined;
  const offX = off ? Number(off.getAttribute("x")) || 0 : 0;
  const offY = off ? Number(off.getAttribute("y")) || 0 : 0;
  const extCx = ext ? Number(ext.getAttribute("cx")) : NaN;
  const extCy = ext ? Number(ext.getAttribute("cy")) : NaN;
  if (!Number.isFinite(extCx) || !Number.isFinite(extCy)) {
    return null;
  }
  const graphic = firstChildElement(frame, A_NAMESPACE, "graphic");
  const graphicData = graphic ? firstChildElement(graphic, A_NAMESPACE, "graphicData") : undefined;
  const chartRef = graphicData ? firstChildElement(graphicData, C_NAMESPACE, "chart") : undefined;
  const relId = chartRef ? chartRef.getAttributeNS(R_NAMESPACE, "id") : null;
  if (!relId) {
    return null;
  }
  const rel = ctx.mainRels.get(relId);
  if (!rel || rel.type !== CHART_RELATIONSHIP) {
    return null;
  }
  const partPath = resolvePartPath(ctx.mainPartPath, rel.target);
  const bytes = ctx.parts.get(partPath);
  if (!bytes) {
    return null;
  }
  const parsed = parseChartXml(new TextDecoder("utf-8").decode(bytes));
  const cNvPr = Array.from(frame.children).find((el) => el.localName === "cNvPr");
  const altText = cNvPr?.getAttribute("descr") || undefined;
  const chart: Chart = {
    type: "chart",
    chartType: parsed?.chartType ?? "columnClustered",
    categories: parsed?.categories ?? [],
    series: parsed?.series ?? [],
    options: parsed?.options ?? { legend: "none", dataLabels: false, gridlines: false },
    widthCm: emuToCm(extCx * transform.scaleX),
    heightCm: emuToCm(extCy * transform.scaleY),
    wrap: "square",
    offsetXCm: emuToCm((offX - transform.chOffX) * transform.scaleX),
    offsetYCm: emuToCm((offY - transform.chOffY) * transform.scaleY),
    sourcePart: partPath,
  };
  if (!parsed) {
    chart.unsupportedChartXml = new TextDecoder("utf-8").decode(bytes);
  }
  if (altText) {
    chart.altText = altText;
  }
  return chart;
}

export interface GroupShapeTextContext {
  parseParagraph: (pEl: Element) => Paragraph;
  /** Media-resolution context for `pic:pic` image members (SPEC-024/-057);
   * omit when the caller has no media pipeline (image members then seal). */
  mainRels?: Map<string, Relationship>;
  mainPartPath?: string;
  parts?: Map<string, Uint8Array>;
  /** Sink for the MediaPart of every parsed image member, so parseDocx adds
   * them to `DocumentModel.media` exactly as it does for standalone images. */
  collectMedia?: (part: MediaPart) => void;
}

interface ChildTransform {
  scaleX: number;
  scaleY: number;
  chOffX: number;
  chOffY: number;
}

/** Parses the members of a `wpg:wgp`/`wpg:grpSp` element, applying the child
 * coordinate transform. Returns null when any member is a kind this reader
 * does not model (so the caller seals the container). */
function parseMembers(wgp: Element, transform: ChildTransform, ctx: GroupShapeTextContext): DrawingMember[] | null {
  const members: DrawingMember[] = [];
  for (const child of Array.from(wgp.children)) {
    if (child.namespaceURI === WPS_NAMESPACE && child.localName === "wsp") {
      const shape = parseGroupMemberWsp(child, transform, { parseParagraph: ctx.parseParagraph });
      if (!shape) {
        return null;
      }
      members.push(shape);
    } else if (child.namespaceURI === PIC_NAMESPACE && child.localName === "pic") {
      if (!ctx.mainRels || ctx.mainPartPath === undefined || !ctx.parts) {
        return null;
      }
      const parsed = parseGroupMemberImage(child, transform, ctx.mainRels, ctx.mainPartPath, ctx.parts);
      if (!parsed) {
        return null;
      }
      ctx.collectMedia?.(parsed.media);
      members.push(parsed.image);
    } else if (child.namespaceURI === WPG_NAMESPACE && child.localName === "grpSp") {
      const nested = parseNestedGroup(child, transform, ctx);
      if (!nested) {
        return null;
      }
      members.push(nested);
    } else if (child.localName === "graphicFrame") {
      // A chart member (SPEC-057 AC-55): a `graphicFrame` holding a `c:chart`.
      const chart = parseGroupMemberChart(child, transform, ctx);
      if (!chart) {
        return null;
      }
      members.push(chart);
    } else if (
      (child.namespaceURI === WPG_NAMESPACE && (child.localName === "cNvGrpSpPr" || child.localName === "grpSpPr")) ||
      (child.namespaceURI === WPS_NAMESPACE && child.localName === "cNvGrpSpPr") ||
      (child.namespaceURI === WP_NAMESPACE && child.localName === "cNvPr") ||
      child.localName === "cNvPr"
    ) {
      // Non-member group properties — skip.
      continue;
    } else {
      // A picture, a chart graphicFrame, or anything else this reader does not
      // yet model: seal.
      return null;
    }
  }
  return members;
}

/** The child transform a group's `a:xfrm` defines: `s = ext/chExt`, with the
 * child-space origin `chOff`. */
function childTransformOf(xfrm: Element | undefined): ChildTransform {
  const ext = xfrm ? firstChildElement(xfrm, A_NAMESPACE, "ext") : undefined;
  const chExt = xfrm ? firstChildElement(xfrm, A_NAMESPACE, "chExt") : undefined;
  const chOff = xfrm ? firstChildElement(xfrm, A_NAMESPACE, "chOff") : undefined;
  const extCx = ext ? Number(ext.getAttribute("cx")) : NaN;
  const extCy = ext ? Number(ext.getAttribute("cy")) : NaN;
  const chExtCx = chExt ? Number(chExt.getAttribute("cx")) : NaN;
  const chExtCy = chExt ? Number(chExt.getAttribute("cy")) : NaN;
  const scaleX = Number.isFinite(extCx) && Number.isFinite(chExtCx) && chExtCx !== 0 ? extCx / chExtCx : 1;
  const scaleY = Number.isFinite(extCy) && Number.isFinite(chExtCy) && chExtCy !== 0 ? extCy / chExtCy : 1;
  const chOffX = chOff ? Number(chOff.getAttribute("x")) || 0 : 0;
  const chOffY = chOff ? Number(chOff.getAttribute("y")) || 0 : 0;
  return { scaleX, scaleY, chOffX, chOffY };
}

/** A nested `wpg:grpSp` member: its own `a:xfrm` positions it (in the parent's
 * child space, so the parent transform applies to its off/ext) and defines its
 * own child space for its members. */
function parseNestedGroup(grpSp: Element, parent: ChildTransform, ctx: GroupShapeTextContext): DrawingGroup | null {
  const grpSpPr = firstChildElement(grpSp, WPG_NAMESPACE, "grpSpPr");
  const xfrm = grpSpPr ? firstChildElement(grpSpPr, A_NAMESPACE, "xfrm") : undefined;
  const off = xfrm ? firstChildElement(xfrm, A_NAMESPACE, "off") : undefined;
  const ext = xfrm ? firstChildElement(xfrm, A_NAMESPACE, "ext") : undefined;
  const offX = off ? Number(off.getAttribute("x")) || 0 : 0;
  const offY = off ? Number(off.getAttribute("y")) || 0 : 0;
  const extCx = ext ? Number(ext.getAttribute("cx")) : NaN;
  const extCy = ext ? Number(ext.getAttribute("cy")) : NaN;
  if (!Number.isFinite(extCx) || !Number.isFinite(extCy)) {
    return null;
  }
  const rotAttr = xfrm?.getAttribute("rot");
  const rotationDeg = rotAttr ? Math.round(Number(rotAttr) / 60000) : 0;
  const inner = childTransformOf(xfrm);
  const members = parseMembers(grpSpPr ? grpSp : grpSp, inner, ctx);
  if (!members || members.length === 0) {
    return null;
  }
  const cNvPr = firstChildElement(grpSp, WPG_NAMESPACE, "cNvPr");
  const altText = cNvPr?.getAttribute("descr") || undefined;
  const group: DrawingGroup = {
    type: "drawingGroup",
    widthCm: emuToCm(extCx * parent.scaleX),
    heightCm: emuToCm(extCy * parent.scaleY),
    wrap: "square",
    offsetXCm: emuToCm((offX - parent.chOffX) * parent.scaleX),
    offsetYCm: emuToCm((offY - parent.chOffY) * parent.scaleY),
    members,
  };
  if (rotationDeg) {
    group.rotationDeg = ((rotationDeg % 360) + 360) % 360;
  }
  if (altText) {
    group.altText = altText;
  }
  return group;
}

/** Parses a `w:drawing` element into a DrawingGroup, or null when it is not a
 * `wpg:wgp` group. Falls to a sealed container (preservedSubtreeXml set) when
 * it holds a member kind this reader does not yet model. */
export function parseGroupDrawing(drawingEl: Element, ctx: GroupShapeTextContext): DrawingGroup | null {
  const containers = Array.from(drawingEl.children).filter((el) => el.namespaceURI === WP_NAMESPACE);
  if (containers.length !== 1) {
    return null;
  }
  const container = containers[0];
  const isAnchor = container.localName === "anchor";
  if (!isAnchor && container.localName !== "inline") {
    return null;
  }
  const graphic = firstChildElement(container, A_NAMESPACE, "graphic");
  const graphicData = graphic ? firstChildElement(graphic, A_NAMESPACE, "graphicData") : undefined;
  if (!graphicData || !(graphicData.getAttribute("uri") ?? "").includes(GROUP_URI)) {
    return null;
  }
  const wgp = firstChildElement(graphicData, WPG_NAMESPACE, "wgp");
  if (!wgp) {
    return null;
  }
  const grpSpPr = firstChildElement(wgp, WPG_NAMESPACE, "grpSpPr");
  const xfrm = grpSpPr ? firstChildElement(grpSpPr, A_NAMESPACE, "xfrm") : undefined;
  const ext = xfrm ? firstChildElement(xfrm, A_NAMESPACE, "ext") : undefined;
  const extCx = ext ? Number(ext.getAttribute("cx")) : NaN;
  const extCy = ext ? Number(ext.getAttribute("cy")) : NaN;
  if (!Number.isFinite(extCx) || !Number.isFinite(extCy)) {
    return null;
  }
  const rotAttr = xfrm?.getAttribute("rot");
  const rotationDeg = rotAttr ? Math.round(Number(rotAttr) / 60000) : 0;
  const transform = childTransformOf(xfrm);

  // Anchor-level fields (same rules as a single shape).
  const docPr = firstChildElement(container, WP_NAMESPACE, "docPr");
  const altText = docPr?.getAttribute("descr") || undefined;
  let wrap: DrawingGroup["wrap"] = "inline";
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
      wrap = "square";
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

  const members = parseMembers(wgp, transform, ctx);
  if (!members || members.length < 2) {
    // A member kind this reader does not yet model (a picture, a chart frame):
    // keep the drawing on SPEC-008 opaque preservation, exactly as before this
    // spec — preserved verbatim, the caller's opaque-inline fallback handles
    // it. (A childless or single-child group also yields no container.)
    return null;
  }
  const group: DrawingGroup = {
    type: "drawingGroup",
    widthCm: emuToCm(extCx),
    heightCm: emuToCm(extCy),
    wrap,
    members,
  };
  if (rotationDeg) {
    group.rotationDeg = ((rotationDeg % 360) + 360) % 360;
  }
  if (altText) {
    group.altText = altText;
  }
  if (offsetXCm !== undefined) {
    group.offsetXCm = offsetXCm;
  }
  if (offsetYCm !== undefined) {
    group.offsetYCm = offsetYCm;
  }
  if (wrap !== "inline") {
    group.relativeToH = "column";
    group.relativeToV = "paragraph";
  }
  if (zOrder !== undefined && zOrder !== 0) {
    group.zOrder = zOrder;
  }
  return group;
}
