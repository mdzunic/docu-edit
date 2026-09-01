// SPEC-070: reads/writes a Connector as a DrawingML `wps:cxnSp` inside a
// `wp:anchor`/`wp:inline` — the connector-flavoured sibling of shape.ts's
// `wps:wsp` path. Attachment travels as `a:stCxn`/`a:endCxn` `@id`/`@idx`
// referencing a `wp:docPr/@id`; no identifier enters the model — the reader
// resolves each id to the target's block index within the section
// ("Identity without ids"), the writer mints fresh ids per document. The
// exact endpoint coordinates ride in our own `de:` extension namespace so a
// free end round-trips precisely (Word reads the box corners, ignoring the
// unknown attributes).
import type { Block, Connector, ConnectorKind, Section } from "../documentModel";
import { isConnector } from "../documentModel";
import type { ArrowheadKind, ArrowheadSize, ShapeOutline } from "../core/shapes";
import { escapeXmlAttribute, firstChildElement } from "../core/xml";
import { cssColorToOoxml } from "../core/colors";
import { cmToEmu, emuToCm } from "../core/units";
import { outerShdwXml, parseShapeStyle, splitPreservedStyle, srgbClrXml } from "./shape";

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const DE_NAMESPACE = "urn:docu-edit:shape";

const EMU_PER_POINT = 12700;
const ptToEmu = (pt: number): number => Math.round(pt * EMU_PER_POINT);
const degToOoxmlRot = (deg: number): number => Math.round(deg * 60000);
const ooxmlRotToDeg = (rot: number): number => Math.round(rot / 60000);

const KIND_TO_PRST: Record<ConnectorKind, string> = {
  straight: "straightConnector1",
  elbow: "bentConnector3",
  curved: "curvedConnector3",
};
const PRST_TO_KIND: Record<string, ConnectorKind> = {
  straightConnector1: "straight",
  bentConnector3: "elbow",
  curvedConnector3: "curved",
};

/** ArrowheadKind ↔ OOXML `a:headEnd`/`a:tailEnd @type`. `open` maps to
 * DrawingML's `arrow`; the rest share their names. */
const ARROW_TO_OOXML: Record<ArrowheadKind, string> = {
  none: "none",
  triangle: "triangle",
  open: "arrow",
  stealth: "stealth",
  diamond: "diamond",
  oval: "oval",
};
const OOXML_TO_ARROW: Record<string, ArrowheadKind> = {
  none: "none",
  triangle: "triangle",
  arrow: "open",
  stealth: "stealth",
  diamond: "diamond",
  oval: "oval",
};

/** Pending, not-yet-resolved endpoint attachment references (the referenced
 * `wp:docPr/@id` and site index), keyed by the parsed Connector object. Read
 * out by resolveDocxConnectorAttachments once every block's own docPr id is
 * known. A WeakMap so the model itself never carries the transient id. */
const pendingRefs = new WeakMap<Connector, { start?: { docPrId: number; site: number }; end?: { docPrId: number; site: number } }>();

/** The docPr id stamped on a parsed drawing block during read — see
 * stampConnectorTargetId. Transient; cleared by resolveDocxConnectorAttachments. */
const blockDocPrId = new WeakMap<Block, number>();

/** Records the `wp:docPr/@id` a drawing block was read under, so a connector's
 * `a:stCxn`/`a:endCxn` id can later be resolved to that block's index. Called
 * by reader.ts at the drawing dispatch for every drawing kind. */
export function stampConnectorTargetId(drawingEl: Element, block: Block): void {
  const id = docPrIdOf(drawingEl);
  if (id !== undefined) {
    blockDocPrId.set(block, id);
  }
}

function docPrIdOf(drawingEl: Element): number | undefined {
  const container = Array.from(drawingEl.children).find((el) => el.namespaceURI === WP_NAMESPACE);
  const docPr = container ? firstChildElement(container, WP_NAMESPACE, "docPr") : undefined;
  const raw = docPr?.getAttribute("id");
  if (raw === null || raw === undefined) {
    return undefined;
  }
  const n = Number(raw);
  return Number.isFinite(n) ? n : undefined;
}

function parseArrow(ln: Element | undefined, localName: "headEnd" | "tailEnd"): ArrowheadKind | undefined {
  if (!ln) {
    return undefined;
  }
  const end = firstChildElement(ln, A_NAMESPACE, localName);
  const type = end?.getAttribute("type");
  if (type === null || type === undefined) {
    return undefined;
  }
  return OOXML_TO_ARROW[type] ?? "triangle";
}

/** The explicit arrow size (SPEC-073 AC-14) an end carries, or undefined
 * for a bare end — absent means "medium" everywhere downstream. */
function parseArrowSize(ln: Element | undefined, localName: "headEnd" | "tailEnd"): ArrowheadSize | undefined {
  const end = ln ? firstChildElement(ln, A_NAMESPACE, localName) : undefined;
  const len = end?.getAttribute("len") ?? end?.getAttribute("w");
  if (len === "sm") {
    return "small";
  }
  if (len === "lg") {
    return "large";
  }
  return len === "med" ? "medium" : undefined;
}

/** Parses a `w:drawing` holding a `wps:cxnSp` into a Connector, registering
 * its attachment references for later resolution. Returns null when the
 * drawing is not a single connector this spec models. */
export function parseConnectorDrawing(drawingEl: Element): { connector: Connector } | null {
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
  const cxnSp = graphicData ? firstChildElement(graphicData, WPS_NAMESPACE, "cxnSp") : undefined;
  if (!cxnSp) {
    return null;
  }
  const spPr = firstChildElement(cxnSp, WPS_NAMESPACE, "spPr");
  if (!spPr) {
    return null;
  }
  const prst = firstChildElement(spPr, A_NAMESPACE, "prstGeom")?.getAttribute("prst") ?? "";
  const kind = PRST_TO_KIND[prst] ?? "straight";

  const xfrm = firstChildElement(spPr, A_NAMESPACE, "xfrm");
  const rotAttr = xfrm?.getAttribute("rot");
  const rotationDeg = rotAttr ? ooxmlRotToDeg(Number(rotAttr)) : 0;

  const docPr = firstChildElement(container, WP_NAMESPACE, "docPr");
  const altText = docPr?.getAttribute("descr") || undefined;
  const groupOrdinalAttr = docPr?.getAttributeNS(DE_NAMESPACE, "group");
  const groupOrdinal = groupOrdinalAttr ? Number(groupOrdinalAttr) : undefined;

  const widthCm = emuToCm(cx);
  const heightCm = emuToCm(cy);

  // Exact endpoints from our de: extension, else the box corners (± flip).
  const deNum = (attr: string): number | undefined => {
    const v = docPr?.getAttributeNS(DE_NAMESPACE, attr);
    if (v === null || v === undefined) {
      return undefined;
    }
    const n = Number(v);
    return Number.isFinite(n) ? emuToCm(n) : undefined;
  };
  const flipH = xfrm?.getAttribute("flipH") === "1";
  const flipV = xfrm?.getAttribute("flipV") === "1";
  const start = {
    xCm: deNum("x1") ?? (flipH ? widthCm : 0),
    yCm: deNum("y1") ?? (flipV ? heightCm : 0),
  };
  const end = {
    xCm: deNum("x2") ?? (flipH ? 0 : widthCm),
    yCm: deNum("y2") ?? (flipV ? 0 : heightCm),
  };

  let wrap: Connector["wrap"] = "inline";
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
    const posH = firstChildElement(container, WP_NAMESPACE, "positionH");
    const offH = posH ? firstChildElement(posH, WP_NAMESPACE, "posOffset") : undefined;
    if (offH?.textContent) {
      offsetXCm = emuToCm(Number(offH.textContent));
    }
    const posV = firstChildElement(container, WP_NAMESPACE, "positionV");
    const offV = posV ? firstChildElement(posV, WP_NAMESPACE, "posOffset") : undefined;
    if (offV?.textContent) {
      offsetYCm = emuToCm(Number(offV.textContent));
    }
    const rh = container.getAttribute("relativeHeight");
    if (rh) {
      zOrder = Number(rh) - 251658240;
    }
  }

  const connector: Connector = {
    type: "connector",
    kind,
    widthCm,
    heightCm,
    wrap,
    start,
    end,
  };
  if (rotationDeg) {
    connector.rotationDeg = rotationDeg;
  }
  if (altText) {
    connector.altText = altText;
  }
  if (groupOrdinal !== undefined && Number.isFinite(groupOrdinal)) {
    connector.groupOrdinal = groupOrdinal;
  }
  if (offsetXCm !== undefined) {
    connector.offsetXCm = offsetXCm;
  }
  if (offsetYCm !== undefined) {
    connector.offsetYCm = offsetYCm;
  }
  if (wrap !== "inline") {
    connector.relativeToH = "column";
    connector.relativeToV = "paragraph";
  }
  if (zOrder !== undefined) {
    connector.zOrder = zOrder;
  }
  // Adjustment (bent/curved): the single adj1 guide, as a 0-1 fraction.
  const avLst = firstChildElement(firstChildElement(spPr, A_NAMESPACE, "prstGeom") ?? spPr, A_NAMESPACE, "avLst");
  const gd = avLst ? firstChildElement(avLst, A_NAMESPACE, "gd") : undefined;
  const gdVal = gd?.getAttribute("fmla");
  const gdMatch = gdVal ? /^val\s+(-?\d+)$/.exec(gdVal.trim()) : null;
  if (gdMatch) {
    connector.adjustments = [Number(gdMatch[1]) / 100000];
  }
  // SPEC-073: the widened outline (alpha/join/cap/compound/dash), shadow and
  // unmodelled-property preservation share the shape styling parser. A
  // connector's arrowheads stay on its own startArrow/endArrow fields
  // (SPEC-070), so any arrows the outline parser captured are lifted off.
  const style = parseShapeStyle(spPr);
  if (style.outline) {
    delete style.outline.startArrow;
    delete style.outline.endArrow;
    connector.outline = style.outline;
  }
  if (style.shadow) {
    connector.shadow = style.shadow;
  }
  if (style.preserved.length > 0) {
    connector.unsupportedStyleXml = style.preserved.join("");
  }
  const ln = firstChildElement(spPr, A_NAMESPACE, "ln");
  const startArrow = parseArrow(ln, "headEnd");
  const endArrow = parseArrow(ln, "tailEnd");
  if (startArrow) {
    connector.startArrow = startArrow;
  }
  if (endArrow) {
    connector.endArrow = endArrow;
  }
  const startSize = parseArrowSize(ln, "headEnd");
  if (startSize) {
    connector.startArrowSize = startSize;
  }
  const endSize = parseArrowSize(ln, "tailEnd");
  if (endSize) {
    connector.endArrowSize = endSize;
  }

  // Attachment references (a:stCxn/a:endCxn on wps:cNvCnPr).
  const cNvCnPr = firstChildElement(cxnSp, WPS_NAMESPACE, "cNvCnPr");
  const refs: { start?: { docPrId: number; site: number }; end?: { docPrId: number; site: number } } = {};
  if (cNvCnPr) {
    const st = firstChildElement(cNvCnPr, A_NAMESPACE, "stCxn");
    const en = firstChildElement(cNvCnPr, A_NAMESPACE, "endCxn");
    if (st) {
      refs.start = { docPrId: Number(st.getAttribute("id")), site: Number(st.getAttribute("idx")) };
    }
    if (en) {
      refs.end = { docPrId: Number(en.getAttribute("id")), site: Number(en.getAttribute("idx")) };
    }
  }
  if (refs.start || refs.end) {
    pendingRefs.set(connector, refs);
  }
  return { connector };
}

/** Resolves every connector's pending `a:stCxn`/`a:endCxn` id references to
 * the block index of the target within the connector's section, using the
 * docPr ids stamped during read. Clears all transient state. Called once by
 * parseDocx after the whole model is assembled. */
export function resolveDocxConnectorAttachments(model: { sections: Section[] }): void {
  for (const section of model.sections) {
    const idToIndex = new Map<number, number>();
    section.blocks.forEach((block, index) => {
      const id = blockDocPrId.get(block);
      if (id !== undefined) {
        idToIndex.set(id, index);
      }
    });
    for (const block of section.blocks) {
      if (!isConnector(block)) {
        continue;
      }
      const refs = pendingRefs.get(block);
      if (!refs) {
        continue;
      }
      if (refs.start) {
        const idx = idToIndex.get(refs.start.docPrId);
        if (idx !== undefined) {
          block.startAttachment = { target: { block: idx }, site: refs.start.site };
        }
      }
      if (refs.end) {
        const idx = idToIndex.get(refs.end.docPrId);
        if (idx !== undefined) {
          block.endAttachment = { target: { block: idx }, site: refs.end.site };
        }
      }
      pendingRefs.delete(block);
    }
  }
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

function arrowXml(kind: ArrowheadKind | undefined, localName: "headEnd" | "tailEnd", size?: ArrowheadSize): string {
  if (kind === undefined) {
    return "";
  }
  if (size === undefined) {
    // Bare form, exactly as SPEC-070 shipped — byte-stable for connectors
    // that never chose a size.
    return `<a:${localName} type="${ARROW_TO_OOXML[kind]}"/>`;
  }
  const sz = size === "small" ? "sm" : size === "large" ? "lg" : "med";
  return `<a:${localName} type="${ARROW_TO_OOXML[kind]}" w="${sz}" len="${sz}"/>`;
}

function connectorOutlineXml(connector: Connector): string {
  const outline = connector.outline;
  const arrows =
    arrowXml(connector.startArrow, "headEnd", connector.startArrowSize) +
    arrowXml(connector.endArrow, "tailEnd", connector.endArrowSize);
  if (!outline) {
    return `<a:ln>${arrows || "<a:noFill/>"}</a:ln>`;
  }
  const color = cssColorToOoxml(outline.color);
  if (!color) {
    return `<a:ln>${arrows || "<a:noFill/>"}</a:ln>`;
  }
  const capAttr = outline.cap === "round" ? ' cap="rnd"' : outline.cap === "square" ? ' cap="sq"' : "";
  const cmpd =
    outline.compound === "double"
      ? "dbl"
      : outline.compound === "thickThin"
        ? "thickThin"
        : outline.compound === "thinThick"
          ? "thinThick"
          : "";
  const joinXml = outline.join === "round" ? "<a:round/>" : outline.join === "bevel" ? "<a:bevel/>" : "";
  return (
    `<a:ln w="${ptToEmu(outline.widthPt)}"${capAttr}${cmpd ? ` cmpd="${cmpd}"` : ""}>` +
    `<a:solidFill>${srgbClrXml(outline.color, outline.alpha)}</a:solidFill>` +
    dashXml(outline.dash) +
    joinXml +
    arrows +
    "</a:ln>"
  );
}

/** How the writer resolves a connector's attachment target to the docPr id it
 * will assign that target — see planConnectorTargetIds in writer.ts. */
export interface ConnectorTargetIds {
  start?: number;
  end?: number;
}

/** Serialises a Connector + its assigned docPrId into a `w:drawing` holding a
 * `wps:cxnSp`. `targetIds` gives the docPr ids of the attached shapes so
 * `a:stCxn`/`a:endCxn` can name them. */
export function connectorDrawingXml(connector: Connector, docPrId: number, targetIds: ConnectorTargetIds): string {
  const cx = cmToEmu(connector.widthCm);
  const cy = cmToEmu(Math.max(0, connector.heightCm));
  const id = docPrId;
  const name = `Connector ${id}`;
  const descrAttr = connector.altText ? ` descr="${escapeXmlAttribute(connector.altText)}"` : "";
  const groupAttr = connector.groupOrdinal !== undefined ? ` de:group="${connector.groupOrdinal}"` : "";
  const x1 = cmToEmu(connector.start.xCm);
  const y1 = cmToEmu(connector.start.yCm);
  const x2 = cmToEmu(connector.end.xCm);
  const y2 = cmToEmu(connector.end.yCm);
  const docPr =
    `<wp:docPr id="${id}" name="${escapeXmlAttribute(name)}"${descrAttr}` +
    ` xmlns:de="${DE_NAMESPACE}"${groupAttr}` +
    ` de:x1="${x1}" de:y1="${y1}" de:x2="${x2}" de:y2="${y2}"/>`;

  const flipH = connector.start.xCm > connector.end.xCm ? ' flipH="1"' : "";
  const flipV = connector.start.yCm > connector.end.yCm ? ' flipV="1"' : "";
  const rotAttr = connector.rotationDeg ? ` rot="${degToOoxmlRot(connector.rotationDeg)}"` : "";
  const prst = KIND_TO_PRST[connector.kind];
  const adj = connector.kind !== "straight" && connector.adjustments?.[0] !== undefined
    ? `<a:avLst><a:gd name="adj1" fmla="val ${Math.round(connector.adjustments[0] * 100000)}"/></a:avLst>`
    : "<a:avLst/>";

  const stCxn = connector.startAttachment && targetIds.start !== undefined
    ? `<a:stCxn id="${targetIds.start}" idx="${connector.startAttachment.site}"/>`
    : "";
  const endCxn = connector.endAttachment && targetIds.end !== undefined
    ? `<a:endCxn id="${targetIds.end}" idx="${connector.endAttachment.site}"/>`
    : "";

  const preservedStyle = connector.unsupportedStyleXml
    ? splitPreservedStyle(connector.unsupportedStyleXml)
    : { fill: [] as string[], ln: [] as string[], effects: [] as string[], other: [] as string[] };
  const cxnSp =
    `<wps:cxnSp xmlns:wps="${WPS_NAMESPACE}">` +
    `<wps:cNvCnPr>${stCxn}${endCxn}</wps:cNvCnPr>` +
    "<wps:spPr>" +
    `<a:xfrm${rotAttr}${flipH}${flipV}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${prst}">${adj}</a:prstGeom>` +
    preservedStyle.fill.join("") +
    (preservedStyle.ln.length > 0 ? preservedStyle.ln.join("") : connectorOutlineXml(connector)) +
    (connector.shadow || preservedStyle.effects.length > 0
      ? `<a:effectLst>${connector.shadow ? outerShdwXml(connector.shadow) : ""}${preservedStyle.effects.join("")}</a:effectLst>`
      : "") +
    preservedStyle.other.join("") +
    "</wps:spPr>" +
    "</wps:cxnSp>";

  const graphic =
    `<a:graphic xmlns:a="${A_NAMESPACE}">` +
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    cxnSp +
    "</a:graphicData>" +
    "</a:graphic>";

  if (connector.wrap === "inline") {
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

  const behindDoc = connector.wrap === "behind" ? "1" : "0";
  let wrapXml: string;
  if (connector.wrap === "square") {
    wrapXml = '<wp:wrapSquare wrapText="bothSides"/>';
  } else if (connector.wrap === "topAndBottom") {
    wrapXml = "<wp:wrapTopAndBottom/>";
  } else {
    wrapXml = "<wp:wrapNone/>";
  }
  const offsetXEmu = cmToEmu(connector.offsetXCm ?? 0);
  const offsetYEmu = cmToEmu(connector.offsetYCm ?? 0);
  const relativeHeight = 251658240 + (connector.zOrder ?? 0);
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

/** A connector emitted at its own block slot (non-absorbed) needs its own
 * `w:p` wrapper, like shapeParagraphXml. */
export function connectorParagraphXml(connector: Connector, docPrId: number, targetIds: ConnectorTargetIds): string {
  return `<w:p><w:r>${connectorDrawingXml(connector, docPrId, targetIds)}</w:r></w:p>`;
}

/** Resolves the docPr ids of a connector's attachment targets from the map
 * built by planConnectorTargetIds (`${section}:${block}` → docPrId). */
export function connectorTargetIds(
  section: number,
  connector: Connector,
  docPrIdByBlock: Map<string, number>,
): ConnectorTargetIds {
  const ids: ConnectorTargetIds = {};
  if (connector.startAttachment) {
    ids.start = docPrIdByBlock.get(`${section}:${connector.startAttachment.target.block}`);
  }
  if (connector.endAttachment) {
    ids.end = docPrIdByBlock.get(`${section}:${connector.endAttachment.target.block}`);
  }
  return ids;
}

/** Whether a section's block, when it is a connector, is what a given drawing
 * element is (used by callers testing dispatch). */
export function sectionHasConnector(section: Section): boolean {
  return section.blocks.some(isConnector);
}
