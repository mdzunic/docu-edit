// SPEC-070: reads/writes a Connector as an ODF `draw:connector` — the
// connector-flavoured sibling of shape.ts's `draw:line`. Endpoints ride in
// `svg:x1/y1/x2/y2` (always present); attachment travels as
// `draw:start-shape`/`draw:end-shape` naming a target's `draw:id` plus a
// `draw:start-glue-point`/`draw:end-glue-point` site index. No identifier
// enters the model — the reader resolves each `draw:id` to the target's block
// index within the section, the writer mints ids per drawing. Arrowheads and
// the routing adjustment ride in our own `de:` extension namespace so they
// round-trip exactly (LibreOffice ignores the unknown attributes).
import type { Connector, ConnectorKind, Section } from "../documentModel";
import { isConnector } from "../documentModel";
import type { ArrowheadKind, ArrowheadSize } from "../core/shapes";
import { escapeXmlAttribute, firstChildElement, parseXml } from "../core/xml";
import { cmToOdfLength, odfLengthToCm } from "./units";
import { DRAW_NAMESPACE, STYLE_NAMESPACE, SVG_NAMESPACE, drawAttr, svgAttr, textAttr } from "../core/odfXml";
import {
  createOdfDrawingDefs,
  mergePreservedGraphicProps,
  outlineStyleAttrs,
  parseGraphicStyleProps,
  shadowStyleAttrs,
  wrapStyleAttrs,
  type OdfShapeStyleWriteContext,
} from "./shape";

const DE_NAMESPACE = "urn:docu-edit:shape";
const SHAPE_XMLNS = ` xmlns:draw="${DRAW_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}"`;

const KIND_TO_DRAW_TYPE: Record<ConnectorKind, string> = {
  straight: "line",
  elbow: "standard",
  curved: "curve",
};
const DRAW_TYPE_TO_KIND: Record<string, ConnectorKind> = {
  line: "straight",
  standard: "elbow",
  curve: "curved",
  curved: "curved",
  lines: "elbow",
};

/** Pending, not-yet-resolved attachment references, keyed by the parsed
 * Connector. Read out by resolveOdfConnectorAttachments once every drawing's
 * draw:id is known. */
const pendingRefs = new WeakMap<Connector, { start?: { id: string; site: number }; end?: { id: string; site: number } }>();

/** The draw:id a parsed drawing element carried, keyed by its block object —
 * stamped at the reader dispatch for connector resolution. */
const blockDrawId = new WeakMap<object, string>();

/** Records the `draw:id`/`xml:id` a drawing element was read under. Called by
 * reader.ts for every drawing kind. */
export function stampOdfConnectorTargetId(el: Element, block: object): void {
  const id = el.getAttributeNS(DRAW_NAMESPACE, "id") ?? el.getAttribute("draw:id") ?? el.getAttribute("xml:id");
  if (id) {
    blockDrawId.set(block, id);
  }
}

/** Parses a `draw:connector` into a Connector, registering its attachment
 * references for later resolution. Returns null for any other element. */
export function parseConnectorElement(
  el: Element,
  contentDoc: Document,
  stylesDoc: Document | undefined,
): { connector: Connector } | null {
  if (el.namespaceURI !== DRAW_NAMESPACE || el.localName !== "connector") {
    return null;
  }
  const x1 = odfLengthToCm(svgAttr(el, "x1")) ?? 0;
  const y1 = odfLengthToCm(svgAttr(el, "y1")) ?? 0;
  const x2 = odfLengthToCm(svgAttr(el, "x2")) ?? 0;
  const y2 = odfLengthToCm(svgAttr(el, "y2")) ?? 0;
  const offsetXCm = Math.min(x1, x2);
  const offsetYCm = Math.min(y1, y2);
  const widthCm = Math.abs(x2 - x1);
  const heightCm = Math.abs(y2 - y1);

  const drawType = drawAttr(el, "type") ?? "line";
  const kind = DRAW_TYPE_TO_KIND[drawType] ?? "straight";

  const anchorType = textAttr(el, "anchor-type");
  const wrap = anchorType === null || anchorType === "as-char" ? "inline" : "square";

  const connector: Connector = {
    type: "connector",
    kind,
    widthCm,
    heightCm,
    wrap,
    start: { xCm: x1 - offsetXCm, yCm: y1 - offsetYCm },
    end: { xCm: x2 - offsetXCm, yCm: y2 - offsetYCm },
  };
  if (wrap !== "inline") {
    connector.relativeToH = "column";
    connector.relativeToV = "paragraph";
    connector.offsetXCm = offsetXCm;
    connector.offsetYCm = offsetYCm;
  }

  const descEl = firstChildElement(el, SVG_NAMESPACE, "desc");
  const altText = descEl?.textContent || undefined;
  if (altText) {
    connector.altText = altText;
  }
  const groupOrdinalAttr = el.getAttributeNS(DE_NAMESPACE, "group");
  if (groupOrdinalAttr && Number.isFinite(Number(groupOrdinalAttr))) {
    connector.groupOrdinal = Number(groupOrdinalAttr);
  }
  const rotationAttr = el.getAttributeNS(DE_NAMESPACE, "rotation");
  if (rotationAttr && Number(rotationAttr)) {
    connector.rotationDeg = Number(rotationAttr);
  }
  const adjAttr = el.getAttributeNS(DE_NAMESPACE, "adjust");
  if (adjAttr && Number.isFinite(Number(adjAttr))) {
    connector.adjustments = [Number(adjAttr)];
  }
  const startArrow = el.getAttributeNS(DE_NAMESPACE, "start-arrow") as ArrowheadKind | null;
  const endArrow = el.getAttributeNS(DE_NAMESPACE, "end-arrow") as ArrowheadKind | null;
  if (startArrow) {
    connector.startArrow = startArrow;
  }
  if (endArrow) {
    connector.endArrow = endArrow;
  }
  // Arrowhead sizes (SPEC-073) ride in the same de: extension namespace as
  // the kinds; absent means unset (the renderer's "medium" default).
  const startArrowSize = el.getAttributeNS(DE_NAMESPACE, "start-arrow-size") as ArrowheadSize | null;
  if (startArrowSize === "small" || startArrowSize === "medium" || startArrowSize === "large") {
    connector.startArrowSize = startArrowSize;
  }
  const endArrowSize = el.getAttributeNS(DE_NAMESPACE, "end-arrow-size") as ArrowheadSize | null;
  if (endArrowSize === "small" || endArrowSize === "medium" || endArrowSize === "large") {
    connector.endArrowSize = endArrowSize;
  }
  // SPEC-073: outline (widened vocabulary), shadow and the verbatim
  // preserved leftovers, through the same shared parse shapes use. Markers
  // are "consume": a connector's style markers are regenerated from its own
  // de:start-arrow/de:end-arrow fields on every save. The parsed wrap is
  // ignored — a connector's wrap stays anchor-derived (SPEC-070).
  const style = parseGraphicStyleProps(el, contentDoc, stylesDoc, { markers: "consume", parseFill: false });
  if (style.outline) {
    connector.outline = style.outline;
  }
  if (style.shadow) {
    connector.shadow = style.shadow;
  }
  if (style.unsupportedStyleXml) {
    connector.unsupportedStyleXml = style.unsupportedStyleXml;
  }

  // Attachment references.
  const startShape = drawAttr(el, "start-shape");
  const endShape = drawAttr(el, "end-shape");
  const refs: { start?: { id: string; site: number }; end?: { id: string; site: number } } = {};
  if (startShape) {
    refs.start = { id: startShape, site: Number(drawAttr(el, "start-glue-point") ?? "8") };
  }
  if (endShape) {
    refs.end = { id: endShape, site: Number(drawAttr(el, "end-glue-point") ?? "8") };
  }
  if (refs.start || refs.end) {
    pendingRefs.set(connector, refs);
  }
  return { connector };
}

/** Resolves every connector's draw:start-shape/draw:end-shape id references to
 * the target's block index within the connector's section. Clears transient
 * state. Called by parseOdt after the whole model is assembled. */
export function resolveOdfConnectorAttachments(model: { sections: Section[] }): void {
  for (const section of model.sections) {
    const idToIndex = new Map<string, number>();
    section.blocks.forEach((block, index) => {
      const id = blockDrawId.get(block);
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
        const idx = idToIndex.get(refs.start.id);
        if (idx !== undefined) {
          block.startAttachment = { target: { block: idx }, site: refs.start.site };
        }
      }
      if (refs.end) {
        const idx = idToIndex.get(refs.end.id);
        if (idx !== undefined) {
          block.endAttachment = { target: { block: idx }, site: refs.end.site };
        }
      }
      pendingRefs.delete(block);
    }
  }
}

function parsePreservedConnectorStyle(xml: string | undefined): Element | undefined {
  if (!xml) {
    return undefined;
  }
  try {
    const el = parseXml(xml).documentElement;
    return el.namespaceURI === STYLE_NAMESPACE && el.localName === "graphic-properties" ? el : undefined;
  } catch {
    return undefined;
  }
}

function connectorGraphicPropertiesXml(connector: Connector, ctx: OdfShapeStyleWriteContext): string {
  const defs = ctx.defs ?? createOdfDrawingDefs();
  const strokeAttrs = connector.outline ? outlineStyleAttrs(connector.outline, defs) : ' draw:stroke="none"';
  const markerEnd = connector.endArrow && connector.endArrow !== "none" ? ' draw:marker-end="Arrow" draw:marker-end-width="0.3cm"' : "";
  const markerStart = connector.startArrow && connector.startArrow !== "none" ? ' draw:marker-start="Arrow" draw:marker-start-width="0.3cm"' : "";
  const shadowAttrs = connector.shadow ? shadowStyleAttrs(connector.shadow) : "";
  const deXmlns = shadowAttrs ? ` xmlns:de="${DE_NAMESPACE}"` : "";
  // topAndBottom is not a connector wrap (SPEC-070); the shared helper's
  // square/behind/infront branches are exactly the historic emission.
  const wrapAttrs = connector.wrap === "topAndBottom" ? "" : wrapStyleAttrs(connector.wrap);
  const base = `<style:graphic-properties${deXmlns}${strokeAttrs}${markerStart}${markerEnd}${shadowAttrs}${wrapAttrs} style:vertical-pos="from-top" style:vertical-rel="paragraph" style:horizontal-pos="from-left" style:horizontal-rel="paragraph"/>`;
  const preserved = ctx.includePreserved !== false ? parsePreservedConnectorStyle(connector.unsupportedStyleXml) : undefined;
  return preserved ? mergePreservedGraphicProps(base, preserved, defs) : base;
}

/** One generated graphic-family `style:style` for a connector. */
export function connectorGraphicStyleXml(name: string, connector: Connector, ctx: OdfShapeStyleWriteContext = {}): string {
  return `<style:style style:name="${escapeXmlAttribute(name)}"${SHAPE_XMLNS} style:family="graphic" style:parent-style-name="Graphics">${connectorGraphicPropertiesXml(connector, ctx)}</style:style>`;
}

/** How a connector's attachment targets resolve to their draw:ids. */
export interface ConnectorDrawIds {
  self?: string;
  start?: string;
  end?: string;
}

/** Serialises a Connector into a `draw:connector`. `ids.start`/`ids.end` are
 * the draw:ids of the attached shapes; `ids.self` this connector's own id. */
export function connectorElementXml(connector: Connector, styleName: string, ids: ConnectorDrawIds): string {
  const anchorType = connector.wrap === "inline" ? "as-char" : "paragraph";
  const originX = connector.offsetXCm ?? 0;
  const originY = connector.offsetYCm ?? 0;
  const x1 = originX + connector.start.xCm;
  const y1 = originY + connector.start.yCm;
  const x2 = originX + connector.end.xCm;
  const y2 = originY + connector.end.yCm;
  const descXml = connector.altText ? `<svg:desc>${escapeXmlAttribute(connector.altText)}</svg:desc>` : "";
  const idAttr = ids.self ? ` draw:id="${escapeXmlAttribute(ids.self)}" xml:id="${escapeXmlAttribute(ids.self)}"` : "";
  const groupAttr = connector.groupOrdinal !== undefined ? ` de:group="${connector.groupOrdinal}"` : "";
  const rotationAttr = connector.rotationDeg ? ` de:rotation="${connector.rotationDeg}"` : "";
  const adjustAttr = connector.adjustments?.[0] !== undefined ? ` de:adjust="${connector.adjustments[0]}"` : "";
  const startArrowAttr = connector.startArrow ? ` de:start-arrow="${connector.startArrow}"` : "";
  const endArrowAttr = connector.endArrow ? ` de:end-arrow="${connector.endArrow}"` : "";
  const startArrowSizeAttr = connector.startArrowSize ? ` de:start-arrow-size="${connector.startArrowSize}"` : "";
  const endArrowSizeAttr = connector.endArrowSize ? ` de:end-arrow-size="${connector.endArrowSize}"` : "";
  const startShapeAttr = connector.startAttachment && ids.start
    ? ` draw:start-shape="${escapeXmlAttribute(ids.start)}" draw:start-glue-point="${connector.startAttachment.site}"`
    : "";
  const endShapeAttr = connector.endAttachment && ids.end
    ? ` draw:end-shape="${escapeXmlAttribute(ids.end)}" draw:end-glue-point="${connector.endAttachment.site}"`
    : "";

  return (
    `<draw:connector${SHAPE_XMLNS} xmlns:xml="http://www.w3.org/XML/1998/namespace" text:anchor-type="${anchorType}"` +
    ` draw:style-name="${escapeXmlAttribute(styleName)}" draw:type="${KIND_TO_DRAW_TYPE[connector.kind]}"${idAttr}` +
    ` xmlns:de="${DE_NAMESPACE}"${groupAttr}${rotationAttr}${adjustAttr}${startArrowAttr}${endArrowAttr}${startArrowSizeAttr}${endArrowSizeAttr}${startShapeAttr}${endShapeAttr}` +
    ` svg:x1="${cmToOdfLength(x1)}" svg:y1="${cmToOdfLength(y1)}" svg:x2="${cmToOdfLength(x2)}" svg:y2="${cmToOdfLength(y2)}">` +
    descXml +
    "</draw:connector>"
  );
}

/** A stable draw:id for the drawing at (section, block) — deterministic so a
 * connector and its target agree without a shared counter. */
export function odfDrawId(section: number, block: number): string {
  return `de-drawing-${section}-${block}`;
}

/** Whether any block in the section is a connector (dispatch pre-check helper). */
export function sectionHasConnector(section: Section): boolean {
  return section.blocks.some(isConnector);
}
