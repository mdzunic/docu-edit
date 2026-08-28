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
import type { ArrowheadKind, ShapeOutline } from "../core/shapes";
import { escapeXmlAttribute, firstChildElement } from "../core/xml";
import { cmToOdfLength, odfLengthToCm } from "./units";
import { DRAW_NAMESPACE, STYLE_NAMESPACE, SVG_NAMESPACE, drawAttr, svgAttr, textAttr } from "../core/odfXml";

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

function resolveGraphicStroke(
  el: Element,
  contentDoc: Document,
  stylesDoc: Document | undefined,
): ShapeOutline | undefined {
  const styleName = drawAttr(el, "style-name");
  const find = (doc: Document | undefined): Element | undefined => {
    if (!doc || !styleName) {
      return undefined;
    }
    for (const styleEl of Array.from(doc.getElementsByTagNameNS(STYLE_NAMESPACE, "style"))) {
      if (styleEl.getAttributeNS(STYLE_NAMESPACE, "name") === styleName) {
        return firstChildElement(styleEl, STYLE_NAMESPACE, "graphic-properties");
      }
    }
    return undefined;
  };
  const props = find(contentDoc) ?? find(stylesDoc);
  if (!props) {
    return undefined;
  }
  const strokeAttr = drawAttr(props, "stroke");
  if (!strokeAttr || strokeAttr === "none") {
    return undefined;
  }
  const color = svgAttr(props, "stroke-color") ?? "#000000";
  const widthCm = odfLengthToCm(svgAttr(props, "stroke-width")) ?? 0.03;
  const widthPt = (widthCm / 2.54) * 72;
  const dash: ShapeOutline["dash"] = strokeAttr === "dash" ? "dash" : "solid";
  return { color: color.toLowerCase(), widthPt, dash };
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
  const outline = resolveGraphicStroke(el, contentDoc, stylesDoc);
  if (outline) {
    connector.outline = outline;
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

function connectorGraphicPropertiesXml(connector: Connector): string {
  let strokeAttrs = ' draw:stroke="none"';
  if (connector.outline) {
    const widthCm = (connector.outline.widthPt / 72) * 2.54;
    strokeAttrs =
      ` draw:stroke="${connector.outline.dash === "solid" ? "solid" : "dash"}"` +
      ` svg:stroke-color="${escapeXmlAttribute(connector.outline.color)}"` +
      ` svg:stroke-width="${cmToOdfLength(widthCm)}"`;
  }
  const markerEnd = connector.endArrow && connector.endArrow !== "none" ? ' draw:marker-end="Arrow" draw:marker-end-width="0.3cm"' : "";
  const markerStart = connector.startArrow && connector.startArrow !== "none" ? ' draw:marker-start="Arrow" draw:marker-start-width="0.3cm"' : "";
  let wrapAttrs = "";
  if (connector.wrap === "square") {
    wrapAttrs = ' style:wrap="parallel" style:number-wrapped-paragraphs="no-limit"';
  } else if (connector.wrap === "behind") {
    wrapAttrs = ' style:wrap="run-through" style:run-through="background"';
  } else if (connector.wrap === "infront") {
    wrapAttrs = ' style:wrap="run-through" style:run-through="foreground"';
  }
  return `<style:graphic-properties${strokeAttrs}${markerStart}${markerEnd}${wrapAttrs} style:vertical-pos="from-top" style:vertical-rel="paragraph" style:horizontal-pos="from-left" style:horizontal-rel="paragraph"/>`;
}

/** One generated graphic-family `style:style` for a connector. */
export function connectorGraphicStyleXml(name: string, connector: Connector): string {
  return `<style:style style:name="${escapeXmlAttribute(name)}"${SHAPE_XMLNS} style:family="graphic" style:parent-style-name="Graphics">${connectorGraphicPropertiesXml(connector)}</style:style>`;
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
  const startShapeAttr = connector.startAttachment && ids.start
    ? ` draw:start-shape="${escapeXmlAttribute(ids.start)}" draw:start-glue-point="${connector.startAttachment.site}"`
    : "";
  const endShapeAttr = connector.endAttachment && ids.end
    ? ` draw:end-shape="${escapeXmlAttribute(ids.end)}" draw:end-glue-point="${connector.endAttachment.site}"`
    : "";

  return (
    `<draw:connector${SHAPE_XMLNS} xmlns:xml="http://www.w3.org/XML/1998/namespace" text:anchor-type="${anchorType}"` +
    ` draw:style-name="${escapeXmlAttribute(styleName)}" draw:type="${KIND_TO_DRAW_TYPE[connector.kind]}"${idAttr}` +
    ` xmlns:de="${DE_NAMESPACE}"${groupAttr}${rotationAttr}${adjustAttr}${startArrowAttr}${endArrowAttr}${startShapeAttr}${endShapeAttr}` +
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
