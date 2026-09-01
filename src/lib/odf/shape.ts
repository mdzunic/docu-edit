// Reads/writes the ODF shape markup SPEC-051 models: `draw:custom-shape`
// (area geometries) and `draw:line` (Line/Arrow), the shape-flavoured
// siblings of drawing.ts's `draw:frame`/`draw:image` for pictures. Mirrors
// drawing.ts's own split (parse.../...Xml) so the two modules read side by
// side; this one never touches Image, drawing.ts never touches Shape.
import type { Block, ImageWrapMode, Paragraph, Shape } from "../documentModel";
import type { ShapeFill, ShapeGeometry, ShapeOutline } from "../core/shapes";
import { LINE_GEOMETRIES, shapeFillBaseColor } from "../core/shapes";
import { adjustmentToGuideValue, guideValueToAdjustment } from "../core/shapeCatalogue";
import {
  GEOMETRY_TO_DRAW_TYPE,
  DRAW_TYPE_TO_GEOMETRY as DRAW_TYPE_TO_GEOMETRY_TABLE,
} from "../core/drawingMl";
import { enhancedPathString, enhancedPathViewBox, parseEnhancedPath } from "../core/enhancedPath";
import { escapeXmlAttribute, firstChildElement } from "../core/xml";
import { cmToOdfLength, odfLengthToCm } from "./units";
import {
  DRAW_NAMESPACE,
  STYLE_NAMESPACE,
  SVG_NAMESPACE,
  TEXT_NAMESPACE,
  drawAttr,
  styleAttr,
  svgAttr,
  textAttr,
} from "../core/odfXml";

// Our own extension namespace, used to round-trip groupOrdinal and rotation
// (ODF's real `draw:transform` is a general affine-matrix string; encoding
// our whole-degree rotation through it is unnecessary complexity when we
// only ever need to round-trip through our own reader/writer — see the
// implement summary). Unknown-namespaced attributes are ignored by
// LibreOffice/Word, so this never corrupts a file another app opens.
const DE_NAMESPACE = "urn:docu-edit:shape";

// Declared locally on every shape-related element rather than on the
// document root, so content.xml's bytes are unchanged for documents without
// shapes (the TABLE_XMLNS/FRAME_XMLNS pattern from writer.ts/drawing.ts).
// Needed on both the style:style (draw:fill/stroke, svg:stroke-*) and the
// draw:custom-shape/draw:line element itself (svg:width/height/x/y, etc).
const SHAPE_XMLNS = ` xmlns:draw="${DRAW_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}"`;

// The ODF draw:type maps now live in core/drawingMl.ts (SPEC-079's two-phase
// move) so the `.odp` reader shares them; this module re-imports them (the
// reverse map aliased to its historical name), its behaviour unchanged.

/** Everything needed to parse a shape's text:p children into
 * `Shape.text.blocks` — injected by the caller (reader.ts) so this module
 * stays independent of the paragraph parser's own dependency graph. */
export interface OdfShapeTextContext {
  parseParagraph: (pEl: Element) => Paragraph;
}

function vAlignOf(el: Element): "top" | "middle" | "bottom" {
  const anchor = drawAttr(el, "text-vertical-align") ?? styleAttr(el, "vertical-align");
  if (anchor === "middle" || anchor === "center") {
    return "middle";
  }
  if (anchor === "bottom") {
    return "bottom";
  }
  return "top";
}

function parseShapeText(shapeEl: Element, ctx: OdfShapeTextContext): Shape["text"] | undefined {
  const textBox = firstChildElement(shapeEl, DRAW_NAMESPACE, "text-box");
  const container = textBox ?? shapeEl;
  const paragraphs = Array.from(container.children).filter(
    (child) => child.namespaceURI === TEXT_NAMESPACE && (child.localName === "p" || child.localName === "h"),
  );
  if (paragraphs.length === 0) {
    return undefined;
  }
  const blocks: Block[] = paragraphs.map((p) => ctx.parseParagraph(p));
  return { blocks, vAlign: vAlignOf(textBox ?? shapeEl) };
}

function resolveGraphicStyleProps(
  el: Element,
  contentDoc: Document,
  stylesDoc: Document | undefined,
): { fill?: ShapeFill; outline?: ShapeOutline; wrap: ImageWrapMode } {
  const styleName = drawAttr(el, "style-name");
  let props: Element | undefined;
  const find = (doc: Document | undefined) => {
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
  props = find(contentDoc) ?? find(stylesDoc);

  let fill: ShapeFill | undefined;
  let outline: ShapeOutline | undefined;
  let wrap: ImageWrapMode = "inline";
  if (props) {
    // draw:fill/draw:fill-color/draw:stroke live in the draw: namespace;
    // stroke colour/width reuse SVG's own attribute names (svg: namespace,
    // ODF's convention for stroke details) — matches graphicPropertiesXml's
    // own emission below, which is what a same-format save must read back.
    const fillAttr = drawAttr(props, "fill");
    if (fillAttr === "solid") {
      const color = drawAttr(props, "fill-color");
      if (color) {
        fill = { kind: "solid", color: color.toLowerCase() };
      }
    }
    const strokeAttr = drawAttr(props, "stroke");
    if (strokeAttr && strokeAttr !== "none") {
      const color = svgAttr(props, "stroke-color") ?? "#000000";
      const widthCm = odfLengthToCm(svgAttr(props, "stroke-width")) ?? 0.03;
      const widthPt = (widthCm / 2.54) * 72;
      const dash: ShapeOutline["dash"] = strokeAttr === "dash" ? "dash" : "solid";
      outline = { color: color.toLowerCase(), widthPt, dash };
    }
    const wrapAttr = styleAttr(props, "wrap");
    const runThrough = styleAttr(props, "run-through");
    if (wrapAttr === "parallel" || wrapAttr === "dynamic") {
      wrap = "square";
    } else if (wrapAttr === "run-through" && runThrough === "background") {
      wrap = "behind";
    } else if (wrapAttr === "run-through") {
      wrap = "infront";
    } else if (wrapAttr === "none") {
      wrap = "topAndBottom";
    }
  }
  const anchorType = textAttr(el, "anchor-type");
  if (anchorType === null || anchorType === "as-char") {
    wrap = "inline";
  }
  return { fill, outline, wrap };
}

export interface ParsedShape {
  shape: Shape;
}

/** Parses a `draw:custom-shape` or `draw:line` element into a Shape, or
 * null when its geometry/type is outside the catalogue this spec models —
 * the caller's SPEC-008 opaque fallback then preserves it verbatim. */
export function parseShapeElement(
  el: Element,
  contentDoc: Document,
  stylesDoc: Document | undefined,
  ctx: OdfShapeTextContext,
): ParsedShape | null {
  const isLine = el.namespaceURI === DRAW_NAMESPACE && el.localName === "line";
  const isCustom = el.namespaceURI === DRAW_NAMESPACE && el.localName === "custom-shape";
  if (!isLine && !isCustom) {
    return null;
  }

  const { fill, outline, wrap } = resolveGraphicStyleProps(el, contentDoc, stylesDoc);
  const descEl = firstChildElement(el, SVG_NAMESPACE, "desc");
  const altText = descEl?.textContent || undefined;
  const groupOrdinalAttr = el.getAttributeNS(DE_NAMESPACE, "group");
  const groupOrdinal = groupOrdinalAttr ? Number(groupOrdinalAttr) : undefined;
  const rotationAttr = el.getAttributeNS(DE_NAMESPACE, "rotation");
  const rotationDeg = rotationAttr ? Number(rotationAttr) : undefined;

  let widthCm: number;
  let heightCm: number;
  let geometry: ShapeGeometry;
  let unsupportedGeometryXml: string | undefined;
  let path: Shape["path"];
  let pathSource: Shape["pathSource"];
  let adjustments: number[] = [];

  let lineOffsetXCm: number | undefined;
  let lineOffsetYCm: number | undefined;
  if (isLine) {
    const x1 = odfLengthToCm(svgAttr(el, "x1")) ?? 0;
    const y1 = odfLengthToCm(svgAttr(el, "y1")) ?? 0;
    const x2 = odfLengthToCm(svgAttr(el, "x2")) ?? 0;
    const y2 = odfLengthToCm(svgAttr(el, "y2")) ?? 0;
    widthCm = Math.abs(x2 - x1);
    heightCm = Math.abs(y2 - y1);
    // shapeElementXml always writes x1/y1 as the box origin (x2/y2 = origin
    // + width/height) — a draw:line has no separate svg:x/svg:y, unlike
    // draw:custom-shape below.
    lineOffsetXCm = x1;
    lineOffsetYCm = y1;
    geometry = drawAttr(el, "marker-end") ? "arrow" : "line";
  } else {
    const w = odfLengthToCm(svgAttr(el, "width"));
    const h = odfLengthToCm(svgAttr(el, "height"));
    if (w === undefined || h === undefined) {
      return null;
    }
    widthCm = w;
    heightCm = h;
    const enhancedGeometry = firstChildElement(el, DRAW_NAMESPACE, "enhanced-geometry");
    const drawType = enhancedGeometry ? drawAttr(enhancedGeometry, "type") : null;
    const mapped = drawType ? DRAW_TYPE_TO_GEOMETRY_TABLE[drawType] : undefined;
    geometry = mapped ?? "rect";
    if (!mapped) {
      // A non-primitive enhanced-geometry with an @draw:enhanced-path is a
      // path shape (SPEC-071 AC-5) — parse it and render its real outline,
      // unless it is formula-driven or uses an unmodelled command, in which
      // case it stays on SPEC-051's verbatim preservation path.
      const enhancedPath = enhancedGeometry ? drawAttr(enhancedGeometry, "enhanced-path") : null;
      const viewBox = enhancedGeometry ? svgAttr(enhancedGeometry, "viewBox") : null;
      const vb = viewBox ? viewBox.trim().split(/\s+/).map(Number) : null;
      const parsed =
        enhancedPath && vb && vb.length === 4 && Number.isFinite(vb[2]) && Number.isFinite(vb[3])
          ? parseEnhancedPath(enhancedPath, vb[2] - vb[0], vb[3] - vb[1])
          : null;
      if (parsed) {
        path = parsed.path;
        const smoothAttr = enhancedGeometry ? enhancedGeometry.getAttributeNS(DE_NAMESPACE, "smooth") : null;
        if (smoothAttr) {
          path.smooth = smoothAttr.split("").map((c) => c === "1");
        }
        if (parsed.converted && enhancedGeometry) {
          pathSource = { format: "odt" as const, xml: new XMLSerializer().serializeToString(enhancedGeometry) };
        }
      } else {
        unsupportedGeometryXml = enhancedGeometry ? new XMLSerializer().serializeToString(enhancedGeometry) : "";
      }
    } else if (enhancedGeometry) {
      // draw:modifiers is a space-separated list of adjustment values in the
      // shape's coordinate space, normalised to fractions (SPEC-069 AC-6b).
      const modifiers = drawAttr(enhancedGeometry, "modifiers");
      if (modifiers) {
        adjustments = modifiers
          .trim()
          .split(/\s+/)
          .map(Number)
          .filter((n) => !Number.isNaN(n))
          .map(guideValueToAdjustment);
      }
    }
  }

  const shape: Shape = { type: "shape", geometry, widthCm, heightCm, wrap };
  if (path) {
    shape.path = path;
    if (pathSource) {
      shape.pathSource = pathSource;
    }
  } else if (unsupportedGeometryXml !== undefined) {
    shape.unsupportedGeometryXml = unsupportedGeometryXml;
  }
  if (adjustments.length > 0) {
    shape.adjustments = adjustments;
  }
  if (fill) {
    shape.fill = fill;
  }
  if (outline) {
    shape.outline = outline;
  }
  if (altText) {
    shape.altText = altText;
  }
  if (rotationDeg) {
    shape.rotationDeg = rotationDeg;
  }
  if (groupOrdinal !== undefined && Number.isFinite(groupOrdinal)) {
    shape.groupOrdinal = groupOrdinal;
  }
  if (wrap !== "inline") {
    shape.relativeToH = "column";
    shape.relativeToV = "paragraph";
    const offsetXCm = isLine ? lineOffsetXCm : odfLengthToCm(svgAttr(el, "x"));
    const offsetYCm = isLine ? lineOffsetYCm : odfLengthToCm(svgAttr(el, "y"));
    if (offsetXCm !== undefined) {
      shape.offsetXCm = offsetXCm;
    }
    if (offsetYCm !== undefined) {
      shape.offsetYCm = offsetYCm;
    }
  }
  if (!isLine) {
    const text = parseShapeText(el, ctx);
    if (text) {
      shape.text = text;
    }
  }
  return { shape };
}

function graphicPropertiesXml(shape: Shape): string {
  const fillColor = shapeFillBaseColor(shape.fill);
  const fillAttrs =
    shape.fill && shape.fill.kind !== "none" && fillColor
      ? ` draw:fill="solid" draw:fill-color="${escapeXmlAttribute(fillColor)}"`
      : ' draw:fill="none"';
  let strokeAttrs = ' draw:stroke="none"';
  if (shape.outline) {
    const widthCm = (shape.outline.widthPt / 72) * 2.54;
    strokeAttrs =
      ` draw:stroke="${shape.outline.dash === "solid" ? "solid" : "dash"}"` +
      ` svg:stroke-color="${escapeXmlAttribute(shape.outline.color)}"` +
      ` svg:stroke-width="${cmToOdfLength(widthCm)}"`;
  }
  let wrapAttrs = "";
  if (shape.wrap === "square") {
    wrapAttrs = ' style:wrap="parallel" style:number-wrapped-paragraphs="no-limit"';
  } else if (shape.wrap === "topAndBottom") {
    wrapAttrs = ' style:wrap="none"';
  } else if (shape.wrap === "behind") {
    wrapAttrs = ' style:wrap="run-through" style:run-through="background"';
  } else if (shape.wrap === "infront") {
    wrapAttrs = ' style:wrap="run-through" style:run-through="foreground"';
  }
  return `<style:graphic-properties${fillAttrs}${strokeAttrs}${wrapAttrs} style:vertical-pos="from-top" style:vertical-rel="paragraph" style:horizontal-pos="from-left" style:horizontal-rel="paragraph"/>`;
}

/** One generated graphic-family `style:style` for a shape — each shape gets
 * its own automatic style (no cross-shape dedup, unlike AutomaticStyles'
 * cached per-wrap-mode image styles) since fill/outline/wrap vary per shape. */
export function shapeGraphicStyleXml(name: string, shape: Shape): string {
  return `<style:style style:name="${escapeXmlAttribute(name)}"${SHAPE_XMLNS} style:family="graphic" style:parent-style-name="Graphics">${graphicPropertiesXml(shape)}</style:style>`;
}

export interface OdfShapeTextWriteContext {
  paragraphXml: (paragraph: Paragraph) => string;
}

function textBoxXml(shape: Shape, ctx: OdfShapeTextWriteContext): string {
  if (!shape.text) {
    return "";
  }
  const paragraphs = shape.text.blocks
    .map((block) => (block.type === "paragraph" ? ctx.paragraphXml(block) : ""))
    .join("");
  const vAlign =
    shape.text.vAlign === "middle" ? "middle" : shape.text.vAlign === "bottom" ? "bottom" : "top";
  return `<draw:text-box draw:text-vertical-align="${vAlign}">${paragraphs || "<text:p/>"}</draw:text-box>`;
}

/** Serialises a Shape into `draw:custom-shape` (area geometries) or
 * `draw:line` (Line/Arrow) — the shape-flavoured sibling of drawing.ts's
 * drawFrameXml. `styleName` is this shape's own generated graphic style
 * (see shapeGraphicStyleXml). */
export function shapeElementXml(shape: Shape, styleName: string, ctx: OdfShapeTextWriteContext, drawId?: string): string {
  const anchorType = shape.wrap === "inline" ? "as-char" : "paragraph";
  const descXml = shape.altText ? `<svg:desc>${escapeXmlAttribute(shape.altText)}</svg:desc>` : "";
  const groupAttr = shape.groupOrdinal !== undefined ? ` de:group="${shape.groupOrdinal}"` : "";
  const rotationAttr = shape.rotationDeg ? ` de:rotation="${shape.rotationDeg}"` : "";
  const deXmlns = groupAttr || rotationAttr ? ` xmlns:de="${DE_NAMESPACE}"` : "";
  // SPEC-070: a shape referenced by a connector's draw:start-shape/end-shape
  // carries the matching draw:id (also as xml:id, which ODF requires as the
  // canonical id) so the reference resolves on reopen.
  const idAttr = drawId
    ? ` xmlns:xml="http://www.w3.org/XML/1998/namespace" draw:id="${escapeXmlAttribute(drawId)}" xml:id="${escapeXmlAttribute(drawId)}"`
    : "";
  const xCm = shape.wrap !== "inline" ? shape.offsetXCm ?? 0 : undefined;
  const yCm = shape.wrap !== "inline" ? shape.offsetYCm ?? 0 : undefined;
  const posAttrs = xCm !== undefined && yCm !== undefined ? ` svg:x="${cmToOdfLength(xCm)}" svg:y="${cmToOdfLength(yCm)}"` : "";

  if (LINE_GEOMETRIES.has(shape.geometry) && !shape.unsupportedGeometryXml) {
    const markerAttr = shape.geometry === "arrow" ? ' draw:marker-end="Arrow" draw:marker-end-width="0.3cm"' : "";
    return (
      `<draw:line${SHAPE_XMLNS} text:anchor-type="${anchorType}" draw:style-name="${escapeXmlAttribute(styleName)}"${idAttr}${deXmlns}${groupAttr}${rotationAttr}` +
      ` svg:x1="${cmToOdfLength(xCm ?? 0)}" svg:y1="${cmToOdfLength(yCm ?? 0)}"` +
      ` svg:x2="${cmToOdfLength((xCm ?? 0) + shape.widthCm)}" svg:y2="${cmToOdfLength((yCm ?? 0) + shape.heightCm)}"` +
      `${markerAttr}>` +
      descXml +
      "</draw:line>"
    );
  }

  const modifiersAttr =
    shape.adjustments && shape.adjustments.length > 0
      ? ` draw:modifiers="${shape.adjustments.map(adjustmentToGuideValue).join(" ")}"`
      : "";
  // A path shape writes as a non-primitive enhanced-geometry with a matching
  // @svg:viewBox on the shared 21600 grid (AC-27). A path reduced from arcs/
  // quadratics on read re-emits its verbatim source on a same-format save.
  const smoothAttr =
    shape.path?.smooth && shape.path.smooth.length > 0
      ? ` xmlns:de="${DE_NAMESPACE}" de:smooth="${shape.path.smooth.map((b) => (b ? "1" : "0")).join("")}"`
      : "";
  const geometryXml = shape.path
    ? shape.pathSource?.format === "odt"
      ? shape.pathSource.xml
      : `<draw:enhanced-geometry${smoothAttr} svg:viewBox="${enhancedPathViewBox()}" draw:type="non-primitive" draw:enhanced-path="${escapeXmlAttribute(enhancedPathString(shape.path))}"/>`
    : (shape.unsupportedGeometryXml ?? `<draw:enhanced-geometry draw:type="${GEOMETRY_TO_DRAW_TYPE[shape.geometry]}"${modifiersAttr}/>`);
  return (
    `<draw:custom-shape${SHAPE_XMLNS} text:anchor-type="${anchorType}" draw:style-name="${escapeXmlAttribute(styleName)}"${idAttr}${deXmlns}${groupAttr}${rotationAttr}` +
    `${posAttrs} svg:width="${cmToOdfLength(shape.widthCm)}" svg:height="${cmToOdfLength(shape.heightCm)}">` +
    descXml +
    textBoxXml(shape, ctx) +
    geometryXml +
    "</draw:custom-shape>"
  );
}
