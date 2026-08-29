// The `.odp` -> PresentationModel pipeline (SPEC-079): verify the package's
// `mimetype`, read the master page's page size, and walk content.xml's
// office:presentation draw:pages into WYSIWYG `Slide`s — frames with
// geometry, text with effective formatting resolved through ODF style chains
// (style:style parent-style-name across content.xml automatic styles and
// styles.xml), images, shapes, groups and foreign placeholder boxes. The
// whole package is retained verbatim in `opaque` so a no-edit download stays
// byte-identical. Reading is tolerant: an unreadable page still yields a
// slide.
import { bytesToBase64Async, type OpaquePart } from "../../core/packageParts";
import { readOdfMimetype } from "../../core/packageDetect";
import { parseOdfCustomProperties, parseOdfMeta } from "../../core/odfMeta";
import type { DocumentProperties } from "../../core/documentProperties";
import type { MediaPart } from "../../core/media";
import { contentTypeForPartName } from "../../core/media";
import { unzip, ZipError } from "../../core/zip";
import { escapeXmlAttribute, firstChildElement, parseXml, serializeElement, XmlParseError } from "../../core/xml";
import {
  DRAW_NAMESPACE,
  OFFICE_NAMESPACE,
  STYLE_NAMESPACE,
  TABLE_NAMESPACE,
  TEXT_NAMESPACE,
  drawAttr,
  foAttr,
  odfColorToCss,
  odfLengthToCm,
  odfLengthToPt,
  presentationAttr,
  smilAttr,
  styleAttr,
  svgAttr,
  tableAttr,
  xlinkAttr,
} from "../../core/odfXml";
import { DRAW_TYPE_TO_GEOMETRY } from "../../core/drawingMl";
import type { ShapeGeometry } from "../../core/shapes";
import { readOdpTransition, type OdpTransitionAttrs } from "../transitions";
import { readOdpAnim } from "./anim";

const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
import {
  DEFAULT_SLIDE_HEIGHT_CM,
  DEFAULT_SLIDE_WIDTH_CM,
  DEFAULT_TABLE_CELL_PADDING_CM,
  type LayoutSlot,
  type ParsedPresentation,
  type PlaceholderRole,
  type Slide,
  type SlideFrame,
  type SlideLayout,
  type SlideParagraph,
  type SlideRun,
  type SlideTableCell,
  type SlideTableRow,
  type TableFrame,
} from "../model";
import { computeOwnerGrid, type BorderSide, type CellBorders } from "../../core/tables";
import { classifyLayout } from "../layouts";

export class OdpOpenError extends Error {}

export const INVALID_ODP_MESSAGE = "This file isn't an OpenDocument presentation we can open";

/** The package mimetype identifying an OpenDocument presentation. */
export const ODP_MIMETYPE = "application/vnd.oasis.opendocument.presentation";
/** An `.otp` template's mimetype — same package shape, opened the same way. */
export const OTP_MIMETYPE = "application/vnd.oasis.opendocument.presentation-template";

const DEFAULT_FONT_SIZE_PT = 18;

function decodeUtf8(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

interface RunProps {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: string;
  sizePt?: number;
  color?: string;
}

interface GraphicProps {
  fill?: string; // "solid" | "none" | ...
  fillColor?: string;
  stroke?: string; // "solid" | "none" | "dash"
  strokeColor?: string;
  strokeWidthPt?: number;
}

interface StyleEntry {
  parent?: string;
  text: RunProps;
  align?: SlideParagraph["align"];
  graphic: GraphicProps;
}

/** Indexes every `style:style` (by name) from content.xml's automatic styles
 * and styles.xml so a run can resolve through its `parent-style-name` chain. */
class StyleResolver {
  private readonly styles = new Map<string, StyleEntry>();
  private readonly cache = new Map<string, RunProps>();
  private readonly drawingPageTransitions = new Map<string, OdpTransitionAttrs>();
  private readonly rawStyles = new Map<string, Element>();

  /** The raw `<style:style>` element for a name, for the table reader to read
   * table-column/row/cell properties the general resolver does not model
   * (SPEC-091). */
  raw(name: string | null | undefined): Element | undefined {
    return name ? this.rawStyles.get(name) : undefined;
  }

  add(styleEl: Element): void {
    const name = styleAttr(styleEl, "name");
    if (!name) {
      return;
    }
    this.rawStyles.set(name, styleEl);
    const parent = styleAttr(styleEl, "parent-style-name") ?? undefined;
    const entry: StyleEntry = { parent, text: {}, graphic: {} };
    const textProps = firstChildElement(styleEl, STYLE_NAMESPACE, "text-properties");
    if (textProps) {
      entry.text = readTextProps(textProps);
    }
    const paraProps = firstChildElement(styleEl, STYLE_NAMESPACE, "paragraph-properties");
    const align = paraProps ? alignOf(foAttr(paraProps, "text-align")) : undefined;
    if (align) {
      entry.align = align;
    }
    const graphicProps = firstChildElement(styleEl, STYLE_NAMESPACE, "graphic-properties");
    if (graphicProps) {
      entry.graphic = readGraphicProps(graphicProps);
    }
    if (styleAttr(styleEl, "family") === "drawing-page") {
      const dpProps = firstChildElement(styleEl, STYLE_NAMESPACE, "drawing-page-properties");
      if (dpProps) {
        this.drawingPageTransitions.set(name, readDrawingPageTransition(dpProps));
      }
    }
    this.styles.set(name, entry);
  }

  /** The transition attributes carried by a page's drawing-page style, or
   * undefined when the style declares no `smil:type` (a plain page). */
  drawingPageTransition(name: string | null | undefined): OdpTransitionAttrs | undefined {
    if (!name) {
      return undefined;
    }
    const attrs = this.drawingPageTransitions.get(name);
    return attrs && attrs.smilType ? attrs : undefined;
  }

  private resolveGraphic(name: string | null | undefined): GraphicProps {
    if (!name) {
      return {};
    }
    const entry = this.styles.get(name);
    if (!entry) {
      return {};
    }
    const base = entry.parent ? this.resolveGraphic(entry.parent) : {};
    return { ...base, ...entry.graphic };
  }

  /** The resolved solid fill colour of a graphic style, or undefined when the
   * style has no fill (draw:fill="none") or no colour. */
  fillColor(name: string | null | undefined): string | undefined {
    const graphic = this.resolveGraphic(name);
    if (graphic.fill === "none") {
      return undefined;
    }
    return graphic.fillColor;
  }

  /** The resolved stroke of a graphic style, or undefined for no stroke. */
  stroke(name: string | null | undefined): { color: string; widthPt: number } | undefined {
    const graphic = this.resolveGraphic(name);
    if (graphic.stroke === "none" || !graphic.strokeColor) {
      return undefined;
    }
    return { color: graphic.strokeColor, widthPt: graphic.strokeWidthPt ?? 0.75 };
  }

  /** The effective run properties of a style, merging its parent chain. */
  resolve(name: string | null | undefined): RunProps {
    if (!name) {
      return {};
    }
    const cached = this.cache.get(name);
    if (cached) {
      return cached;
    }
    const entry = this.styles.get(name);
    if (!entry) {
      return {};
    }
    const base = entry.parent ? this.resolve(entry.parent) : {};
    const merged = mergeProps(base, entry.text);
    this.cache.set(name, merged);
    return merged;
  }

  align(name: string | null | undefined): SlideParagraph["align"] {
    let current = name ?? undefined;
    const seen = new Set<string>();
    while (current && !seen.has(current)) {
      seen.add(current);
      const entry = this.styles.get(current);
      if (entry?.align) {
        return entry.align;
      }
      current = entry?.parent;
    }
    return undefined;
  }
}

function readTextProps(textProps: Element): RunProps {
  const props: RunProps = {};
  const size = odfLengthToPt(foAttr(textProps, "font-size"));
  if (size !== undefined) {
    props.sizePt = size;
  }
  const weight = foAttr(textProps, "font-weight");
  if (weight) {
    props.bold = weight === "bold" || Number(weight) >= 600;
  }
  const style = foAttr(textProps, "font-style");
  if (style) {
    props.italic = style === "italic" || style === "oblique";
  }
  const underline = styleAttr(textProps, "text-underline-style");
  if (underline) {
    props.underline = underline !== "none";
  }
  const color = odfColorToCss(foAttr(textProps, "color"));
  if (color) {
    props.color = color;
  }
  const font = styleAttr(textProps, "font-name") ?? foAttr(textProps, "font-family");
  if (font) {
    props.fontFamily = font;
  }
  return props;
}

function readGraphicProps(props: Element): GraphicProps {
  const graphic: GraphicProps = {};
  const fill = drawAttr(props, "fill");
  if (fill) graphic.fill = fill;
  const fillColor = odfColorToCss(drawAttr(props, "fill-color"));
  if (fillColor) graphic.fillColor = fillColor;
  const stroke = drawAttr(props, "stroke");
  if (stroke) graphic.stroke = stroke;
  const strokeColor = odfColorToCss(svgAttr(props, "stroke-color"));
  if (strokeColor) graphic.strokeColor = strokeColor;
  const widthCm = odfLengthToCm(svgAttr(props, "stroke-width"));
  if (widthCm !== undefined) graphic.strokeWidthPt = (widthCm / 2.54) * 72;
  return graphic;
}

/** Reads a page's `style:drawing-page-properties` transition attributes into
 * the format-neutral bag `readOdpTransition` maps (SPEC-085). */
function readDrawingPageTransition(props: Element): OdpTransitionAttrs {
  const attrs: OdpTransitionAttrs = {};
  const smilType = smilAttr(props, "type");
  if (smilType) attrs.smilType = smilType;
  const smilSubtype = smilAttr(props, "subtype");
  if (smilSubtype) attrs.smilSubtype = smilSubtype;
  const smilDirection = smilAttr(props, "direction");
  if (smilDirection) attrs.smilDirection = smilDirection;
  const smilDur = smilAttr(props, "dur");
  if (smilDur) attrs.smilDur = smilDur;
  const speed = presentationAttr(props, "transition-speed");
  if (speed) attrs.transitionSpeed = speed;
  const transitionType = presentationAttr(props, "transition-type");
  if (transitionType) attrs.transitionType = transitionType;
  const duration = presentationAttr(props, "duration");
  if (duration) attrs.presentationDuration = duration;
  return attrs;
}

function mergeProps(base: RunProps, over: RunProps): RunProps {
  return {
    bold: over.bold ?? base.bold,
    italic: over.italic ?? base.italic,
    underline: over.underline ?? base.underline,
    fontFamily: over.fontFamily ?? base.fontFamily,
    sizePt: over.sizePt ?? base.sizePt,
    color: over.color ?? base.color,
  };
}

function alignOf(value: string | null): SlideParagraph["align"] {
  if (value === "center") return "center";
  if (value === "end" || value === "right") return "right";
  if (value === "justify") return "justify";
  if (value === "start" || value === "left") return "left";
  return undefined;
}

function roleOf(cls: string | null): PlaceholderRole | undefined {
  if (!cls) {
    return undefined;
  }
  if (cls === "title") return "title";
  if (cls === "subtitle") return "subtitle";
  if (cls === "outline" || cls === "body" || cls === "text" || cls === "notes") return "body";
  return "other";
}

function runFromProps(text: string, props: RunProps): SlideRun {
  const run: SlideRun = { text };
  if (props.bold) run.bold = true;
  if (props.italic) run.italic = true;
  if (props.underline) run.underline = true;
  if (props.fontFamily) run.fontFamily = props.fontFamily;
  if (props.sizePt !== undefined && props.sizePt !== DEFAULT_FONT_SIZE_PT) run.sizePt = props.sizePt;
  if (props.color) run.color = props.color;
  return run;
}

/** Parses a `draw:text-box`'s `text:p`/`text:list` paragraphs into
 * SlideParagraphs, resolving each run through the style chain. */
function parseTextBox(textBox: Element, role: PlaceholderRole | undefined, styles: StyleResolver): SlideParagraph[] {
  const paragraphs: SlideParagraph[] = [];
  const walk = (container: Element): void => {
    for (const child of Array.from(container.children)) {
      if (child.namespaceURI !== TEXT_NAMESPACE) {
        continue;
      }
      if (child.localName === "p") {
        const para = parseParagraph(child, role, styles);
        if (para.runs.length > 0) {
          paragraphs.push(para);
        }
      } else if (child.localName === "list" || child.localName === "list-item") {
        walk(child);
      }
    }
  };
  walk(textBox);
  return paragraphs;
}

function parseParagraph(p: Element, role: PlaceholderRole | undefined, styles: StyleResolver): SlideParagraph {
  const paraStyleName = textStyleName(p);
  const paraProps = styles.resolve(paraStyleName);
  const runs: SlideRun[] = [];
  for (const node of Array.from(p.childNodes)) {
    if (node.nodeType === 3) {
      const text = node.textContent ?? "";
      if (text.length > 0) {
        runs.push(runFromProps(text, paraProps));
      }
    } else if (node.nodeType === 1) {
      const el = node as Element;
      if (el.namespaceURI === TEXT_NAMESPACE && el.localName === "span") {
        const spanProps = mergeProps(paraProps, styles.resolve(textStyleName(el)));
        const text = el.textContent ?? "";
        if (text.length > 0) {
          runs.push(runFromProps(text, spanProps));
        }
      } else if (el.namespaceURI === TEXT_NAMESPACE && el.localName === "a") {
        // A hyperlink renders as plain styled text (out of scope: real links).
        const text = el.textContent ?? "";
        if (text.length > 0) {
          runs.push(runFromProps(text, paraProps));
        }
      }
    }
  }
  const para: SlideParagraph = { runs };
  const align = styles.align(paraStyleName);
  if (align) para.align = align;
  const bullet = role === "body" ? "disc" : undefined;
  if (bullet) para.bullet = bullet;
  return para;
}

function textStyleName(el: Element): string | null {
  return el.getAttributeNS(TEXT_NAMESPACE, "style-name");
}

interface Geometry {
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
}

function geometryOf(el: Element): Geometry | undefined {
  const x = odfLengthToCm(svgAttr(el, "x"));
  const y = odfLengthToCm(svgAttr(el, "y"));
  const w = odfLengthToCm(svgAttr(el, "width"));
  const h = odfLengthToCm(svgAttr(el, "height"));
  if (x === undefined || y === undefined || w === undefined || h === undefined) {
    return undefined;
  }
  return { xCm: x, yCm: y, widthCm: w, heightCm: h };
}

interface OdpContext {
  styles: StyleResolver;
  masterPh: Map<string, Geometry>;
  media: Map<string, MediaPart>;
  parts: Map<string, Uint8Array>;
}

function frameBase(geom: Geometry): { id: string; xCm: number; yCm: number; widthCm: number; heightCm: number } {
  return { id: crypto.randomUUID(), xCm: geom.xCm, yCm: geom.yCm, widthCm: geom.widthCm, heightCm: geom.heightCm };
}

/** Parses one shape element into a frame, or null for a non-shape. */
function parseFrameChild(el: Element, ctx: OdpContext): SlideFrame | null {
  if (el.namespaceURI !== DRAW_NAMESPACE) {
    return null;
  }
  switch (el.localName) {
    case "frame":
      return parseDrawFrame(el, ctx);
    case "custom-shape":
      return parseCustomShape(el, ctx);
    case "rect":
    case "ellipse":
    case "circle":
      return parseBasicShape(el, ctx);
    case "line":
      return parseLine(el, ctx);
    case "g":
      return parseGroup(el, ctx);
    default:
      return foreignOf(el, "Object");
  }
}

function graphicFillColor(el: Element, ctx: OdpContext): string | undefined {
  const styleName = drawAttr(el, "style-name");
  return styleName ? ctx.styles.fillColor(styleName) : undefined;
}

function parseDrawFrame(frame: Element, ctx: OdpContext): SlideFrame {
  const cls = presentationAttr(frame, "class");
  const role = roleOf(cls);
  let geom = geometryOf(frame);
  if (!geom && cls) {
    geom = ctx.masterPh.get(cls); // AC-21: master page placeholder fallback.
  }
  const base = frameBase(geom ?? { xCm: 0, yCm: 0, widthCm: 0, heightCm: 0 });

  const image = firstChildElement(frame, DRAW_NAMESPACE, "image");
  if (image) {
    const href = xlinkAttr(image, "href");
    const partPath = href ? normaliseHref(href) : undefined;
    if (partPath && collectMedia(partPath, ctx)) {
      const out: SlideFrame = { ...base, kind: "image", mediaId: partPath };
      const stroke = ctx.styles.stroke(drawAttr(frame, "style-name"));
      if (stroke) out.stroke = stroke;
      return out;
    }
    return foreignAt(frame, base, "Media");
  }

  const table = firstChildElement(frame, TABLE_NAMESPACE, "table");
  if (table) {
    const parsed = parseOdpTable(table, base, frame, ctx);
    if (parsed) {
      return parsed;
    }
    return foreignAt(frame, base, "Table"); // out of subset (AC-3b)
  }

  const textBox = firstChildElement(frame, DRAW_NAMESPACE, "text-box");
  if (textBox) {
    const paragraphs = parseTextBox(textBox, role, ctx.styles);
    const out: SlideFrame = { ...base, kind: "text", paragraphs };
    if (role) out.role = role;
    const fill = graphicFillColor(frame, ctx);
    if (fill) out.fill = { type: "solid", color: fill };
    const stroke = ctx.styles.stroke(drawAttr(frame, "style-name"));
    if (stroke) out.stroke = stroke;
    return out;
  }
  return foreignAt(frame, base, "Object");
}

// --- table (SPEC-091) -------------------------------------------------------

function round3(v: number): number {
  return Math.round(v * 1000) / 1000;
}

/** Parses `fo:border`-shaped value ("0.5pt solid #000000") into a BorderSide. */
function parseFoBorder(value: string | null | undefined): BorderSide | undefined {
  if (!value) {
    return undefined;
  }
  const trimmed = value.trim();
  if (trimmed === "none" || /^0(cm|pt|mm|in)?$/.test(trimmed)) {
    return { style: "none", widthPt: 0, color: "#000000" };
  }
  let widthPt = 1;
  let style: BorderSide["style"] = "single";
  let color = "#000000";
  for (const part of trimmed.split(/\s+/)) {
    const len = odfLengthToPt(part);
    if (len !== undefined) {
      widthPt = len;
      continue;
    }
    if (part === "solid") {
      style = "single";
    } else if (part === "double" || part === "dashed" || part === "dotted" || part === "none") {
      style = part;
    } else {
      const c = odfColorToCss(part);
      if (c) color = c;
    }
  }
  return { style, widthPt, color };
}

const MODELLED_CELL_PROP_ATTRS = new Set([
  "background-color",
  "padding",
  "border",
  "border-left",
  "border-right",
  "border-top",
  "border-bottom",
  "vertical-align",
]);

/** Reads a `table-cell` style's modelled fill / padding / valign / borders into
 * `cell`, and returns whether the style is fully modelled (so no
 * `odpCellStyleName` reference need survive — AC-3c). */
function readCellStyle(styleName: string | undefined, cell: SlideTableCell, ctx: OdpContext): boolean {
  const raw = ctx.styles.raw(styleName);
  if (!raw) {
    return true;
  }
  let fullyModelled = !styleAttr(raw, "parent-style-name");
  const cp = firstChildElement(raw, STYLE_NAMESPACE, "table-cell-properties");
  for (const child of Array.from(raw.children)) {
    if (child.namespaceURI === STYLE_NAMESPACE && child.localName === "table-cell-properties") {
      continue;
    }
    fullyModelled = false; // paragraph/text props etc. are not modelled here
  }
  if (!cp) {
    return fullyModelled;
  }
  for (const attr of Array.from(cp.attributes)) {
    if (!MODELLED_CELL_PROP_ATTRS.has(attr.localName ?? attr.name)) {
      fullyModelled = false;
    }
  }
  const bg = foAttr(cp, "background-color");
  if (bg === "transparent" || bg === "none") {
    cell.fill = { type: "none" };
  } else if (bg) {
    const c = odfColorToCss(bg);
    if (c) cell.fill = { type: "solid", color: c };
  }
  const pad = odfLengthToCm(foAttr(cp, "padding"));
  if (pad !== undefined && Math.abs(pad - DEFAULT_TABLE_CELL_PADDING_CM) > 1e-6) {
    cell.paddingCm = round3(pad);
  }
  const va = styleAttr(cp, "vertical-align");
  if (va === "middle") cell.verticalAlign = "center";
  else if (va === "bottom") cell.verticalAlign = "bottom";
  const all = parseFoBorder(foAttr(cp, "border"));
  const borders: CellBorders = {};
  const top = parseFoBorder(foAttr(cp, "border-top")) ?? all;
  const bottom = parseFoBorder(foAttr(cp, "border-bottom")) ?? all;
  const left = parseFoBorder(foAttr(cp, "border-left")) ?? all;
  const right = parseFoBorder(foAttr(cp, "border-right")) ?? all;
  if (top) borders.top = top;
  if (bottom) borders.bottom = bottom;
  if (left) borders.left = left;
  if (right) borders.right = right;
  if (Object.keys(borders).length > 0) cell.borders = borders;
  return fullyModelled;
}

/** The column width (cm) a `table-column` style declares, plus whether it is
 * fully modelled (only `style:column-width`). */
function readColumnStyle(styleName: string | undefined, ctx: OdpContext): { widthCm?: number; fullyModelled: boolean } {
  const raw = ctx.styles.raw(styleName);
  if (!raw) {
    return { fullyModelled: true };
  }
  let fullyModelled = !styleAttr(raw, "parent-style-name");
  const cp = firstChildElement(raw, STYLE_NAMESPACE, "table-column-properties");
  if (!cp) {
    return { fullyModelled };
  }
  for (const attr of Array.from(cp.attributes)) {
    if ((attr.localName ?? attr.name) !== "column-width") {
      fullyModelled = false;
    }
  }
  return { widthCm: odfLengthToCm(styleAttr(cp, "column-width")), fullyModelled };
}

/** The row height (cm) a `table-row` style declares, plus whether it is fully
 * modelled (only a row-height/min-row-height). */
function readRowStyle(styleName: string | undefined, ctx: OdpContext): { heightCm?: number; fullyModelled: boolean } {
  const raw = ctx.styles.raw(styleName);
  if (!raw) {
    return { fullyModelled: true };
  }
  let fullyModelled = !styleAttr(raw, "parent-style-name");
  const cp = firstChildElement(raw, STYLE_NAMESPACE, "table-row-properties");
  if (!cp) {
    return { fullyModelled };
  }
  const allowed = new Set(["row-height", "min-row-height", "use-optimal-row-height"]);
  for (const attr of Array.from(cp.attributes)) {
    if (!allowed.has(attr.localName ?? attr.name)) {
      fullyModelled = false;
    }
  }
  const h = odfLengthToCm(styleAttr(cp, "row-height")) ?? odfLengthToCm(styleAttr(cp, "min-row-height"));
  return { heightCm: h, fullyModelled };
}

function parseOdpCells(row: Element, colCount: number, ctx: OdpContext): SlideTableCell[] | undefined {
  const cells: SlideTableCell[] = [];
  for (const el of Array.from(row.children)) {
    if (el.namespaceURI !== TABLE_NAMESPACE) {
      continue;
    }
    if (el.localName === "covered-table-cell") {
      cells.push({ paragraphs: [], covered: true });
    } else if (el.localName === "table-cell") {
      const cell: SlideTableCell = { paragraphs: parseTextBox(el, undefined, ctx.styles) };
      const colSpan = Number(tableAttr(el, "number-columns-spanned") ?? "1");
      const rowSpan = Number(tableAttr(el, "number-rows-spanned") ?? "1");
      if (colSpan > 1) cell.colSpan = colSpan;
      if (rowSpan > 1) cell.rowSpan = rowSpan;
      const styleName = tableAttr(el, "style-name") ?? undefined;
      const fullyModelled = readCellStyle(styleName, cell, ctx);
      if (styleName && !fullyModelled) {
        cell.odpCellStyleName = styleName;
      }
      cells.push(cell);
    }
  }
  return cells.length === colCount ? cells : undefined;
}

/** Parses an in-subset `table:table` into a `TableFrame`, or `undefined` when
 * out of subset (no columns, no rows, an irregular row, or an inconsistent
 * merge map — AC-3b). Unmodelled style markup is preserved by name in
 * odpTableStyleName/odpColumnStyleNames/odpRowStyleName/odpCellStyleName. */
function parseOdpTable(
  table: Element,
  base: ReturnType<typeof frameBase>,
  frameEl: Element,
  ctx: OdpContext,
): TableFrame | undefined {
  const columnWidthsCm: number[] = [];
  const columnStyleNames: (string | undefined)[] = [];
  const rows: SlideTableRow[] = [];
  const rowStyleNames: (string | undefined)[] = [];
  let sawHeader = false;

  const walk = (parent: Element, inHeader: boolean): boolean => {
    for (const el of Array.from(parent.children)) {
      if (el.namespaceURI !== TABLE_NAMESPACE) {
        continue;
      }
      switch (el.localName) {
        case "table-column": {
          const styleName = tableAttr(el, "style-name") ?? undefined;
          const { widthCm, fullyModelled } = readColumnStyle(styleName, ctx);
          const repeat = Number(tableAttr(el, "number-columns-repeated") ?? "1");
          for (let i = 0; i < repeat; i++) {
            columnWidthsCm.push(widthCm !== undefined ? round3(widthCm) : 0);
            columnStyleNames.push(styleName && !fullyModelled ? styleName : undefined);
          }
          break;
        }
        case "table-columns":
        case "table-column-group":
          if (!walk(el, inHeader)) return false;
          break;
        case "table-header-rows":
          sawHeader = true;
          if (!walk(el, true)) return false;
          break;
        case "table-rows":
        case "table-row-group":
          if (!walk(el, inHeader)) return false;
          break;
        case "table-row": {
          const styleName = tableAttr(el, "style-name") ?? undefined;
          const { heightCm, fullyModelled } = readRowStyle(styleName, ctx);
          const cells = parseOdpCells(el, columnWidthsCm.length, ctx);
          if (!cells) {
            return false; // irregular row
          }
          const row: SlideTableRow = { cells };
          if (heightCm !== undefined) row.heightCm = round3(heightCm);
          if (inHeader) row.isHeader = true;
          rows.push(row);
          rowStyleNames.push(styleName && !fullyModelled ? styleName : undefined);
          break;
        }
        default:
          break;
      }
    }
    return true;
  };

  if (!walk(table, false)) {
    return undefined;
  }
  if (columnWidthsCm.length === 0 || rows.length === 0) {
    return undefined;
  }
  const frame: TableFrame = {
    ...base,
    xCm: round3(base.xCm),
    yCm: round3(base.yCm),
    widthCm: round3(base.widthCm),
    heightCm: round3(base.heightCm),
    kind: "table",
    rows,
    columnWidthsCm,
  };
  if (!isOdpConsistentMergeGrid(frame)) {
    return undefined;
  }
  if (sawHeader && rows[0]?.isHeader) {
    frame.firstRow = true; // the header look survives the round trip
  }
  if (columnStyleNames.some((n) => n !== undefined)) {
    frame.odpColumnStyleNames = columnStyleNames;
  }
  rows.forEach((r, i) => {
    if (rowStyleNames[i]) r.odpRowStyleName = rowStyleNames[i];
  });
  const tableStyle = tableAttr(table, "style-name") ?? undefined;
  if (tableStyle) {
    frame.odpTableStyleName = tableStyle; // table-level props are all unmodelled
  }
  void frameEl;
  return frame;
}

function isOdpConsistentMergeGrid(frame: TableFrame): boolean {
  const grid = computeOwnerGrid(frame);
  for (let r = 0; r < frame.rows.length; r++) {
    for (let c = 0; c < frame.columnWidthsCm.length; c++) {
      const owner = grid[r][c];
      const isOwnerSlot = owner.originRow === r && owner.originCol === c;
      if (isOwnerSlot === Boolean(frame.rows[r].cells[c].covered)) {
        return false;
      }
    }
  }
  return true;
}

/** Test seam (SPEC-091 golden): parse a standalone `draw:frame` table XML
 * string into a `TableFrame` (or undefined), with its automatic styles supplied
 * inline in an `office:automatic-styles` wrapper. */
export function parseOdpTableFrameForTest(drawFrameXml: string, autoStylesXml = ""): TableFrame | undefined {
  const wrapped =
    `<office:document-content ${ODP_TEST_NS}>` +
    `<office:automatic-styles>${autoStylesXml}</office:automatic-styles>` +
    `<office:body><office:presentation><draw:page>${drawFrameXml}</draw:page></office:presentation></office:body>` +
    `</office:document-content>`;
  const doc = parseXml(wrapped);
  const styles = new StyleResolver();
  addStylesFrom(doc.documentElement, styles);
  const ctx: OdpContext = { styles, masterPh: new Map(), media: new Map(), parts: new Map() };
  const frameEl = doc.documentElement.getElementsByTagNameNS(DRAW_NAMESPACE, "frame")[0];
  const frame = parseDrawFrame(frameEl, ctx);
  return frame.kind === "table" ? frame : undefined;
}

const ODP_TEST_NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
  `xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
  `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
  `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ` +
  `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ` +
  `xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"`;

function parseCustomShape(el: Element, ctx: OdpContext): SlideFrame {
  const geom = geometryOf(el);
  const base = frameBase(geom ?? { xCm: 0, yCm: 0, widthCm: 0, heightCm: 0 });
  const enhanced = firstChildElement(el, DRAW_NAMESPACE, "enhanced-geometry");
  const drawType = enhanced ? drawAttr(enhanced, "type") : null;
  const geometry = drawType ? DRAW_TYPE_TO_GEOMETRY[drawType] : undefined;
  if (!geom || !geometry) {
    return foreignAt(el, base, "Object");
  }
  return shapeFrame(el, base, geometry, ctx);
}

function parseBasicShape(el: Element, ctx: OdpContext): SlideFrame {
  const geom = geometryOf(el);
  const base = frameBase(geom ?? { xCm: 0, yCm: 0, widthCm: 0, heightCm: 0 });
  const geometry = el.localName === "rect" ? "rect" : "ellipse";
  if (!geom) {
    return foreignAt(el, base, "Object");
  }
  return shapeFrame(el, base, geometry, ctx);
}

function shapeFrame(
  el: Element,
  base: ReturnType<typeof frameBase>,
  geometry: ShapeGeometry,
  ctx: OdpContext,
): SlideFrame {
  const frame: SlideFrame = { ...base, kind: "shape", geometry };
  const fill = graphicFillColor(el, ctx);
  if (fill) frame.fill = { type: "solid", color: fill };
  const stroke = ctx.styles.stroke(drawAttr(el, "style-name"));
  if (stroke) frame.stroke = stroke;
  const textBox = firstChildElement(el, DRAW_NAMESPACE, "text-box") ?? el;
  const paragraphs = parseTextBox(textBox, undefined, ctx.styles);
  if (paragraphs.length > 0) frame.paragraphs = paragraphs;
  return frame;
}

function parseLine(el: Element, ctx: OdpContext): SlideFrame {
  const x1 = odfLengthToCm(svgAttr(el, "x1")) ?? 0;
  const y1 = odfLengthToCm(svgAttr(el, "y1")) ?? 0;
  const x2 = odfLengthToCm(svgAttr(el, "x2")) ?? 0;
  const y2 = odfLengthToCm(svgAttr(el, "y2")) ?? 0;
  const base = frameBase({
    xCm: Math.min(x1, x2),
    yCm: Math.min(y1, y2),
    widthCm: Math.abs(x2 - x1),
    heightCm: Math.abs(y2 - y1),
  });
  const frame: SlideFrame = { ...base, kind: "shape", geometry: "line" };
  const stroke = ctx.styles.stroke(drawAttr(el, "style-name"));
  if (stroke) frame.stroke = stroke;
  return frame;
}

function parseGroup(el: Element, ctx: OdpContext): SlideFrame {
  const children: SlideFrame[] = [];
  for (const child of Array.from(el.children)) {
    const frame = parseFrameChild(child, ctx);
    if (frame) {
      children.push(frame);
    }
  }
  // A draw:g has no box of its own; its bounds are its children's union.
  const bounds = unionBounds(children);
  return { ...frameBase(bounds), kind: "group", children, xml: serializeElement(el) };
}

function unionBounds(frames: SlideFrame[]): Geometry {
  if (frames.length === 0) {
    return { xCm: 0, yCm: 0, widthCm: 0, heightCm: 0 };
  }
  const minX = Math.min(...frames.map((f) => f.xCm));
  const minY = Math.min(...frames.map((f) => f.yCm));
  const maxX = Math.max(...frames.map((f) => f.xCm + f.widthCm));
  const maxY = Math.max(...frames.map((f) => f.yCm + f.heightCm));
  return { xCm: minX, yCm: minY, widthCm: maxX - minX, heightCm: maxY - minY };
}

function foreignOf(el: Element, label: string): SlideFrame {
  const geom = geometryOf(el);
  const base = frameBase(geom ?? { xCm: 0, yCm: 0, widthCm: 0, heightCm: 0 });
  return foreignAt(el, base, label);
}

function foreignAt(el: Element, base: ReturnType<typeof frameBase>, label: string): SlideFrame {
  const frame: SlideFrame = { ...base, kind: "foreign", xml: serializeElement(el), label };
  if (!(base.widthCm > 0 && base.heightCm > 0)) {
    frame.noBounds = true;
  }
  return frame;
}

function normaliseHref(href: string): string {
  return href.replace(/^\.\//, "").replace(/^\//, "");
}

function collectMedia(partPath: string, ctx: OdpContext): boolean {
  if (ctx.media.has(partPath)) {
    return true;
  }
  const raw = ctx.parts.get(partPath);
  if (!raw) {
    return false;
  }
  const contentType = contentTypeForPartName(partPath) ?? "image/png";
  ctx.media.set(partPath, { mediaId: partPath, contentType, data: base64Sync(raw) });
  return true;
}

function base64Sync(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return typeof btoa === "function" ? btoa(binary) : Buffer.from(bytes).toString("base64");
}

/** Parses an `.odp` file's bytes into the slide model. */
export async function parseOdp(
  data: ArrayBuffer | Uint8Array,
  preunzipped?: Map<string, Uint8Array>,
): Promise<ParsedPresentation> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  // Reuse the parts the open funnel already inflated for detection (see
  // parseDocx). Direct callers pass only `data` and unzip here as before.
  let parts: Map<string, Uint8Array>;
  if (preunzipped) {
    parts = preunzipped;
  } else {
    try {
      parts = await unzip(bytes);
    } catch (error) {
      if (error instanceof ZipError) {
        throw new OdpOpenError(INVALID_ODP_MESSAGE);
      }
      throw error;
    }
  }
  const mimetype = readOdfMimetype(parts);
  if (mimetype !== ODP_MIMETYPE && mimetype !== OTP_MIMETYPE) {
    throw new OdpOpenError(INVALID_ODP_MESSAGE);
  }

  const contentBytes = parts.get("content.xml");
  if (!contentBytes) {
    throw new OdpOpenError(INVALID_ODP_MESSAGE);
  }
  let contentDoc: Document;
  try {
    contentDoc = parseXml(decodeUtf8(contentBytes));
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw new OdpOpenError(INVALID_ODP_MESSAGE);
    }
    throw error;
  }
  const body = firstChildElement(contentDoc.documentElement, OFFICE_NAMESPACE, "body");
  const presentation = body && firstChildElement(body, OFFICE_NAMESPACE, "presentation");
  if (!presentation) {
    throw new OdpOpenError(INVALID_ODP_MESSAGE);
  }

  // Styles: content.xml automatic styles + styles.xml (styles + automatic).
  const styles = new StyleResolver();
  addStylesFrom(contentDoc.documentElement, styles);
  const stylesBytes = parts.get("styles.xml");
  let stylesDoc: Document | undefined;
  if (stylesBytes) {
    try {
      stylesDoc = parseXml(decodeUtf8(stylesBytes));
      addStylesFrom(stylesDoc.documentElement, styles);
    } catch {
      stylesDoc = undefined;
    }
  }

  const masterPh = buildMasterPlaceholders(stylesDoc);
  const { widthCm, heightCm } = pageSize(stylesDoc, contentDoc);

  const media = new Map<string, MediaPart>();
  const ctx: OdpContext = { styles, masterPh, media, parts };

  // Split content.xml around its draw:pages (SPEC-080): the writer's
  // splice-equality envelope + each page's EXACT source substring, so
  // `prefix + pages.join("") + suffix` reproduces content.xml byte-for-byte.
  const contentText = decodeUtf8(contentBytes);
  const { prefix, suffix, pageSources } = splitContentEnvelope(contentText);

  // The read-derived layout catalogue (SPEC-089 AC-2/3), keyed
  // `${master}::${pageLayout}` (or `${master}::` for a page naming no layout).
  const pageLayoutStyles = indexPageLayoutStyles(stylesDoc, contentDoc);
  const layoutsByKey = new Map<string, SlideLayout>();

  const slides: Slide[] = [];
  let pageIndex = 0;
  for (const child of Array.from(presentation.children)) {
    if (child.namespaceURI === DRAW_NAMESPACE && child.localName === "page") {
      const slide = parsePage(child, ctx);
      if (pageSources[pageIndex] !== undefined) {
        slide.sourceXml = pageSources[pageIndex];
      }
      deriveOdpLayout(child, slide, ctx, pageLayoutStyles, layoutsByKey);
      slides.push(slide);
      pageIndex += 1;
    }
  }
  const layouts = [...layoutsByKey.values()];

  // Document properties from meta.xml.
  const metaBytes = parts.get("meta.xml");
  const metaText = metaBytes ? decodeUtf8(metaBytes) : undefined;
  const properties: DocumentProperties = metaText ? parseOdfMeta(metaText) : {};
  if (metaText) {
    const { customProperties, unsupportedCustomProperties } = parseOdfCustomProperties(metaText);
    if (customProperties.length > 0) {
      properties.customProperties = customProperties;
    }
    if (unsupportedCustomProperties.length > 0) {
      properties.unsupportedCustomProperties = unsupportedCustomProperties;
    }
  }

  const opaque: OpaquePart[] = [];
  for (const [name, partBytes] of parts) {
    opaque.push({ name, data: await bytesToBase64Async(partBytes) });
  }

  return {
    slides,
    layouts,
    slideWidthCm: widthCm,
    slideHeightCm: heightCm,
    media: [...media.values()],
    opaque,
    properties,
    odpContentEnvelope: { prefix, suffix },
  };
}

/** Splits a content.xml string into everything before its first `draw:page`,
 * everything from the `office:presentation` close onwards, and each page's
 * exact source substring (the region between pages sliced at every
 * `<draw:page` boundary, so trailing whitespace stays with its page and the
 * three pieces recombine byte-for-byte). A deck with no pages yields empty
 * `pageSources` and a prefix/suffix that already recombine. */
function splitContentEnvelope(content: string): { prefix: string; suffix: string; pageSources: string[] } {
  const closeIdx = content.lastIndexOf("</office:presentation>");
  const suffixStart = closeIdx >= 0 ? closeIdx : content.length;
  // `<draw:page` must start a `draw:page` element, NOT `<draw:page-thumbnail`
  // (which a `presentation:notes` block carries): only accept it when followed
  // by a name-terminating character.
  const isPageStart = (at: number): boolean => {
    const after = content["<draw:page".length + at];
    return after === " " || after === ">" || after === "/" || after === "\t" || after === "\n" || after === "\r";
  };
  const nextPage = (from: number): number => {
    for (let idx = content.indexOf("<draw:page", from); idx !== -1; idx = content.indexOf("<draw:page", idx + 1)) {
      if (idx >= suffixStart) {
        return -1;
      }
      if (isPageStart(idx)) {
        return idx;
      }
    }
    return -1;
  };
  const firstPage = nextPage(0);
  if (firstPage < 0) {
    return { prefix: content.slice(0, suffixStart), suffix: content.slice(suffixStart), pageSources: [] };
  }
  const starts: number[] = [];
  for (let idx = firstPage; idx !== -1; idx = nextPage(idx + 1)) {
    starts.push(idx);
  }
  const pageSources = starts.map((start, i) =>
    content.slice(start, i + 1 < starts.length ? starts[i + 1] : suffixStart),
  );
  return { prefix: content.slice(0, firstPage), suffix: content.slice(suffixStart), pageSources };
}

function parsePage(page: Element, ctx: OdpContext): Slide {
  const frames: SlideFrame[] = [];
  // wire id (`xml:id`/`draw:id`) → minted frame id, so an `anim:*` tree's
  // `smil:targetElement` resolves back to the frame it targets (SPEC-086).
  const frameForWire = new Map<string, string>();
  for (const child of Array.from(page.children)) {
    const frame = parseFrameChild(child, ctx);
    if (frame) {
      frames.push(frame);
      const wireId = child.getAttributeNS(XML_NAMESPACE, "id") ?? drawAttr(child, "id");
      if (wireId) {
        frameForWire.set(wireId, frame.id);
      }
    }
  }
  const slide: Slide = { id: crypto.randomUUID(), frames, sourceXml: serializeElement(page) };
  const preserved: Record<string, string> = {};
  const pageAttrs = Array.from(page.attributes)
    .map((attr) => ` ${attr.name}="${escapeXmlAttribute(attr.value)}"`)
    .join("");
  if (pageAttrs) {
    preserved["@page"] = pageAttrs;
  }
  const animEls: Element[] = [];
  for (const child of Array.from(page.children)) {
    if (child.namespaceURI === PRESENTATION_NS && child.localName === "notes") {
      // SPEC-083: notes move from `preserved` into the modelled field — the
      // reader parses the notes text-box rather than retaining it opaquely.
      const notes = readOdpNotes(child);
      if (notes) {
        slide.notes = notes;
      }
    } else if (isAnimationElement(child)) {
      animEls.push(child);
    }
  }
  // A canonical (or Impress-authored equivalent) `anim:*` tree whose targets all
  // resolve becomes the modelled `animations`; anything out of subset stays
  // preserved verbatim (AC-4).
  if (animEls.length > 0) {
    const animations = readOdpAnim(animEls, (wireId) => frameForWire.get(wireId));
    if (animations) {
      slide.animations = animations;
    } else {
      preserved["anim"] = animEls.map((el) => serializeElement(el)).join("");
    }
  }
  // Slide transition (SPEC-085): the page's drawing-page style carries the smil
  // attributes. A subset combination becomes the modelled field; an
  // out-of-subset one leaves the field unset and marks the honest
  // "original transition kept" state (its bytes survive via the untouched-page
  // path, or via the retained `draw:style-name` when the page rebuilds).
  const transitionAttrs = ctx.styles.drawingPageTransition(drawAttr(page, "style-name"));
  if (transitionAttrs) {
    const transition = readOdpTransition(transitionAttrs);
    if (transition) {
      slide.transition = transition;
    } else {
      preserved["draw:transition"] = "1";
    }
  }
  if (Object.keys(preserved).length > 0) {
    slide.preserved = preserved;
  }
  return slide;
}

/** The plain-text speaker notes of a page (SPEC-083, AC-3/4): the
 * `draw:frame presentation:class="notes"` inside `presentation:notes`; its
 * `draw:text-box` paragraphs map to entries with formatting flattened
 * (`draw:page-thumbnail` is ignored). Trailing empty paragraphs are dropped;
 * `undefined` when nothing text-bearing remains. */
function readOdpNotes(notesEl: Element): string[] | undefined {
  for (const frame of Array.from(notesEl.getElementsByTagNameNS(DRAW_NAMESPACE, "frame"))) {
    if (presentationAttr(frame, "class") !== "notes") {
      continue;
    }
    const textBox = firstChildElement(frame, DRAW_NAMESPACE, "text-box");
    if (!textBox) {
      return undefined;
    }
    const paragraphs = Array.from(textBox.getElementsByTagNameNS(TEXT_NAMESPACE, "p")).map((p) => p.textContent ?? "");
    while (paragraphs.length > 0 && paragraphs[paragraphs.length - 1].length === 0) {
      paragraphs.pop();
    }
    return paragraphs.length > 0 ? paragraphs : undefined;
  }
  return undefined;
}

/** Whether an element is one of a page's animation children (`anim:*`) the
 * page serialiser re-emits verbatim (AC-38). */
function isAnimationElement(el: Element): boolean {
  if (el.prefix === "anim") {
    return true;
  }
  const ns = el.namespaceURI ?? "";
  return ns.includes("animation") || ns.includes("smil");
}

const PRESENTATION_NS = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";

/** Adds every `style:style` under a document root's `office:styles` /
 * `office:automatic-styles` to the resolver. */
function addStylesFrom(root: Element, resolver: StyleResolver): void {
  for (const container of ["automatic-styles", "styles"]) {
    const stylesEl = firstChildElement(root, OFFICE_NAMESPACE, container);
    if (!stylesEl) {
      continue;
    }
    for (const styleEl of Array.from(stylesEl.children)) {
      if (styleEl.namespaceURI === STYLE_NAMESPACE && styleEl.localName === "style") {
        resolver.add(styleEl);
      }
    }
  }
}

/** class → geometry from the master pages' presentation placeholders (AC-21). */
function buildMasterPlaceholders(stylesDoc: Document | undefined): Map<string, Geometry> {
  const map = new Map<string, Geometry>();
  if (!stylesDoc) {
    return map;
  }
  const masterStyles = firstChildElement(stylesDoc.documentElement, OFFICE_NAMESPACE, "master-styles");
  if (!masterStyles) {
    return map;
  }
  const frames = masterStyles.getElementsByTagNameNS(DRAW_NAMESPACE, "frame");
  for (const frame of Array.from(frames)) {
    const cls = presentationAttr(frame, "class");
    const geom = geometryOf(frame);
    if (cls && geom && !map.has(cls)) {
      map.set(cls, geom);
    }
  }
  return map;
}

/** Indexes every `style:presentation-page-layout` by its `style:name`
 * (SPEC-089 AC-3), across styles.xml and content.xml automatic styles. */
function indexPageLayoutStyles(stylesDoc: Document | undefined, contentDoc: Document): Map<string, Element> {
  const map = new Map<string, Element>();
  for (const doc of [stylesDoc, contentDoc]) {
    if (!doc) {
      continue;
    }
    for (const el of Array.from(doc.documentElement.getElementsByTagNameNS(STYLE_NAMESPACE, "presentation-page-layout"))) {
      const name = styleAttr(el, "name");
      if (name && !map.has(name)) {
        map.set(name, el);
      }
    }
  }
  return map;
}

/** The `LayoutSlot[]` a `style:presentation-page-layout` declares (ODF 1.2
 * §9.6.2): one per `presentation:placeholder` child — role from
 * `roleOf(presentation:object)`, geometry from svg:x/y/width/height in cm,
 * `phType`/`phIdx` left undefined (ODF has no `p:ph idx`). */
function odpSlotsFromPageLayout(styleEl: Element): LayoutSlot[] {
  const slots: LayoutSlot[] = [];
  for (const ph of Array.from(styleEl.getElementsByTagNameNS(PRESENTATION_NS, "placeholder"))) {
    const object = presentationAttr(ph, "object");
    const geom = geometryOf(ph);
    slots.push({
      role: roleOf(object) ?? "other",
      xCm: geom?.xCm ?? 0,
      yCm: geom?.yCm ?? 0,
      widthCm: geom?.widthCm ?? 0,
      heightCm: geom?.heightCm ?? 0,
    });
  }
  return slots;
}

/** The `LayoutSlot[]` a page's own placeholder `draw:frame`s imply, for a page
 * that names no page-layout style (SPEC-089 AC-3): each direct-child
 * `draw:frame presentation:class` with a resolvable role, geometry from its own
 * box (or the master placeholder fallback). */
function odpSlotsFromPage(page: Element, ctx: OdpContext): LayoutSlot[] {
  const slots: LayoutSlot[] = [];
  for (const child of Array.from(page.children)) {
    if (child.namespaceURI !== DRAW_NAMESPACE || child.localName !== "frame") {
      continue;
    }
    const cls = presentationAttr(child, "class");
    const role = roleOf(cls);
    if (!role) {
      continue;
    }
    const geom = geometryOf(child) ?? (cls ? ctx.masterPh.get(cls) : undefined);
    slots.push({
      role,
      xCm: geom?.xCm ?? 0,
      yCm: geom?.yCm ?? 0,
      widthCm: geom?.widthCm ?? 0,
      heightCm: geom?.heightCm ?? 0,
    });
  }
  return slots;
}

/** Derives one page's layout catalogue entry and assigns the slide's
 * `layoutKey` (SPEC-089 AC-2/3): a referenced `style:presentation-page-layout`
 * keyed `${master}::${pageLayout}`, else a synthetic entry keyed `${master}::`
 * classified from the page's own placeholder frames. Registers the entry in
 * `byKey` (first page per key wins its slots). */
function deriveOdpLayout(
  page: Element,
  slide: Slide,
  ctx: OdpContext,
  pageLayoutStyles: Map<string, Element>,
  byKey: Map<string, SlideLayout>,
): void {
  const master = drawAttr(page, "master-page-name") ?? "";
  const pageLayoutName = presentationAttr(page, "presentation-page-layout-name");
  if (pageLayoutName) {
    const key = `${master}::${pageLayoutName}`;
    slide.layoutKey = key;
    if (!byKey.has(key)) {
      const styleEl = pageLayoutStyles.get(pageLayoutName);
      const slots = styleEl ? odpSlotsFromPageLayout(styleEl) : [];
      const layout: SlideLayout = { key, kind: classifyLayout(undefined, slots), name: pageLayoutName, slots };
      if (master) layout.masterName = master;
      byKey.set(key, layout);
    }
    return;
  }
  const key = `${master}::`;
  slide.layoutKey = key;
  if (!byKey.has(key)) {
    const slots = odpSlotsFromPage(page, ctx);
    const layout: SlideLayout = { key, kind: classifyLayout(undefined, slots), name: master || "Default", slots };
    if (master) layout.masterName = master;
    byKey.set(key, layout);
  }
}

/** The slide size from the master page's `style:page-layout`
 * (`fo:page-width`/`fo:page-height`), else the default. */
function pageSize(stylesDoc: Document | undefined, contentDoc: Document): { widthCm: number; heightCm: number } {
  const search = (doc: Document | undefined): { widthCm: number; heightCm: number } | undefined => {
    if (!doc) {
      return undefined;
    }
    const layouts = doc.documentElement.getElementsByTagNameNS(STYLE_NAMESPACE, "page-layout-properties");
    for (const props of Array.from(layouts)) {
      const w = odfLengthToCm(foAttr(props, "page-width"));
      const h = odfLengthToCm(foAttr(props, "page-height"));
      if (w !== undefined && h !== undefined) {
        return { widthCm: w, heightCm: h };
      }
    }
    return undefined;
  };
  return (
    search(stylesDoc) ?? search(contentDoc) ?? { widthCm: DEFAULT_SLIDE_WIDTH_CM, heightCm: DEFAULT_SLIDE_HEIGHT_CM }
  );
}
