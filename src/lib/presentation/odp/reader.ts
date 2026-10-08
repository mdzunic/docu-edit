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
import { contentTypeForPartName, MAX_MEDIA_SIZE_BYTES } from "../../core/media";
import { mediaKindForContentType } from "../mediaFrames";
import { unzip, ZipError } from "../../core/zip";
import { childElements, escapeXmlAttribute, firstChildElement, parseXml, serializeElement, XmlParseError } from "../../core/xml";
import {
  DRAW_NAMESPACE,
  FO_NAMESPACE,
  OFFICE_NAMESPACE,
  STYLE_NAMESPACE,
  SVG_NAMESPACE,
  TABLE_NAMESPACE,
  TEXT_NAMESPACE,
  drawAttr,
  foAttr,
  officeAttr,
  odfColorToCss,
  odfLengthToCm,
  odfLengthToPt,
  odfPercentage,
  presentationAttr,
  smilAttr,
  styleAttr,
  svgAttr,
  tableAttr,
  textAttr,
  xlinkAttr,
} from "../../core/odfXml";
import { bulletStyleForChar, type ListMarkerStyle } from "../../core/lists";
import { sanitizeHref } from "../../core/urls";
import { DRAW_TYPE_TO_GEOMETRY } from "../../core/drawingMl";
import { guideValueToAdjustment } from "../../core/shapeCatalogue";
import { normalizeRotationDeg } from "../../core/shapes";
import { CHART_NAMESPACE, parseOdfChartXml } from "../../core/chartPartOdf";
import type { ShapeGeometry } from "../../core/shapes";
import { readOdpTransition, type OdpTransitionAttrs } from "../transitions";
import { readOdpAnim } from "./anim";
import { effectiveHangingCm, effectiveIndentCm } from "../slideLists";
import { isLinkableKind } from "../slideLinks";
import { DEFAULT_INSETS_CM } from "../textBody";

const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
import {
  DEFAULT_SLIDE_HEIGHT_CM,
  DEFAULT_SLIDE_WIDTH_CM,
  DEFAULT_TABLE_CELL_PADDING_CM,
  type DeckHeaderFooter,
  type LayoutSlot,
  type ParsedPresentation,
  type MediaFrame,
  type PlaceholderRole,
  type Slide,
  type SlideFieldPlaceholder,
  type SlideFrame,
  type SlideLayout,
  type ShapeFrame,
  type SlideMediaKind,
  type SlideParagraph,
  type SlideRun,
  type SlideTableCell,
  type SlideTableRow,
  type TableFrame,
  type TextBodyProps,
} from "../model";
import { DEFAULT_DATE_FORMAT, defaultFieldRect } from "../headerFooter";
import { formatFromOdtDateStyleName } from "../../core/dateTimeFormats";
import { computeOwnerGrid, type BorderSide, type CellBorders } from "../../core/tables";
import { classifyLayout } from "../layouts";
import { DEFAULT_THEME, type ModelColor } from "../../core/theme";
import { readOdpTheme, themeParentSlot } from "./theme";

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
  /** SPEC-094: a `de-theme-text-*`-parented style's colour reads as a ref. */
  color?: ModelColor;
  /** SPEC-097: `fo:letter-spacing` in points ("normal" reads as an explicit
   *  0 so it overrides a parent's spacing). */
  letterSpacingPt?: number;
}

interface GraphicProps {
  fill?: string; // "solid" | "none" | ...
  /** SPEC-094: a `de-theme-*`-parented style's fill reads as a ref. */
  fillColor?: ModelColor;
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
  private readonly drawingPageDisplays = new Map<string, OdpDisplayFlags>();
  private readonly rawStyles = new Map<string, Element>();
  private readonly listStyles = new Map<string, Element>();

  /** The raw `<style:style>` element for a name, for the table reader to read
   * table-column/row/cell properties the general resolver does not model
   * (SPEC-091). */
  raw(name: string | null | undefined): Element | undefined {
    return name ? this.rawStyles.get(name) : undefined;
  }

  add(styleEl: Element): void {
    const name = styleAttr(styleEl, "name");
    // First registration wins: content.xml's automatic styles are added
    // before styles.xml's, and Impress reuses the same automatic names (P1,
    // T1, …) in both parts — a slide's paragraph means content.xml's.
    if (!name || this.styles.has(name)) {
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
    // SPEC-094: a style parented on a `de-theme-*` scheme style carries a
    // colour REFERENCE — the literal it also declares is the resolved value
    // the writer put beside the link for other readers.
    const themeSlot = themeParentSlot(parent);
    if (themeSlot) {
      if (entry.text.color !== undefined) {
        entry.text.color = { theme: themeSlot };
      }
      if (entry.graphic.fillColor !== undefined) {
        entry.graphic.fillColor = { theme: themeSlot };
      }
    }
    if (styleAttr(styleEl, "family") === "drawing-page") {
      const dpProps = firstChildElement(styleEl, STYLE_NAMESPACE, "drawing-page-properties");
      if (dpProps) {
        this.drawingPageTransitions.set(name, readDrawingPageTransition(dpProps));
        this.drawingPageDisplays.set(name, readDrawingPageDisplay(dpProps));
      }
    }
    this.styles.set(name, entry);
  }

  /** Indexes a `text:list-style` by name (SPEC-097); first registration wins,
   * as for `style:style`. */
  addListStyle(listStyleEl: Element): void {
    const name = styleAttr(listStyleEl, "name");
    if (name && !this.listStyles.has(name)) {
      this.listStyles.set(name, listStyleEl);
    }
  }

  /** The `text:list-style` element of a name, or undefined. */
  listStyle(name: string | null | undefined): Element | undefined {
    return name ? this.listStyles.get(name) : undefined;
  }

  /** Walks a style's parent chain from the style itself outwards, handing
   * each style's `style:{propsLocal}` element to `pick`; the first defined
   * result wins (SPEC-097's paragraph and text-body properties resolve the
   * way `align` does). */
  chainPick<T>(name: string | null | undefined, propsLocal: string, pick: (props: Element) => T | undefined): T | undefined {
    let current = name ?? undefined;
    const seen = new Set<string>();
    while (current && !seen.has(current)) {
      seen.add(current);
      const styleEl = this.rawStyles.get(current);
      const props = styleEl ? firstChildElement(styleEl, STYLE_NAMESPACE, propsLocal) : undefined;
      if (props) {
        const value = pick(props);
        if (value !== undefined) {
          return value;
        }
      }
      current = this.styles.get(current)?.parent;
    }
    return undefined;
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

  /** The `presentation:display-*` field flags a page's drawing-page style
   * carries (SPEC-093); an absent flag is left undefined (read as "show"). */
  drawingPageDisplay(name: string | null | undefined): OdpDisplayFlags {
    return (name ? this.drawingPageDisplays.get(name) : undefined) ?? {};
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
  fillColor(name: string | null | undefined): ModelColor | undefined {
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
  const spacing = foAttr(textProps, "letter-spacing");
  if (spacing === "normal") {
    props.letterSpacingPt = 0;
  } else {
    const pt = odfLengthToPt(spacing);
    if (pt !== undefined) {
      // −10…10 pt in 0.1 steps (SPEC-097) — the cm wire's r3 rounding
      // (≤ 0.015pt) snaps back onto the step.
      props.letterSpacingPt = Math.round(Math.min(10, Math.max(-10, pt)) * 10) / 10;
    }
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

/** A page's `presentation:display-footer`/`display-page-number`/
 * `display-date-time` flags (SPEC-093); absent leaves the field undefined,
 * read as "show" (the ODF default — confirmed by the golden fixture). */
interface OdpDisplayFlags {
  footer?: boolean;
  pageNumber?: boolean;
  dateTime?: boolean;
}

function readDrawingPageDisplay(props: Element): OdpDisplayFlags {
  const flags: OdpDisplayFlags = {};
  const footer = presentationAttr(props, "display-footer");
  if (footer !== null) flags.footer = footer === "true";
  const pageNumber = presentationAttr(props, "display-page-number");
  if (pageNumber !== null) flags.pageNumber = pageNumber === "true";
  const dateTime = presentationAttr(props, "display-date-time");
  if (dateTime !== null) flags.dateTime = dateTime === "true";
  return flags;
}

/** One `presentation:date-time-decl` (SPEC-093): its source, the auto date's
 * `style:data-style-name`, and the fixed text. */
interface OdpDateDecl {
  source: "fixed" | "current-date";
  dataStyleName?: string;
  text: string;
}

/** The deck's `office:presentation` field declarations (SPEC-093): footer name
 * → text, and date-time name → decl. `presentation:header-decl` is preserved
 * untouched (it rides the envelope prefix) and is not modelled (AC-95). */
interface OdpDecls {
  footers: Map<string, string>;
  dateTimes: Map<string, OdpDateDecl>;
}

function readDeclarations(presentation: Element): OdpDecls {
  const footers = new Map<string, string>();
  const dateTimes = new Map<string, OdpDateDecl>();
  for (const child of Array.from(presentation.children)) {
    if (child.namespaceURI !== PRESENTATION_NS) {
      continue;
    }
    const name = presentationAttr(child, "name");
    if (!name) {
      continue;
    }
    if (child.localName === "footer-decl") {
      footers.set(name, child.textContent ?? "");
    } else if (child.localName === "date-time-decl") {
      const decl: OdpDateDecl = {
        source: presentationAttr(child, "source") === "fixed" ? "fixed" : "current-date",
        text: child.textContent ?? "",
      };
      const dataStyle = styleAttr(child, "data-style-name");
      if (dataStyle) {
        decl.dataStyleName = dataStyle;
      }
      dateTimes.set(name, decl);
    }
  }
  return { footers, dateTimes };
}

function mergeProps(base: RunProps, over: RunProps): RunProps {
  return {
    bold: over.bold ?? base.bold,
    italic: over.italic ?? base.italic,
    underline: over.underline ?? base.underline,
    fontFamily: over.fontFamily ?? base.fontFamily,
    sizePt: over.sizePt ?? base.sizePt,
    color: over.color ?? base.color,
    letterSpacingPt: over.letterSpacingPt ?? base.letterSpacingPt,
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

function runFromProps(text: string, props: RunProps, link?: RunLink): SlideRun {
  const run: SlideRun = { text };
  if (props.bold) run.bold = true;
  if (props.italic) run.italic = true;
  if (props.underline) run.underline = true;
  if (props.fontFamily) run.fontFamily = props.fontFamily;
  if (props.sizePt !== undefined && props.sizePt !== DEFAULT_FONT_SIZE_PT) run.sizePt = props.sizePt;
  if (props.color) run.color = props.color;
  if (props.letterSpacingPt) run.letterSpacingPt = props.letterSpacingPt;
  if (link) {
    run.linkHref = link.href;
    if (link.tooltip) run.linkTooltip = link.tooltip;
  }
  return run;
}

/** A run's hyperlink while walking a `text:a` (SPEC-097): a
 * `sanitizeHref`-clean external URL plus its `office:title`. */
interface RunLink {
  href: string;
  tooltip?: string;
}

/** The modelled link of an `xlink:href` (text:a / draw:a): sanitised, or null
 * for a dangerous/empty URL (read as plain text) and for an in-deck `#…`
 * target — the slide-jump link form is golden-pending (textDepthGolden.ts),
 * so it keeps today's plain-text reading rather than a guessed `linkSlide`. */
function linkHrefOf(raw: string | null): string | null {
  if (raw === null || raw.trim().startsWith("#")) {
    return null;
  }
  return sanitizeHref(raw);
}

/** The list a paragraph sits in while walking a text body (SPEC-097): the
 * resolved `text:list-style` (undefined = unresolvable → disc) and the
 * 0-based nesting depth (= the slide paragraph level). */
interface ListScope {
  style: Element | undefined;
  depth: number;
}

/** Where a paragraph sits: outside any list, in a `text:list-header` (an
 * explicitly unnumbered paragraph), or in a `text:list-item`. */
type ListPlace =
  | { kind: "none" }
  | { kind: "header" }
  | { kind: "item"; scope: ListScope; itemStart?: number };

/** Parses a `draw:text-box`'s `text:p`/`text:list` paragraphs into
 * SlideParagraphs, resolving each run through the style chain. A `text:list`
 * nests by depth into the paragraph `level` and its `text:list-style` level
 * styles into the list fields (SPEC-097); a `text:list-header` paragraph is
 * explicitly not a list item. */
function parseTextBox(textBox: Element, role: PlaceholderRole | undefined, ctx: OdpContext): SlideParagraph[] {
  const paragraphs: SlideParagraph[] = [];
  const walk = (container: Element, place: ListPlace): void => {
    for (const child of Array.from(container.children)) {
      if (child.namespaceURI !== TEXT_NAMESPACE) {
        continue;
      }
      if (child.localName === "p") {
        const para = parseParagraph(child, role, ctx, place);
        if (para.runs.length > 0) {
          paragraphs.push(para);
        }
      } else if (child.localName === "list") {
        const parent = place.kind === "item" ? place.scope : undefined;
        const styleName = textAttr(child, "style-name");
        // An own style name wins; a nested list without one inherits its
        // container's (ODF 1.2 §19.880); an unresolvable name reads as disc.
        const style = styleName !== null ? ctx.styles.listStyle(styleName) : parent?.style;
        const scope: ListScope = { style, depth: parent ? parent.depth + 1 : 0 };
        for (const entry of Array.from(child.children)) {
          if (entry.namespaceURI !== TEXT_NAMESPACE) {
            continue;
          }
          if (entry.localName === "list-item") {
            const start = Number(textAttr(entry, "start-value"));
            walk(entry, { kind: "item", scope, itemStart: Number.isInteger(start) && start > 1 ? start : undefined });
          } else if (entry.localName === "list-header") {
            walk(entry, { kind: "header" });
          }
        }
      } else if (child.localName === "list-item") {
        // A stray list-item outside a text:list: today's flattening.
        walk(child, place);
      }
    }
  };
  walk(textBox, { kind: "none" });
  return paragraphs;
}

/** One `text:list-level-style-*` read into the paragraph list fields, plus
 * the level's own indent/hanging (the paragraph style's fo:margin-left/
 * fo:text-indent win over these). */
interface OdpListLevel {
  bullet: "disc" | "number";
  bulletStyle?: ListMarkerStyle;
  bulletChar?: string;
  bulletImageId?: string;
  startValue?: number;
  indentCm?: number;
  hangingCm?: number;
}

const NUM_FORMAT_STYLES: Record<string, ListMarkerStyle> = {
  "1": "decimal",
  a: "lowerAlpha",
  A: "upperAlpha",
  i: "lowerRoman",
  I: "upperRoman",
};

/** The level style (1-based `text:level`) of a list style for a 0-based
 * depth, read into list fields; a missing style or level reads as disc. */
function readListLevel(listStyle: Element | undefined, depth: number, ctx: OdpContext): OdpListLevel {
  const wanted = String(depth + 1);
  const levelEl = listStyle
    ? Array.from(listStyle.children).find(
        (el) => el.namespaceURI === TEXT_NAMESPACE && el.localName.startsWith("list-level-style-") && textAttr(el, "level") === wanted,
      )
    : undefined;
  if (!levelEl) {
    return { bullet: "disc" };
  }
  const level: OdpListLevel = { bullet: "disc" };
  if (levelEl.localName === "list-level-style-number") {
    level.bullet = "number";
    const style = NUM_FORMAT_STYLES[styleAttr(levelEl, "num-format") ?? "1"] ?? "decimal";
    if (style !== "decimal") level.bulletStyle = style;
    const start = Number(textAttr(levelEl, "start-value"));
    if (Number.isInteger(start) && start >= 1) level.startValue = Math.min(start, 32767);
  } else if (levelEl.localName === "list-level-style-image") {
    const href = xlinkAttr(levelEl, "href");
    const partPath = href ? normaliseHref(href) : undefined;
    if (partPath && collectMedia(partPath, ctx)) level.bulletImageId = partPath;
  } else if (levelEl.localName === "list-level-style-bullet") {
    const char = textAttr(levelEl, "bullet-char") ?? "";
    const style = char ? bulletStyleForChar(char) : undefined;
    if (style) {
      if (style !== "disc") level.bulletStyle = style;
    } else if (char) {
      level.bulletChar = char;
    }
  }
  // The level's indents: label-alignment mode (ODF 1.2, Impress) carries the
  // text edge and a negative first-line indent; the older position mode
  // carries space-before + min-label-width.
  const props = firstChildElement(levelEl, STYLE_NAMESPACE, "list-level-properties");
  if (props) {
    const align = firstChildElement(props, STYLE_NAMESPACE, "list-level-label-alignment");
    if (align) {
      const margin = odfLengthToCm(foAttr(align, "margin-left"));
      const indent = odfLengthToCm(foAttr(align, "text-indent"));
      if (margin !== undefined) level.indentCm = round3(margin);
      if (indent !== undefined && indent <= 0) level.hangingCm = round3(-indent);
    } else {
      const before = odfLengthToCm(textAttr(props, "space-before"));
      const label = odfLengthToCm(textAttr(props, "min-label-width"));
      if (before !== undefined || label !== undefined) {
        level.indentCm = round3((before ?? 0) + (label ?? 0));
        level.hangingCm = round3(label ?? 0);
      }
    }
  }
  return level;
}

/** The paragraph-properties attributes SPEC-097 models — everything else on a
 * paragraph's own automatic style rides `odpParaProps`. A positive
 * fo:text-indent (a first-line indent) is not modelled and is kept there. */
const MODELLED_PARA_ATTRS = new Set([
  "fo:text-align",
  "fo:line-height",
  "style:line-height-at-least",
  "fo:margin-top",
  "fo:margin-bottom",
  "fo:margin-left",
]);

/** The prefixes the writer's documents bind, by namespace — a preserved
 * attribute is re-spelt with its canonical prefix so the writer's
 * modelled-wins de-duplication (by name) can never emit one attribute twice
 * under two spellings. Other namespaces keep the source's own prefix (bound
 * on the retained content.xml root). */
const CANONICAL_PREFIX: Record<string, string> = {
  [FO_NAMESPACE]: "fo",
  [STYLE_NAMESPACE]: "style",
  [TEXT_NAMESPACE]: "text",
  [DRAW_NAMESPACE]: "draw",
  [SVG_NAMESPACE]: "svg",
  [OFFICE_NAMESPACE]: "office",
};

function canonicalAttrName(attr: Attr): string {
  const prefix = attr.namespaceURI ? CANONICAL_PREFIX[attr.namespaceURI] : undefined;
  return prefix ? `${prefix}:${attr.localName}` : attr.name;
}

/** The unmodelled remainder of a `style:paragraph-properties` element as
 * SPEC-097's `odpParaProps` string: its other attributes in source order
 * (`name="value"`, space-separated, values re-escaped) followed directly by
 * its child elements verbatim (e.g. `style:tab-stops`) — the attribute part
 * never contains "<", so the first "<" starts the children. Undefined when
 * nothing is left. */
function unmodelledParaProps(props: Element): string | undefined {
  const attrs: string[] = [];
  for (const attr of Array.from(props.attributes)) {
    if (!attr.prefix || attr.prefix === "xmlns") {
      continue;
    }
    const name = canonicalAttrName(attr);
    if (MODELLED_PARA_ATTRS.has(name)) {
      continue;
    }
    if (name === "fo:text-indent") {
      const cm = odfLengthToCm(attr.value);
      if (cm !== undefined && cm <= 0) {
        continue; // a hanging (or zero) indent is modelled as hangingCm
      }
    }
    attrs.push(`${name}="${escapeXmlAttribute(attr.value)}"`);
  }
  const children = Array.from(props.children).map((child) => serializeElement(child)).join("");
  const out = attrs.join(" ") + children;
  return out.length > 0 ? out : undefined;
}

/** Points from an ODF length, snapped to 0.1pt — the cm wire's r3 rounding
 * (≤ 0.015pt) lands back on the step a spacing value was authored at. */
function spacingPt(value: string | null): number | undefined {
  const pt = odfLengthToPt(value);
  return pt === undefined ? undefined : Math.round(pt * 10) / 10;
}

function parseParagraph(p: Element, role: PlaceholderRole | undefined, ctx: OdpContext, place: ListPlace = { kind: "none" }): SlideParagraph {
  const styles = ctx.styles;
  const paraStyleName = textStyleName(p);
  const paraProps = styles.resolve(paraStyleName);
  const runs: SlideRun[] = [];
  // Inline content: text nodes, spans (their own style over the paragraph's)
  // and text:a hyperlinks. Inside a span or link any other element (a field,
  // text:s, …) contributes its text content, as the flattening always did;
  // directly under the paragraph such elements are skipped.
  const walkInline = (parent: Element, props: RunProps, link: RunLink | undefined, nested: boolean): void => {
    for (const node of Array.from(parent.childNodes)) {
      if (node.nodeType === 3) {
        const text = node.textContent ?? "";
        if (text.length > 0) {
          runs.push(runFromProps(text, props, link));
        }
      } else if (node.nodeType === 1) {
        const el = node as Element;
        if (el.namespaceURI === TEXT_NAMESPACE && el.localName === "span") {
          walkInline(el, mergeProps(props, styles.resolve(textStyleName(el))), link, true);
        } else if (el.namespaceURI === TEXT_NAMESPACE && el.localName === "a") {
          // SPEC-097: the href becomes the run link (sanitised; a dangerous,
          // empty or in-deck `#…` target keeps the plain-text reading).
          const href = linkHrefOf(xlinkAttr(el, "href"));
          const title = officeAttr(el, "title");
          walkInline(el, props, href ? { href, tooltip: title ?? undefined } : link, true);
        } else if (nested) {
          const text = el.textContent ?? "";
          if (text.length > 0) {
            runs.push(runFromProps(text, props, link));
          }
        }
      }
    }
  };
  walkInline(p, paraProps, undefined, false);
  const para: SlideParagraph = { runs };
  const align = styles.align(paraStyleName);
  if (align) para.align = align;

  // List membership (SPEC-097): a list item reads its level style; a list
  // header is explicitly unnumbered; a body-role paragraph outside any list
  // keeps the placeholder's implicit disc.
  let level: OdpListLevel | undefined;
  if (place.kind === "item") {
    level = readListLevel(place.scope.style, place.scope.depth, ctx);
    para.bullet = level.bullet;
    if (level.bulletStyle) para.bulletStyle = level.bulletStyle;
    if (level.bulletChar) para.bulletChar = level.bulletChar;
    if (level.bulletImageId) para.bulletImageId = level.bulletImageId;
    const depth = Math.min(place.scope.depth, 8);
    if (depth > 0) para.level = depth;
    // Every item of a level carries the level's start value (as every pptx
    // `a:buAutoNum` carries its startAt); only a run's first item seeds the
    // derived numbers. A per-item `text:start-value` (one of the two
    // golden-pending restart forms) is read tolerantly for its own item.
    const start = place.itemStart ?? level.startValue;
    if (start !== undefined && start !== 1) para.startAt = Math.min(start, 32767);
  } else if (place.kind === "none" && role === "body") {
    para.bullet = "disc";
  }

  // Spacing (SPEC-097), resolved along the paragraph style's parent chain.
  const lineHeight = styles.chainPick(paraStyleName, "paragraph-properties", (props) => {
    const value = foAttr(props, "line-height");
    if (value !== null) return { value, atLeast: false };
    const least = styleAttr(props, "line-height-at-least");
    return least !== null ? { value: least, atLeast: true } : undefined;
  });
  if (lineHeight) {
    const percent = lineHeight.atLeast ? undefined : odfPercentage(lineHeight.value);
    if (percent !== undefined) {
      const multiple = round3(percent / 100);
      if (multiple !== 1) para.lineSpacing = multiple;
    } else {
      const pt = odfLengthToPt(lineHeight.value);
      if (pt !== undefined && pt > 0) {
        para.lineSpacingPt = round3(pt);
        if (!lineHeight.atLeast) para.lineSpacingRule = "exact";
      }
    }
  }
  const before = spacingPt(styles.chainPick(paraStyleName, "paragraph-properties", (props) => foAttr(props, "margin-top") ?? undefined) ?? null);
  if (before) para.spaceBeforePt = before;
  const after = spacingPt(styles.chainPick(paraStyleName, "paragraph-properties", (props) => foAttr(props, "margin-bottom") ?? undefined) ?? null);
  if (after) para.spaceAfterPt = after;

  // Indents: the paragraph style's fo:margin-left / negative fo:text-indent,
  // else the list level's; stored only where they differ from slideLists'
  // effective defaults for this paragraph's level and list membership.
  const marginCm = odfLengthToCm(styles.chainPick(paraStyleName, "paragraph-properties", (props) => foAttr(props, "margin-left") ?? undefined));
  const indentCm = marginCm !== undefined ? round3(marginCm) : level?.indentCm;
  const textIndentCm = odfLengthToCm(styles.chainPick(paraStyleName, "paragraph-properties", (props) => foAttr(props, "text-indent") ?? undefined));
  const hangingCm = textIndentCm !== undefined ? (textIndentCm <= 0 ? round3(-textIndentCm) : undefined) : level?.hangingCm;
  if (indentCm !== undefined && Math.abs(indentCm - effectiveIndentCm(para)) > 0.0005) para.indentCm = indentCm;
  if (hangingCm !== undefined && Math.abs(hangingCm - effectiveHangingCm(para)) > 0.0005) para.hangingCm = hangingCm;

  // Everything else on the paragraph's OWN automatic style is preserved.
  const ownStyle = styles.raw(paraStyleName);
  const ownProps = ownStyle ? firstChildElement(ownStyle, STYLE_NAMESPACE, "paragraph-properties") : undefined;
  const preserved = ownProps ? unmodelledParaProps(ownProps) : undefined;
  if (preserved) para.odpParaProps = preserved;
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
  /** The deck's field declarations (SPEC-093), resolved per page. */
  decls: OdpDecls;
  /** Deck slide size, for the `DEFAULT_FIELD_BAND` geometry fallback. */
  deckWidthCm: number;
  deckHeightCm: number;
}

function frameBase(geom: Geometry): { id: string; xCm: number; yCm: number; widthCm: number; heightCm: number } {
  return { id: crypto.randomUUID(), xCm: geom.xCm, yCm: geom.yCm, widthCm: geom.widthCm, heightCm: geom.heightCm };
}

/** SPEC-096: the inverse of the serialiser's `draw:transform="rotate(r)"` —
 * ODF's rotate is counter-clockwise radians, so `rotationDeg =
 * normalizeRotationDeg(-r × 180/π)`. A `draw:transform` carrying anything but
 * a lone `rotate(…)` (a translate, skew or matrix) is NOT modelled: rotation
 * stays 0 and the frame parses exactly as it does today. */
function rotationOf(el: Element): number {
  const transform = drawAttr(el, "transform");
  if (!transform) {
    return 0;
  }
  const match = /^\s*rotate\s*\(\s*(-?[0-9.eE+-]+)\s*\)\s*$/.exec(transform);
  if (!match) {
    return 0;
  }
  const r = Number(match[1]);
  if (!Number.isFinite(r)) {
    return 0;
  }
  return normalizeRotationDeg((-r * 180) / Math.PI);
}

/** Stamps a parsed frame with the element's modelled rotation, if any. */
function withRotation<T extends { rotationDeg?: number }>(frame: T, el: Element): T {
  const deg = rotationOf(el);
  if (deg !== 0) {
    frame.rotationDeg = deg;
  }
  return frame;
}

/** The deck field kind a `presentation:class` denotes (SPEC-093), or undefined. */
function fieldClassKind(cls: string | null): "footer" | "slideNumber" | "date" | undefined {
  if (cls === "footer") return "footer";
  if (cls === "page-number") return "slideNumber";
  if (cls === "date-time") return "date";
  return undefined;
}

/** Parses one shape element into a frame, or null for a non-shape. */
function parseFrameChild(el: Element, ctx: OdpContext): SlideFrame | null {
  if (el.namespaceURI !== DRAW_NAMESPACE) {
    return null;
  }
  // SPEC-093: a page-level footer/page-number/date-time frame is not an
  // ordinary frame — the page's field entries are synthesised from its display
  // flags + declarations (odpPageFields), so drop it here.
  if (el.localName === "frame" && fieldClassKind(presentationAttr(el, "class")) !== undefined) {
    return null;
  }
  switch (el.localName) {
    case "a":
      return parseFrameLink(el, ctx);
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

/** The one drawing element a `draw:a` wraps, or undefined when it wraps
 * anything else (none, several, or non-draw content). */
function linkedShapeOf(a: Element): Element | undefined {
  const children = Array.from(a.children);
  return children.length === 1 && children[0].namespaceURI === DRAW_NAMESPACE ? children[0] : undefined;
}

/** A `draw:a` wrapping one shape element (SPEC-097): the shape parses as
 * usual and, for a linkable kind (text/image/shape — isLinkableKind), carries
 * the sanitised `xlink:href` as its frame link (+ `office:title` tooltip). A
 * dangerous href is stripped (the frame reads unlinked). An in-deck `#…`
 * target (the golden-pending slide-jump form), a non-linkable child (group,
 * table, chart, media) or an unexpected shape of wrapper keeps the whole
 * `draw:a` verbatim as a foreign frame — today's reading — so its markup
 * survives. */
function parseFrameLink(a: Element, ctx: OdpContext): SlideFrame | null {
  const inner = linkedShapeOf(a);
  const base = frameBase((inner && geometryOf(inner)) ?? { xCm: 0, yCm: 0, widthCm: 0, heightCm: 0 });
  const raw = xlinkAttr(a, "href");
  if (!inner || (raw !== null && raw.trim().startsWith("#"))) {
    return foreignAt(a, base, "Object");
  }
  const frame = parseFrameChild(inner, ctx);
  if (!frame) {
    return null;
  }
  if (!isLinkableKind(frame.kind)) {
    return foreignAt(a, base, "Object");
  }
  const href = linkHrefOf(raw);
  if (href) {
    frame.linkHref = href;
    const title = officeAttr(a, "title");
    if (title) frame.linkTooltip = title;
  }
  return frame;
}

/** The text-body graphic properties SPEC-097 does not model, preserved in
 * `odpBodyProps` in this fixed order (`draw:textarea-vertical-align` only for
 * a value outside top/middle/bottom). The vertical-text frame's
 * `style:writing-mode` rides here, as do both golden-pending shrink-on-
 * overflow candidates (`draw:fit-to-size`, `style:shrink-to-fit`). */
const PRESERVED_BODY_ATTRS: ReadonlyArray<readonly [string, string, string]> = [
  [STYLE_NAMESPACE, "writing-mode", "style:writing-mode"],
  [DRAW_NAMESPACE, "fit-to-size", "draw:fit-to-size"],
  [STYLE_NAMESPACE, "shrink-to-fit", "style:shrink-to-fit"],
  [DRAW_NAMESPACE, "auto-grow-width", "draw:auto-grow-width"],
  [DRAW_NAMESPACE, "textarea-horizontal-align", "draw:textarea-horizontal-align"],
  [DRAW_NAMESPACE, "textarea-vertical-align", "draw:textarea-vertical-align"],
];

const ANCHORS = new Set(["top", "middle", "bottom"]);

/** A text/shape frame's text-body properties (SPEC-097), read from its
 * graphic style along the `draw:style-name` parent chain: fo:padding-* (or
 * the fo:padding shorthand) → insetsCm when any side differs from the pptx
 * defaults, draw:textarea-vertical-align → anchor when not the kind's default,
 * draw:auto-grow-height="true" → "resizeShape", fo:wrap-option="no-wrap" →
 * wrap false, style:columns → columns; the unmodelled rest → odpBodyProps. */
function readTextBody(el: Element, kind: "text" | "shape", ctx: OdpContext): { body?: TextBodyProps; odpBodyProps?: string } {
  const name = drawAttr(el, "style-name");
  if (!name) {
    return {};
  }
  const pick = <T>(fn: (props: Element) => T | undefined): T | undefined => ctx.styles.chainPick(name, "graphic-properties", fn);
  const body: TextBodyProps = {};
  const side = (edge: string): number | undefined =>
    odfLengthToCm(pick((props) => foAttr(props, `padding-${edge}`) ?? foAttr(props, "padding") ?? undefined));
  const left = side("left");
  const top = side("top");
  const right = side("right");
  const bottom = side("bottom");
  if (left !== undefined || top !== undefined || right !== undefined || bottom !== undefined) {
    const insets = {
      left: round3(left ?? DEFAULT_INSETS_CM.left),
      top: round3(top ?? DEFAULT_INSETS_CM.top),
      right: round3(right ?? DEFAULT_INSETS_CM.right),
      bottom: round3(bottom ?? DEFAULT_INSETS_CM.bottom),
    };
    if (
      insets.left !== DEFAULT_INSETS_CM.left ||
      insets.top !== DEFAULT_INSETS_CM.top ||
      insets.right !== DEFAULT_INSETS_CM.right ||
      insets.bottom !== DEFAULT_INSETS_CM.bottom
    ) {
      body.insetsCm = insets;
    }
  }
  const anchor = pick((props) => drawAttr(props, "textarea-vertical-align") ?? undefined);
  if (anchor && ANCHORS.has(anchor) && anchor !== (kind === "shape" ? "middle" : "top")) {
    body.anchor = anchor as NonNullable<TextBodyProps["anchor"]>;
  }
  if (pick((props) => drawAttr(props, "auto-grow-height") ?? undefined) === "true") {
    body.autofit = "resizeShape";
  }
  if (pick((props) => foAttr(props, "wrap-option") ?? undefined) === "no-wrap") {
    body.wrap = false;
  }
  const columns = pick((props) => firstChildElement(props, STYLE_NAMESPACE, "columns"));
  if (columns) {
    const count = Number(foAttr(columns, "column-count"));
    const gapCm = odfLengthToCm(foAttr(columns, "column-gap")) ?? 0;
    if (Number.isInteger(count) && count >= 1 && (count > 1 || gapCm > 0)) {
      body.columns = { count: Math.min(count, 16), gapCm: round3(Math.max(0, gapCm)) };
    }
  }
  const preserved: string[] = [];
  for (const [ns, local, qname] of PRESERVED_BODY_ATTRS) {
    const value = pick((props) => props.getAttributeNS(ns, local) ?? undefined);
    if (value === undefined || (local === "textarea-vertical-align" && ANCHORS.has(value))) {
      continue;
    }
    preserved.push(`${qname}="${escapeXmlAttribute(value)}"`);
  }
  const out: { body?: TextBodyProps; odpBodyProps?: string } = {};
  if (Object.keys(body).length > 0) out.body = body;
  if (preserved.length > 0) out.odpBodyProps = preserved.join(" ");
  return out;
}

function graphicFillColor(el: Element, ctx: OdpContext): ModelColor | undefined {
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

  // A `draw:plugin` (or a media-typed `draw:object`) is an audio/video clip
  // (SPEC-095), checked before the chart-object branch so a media object is
  // never mistaken for a chart.
  const media = parseMediaFrame(frame, base, ctx);
  if (media) {
    return withRotation(media, frame);
  }

  // A `draw:object` (an embedded chart) is checked before `draw:image` so a
  // chart's sibling `ObjectReplacements/…` preview image is never mistaken for
  // the frame's picture (SPEC-092 AC-9a): it stays an untouched preserved part.
  const object = firstChildElement(frame, DRAW_NAMESPACE, "object");
  if (object) {
    const chart = parseChartObject(frame, object, base, ctx);
    if (chart) {
      return chart;
    }
    return foreignAt(frame, base, "Object"); // unresolvable object (AC-10)
  }

  const image = firstChildElement(frame, DRAW_NAMESPACE, "image");
  if (image) {
    const href = xlinkAttr(image, "href");
    const partPath = href ? normaliseHref(href) : undefined;
    if (partPath && collectMedia(partPath, ctx)) {
      const out: SlideFrame = { ...base, kind: "image", mediaId: partPath };
      const stroke = ctx.styles.stroke(drawAttr(frame, "style-name"));
      if (stroke) out.stroke = stroke;
      return withRotation(out, frame);
    }
    return foreignAt(frame, base, "Media");
  }

  const table = firstChildElement(frame, TABLE_NAMESPACE, "table");
  if (table) {
    const parsed = parseOdpTable(table, base, frame, ctx);
    if (parsed) {
      return withRotation(parsed, frame);
    }
    return foreignAt(frame, base, "Table"); // out of subset (AC-3b)
  }

  const textBox = firstChildElement(frame, DRAW_NAMESPACE, "text-box");
  if (textBox) {
    const paragraphs = parseTextBox(textBox, role, ctx);
    const out: SlideFrame = { ...base, kind: "text", paragraphs };
    if (role) out.role = role;
    const fill = graphicFillColor(frame, ctx);
    if (fill) out.fill = { type: "solid", color: fill };
    const stroke = ctx.styles.stroke(drawAttr(frame, "style-name"));
    if (stroke) out.stroke = stroke;
    const { body, odpBodyProps } = readTextBody(frame, "text", ctx);
    if (body) out.body = body;
    if (odpBodyProps) out.odpBodyProps = odpBodyProps;
    return withRotation(out, frame);
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

function parseOdpCells(row: Element, ctx: OdpContext): SlideTableCell[] {
  const cells: SlideTableCell[] = [];
  for (const el of Array.from(row.children)) {
    if (el.namespaceURI !== TABLE_NAMESPACE) {
      continue;
    }
    if (el.localName === "covered-table-cell") {
      cells.push({ paragraphs: [], covered: true });
    } else if (el.localName === "table-cell") {
      const cell: SlideTableCell = { paragraphs: parseTextBox(el, undefined, ctx) };
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
  return cells;
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
          const row: SlideTableRow = { cells: parseOdpCells(el, ctx) };
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
  if (rows.length === 0) {
    return undefined;
  }
  // With declared table-columns the grid must match them; with none (some
  // producers omit them) infer the count from the first row and pad the widths
  // with an equal share of the frame width (the modelled widths matter only
  // once the table is edited).
  let colCount = columnWidthsCm.length;
  if (colCount === 0) {
    colCount = rows[0].cells.length;
    while (columnWidthsCm.length < colCount) {
      columnWidthsCm.push(colCount > 0 ? round3(base.widthCm / colCount) : 0);
      columnStyleNames.push(undefined);
    }
  }
  if (colCount === 0 || rows.some((r) => r.cells.length !== colCount)) {
    return undefined; // irregular grid (AC-3b)
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
  const ctx: OdpContext = {
    styles,
    masterPh: new Map(),
    media: new Map(),
    parts: new Map(),
    decls: { footers: new Map(), dateTimes: new Map() },
    deckWidthCm: DEFAULT_SLIDE_WIDTH_CM,
    deckHeightCm: DEFAULT_SLIDE_HEIGHT_CM,
  };
  const frameEl = doc.documentElement.getElementsByTagNameNS(DRAW_NAMESPACE, "frame")[0];
  const frame = parseDrawFrame(frameEl, ctx);
  return frame.kind === "table" ? frame : undefined;
}

/** Test seam (SPEC-092): parses a chart `draw:frame` against a supplied
 * `{ partPath → bytes }` parts map. */
export function parseChartDrawFrameForTest(drawFrameXml: string, parts: Record<string, Uint8Array>): SlideFrame {
  const wrapped =
    `<office:document-content ${ODP_TEST_NS} xmlns:xlink="http://www.w3.org/1999/xlink">` +
    `<office:body><office:presentation><draw:page>${drawFrameXml}</draw:page></office:presentation></office:body>` +
    `</office:document-content>`;
  const doc = parseXml(wrapped);
  const ctx: OdpContext = {
    styles: new StyleResolver(),
    masterPh: new Map(),
    media: new Map(),
    parts: new Map(Object.entries(parts)),
    decls: { footers: new Map(), dateTimes: new Map() },
    deckWidthCm: DEFAULT_SLIDE_WIDTH_CM,
    deckHeightCm: DEFAULT_SLIDE_HEIGHT_CM,
  };
  const frameEl = doc.documentElement.getElementsByTagNameNS(DRAW_NAMESPACE, "frame")[0];
  return parseDrawFrame(frameEl, ctx);
}

/** Test seam (SPEC-095 golden): parses a standalone media `draw:frame` against
 * a supplied `{ partPath → bytes }` parts map, so the golden can assert
 * draw:plugin → MediaFrame, a tolerant variant, and preserve-don't-parse. */
export function parseMediaFrameForTest(drawFrameXml: string, parts: Record<string, Uint8Array> = {}): SlideFrame {
  const wrapped =
    `<office:document-content ${ODP_TEST_NS} xmlns:xlink="http://www.w3.org/1999/xlink">` +
    `<office:body><office:presentation><draw:page>${drawFrameXml}</draw:page></office:presentation></office:body>` +
    `</office:document-content>`;
  const doc = parseXml(wrapped);
  const ctx: OdpContext = {
    styles: new StyleResolver(),
    masterPh: new Map(),
    media: new Map(),
    parts: new Map(Object.entries(parts)),
    decls: { footers: new Map(), dateTimes: new Map() },
    deckWidthCm: DEFAULT_SLIDE_WIDTH_CM,
    deckHeightCm: DEFAULT_SLIDE_HEIGHT_CM,
  };
  const frameEl = doc.documentElement.getElementsByTagNameNS(DRAW_NAMESPACE, "frame")[0];
  return parseDrawFrame(frameEl, ctx);
}

/** Test seam (SPEC-096 golden): parses one standalone shape/group element
 * (`draw:g`, `draw:custom-shape`, `draw:rect`…) into a SlideFrame, so the
 * goldens can assert the reader accepts the canonical bytes plus
 * attribute-reordered variants. `styles` is optional automatic-styles markup
 * (e.g. the golden's `de-G1`) resolved for fills/strokes. */
export function parseFrameChildForTest(elementXml: string, styles = ""): SlideFrame | null {
  const wrapped =
    `<office:document-content ${ODP_TEST_NS} xmlns:xlink="http://www.w3.org/1999/xlink">` +
    `<office:automatic-styles>${styles}</office:automatic-styles>` +
    `<office:body><office:presentation><draw:page>${elementXml}</draw:page></office:presentation></office:body>` +
    `</office:document-content>`;
  const doc = parseXml(wrapped);
  const resolver = new StyleResolver();
  addStylesFrom(doc.documentElement, resolver);
  const ctx: OdpContext = {
    styles: resolver,
    masterPh: new Map(),
    media: new Map(),
    parts: new Map(),
    decls: { footers: new Map(), dateTimes: new Map() },
    deckWidthCm: DEFAULT_SLIDE_WIDTH_CM,
    deckHeightCm: DEFAULT_SLIDE_HEIGHT_CM,
  };
  const pages = doc.documentElement.getElementsByTagNameNS(DRAW_NAMESPACE, "page");
  const first = pages[0]?.firstElementChild;
  return first ? parseFrameChild(first, ctx) : null;
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
  const frame = shapeFrame(el, base, geometry, ctx);
  // SPEC-096: `draw:modifiers` parses back to the stored adjustment fractions
  // (the serialiser's own 100000ths convention — see modifiersAttr).
  const modifiers = enhanced ? drawAttr(enhanced, "modifiers") : null;
  if (modifiers && frame.kind === "shape") {
    const adjustments = modifiers
      .trim()
      .split(/\s+/)
      .map(Number)
      .filter((n) => Number.isFinite(n))
      .map(guideValueToAdjustment);
    if (adjustments.length > 0) {
      frame.adjustments = adjustments;
    }
  }
  return frame;
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
  const frame: ShapeFrame = { ...base, kind: "shape", geometry };
  const fill = graphicFillColor(el, ctx);
  if (fill) frame.fill = { type: "solid", color: fill };
  const stroke = ctx.styles.stroke(drawAttr(el, "style-name"));
  if (stroke) frame.stroke = stroke;
  const textBox = firstChildElement(el, DRAW_NAMESPACE, "text-box") ?? el;
  const paragraphs = parseTextBox(textBox, undefined, ctx);
  if (paragraphs.length > 0) frame.paragraphs = paragraphs;
  const { body, odpBodyProps } = readTextBody(el, "shape", ctx);
  if (body) frame.body = body;
  if (odpBodyProps) frame.odpBodyProps = odpBodyProps;
  return withRotation(frame, el);
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

/** Resolves a `draw:frame`'s `draw:object` reference — `xlink:href="./Object 1"`
 * to `Object 1/content.xml` in the package — into a `ChartFrame` (SPEC-092
 * AC-9). A parseable in-catalogue object becomes a painted chart; a valid
 * `office:chart` outside the catalogue becomes an unsupported placeholder
 * carrying the verbatim `content.xml` (AC-9b); an unresolvable href, or a
 * resolved part that is not an `office:chart` at all, returns null so the
 * caller keeps today's "Object" ForeignFrame (AC-10). `xml` retains the source
 * `draw:frame` so an untouched chart re-emits byte-identically. */
function parseChartObject(
  frame: Element,
  object: Element,
  base: ReturnType<typeof frameBase>,
  ctx: OdpContext,
): SlideFrame | null {
  const href = xlinkAttr(object, "href");
  if (!href) {
    return null;
  }
  const objectDir = href.replace(/^\.\//, "").replace(/\/$/, "");
  const partPath = `${objectDir}/content.xml`;
  const bytes = ctx.parts.get(partPath);
  if (!bytes) {
    return null;
  }
  const xmlText = new TextDecoder("utf-8").decode(bytes);
  const parsed = parseOdfChartXml(xmlText);
  if (parsed) {
    return { ...base, kind: "chart", ...parsed, sourcePart: partPath, xml: serializeElement(frame) };
  }
  if (!isOfficeChart(xmlText)) {
    return null; // not a chart object at all — foreign "Object" (AC-10)
  }
  return {
    ...base,
    kind: "chart",
    chartType: "columnClustered",
    categories: [],
    series: [],
    options: { legend: "none", dataLabels: false, gridlines: false },
    unsupportedChartXml: xmlText,
    sourcePart: partPath,
    xml: serializeElement(frame),
  };
}

/** True when `xmlText` is an `office:document-content` holding an
 * `office:chart`/`chart:chart` — the test separating AC-9b (out-of-catalogue
 * chart) from AC-10 (not a chart object). */
function isOfficeChart(xmlText: string): boolean {
  let doc: Document;
  try {
    doc = parseXml(xmlText);
  } catch (err) {
    if (err instanceof XmlParseError) {
      return false;
    }
    throw err;
  }
  const body = firstChildElement(doc.documentElement, OFFICE_NAMESPACE, "body");
  const officeChart = body && firstChildElement(body, OFFICE_NAMESPACE, "chart");
  return !!officeChart && !!firstChildElement(officeChart, CHART_NAMESPACE, "chart");
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

/** The `draw:plugin` or media-typed `draw:object` inside a `draw:frame`, else
 * undefined (SPEC-095). A `draw:object` is media only when its `draw:mime-type`
 * is `video/*` / `audio/*`, so a chart object is never caught here. */
function mediaChildOf(frame: Element): { el: Element; mimeType: string | undefined } | undefined {
  const plugin = firstChildElement(frame, DRAW_NAMESPACE, "plugin");
  if (plugin) {
    return { el: plugin, mimeType: drawAttr(plugin, "mime-type") ?? undefined };
  }
  const object = firstChildElement(frame, DRAW_NAMESPACE, "object");
  const mime = object ? drawAttr(object, "mime-type") : null;
  if (object && mime && (mime.startsWith("video/") || mime.startsWith("audio/"))) {
    return { el: object, mimeType: mime };
  }
  return undefined;
}

/** A `draw:frame` carrying `draw:plugin`/media `draw:object` becomes a
 * MediaFrame (AC-5/6): `xlink:href` through `normaliseHref`, `draw:mime-type` →
 * contentType, `presentation:play-full-screen` → fullScreen, the `draw:param`
 * rows → playback flags, a sibling `draw:image` → posterMediaId; the verbatim
 * source is kept for byte-exact re-emit. Undefined for a non-media frame. */
function parseMediaFrame(
  frame: Element,
  base: ReturnType<typeof frameBase>,
  ctx: OdpContext,
): MediaFrame | undefined {
  const media = mediaChildOf(frame);
  if (!media) {
    return undefined;
  }
  const href = xlinkAttr(media.el, "href");
  const partPath = href ? normaliseHref(href) : undefined;
  const contentType = media.mimeType ?? (partPath ? contentTypeForPartName(partPath) : undefined);
  const kind: SlideMediaKind = mediaKindForContentType(contentType);
  const out: MediaFrame = { ...base, kind: "media", media: kind };

  const raw = partPath ? ctx.parts.get(partPath) : undefined;
  if (partPath && raw) {
    if (raw.length > MAX_MEDIA_SIZE_BYTES) {
      out.oversize = true;
    } else {
      if (!ctx.media.has(partPath)) {
        ctx.media.set(partPath, {
          mediaId: partPath,
          contentType: contentType ?? "application/octet-stream",
          data: base64Sync(raw),
        });
      }
      out.mediaId = partPath;
    }
  } else if (href) {
    // Resolves outside the package — a linked clip, never fetched (AC-32).
    out.linkTarget = href;
  }
  if (contentType) {
    out.contentType = contentType;
  }
  if (presentationAttr(frame, "play-full-screen") === "true") {
    out.fullScreen = true;
  }
  for (const param of childElements(media.el, DRAW_NAMESPACE, "param")) {
    const name = drawAttr(param, "name");
    const value = drawAttr(param, "value");
    if (name === "Loop" && value === "true") {
      out.loop = true;
    } else if (name === "Mute" && value === "true") {
      out.muted = true;
    } else if (name === "AutoPlay" && value === "true") {
      out.autoStart = true;
    } else if (name === "Volume" && value) {
      const n = Number(value);
      if (Number.isFinite(n)) {
        out.volume = n;
      }
    } else if (name === "VolumeDB" && value && out.volume === undefined) {
      // Impress emits decibels (0 = full, negative = quieter); map to 0..100.
      const db = Number(value);
      if (Number.isFinite(db)) {
        out.volume = Math.max(0, Math.min(100, Math.round(100 * 10 ** (db / 20))));
      }
    }
  }

  const posterImage = firstChildElement(frame, DRAW_NAMESPACE, "image");
  if (posterImage) {
    const posterHref = xlinkAttr(posterImage, "href");
    const posterPath = posterHref ? normaliseHref(posterHref) : undefined;
    if (posterPath && collectMedia(posterPath, ctx)) {
      out.posterMediaId = posterPath;
    }
  }

  const stroke = ctx.styles.stroke(drawAttr(frame, "style-name"));
  if (stroke) {
    out.stroke = stroke;
  }
  out.sourceXml = serializeElement(frame);
  return out;
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

  // SPEC-094: the deck theme rides styles.xml's `de-theme-*` block.
  const theme = stylesBytes ? readOdpTheme(decodeUtf8(stylesBytes), DEFAULT_THEME) : undefined;

  const masterPh = buildMasterPlaceholders(stylesDoc);
  const { widthCm, heightCm } = pageSize(stylesDoc, contentDoc);
  const decls = readDeclarations(presentation);

  const media = new Map<string, MediaPart>();
  const ctx: OdpContext = { styles, masterPh, media, parts, decls, deckWidthCm: widthCm, deckHeightCm: heightCm };

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

  const headerFooter = assembleOdpHeaderFooter(slides, decls);

  const parsed: ParsedPresentation = {
    slides,
    layouts,
    slideWidthCm: widthCm,
    slideHeightCm: heightCm,
    headerFooter,
    media: [...media.values()],
    opaque,
    properties,
    odpContentEnvelope: { prefix, suffix },
  };
  if (theme) {
    parsed.theme = theme;
  }
  return parsed;
}

/** Assembles the deck-level `headerFooter` from the odp declarations + read
 * slides (SPEC-093): the footer default is the first footer decl (divergent
 * pages keep a per-slide override); the date default comes from the auto date
 * decl's data style (or a deck-wide fixed decl when no auto one exists). */
function assembleOdpHeaderFooter(slides: Slide[], decls: OdpDecls): DeckHeaderFooter | undefined {
  const hf: DeckHeaderFooter = {};
  const firstFooter = [...decls.footers.values()][0];
  if (firstFooter !== undefined && firstFooter !== "") {
    hf.footerText = firstFooter;
    for (const slide of slides) {
      for (const field of slide.fields ?? []) {
        if (field.kind === "footer" && field.text === firstFooter) {
          delete field.text;
        }
      }
    }
  }
  const dates = [...decls.dateTimes.values()];
  const autoDecl = dates.find((decl) => decl.source === "current-date");
  const fixedDecl = dates.find((decl) => decl.source === "fixed");
  if (autoDecl) {
    const format = autoDecl.dataStyleName ? formatFromOdtDateStyleName(autoDecl.dataStyleName) : null;
    if (format && format !== DEFAULT_DATE_FORMAT) {
      hf.dateFormat = format;
    }
  } else if (fixedDecl) {
    hf.dateSource = "fixed";
    if (fixedDecl.text) {
      hf.dateFixedText = fixedDecl.text;
    }
    for (const slide of slides) {
      for (const field of slide.fields ?? []) {
        if (field.kind === "date" && field.text === fixedDecl.text) {
          delete field.text;
        }
      }
    }
  }
  return Object.keys(hf).length > 0 ? hf : undefined;
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
      // A linked frame's wire id sits on the shape inside its draw:a.
      const idEl = child.namespaceURI === DRAW_NAMESPACE && child.localName === "a" ? (linkedShapeOf(child) ?? child) : child;
      const wireId = idEl.getAttributeNS(XML_NAMESPACE, "id") ?? drawAttr(idEl, "id");
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
  const fields = odpPageFields(page, ctx);
  if (fields.length > 0) {
    slide.fields = fields;
  }
  return slide;
}

/** The deck field placeholders a page shows (SPEC-093): synthesised from its
 * `presentation:display-*` flags + `use-*-name` declarations, geometry from the
 * master placeholders (`OdpContext.masterPh`) with a `DEFAULT_FIELD_BAND`
 * fallback. Presence in the returned list IS the page's visibility. */
function odpPageFields(page: Element, ctx: OdpContext): SlideFieldPlaceholder[] {
  const fields: SlideFieldPlaceholder[] = [];
  const display = ctx.styles.drawingPageDisplay(drawAttr(page, "style-name"));
  const rect = (kind: SlideFieldPlaceholder["kind"], cls: string): Pick<SlideFieldPlaceholder, "xCm" | "yCm" | "widthCm" | "heightCm"> => {
    const master = ctx.masterPh.get(cls);
    if (master) {
      return { xCm: master.xCm, yCm: master.yCm, widthCm: master.widthCm, heightCm: master.heightCm };
    }
    const band = defaultFieldRect(kind, ctx.deckWidthCm, ctx.deckHeightCm);
    return { xCm: band.xCm, yCm: band.yCm, widthCm: band.widthCm, heightCm: band.heightCm };
  };
  const useFooter = presentationAttr(page, "use-footer-name");
  if (useFooter && display.footer !== false && ctx.masterPh.has("footer")) {
    const entry: SlideFieldPlaceholder = { kind: "footer", ...rect("footer", "footer"), align: "center" };
    const text = ctx.decls.footers.get(useFooter);
    if (text) {
      entry.text = text;
    }
    fields.push(entry);
  }
  if (display.pageNumber !== false && ctx.masterPh.has("page-number")) {
    fields.push({ kind: "slideNumber", ...rect("slideNumber", "page-number"), align: "right" });
  }
  const useDate = presentationAttr(page, "use-date-time-name");
  if (useDate && display.dateTime !== false && ctx.masterPh.has("date-time")) {
    // The date sits left in the band (left is the default alignment → absent).
    const entry: SlideFieldPlaceholder = { kind: "date", ...rect("date", "date-time") };
    const decl = ctx.decls.dateTimes.get(useDate);
    if (decl) {
      if (decl.source === "fixed") {
        // A fixed date on this page — a per-slide literal (the deck stays auto
        // unless every page is fixed, folded to the deck in assembleOdpHeaderFooter).
        if (decl.text) {
          entry.text = decl.text;
        }
      } else if (decl.dataStyleName && !formatFromOdtDateStyleName(decl.dataStyleName)) {
        // A foreign data style (Impress's D1) is out-of-subset — preserve it,
        // render the default pattern (ODF caches nothing).
        entry.wireStyleName = decl.dataStyleName;
      }
    }
    fields.push(entry);
  }
  return fields;
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
      } else if (styleEl.namespaceURI === TEXT_NAMESPACE && styleEl.localName === "list-style") {
        resolver.addListStyle(styleEl);
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

/** The slide size from the master page's `style:page-layout`
 * (`fo:page-width`/`fo:page-height`), else the default. */
function pageSize(stylesDoc: Document | undefined, contentDoc: Document): { widthCm: number; heightCm: number } {
  return (
    search(stylesDoc) ?? search(contentDoc) ?? { widthCm: DEFAULT_SLIDE_WIDTH_CM, heightCm: DEFAULT_SLIDE_HEIGHT_CM }
  );
}
