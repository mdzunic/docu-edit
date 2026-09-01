// The `.pptx` -> PresentationModel pipeline (SPEC-079): unzip the OPC
// package, follow the root relationships to the presentation part (never
// assuming its path), read the slide size, and walk `p:sldIdLst` to each
// slide part. Each slide is resolved into a WYSIWYG `Slide` — frames with
// effective geometry (placeholder inheritance through layout → master),
// effective text formatting (run → shape lstStyle → master txStyles → theme
// defaults), resolved scheme colours and background — with the whole package
// retained verbatim in `opaque` so a no-edit download stays byte-identical.
// Reading is as tolerant as the preparation reader: an unreadable slide still
// yields an (empty) slide.
import { bytesToBase64Async, type OpaquePart } from "../../core/packageParts";
import { detectPackage } from "../../core/packageDetect";
import { parseCoreProperties, parseCustomProperties } from "../../core/opcProperties";
import type { DocumentProperties } from "../../core/documentProperties";
import type { MediaPart } from "../../core/media";
import { contentTypeForPartName, MAX_MEDIA_SIZE_BYTES } from "../../core/media";
import { mediaKindForContentType } from "../mediaFrames";
import {
  AUDIO_RELATIONSHIP,
  DE_MEDIA_NAMESPACE,
  MEDIA_RELATIONSHIP,
  P14_NAMESPACE,
  VIDEO_RELATIONSHIP,
} from "./mediaWire";
import { cmToEmu, emuToCm } from "../../core/units";
import { unzip, ZipError } from "../../core/zip";
import {
  DRAWINGML_NAMESPACE,
  firstA,
  childrenA,
  parseXfrm,
  parseAvLstAdjustments,
  parseGroupChildMapping,
  mapGroupChild,
  PRST_TO_GEOMETRY,
  type DrawingXfrm,
  type GroupChildMapping,
} from "../../core/drawingMl";
import {
  childElements,
  escapeXmlAttribute,
  firstChildElement,
  parseRelationships,
  parseXml,
  relsPathFor,
  resolvePartPath,
  serializeElement,
  XmlParseError,
  type Relationship,
} from "../../core/xml";
import {
  DEFAULT_SLIDE_HEIGHT_CM,
  DEFAULT_SLIDE_WIDTH_CM,
  DEFAULT_TABLE_CELL_PADDING_CM,
  type DeckHeaderFooter,
  type DeckTheme,
  type LayoutSlot,
  type ParsedPresentation,
  type PlaceholderRole,
  type Slide,
  type SlideFieldKind,
  type SlideFieldPlaceholder,
  type SlideFill,
  type MediaFrame,
  type SlideFrame,
  type SlideLayout,
  type SlideMediaKind,
  type SlideParagraph,
  type SlideRun,
  type SlideTableCell,
  type SlideTableRow,
  type TableFrame,
} from "../model";
import { DEFAULT_DATE_FORMAT, defaultFieldRect } from "../headerFooter";
import { KIND_OF_PH_TYPE, PPTX_DATETIME_TO_FORMAT } from "./fields";
import { computeOwnerGrid, type BorderSide, type CellBorders } from "../../core/tables";
import { CHART_RELATIONSHIP, C_NAMESPACE, parseChartXml } from "../../core/chartPartOoxml";
import { parseClrMap, parseClrScheme, resolveFirstColor, type ColorContext } from "./color";
import { colorRefFrom, fontRefFrom, readDeckTheme } from "./theme";
import { DEFAULT_THEME, type ModelColor, type ModelFont } from "../../core/theme";
import { classifyLayout } from "../layouts";
import { readPptxTransition } from "../transitions";
import { readPptxTiming } from "./timing";

export class PptxOpenError extends Error {}

export const INVALID_PPTX_MESSAGE = "This file isn't a PowerPoint presentation we can open";

/** The main part's declared content type that makes an OPC package a
 * presentation (ECMA-376 PresentationML). */
export const PPTX_MAIN_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

/** The file's MIME type, for downloads. */
export const PPTX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

const P_NAMESPACE = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NAMESPACE = DRAWINGML_NAMESPACE;
const R_ATTRIBUTE_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SLIDE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_LAYOUT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const SLIDE_MASTER_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const THEME_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const IMAGE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const NOTES_SLIDE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const CORE_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const CUSTOM_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

const DEFAULT_FONT_SIZE_PT = 18;
const DEFAULT_COLOR = "#000000";

function decodeUtf8(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

function pChild(parent: Element, localName: string): Element | undefined {
  return firstChildElement(parent, P_NAMESPACE, localName);
}

/** The placeholder role a `p:ph@type` denotes. Absent type = a body/object
 * placeholder. */
function roleOf(type: string | undefined): PlaceholderRole {
  if (type === "title" || type === "ctrTitle") {
    return "title";
  }
  if (type === "subTitle") {
    return "subtitle";
  }
  if (type === undefined || type === "body") {
    return "body";
  }
  return "other";
}

/** The (type+idx) key a placeholder inherits geometry/style by (AC-9): title
 * and ctrTitle share a key so a slide's `ctrTitle` still matches a layout's
 * `title`. Exported for `layouts.ts`'s exact-match pass (SPEC-089). */
export function phKey(type: string | undefined, idx: string | undefined): string {
  const normalisedType = type === "ctrTitle" ? "title" : type ?? "body";
  return `${normalisedType}#${idx ?? ""}`;
}

interface PlaceholderInfo {
  xfrm?: DrawingXfrm;
  lstStyle?: Element;
  role: PlaceholderRole;
}

/** A resolved run-property set (partial: only the fields a source declares).
 * SPEC-094: colour/font may be a scheme reference when the deck has a theme. */
interface RunProps {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: ModelFont;
  sizePt?: number;
  color?: ModelColor;
}

/** Everything a slide's frames resolve their effective values against. */
interface SlideContext {
  color: ColorContext;
  /** Present when the deck's theme was captured AND this slide's chain ends in
   * that same theme part: `a:schemeClr`/`+mj-lt`/`+mn-lt` then read as
   * references through this clrMap (SPEC-094 AC-14/16). Absent = literals. */
  themeCapture?: { map: Record<string, string> };
  layoutPh: Map<string, PlaceholderInfo>;
  masterPh: Map<string, PlaceholderInfo>;
  /** Master `p:txStyles`: role → level (0-based) → defRPr element. */
  txStyles: Map<PlaceholderRole, Element[]>;
  themeMinorLatin?: string;
  slideRels: Map<string, Relationship>;
  slidePartPath: string;
  media: Map<string, MediaPart>;
  /** Every package part's raw bytes — the chart-part lookup (SPEC-092 AC-7a)
   *  resolves a `p:graphicFrame`'s `c:chart r:id` through here. */
  parts: Map<string, Uint8Array>;
}

/** The first colour element (`a:srgbClr`/`a:sysClr`/`a:schemeClr`) inside a
 * fill container, or undefined. */
function firstColorElement(container: Element): Element | undefined {
  for (const child of Array.from(container.children)) {
    if (child.namespaceURI === A_NAMESPACE && ["srgbClr", "sysClr", "schemeClr"].includes(child.localName)) {
      return child;
    }
  }
  return undefined;
}

/** Reads a fill container's solid colour: SPEC-079's resolution stays the
 * normative effective value, and when the deck has a captured theme an
 * `a:schemeClr` additionally records the reference (SPEC-094 AC-14) — any
 * transform outside lumMod/lumOff/shade/tint keeps it a literal (AC-15). */
function readSolidColor(
  container: Element,
  color: ColorContext,
  capture: { map: Record<string, string> } | undefined,
): ModelColor | undefined {
  const resolved = resolveFirstColor(container, color);
  if (!capture) {
    return resolved;
  }
  const el = firstColorElement(container);
  if (el && el.localName === "schemeClr") {
    const ref = colorRefFrom(el, capture.map);
    if (ref) {
      return ref;
    }
  }
  return resolved;
}

/** Reads a run/defRPr-shaped element's own declared properties. */
function readRunProps(
  rPr: Element | undefined,
  color: ColorContext,
  capture?: { map: Record<string, string> },
): RunProps {
  if (!rPr) {
    return {};
  }
  const props: RunProps = {};
  const b = rPr.getAttribute("b");
  if (b !== null) {
    props.bold = b === "1" || b === "true";
  }
  const i = rPr.getAttribute("i");
  if (i !== null) {
    props.italic = i === "1" || i === "true";
  }
  const u = rPr.getAttribute("u");
  if (u !== null) {
    props.underline = u !== "none";
  }
  const sz = rPr.getAttribute("sz");
  if (sz !== null && Number.isFinite(Number(sz))) {
    props.sizePt = Number(sz) / 100;
  }
  const latin = firstA(rPr, "latin");
  const typeface = latin?.getAttribute("typeface");
  if (typeface) {
    // `+mj-lt`/`+mn-lt` become theme font references on a themed deck
    // (SPEC-094 AC-16); on an unthemed one the baked minor-latin literal
    // stands in (nothing to reference).
    const ref = capture ? fontRefFrom(typeface) : undefined;
    props.fontFamily = ref ?? typeface;
  }
  const solidFill = firstA(rPr, "solidFill");
  if (solidFill) {
    const value = readSolidColor(solidFill, color, capture);
    if (value) {
      props.color = value;
    }
  }
  return props;
}

/** Merges a lower-precedence property set under a higher one (higher wins for
 * any field it declares). */
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

/** The `a:defRPr` of a level in an `a:lstStyle` (`a:lvl{n}pPr/a:defRPr`). */
function lstStyleDefRPr(lstStyle: Element | undefined, level: number): Element | undefined {
  if (!lstStyle) {
    return undefined;
  }
  const lvl = firstA(lstStyle, `lvl${level + 1}pPr`);
  return lvl ? firstA(lvl, "defRPr") : undefined;
}

/** Resolves one run's effective formatting through the full chain
 * (AC-13..16): run rPr → shape lstStyle → layout/master placeholder chain →
 * theme minor latin + built-in defaults. */
function resolveRun(
  text: string,
  rPr: Element | undefined,
  level: number,
  role: PlaceholderRole | undefined,
  shapeLstStyle: Element | undefined,
  ctx: SlideContext,
): SlideRun {
  // Built-in defaults (SPEC-079 AC-16); on a themed deck the inherited default
  // font is a `minor` reference rather than the baked literal (SPEC-094 AC-16).
  let props: RunProps = {
    fontFamily: ctx.themeCapture ? { theme: "minor" } : ctx.themeMinorLatin,
    sizePt: DEFAULT_FONT_SIZE_PT,
    color: DEFAULT_COLOR,
  };
  // Master placeholder chain: txStyles by role + level (AC-15), then the
  // layout/master placeholder's own lstStyle.
  if (role) {
    const styleLevels = ctx.txStyles.get(role);
    props = mergeProps(props, readRunProps(styleLevels?.[level], ctx.color, ctx.themeCapture));
  }
  // Shape's own lstStyle (AC-14).
  props = mergeProps(props, readRunProps(lstStyleDefRPr(shapeLstStyle, level), ctx.color, ctx.themeCapture));
  // The run's own rPr (AC-13).
  props = mergeProps(props, readRunProps(rPr, ctx.color, ctx.themeCapture));

  const run: SlideRun = { text };
  if (props.bold) run.bold = true;
  if (props.italic) run.italic = true;
  if (props.underline) run.underline = true;
  if (props.fontFamily) run.fontFamily = props.fontFamily;
  if (props.sizePt !== undefined && props.sizePt !== DEFAULT_FONT_SIZE_PT) run.sizePt = props.sizePt;
  if (props.color && props.color !== DEFAULT_COLOR) run.color = props.color;
  return run;
}

/** The bullet a paragraph renders (AC / Reader rules): explicit buChar/
 * buAutoNum/buNone win; else body-role paragraphs default to "disc", title/
 * subtitle to none. */
function bulletOf(pPr: Element | undefined, role: PlaceholderRole | undefined): SlideParagraph["bullet"] {
  if (pPr) {
    if (firstA(pPr, "buNone")) {
      return undefined;
    }
    if (firstA(pPr, "buChar")) {
      return "disc";
    }
    if (firstA(pPr, "buAutoNum")) {
      return "number";
    }
  }
  return role === "body" ? "disc" : undefined;
}

function alignOf(pPr: Element | undefined): SlideParagraph["align"] {
  const algn = pPr?.getAttribute("algn");
  if (algn === "ctr") return "center";
  if (algn === "r") return "right";
  if (algn === "just") return "justify";
  return undefined;
}

/** Parses a `p:txBody` into paragraphs, splitting each `a:p` at explicit
 * `a:br` breaks so every entry is one rendered line. */
function parseTxBody(
  txBody: Element,
  role: PlaceholderRole | undefined,
  ctx: SlideContext,
): SlideParagraph[] {
  const shapeLstStyle = firstA(txBody, "lstStyle");
  const paragraphs: SlideParagraph[] = [];
  for (const p of childrenA(txBody, "p")) {
    const pPr = firstA(p, "pPr");
    const level = Number(pPr?.getAttribute("lvl") ?? "0") || 0;
    const bullet = bulletOf(pPr, role);
    const align = alignOf(pPr);
    // A paragraph splits into visual lines at each a:br.
    let currentRuns: SlideRun[] = [];
    const flush = (): void => {
      const para: SlideParagraph = { runs: currentRuns };
      if (align) para.align = align;
      if (bullet) para.bullet = bullet;
      if (level) para.level = level;
      paragraphs.push(para);
      currentRuns = [];
    };
    for (const child of Array.from(p.children)) {
      if (child.namespaceURI !== A_NAMESPACE) {
        continue;
      }
      if (child.localName === "br") {
        flush();
      } else if (child.localName === "r" || child.localName === "fld") {
        const t = firstA(child, "t");
        const text = t?.textContent ?? "";
        if (text.length === 0) {
          continue;
        }
        currentRuns.push(resolveRun(text, firstA(child, "rPr"), level, role, shapeLstStyle, ctx));
      }
    }
    flush();
  }
  // Drop trailing wholly-empty paragraphs, but keep an all-empty body as [].
  return paragraphs.filter((para) => para.runs.length > 0);
}

/** Resolves a `p:spPr` fill (`a:solidFill`) to a SlideFill, or undefined when
 * the shape carries none. */
function fillOf(spPr: Element | undefined, ctx: SlideContext): SlideFill | undefined {
  if (!spPr) {
    return undefined;
  }
  if (firstA(spPr, "noFill")) {
    return { type: "none" };
  }
  const solidFill = firstA(spPr, "solidFill");
  if (solidFill) {
    const color = readSolidColor(solidFill, ctx.color, ctx.themeCapture);
    if (color) {
      return { type: "solid", color };
    }
  }
  const grad = firstA(spPr, "gradFill");
  if (grad) {
    // Gradient approximation: the first stop's colour (AC-26).
    const gsLst = firstA(grad, "gsLst");
    const firstGs = gsLst ? firstA(gsLst, "gs") : undefined;
    const color = firstGs ? resolveFirstColor(firstGs, ctx.color) : undefined;
    if (color) {
      return { type: "solid", color };
    }
  }
  if (firstA(spPr, "blipFill") || firstA(spPr, "pattFill")) {
    return { type: "solid", color: "#ffffff" };
  }
  return undefined;
}

/** The stroke a `p:spPr/a:ln` describes, or undefined for none. */
function strokeOf(spPr: Element | undefined, ctx: SlideContext): { color: ModelColor; widthPt: number } | undefined {
  if (!spPr) {
    return undefined;
  }
  const ln = firstA(spPr, "ln");
  if (!ln || firstA(ln, "noFill")) {
    return undefined;
  }
  const solidFill = firstA(ln, "solidFill");
  const color = solidFill ? readSolidColor(solidFill, ctx.color, ctx.themeCapture) : undefined;
  if (!color) {
    return undefined;
  }
  const w = ln.getAttribute("w");
  const widthPt = w ? Number(w) / 12700 : 0.75;
  return { color, widthPt };
}

/** The geometry from a `p:spPr`'s own `a:xfrm`, else the inherited
 * placeholder's (layout then master, AC-9/10). */
function geometryFor(
  ownXfrm: DrawingXfrm | undefined,
  phType: string | undefined,
  phIdx: string | undefined,
  role: PlaceholderRole | undefined,
  ctx: SlideContext,
): { xCm: number; yCm: number; widthCm: number; heightCm: number; rotationDeg: number } {
  let xfrm = ownXfrm;
  if (!xfrm || xfrm.xEmu === undefined || xfrm.cxEmu === undefined) {
    const inherited = inheritedPlaceholder(phType, phIdx, role, ctx)?.xfrm;
    if (inherited) {
      xfrm = { ...inherited, ...ownXfrm };
      // A partial own xfrm (rotation only) keeps the inherited offset/extent.
      if (ownXfrm?.xEmu === undefined) {
        xfrm.xEmu = inherited.xEmu;
        xfrm.yEmu = inherited.yEmu;
      }
      if (ownXfrm?.cxEmu === undefined) {
        xfrm.cxEmu = inherited.cxEmu;
        xfrm.cyEmu = inherited.cyEmu;
      }
    }
  }
  return {
    xCm: emuToCm(xfrm?.xEmu ?? 0),
    yCm: emuToCm(xfrm?.yEmu ?? 0),
    widthCm: emuToCm(xfrm?.cxEmu ?? 0),
    heightCm: emuToCm(xfrm?.cyEmu ?? 0),
    rotationDeg: xfrm?.rotationDeg ?? ownXfrm?.rotationDeg ?? 0,
  };
}

function inheritedPlaceholder(
  phType: string | undefined,
  phIdx: string | undefined,
  role: PlaceholderRole | undefined,
  ctx: SlideContext,
): PlaceholderInfo | undefined {
  // Match by (type + idx) — layout first, then master (AC-9/10). A title
  // placeholder also matches ignoring idx (titles rarely carry one).
  void role;
  const key = phKey(phType, phIdx);
  return ctx.layoutPh.get(key) ?? ctx.masterPh.get(key);
}

function frameBase(
  geom: { xCm: number; yCm: number; widthCm: number; heightCm: number; rotationDeg: number },
): { id: string; xCm: number; yCm: number; widthCm: number; heightCm: number; rotationDeg?: number } {
  const base = {
    id: crypto.randomUUID(),
    xCm: geom.xCm,
    yCm: geom.yCm,
    widthCm: geom.widthCm,
    heightCm: geom.heightCm,
  };
  return geom.rotationDeg ? { ...base, rotationDeg: geom.rotationDeg } : base;
}

/** The special field `p:ph@type` a `p:sp` carries (`ftr`/`sldNum`/`dt`), or
 * undefined — SPEC-093 diverts these three out of the frame list. */
function fieldPhType(sp: Element): "ftr" | "sldNum" | "dt" | undefined {
  const nvSpPr = pChild(sp, "nvSpPr");
  const nvPr = nvSpPr && pChild(nvSpPr, "nvPr");
  const ph = nvPr && pChild(nvPr, "ph");
  const type = ph?.getAttribute("type") ?? undefined;
  return type === "ftr" || type === "sldNum" || type === "dt" ? type : undefined;
}

/** Explicit run look (size/colour/font) declared on a field's `a:fld`/`a:r`
 * `a:rPr`, kept only when present so a canonical (look-less) field round-trips
 * to the same absent-look entry. */
function fieldLook(
  rPr: Element | undefined,
  ctx: SlideContext,
): Pick<SlideFieldPlaceholder, "sizePt" | "color" | "fontFamily"> {
  // No capture context: a field's look keeps plain literals (SlideFieldPlaceholder
  // is not widened by SPEC-094).
  const props = readRunProps(rPr, ctx.color);
  const look: Pick<SlideFieldPlaceholder, "sizePt" | "color" | "fontFamily"> = {};
  if (props.sizePt !== undefined) look.sizePt = props.sizePt;
  if (typeof props.color === "string") look.color = props.color;
  if (typeof props.fontFamily === "string") look.fontFamily = props.fontFamily;
  return look;
}

/** Parses a footer/slide-number/date placeholder `p:sp` into a `Slide.fields`
 * entry (SPEC-093 AC-7/13): geometry resolves slide-own `a:xfrm` → the
 * layout/master placeholder map → `DEFAULT_FIELD_BAND`; the cached `a:t` and
 * `a:fld@type` give the value/wire fields. Returns the entry plus any in-subset
 * auto-date format to lift to the deck. */
function parseFieldPlaceholder(
  sp: Element,
  kind: SlideFieldKind,
  ctx: SlideContext,
  deckWidthCm: number,
  deckHeightCm: number,
): { entry: SlideFieldPlaceholder; deckDateFormat?: string } {
  const nvSpPr = pChild(sp, "nvSpPr");
  const nvPr = nvSpPr && pChild(nvSpPr, "nvPr");
  const ph = nvPr && pChild(nvPr, "ph");
  const phType = ph?.getAttribute("type") ?? undefined;
  const phIdx = ph?.getAttribute("idx") ?? undefined;
  const role = roleOf(phType);
  const spPr = pChild(sp, "spPr");
  const ownXfrm = parseXfrm(spPr ? firstA(spPr, "xfrm") : undefined);
  const inherited = inheritedPlaceholder(phType, phIdx, role, ctx);
  let rect: { xCm: number; yCm: number; widthCm: number; heightCm: number };
  if (ownXfrm || inherited) {
    const geom = geometryFor(ownXfrm, phType, phIdx, role, ctx);
    rect = { xCm: geom.xCm, yCm: geom.yCm, widthCm: geom.widthCm, heightCm: geom.heightCm };
  } else {
    const band = defaultFieldRect(kind, deckWidthCm, deckHeightCm);
    rect = { xCm: band.xCm, yCm: band.yCm, widthCm: band.widthCm, heightCm: band.heightCm };
  }
  const entry: SlideFieldPlaceholder = { kind, ...rect };
  const txBody = pChild(sp, "txBody");
  const p = txBody ? firstA(txBody, "p") : undefined;
  const align = alignOf(p ? firstA(p, "pPr") : undefined);
  if (align === "center" || align === "right") {
    entry.align = align;
  }
  const fld = p ? firstA(p, "fld") : undefined;
  const run = p ? firstA(p, "r") : undefined;
  const rPr = fld ? firstA(fld, "rPr") : run ? firstA(run, "rPr") : undefined;
  Object.assign(entry, fieldLook(rPr, ctx));
  const cachedEl = fld ? firstA(fld, "t") : run ? firstA(run, "t") : undefined;
  const cached = cachedEl?.textContent ?? "";
  let deckDateFormat: string | undefined;
  if (kind === "footer") {
    if (cached) {
      entry.text = cached;
    }
  } else if (kind === "date") {
    const fldType = fld?.getAttribute("type") ?? undefined;
    if (fldType && fldType !== "slidenum") {
      const mapped = PPTX_DATETIME_TO_FORMAT[fldType];
      if (mapped) {
        deckDateFormat = mapped; // in-subset auto: lift to deck, recomputed live
      } else {
        entry.wireType = fldType; // out-of-subset: preserve, render from cache
        if (cached) {
          entry.text = cached;
        }
      }
    } else if (cached) {
      // A plain `a:r` date is a fixed literal on this slide.
      entry.text = cached;
    }
  }
  // slideNumber: the value is live from index; nothing cached is kept.
  return { entry, deckDateFormat };
}

/** Parses one `p:sp` (a shape or placeholder) into a frame. */
function parseSp(sp: Element, ctx: SlideContext): SlideFrame {
  const nvSpPr = pChild(sp, "nvSpPr");
  const nvPr = nvSpPr && pChild(nvSpPr, "nvPr");
  const ph = nvPr && pChild(nvPr, "ph");
  const phType = ph?.getAttribute("type") ?? undefined;
  const phIdx = ph?.getAttribute("idx") ?? undefined;
  const isPlaceholder = ph !== undefined && ph !== null;
  const role = isPlaceholder ? roleOf(phType) : undefined;

  const spPr = pChild(sp, "spPr");
  const ownXfrm = parseXfrm(spPr ? firstA(spPr, "xfrm") : undefined);
  const geom = geometryFor(ownXfrm, phType, phIdx, role, ctx);
  const base = frameBase(geom);

  const txBody = pChild(sp, "txBody");
  const prstGeom = spPr ? firstA(spPr, "prstGeom") : undefined;
  const prst = prstGeom?.getAttribute("prst") ?? "";

  if (isPlaceholder) {
    const paragraphs = txBody ? parseTxBody(txBody, role, ctx) : [];
    const frame: SlideFrame = { ...base, kind: "text", role, paragraphs };
    if (phType) frame.phType = phType;
    if (phIdx !== undefined) frame.phIdx = Number(phIdx);
    const fill = fillOf(spPr, ctx);
    if (fill) frame.fill = fill;
    const stroke = strokeOf(spPr, ctx);
    if (stroke) frame.stroke = stroke;
    return frame;
  }

  // Non-placeholder shape: a rect (or no) prstGeom with text is a text box;
  // a known preset is a shape; anything else preserved as a foreign box.
  const geometry = prst ? PRST_TO_GEOMETRY[prst] : undefined;
  if ((prst === "" || prst === "rect") && txBody) {
    const paragraphs = parseTxBody(txBody, undefined, ctx);
    const frame: SlideFrame = { ...base, kind: "text", paragraphs };
    const fill = fillOf(spPr, ctx);
    if (fill) frame.fill = fill;
    const stroke = strokeOf(spPr, ctx);
    if (stroke) frame.stroke = stroke;
    return frame;
  }
  if (geometry) {
    const frame: SlideFrame = { ...base, kind: "shape", geometry };
    const adjustments = prstGeom ? parseAvLstAdjustments(prstGeom) : [];
    if (adjustments.length > 0) frame.adjustments = adjustments;
    const fill = fillOf(spPr, ctx);
    if (fill) frame.fill = fill;
    const stroke = strokeOf(spPr, ctx);
    if (stroke) frame.stroke = stroke;
    if (txBody) {
      const paragraphs = parseTxBody(txBody, undefined, ctx);
      if (paragraphs.length > 0) frame.paragraphs = paragraphs;
    }
    return frame;
  }
  // Unknown/custom geometry → foreign "Object".
  return foreignFrame(sp, base, "Object");
}

function foreignFrame(
  el: Element,
  base: ReturnType<typeof frameBase>,
  label: string,
): SlideFrame {
  const hasBounds = base.widthCm > 0 && base.heightCm > 0;
  const frame: SlideFrame = { ...base, kind: "foreign", xml: serializeElement(el), label };
  if (!hasBounds) {
    frame.noBounds = true;
  }
  return frame;
}

/** Resolves a `p:pic` into a MediaFrame (SPEC-095), else an ImageFrame, else
 * the SPEC-079 "Media" ForeignFrame. */
function parsePic(pic: Element, ctx: SlideContext): SlideFrame {
  const spPr = pChild(pic, "spPr");
  const ownXfrm = parseXfrm(spPr ? firstA(spPr, "xfrm") : undefined);
  const geom = geometryFor(ownXfrm, undefined, undefined, undefined, ctx);
  const base = frameBase(geom);

  const media = parseMediaPic(pic, base, spPr, ctx);
  if (media) {
    return media;
  }

  const blipFill = pChild(pic, "blipFill");
  const blip = blipFill ? firstA(blipFill, "blip") : undefined;
  const embed = blip?.getAttributeNS(R_ATTRIBUTE_NAMESPACE, "embed");
  const rel = embed ? ctx.slideRels.get(embed) : undefined;
  if (rel && rel.type === IMAGE_RELATIONSHIP) {
    const partPath = resolvePartPath(ctx.slidePartPath, rel.target);
    if (ctx.media.has(partPath)) {
      const frame: SlideFrame = { ...base, kind: "image", mediaId: partPath };
      const stroke = strokeOf(spPr, ctx);
      if (stroke) frame.stroke = stroke;
      return frame;
    }
  }
  return foreignFrame(pic, base, "Media");
}

/** The `p14:media` element under a `p:pic`'s `p:nvPr` (its modern embedded
 * pointer), or undefined. */
function firstP14Media(nvPr: Element): Element | undefined {
  return Array.from(nvPr.getElementsByTagNameNS(P14_NAMESPACE, "media"))[0];
}

/** The docu-edit playback-flags recorded in a rebuilt media `p:pic` (SPEC-095),
 * applied to `frame` when present; a no-op for a clip that carries none (a
 * verbatim clip from another app). Mirrors `mediaFlagsExtXml`. */
function applyMediaFlags(frame: MediaFrame, nvPr: Element): void {
  const el = Array.from(nvPr.getElementsByTagNameNS(DE_MEDIA_NAMESPACE, "media"))[0];
  if (!el) {
    return;
  }
  if (el.getAttribute("autoStart") === "1") {
    frame.autoStart = true;
  }
  if (el.getAttribute("loop") === "1") {
    frame.loop = true;
  }
  if (el.getAttribute("mute") === "1") {
    frame.muted = true;
  }
  const vol = el.getAttribute("vol");
  if (vol !== null && vol !== "") {
    const n = Number(vol);
    if (Number.isFinite(n)) {
      frame.volume = n;
    }
  }
  if (el.getAttribute("fullScreen") === "1") {
    frame.fullScreen = true;
  }
}

/** A `p:pic` carrying `a:videoFile`/`a:audioFile` or a `p14:media` extension is
 * an audio/video clip (AC-2): resolves the clip part (or, unresolved, its
 * link/oversize state), the poster blip and the modelled flags, and keeps the
 * verbatim source for byte-exact re-emit. Returns undefined for a `p:pic` that
 * is not media (parsePic then treats it as an image or a "Media" ForeignFrame). */
function parseMediaPic(
  pic: Element,
  base: ReturnType<typeof frameBase>,
  spPr: Element | undefined,
  ctx: SlideContext,
): MediaFrame | undefined {
  const nvPicPr = pChild(pic, "nvPicPr");
  const nvPr = nvPicPr ? pChild(nvPicPr, "nvPr") : undefined;
  if (!nvPr) {
    return undefined;
  }
  const videoFile = firstA(nvPr, "videoFile");
  const audioFile = firstA(nvPr, "audioFile");
  const p14media = firstP14Media(nvPr);
  if (!videoFile && !audioFile && !p14media) {
    return undefined;
  }

  const R = R_ATTRIBUTE_NAMESPACE;
  const fileEl = videoFile ?? audioFile;
  // Resolution order: the modern p14:media embed, then the transitional
  // videoFile/audioFile embed or link.
  const rids = [
    p14media?.getAttributeNS(R, "embed"),
    fileEl?.getAttributeNS(R, "embed"),
    fileEl?.getAttributeNS(R, "link"),
  ];
  let mediaId: string | undefined;
  let linkTarget: string | undefined;
  let oversize = false;
  let contentType: string | undefined;
  for (const rid of rids) {
    if (!rid) {
      continue;
    }
    const rel = ctx.slideRels.get(rid);
    if (!rel) {
      continue;
    }
    const partPath = resolvePartPath(ctx.slidePartPath, rel.target);
    if (ctx.media.has(partPath)) {
      mediaId = partPath;
      contentType = ctx.media.get(partPath)?.contentType;
      break;
    }
    if (ctx.parts.has(partPath)) {
      // In the package but not taken into the model — over the ceiling
      // (collectSlideMedia skipped it): model as oversize, ride out via opaque.
      oversize = true;
      contentType = contentTypeForPartName(partPath);
      break;
    }
    // No such part: a linked (external) clip — a placeholder, never fetched.
    linkTarget = rel.target;
    break;
  }

  const kind: SlideMediaKind = audioFile ? "audio" : videoFile ? "video" : mediaKindForContentType(contentType);
  const frame: MediaFrame = { ...base, kind: "media", media: kind };
  if (mediaId) {
    frame.mediaId = mediaId;
  }
  if (linkTarget) {
    frame.linkTarget = linkTarget;
  }
  if (oversize) {
    frame.oversize = true;
  }
  if (contentType) {
    frame.contentType = contentType;
  }

  const blipFill = pChild(pic, "blipFill");
  const blip = blipFill ? firstA(blipFill, "blip") : undefined;
  const posterEmbed = blip?.getAttributeNS(R, "embed");
  if (posterEmbed) {
    const posterRel = ctx.slideRels.get(posterEmbed);
    if (posterRel && posterRel.type === IMAGE_RELATIONSHIP) {
      const posterPath = resolvePartPath(ctx.slidePartPath, posterRel.target);
      if (ctx.media.has(posterPath)) {
        frame.posterMediaId = posterPath;
      }
    }
  }

  applyMediaFlags(frame, nvPr);
  const stroke = strokeOf(spPr, ctx);
  if (stroke) {
    frame.stroke = stroke;
  }
  // Verbatim source, re-emitted byte-for-byte until a command edits a field.
  frame.sourceXml = serializeElement(pic);
  return frame;
}

/** Parses a `p:graphicFrame`: an in-subset DrawingML table (`a:tbl`) becomes a
 * modelled `TableFrame` (SPEC-091 AC-2a); a chart, an out-of-subset table
 * (AC-2b), or anything else stays a verbatim ForeignFrame — tables never
 * half-parse. */
function parseGraphicFrame(gf: Element, ctx: SlideContext): SlideFrame {
  const xfrm = parseXfrm(pChild(gf, "xfrm"));
  const geom = geometryFor(xfrm, undefined, undefined, undefined, ctx);
  const base = frameBase(geom);
  const graphic = firstA(gf, "graphic");
  const graphicData = graphic ? firstA(graphic, "graphicData") : undefined;
  const uri = graphicData?.getAttribute("uri") ?? "";
  const tbl = graphicData ? firstA(graphicData, "tbl") : undefined;
  if (tbl) {
    const table = parseTbl(tbl, base, gf, ctx);
    if (table) {
      return table;
    }
    return foreignFrame(gf, base, "Table"); // out of subset (AC-2b)
  }
  if (uri === C_NAMESPACE && graphicData) {
    const chart = parseChartGraphicFrame(gf, graphicData, base, ctx);
    if (chart) {
      return chart;
    }
  }
  const label = uri.includes("chart") ? "Chart" : "Object";
  return foreignFrame(gf, base, label);
}

/** Resolves a chart `p:graphicFrame` — `<c:chart r:id>` through the slide's
 * rels to `ppt/charts/chartN.xml` — into a `ChartFrame` (SPEC-092 AC-7). A
 * parseable in-catalogue part becomes a painted chart; a valid `c:chartSpace`
 * outside the catalogue becomes an unsupported placeholder chart carrying the
 * verbatim part text (AC-8); an unresolvable relationship, or a resolved part
 * that is not a `c:chartSpace` at all, returns null so the caller keeps
 * today's "Chart" ForeignFrame (AC-8a). Either way `xml` retains the source
 * `p:graphicFrame` so an untouched chart re-emits byte-identically. */
function parseChartGraphicFrame(
  gf: Element,
  graphicData: Element,
  base: ReturnType<typeof frameBase>,
  ctx: SlideContext,
): SlideFrame | null {
  const chartRef = firstChildElement(graphicData, C_NAMESPACE, "chart");
  const relId = chartRef?.getAttributeNS(R_ATTRIBUTE_NAMESPACE, "id");
  if (!relId) {
    return null;
  }
  const rel = ctx.slideRels.get(relId);
  if (!rel || rel.type !== CHART_RELATIONSHIP) {
    return null;
  }
  const partPath = resolvePartPath(ctx.slidePartPath, rel.target);
  const bytes = ctx.parts.get(partPath);
  if (!bytes) {
    return null;
  }
  const xmlText = new TextDecoder("utf-8").decode(bytes);
  const parsed = parseChartXml(xmlText);
  if (parsed) {
    return { ...base, kind: "chart", ...parsed, sourcePart: partPath, xml: serializeElement(gf) };
  }
  if (!isChartSpace(xmlText)) {
    return null; // not a chart part at all — foreign "Chart" (AC-8a)
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
    xml: serializeElement(gf),
  };
}

/** True when `xmlText` is a well-formed `c:chartSpace` document — the test
 * that separates AC-8 (out-of-catalogue chart) from AC-8a (not a chart). */
function isChartSpace(xmlText: string): boolean {
  let doc: Document;
  try {
    doc = parseXml(xmlText);
  } catch (err) {
    if (err instanceof XmlParseError) {
      return false;
    }
    throw err;
  }
  const root = doc.documentElement;
  return root.namespaceURI === C_NAMESPACE && root.localName === "chartSpace";
}

const TRUE_ATTR = new Set(["1", "true", "on"]);

function boolAttr(el: Element, name: string): boolean {
  const v = el.getAttribute(name);
  return v !== null && TRUE_ATTR.has(v);
}

/** One `a:lnL/R/T/B` edge → a modelled `BorderSide`, or undefined when absent
 * (no explicit edge). A child `a:noFill` is the "none" style. */
function parseBorderSide(ln: Element | undefined, color: ColorContext): BorderSide | undefined {
  if (!ln) {
    return undefined;
  }
  if (firstA(ln, "noFill")) {
    return { style: "none", widthPt: 0, color: "#000000" };
  }
  const solidFill = firstA(ln, "solidFill");
  const resolved = solidFill ? resolveFirstColor(solidFill, color) : undefined;
  const w = ln.getAttribute("w");
  const widthPt = w ? Number(w) / 12700 : 1;
  const cmpd = ln.getAttribute("cmpd");
  const dash = firstA(ln, "prstDash")?.getAttribute("val");
  let style: BorderSide["style"] = "single";
  if (cmpd === "dbl") {
    style = "double";
  } else if (dash === "dash") {
    style = "dashed";
  } else if (dash === "sysDot" || dash === "dot") {
    style = "dotted";
  }
  return { style, widthPt, color: resolved ?? "#000000" };
}

/** Reads a cell's `a:tcPr` into the modelled fill / borders / padding / anchor,
 * dropping the renderer defaults (0.13cm padding, top anchor) so a plain cell
 * stays plain (AC-2a). */
function applyTcPr(cell: SlideTableCell, tcPr: Element, ctx: SlideContext): void {
  const anchor = tcPr.getAttribute("anchor");
  if (anchor === "ctr") {
    cell.verticalAlign = "center";
  } else if (anchor === "b") {
    cell.verticalAlign = "bottom";
  }
  const marL = tcPr.getAttribute("marL");
  if (marL !== null) {
    const cm = Math.round(emuToCm(Number(marL)) * 1000) / 1000;
    if (Math.abs(cm - DEFAULT_TABLE_CELL_PADDING_CM) > 1e-6) {
      cell.paddingCm = cm;
    }
  }
  const borders: CellBorders = {};
  const left = parseBorderSide(firstA(tcPr, "lnL"), ctx.color);
  const right = parseBorderSide(firstA(tcPr, "lnR"), ctx.color);
  const top = parseBorderSide(firstA(tcPr, "lnT"), ctx.color);
  const bottom = parseBorderSide(firstA(tcPr, "lnB"), ctx.color);
  if (left) borders.left = left;
  if (right) borders.right = right;
  if (top) borders.top = top;
  if (bottom) borders.bottom = bottom;
  if (Object.keys(borders).length > 0) {
    cell.borders = borders;
  }
  if (firstA(tcPr, "noFill")) {
    cell.fill = { type: "none" };
  } else {
    const solidFill = firstA(tcPr, "solidFill");
    if (solidFill) {
      const color = readSolidColor(solidFill, ctx.color, ctx.themeCapture);
      if (color) {
        cell.fill = { type: "solid", color };
      }
    }
  }
}

/** Parses an in-subset `a:tbl` into a `TableFrame`, or `undefined` when it is
 * out of subset per the enumerated abort list — a missing/empty grid, no rows,
 * a row whose cells do not total the column count, or an inconsistent merge map
 * (AC-2b). Unmodelled markup is preserved verbatim in pptxTblPr/pptxTcPr/
 * pptxGraphicFramePr (AC-2c). */
function parseTbl(
  tbl: Element,
  base: ReturnType<typeof frameBase>,
  gf: Element,
  ctx: SlideContext,
): TableFrame | undefined {
  const tblGrid = firstA(tbl, "tblGrid");
  const gridCols = tblGrid ? childrenA(tblGrid, "gridCol") : [];
  if (gridCols.length === 0) {
    return undefined; // no columns
  }
  const columnWidthsCm = gridCols.map((col) => Math.round(emuToCm(Number(col.getAttribute("w") ?? "0")) * 1000) / 1000);
  const colCount = columnWidthsCm.length;
  const trs = childrenA(tbl, "tr");
  if (trs.length === 0) {
    return undefined; // no rows
  }
  const rows: SlideTableRow[] = [];
  for (const tr of trs) {
    const tcs = childrenA(tr, "tc");
    if (tcs.length !== colCount) {
      return undefined; // irregular grid (AC-2b)
    }
    const cells: SlideTableCell[] = tcs.map((tc) => {
      if (boolAttr(tc, "hMerge") || boolAttr(tc, "vMerge")) {
        return { paragraphs: [], covered: true };
      }
      const cell: SlideTableCell = { paragraphs: [] };
      const txBody = firstA(tc, "txBody");
      if (txBody) {
        cell.paragraphs = parseTxBody(txBody, undefined, ctx);
      }
      const gridSpan = Number(tc.getAttribute("gridSpan") ?? "1");
      const rowSpan = Number(tc.getAttribute("rowSpan") ?? "1");
      if (gridSpan > 1) cell.colSpan = gridSpan;
      if (rowSpan > 1) cell.rowSpan = rowSpan;
      const tcPr = firstA(tc, "tcPr");
      if (tcPr) {
        applyTcPr(cell, tcPr, ctx);
        if (!tcPrFullyModelled(tcPr)) {
          cell.pptxTcPr = serializeElement(tcPr); // unmodelled markup survives (AC-2c)
        }
      }
      return cell;
    });
    const row: SlideTableRow = { cells };
    const h = tr.getAttribute("h");
    if (h !== null) {
      row.heightCm = Math.round(emuToCm(Number(h)) * 1000) / 1000;
    }
    rows.push(row);
  }
  const r3 = (v: number): number => Math.round(v * 1000) / 1000;
  const frame: TableFrame = {
    ...base,
    xCm: r3(base.xCm),
    yCm: r3(base.yCm),
    widthCm: r3(base.widthCm),
    heightCm: r3(base.heightCm),
    kind: "table",
    rows,
    columnWidthsCm,
  };
  if (!isConsistentMergeGrid(frame)) {
    return undefined; // inconsistent merges (AC-2b)
  }
  const tblPr = firstA(tbl, "tblPr");
  if (tblPr) {
    if (boolAttr(tblPr, "firstRow")) frame.firstRow = true;
    if (boolAttr(tblPr, "bandRow")) frame.bandRow = true;
    const styleId = firstA(tblPr, "tableStyleId")?.textContent ?? undefined;
    if (styleId) frame.styleId = styleId;
    if (!tblPrFullyModelled(tblPr)) {
      frame.pptxTblPr = serializeElement(tblPr); // unmodelled look markup survives (AC-2c)
    }
  }
  const nvGf = pChild(gf, "nvGraphicFramePr");
  if (nvGf && !nvGfFullyModelled(nvGf)) {
    frame.pptxGraphicFramePr = serializeElement(nvGf); // unmodelled non-visual props survive (AC-2c)
  }
  return frame;
}

// --- "fully modelled" predicates (SPEC-091) ---------------------------------
// A verbatim string is captured ONLY when the source carries markup the model
// does not represent — so a table this editor authored (or a fully-modelled
// import) round-trips through the model with no preservation string, and the
// golden reader test decodes to the clean model (AC-16b). The docxTcPr rule.

function attrsWithin(el: Element, allowed: Set<string>): boolean {
  return Array.from(el.attributes).every((a) => allowed.has(a.localName ?? a.name));
}

function aChildrenWithin(el: Element, allowed: Set<string>): boolean {
  return Array.from(el.children).every(
    (c) => c.namespaceURI === A_NAMESPACE && allowed.has(c.localName),
  );
}

const MODELLED_TCPR_ATTRS = new Set(["marL", "marR", "marT", "marB", "anchor"]);
const MODELLED_TCPR_CHILDREN = new Set(["lnL", "lnR", "lnT", "lnB", "solidFill", "noFill"]);
const MODELLED_LN_ATTRS = new Set(["w", "cmpd", "cap", "algn"]);
const MODELLED_LN_CHILDREN = new Set(["solidFill", "noFill", "prstDash"]);

function lnFullyModelled(ln: Element): boolean {
  if (!attrsWithin(ln, MODELLED_LN_ATTRS) || !aChildrenWithin(ln, MODELLED_LN_CHILDREN)) {
    return false;
  }
  const solidFill = firstA(ln, "solidFill");
  return !solidFill || aChildrenWithin(solidFill, new Set(["srgbClr"]));
}

function tcPrFullyModelled(tcPr: Element): boolean {
  if (!attrsWithin(tcPr, MODELLED_TCPR_ATTRS)) {
    return false;
  }
  const mar = ["marL", "marR", "marT", "marB"].map((a) => tcPr.getAttribute(a)).filter((v) => v !== null);
  if (mar.length > 0 && (mar.length !== 4 || new Set(mar).size !== 1)) {
    return false; // asymmetric padding is not representable by the uniform model
  }
  if (!aChildrenWithin(tcPr, MODELLED_TCPR_CHILDREN)) {
    return false;
  }
  for (const child of Array.from(tcPr.children)) {
    if (child.localName.startsWith("ln") && !lnFullyModelled(child)) {
      return false;
    }
    if (child.localName === "solidFill" && !aChildrenWithin(child, new Set(["srgbClr"]))) {
      return false;
    }
  }
  return true;
}

function tblPrFullyModelled(tblPr: Element): boolean {
  return attrsWithin(tblPr, new Set(["firstRow", "bandRow"])) && aChildrenWithin(tblPr, new Set(["tableStyleId"]));
}

function nvGfFullyModelled(nvGf: Element): boolean {
  for (const child of Array.from(nvGf.children)) {
    if (child.namespaceURI !== P_NAMESPACE) {
      return false;
    }
    if (child.localName === "cNvPr") {
      if (!attrsWithin(child, new Set(["id", "name"])) || child.children.length > 0) {
        return false;
      }
    } else if (child.localName === "cNvGraphicFramePr") {
      for (const gc of Array.from(child.children)) {
        if (gc.namespaceURI !== A_NAMESPACE || gc.localName !== "graphicFrameLocks" || !attrsWithin(gc, new Set(["noGrp"]))) {
          return false;
        }
      }
    } else if (child.localName === "nvPr") {
      if (child.children.length > 0) {
        return false;
      }
    } else {
      return false;
    }
  }
  return true;
}

/** Whether every covered slot is under a spanning owner and every owner's span
 * exactly covers its rectangle — the merge map the writer will reproduce. */
function isConsistentMergeGrid(frame: TableFrame): boolean {
  const grid = computeOwnerGrid(frame);
  for (let r = 0; r < frame.rows.length; r++) {
    for (let c = 0; c < frame.columnWidthsCm.length; c++) {
      const owner = grid[r][c];
      const isOwnerSlot = owner.originRow === r && owner.originCol === c;
      const cell = frame.rows[r].cells[c];
      if (isOwnerSlot === Boolean(cell.covered)) {
        return false;
      }
    }
  }
  return true;
}

/** Test seam (SPEC-091 golden): parse a standalone `p:graphicFrame` table XML
 * string into a `TableFrame` (or undefined when out of subset), through a
 * minimal context — table cells carry explicit run properties, so no theme is
 * needed. */
export function parseTableGraphicFrameForTest(graphicFrameXml: string): TableFrame | undefined {
  // A bare graphicFrame carries no xmlns (its slide root normally declares p/a);
  // inject them when absent so the string parses standalone.
  const xml = graphicFrameXml.includes("xmlns:p")
    ? graphicFrameXml
    : graphicFrameXml.replace(
        "<p:graphicFrame>",
        `<p:graphicFrame xmlns:p="${P_NAMESPACE}" xmlns:a="${A_NAMESPACE}" xmlns:r="${R_ATTRIBUTE_NAMESPACE}">`,
      );
  const doc = parseXml(xml);
  const gf = doc.documentElement;
  const ctx: SlideContext = {
    color: { scheme: {}, map: {} },
    layoutPh: new Map(),
    masterPh: new Map(),
    txStyles: new Map(),
    slideRels: new Map(),
    slidePartPath: "",
    media: new Map(),
    parts: new Map(),
  };
  const frame = parseGraphicFrame(gf, ctx);
  return frame.kind === "table" ? frame : undefined;
}

/** Test seam (SPEC-096 golden): parses a standalone `p:sp`/`p:grpSp` string
 * into a SlideFrame through a minimal context, so the group/adjustment
 * goldens can assert the reader accepts the canonical bytes plus
 * attribute-reordered variants. */
export function parseShapeTreeChildForTest(elementXml: string): SlideFrame | null {
  const xml = elementXml.includes("xmlns:p")
    ? elementXml
    : elementXml.replace(/^<(p:sp|p:grpSp)>/, `<$1 xmlns:p="${P_NAMESPACE}" xmlns:a="${A_NAMESPACE}">`);
  const doc = parseXml(xml);
  const ctx: SlideContext = {
    color: { scheme: {}, map: {} },
    layoutPh: new Map(),
    masterPh: new Map(),
    txStyles: new Map(),
    slideRels: new Map(),
    slidePartPath: "",
    media: new Map(),
    parts: new Map(),
  };
  return parseShapeTreeChild(doc.documentElement, ctx);
}

/** Test seam (SPEC-093): parses one field placeholder `p:sp` into a
 * `SlideFieldPlaceholder`, so the golden check can assert the reader accepts the
 * canonical bytes plus attribute-reordered variants. */
export function parseFieldPlaceholderForTest(
  spXml: string,
  kind: SlideFieldKind,
  deckWidthCm = DEFAULT_SLIDE_WIDTH_CM,
  deckHeightCm = DEFAULT_SLIDE_HEIGHT_CM,
): SlideFieldPlaceholder {
  const xml = spXml.includes("xmlns:p")
    ? spXml
    : spXml.replace("<p:sp>", `<p:sp xmlns:p="${P_NAMESPACE}" xmlns:a="${A_NAMESPACE}">`);
  const doc = parseXml(xml);
  const ctx: SlideContext = {
    color: { scheme: {}, map: {} },
    layoutPh: new Map(),
    masterPh: new Map(),
    txStyles: new Map(),
    slideRels: new Map(),
    slidePartPath: "",
    media: new Map(),
    parts: new Map(),
  };
  return parseFieldPlaceholder(doc.documentElement, kind, ctx, deckWidthCm, deckHeightCm).entry;
}

/** Test seam (SPEC-092): parses a chart `p:graphicFrame` against a supplied
 * `{ relId → partPath }` rel map and `{ partPath → bytes }` parts map. */
export function parseChartGraphicFrameForTest(
  graphicFrameXml: string,
  rels: Record<string, string>,
  parts: Record<string, Uint8Array>,
): SlideFrame {
  const xml = graphicFrameXml.includes("xmlns:p")
    ? graphicFrameXml
    : graphicFrameXml.replace(
        "<p:graphicFrame>",
        `<p:graphicFrame xmlns:p="${P_NAMESPACE}" xmlns:a="${A_NAMESPACE}" xmlns:r="${R_ATTRIBUTE_NAMESPACE}" xmlns:c="${C_NAMESPACE}">`,
      );
  const gf = parseXml(xml).documentElement;
  const slideRels = new Map<string, Relationship>(
    Object.entries(rels).map(([id, target]) => [id, { id, type: CHART_RELATIONSHIP, target }]),
  );
  const ctx: SlideContext = {
    color: { scheme: {}, map: {} },
    layoutPh: new Map(),
    masterPh: new Map(),
    txStyles: new Map(),
    slideRels,
    slidePartPath: "ppt/slides/slide1.xml",
    media: new Map(),
    parts: new Map(Object.entries(parts)),
  };
  return parseGraphicFrame(gf, ctx);
}

/** Test seam (SPEC-095 golden): parses a standalone media `p:pic` against
 * supplied slide rels, collected media parts and raw package parts, so the
 * media golden can assert the reader accepts the canonical bytes plus a
 * PowerPoint-flavoured variant, and degrades an unrecognised `p:pic` to the
 * "Media" ForeignFrame. */
export function parseMediaPicForTest(
  picXml: string,
  rels: Relationship[],
  media: MediaPart[],
  parts: Record<string, Uint8Array> = {},
): SlideFrame {
  // `xmlns:p="` specifically (not `xmlns:p14="`, which a media p:pic carries).
  const xml = picXml.includes(`xmlns:p="`)
    ? picXml
    : picXml.replace(
        "<p:pic>",
        `<p:pic xmlns:p="${P_NAMESPACE}" xmlns:a="${A_NAMESPACE}" xmlns:r="${R_ATTRIBUTE_NAMESPACE}" xmlns:p14="${P14_NAMESPACE}">`,
      );
  const pic = parseXml(xml).documentElement;
  const ctx: SlideContext = {
    color: { scheme: {}, map: {} },
    layoutPh: new Map(),
    masterPh: new Map(),
    txStyles: new Map(),
    slideRels: new Map(rels.map((rel) => [rel.id, rel])),
    slidePartPath: "ppt/slides/slide1.xml",
    media: new Map(media.map((part) => [part.mediaId, part])),
    parts: new Map(Object.entries(parts)),
  };
  return parsePic(pic, ctx);
}

/** Parses a `p:cxnSp` connector: a line/straight connector is a line shape,
 * anything else a foreign box. */
function parseCxnSp(cxn: Element, ctx: SlideContext): SlideFrame {
  const spPr = pChild(cxn, "spPr");
  const ownXfrm = parseXfrm(spPr ? firstA(spPr, "xfrm") : undefined);
  const geom = geometryFor(ownXfrm, undefined, undefined, undefined, ctx);
  const base = frameBase(geom);
  const prstGeom = spPr ? firstA(spPr, "prstGeom") : undefined;
  const prst = prstGeom?.getAttribute("prst") ?? "";
  if (prst === "line" || prst === "straightConnector1") {
    const frame: SlideFrame = { ...base, kind: "shape", geometry: "line" };
    const stroke = strokeOf(spPr, ctx);
    if (stroke) frame.stroke = stroke;
    return frame;
  }
  return foreignFrame(cxn, base, "Object");
}

/** Parses a `p:grpSp` group into a GroupFrame whose children are flattened
 * into slide space via the group's chOff/chExt mapping (AC-37). */
function parseGrpSp(grp: Element, ctx: SlideContext): SlideFrame {
  const grpSpPr = pChild(grp, "grpSpPr");
  const mapping = grpSpPr ? parseGroupChildMapping(grpSpPr) : undefined;
  const groupXfrm = grpSpPr ? parseXfrm(firstA(grpSpPr, "xfrm")) : undefined;
  const groupGeom = geometryFor(groupXfrm, undefined, undefined, undefined, ctx);
  const children: SlideFrame[] = [];
  for (const child of Array.from(grp.children)) {
    if (child.namespaceURI !== P_NAMESPACE) {
      continue;
    }
    const frame = parseShapeTreeChild(child, ctx);
    if (frame) {
      children.push(mapping ? mapChildFrame(frame, mapping) : frame);
    }
  }
  // A group missing its own xfrm (e.g. an empty `<p:grpSpPr/>`) yields 0×0
  // geometry and an unclickable hit box; fall back to the children's union in
  // slide space, matching the ODP reader (AC-31).
  const geom =
    groupXfrm?.xEmu === undefined || groupXfrm?.cxEmu === undefined
      ? { ...unionBounds(children), rotationDeg: groupGeom.rotationDeg }
      : groupGeom;
  const base = frameBase(geom);
  return { ...base, kind: "group", children, xml: serializeElement(grp) };
}

/** Union of child frames' axis-aligned bounds in slide space, in cm. Used as a
 * group's own box when its `grpSpPr` carries no usable `xfrm` (AC-31). */
function unionBounds(
  frames: SlideFrame[],
): { xCm: number; yCm: number; widthCm: number; heightCm: number } {
  if (frames.length === 0) {
    return { xCm: 0, yCm: 0, widthCm: 0, heightCm: 0 };
  }
  const minX = Math.min(...frames.map((f) => f.xCm));
  const minY = Math.min(...frames.map((f) => f.yCm));
  const maxX = Math.max(...frames.map((f) => f.xCm + f.widthCm));
  const maxY = Math.max(...frames.map((f) => f.yCm + f.heightCm));
  return { xCm: minX, yCm: minY, widthCm: maxX - minX, heightCm: maxY - minY };
}

/** Applies a group's child mapping to a fully-parsed child frame's own
 * geometry (already in the group's child coordinate space, AC-37). Nested
 * group children are remapped by the same mapping recursively so a one-level
 * group lands its children exactly in slide space. */
function mapChildFrame(frame: SlideFrame, mapping: GroupChildMapping): SlideFrame {
  const mapped = mapGroupChild(
    mapping,
    cmToEmu(frame.xCm),
    cmToEmu(frame.yCm),
    cmToEmu(frame.widthCm),
    cmToEmu(frame.heightCm),
  );
  const next = {
    ...frame,
    xCm: emuToCm(mapped.xEmu),
    yCm: emuToCm(mapped.yEmu),
    widthCm: emuToCm(mapped.cxEmu),
    heightCm: emuToCm(mapped.cyEmu),
  } as SlideFrame;
  if (next.kind === "group") {
    next.children = next.children.map((child) => mapChildFrame(child, mapping));
  }
  return next;
}

/** Parses one child of a `p:spTree`/`p:grpSp` into a frame, or null for a
 * container element (nvGrpSpPr/grpSpPr) that is not itself a frame. */
function parseShapeTreeChild(child: Element, ctx: SlideContext): SlideFrame | null {
  switch (child.localName) {
    case "sp":
      return parseSp(child, ctx);
    case "pic":
      return parsePic(child, ctx);
    case "graphicFrame":
      return parseGraphicFrame(child, ctx);
    case "cxnSp":
      return parseCxnSp(child, ctx);
    case "grpSp":
      return parseGrpSp(child, ctx);
    default:
      return null;
  }
}

/** Builds the (type+idx)→placeholder map from a layout/master `p:cSld/p:spTree`. */
function buildPlaceholderMap(spTree: Element | undefined): Map<string, PlaceholderInfo> {
  const map = new Map<string, PlaceholderInfo>();
  if (!spTree) {
    return map;
  }
  for (const sp of childElements(spTree, P_NAMESPACE, "sp")) {
    const nvSpPr = pChild(sp, "nvSpPr");
    const nvPr = nvSpPr && pChild(nvSpPr, "nvPr");
    const ph = nvPr && pChild(nvPr, "ph");
    if (!ph) {
      continue;
    }
    const phType = ph.getAttribute("type") ?? undefined;
    const phIdx = ph.getAttribute("idx") ?? undefined;
    const spPr = pChild(sp, "spPr");
    const txBody = pChild(sp, "txBody");
    map.set(phKey(phType, phIdx), {
      xfrm: parseXfrm(spPr ? firstA(spPr, "xfrm") : undefined),
      lstStyle: txBody ? firstA(txBody, "lstStyle") : undefined,
      role: roleOf(phType),
    });
  }
  return map;
}

/** Reads a master `p:txStyles` into role → level → defRPr. */
function buildTxStyles(masterRoot: Element | undefined): Map<PlaceholderRole, Element[]> {
  const styles = new Map<PlaceholderRole, Element[]>();
  if (!masterRoot) {
    return styles;
  }
  const txStyles = pChild(masterRoot, "txStyles");
  if (!txStyles) {
    return styles;
  }
  const read = (styleEl: Element | undefined): Element[] => {
    const levels: Element[] = [];
    if (!styleEl) {
      return levels;
    }
    for (let i = 0; i < 9; i++) {
      const lvl = firstA(styleEl, `lvl${i + 1}pPr`);
      const defRPr = lvl ? firstA(lvl, "defRPr") : undefined;
      if (defRPr) {
        levels[i] = defRPr;
      }
    }
    return levels;
  };
  styles.set("title", read(pChild(txStyles, "titleStyle")));
  const body = read(pChild(txStyles, "bodyStyle"));
  styles.set("body", body);
  styles.set("subtitle", body);
  styles.set("other", read(pChild(txStyles, "otherStyle")));
  return styles;
}

/** The theme minor-latin typeface, the built-in font default (AC-16). */
function themeMinorLatin(themeRoot: Element | undefined): string | undefined {
  if (!themeRoot) {
    return undefined;
  }
  const elements = firstA(themeRoot, "themeElements");
  const fontScheme = elements && firstA(elements, "fontScheme");
  const minor = fontScheme && firstA(fontScheme, "minorFont");
  const latin = minor && firstA(minor, "latin");
  return latin?.getAttribute("typeface") || undefined;
}

/** Resolves the slide's effective background (AC-22..28): explicit slide
 * `p:bg`, else the layout's, else the master's — solid exactly, gradient as
 * first stop, image/pattern as white; never throws. */
function backgroundOf(
  slideRoot: Element,
  layoutRoot: Element | undefined,
  masterRoot: Element | undefined,
  ctx: SlideContext,
): SlideFill | undefined {
  for (const root of [slideRoot, layoutRoot, masterRoot]) {
    if (!root) {
      continue;
    }
    const cSld = pChild(root, "cSld");
    const bg = cSld && pChild(cSld, "bg");
    if (!bg) {
      continue;
    }
    const bgPr = pChild(bg, "bgPr");
    if (bgPr) {
      const fill = fillOf(bgPr, ctx);
      if (fill) {
        return fill;
      }
    }
    // p:bgRef → a scheme colour reference; approximate as its resolved colour.
    const bgRef = pChild(bg, "bgRef");
    if (bgRef) {
      const color = resolveFirstColor(bgRef, ctx.color);
      if (color) {
        return { type: "solid", color };
      }
    }
  }
  return undefined;
}

/** Strips trailing wholly-empty paragraphs from a notes list, returning
 * `undefined` when nothing text-bearing remains (`[]` is never produced). */
function trimNotes(paragraphs: string[]): string[] | undefined {
  const trimmed = paragraphs.slice();
  while (trimmed.length > 0 && trimmed[trimmed.length - 1].length === 0) {
    trimmed.pop();
  }
  return trimmed.length > 0 ? trimmed : undefined;
}

/** The plain-text speaker notes of a slide (SPEC-083, AC-2/4): follow the
 * slide's `notesSlide` relationship, read the `p:sp` whose `p:ph type="body"`
 * — each `a:p` maps to one entry (runs joined, formatting flattened) — and
 * drop trailing empty paragraphs. Undefined when there is no notes part or its
 * body carries no text; other shapes (slide-image placeholder, headers) are
 * ignored. */
function readNotes(
  slideRels: Map<string, Relationship>,
  slidePath: string,
  parseDoc: (path: string) => Document | undefined,
): string[] | undefined {
  const rel = Array.from(slideRels.values()).find((r) => r.type === NOTES_SLIDE_RELATIONSHIP);
  if (!rel) {
    return undefined;
  }
  const notesRoot = parseDoc(resolvePartPath(slidePath, rel.target))?.documentElement;
  if (!notesRoot) {
    return undefined;
  }
  const cSld = pChild(notesRoot, "cSld");
  const spTree = cSld && pChild(cSld, "spTree");
  if (!spTree) {
    return undefined;
  }
  for (const sp of childElements(spTree, P_NAMESPACE, "sp")) {
    const nvSpPr = pChild(sp, "nvSpPr");
    const nvPr = nvSpPr && pChild(nvSpPr, "nvPr");
    const ph = nvPr && pChild(nvPr, "ph");
    if (!ph || ph.getAttribute("type") !== "body") {
      continue;
    }
    const txBody = pChild(sp, "txBody");
    if (!txBody) {
      return undefined;
    }
    return trimNotes(childrenA(txBody, "p").map((p) => p.textContent ?? ""));
  }
  return undefined;
}

interface ChainDocs {
  layoutRoot?: Element;
  masterRoot?: Element;
  themeRoot?: Element;
  /** The resolved theme part path, when the chain reached one (SPEC-094). */
  themePath?: string;
}

/** Parses a `.pptx` file's bytes into the slide model. Throws
 * `PptxOpenError` for anything that isn't a valid, openable `.pptx`. */
export async function parsePptx(
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
        throw new PptxOpenError(INVALID_PPTX_MESSAGE);
      }
      throw error;
    }
  }
  const info = detectPackage(parts);
  if (info.container !== "opc" || info.mainContentType !== PPTX_MAIN_CONTENT_TYPE) {
    throw new PptxOpenError(INVALID_PPTX_MESSAGE);
  }

  const readPart = (path: string): string | undefined => {
    const raw = parts.get(path);
    return raw ? decodeUtf8(raw) : undefined;
  };
  const parseDoc = (path: string): Document | undefined => {
    const text = readPart(path);
    if (!text) {
      return undefined;
    }
    try {
      return parseXml(text);
    } catch {
      return undefined;
    }
  };
  const relsFor = (path: string): Map<string, Relationship> => {
    const text = readPart(relsPathFor(path));
    return text ? parseRelationships(text) : new Map();
  };

  const mainText = readPart(info.mainPartPath);
  if (!mainText) {
    throw new PptxOpenError(INVALID_PPTX_MESSAGE);
  }
  let mainDoc: Document;
  try {
    mainDoc = parseXml(mainText);
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw new PptxOpenError(INVALID_PPTX_MESSAGE);
    }
    throw error;
  }
  if (mainDoc.documentElement.namespaceURI !== P_NAMESPACE || mainDoc.documentElement.localName !== "presentation") {
    throw new PptxOpenError(INVALID_PPTX_MESSAGE);
  }

  // Slide size (AC-2).
  const sldSz = pChild(mainDoc.documentElement, "sldSz");
  const slideWidthCm = sldSz ? emuToCm(Number(sldSz.getAttribute("cx"))) : DEFAULT_SLIDE_WIDTH_CM;
  const slideHeightCm = sldSz ? emuToCm(Number(sldSz.getAttribute("cy"))) : DEFAULT_SLIDE_HEIGHT_CM;

  // Slide-number origin (SPEC-090): `firstSlideNum` on the p:presentation start
  // tag; absent or 1 means the deck counts from 1.
  const firstSlideNumRaw = mainDoc.documentElement.getAttribute("firstSlideNum");
  const firstSlideNumParsed = firstSlideNumRaw !== null && firstSlideNumRaw !== "" ? Number(firstSlideNumRaw) : NaN;
  const firstSlideNumber = Number.isInteger(firstSlideNumParsed) ? firstSlideNumParsed : undefined;

  // Title-slide field policy (SPEC-093): `showSpecialPlsOnTitleSld="0"` hides
  // every field on the title slide; absent or "1" = show.
  const showSpecialPlsRaw = mainDoc.documentElement.getAttribute("showSpecialPlsOnTitleSld");

  const mainRels = relsFor(info.mainPartPath);

  // Collect every media part up-front, keyed by its own path (AC-31/32).
  const media = new Map<string, MediaPart>();

  const deps: BuildDeps = {
    readPart,
    themeHolder: {},
    parseDoc,
    relsFor,
    media,
    parts,
    deckWidthCm: slideWidthCm,
    deckHeightCm: slideHeightCm,
  };

  const slides: Slide[] = [];
  let deckDateFormat: string | undefined;
  const sldIdLst = pChild(mainDoc.documentElement, "sldIdLst");
  for (const sldId of sldIdLst ? childElements(sldIdLst, P_NAMESPACE, "sldId") : []) {
    const relId = sldId.getAttributeNS(R_ATTRIBUTE_NAMESPACE, "id");
    const rel = relId ? mainRels.get(relId) : undefined;
    if (!rel || rel.type !== SLIDE_RELATIONSHIP) {
      continue;
    }
    const slidePath = resolvePartPath(info.mainPartPath, rel.target);
    const slideText = readPart(slidePath);
    if (slideText === undefined) {
      continue;
    }
    let slideDoc: Document;
    try {
      slideDoc = parseXml(slideText);
    } catch {
      // Unparseable slide still counts (AC-61): an empty slide.
      slides.push({ id: crypto.randomUUID(), frames: [], partName: slidePath });
      continue;
    }
    const built = buildSlide(slideDoc, slidePath, deps);
    slides.push(built.slide);
    if (built.deckDateFormat && !deckDateFormat) {
      deckDateFormat = built.deckDateFormat;
    }
  }

  const headerFooter = assembleHeaderFooter(slides, deckDateFormat, showSpecialPlsRaw);

  // The read-derived layout catalogue (SPEC-089 AC-2).
  const layouts = buildLayoutCatalogue(mainDoc, mainRels, info.mainPartPath, slides, deps);

  // Document properties (SPEC-041), resolved through the root relationships.
  let properties: DocumentProperties = {};
  const rootRelsText = readPart("_rels/.rels");
  if (rootRelsText) {
    const rootRels = parseRelationships(rootRelsText);
    const coreRel = Array.from(rootRels.values()).find((r) => r.type === CORE_PROPERTIES_RELATIONSHIP);
    const coreText = coreRel ? readPart(resolvePartPath("", coreRel.target)) : undefined;
    if (coreText) {
      properties = parseCoreProperties(coreText);
    }
    const customRel = Array.from(rootRels.values()).find((r) => r.type === CUSTOM_PROPERTIES_RELATIONSHIP);
    const customText = customRel ? readPart(resolvePartPath("", customRel.target)) : undefined;
    if (customText) {
      const { customProperties, unsupportedCustomProperties } = parseCustomProperties(customText);
      if (customProperties.length > 0) {
        properties.customProperties = customProperties;
      }
      if (unsupportedCustomProperties.length > 0) {
        properties.unsupportedCustomProperties = unsupportedCustomProperties;
      }
    }
  }

  // Retain the whole package, in package order (SPEC-008 wholesale).
  const opaque: OpaquePart[] = [];
  for (const [name, partBytes] of parts) {
    opaque.push({ name, data: await bytesToBase64Async(partBytes) });
  }

  const parsed: ParsedPresentation = {
    slides,
    layouts,
    slideWidthCm,
    slideHeightCm,
    firstSlideNumber,
    headerFooter,
    media: [...media.values()],
    opaque,
    properties,
  };
  if (deps.themeHolder.theme) {
    parsed.theme = deps.themeHolder.theme;
  }
  return parsed;
}

/** Assembles the deck-level `headerFooter` from the read slides (SPEC-093):
 * lifts the common footer text to the deck (divergent slides keep a per-slide
 * override), takes the auto-date format from the first in-subset date field,
 * and reads the title-slide policy from `showSpecialPlsOnTitleSld`. */
function assembleHeaderFooter(
  slides: Slide[],
  deckDateFormat: string | undefined,
  showSpecialPlsRaw: string | null,
): DeckHeaderFooter | undefined {
  let deckFooterText: string | undefined;
  for (const slide of slides) {
    const footer = slide.fields?.find((field) => field.kind === "footer" && field.text !== undefined && field.text !== "");
    if (footer) {
      deckFooterText = footer.text;
      break;
    }
  }
  if (deckFooterText !== undefined) {
    for (const slide of slides) {
      for (const field of slide.fields ?? []) {
        if (field.kind === "footer" && field.text === deckFooterText) {
          delete field.text;
        }
      }
    }
  }
  const hf: DeckHeaderFooter = {};
  if (deckFooterText !== undefined) {
    hf.footerText = deckFooterText;
  }
  if (deckDateFormat !== undefined && deckDateFormat !== DEFAULT_DATE_FORMAT) {
    hf.dateFormat = deckDateFormat;
  }
  if (showSpecialPlsRaw === "0") {
    hf.showOnTitleSlide = false;
  }
  return Object.keys(hf).length > 0 ? hf : undefined;
}

interface BuildDeps {
  readPart: (path: string) => string | undefined;
  /** Mutable deck-theme holder (SPEC-094): the FIRST master's theme is the
   * deck's editable one; later chains reaching a different part read literals. */
  themeHolder: { theme?: DeckTheme };
  parseDoc: (path: string) => Document | undefined;
  relsFor: (path: string) => Map<string, Relationship>;
  media: Map<string, MediaPart>;
  parts: Map<string, Uint8Array>;
  /** Deck slide size, so a field placeholder with no resolvable geometry falls
   *  back to `DEFAULT_FIELD_BAND` (SPEC-093 AC-13). */
  deckWidthCm: number;
  deckHeightCm: number;
}

/** Resolves a slide's layout/master/theme chain and parses its shape tree.
 * Returns the slide plus any in-subset auto-date format found on it, which
 * `parsePptx` lifts to the deck-level `headerFooter` (SPEC-093). */
function buildSlide(
  slideDoc: Document,
  slidePath: string,
  deps: BuildDeps,
): { slide: Slide; deckDateFormat?: string } {
  const slideRoot = slideDoc.documentElement;
  const slideRels = deps.relsFor(slidePath);

  // slide → layout → master → theme, each through its own .rels (AC: never
  // assume part paths).
  const chain = resolveChain(slidePath, slideRels, deps);

  const color: ColorContext = {
    scheme: chain.themeRoot ? parseClrScheme(chain.themeRoot) : {},
    map: chain.masterRoot ? parseClrMap(chain.masterRoot) : {},
  };

  // SPEC-094: the first master's theme becomes the deck theme; a slide whose
  // chain ends in THAT part reads its scheme colours/fonts as references. A
  // slide hanging off a different theme part (a second master) keeps literals,
  // so its rendered colours cannot drift when the deck theme is edited.
  if (!deps.themeHolder.theme && chain.themeRoot && chain.themePath) {
    deps.themeHolder.theme = readDeckTheme(chain.themeRoot, chain.masterRoot, chain.themePath, DEFAULT_THEME);
  }
  const deckTheme = deps.themeHolder.theme;
  const themeCapture =
    deckTheme && chain.themePath === deckTheme.partName ? { map: deckTheme.colorMap ?? {} } : undefined;

  const ctx: SlideContext = {
    color,
    themeCapture,
    layoutPh: buildPlaceholderMap(chain.layoutRoot ? layoutSpTree(chain.layoutRoot) : undefined),
    masterPh: buildPlaceholderMap(chain.masterRoot ? layoutSpTree(chain.masterRoot) : undefined),
    txStyles: buildTxStyles(chain.masterRoot),
    themeMinorLatin: themeMinorLatin(chain.themeRoot),
    slideRels,
    slidePartPath: slidePath,
    media: deps.media,
    parts: deps.parts,
  };

  // Populate media referenced by this slide's image rels.
  collectSlideMedia(slideRels, slidePath, deps);

  const cSld = pChild(slideRoot, "cSld");
  const spTree = cSld && pChild(cSld, "spTree");
  const frames: SlideFrame[] = [];
  // `p:cNvPr id` → minted frame id, so a `p:timing`'s `p:spTgt spid` resolves
  // back to the frame it targets (SPEC-086).
  const frameForSpid = new Map<string, string>();
  // SPEC-093: footer/slide-number/date placeholders are diverted out of the
  // frame list into `fields`, their source `p:sp` kept under `p:ph:{type}`.
  const fields: SlideFieldPlaceholder[] = [];
  const fieldPreserved: Record<string, string> = {};
  let deckDateFormat: string | undefined;
  if (spTree) {
    for (const child of Array.from(spTree.children)) {
      if (child.namespaceURI !== P_NAMESPACE) {
        continue;
      }
      if (child.localName === "sp") {
        const phType = fieldPhType(child);
        if (phType) {
          const parsed = parseFieldPlaceholder(child, KIND_OF_PH_TYPE[phType], ctx, deps.deckWidthCm, deps.deckHeightCm);
          fields.push(parsed.entry);
          fieldPreserved[`p:ph:${phType}`] = serializeElement(child);
          if (parsed.deckDateFormat && !deckDateFormat) {
            deckDateFormat = parsed.deckDateFormat;
          }
          continue;
        }
      }
      const frame = parseShapeTreeChild(child, ctx);
      if (frame) {
        frames.push(frame);
        const spid = firstCNvPrId(child);
        if (spid) {
          frameForSpid.set(spid, frame.id);
        }
      }
    }
  }

  const slide: Slide = { id: crypto.randomUUID(), frames, partName: slidePath };
  // The slide's layout key is its layout part's absolute path (SPEC-089 AC-2),
  // which the catalogue also keys by, so `layoutFor` resolves after re-parse.
  const layoutRel = Array.from(slideRels.values()).find((r) => r.type === SLIDE_LAYOUT_RELATIONSHIP);
  if (layoutRel) {
    slide.layoutKey = resolvePartPath(slidePath, layoutRel.target);
  }
  const background = backgroundOf(slideRoot, chain.layoutRoot, chain.masterRoot, ctx);
  if (background) {
    slide.background = background;
  }
  const preserved = collectPreserved(slideRoot);
  // A subset transition becomes the modelled field; its source XML is then
  // dropped from `preserved` so it is never emitted twice. An out-of-subset
  // transition stays in `preserved` untouched (SPEC-085 AC-2/AC-3).
  if (preserved["p:transition"]) {
    const transitionEl = pChild(slideRoot, "transition");
    const transition = transitionEl ? readPptxTransition(transitionEl) : undefined;
    if (transition) {
      slide.transition = transition;
      delete preserved["p:transition"];
    }
  }
  // A canonical (or PowerPoint-authored equivalent) `p:timing` whose targets all
  // resolve becomes the modelled `animations`; its source XML is then dropped
  // from `preserved`. Anything out of subset stays preserved verbatim (AC-4).
  if (preserved["p:timing"]) {
    const timingEl = pChild(slideRoot, "timing");
    const animations = timingEl ? readPptxTiming(timingEl, (spid) => frameForSpid.get(spid)) : undefined;
    if (animations) {
      slide.animations = animations;
      delete preserved["p:timing"];
    }
  }
  Object.assign(preserved, fieldPreserved);
  if (Object.keys(preserved).length > 0) {
    slide.preserved = preserved;
  }
  if (fields.length > 0) {
    slide.fields = fields;
  }
  const notes = readNotes(slideRels, slidePath, deps.parseDoc);
  if (notes) {
    slide.notes = notes;
  }
  return { slide, deckDateFormat };
}

function layoutSpTree(root: Element): Element | undefined {
  const cSld = pChild(root, "cSld");
  return cSld ? pChild(cSld, "spTree") : undefined;
}

/** The `LayoutSlot[]` a `p:sldLayout` root declares: one per placeholder `p:sp`,
 * with role from `p:ph@type`, verbatim `phType`/`phIdx`, and cm geometry from
 * its own `a:xfrm` (SPEC-089 AC-2). Placeholders without an `a:xfrm` (geometry
 * inherited from the master) contribute a zero-geometry slot. */
function layoutSlotsFrom(layoutRoot: Element): LayoutSlot[] {
  const spTree = layoutSpTree(layoutRoot);
  if (!spTree) {
    return [];
  }
  const slots: LayoutSlot[] = [];
  for (const sp of childElements(spTree, P_NAMESPACE, "sp")) {
    const nvSpPr = pChild(sp, "nvSpPr");
    const nvPr = nvSpPr && pChild(nvSpPr, "nvPr");
    const ph = nvPr && pChild(nvPr, "ph");
    if (!ph) {
      continue;
    }
    const phType = ph.getAttribute("type") ?? undefined;
    const phIdx = ph.getAttribute("idx") ?? undefined;
    const spPr = pChild(sp, "spPr");
    const xfrm = parseXfrm(spPr ? firstA(spPr, "xfrm") : undefined);
    const slot: LayoutSlot = {
      role: roleOf(phType),
      xCm: emuToCm(xfrm?.xEmu ?? 0),
      yCm: emuToCm(xfrm?.yEmu ?? 0),
      widthCm: emuToCm(xfrm?.cxEmu ?? 0),
      heightCm: emuToCm(xfrm?.cyEmu ?? 0),
    };
    if (phType) slot.phType = phType;
    if (phIdx !== undefined) slot.phIdx = Number(phIdx);
    slots.push(slot);
  }
  return slots;
}

/** Parses one layout part into a `SlideLayout`, keyed by its absolute part
 * path. Returns undefined when the part is missing/unparseable. */
function layoutEntryFor(layoutPath: string, masterName: string | undefined, deps: BuildDeps): SlideLayout | undefined {
  const root = deps.parseDoc(layoutPath)?.documentElement;
  if (!root) {
    return undefined;
  }
  const cSld = pChild(root, "cSld");
  const name = cSld?.getAttribute("name") ?? layoutPath.split("/").pop() ?? layoutPath;
  const type = root.getAttribute("type") ?? undefined;
  const slots = layoutSlotsFrom(root);
  const layout: SlideLayout = { key: layoutPath, kind: classifyLayout(type, slots), name, slots };
  if (masterName) {
    layout.masterName = masterName;
  }
  return layout;
}

/** Derives the deck's layout catalogue (SPEC-089 AC-2): every layout reached
 * from a master's `p:sldLayoutIdLst` plus every layout a slide references,
 * keyed by absolute part path and de-duplicated. Master order first, then any
 * stray slide layout, so the gallery lists them deck-stably. */
function buildLayoutCatalogue(
  mainDoc: Document,
  mainRels: Map<string, Relationship>,
  mainPartPath: string,
  slides: Slide[],
  deps: BuildDeps,
): SlideLayout[] {
  const byKey = new Map<string, SlideLayout>();
  const add = (layoutPath: string, masterName: string | undefined): void => {
    if (byKey.has(layoutPath)) {
      return;
    }
    const entry = layoutEntryFor(layoutPath, masterName, deps);
    if (entry) {
      byKey.set(layoutPath, entry);
    }
  };
  // Every master's layouts, in master → sldLayoutIdLst order.
  const sldMasterIdLst = pChild(mainDoc.documentElement, "sldMasterIdLst");
  for (const sldMasterId of sldMasterIdLst ? childElements(sldMasterIdLst, P_NAMESPACE, "sldMasterId") : []) {
    const relId = sldMasterId.getAttributeNS(R_ATTRIBUTE_NAMESPACE, "id");
    const rel = relId ? mainRels.get(relId) : undefined;
    if (!rel || rel.type !== SLIDE_MASTER_RELATIONSHIP) {
      continue;
    }
    const masterPath = resolvePartPath(mainPartPath, rel.target);
    const masterRoot = deps.parseDoc(masterPath)?.documentElement;
    if (!masterRoot) {
      continue;
    }
    const masterRels = deps.relsFor(masterPath);
    const idLst = pChild(masterRoot, "sldLayoutIdLst");
    for (const layoutId of idLst ? childElements(idLst, P_NAMESPACE, "sldLayoutId") : []) {
      const lRelId = layoutId.getAttributeNS(R_ATTRIBUTE_NAMESPACE, "id");
      const lRel = lRelId ? masterRels.get(lRelId) : undefined;
      if (lRel && lRel.type === SLIDE_LAYOUT_RELATIONSHIP) {
        add(resolvePartPath(masterPath, lRel.target), masterPath);
      }
    }
  }
  // Any layout a slide references that no master listed (defensive).
  for (const slide of slides) {
    if (slide.layoutKey && !byKey.has(slide.layoutKey)) {
      add(slide.layoutKey, undefined);
    }
  }
  return [...byKey.values()];
}

/** Copies this slide's referenced image AND audio/video parts into the shared
 * media map (SPEC-095). A clip already in the package over `MAX_MEDIA_SIZE_BYTES`
 * is deliberately NOT taken into the model — it would bloat the single media
 * record — so it rides out through `opaque` and the reader models it as
 * `oversize` (AC-11). */
function collectSlideMedia(slideRels: Map<string, Relationship>, slidePath: string, deps: BuildDeps): void {
  for (const rel of slideRels.values()) {
    const isImage = rel.type === IMAGE_RELATIONSHIP;
    const isClip =
      rel.type === VIDEO_RELATIONSHIP || rel.type === AUDIO_RELATIONSHIP || rel.type === MEDIA_RELATIONSHIP;
    if (!isImage && !isClip) {
      continue;
    }
    const partPath = resolvePartPath(slidePath, rel.target);
    if (deps.media.has(partPath)) {
      continue;
    }
    const raw = deps.parts.get(partPath);
    if (!raw) {
      continue;
    }
    if (isClip && raw.length > MAX_MEDIA_SIZE_BYTES) {
      continue;
    }
    const contentType = contentTypeForPartName(partPath) ?? (isClip ? "application/octet-stream" : "image/png");
    deps.media.set(partPath, { mediaId: partPath, contentType, data: base64Sync(raw) });
  }
}

function base64Sync(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  // btoa exists in both the browser and Node's global scope (Vitest/jsdom).
  return typeof btoa === "function" ? btoa(binary) : Buffer.from(bytes).toString("base64");
}

function resolveChain(slidePath: string, slideRels: Map<string, Relationship>, deps: BuildDeps): ChainDocs {
  const layoutRel = Array.from(slideRels.values()).find((r) => r.type === SLIDE_LAYOUT_RELATIONSHIP);
  if (!layoutRel) {
    return {};
  }
  const layoutPath = resolvePartPath(slidePath, layoutRel.target);
  const layoutRoot = deps.parseDoc(layoutPath)?.documentElement;
  const layoutRels = deps.relsFor(layoutPath);
  const masterRel = Array.from(layoutRels.values()).find((r) => r.type === SLIDE_MASTER_RELATIONSHIP);
  let masterRoot: Element | undefined;
  let themeRoot: Element | undefined;
  let themePath: string | undefined;
  if (masterRel) {
    const masterPath = resolvePartPath(layoutPath, masterRel.target);
    masterRoot = deps.parseDoc(masterPath)?.documentElement;
    const masterRels = deps.relsFor(masterPath);
    const themeRel = Array.from(masterRels.values()).find((r) => r.type === THEME_RELATIONSHIP);
    if (themeRel) {
      themePath = resolvePartPath(masterPath, themeRel.target);
      themeRoot = deps.parseDoc(themePath)?.documentElement;
    }
  }
  return { layoutRoot, masterRoot, themeRoot, themePath };
}

/** Captures the verbatim preservation fields SPEC-081 re-emits: the root
 * `p:sld` attribute string (under `@root`, so its `xmlns`/`showMasterSp`/…
 * survive), plus `p:bg`, `p:transition`, `p:timing`, `p:clrMapOvr`,
 * `p:extLst`. Keyed by element name. */
/** A shape-tree child's own `p:cNvPr id` (the first `p:cNvPr` in it — a group's
 * own precedes its children's), used to resolve `p:timing` targets (SPEC-086).*/
function firstCNvPrId(child: Element): string | undefined {
  const list = child.getElementsByTagNameNS(P_NAMESPACE, "cNvPr");
  return list.length > 0 ? list[0].getAttribute("id") ?? undefined : undefined;
}

function collectPreserved(slideRoot: Element): Record<string, string> {
  const preserved: Record<string, string> = {};
  const rootAttrs = Array.from(slideRoot.attributes)
    .map((attr) => ` ${attr.name}="${escapeXmlAttribute(attr.value)}"`)
    .join("");
  if (rootAttrs) {
    preserved["@root"] = rootAttrs;
  }
  const cSld = pChild(slideRoot, "cSld");
  const bg = cSld && pChild(cSld, "bg");
  if (bg) {
    preserved["p:bg"] = serializeElement(bg);
  }
  for (const name of ["transition", "timing", "clrMapOvr", "extLst"]) {
    const el = pChild(slideRoot, name);
    if (el) {
      preserved[`p:${name}`] = serializeElement(el);
    }
  }
  return preserved;
}
