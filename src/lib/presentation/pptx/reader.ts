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
import { contentTypeForPartName } from "../../core/media";
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
  type ParsedPresentation,
  type PlaceholderRole,
  type Slide,
  type SlideFill,
  type SlideFrame,
  type SlideParagraph,
  type SlideRun,
} from "../model";
import { parseClrMap, parseClrScheme, resolveFirstColor, type ColorContext } from "./color";
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
 * `title`. */
function phKey(type: string | undefined, idx: string | undefined): string {
  const normalisedType = type === "ctrTitle" ? "title" : type ?? "body";
  return `${normalisedType}#${idx ?? ""}`;
}

interface PlaceholderInfo {
  xfrm?: DrawingXfrm;
  lstStyle?: Element;
  role: PlaceholderRole;
}

/** A resolved run-property set (partial: only the fields a source declares). */
interface RunProps {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: string;
  sizePt?: number;
  color?: string;
}

/** Everything a slide's frames resolve their effective values against. */
interface SlideContext {
  color: ColorContext;
  layoutPh: Map<string, PlaceholderInfo>;
  masterPh: Map<string, PlaceholderInfo>;
  /** Master `p:txStyles`: role → level (0-based) → defRPr element. */
  txStyles: Map<PlaceholderRole, Element[]>;
  themeMinorLatin?: string;
  slideRels: Map<string, Relationship>;
  slidePartPath: string;
  media: Map<string, MediaPart>;
}

/** Reads a run/defRPr-shaped element's own declared properties. */
function readRunProps(rPr: Element | undefined, color: ColorContext): RunProps {
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
    props.fontFamily = typeface;
  }
  const solidFill = firstA(rPr, "solidFill");
  if (solidFill) {
    const resolved = resolveFirstColor(solidFill, color);
    if (resolved) {
      props.color = resolved;
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
  // Built-in defaults (AC-16).
  let props: RunProps = {
    fontFamily: ctx.themeMinorLatin,
    sizePt: DEFAULT_FONT_SIZE_PT,
    color: DEFAULT_COLOR,
  };
  // Master placeholder chain: txStyles by role + level (AC-15), then the
  // layout/master placeholder's own lstStyle.
  if (role) {
    const styleLevels = ctx.txStyles.get(role);
    props = mergeProps(props, readRunProps(styleLevels?.[level], ctx.color));
  }
  // Shape's own lstStyle (AC-14).
  props = mergeProps(props, readRunProps(lstStyleDefRPr(shapeLstStyle, level), ctx.color));
  // The run's own rPr (AC-13).
  props = mergeProps(props, readRunProps(rPr, ctx.color));

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
    const color = resolveFirstColor(solidFill, ctx.color);
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
function strokeOf(spPr: Element | undefined, ctx: SlideContext): { color: string; widthPt: number } | undefined {
  if (!spPr) {
    return undefined;
  }
  const ln = firstA(spPr, "ln");
  if (!ln || firstA(ln, "noFill")) {
    return undefined;
  }
  const solidFill = firstA(ln, "solidFill");
  const color = solidFill ? resolveFirstColor(solidFill, ctx.color) : undefined;
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

/** Resolves a `p:pic` blip into a media part and an ImageFrame. */
function parsePic(pic: Element, ctx: SlideContext): SlideFrame {
  const spPr = pChild(pic, "spPr");
  const ownXfrm = parseXfrm(spPr ? firstA(spPr, "xfrm") : undefined);
  const geom = geometryFor(ownXfrm, undefined, undefined, undefined, ctx);
  const base = frameBase(geom);
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

/** Parses a `p:graphicFrame` (table/chart/other) into a foreign frame. */
function parseGraphicFrame(gf: Element, ctx: SlideContext): SlideFrame {
  const xfrm = parseXfrm(pChild(gf, "xfrm"));
  const geom = geometryFor(xfrm, undefined, undefined, undefined, ctx);
  const base = frameBase(geom);
  const graphic = firstA(gf, "graphic");
  const graphicData = graphic ? firstA(graphic, "graphicData") : undefined;
  const uri = graphicData?.getAttribute("uri") ?? "";
  let label = "Object";
  if (graphicData && firstA(graphicData, "tbl")) {
    label = "Table";
  } else if (uri.includes("chart")) {
    label = "Chart";
  }
  return foreignFrame(gf, base, label);
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

  const mainRels = relsFor(info.mainPartPath);

  // Collect every media part up-front, keyed by its own path (AC-31/32).
  const media = new Map<string, MediaPart>();

  const slides: Slide[] = [];
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
    slides.push(
      buildSlide(slideDoc, slidePath, {
        readPart,
        parseDoc,
        relsFor,
        media,
        parts,
      }),
    );
  }

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

  return { slides, slideWidthCm, slideHeightCm, media: [...media.values()], opaque, properties };
}

interface BuildDeps {
  readPart: (path: string) => string | undefined;
  parseDoc: (path: string) => Document | undefined;
  relsFor: (path: string) => Map<string, Relationship>;
  media: Map<string, MediaPart>;
  parts: Map<string, Uint8Array>;
}

/** Resolves a slide's layout/master/theme chain and parses its shape tree. */
function buildSlide(slideDoc: Document, slidePath: string, deps: BuildDeps): Slide {
  const slideRoot = slideDoc.documentElement;
  const slideRels = deps.relsFor(slidePath);

  // slide → layout → master → theme, each through its own .rels (AC: never
  // assume part paths).
  const chain = resolveChain(slidePath, slideRels, deps);

  const color: ColorContext = {
    scheme: chain.themeRoot ? parseClrScheme(chain.themeRoot) : {},
    map: chain.masterRoot ? parseClrMap(chain.masterRoot) : {},
  };

  const ctx: SlideContext = {
    color,
    layoutPh: buildPlaceholderMap(chain.layoutRoot ? layoutSpTree(chain.layoutRoot) : undefined),
    masterPh: buildPlaceholderMap(chain.masterRoot ? layoutSpTree(chain.masterRoot) : undefined),
    txStyles: buildTxStyles(chain.masterRoot),
    themeMinorLatin: themeMinorLatin(chain.themeRoot),
    slideRels,
    slidePartPath: slidePath,
    media: deps.media,
  };

  // Populate media referenced by this slide's image rels.
  collectSlideMedia(slideRels, slidePath, deps);

  const cSld = pChild(slideRoot, "cSld");
  const spTree = cSld && pChild(cSld, "spTree");
  const frames: SlideFrame[] = [];
  // `p:cNvPr id` → minted frame id, so a `p:timing`'s `p:spTgt spid` resolves
  // back to the frame it targets (SPEC-086).
  const frameForSpid = new Map<string, string>();
  if (spTree) {
    for (const child of Array.from(spTree.children)) {
      if (child.namespaceURI !== P_NAMESPACE) {
        continue;
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
  if (Object.keys(preserved).length > 0) {
    slide.preserved = preserved;
  }
  const notes = readNotes(slideRels, slidePath, deps.parseDoc);
  if (notes) {
    slide.notes = notes;
  }
  return slide;
}

function layoutSpTree(root: Element): Element | undefined {
  const cSld = pChild(root, "cSld");
  return cSld ? pChild(cSld, "spTree") : undefined;
}

/** Copies this slide's referenced image parts into the shared media map. */
function collectSlideMedia(slideRels: Map<string, Relationship>, slidePath: string, deps: BuildDeps): void {
  for (const rel of slideRels.values()) {
    if (rel.type !== IMAGE_RELATIONSHIP) {
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
    const contentType = contentTypeForPartName(partPath) ?? "image/png";
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
  if (masterRel) {
    const masterPath = resolvePartPath(layoutPath, masterRel.target);
    masterRoot = deps.parseDoc(masterPath)?.documentElement;
    const masterRels = deps.relsFor(masterPath);
    const themeRel = Array.from(masterRels.values()).find((r) => r.type === THEME_RELATIONSHIP);
    if (themeRel) {
      themeRoot = deps.parseDoc(resolvePartPath(masterPath, themeRel.target))?.documentElement;
    }
  }
  return { layoutRoot, masterRoot, themeRoot };
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
