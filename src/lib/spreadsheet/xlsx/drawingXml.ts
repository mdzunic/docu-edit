// The xlsx sheet drawing layer on the wire (SPEC-131 "Reading xlsx
// drawings", "MCE wrappers", "Writing", "Canonical new drawings"): THIS FILE
// is the normative source for how a worksheet's `xdr:wsDr` drawing part
// becomes Sheet.drawings and how a touched layer is written back.
//
// Reading. The worksheet's `<drawing r:id>` (a CT_Worksheet child) is
// followed through the sheet's relationships to its `xdr:wsDr` part. The part
// is parsed with parseXml for content and its top-level children are sliced
// with the scanner (xlsx/elementTree.ts) into each drawing's `sourceXml` — the
// child with the whitespace before it. Every top-level child is a drawing,
// ids "dr1", "dr2"… in document order:
// | Child | Anchor |
// | xdr:twoCellAnchor | twoCell, `editAs` (absent → "twoCell", always set) |
// | xdr:oneCellAnchor | oneCell: from + ext |
// | xdr:absoluteAnchor | absolute: pos + ext |
// | mc:AlternateContent | the first branch (Choices, then the Fallback) holding an anchor |
// | anything else | a hidden foreign "Object" at 0,0 — kept, never painted |
// Every anchor goes through drawings.ts withDerivedSize. What an anchor holds:
// | Element | Content |
// | xdr:pic | image: `a:blip@r:embed` → the part's rels → one MediaPart per package path, shared (mediaId = the path, the bytes deferred with registerPendingMedia); `a:srcRect` → crop; `r:link` → foreign "Linked picture" (the embedded copy its face); an unresolvable embed → foreign "Picture" |
// | xdr:sp, xdr:cxnSp | shape: `a:prstGeom@prst` via PRST_TO_GEOMETRY (unknown or `a:custGeom` → "rect" + unsupportedGeometryXml), avLst adjustments, fill and outline from `xdr:spPr` else the `xdr:style` refs, `xdr:txBody` → DrawingText |
// | xdr:grpSp | group: members normalised to the group box through parseGroupChildMapping/mapGroupChild; a chart frame → foreign "Chart"; a nested group or any other frame → foreign "Object"; a hidden member is left out |
// | xdr:graphicFrame | chart (SPEC-132): `c:chart@r:id` → the part, read by xlsx/chartXml.ts readXlsxChart into a SheetChartContent — `visual` (out of catalogue: unsupportedChartXml), `refs`, `readOnlyReason`, `sourcePart` (never `uncached`: open re-derives it); a relationship or part that does not resolve → foreign "Chart"; SmartArt → foreign "SmartArt", else foreign "Object" |
// | xdr:contentPart | foreign "Ink" |
// `xdr:cNvPr` gives name / descr (altText) / hidden; `a:xfrm@rot` rotationDeg.
//
// Chart sheets (SPEC-132 "Chart sheets"). A chartsheet's `<drawing r:id>` is
// followed the same way (readXlsxChartSheet): when its part holds exactly one
// top-level child, an anchor (Excel writes `xdr:absoluteAnchor`) whose
// graphicFrame reads as a chart, that drawing is the chart sheet's one
// drawing; anything else leaves the sheet foreign. The part a chart drawing
// relates is the one its anchor's `c:chart@r:id` names in its sheet's drawing
// part's `.rels` (sheetChartPart) — its `sourcePart` for a drawing that has
// no retained XML.
//
// MCE. UNDERSTOOD_SHEET_DRAWING_NAMESPACES is empty: selectAlternateContent
// always reads the Fallback, and the drawing is `foreign` with the WHOLE
// wrapper as its sourceXml. Label by the first Choice's `Requires`
// namespaces: chartex (cx1, cx2…) → "Chart"; sle15 → "Slicer"; aink → "Ink";
// a14 with an `a14:compatExt` → "Form control", with an `a14:m` → "Equation";
// else "Object". name/altText come from the first branch whose cNvPr has a
// name; `hidden` only from the branch read (Excel's form control hides its
// Choice stand-in, its real face being VML). The face is the read branch's
// `xdr:pic` or catalogue `xdr:sp` unless hidden — never for chartex.
//
// Writing (writeDrawingPart). A `drawingsDirty` sheet's part is rebuilt as
// the source up to and including the `xdr:wsDr` start tag, each drawing's
// anchor element in z-order, and the source from its last child on. An
// imported drawing is its sourceXml with only these splices: `xdr:from`/
// `xdr:to` rewritten (`<xdr:col>…<xdr:colOff>…<xdr:row>…<xdr:rowOff>…`, the
// source's prefix) where drawings.ts writtenPoints differs from the source,
// `xdr:pos`/`xdr:ext` attributes, `editAs`, `cNvPr` name/descr/hidden where
// the model differs from what reading the source gives, and — once the
// anchor moved — the content element's own `a:off`/`a:ext` (in `xdr:spPr` or
// `xdr:grpSpPr`, `xdr:xfrm` in a frame) set to the anchor's EMU box unless
// all four were 0 in the source. Inside an `mc:AlternateContent` every
// branch is patched alike. The content element is never regenerated; an
// untouched drawing is byte-identical. A drawing created here is written in
// the canonical form xlsx/drawingGolden.ts fixes — SPEC-132: a created chart
// too, its anchor holding chartXml.ts's chartGraphicFrameXml (the alt text
// and hidden flag spliced into its cNvPr) around the part the writer's chart
// planner (xlsx/chartParts.ts) names for it. The part's `.rels` keeps a
// relationship unless the source's drawings used it and no remaining one
// does, and gains an image relationship per new picture part and a chart
// relationship per new chart part (rId = 1 + highest). A chart sheet's
// drawing part is never rewritten: only its chart part is, by the chart
// write paths.
import { CHART_RELATIONSHIP } from "../../core/chartPartOoxml";
import { DRAWINGML_NAMESPACE, GEOMETRY_TO_PRST, mapGroupChild, parseAvLstAdjustments, parseGroupChildMapping, PRST_TO_GEOMETRY, type GroupChildMapping } from "../../core/drawingMl";
import { contentTypeForPartName, registerPendingMedia, type MediaPart } from "../../core/media";
import { adjustmentGuideName, adjustmentToGuideValue } from "../../core/shapeCatalogue";
import { shapeFillBaseColor, type ShapeFill, type ShapeOutline, type ShapeVisual } from "../../core/shapes";
import type { ModelColor } from "../../core/theme";
import { schemeClrRefFrom, schemeClrXmlFor, schemeSlotFromVal, type DrawingMlColorRef } from "../../core/themeDrawingMl";
import { MARKUP_COMPATIBILITY_NAMESPACE, parseXml, relsPathFor, selectAlternateContent, serializeElement, XmlParseError } from "../../core/xml";
import { scanXml, XmlScanError, type XmlScanAttribute } from "../../core/xmlScanner";
import { anchorBehaviour, drawingRectPx, EMU_PER_PX, pointFromPx, withDerivedSize, writtenPoints } from "../drawings";
import { buildGridMetrics, type GridMetrics } from "../geometry";
import type { AnchorPoint, DrawingAnchor, DrawingText, Sheet, SheetChartContent, SheetDrawing, SheetDrawingContent } from "../model";
import { applySplices, escapeAttributeText, escapeAttributeValue, escapeElementText, setAttribute, type Splice } from "../xmlSplice";
import { chartGraphicFrameXml, readXlsxChart } from "./chartXml";
import { elementTree, type XmlNode } from "./elementTree";
import {
  contentTypeIn,
  nextRelationshipId,
  partTextOf,
  readContentTypes,
  readRels,
  relationshipKind,
  relationshipsOf,
  relationshipXml,
  relativeTarget,
  spliceRels,
  targetPart,
  type ContentTypes,
  type PartLookup,
  type RelEntry,
} from "./package";

// --- namespaces and constants ------------------------------------------------------------

const XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const XDR_NAMESPACES: ReadonlySet<string> = new Set([XDR_NS, "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"]);
const A_NS = DRAWINGML_NAMESPACE;
const A_NAMESPACES: ReadonlySet<string> = new Set([A_NS, "http://purl.oclc.org/ooxml/drawingml/main"]);
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_NAMESPACES: ReadonlySet<string> = new Set([REL_NS, "http://purl.oclc.org/ooxml/officeDocument/relationships"]);
const SPREADSHEETML_NAMESPACES: ReadonlySet<string> = new Set(["http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main"]);
const MC_NS = MARKUP_COMPATIBILITY_NAMESPACE;
const CHART_URIS: ReadonlySet<string> = new Set(["http://schemas.openxmlformats.org/drawingml/2006/chart", "http://purl.oclc.org/ooxml/drawingml/chart"]);
const DIAGRAM_URIS: ReadonlySet<string> = new Set(["http://schemas.openxmlformats.org/drawingml/2006/diagram", "http://purl.oclc.org/ooxml/drawingml/diagram"]);
const A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main";
const SLE15_NS = "http://schemas.microsoft.com/office/drawing/2012/slicer";
const AINK_NS = "http://schemas.microsoft.com/office/drawing/2016/ink";

/** The Office-2010+ drawing namespaces whose `mc:Choice` this product reads
 * instead of the Fallback: none (see the module header). A future spec that
 * renders a wrapped kind better than its Fallback adds its namespace here. */
export const UNDERSTOOD_SHEET_DRAWING_NAMESPACES: ReadonlySet<string> = new Set<string>();

/** A drawing part's content type (its `[Content_Types].xml` Override). */
export const DRAWING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml";
/** The worksheet → drawing relationship type. */
export const DRAWING_RELATIONSHIP_TYPE = `${REL_NS}/drawing`;
/** The drawing → picture relationship type. */
export const IMAGE_RELATIONSHIP_TYPE = `${REL_NS}/image`;

const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const ANCHOR_LOCALS: ReadonlySet<string> = new Set(["twoCellAnchor", "oneCellAnchor", "absoluteAnchor"]);
const CONTENT_LOCALS: ReadonlySet<string> = new Set(["pic", "sp", "cxnSp", "grpSp", "graphicFrame", "contentPart"]);
const FILL_LOCALS: ReadonlySet<string> = new Set(["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"]);
/** EMU per point (12,700) — `a:ln@w`. */
const EMU_PER_PT = 12700;
/** The Office theme's line styles (`a:lnStyleLst`), by `lnRef@idx`, pt. */
const THEME_LINE_WIDTH_PT: Readonly<Record<number, number>> = { 1: 0.5, 2: 1, 3: 1.5 };
/** DrawingML's line width when `a:ln` gives none and no style does: 0.75 pt. */
const DEFAULT_LINE_WIDTH_PT = 0.75;

/** `a:prstDash@val` → the model's dash. */
const DASH_BY_PRESET: Readonly<Record<string, ShapeOutline["dash"]>> = {
  solid: "solid",
  dash: "dash",
  sysDash: "dash",
  dot: "dot",
  sysDot: "dot",
  dashDot: "dashDot",
  sysDashDot: "dashDot",
  sysDashDotDot: "dashDot",
  lgDash: "longDash",
  lgDashDot: "longDashDot",
  lgDashDotDot: "longDashDot",
};
/** The model's dash → the `a:prstDash@val` a canonical outline writes. */
const PRESET_BY_DASH: Readonly<Record<ShapeOutline["dash"], string | undefined>> = {
  solid: undefined,
  dash: "dash",
  dot: "sysDot",
  dashDot: "dashDot",
  longDash: "lgDash",
  longDashDot: "lgDashDot",
};

/** The common `a:prstClr` names (ST_PresetColorVal) → hex — Excel's slicer
 * Fallback paints `white` on `green`. Others read as no colour. */
const PRESET_COLORS: Readonly<Record<string, string>> = {
  black: "#000000",
  white: "#ffffff",
  red: "#ff0000",
  green: "#008000",
  lime: "#00ff00",
  blue: "#0000ff",
  yellow: "#ffff00",
  cyan: "#00ffff",
  aqua: "#00ffff",
  magenta: "#ff00ff",
  fuchsia: "#ff00ff",
  gray: "#808080",
  grey: "#808080",
  silver: "#c0c0c0",
  maroon: "#800000",
  navy: "#000080",
  olive: "#808000",
  purple: "#800080",
  teal: "#008080",
  orange: "#ffa500",
  ltGray: "#d3d3d3",
  dkGray: "#a9a9a9",
};

// --- DOM helpers ------------------------------------------------------------------------------

function isXdr(el: Element, local?: string): boolean {
  return el.namespaceURI !== null && XDR_NAMESPACES.has(el.namespaceURI) && (local === undefined || el.localName === local);
}

function isA(el: Element, local?: string): boolean {
  return el.namespaceURI !== null && A_NAMESPACES.has(el.namespaceURI) && (local === undefined || el.localName === local);
}

function isMc(el: Element, local: string): boolean {
  return el.namespaceURI === MC_NS && el.localName === local;
}

function xdrChild(parent: Element, local: string): Element | undefined {
  return Array.from(parent.children).find((child) => isXdr(child, local));
}

function aChild(parent: Element, local: string): Element | undefined {
  return Array.from(parent.children).find((child) => isA(child, local));
}

function aChildren(parent: Element, local: string): Element[] {
  return Array.from(parent.children).filter((child) => isA(child, local));
}

/** An attribute in the officeDocument relationships namespace. */
function relAttr(el: Element, local: string): string | undefined {
  for (const attr of Array.from(el.attributes)) {
    if (attr.localName === local && attr.namespaceURI !== null && REL_NAMESPACES.has(attr.namespaceURI) && attr.value !== "") {
      return attr.value;
    }
  }
  return undefined;
}

const isOn = (value: string | null | undefined): boolean => value === "1" || value === "true";

/** A finite number attribute (undefined when absent or not a number). */
function numberAttr(el: Element | undefined, name: string): number | undefined {
  const raw = el?.getAttribute(name);
  if (raw === null || raw === undefined || raw.trim() === "") {
    return undefined;
  }
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

/** EMU per universal-measure unit (strict ST_Coordinate: "2.5cm", "1in"…). */
const EMU_PER_UNIT: Readonly<Record<string, number>> = { mm: 36000, cm: 360000, in: 914400, pt: 12700, pc: 152400, pi: 152400 };

/** An ST_Coordinate as whole EMU (0 when unreadable). */
function coordinate(text: string | null | undefined): number {
  const raw = (text ?? "").trim();
  const plain = Number(raw);
  if (raw !== "" && Number.isFinite(plain)) {
    return Math.round(plain);
  }
  const measure = /^(-?\d+(?:\.\d+)?)(mm|cm|in|pt|pc|pi)$/.exec(raw);
  return measure ? Math.round(Number(measure[1]) * EMU_PER_UNIT[measure[2]!]!) : 0;
}

// --- reading: anchors -------------------------------------------------------------------------

const ZERO_POINT: AnchorPoint = { col: 0, colOffEmu: 0, row: 0, rowOffEmu: 0 };

function readPoint(el: Element | undefined): AnchorPoint {
  if (!el) {
    return { ...ZERO_POINT };
  }
  const part = (local: string): number => Math.max(0, coordinate(xdrChild(el, local)?.textContent));
  return { col: part("col"), colOffEmu: part("colOff"), row: part("row"), rowOffEmu: part("rowOff") };
}

function readExt(el: Element | undefined): { cx: number; cy: number } {
  return { cx: Math.max(0, coordinate(el?.getAttribute("cx"))), cy: Math.max(0, coordinate(el?.getAttribute("cy"))) };
}

const EDIT_AS: ReadonlySet<string> = new Set(["twoCell", "oneCell", "absolute"]);

/** An anchor element's anchor, as written (before withDerivedSize). */
function readAnchorElement(el: Element): DrawingAnchor {
  if (el.localName === "oneCellAnchor") {
    return { kind: "oneCell", from: readPoint(xdrChild(el, "from")), extEmu: readExt(xdrChild(el, "ext")) };
  }
  if (el.localName === "absoluteAnchor") {
    const pos = xdrChild(el, "pos");
    return { kind: "absolute", posEmu: { x: coordinate(pos?.getAttribute("x")), y: coordinate(pos?.getAttribute("y")) }, extEmu: readExt(xdrChild(el, "ext")) };
  }
  const editAs = el.getAttribute("editAs") ?? "twoCell";
  return {
    kind: "twoCell",
    editAs: (EDIT_AS.has(editAs) ? editAs : "twoCell") as DrawingAnchor["editAs"],
    from: readPoint(xdrChild(el, "from")),
    to: readPoint(xdrChild(el, "to")),
  };
}

const isAnchorElement = (el: Element): boolean => isXdr(el) && ANCHOR_LOCALS.has(el.localName);
const isContentElement = (el: Element): boolean => isXdr(el) && CONTENT_LOCALS.has(el.localName);

/** The anchor's content: its drawing element or an `mc:AlternateContent`. */
function anchorContent(anchor: Element): Element | undefined {
  return Array.from(anchor.children).find((child) => isContentElement(child) || isMc(child, "AlternateContent"));
}

/** An unanchored element's stand-in anchor: absolute at 0,0, no size. */
const unanchored = (): DrawingAnchor => ({ kind: "absolute", posEmu: { x: 0, y: 0 }, extEmu: { cx: 0, cy: 0 } });

// --- reading: base fields ---------------------------------------------------------------------

interface DrawingBase {
  name: string;
  altText?: string;
  hidden?: true;
}

/** A drawing element's `cNvPr` (inside its `nv…Pr`, any namespace). */
function cNvPrOf(content: Element | undefined): Element | undefined {
  const nv = content ? Array.from(content.children).find((child) => /^nv\w*Pr$/.test(child.localName)) : undefined;
  return nv ? Array.from(nv.children).find((child) => child.localName === "cNvPr") : undefined;
}

function baseOf(cNvPr: Element | undefined): DrawingBase {
  const base: DrawingBase = { name: cNvPr?.getAttribute("name") ?? "" };
  const descr = cNvPr?.getAttribute("descr");
  if (descr) {
    base.altText = descr;
  }
  if (isOn(cNvPr?.getAttribute("hidden"))) {
    base.hidden = true;
  }
  return base;
}

/** A drawing element's own `a:xfrm` (`xdr:xfrm` for a frame). */
function xfrmOf(content: Element): Element | undefined {
  const own = Array.from(content.children).find((child) => child.localName === "xfrm");
  if (own) {
    return own;
  }
  const props = xdrChild(content, content.localName === "grpSp" ? "grpSpPr" : "spPr");
  return props ? aChild(props, "xfrm") : undefined;
}

/** `xfrm@rot` in whole degrees clockwise, undefined when 0. */
function rotationOf(content: Element): number | undefined {
  const rot = numberAttr(xfrmOf(content), "rot");
  const deg = rot === undefined ? 0 : Math.round(rot / 60000);
  return deg === 0 ? undefined : deg;
}

/** An `xfrm`'s `a:off`/`a:ext` box, EMU (undefined when it has neither). */
function xfrmBox(xfrm: Element | undefined): { x: number; y: number; cx: number; cy: number } | undefined {
  const off = xfrm ? aChild(xfrm, "off") : undefined;
  const ext = xfrm ? aChild(xfrm, "ext") : undefined;
  if (!off && !ext) {
    return undefined;
  }
  return { x: numberAttr(off, "x") ?? 0, y: numberAttr(off, "y") ?? 0, cx: numberAttr(ext, "cx") ?? 0, cy: numberAttr(ext, "cy") ?? 0 };
}

// --- reading: colours, fills, outlines, text ---------------------------------------------------

const hexColor = (value: string | null | undefined): string | undefined => (value && /^[0-9a-fA-F]{6}$/.test(value) ? `#${value.toLowerCase()}` : undefined);

/** The colour an element's first DrawingML colour child names. */
function colorIn(parent: Element | undefined): ModelColor | undefined {
  for (const child of parent ? Array.from(parent.children) : []) {
    if (!isA(child)) {
      continue;
    }
    switch (child.localName) {
      case "srgbClr":
        return hexColor(child.getAttribute("val"));
      case "schemeClr": {
        const ref = schemeClrRefFrom(child);
        if (ref) {
          return ref;
        }
        const slot = schemeSlotFromVal(child.getAttribute("val"));
        return slot ? { theme: slot } : undefined;
      }
      case "sysClr":
        return hexColor(child.getAttribute("lastClr")) ?? (child.getAttribute("val") === "window" ? "#ffffff" : "#000000");
      case "prstClr":
        return PRESET_COLORS[child.getAttribute("val") ?? ""];
      case "scrgbClr": {
        const channel = (name: string): string =>
          Math.round((Math.min(100000, Math.max(0, numberAttr(child, name) ?? 0)) / 100000) * 255)
            .toString(16)
            .padStart(2, "0");
        return `#${channel("r")}${channel("g")}${channel("b")}`;
      }
      default:
        continue;
    }
  }
  return undefined;
}

/** The colour a fill element paints with: solid's own, a gradient's first
 * stop, a pattern's foreground. */
function fillColor(fill: Element): ModelColor | undefined {
  if (isA(fill, "solidFill")) {
    return colorIn(fill);
  }
  if (isA(fill, "gradFill")) {
    const stops = aChild(fill, "gsLst");
    const first = stops ? aChildren(stops, "gs").toSorted((a, b) => (numberAttr(a, "pos") ?? 0) - (numberAttr(b, "pos") ?? 0))[0] : undefined;
    return colorIn(first);
  }
  if (isA(fill, "pattFill")) {
    return colorIn(aChild(fill, "fgClr"));
  }
  return undefined;
}

const firstFill = (parent: Element | undefined): Element | undefined => (parent ? Array.from(parent.children).find((child) => isA(child) && FILL_LOCALS.has(child.localName)) : undefined);

/** A shape's fill: `xdr:spPr`'s own, else its style's `fillRef`. */
function fillFrom(spPr: Element | undefined, style: Element | undefined): ShapeFill | undefined {
  const own = firstFill(spPr);
  if (own) {
    if (isA(own, "noFill")) {
      return { kind: "none" };
    }
    const color = fillColor(own);
    return color === undefined ? undefined : { kind: "solid", color };
  }
  const ref = style ? aChild(style, "fillRef") : undefined;
  if (!ref || (numberAttr(ref, "idx") ?? 0) === 0) {
    return undefined;
  }
  const color = colorIn(ref);
  return color === undefined ? undefined : { kind: "solid", color };
}

/** A shape's outline: `a:ln` (its width, fill and dash), its colour from
 * the style's `lnRef` when `a:ln` has no fill of its own. */
function outlineFrom(spPr: Element | undefined, style: Element | undefined): ShapeOutline | undefined {
  const ln = spPr ? aChild(spPr, "ln") : undefined;
  const ref = style ? aChild(style, "lnRef") : undefined;
  const refIdx = numberAttr(ref, "idx") ?? 0;
  const lnFill = firstFill(ln);
  if (lnFill && isA(lnFill, "noFill")) {
    return undefined;
  }
  const color = lnFill ? fillColor(lnFill) : refIdx > 0 ? colorIn(ref) : undefined;
  if (color === undefined) {
    return undefined;
  }
  const w = numberAttr(ln, "w");
  const dashEl = ln ? aChild(ln, "prstDash") : undefined;
  return {
    color,
    widthPt: w !== undefined ? w / EMU_PER_PT : (THEME_LINE_WIDTH_PT[refIdx] ?? DEFAULT_LINE_WIDTH_PT),
    dash: DASH_BY_PRESET[dashEl?.getAttribute("val") ?? "solid"] ?? "solid",
  };
}

/** A paragraph's runs, in order, `mc:AlternateContent` read through its
 * Fallback: the text and the first `a:r`. */
function paragraphText(p: Element, onRun: (run: Element) => void): string {
  let text = "";
  const visit = (children: readonly Element[]): void => {
    for (const child of children) {
      if (isA(child, "r") || isA(child, "fld")) {
        text += aChild(child, "t")?.textContent ?? "";
        if (isA(child, "r")) {
          onRun(child);
        }
      } else if (isA(child, "br")) {
        text += "\n";
      } else if (isMc(child, "AlternateContent")) {
        visit(selectAlternateContent(child, UNDERSTOOD_SHEET_DRAWING_NAMESPACES));
      }
    }
  };
  visit(Array.from(p.children));
  return text;
}

const ALIGN: Readonly<Record<string, "left" | "center" | "right">> = { l: "left", ctr: "center", r: "right" };

/** `xdr:txBody` → DrawingText; undefined when every paragraph is empty. */
function textFrom(txBody: Element, style: Element | undefined): DrawingText | undefined {
  let firstRun: Element | undefined;
  const paragraphs: DrawingText["paragraphs"] = aChildren(txBody, "p").map((p) => {
    const text = paragraphText(p, (run) => {
      firstRun ??= run;
    });
    const pPr = aChild(p, "pPr");
    const align = ALIGN[pPr?.getAttribute("algn") ?? ""];
    return align ? { text, align } : { text };
  });
  if (paragraphs.every((paragraph) => paragraph.text === "")) {
    return undefined;
  }
  const anchorAttr = aChild(txBody, "bodyPr")?.getAttribute("anchor");
  const text: DrawingText = { paragraphs, anchor: anchorAttr === "ctr" ? "middle" : anchorAttr === "b" ? "bottom" : "top" };
  const rPr = firstRun ? aChild(firstRun, "rPr") : undefined;
  const sz = numberAttr(rPr, "sz");
  if (sz !== undefined) {
    text.sizePt = sz / 100;
  }
  const runFill = firstFill(rPr);
  const color = (runFill ? fillColor(runFill) : undefined) ?? colorIn(style ? aChild(style, "fontRef") : undefined);
  if (color !== undefined) {
    text.color = color;
  }
  if (isOn(rPr?.getAttribute("b"))) {
    text.bold = true;
  }
  return text;
}

/** A shape's visual subset ("The drawing model"). */
function shapeVisualOf(content: Element): ShapeVisual {
  const spPr = xdrChild(content, "spPr");
  const style = xdrChild(content, "style");
  const prstGeom = spPr ? aChild(spPr, "prstGeom") : undefined;
  const custGeom = spPr ? aChild(spPr, "custGeom") : undefined;
  const prst = prstGeom?.getAttribute("prst") ?? "";
  let visual: ShapeVisual;
  if (prstGeom && Object.hasOwn(PRST_TO_GEOMETRY, prst)) {
    visual = { geometry: PRST_TO_GEOMETRY[prst]! };
    const adjustments = parseAvLstAdjustments(prstGeom);
    if (adjustments.length > 0) {
      visual.adjustments = adjustments;
    }
  } else if (prstGeom || custGeom) {
    visual = { geometry: "rect", unsupportedGeometryXml: serializeElement((prstGeom ?? custGeom)!) };
  } else {
    visual = { geometry: "rect" };
  }
  const fill = fillFrom(spPr, style);
  if (fill) {
    visual.fill = fill;
  }
  const outline = outlineFrom(spPr, style);
  if (outline) {
    visual.outline = outline;
  }
  return visual;
}

// --- reading: content ------------------------------------------------------------------------

/** What reading one drawing part needs. */
interface ReadContext {
  parts: PartLookup;
  /** The drawing part's name and its relationships by id. */
  drawingPart: string;
  rels: ReadonlyMap<string, RelEntry>;
  /** The package's images, one MediaPart per path, shared across sheets. */
  media: Map<string, MediaPart>;
  contentTypes(): ContentTypes | undefined;
}

/** The shared MediaPart of a picture relationship's target, registered on
 * first sight (its bytes deferred); undefined when it does not resolve. */
function mediaFor(rId: string, ctx: ReadContext): MediaPart | undefined {
  const rel = ctx.rels.get(rId);
  const path = rel ? targetPart(ctx.drawingPart, rel) : undefined;
  if (path === undefined) {
    return undefined;
  }
  const known = ctx.media.get(path);
  if (known) {
    return known;
  }
  const bytes = ctx.parts.get(path);
  if (!bytes) {
    return undefined;
  }
  const types = ctx.contentTypes();
  const part: MediaPart = { mediaId: path, contentType: (types && contentTypeIn(types, path)) ?? contentTypeForPartName(path) ?? "application/octet-stream", data: "" };
  registerPendingMedia(part, bytes);
  ctx.media.set(path, part);
  return part;
}

type PaintContent = Exclude<SheetDrawingContent, { kind: "group" | "chart" | "foreign" }>;
type MemberContent = Exclude<SheetDrawingContent, { kind: "group" | "chart" }>;

/** `a:srcRect` (1/1000 of a percent per side) → crop fractions. */
function cropOf(blipFill: Element): { l: number; t: number; r: number; b: number } | undefined {
  const rect = aChild(blipFill, "srcRect");
  if (!rect) {
    return undefined;
  }
  const side = (name: string): number => (numberAttr(rect, name) ?? 0) / 100000;
  const crop = { l: side("l"), t: side("t"), r: side("r"), b: side("b") };
  return crop.l === 0 && crop.t === 0 && crop.r === 0 && crop.b === 0 ? undefined : crop;
}

function readPicture(pic: Element, ctx: ReadContext): Extract<SheetDrawingContent, { kind: "image" | "foreign" }> {
  const blipFill = xdrChild(pic, "blipFill");
  const blip = blipFill ? aChild(blipFill, "blip") : undefined;
  const embed = blip ? relAttr(blip, "embed") : undefined;
  const media = embed === undefined ? undefined : mediaFor(embed, ctx);
  let image: Extract<SheetDrawingContent, { kind: "image" }> | undefined;
  if (media) {
    image = { kind: "image", mediaId: media.mediaId };
    const crop = blipFill ? cropOf(blipFill) : undefined;
    if (crop) {
      image.crop = crop;
    }
  }
  if (blip && relAttr(blip, "link") !== undefined) {
    return image ? { kind: "foreign", label: "Linked picture", face: image } : { kind: "foreign", label: "Linked picture" };
  }
  return image ?? { kind: "foreign", label: "Picture" };
}

function readShape(sp: Element): Extract<SheetDrawingContent, { kind: "shape" }> {
  const visual = shapeVisualOf(sp);
  const txBody = xdrChild(sp, "txBody");
  const text = txBody ? textFrom(txBody, xdrChild(sp, "style")) : undefined;
  return text ? { kind: "shape", visual, text } : { kind: "shape", visual };
}

function readGraphicFrame(frame: Element, ctx: ReadContext): SheetDrawingContent {
  const graphic = aChild(frame, "graphic");
  const data = graphic ? aChild(graphic, "graphicData") : undefined;
  const uri = data?.getAttribute("uri") ?? "";
  if (CHART_URIS.has(uri)) {
    const chart = data ? Array.from(data.children).find((child) => child.localName === "chart") : undefined;
    const rId = chart ? relAttr(chart, "id") : undefined;
    const rel = rId === undefined ? undefined : ctx.rels.get(rId);
    const sourcePart = rel ? targetPart(ctx.drawingPart, rel) : undefined;
    const partText = sourcePart === undefined ? undefined : partTextOf(ctx.parts, sourcePart);
    if (sourcePart === undefined || partText === undefined) {
      return { kind: "foreign", label: "Chart" };
    }
    const read = readXlsxChart(partText);
    const content: SheetChartContent = { kind: "chart", visual: read.visual, sourcePart };
    if (read.refs) {
      content.refs = read.refs;
    }
    if (read.readOnlyReason !== undefined) {
      content.readOnlyReason = read.readOnlyReason;
    }
    return content;
  }
  return { kind: "foreign", label: DIAGRAM_URIS.has(uri) ? "SmartArt" : "Object" };
}

/** A group member's box in the group's child space, EMU. */
function memberBox(member: Element): { x: number; y: number; cx: number; cy: number } | undefined {
  if (isMc(member, "AlternateContent")) {
    const branches = Array.from(member.children).filter((child) => isMc(child, "Choice") || isMc(child, "Fallback"));
    for (const branch of branches) {
      const inner = Array.from(branch.children).find(isContentElement);
      const box = inner ? xfrmBox(xfrmOf(inner)) : undefined;
      if (box) {
        return box;
      }
    }
    return undefined;
  }
  return xfrmBox(xfrmOf(member));
}

/** A group member's content (see the module header). */
function memberContent(member: Element, ctx: ReadContext): MemberContent {
  if (isXdr(member, "pic")) {
    return readPicture(member, ctx);
  }
  if (isXdr(member, "sp") || isXdr(member, "cxnSp")) {
    const shape = readShape(member);
    const rotation = rotationOf(member);
    return rotation === undefined ? shape : { ...shape, visual: { ...shape.visual, rotationDeg: rotation } };
  }
  if (isXdr(member, "graphicFrame") && isChartFrame(member)) {
    return { kind: "foreign", label: "Chart" }; // a group member is never a modelled chart
  }
  return { kind: "foreign", label: isXdr(member, "contentPart") ? "Ink" : "Object" };
}

/** Whether a graphicFrame holds a DrawingML chart. */
function isChartFrame(frame: Element): boolean {
  const graphic = aChild(frame, "graphic");
  const data = graphic ? aChild(graphic, "graphicData") : undefined;
  return CHART_URIS.has(data?.getAttribute("uri") ?? "");
}

function readGroup(grpSp: Element, ctx: ReadContext): Extract<SheetDrawingContent, { kind: "group" }> {
  const grpSpPr = xdrChild(grpSp, "grpSpPr");
  const mapping: GroupChildMapping | undefined = grpSpPr ? parseGroupChildMapping(grpSpPr) : undefined;
  const memberElements = Array.from(grpSp.children).filter((child) => isContentElement(child) || isMc(child, "AlternateContent"));
  const read = memberElements.flatMap((member) => {
    const box = memberBox(member);
    const hidden = isContentElement(member) && isOn(cNvPrOf(member)?.getAttribute("hidden"));
    if (!box || hidden) {
      return [];
    }
    const mapped = mapping ? mapGroupChild(mapping, box.x, box.y, box.cx, box.cy) : { xEmu: box.x, yEmu: box.y, cxEmu: box.cx, cyEmu: box.cy };
    return [{ mapped, content: memberContent(member, ctx) }];
  });
  // The group's box: its own off/ext, else the members' union.
  const own = grpSpPr ? xfrmBox(aChild(grpSpPr, "xfrm")) : undefined;
  let frame = own && own.cx > 0 && own.cy > 0 ? own : undefined;
  if (!frame && read.length > 0) {
    const left = Math.min(...read.map((m) => m.mapped.xEmu));
    const top = Math.min(...read.map((m) => m.mapped.yEmu));
    const right = Math.max(...read.map((m) => m.mapped.xEmu + m.mapped.cxEmu));
    const bottom = Math.max(...read.map((m) => m.mapped.yEmu + m.mapped.cyEmu));
    frame = { x: left, y: top, cx: right - left, cy: bottom - top };
  }
  const box = frame ?? { x: 0, y: 0, cx: 0, cy: 0 };
  return {
    kind: "group",
    members: read.map(({ mapped, content }) => ({
      x: fraction(mapped.xEmu - box.x, box.cx),
      y: fraction(mapped.yEmu - box.y, box.cy),
      w: fraction(mapped.cxEmu, box.cx),
      h: fraction(mapped.cyEmu, box.cy),
      content,
    })),
  };
}

/** `value` as a fraction of `size` (0 for an empty size). */
function fraction(value: number, size: number): number {
  return size > 0 ? value / size : 0;
}

/** A drawing element's content (see the module header). */
function readContent(content: Element, ctx: ReadContext): SheetDrawingContent {
  switch (content.localName) {
    case "pic":
      return readPicture(content, ctx);
    case "sp":
    case "cxnSp":
      return readShape(content);
    case "grpSp":
      return readGroup(content, ctx);
    case "graphicFrame":
      return readGraphicFrame(content, ctx);
    default:
      return { kind: "foreign", label: "Ink" };
  }
}

// --- reading: MCE ------------------------------------------------------------------------------

/** One branch of an `mc:AlternateContent`: its anchor (at the wsDr level)
 * and its drawing element. */
interface Branch {
  element: Element;
  anchor?: Element;
  content?: Element;
}

/** The wrapper's branches, Choices then the Fallback. `outer` is the anchor
 * a wrapper inside an anchor sits in (its branches hold content only). */
function branchesOf(mc: Element, outer: Element | undefined): Branch[] {
  const choices = Array.from(mc.children).filter((child) => isMc(child, "Choice"));
  const fallbacks = Array.from(mc.children).filter((child) => isMc(child, "Fallback"));
  return [...choices, ...fallbacks.slice(0, 1)].map((element) => {
    const children = Array.from(element.children);
    if (outer) {
      return { element, anchor: outer, content: children.find(isContentElement) };
    }
    const anchor = children.find(isAnchorElement);
    const inner = anchor ? anchorContent(anchor) : undefined;
    return { element, anchor, content: inner && isContentElement(inner) ? inner : undefined };
  });
}

/** The label a wrapper's first Choice gives it (see the module header). */
function mceLabel(choice: Element | undefined): string {
  if (!choice) {
    return "Object";
  }
  const namespaces = (choice.getAttribute("Requires") ?? "")
    .split(/\s+/)
    .filter(Boolean)
    .map((prefix) => choice.lookupNamespaceURI(prefix) ?? "");
  if (namespaces.some((ns) => ns.endsWith("/chartex"))) {
    return "Chart";
  }
  if (namespaces.includes(SLE15_NS)) {
    return "Slicer";
  }
  if (namespaces.includes(AINK_NS)) {
    return "Ink";
  }
  if (namespaces.includes(A14_NS)) {
    if (choice.getElementsByTagNameNS(A14_NS, "compatExt").length > 0) {
      return "Form control";
    }
    if (choice.getElementsByTagNameNS(A14_NS, "m").length > 0) {
      return "Equation";
    }
  }
  return "Object";
}

/** The face a read branch's drawing element paints, if it has one. */
function faceOf(content: Element | undefined, ctx: ReadContext): PaintContent | undefined {
  if (!content || isOn(cNvPrOf(content)?.getAttribute("hidden"))) {
    return undefined;
  }
  if (isXdr(content, "pic")) {
    const picture = readPicture(content, ctx);
    return picture.kind === "image" ? picture : undefined;
  }
  if (isXdr(content, "sp")) {
    const shape = readShape(content);
    return shape.visual.unsupportedGeometryXml === undefined ? shape : undefined;
  }
  return undefined;
}

/** What reading one top-level child gives, less its id and sourceXml. */
type ReadDrawing = Omit<SheetDrawing, "id" | "sourceXml">;

/** An `mc:AlternateContent` drawing (see the module header). */
function readAlternateContent(mc: Element, ctx: ReadContext | undefined, outer?: Element): ReadDrawing {
  const branches = branchesOf(mc, outer);
  const read = selectAlternateContent(mc, UNDERSTOOD_SHEET_DRAWING_NAMESPACES);
  const readBranch = branches.find((branch) => read.length > 0 && Array.from(branch.element.children).includes(read[0]!));
  const label = mceLabel(branches.find((branch) => isMc(branch.element, "Choice"))?.element);
  const geometry = branches.find((branch) => branch.anchor !== undefined);
  const named = branches.find((branch) => (cNvPrOf(branch.content)?.getAttribute("name") ?? "") !== "") ?? geometry;
  const base = baseOf(cNvPrOf(named?.content));
  const drawing: ReadDrawing = { anchor: geometry?.anchor ? readAnchorElement(geometry.anchor) : unanchored(), name: base.name, content: { kind: "foreign", label } };
  if (base.altText !== undefined) {
    drawing.altText = base.altText;
  }
  if (isOn(cNvPrOf(readBranch?.content)?.getAttribute("hidden")) || !geometry?.anchor) {
    drawing.hidden = true;
  }
  const rotation = geometry?.content ? rotationOf(geometry.content) : undefined;
  if (rotation !== undefined) {
    drawing.rotationDeg = rotation;
  }
  const face = ctx && label !== "Chart" ? faceOf(readBranch?.content, ctx) : undefined;
  if (face) {
    drawing.content = { kind: "foreign", label, face };
  }
  return drawing;
}

/** One top-level child of `xdr:wsDr`. Without a context (the writer
 * re-reading a source) only the anchor and the base fields are read. */
function readTopLevel(el: Element, ctx: ReadContext | undefined): ReadDrawing {
  if (isMc(el, "AlternateContent")) {
    return readAlternateContent(el, ctx);
  }
  if (!isAnchorElement(el)) {
    return { anchor: unanchored(), name: "", hidden: true, content: { kind: "foreign", label: "Object" } };
  }
  const anchor = readAnchorElement(el);
  const content = anchorContent(el);
  if (content && isMc(content, "AlternateContent")) {
    return { ...readAlternateContent(content, ctx, el), anchor };
  }
  const base = baseOf(cNvPrOf(content));
  const drawing: ReadDrawing = { anchor, ...base, content: content && ctx ? readContent(content, ctx) : { kind: "foreign", label: "Object" } };
  const rotation = content ? rotationOf(content) : undefined;
  if (rotation !== undefined) {
    drawing.rotationDeg = rotation;
  }
  return drawing;
}

// --- reading: the part -------------------------------------------------------------------------

/** The drawings of one `xdr:wsDr` part's text (see the module header);
 * [] when it is not a well-formed drawing part. */
function readDrawingPartXml(xml: string, ctx: ReadContext, metrics: GridMetrics): SheetDrawing[] {
  const root = elementTree(xml)?.[0];
  if (!root || root.local !== "wsDr" || root.ns === undefined || !XDR_NAMESPACES.has(root.ns)) {
    return [];
  }
  const elements = Array.from(parseXml(xml).documentElement.children);
  let at = root.tagEnd;
  return root.children.map((node, index) => {
    const sourceXml = xml.slice(at, node.end);
    at = node.end;
    const read = readTopLevel(elements[index]!, ctx);
    return { id: `dr${index + 1}`, ...read, anchor: withDerivedSize(read.anchor, metrics), sourceXml };
  });
}

/** The root of a worksheet part and its children after `sheetData` (all of
 * them when the scan cannot start there): local name, offsets and
 * CT_Worksheet rank — read without scanning the cells. */
export interface WorksheetTail {
  prefix: string;
  rootTag: string;
  children: { local: string; main: boolean; rank: number; start: number; end: number; attrs: readonly XmlScanAttribute[] }[];
  rootClose: number;
}

/** CT_Worksheet's children from `sheetData` on, in schema order. */
const WORKSHEET_TAIL_ORDER = [
  "sheetData",
  "sheetCalcPr",
  "sheetProtection",
  "protectedRanges",
  "scenarios",
  "autoFilter",
  "sortState",
  "dataConsolidate",
  "customSheetViews",
  "mergeCells",
  "phoneticPr",
  "conditionalFormatting",
  "dataValidations",
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "drawingHF",
  "picture",
  "oleObjects",
  "controls",
  "webPublishItems",
  "tableParts",
  "extLst",
];

function scanWorksheetTail(xml: string, text: string, offset: number): WorksheetTail {
  const tail: WorksheetTail = { prefix: "", rootTag: "", children: [], rootClose: xml.length };
  let depth = 0;
  let open: WorksheetTail["children"][number] | null = null;
  scanXml(text, (event) => {
    if (event.type === "open") {
      depth += 1;
      if (depth === 1) {
        tail.prefix = event.name.includes(":") ? `${event.name.slice(0, event.name.indexOf(":"))}:` : "";
        tail.rootTag = text.slice(event.start, event.end);
      } else if (depth === 2) {
        const main = event.ns !== undefined && SPREADSHEETML_NAMESPACES.has(event.ns);
        const child = { local: event.local, main, rank: main ? WORKSHEET_TAIL_ORDER.indexOf(event.local) : -1, start: event.start + offset, end: event.end + offset, attrs: event.attrs };
        tail.children.push(child);
        open = event.selfClosing ? null : child;
      } else if (depth === 4 && open !== null && open.rank < 0 && event.ns !== undefined && SPREADSHEETML_NAMESPACES.has(event.ns)) {
        open.rank = WORKSHEET_TAIL_ORDER.indexOf(event.local); // what an mc:Choice wraps
      }
      if (event.selfClosing) {
        depth -= 1;
      }
    } else if (event.type === "close") {
      if (depth === 2 && open !== null) {
        open.end = event.end + offset;
        open = null;
      } else if (depth === 1) {
        tail.rootClose = event.start + offset;
      }
      depth -= 1;
    }
  });
  return tail;
}

/** A worksheet part's tail (see WorksheetTail): the root start tag plus the
 * text after `</sheetData>`, scanned as one small document. */
export function worksheetTail(xml: string): WorksheetTail {
  let rootTag = "";
  let prefix = "";
  scanXml(xml, (event) => {
    if (event.type === "open") {
      rootTag = xml.slice(event.start, event.end);
      prefix = event.name.includes(":") ? `${event.name.slice(0, event.name.indexOf(":"))}:` : "";
      return "stop";
    }
    return undefined;
  });
  const marker = `</${prefix}sheetData>`;
  const close = xml.lastIndexOf(marker);
  if (close >= 0 && !rootTag.endsWith("/>")) {
    const cut = close + marker.length;
    try {
      const tail = scanWorksheetTail(xml, rootTag + xml.slice(cut), cut - rootTag.length);
      return { ...tail, rootTag };
    } catch (error) {
      if (!(error instanceof XmlScanError)) {
        throw error;
      }
    }
  }
  return scanWorksheetTail(xml, xml, 0);
}

/** The relationship id of a worksheet part's `<drawing r:id>`. */
export function worksheetDrawingRelId(xml: string): string | undefined {
  const drawing = worksheetTail(xml).children.find((child) => child.main && child.local === "drawing");
  return drawing?.attrs.find((attr) => attr.local === "id" && attr.ns !== undefined && REL_NAMESPACES.has(attr.ns))?.value;
}

/** A sheet's drawing part (`<drawing r:id>` through its relationships) and
 * the relationship, when it has one that resolves to a part. `worksheetXml`,
 * when given, is read instead of decoding the sheet part (SPEC-137: the
 * scan's text with the cells cut out — worksheetTail reads only past them). */
export function sheetDrawingPart(parts: PartLookup, sheetPartName: string, worksheetXml?: string): { partName: string; rel: RelEntry } | undefined {
  const rels = relationshipsOf(parts, sheetPartName);
  if (!rels.some((rel) => relationshipKind(rel.type) === "drawing")) {
    return undefined; // the cheap check first: no drawing relationship, no scan
  }
  const xml = worksheetXml ?? partTextOf(parts, sheetPartName);
  let rId: string | undefined;
  try {
    rId = xml === undefined ? undefined : worksheetDrawingRelId(xml);
  } catch (error) {
    if (error instanceof XmlScanError) {
      return undefined;
    }
    throw error;
  }
  const rel = rels.find((entry) => entry.id === rId);
  const partName = rel ? targetPart(sheetPartName, rel) : undefined;
  return rel && partName !== undefined && parts.has(partName) ? { partName, rel } : undefined;
}

/** Reads one sheet's drawing layer (see the module header) from `parts`:
 * its drawings in z-order (anchors sized under the sheet's GridMetrics) and
 * its drawing part's name. Pictures add their images to `media`, one part
 * per package path, shared across every sheet read with the same map. A
 * part that is not well-formed reads as no drawings (it stays in `opaque`). */
export function readXlsxSheetDrawings(
  parts: PartLookup,
  sheetPartName: string,
  sheet: Sheet,
  media: Map<string, MediaPart>,
  worksheetXml?: string,
): { drawings: SheetDrawing[]; drawingPart?: string } {
  const found = sheetDrawingPart(parts, sheetPartName, worksheetXml);
  if (!found) {
    return { drawings: [] };
  }
  const xml = partTextOf(parts, found.partName)!;
  let types: ContentTypes | null | undefined;
  const ctx: ReadContext = {
    parts,
    drawingPart: found.partName,
    rels: new Map(relationshipsOf(parts, found.partName).map((rel) => [rel.id, rel] as const)),
    media,
    contentTypes: () => {
      if (types === undefined) {
        const text = partTextOf(parts, "[Content_Types].xml");
        types = text === undefined ? null : readContentTypes(text);
      }
      return types ?? undefined;
    },
  };
  try {
    return { drawings: readDrawingPartXml(xml, ctx, buildGridMetrics(sheet)), drawingPart: found.partName };
  } catch (error) {
    if (error instanceof XmlParseError || error instanceof XmlScanError) {
      return { drawings: [], drawingPart: found.partName };
    }
    throw error;
  }
}

/** SPEC-132 "Chart sheets": a chartsheet part's layer when the part its
 * `<drawing r:id>` names holds exactly one top-level child — an anchor (Excel
 * writes `xdr:absoluteAnchor`; any anchor kind is accepted) whose
 * graphicFrame reads as a chart (its part resolves) — read exactly as a
 * worksheet's drawing is; undefined otherwise (the sheet stays foreign). */
export function readXlsxChartSheet(parts: PartLookup, chartsheetPartName: string, sheet: Sheet): { drawings: SheetDrawing[]; drawingPart: string } | undefined {
  // A chart drawing embeds no media: a scratch map keeps a rejected layer's
  // pictures out of the workbook's.
  const layer = readXlsxSheetDrawings(parts, chartsheetPartName, sheet, new Map());
  const [only, ...others] = layer.drawings;
  if (!only || others.length > 0 || layer.drawingPart === undefined || only.content.kind !== "chart") {
    return undefined;
  }
  return { drawings: [only], drawingPart: layer.drawingPart };
}

/** The relationship id an anchor's `c:chart` element names. */
function chartRelIdIn(sourceXml: string): string | undefined {
  return /<(?:[\w.-]+:)?chart\b[^>]*?\s[\w.-]+:id\s*=\s*(["'])([^"']*)\1/.exec(sourceXml)?.[2];
}

/** SPEC-132: the part each chart drawing relates — through its sheet's
 * drawing part: the `.rels` entry its anchor's `c:chart@r:id` names (so a
 * duplicated sheet's copy, whose closure copy relates its own chart part,
 * resolves to that part whatever `sourcePart` says) — else the content's
 * `sourcePart`; undefined for any other content, or a part `parts` lacks.
 * Each drawing part's `.rels` is read once. */
export function chartPartResolver(parts: PartLookup): (sheet: Pick<Sheet, "drawingPart">, drawing: SheetDrawing) => string | undefined {
  const relsByPart = new Map<string, readonly RelEntry[]>();
  const relsOf = (drawingPart: string): readonly RelEntry[] => {
    let rels = relsByPart.get(drawingPart);
    if (rels === undefined) {
      try {
        rels = relationshipsOf(parts, drawingPart);
      } catch (error) {
        if (!(error instanceof XmlScanError)) {
          throw error;
        }
        rels = [];
      }
      relsByPart.set(drawingPart, rels);
    }
    return rels;
  };
  return (sheet, drawing) => {
    if (drawing.content.kind !== "chart") {
      return undefined;
    }
    const rId = sheet.drawingPart !== undefined && drawing.sourceXml !== undefined ? chartRelIdIn(drawing.sourceXml) : undefined;
    const rel = rId === undefined ? undefined : relsOf(sheet.drawingPart!).find((entry) => entry.id === rId && relationshipKind(entry.type) === "chart");
    const related = rel ? targetPart(sheet.drawingPart!, rel) : undefined;
    if (related !== undefined && parts.has(related)) {
      return related;
    }
    const named = drawing.content.sourcePart;
    return named !== undefined && parts.has(named) ? named : undefined;
  };
}

// --- writing: the canonical forms (xlsx/drawingGolden.ts) ----------------------------------------

/** A new drawing part's root start tag. */
const WS_DR_OPEN = `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}">`;
const PACKAGE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

type Box = { x: number; y: number; cx: number; cy: number };

/** An anchor's absolute box under `metrics`, whole EMU. */
export function anchorBoxEmu(anchor: DrawingAnchor, metrics: GridMetrics): Box {
  const rect = drawingRectPx(anchor, metrics);
  return {
    x: Math.round(rect.left * EMU_PER_PX),
    y: Math.round(rect.top * EMU_PER_PX),
    cx: Math.round(rect.width * EMU_PER_PX),
    cy: Math.round(rect.height * EMU_PER_PX),
  };
}

/** `<from>`/`<to>` with its four children, `name` the element as written
 * and `prefix` its children's prefix. */
function cornerXml(name: string, prefix: string, point: AnchorPoint): string {
  const child = (local: string, value: number): string => `<${prefix}${local}>${Math.round(value)}</${prefix}${local}>`;
  return `<${name}>${child("col", point.col)}${child("colOff", point.colOffEmu)}${child("row", point.row)}${child("rowOff", point.rowOffEmu)}</${name}>`;
}

/** A colour as a DrawingML colour element: upper-case `srgbClr`, or the
 * scheme reference core/themeDrawingMl.ts writes. */
function colorXml(color: ModelColor): string {
  if (typeof color !== "string") {
    return schemeClrXmlFor(color as DrawingMlColorRef);
  }
  const hex = color.replace(/^#/, "");
  return `<a:srgbClr val="${/^[0-9a-fA-F]{6}$/.test(hex) ? hex.toUpperCase() : "000000"}"/>`;
}

function fillXml(fill: ShapeFill | undefined): string {
  if (!fill) {
    return "";
  }
  if (fill.kind === "none") {
    return `<a:noFill/>`;
  }
  const color = shapeFillBaseColor(fill);
  return color === undefined ? "" : `<a:solidFill>${colorXml(color)}</a:solidFill>`;
}

function lineXml(outline: ShapeOutline | undefined): string {
  if (!outline) {
    return "";
  }
  const dash = PRESET_BY_DASH[outline.dash];
  return `<a:ln w="${Math.round(outline.widthPt * EMU_PER_PT)}"><a:solidFill>${colorXml(outline.color)}</a:solidFill>${dash ? `<a:prstDash val="${dash}"/>` : ""}</a:ln>`;
}

function xfrmXml(box: Box, rotationDeg: number | undefined): string {
  const rot = rotationDeg ? ` rot="${Math.round(rotationDeg * 60000)}"` : "";
  return `<a:xfrm${rot}><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/></a:xfrm>`;
}

function cNvPrXml(id: number, drawing: SheetDrawing): string {
  const descr = drawing.altText === undefined ? "" : ` descr="${escapeAttributeText(drawing.altText)}"`;
  return `<xdr:cNvPr id="${id}" name="${escapeAttributeText(drawing.name)}"${descr}${drawing.hidden ? ` hidden="1"` : ""}/>`;
}

const ALIGN_ATTR: Readonly<Record<string, string>> = { left: "l", center: "ctr", right: "r" };

/** A shape's `xdr:txBody`: empty as Excel writes a fresh shape's, or the
 * text with the first run's size, weight and colour on every run. */
function textBodyXml(text: DrawingText | undefined): string {
  const anchor = text?.anchor === "middle" ? "ctr" : text?.anchor === "bottom" ? "b" : "t";
  const head = `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="${anchor}"/><a:lstStyle/>`;
  if (!text || text.paragraphs.length === 0) {
    return `${head}<a:p><a:endParaRPr lang="en-US" sz="1100"/></a:p></xdr:txBody>`;
  }
  const attrs = ` lang="en-US"${text.sizePt === undefined ? "" : ` sz="${Math.round(text.sizePt * 100)}"`}${text.bold ? ` b="1"` : ""}`;
  const rPr = text.color === undefined ? `<a:rPr${attrs}/>` : `<a:rPr${attrs}><a:solidFill>${colorXml(text.color)}</a:solidFill></a:rPr>`;
  const paragraphs = text.paragraphs.map((paragraph) => {
    const pPr = paragraph.align === undefined ? "" : `<a:pPr algn="${ALIGN_ATTR[paragraph.align]}"/>`;
    if (paragraph.text === "") {
      return `<a:p>${pPr}<a:endParaRPr${attrs}/></a:p>`;
    }
    const runs = paragraph.text
      .split("\n")
      .map((segment) => (segment === "" ? "" : `<a:r>${rPr}<a:t>${escapeElementText(segment)}</a:t></a:r>`))
      .join(`<a:br/>`);
    return `<a:p>${pPr}${runs}</a:p>`;
  });
  return `${head}${paragraphs.join("")}</xdr:txBody>`;
}

/** One `a:srcRect` side, omitted when 0. */
function cropSide(name: string, value: number | undefined): string {
  return value ? ` ${name}="${Math.round(value * 100000)}"` : "";
}

function pictureXml(drawing: SheetDrawing, content: Extract<SheetDrawingContent, { kind: "image" }>, id: number, embedRId: string, box: Box): string {
  const crop = content.crop;
  const srcRect = crop && (crop.l || crop.t || crop.r || crop.b) ? `<a:srcRect${cropSide("l", crop.l)}${cropSide("t", crop.t)}${cropSide("r", crop.r)}${cropSide("b", crop.b)}/>` : "";
  return (
    `<xdr:pic><xdr:nvPicPr>${cNvPrXml(id, drawing)}<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
    `<xdr:blipFill><a:blip xmlns:r="${REL_NS}" r:embed="${escapeAttributeValue(embedRId)}"/>${srcRect}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
    `<xdr:spPr>${xfrmXml(box, drawing.rotationDeg)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`
  );
}

function shapeXml(drawing: SheetDrawing, content: Extract<SheetDrawingContent, { kind: "shape" }>, id: number, box: Box): string {
  const { visual } = content;
  const prst = GEOMETRY_TO_PRST[visual.geometry] ?? "rect";
  const adjustments = visual.adjustments ?? [];
  const avLst =
    adjustments.length === 0
      ? `<a:avLst/>`
      : `<a:avLst>${adjustments.map((value, i) => `<a:gd name="${adjustmentGuideName(i, adjustments.length)}" fmla="val ${adjustmentToGuideValue(value)}"/>`).join("")}</a:avLst>`;
  return (
    `<xdr:sp macro="" textlink=""><xdr:nvSpPr>${cNvPrXml(id, drawing)}<xdr:cNvSpPr/></xdr:nvSpPr>` +
    `<xdr:spPr>${xfrmXml(box, drawing.rotationDeg ?? visual.rotationDeg)}<a:prstGeom prst="${prst}">${avLst}</a:prstGeom>${fillXml(visual.fill)}${lineXml(visual.outline)}</xdr:spPr>` +
    `${textBodyXml(content.text)}</xdr:sp>`
  );
}

/** SPEC-132: a created chart's `xdr:graphicFrame` (chartXml.ts
 * chartGraphicFrameXml — the golden's scaffold) with the drawing's alt text
 * and hidden flag set on its cNvPr, as a picture's and a shape's carry them. */
function chartFrameXml(drawing: SheetDrawing, id: number, chartRId: string): string {
  const frame = chartGraphicFrameXml(id, drawing.name, chartRId);
  const cNvPr = /<(?:[\w.-]+:)?cNvPr\b[^>]*>/.exec(frame);
  if (!cNvPr || (drawing.altText === undefined && !drawing.hidden)) {
    return frame;
  }
  let tag = cNvPr[0];
  if (drawing.altText !== undefined) {
    tag = setAttribute(tag, "descr", drawing.altText, escapeAttributeText);
  }
  if (drawing.hidden) {
    tag = setAttribute(tag, "hidden", "1");
  }
  return frame.slice(0, cNvPr.index) + tag + frame.slice(cNvPr.index + cNvPr[0].length);
}

/** A created drawing's anchor element in the canonical form
 * (xlsx/drawingGolden.ts): its anchor kind as the model has it — twoCell
 * with its editAs (omitted for "twoCell") and drawings.ts writtenPoints,
 * oneCell from + ext, absolute pos + ext — around an `xdr:pic` (an image;
 * `embedRId` its relationship), an `xdr:sp` (a shape) or, SPEC-132, an
 * `xdr:graphicFrame` (a chart; `chartRId` its relationship). `declare` adds
 * namespace declarations to the anchor's start tag (a part whose root does
 * not bind `xdr`/`a`). Undefined for content without a canonical form. */
export function canonicalAnchorXml(drawing: SheetDrawing, metrics: GridMetrics, ids: { cNvPrId: number; embedRId?: string; chartRId?: string }, declare = ""): string | undefined {
  const { anchor, content } = drawing;
  const box = anchorBoxEmu(anchor, metrics);
  let inner: string;
  if (content.kind === "image" && ids.embedRId !== undefined) {
    inner = pictureXml(drawing, content, ids.cNvPrId, ids.embedRId, box);
  } else if (content.kind === "shape") {
    inner = shapeXml(drawing, content, ids.cNvPrId, box);
  } else if (content.kind === "chart" && ids.chartRId !== undefined) {
    inner = chartFrameXml(drawing, ids.cNvPrId, ids.chartRId);
  } else {
    return undefined;
  }
  const body = `${inner}<xdr:clientData/>`;
  if (anchor.kind === "oneCell") {
    const from = anchor.from ?? pointFromPx(box.x / EMU_PER_PX, box.y / EMU_PER_PX, metrics);
    const ext = anchor.extEmu ?? { cx: box.cx, cy: box.cy };
    return `<xdr:oneCellAnchor${declare}>${cornerXml("xdr:from", "xdr:", from)}<xdr:ext cx="${Math.round(ext.cx)}" cy="${Math.round(ext.cy)}"/>${body}</xdr:oneCellAnchor>`;
  }
  if (anchor.kind === "absolute") {
    const pos = anchor.posEmu ?? { x: box.x, y: box.y };
    const ext = anchor.extEmu ?? { cx: box.cx, cy: box.cy };
    return (
      `<xdr:absoluteAnchor${declare}><xdr:pos x="${Math.round(pos.x)}" y="${Math.round(pos.y)}"/><xdr:ext cx="${Math.round(ext.cx)}" cy="${Math.round(ext.cy)}"/>` +
      `${body}</xdr:absoluteAnchor>`
    );
  }
  const { from, to } = writtenPoints(anchor, metrics);
  const editAs = anchor.editAs === undefined || anchor.editAs === "twoCell" ? "" : ` editAs="${anchor.editAs}"`;
  return `<xdr:twoCellAnchor${editAs}${declare}>${cornerXml("xdr:from", "xdr:", from)}${cornerXml("xdr:to", "xdr:", to)}${body}</xdr:twoCellAnchor>`;
}

/** A new drawing part holding `anchors` (canonical anchor elements). */
export function newDrawingPartXml(anchors: readonly string[]): string {
  return `${XML_DECLARATION}${WS_DR_OPEN}${anchors.join("")}</xdr:wsDr>`;
}

/** A new `.rels` part holding `relationships` (`<Relationship>` elements). */
export function newRelsXml(relationships: readonly string[]): string {
  return `${XML_DECLARATION}<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NS}">${relationships.join("")}</Relationships>`;
}

// --- writing: patching an imported drawing -----------------------------------------------------

const isXdrNode = (node: XmlNode, local?: string): boolean => node.ns !== undefined && XDR_NAMESPACES.has(node.ns) && (local === undefined || node.local === local);
const isANode = (node: XmlNode, local?: string): boolean => node.ns !== undefined && A_NAMESPACES.has(node.ns) && (local === undefined || node.local === local);
const isMcNode = (node: XmlNode, local: string): boolean => node.ns === MC_NS && node.local === local;
const isAnchorNode = (node: XmlNode): boolean => isXdrNode(node) && ANCHOR_LOCALS.has(node.local);
const isContentNode = (node: XmlNode): boolean => isXdrNode(node) && CONTENT_LOCALS.has(node.local);

/** An `mc:AlternateContent`'s branches, Choices then the first Fallback. */
function branchNodes(mc: XmlNode): XmlNode[] {
  return [...mc.children.filter((child) => isMcNode(child, "Choice")), ...mc.children.filter((child) => isMcNode(child, "Fallback")).slice(0, 1)];
}

/** A drawing's anchor elements: itself, or each branch's of a wrapper. */
function anchorNodes(top: XmlNode): XmlNode[] {
  if (isAnchorNode(top)) {
    return [top];
  }
  return isMcNode(top, "AlternateContent") ? branchNodes(top).flatMap((branch) => branch.children.filter(isAnchorNode).slice(0, 1)) : [];
}

/** An anchor's drawing elements: its content, or each branch's of a
 * wrapper inside it. */
function contentNodes(anchor: XmlNode): XmlNode[] {
  const content = anchor.children.find((child) => isContentNode(child) || isMcNode(child, "AlternateContent"));
  if (!content) {
    return [];
  }
  return isContentNode(content) ? [content] : branchNodes(content).flatMap((branch) => branch.children.filter(isContentNode).slice(0, 1));
}

function childNode(node: XmlNode, local: string, test: (node: XmlNode, local: string) => boolean = isXdrNode): XmlNode | undefined {
  return node.children.find((child) => test(child, local));
}

function nodeAttr(node: XmlNode | undefined, local: string): string | undefined {
  return node?.attrs.find((attr) => attr.local === local && attr.ns === undefined)?.value;
}

function nodePoint(node: XmlNode | undefined): AnchorPoint | undefined {
  if (!node) {
    return undefined;
  }
  const part = (local: string): number => Math.max(0, coordinate(childNode(node, local)?.text));
  return { col: part("col"), colOffEmu: part("colOff"), row: part("row"), rowOffEmu: part("rowOff") };
}

function samePoint(a: AnchorPoint, b: AnchorPoint): boolean {
  return a.col === b.col && a.row === b.row && Math.round(a.colOffEmu) === b.colOffEmu && Math.round(a.rowOffEmu) === b.rowOffEmu;
}

/** What an anchor element says, as written. */
interface WireAnchor {
  kind: DrawingAnchor["kind"];
  editAs?: NonNullable<DrawingAnchor["editAs"]>;
  from?: AnchorPoint;
  to?: AnchorPoint;
  pos?: { x: number; y: number };
  ext?: { cx: number; cy: number };
}

function wireAnchorOf(node: XmlNode): WireAnchor {
  const kind = node.local === "oneCellAnchor" ? "oneCell" : node.local === "absoluteAnchor" ? "absolute" : "twoCell";
  const pos = childNode(node, "pos");
  const ext = childNode(node, "ext");
  const editAs = nodeAttr(node, "editAs") ?? "twoCell";
  const wire: WireAnchor = { kind, editAs: (EDIT_AS.has(editAs) ? editAs : "twoCell") as WireAnchor["editAs"] };
  const from = nodePoint(childNode(node, "from"));
  const to = nodePoint(childNode(node, "to"));
  if (from) {
    wire.from = from;
  }
  if (to) {
    wire.to = to;
  }
  if (pos) {
    wire.pos = { x: coordinate(nodeAttr(pos, "x")), y: coordinate(nodeAttr(pos, "y")) };
  }
  if (ext) {
    wire.ext = { cx: coordinate(nodeAttr(ext, "cx")), cy: coordinate(nodeAttr(ext, "cy")) };
  }
  return wire;
}

/** What the model asks an anchor element of the source's kind to say. */
function wantedWireAnchor(anchor: DrawingAnchor, source: WireAnchor, metrics: GridMetrics): WireAnchor {
  const box = anchorBoxEmu(anchor, metrics);
  const ext = { cx: Math.round(anchor.extEmu?.cx ?? box.cx), cy: Math.round(anchor.extEmu?.cy ?? box.cy) };
  const corner = (x: number, y: number): AnchorPoint => pointFromPx(x / EMU_PER_PX, y / EMU_PER_PX, metrics);
  if (source.kind === "oneCell") {
    return { kind: "oneCell", from: anchor.kind !== "absolute" && anchor.from ? anchor.from : corner(box.x, box.y), ext };
  }
  if (source.kind === "absolute") {
    return { kind: "absolute", pos: { x: Math.round(anchor.posEmu?.x ?? box.x), y: Math.round(anchor.posEmu?.y ?? box.y) }, ext };
  }
  if (anchor.kind === "twoCell") {
    return { kind: "twoCell", editAs: anchor.editAs ?? "twoCell", ...writtenPoints(anchor, metrics, { from: source.from, to: source.to }) };
  }
  // The model's kind differs from the source's: the box, at the behaviour.
  const behaviour = anchorBehaviour(anchor);
  return {
    kind: "twoCell",
    editAs: behaviour === "move" ? "oneCell" : behaviour === "fixed" ? "absolute" : "twoCell",
    from: corner(box.x, box.y),
    to: corner(box.x + box.cx, box.y + box.cy),
  };
}

/** The `xfrm` of a drawing element (`xdr:xfrm`, or in `spPr`/`grpSpPr`). */
function xfrmNode(content: XmlNode): XmlNode | undefined {
  const own = content.children.find((child) => child.local === "xfrm");
  if (own) {
    return own;
  }
  const props = childNode(content, content.local === "grpSp" ? "grpSpPr" : "spPr");
  return props ? childNode(props, "xfrm", isANode) : undefined;
}

/** A start tag with attributes set (null removes; a typed name or alt text
 * keeps its line breaks and tabs); the splice when it changed. */
function tagSplice(xml: string, node: XmlNode, values: readonly (readonly [string, string | null])[]): Splice | undefined {
  const tag = xml.slice(node.start, node.tagEnd);
  let next = tag;
  for (const [name, value] of values) {
    next = setAttribute(next, name, value, escapeAttributeText);
  }
  return next === tag ? undefined : { start: node.start, end: node.tagEnd, text: next };
}

/** The prefix (colon included) a node's first child is written with, else its own. */
function childPrefix(node: XmlNode): string {
  const name = node.children[0]?.name ?? node.name;
  return name.includes(":") ? name.slice(0, name.indexOf(":") + 1) : "";
}

/** The splices that put the model's anchor into every anchor element of a
 * drawing (each branch alike), compared with the first; whether a point,
 * position or size moved. */
function anchorSplices(xml: string, anchors: readonly XmlNode[], anchor: DrawingAnchor, metrics: GridMetrics, splices: Splice[]): boolean {
  const source = wireAnchorOf(anchors[0]!);
  const want = wantedWireAnchor(anchor, source, metrics);
  const corners = (["from", "to"] as const).filter((local) => want[local] !== undefined && source[local] !== undefined && !samePoint(want[local], source[local]));
  const posMoved = want.pos !== undefined && source.pos !== undefined && (want.pos.x !== source.pos.x || want.pos.y !== source.pos.y);
  const extMoved = want.ext !== undefined && source.ext !== undefined && (want.ext.cx !== source.ext.cx || want.ext.cy !== source.ext.cy);
  const editAs = source.kind === "twoCell" && want.editAs !== undefined && want.editAs !== source.editAs ? want.editAs : undefined;
  for (const node of anchors) {
    for (const local of corners) {
      const corner = childNode(node, local);
      if (corner) {
        splices.push({ start: corner.start, end: corner.end, text: cornerXml(corner.name, childPrefix(corner), want[local]!) });
      }
    }
    const pos = posMoved ? childNode(node, "pos") : undefined;
    const ext = extMoved ? childNode(node, "ext") : undefined;
    const edits = [
      pos ? tagSplice(xml, pos, [["x", String(want.pos!.x)], ["y", String(want.pos!.y)]]) : undefined,
      ext ? tagSplice(xml, ext, [["cx", String(want.ext!.cx)], ["cy", String(want.ext!.cy)]]) : undefined,
      editAs !== undefined && node.local === "twoCellAnchor" ? tagSplice(xml, node, [["editAs", editAs === "twoCell" ? null : editAs]]) : undefined,
    ];
    splices.push(...edits.filter((splice): splice is Splice => splice !== undefined));
  }
  return corners.length > 0 || posMoved || extMoved;
}

/** The splices giving every drawing element's own `a:off`/`a:ext` the
 * anchor's box — skipping one whose four values are all 0 (Excel's
 * graphicFrame convention); `a:chOff`, `a:chExt`, `rot` and flips stay. */
function transformSplices(xml: string, contents: readonly XmlNode[], box: Box, splices: Splice[]): void {
  for (const content of contents) {
    const xfrm = xfrmNode(content);
    const off = xfrm ? childNode(xfrm, "off", isANode) : undefined;
    const ext = xfrm ? childNode(xfrm, "ext", isANode) : undefined;
    const values = [nodeAttr(off, "x"), nodeAttr(off, "y"), nodeAttr(ext, "cx"), nodeAttr(ext, "cy")].map((value) => coordinate(value));
    if (!xfrm || values.every((value) => value === 0)) {
      continue;
    }
    const edits = [
      off ? tagSplice(xml, off, [["x", String(box.x)], ["y", String(box.y)]]) : undefined,
      ext ? tagSplice(xml, ext, [["cx", String(box.cx)], ["cy", String(box.cy)]]) : undefined,
    ];
    splices.push(...edits.filter((splice): splice is Splice => splice !== undefined));
  }
}

/** The base fields reading a drawing's source gives (undefined: it does not
 * parse inside `rootTag`). */
function sourceBase(sourceXml: string, rootTag: string): DrawingBase | undefined {
  const name = /^<([^\s/>]+)/.exec(rootTag)?.[1] ?? "wsDr";
  try {
    const element = parseXml(`${rootTag}${sourceXml}</${name}>`).documentElement.firstElementChild;
    if (!element) {
      return undefined;
    }
    const read = readTopLevel(element, undefined);
    const base: DrawingBase = { name: read.name };
    if (read.altText !== undefined) {
      base.altText = read.altText;
    }
    if (read.hidden) {
      base.hidden = true;
    }
    return base;
  } catch (error) {
    if (error instanceof XmlParseError) {
      return undefined;
    }
    throw error;
  }
}

/** The splices setting `cNvPr` name / descr / hidden where the model
 * differs from what reading the source gives, in every drawing element. */
function cNvPrSplices(drawing: SheetDrawing, xml: string, rootTag: string, contents: readonly XmlNode[], splices: Splice[]): void {
  const read = sourceBase(xml, rootTag);
  if (!read) {
    return;
  }
  const changes: [string, string | null][] = [];
  if (drawing.name !== read.name) {
    changes.push(["name", drawing.name]);
  }
  if (drawing.altText !== read.altText) {
    changes.push(["descr", drawing.altText ?? null]);
  }
  if ((drawing.hidden ?? false) !== (read.hidden ?? false)) {
    changes.push(["hidden", drawing.hidden ? "1" : null]);
  }
  for (const content of changes.length > 0 ? contents : []) {
    const nv = content.children.find((child) => /^nv\w*Pr$/.test(child.local));
    const cNvPr = nv?.children.find((child) => child.local === "cNvPr");
    const splice = cNvPr ? tagSplice(xml, cNvPr, changes) : undefined;
    if (splice) {
      splices.push(splice);
    }
  }
}

/** An imported drawing's anchor element as written (see the module
 * header): its `sourceXml` with only the anchor, the content element's own
 * transform and the `cNvPr` fields that differ from the source patched —
 * byte-identical when nothing does. `rootTag` is the drawing part's
 * `xdr:wsDr` start tag, whose declarations the source is read under. */
export function patchedDrawingXml(drawing: SheetDrawing, rootTag: string, metrics: GridMetrics): string {
  const xml = drawing.sourceXml ?? "";
  const top = elementTree(xml, [rootTag])?.[0];
  const anchors = top ? anchorNodes(top) : [];
  if (anchors.length === 0) {
    return xml;
  }
  const contents = anchors.flatMap(contentNodes);
  const splices: Splice[] = [];
  if (anchorSplices(xml, anchors, drawing.anchor, metrics, splices)) {
    transformSplices(xml, contents, anchorBoxEmu(drawing.anchor, metrics), splices);
  }
  cNvPrSplices(drawing, xml, rootTag, contents, splices);
  return applySplices(xml, splices);
}

// --- writing: the part ------------------------------------------------------------------------

/** Every relationship id an XML fragment's attributes name (any attribute
 * in the r: namespace — embed, link, id, pict…). */
function relIdsIn(xml: string, context: readonly string[]): Set<string> {
  const ids = new Set<string>();
  const visit = (nodes: readonly XmlNode[]): void => {
    for (const node of nodes) {
      for (const attr of node.attrs) {
        if (attr.ns !== undefined && REL_NAMESPACES.has(attr.ns) && attr.value !== "") {
          ids.add(attr.value);
        }
      }
      visit(node.children);
    }
  };
  visit(elementTree(xml, context) ?? []);
  return ids;
}

const CNVPR_ID = /<(?:[\w.-]+:)?cNvPr\b[^>]*?\sid\s*=\s*["'](\d+)["']/g;

/** A drawing part's source and where its children sit. */
interface PartLayout {
  /** Everything up to and including the root start tag. */
  head: string;
  /** Everything from the end of the last source child on. */
  tail: string;
  rootTag: string;
  /** The relationship ids the source's drawings name. */
  referenced: Set<string>;
}

function partLayout(xml: string): PartLayout | undefined {
  const root = elementTree(xml)?.[0];
  if (!root || root.local !== "wsDr" || root.ns === undefined || !XDR_NAMESPACES.has(root.ns)) {
    return undefined;
  }
  if (root.selfClosing) {
    const open = xml.slice(root.start, root.tagEnd).replace(/\s*\/>$/, ">");
    return { head: xml.slice(0, root.start) + open, tail: `</${root.name}>${xml.slice(root.end)}`, rootTag: open, referenced: new Set() };
  }
  const rootTag = xml.slice(root.start, root.tagEnd);
  const last = root.children.at(-1);
  const bodyEnd = last ? last.end : root.tagEnd;
  return { head: xml.slice(0, root.tagEnd), tail: xml.slice(bodyEnd), rootTag, referenced: relIdsIn(xml.slice(root.tagEnd, bodyEnd), [rootTag]) };
}

/** The declarations a canonical anchor needs inside a part whose root does
 * not bind `xdr` and `a` to the transitional namespaces. */
function missingDeclarations(rootTag: string): string {
  const binds = (prefix: string, ns: string): boolean => new RegExp(`\\sxmlns:${prefix}\\s*=\\s*["']${ns.replace(/[.]/g, "\\.")}["']`).test(rootTag);
  return `${binds("xdr", XDR_NS) ? "" : ` xmlns:xdr="${XDR_NS}"`}${binds("a", A_NS) ? "" : ` xmlns:a="${A_NS}"`}`;
}

/** The k a created drawing's name ends with ("Picture 3" → 3). */
function nameNumber(drawing: SheetDrawing): number | undefined {
  const match = /\s(\d+)$/.exec(drawing.name);
  return match ? Number(match[1]) : undefined;
}

/** What writing a drawing part gives. */
export interface DrawingPartWrite {
  xml: string;
  /** The part's `.rels`; undefined when it needs none and had none. */
  rels?: string;
  /** The parts its relationships target after the write. */
  targets: string[];
  /** How many anchor elements it holds. */
  written: number;
}

/** A drawing part rebuilt (see the module header) — the retained `source`
 * the splice base, or a new part without one — holding `drawings` in
 * z-order: an imported drawing patched (canonical when there is no source
 * to patch it into), a created image, shape or chart canonical, any other
 * created drawing left out. A created drawing's cNvPr id is k + 1 (its
 * name's number) when free, else 1 + the highest in the part. `mediaPath`
 * names the package path a created picture's media is written at
 * (undefined: the picture is left out); `chartPart` (SPEC-132) the chart
 * part a created chart relates (undefined, or no `chartPart`: the chart is
 * left out). */
export function writeDrawingPart(options: {
  partName: string;
  source?: { xml: string; rels?: string };
  drawings: readonly SheetDrawing[];
  metrics: GridMetrics;
  mediaPath: (mediaId: string) => string | undefined;
  chartPart?: (drawing: SheetDrawing) => string | undefined;
}): DrawingPartWrite {
  const { partName, drawings, metrics, mediaPath, chartPart } = options;
  const layout = options.source ? partLayout(options.source.xml) : undefined;
  const declare = layout ? missingDeclarations(layout.rootTag) : "";
  const imported = layout ? drawings.filter((drawing) => drawing.sourceXml !== undefined) : [];
  const remaining = new Set(imported.flatMap((drawing) => [...relIdsIn(drawing.sourceXml!, [layout!.rootTag])]));
  const sourceRelsXml = layout ? options.source?.rels : undefined;
  const sourceRels = sourceRelsXml === undefined ? [] : readRels(sourceRelsXml);
  const dropped = new Set([...(layout?.referenced ?? [])].filter((id) => !remaining.has(id)));
  const kept = sourceRels.filter((rel) => !dropped.has(rel.id));
  const relIds = new Set(sourceRels.map((rel) => rel.id));
  const used = new Set(imported.flatMap((drawing) => [...drawing.sourceXml!.matchAll(CNVPR_ID)].map((match) => Number(match[1]))));
  const mint = (k: number | undefined): number => {
    const id = k !== undefined && k >= 1 && !used.has(k + 1) ? k + 1 : Math.max(1, ...used) + 1;
    used.add(id);
    return id;
  };
  const added: string[] = [];
  const relByTarget = new Map<string, string>();
  const relFor = (kind: "image" | "chart", type: string, path: string): string => {
    const key = `${kind} ${path}`;
    let id = relByTarget.get(key) ?? kept.find((rel) => relationshipKind(rel.type) === kind && targetPart(partName, rel) === path)?.id;
    if (id === undefined) {
      id = nextRelationshipId(relIds);
      relIds.add(id);
      added.push(relationshipXml(id, type, relativeTarget(partName, path)));
    }
    relByTarget.set(key, id);
    return id;
  };
  const anchors: string[] = [];
  for (const drawing of drawings) {
    if (layout && drawing.sourceXml !== undefined) {
      anchors.push(patchedDrawingXml(drawing, layout.rootTag, metrics));
      continue;
    }
    const { content } = drawing;
    const path = content.kind === "image" ? mediaPath(content.mediaId) : content.kind === "chart" ? chartPart?.(drawing) : undefined;
    if ((content.kind !== "shape" && path === undefined) || (content.kind !== "image" && content.kind !== "shape" && content.kind !== "chart")) {
      continue;
    }
    const ids = {
      cNvPrId: mint(nameNumber(drawing)),
      ...(path === undefined ? {} : content.kind === "chart" ? { chartRId: relFor("chart", CHART_RELATIONSHIP, path) } : { embedRId: relFor("image", IMAGE_RELATIONSHIP_TYPE, path) }),
    };
    const xml = canonicalAnchorXml(drawing, metrics, ids, declare);
    if (xml !== undefined) {
      anchors.push(xml);
    }
  }
  const xml = layout ? `${layout.head}${anchors.join("")}${layout.tail}` : newDrawingPartXml(anchors);
  const rels = sourceRelsXml !== undefined ? spliceRels(sourceRelsXml, dropped, added) : added.length > 0 ? newRelsXml(added) : undefined;
  const targets = rels === undefined ? [] : readRels(rels).flatMap((rel) => targetPart(partName, rel) ?? []);
  return { xml, ...(rels === undefined ? {} : { rels }), targets, written: anchors.length };
}

// --- writing: the sheet's side -----------------------------------------------------------------

/** The prefix (colon included) a root start tag binds to the relationships
 * namespace; undefined when it binds none. */
function relationshipPrefix(rootTag: string): string | undefined {
  for (const match of rootTag.matchAll(/\sxmlns:([\w.-]+)\s*=\s*["']([^"']*)["']/g)) {
    if (REL_NAMESPACES.has(match[2]!)) {
      return `${match[1]}:`;
    }
  }
  return undefined;
}

/** A worksheet part with its `<drawing r:id>` set to `rId` — replaced, or
 * inserted at its CT_Worksheet place (after `smartTags`, before
 * `legacyDrawing`) — or removed (`rId` null). Every other byte verbatim;
 * the cells are not scanned. */
export function spliceWorksheetDrawing(xml: string, rId: string | null): string {
  const tail = worksheetTail(xml);
  const existing = tail.children.find((child) => child.main && child.local === "drawing");
  if (rId === null) {
    return existing ? xml.slice(0, existing.start) + xml.slice(existing.end) : xml;
  }
  const prefix = relationshipPrefix(tail.rootTag);
  const element = `<${tail.prefix}drawing${prefix === undefined ? ` xmlns:r="${REL_NS}"` : ""} ${prefix ?? "r:"}id="${escapeAttributeValue(rId)}"/>`;
  if (existing) {
    return xml.slice(0, existing.start) + element + xml.slice(existing.end);
  }
  const rank = WORKSHEET_TAIL_ORDER.indexOf("drawing");
  const at = tail.children.find((child) => child.rank > rank)?.start ?? tail.rootClose;
  return xml.slice(0, at) + element + xml.slice(at);
}

/** What writing one sheet's drawing layer changes in its package. */
export interface SheetDrawingWrite {
  /** Part texts written: the drawing part, its `.rels`, the sheet's `.rels`. */
  texts: Map<string, string>;
  /** Relationship edges after the write, by owner part. */
  edges: Map<string, string[]>;
  /** The worksheet's `<drawing r:id>`: the id to set, or null to remove it. */
  worksheetDrawing?: string | null;
  /** A drawing part this write adds (its Override is the writer's). */
  addedPart?: string;
}

/** The relationship type family a `.rels` part uses (transitional unless
 * its officeDocument relationships are strict). */
function relationshipBase(rels: readonly RelEntry[]): string {
  const strict = rels.some((rel) => rel.type.startsWith("http://purl.oclc.org/ooxml/officeDocument/relationships/"));
  return strict ? "http://purl.oclc.org/ooxml/officeDocument/relationships" : REL_NS;
}

/** `xl/drawings/drawing{n}.xml`, n = 1 + the highest in use. */
function nextDrawingPartName(taken: ReadonlySet<string>): string {
  let max = 0;
  for (const name of taken) {
    const match = /^xl\/drawings\/drawing(\d+)\.xml$/.exec(name);
    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  }
  return `xl/drawings/drawing${max + 1}.xml`;
}

/** One `drawingsDirty` sheet's drawing layer written into the package
 * (SPEC-131 "Writing"): its part rebuilt in place; a first drawing on a
 * sheet without a part a new `xl/drawings/drawing{n}.xml` (next free in
 * `taken`, which it joins) with its `.rels`, the sheet relationship and the
 * worksheet's `<drawing>`; the last drawing gone the relationship and the
 * `<drawing>` removed (the part and its `.rels` then unreachable, so the
 * writer drops them). `retained` says whether the sheet part is in `parts`;
 * `chartPart` names the part a created chart relates (writeDrawingPart).
 * Undefined when nothing is written. */
export function writeSheetDrawings(options: {
  parts: PartLookup;
  sheet: Sheet;
  sheetPart: string;
  retained: boolean;
  taken: Set<string>;
  mediaPath: (mediaId: string) => string | undefined;
  chartPart?: (drawing: SheetDrawing) => string | undefined;
}): SheetDrawingWrite | undefined {
  const { parts, sheet, sheetPart, retained, taken, mediaPath, chartPart } = options;
  const found = retained ? sheetDrawingPart(parts, sheetPart) : undefined;
  const drawings = (sheet.drawings ?? []).filter(
    (drawing) => (found && drawing.sourceXml !== undefined) || drawing.content.kind === "image" || drawing.content.kind === "shape" || (drawing.content.kind === "chart" && chartPart !== undefined),
  );
  const sheetRelsPath = relsPathFor(sheetPart);
  const sheetRelsXml = retained ? partTextOf(parts, sheetRelsPath) : undefined;
  const sheetRels = sheetRelsXml === undefined ? [] : readRels(sheetRelsXml);
  const metrics = buildGridMetrics(sheet);
  const out: SheetDrawingWrite = { texts: new Map(), edges: new Map() };
  const sheetTargets = (rels: readonly RelEntry[]): string[] => rels.flatMap((rel) => targetPart(sheetPart, rel) ?? []);
  if (found) {
    const relsXml = partTextOf(parts, relsPathFor(found.partName));
    const source = { xml: partTextOf(parts, found.partName)!, ...(relsXml === undefined ? {} : { rels: relsXml }) };
    const written = writeDrawingPart({ partName: found.partName, source, drawings, metrics, mediaPath, ...(chartPart ? { chartPart } : {}) });
    if (written.written > 0) {
      out.texts.set(found.partName, written.xml);
      if (written.rels !== undefined) {
        out.texts.set(relsPathFor(found.partName), written.rels);
      }
      out.edges.set(found.partName, written.targets);
      return out;
    }
    // The last drawing went: the relationship and the <drawing> go too.
    out.texts.set(sheetRelsPath, spliceRels(sheetRelsXml!, new Set([found.rel.id]), []));
    out.edges.set(sheetPart, sheetTargets(sheetRels.filter((rel) => rel.id !== found.rel.id)));
    out.worksheetDrawing = null;
    return out;
  }
  const partName = nextDrawingPartName(taken);
  const written = writeDrawingPart({ partName, drawings, metrics, mediaPath, ...(chartPart ? { chartPart } : {}) });
  if (written.written === 0) {
    return undefined;
  }
  taken.add(partName);
  out.texts.set(partName, written.xml);
  if (written.rels !== undefined) {
    out.texts.set(relsPathFor(partName), written.rels);
  }
  out.edges.set(partName, written.targets);
  const rId = nextRelationshipId(sheetRels.map((rel) => rel.id));
  const rel = relationshipXml(rId, `${relationshipBase(sheetRels)}/drawing`, relativeTarget(sheetPart, partName));
  out.texts.set(sheetRelsPath, sheetRelsXml === undefined ? newRelsXml([rel]) : spliceRels(sheetRelsXml, new Set(), [rel]));
  out.edges.set(sheetPart, [...sheetTargets(sheetRels), partName]);
  out.worksheetDrawing = rId;
  out.addedPart = partName;
  return out;
}
