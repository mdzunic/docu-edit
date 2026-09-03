// Reads/writes the DrawingML shape markup SPEC-051 models: a `wps:wsp`
// (`prstGeom`/`solidFill`/`ln`/`txbx`) inside a `wp:inline`/`wp:anchor`, the
// same container `drawing.ts` uses for a picture. Mirrors drawing.ts's own
// split (parse.../...Xml, ParsedDrawing-shaped result) so the two modules
// read side by side; this one never touches Image and drawing.ts never
// touches Shape, per SPEC-024's "no shipped call site changes".
import type { Block, Paragraph, Shape } from "../documentModel";
import type {
  ArrowheadKind,
  ArrowheadSize,
  ShapeArrowhead,
  ShapeFill,
  ShapeOutline,
  ShapeShadow,
} from "../core/shapes";
import { LINE_GEOMETRIES } from "../core/shapes";
import { parseXml } from "../core/xml";
import { adjustmentGuideName, adjustmentToGuideValue } from "../core/shapeCatalogue";
import { GEOMETRY_TO_PRST, PRST_TO_GEOMETRY, parseAvLstAdjustments } from "../core/drawingMl";
import { arcToCubics, PATH_GRID, quadraticToCubic, type PathStep, type ShapePath, type ShapeSubpath } from "../core/shapePath";
import { childElements, escapeXmlAttribute, firstChildElement, serializeElement } from "../core/xml";
import { cssColorToOoxml, ooxmlColorToCss } from "../core/colors";
import { isThemeColor, resolveColor, type DocumentTheme, type ModelColor } from "../core/theme";
import { schemeClrXmlFor } from "../core/themeDrawingMl";
import { cmToEmu, emuToCm } from "../core/units";

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
// Our own extension namespace, used only to round-trip groupOrdinal (a
// purely editor-side concept — a flat groupOrdinal tag, not a real wpg:wgp
// group container, see "Specs this changes" / the grouping design note in
// the implement summary). Unknown-namespaced attributes are ignored by
// Word/LibreOffice, so this never corrupts a file another app opens.
const DE_NAMESPACE = "urn:docu-edit:shape";

const EMU_PER_POINT = 12700;
function ptToEmu(pt: number): number {
  return Math.round(pt * EMU_PER_POINT);
}
function emuToPt(emu: number): number {
  return emu / EMU_PER_POINT;
}
// OOXML rotation (a:xfrm/@rot) is 60000ths of a degree.
function degToOoxmlRot(deg: number): number {
  return Math.round(deg * 60000);
}
function ooxmlRotToDeg(rot: number): number {
  return Math.round(rot / 60000);
}

// The OOXML preset-name maps and the `a:avLst` adjustment parser now live in
// core/drawingMl.ts (SPEC-079's two-phase move) so the `.pptx` reader shares
// them; this module re-imports them, its behaviour unchanged.

/** Everything needed to parse a `wps:txbx`'s `w:txbxContent` into
 * `Shape.text.blocks` — injected by the caller (reader.ts) so this module
 * stays independent of the paragraph parser's own dependency graph
 * (styleCatalog/numberingCatalog/links live in reader.ts). */
export interface ShapeTextContext {
  parseParagraph: (pEl: Element) => Paragraph;
  /** Resolves an `a:blip/@r:embed` relationship id to the media part it
   * names (registering the part's bytes with the caller's media collection)
   * — SPEC-073 picture fills. Absent (older callers, unit stubs) preserves
   * any blipFill verbatim instead. */
  resolveBlip?: (embedId: string) => string | null;
}

function vAlignOf(bodyPr: Element | undefined): "top" | "middle" | "bottom" {
  const anchor = bodyPr?.getAttribute("anchor");
  if (anchor === "ctr") {
    return "middle";
  }
  if (anchor === "b") {
    return "bottom";
  }
  return "top";
}

function parseShapeText(wsp: Element, ctx: ShapeTextContext): Shape["text"] | undefined {
  const txbx = firstChildElement(wsp, WPS_NAMESPACE, "txbx");
  if (!txbx) {
    return undefined;
  }
  const content = firstChildElement(txbx, WORD_NAMESPACE, "txbxContent");
  if (!content) {
    return undefined;
  }
  const blocks: Block[] = [];
  for (const child of Array.from(content.children)) {
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "p") {
      blocks.push(ctx.parseParagraph(child));
    }
  }
  if (blocks.length === 0) {
    blocks.push({ type: "paragraph", runs: [] });
  }
  const bodyPr = firstChildElement(wsp, WPS_NAMESPACE, "bodyPr");
  return { blocks, vAlign: vAlignOf(bodyPr) };
}

/** An `a:srgbClr` (+ optional `a:alpha` child, D5's 0..1 opacity) parsed
 * off `parent`'s first solid colour, or null for a colour the model cannot
 * carry (a schemeClr, no srgbClr at all) — the caller then preserves the
 * whole property verbatim (AC-35). */
function parseSolidColor(parent: Element): { color: string; alpha?: number } | null {
  const srgbClr = firstChildElement(parent, A_NAMESPACE, "srgbClr");
  const color = srgbClr ? ooxmlColorToCss(srgbClr.getAttribute("val")) : undefined;
  if (!srgbClr || !color) {
    return null;
  }
  const alphaEl = firstChildElement(srgbClr, A_NAMESPACE, "alpha");
  const amt = alphaEl ? Number(alphaEl.getAttribute("val")) : NaN;
  const parsed: { color: string; alpha?: number } = { color };
  if (Number.isFinite(amt) && amt >= 0 && amt < 100000) {
    parsed.alpha = amt / 100000;
  }
  return parsed;
}

const OOXML_DASH_TO_MODEL: Record<string, ShapeOutline["dash"]> = {
  dash: "dash",
  sysDash: "dot",
  sysDot: "dot",
  dashDot: "dashDot",
  sysDashDot: "dashDot",
  lgDash: "longDash",
  lgDashDot: "longDashDot",
  lgDashDotDot: "longDashDot",
};

const OOXML_ARROW_TO_MODEL: Record<string, ArrowheadKind> = {
  none: "none",
  triangle: "triangle",
  arrow: "open",
  stealth: "stealth",
  diamond: "diamond",
  oval: "oval",
};

const MODEL_ARROW_TO_OOXML: Record<ArrowheadKind, string> = {
  none: "none",
  triangle: "triangle",
  open: "arrow",
  stealth: "stealth",
  diamond: "diamond",
  oval: "oval",
};

const OOXML_ARROW_SIZE: Record<ArrowheadSize, string> = { small: "sm", medium: "med", large: "lg" };

function parseArrowEnd(ln: Element, localName: "headEnd" | "tailEnd"): ShapeArrowhead | undefined {
  const el = firstChildElement(ln, A_NAMESPACE, localName);
  const kind = el ? OOXML_ARROW_TO_MODEL[el.getAttribute("type") ?? ""] : undefined;
  if (!el || !kind || kind === "none") {
    return undefined;
  }
  // Only a sized arrow end reads as an explicit outline arrowhead — a bare
  // `a:tailEnd type="triangle"` is the legacy `arrow`-geometry marker
  // (SPEC-070), handled by the geometry branch in shapeFromWsp.
  const len = el.getAttribute("len") ?? el.getAttribute("w");
  const size: ArrowheadSize = len === "sm" ? "small" : len === "lg" ? "large" : "medium";
  return { kind, size };
}

/** Whether an `a:headEnd`/`a:tailEnd` carries explicit size attributes —
 * the marker distinguishing a SPEC-073 outline arrowhead from SPEC-070's
 * bare `arrow`-geometry tail. */
function arrowEndIsExplicit(ln: Element | undefined, localName: "headEnd" | "tailEnd"): boolean {
  const el = ln ? firstChildElement(ln, A_NAMESPACE, localName) : undefined;
  return !!el && (el.hasAttribute("w") || el.hasAttribute("len"));
}

/** Parses an `a:ln` into the widened outline model, or flags it preserved
 * when its stroke uses vocabulary the model cannot hold (a gradient stroke,
 * a scheme colour). `noFill` and an absent `ln` are simply "no outline". */
function parseLn(ln: Element | undefined): { outline?: ShapeOutline; preserved: boolean } {
  if (!ln) {
    return { preserved: false };
  }
  if (firstChildElement(ln, A_NAMESPACE, "noFill")) {
    return { preserved: false };
  }
  const solidFill = firstChildElement(ln, A_NAMESPACE, "solidFill");
  if (!solidFill && (firstChildElement(ln, A_NAMESPACE, "gradFill") || firstChildElement(ln, A_NAMESPACE, "pattFill"))) {
    return { preserved: true };
  }
  const parsed = solidFill ? parseSolidColor(solidFill) : null;
  if (solidFill && !parsed) {
    return { preserved: true };
  }
  if (!parsed) {
    return { preserved: false };
  }
  const wAttr = ln.getAttribute("w");
  const outline: ShapeOutline = {
    color: parsed.color,
    widthPt: wAttr ? emuToPt(Number(wAttr)) : 0.75,
    dash: OOXML_DASH_TO_MODEL[firstChildElement(ln, A_NAMESPACE, "prstDash")?.getAttribute("val") ?? ""] ?? "solid",
  };
  if (parsed.alpha !== undefined) {
    outline.alpha = parsed.alpha;
  }
  const cap = ln.getAttribute("cap");
  if (cap === "rnd") {
    outline.cap = "round";
  } else if (cap === "sq") {
    outline.cap = "square";
  }
  const cmpd = ln.getAttribute("cmpd");
  if (cmpd === "dbl") {
    outline.compound = "double";
  } else if (cmpd === "thickThin") {
    outline.compound = "thickThin";
  } else if (cmpd === "thinThick") {
    outline.compound = "thinThick";
  }
  if (firstChildElement(ln, A_NAMESPACE, "round")) {
    outline.join = "round";
  } else if (firstChildElement(ln, A_NAMESPACE, "bevel")) {
    outline.join = "bevel";
  }
  const startArrow = parseArrowEnd(ln, "headEnd");
  if (startArrow && arrowEndIsExplicit(ln, "headEnd")) {
    outline.startArrow = startArrow;
  }
  const endArrow = parseArrowEnd(ln, "tailEnd");
  if (endArrow && arrowEndIsExplicit(ln, "tailEnd")) {
    outline.endArrow = endArrow;
  }
  return { outline, preserved: false };
}

/** The `a:avLst` a shape's adjustments serialise to (SPEC-069 AC-6a) —
 * empty when the shape carries none, so a plain shape is byte-unchanged. */
function avLstXml(shape: Shape): string {
  const adjustments = shape.adjustments;
  if (!adjustments || adjustments.length === 0) {
    return "<a:avLst/>";
  }
  const gds = adjustments
    .map((fraction, i) => `<a:gd name="${adjustmentGuideName(i, adjustments.length)}" fmla="val ${adjustmentToGuideValue(fraction)}"/>`)
    .join("");
  return `<a:avLst>${gds}</a:avLst>`;
}

export const FILL_ELEMENT_NAMES = new Set(["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"]);
const EFFECT_CHILD_NAMES = new Set([
  "blur",
  "fillOverlay",
  "glow",
  "innerShdw",
  "outerShdw",
  "prstShdw",
  "reflection",
  "softEdge",
]);

/** Everything the styling parser pulled off one `wps:spPr` (SPEC-073): the
 * modelled properties plus the verbatim leftovers for `unsupportedStyleXml`
 * — only *unparsed* property elements land there (AC-36). */
interface ParsedShapeStyle {
  fill?: ShapeFill;
  outline?: ShapeOutline;
  shadow?: ShapeShadow;
  preserved: string[];
}

function parseGradFill(el: Element): ShapeFill | null {
  const gsLst = firstChildElement(el, A_NAMESPACE, "gsLst");
  if (!gsLst) {
    return null;
  }
  const stops: { position: number; color: string; alpha?: number }[] = [];
  for (const gs of childElements(gsLst, A_NAMESPACE, "gs")) {
    const pos = Number(gs.getAttribute("pos"));
    const parsed = parseSolidColor(gs);
    if (!Number.isFinite(pos) || !parsed) {
      return null;
    }
    const stop: { position: number; color: string; alpha?: number } = {
      position: Math.min(1, Math.max(0, pos / 100000)),
      color: parsed.color,
    };
    if (parsed.alpha !== undefined) {
      stop.alpha = parsed.alpha;
    }
    stops.push(stop);
  }
  if (stops.length < 2) {
    return null;
  }
  stops.sort((a, b) => a.position - b.position);
  const pathEl = firstChildElement(el, A_NAMESPACE, "path");
  if (pathEl) {
    return { kind: "gradient", gradient: { type: "radial", stops } };
  }
  const lin = firstChildElement(el, A_NAMESPACE, "lin");
  const ang = lin ? Number(lin.getAttribute("ang")) : 0;
  return {
    kind: "gradient",
    gradient: { type: "linear", angleDeg: Number.isFinite(ang) ? Math.round(ang / 60000) : 0, stops },
  };
}

const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function parseBlipFill(el: Element, resolveBlip: ((embedId: string) => string | null) | undefined): ShapeFill | null {
  const blip = firstChildElement(el, A_NAMESPACE, "blip");
  const embedId = blip?.getAttributeNS(R_NAMESPACE, "embed");
  const mediaId = embedId && resolveBlip ? resolveBlip(embedId) : null;
  if (!blip || !mediaId) {
    return null;
  }
  const fill: ShapeFill = {
    kind: "picture",
    mediaId,
    mode: firstChildElement(el, A_NAMESPACE, "tile") ? "tile" : "stretch",
  };
  const alphaMod = firstChildElement(blip, A_NAMESPACE, "alphaModFix");
  const amt = alphaMod ? Number(alphaMod.getAttribute("amt")) : NaN;
  if (Number.isFinite(amt) && amt >= 0 && amt < 100000) {
    fill.alpha = amt / 100000;
  }
  return fill;
}

function parsePattFill(el: Element): ShapeFill | null {
  const prst = el.getAttribute("prst");
  const fgClr = firstChildElement(el, A_NAMESPACE, "fgClr");
  const bgClr = firstChildElement(el, A_NAMESPACE, "bgClr");
  const fg = fgClr ? parseSolidColor(fgClr) : null;
  const bg = bgClr ? parseSolidColor(bgClr) : null;
  if (!prst || !fg || !bg) {
    return null;
  }
  return { kind: "pattern", preset: prst, foreground: fg.color, background: bg.color };
}

function parseOuterShdw(el: Element): ShapeShadow | null {
  const parsed = parseSolidColor(el);
  if (!parsed) {
    return null;
  }
  const blurRad = Number(el.getAttribute("blurRad") ?? 0);
  const dist = Number(el.getAttribute("dist") ?? 0);
  const dir = Number(el.getAttribute("dir") ?? 0);
  if (![blurRad, dist, dir].every(Number.isFinite)) {
    return null;
  }
  return {
    color: parsed.color,
    alpha: parsed.alpha ?? 1,
    blurPt: emuToPt(blurRad),
    distancePt: emuToPt(dist),
    angleDeg: Math.round(dir / 60000),
  };
}

/** Parses every styling property off a `spPr` (SPEC-073): fill, outline,
 * shadow — and collects the verbatim XML of every property element the
 * model could not hold: an unparseable fill/outline, non-shadow effects,
 * scene3d/sp3d/extLst, anything unrecognised. Exactly one source per
 * property (AC-36): a parsed property never lands in `preserved`. */
export function parseShapeStyle(
  spPr: Element,
  resolveBlip?: (embedId: string) => string | null,
): ParsedShapeStyle {
  const style: ParsedShapeStyle = { preserved: [] };
  for (const child of Array.from(spPr.children)) {
    if (
      child.namespaceURI !== A_NAMESPACE ||
      child.localName === "xfrm" ||
      child.localName === "prstGeom" ||
      child.localName === "custGeom"
    ) {
      continue;
    }
    if (FILL_ELEMENT_NAMES.has(child.localName)) {
      let fill: ShapeFill | null = null;
      if (child.localName === "noFill") {
        continue; // absence — the model's "no fill"
      }
      if (child.localName === "solidFill") {
        const parsed = parseSolidColor(child);
        if (parsed) {
          fill = { kind: "solid", color: parsed.color };
          if (parsed.alpha !== undefined) {
            fill.alpha = parsed.alpha;
          }
        }
      } else if (child.localName === "gradFill") {
        fill = parseGradFill(child);
      } else if (child.localName === "blipFill") {
        fill = parseBlipFill(child, resolveBlip);
      } else if (child.localName === "pattFill") {
        fill = parsePattFill(child);
      }
      if (fill) {
        style.fill = fill;
      } else {
        style.preserved.push(serializeElement(child));
      }
      continue;
    }
    if (child.localName === "ln") {
      const { outline, preserved } = parseLn(child);
      if (preserved) {
        style.preserved.push(serializeElement(child));
      } else if (outline) {
        style.outline = outline;
      }
      continue;
    }
    if (child.localName === "effectLst") {
      for (const effect of Array.from(child.children)) {
        if (effect.namespaceURI !== A_NAMESPACE) {
          continue;
        }
        const shadow = effect.localName === "outerShdw" && !style.shadow ? parseOuterShdw(effect) : null;
        if (shadow) {
          style.shadow = shadow;
        } else {
          style.preserved.push(serializeElement(effect));
        }
      }
      continue;
    }
    style.preserved.push(serializeElement(child));
  }
  return style;
}

/** One parsed `w:drawing` holding a shape (`wps:wsp`), the shape-flavoured
 * analogue of drawing.ts's ParsedDrawing. */
export interface ParsedShapeDrawing {
  shape: Shape;
}

// --- custom geometry (SPEC-071 `a:custGeom`) --------------------------------

/** A single `a:pt` as a fraction of the path's own `w`/`h` extent, or null
 * when the coordinate is a guide name rather than a number (a formula-driven
 * custGeom the model cannot carry — the caller then preserves it verbatim). */
function custPoint(el: Element | undefined, w: number, h: number): { x: number; y: number } | null {
  if (!el) {
    return null;
  }
  const x = Number(el.getAttribute("x"));
  const y = Number(el.getAttribute("y"));
  if (!Number.isFinite(x) || !Number.isFinite(y)) {
    return null;
  }
  return { x: x / w, y: y / h };
}

/** Parses `a:custGeom/a:pathLst` into a `ShapePath` in 0..1 space, converting
 * `a:arcTo`/`a:quadBezTo` to cubics (SPEC-071 "Primitive reduction"). Returns
 * null — so the caller keeps the SPEC-008 preservation path — when the
 * geometry references guides (formula-driven) or uses a command we do not
 * model. `converted` is true when any arc/quadratic was reduced, so the
 * caller can preserve the source XML for a lossless same-format save. */
function parseCustGeom(custGeom: Element): { path: ShapePath; converted: boolean } | null {
  const pathLst = firstChildElement(custGeom, A_NAMESPACE, "pathLst");
  if (!pathLst) {
    return null;
  }
  const subpaths: ShapeSubpath[] = [];
  let converted = false;
  for (const pathEl of childElements(pathLst, A_NAMESPACE, "path")) {
    const w = Number(pathEl.getAttribute("w")) || PATH_GRID;
    const h = Number(pathEl.getAttribute("h")) || PATH_GRID;
    const steps: PathStep[] = [];
    let cur = { x: 0, y: 0 };
    let closed = false;
    const cmds = Array.from(pathEl.children).filter((c) => c.namespaceURI === A_NAMESPACE);
    for (const cmd of cmds) {
      const pts = Array.from(cmd.children).filter((c) => c.namespaceURI === A_NAMESPACE && c.localName === "pt");
      if (cmd.localName === "moveTo") {
        const p = custPoint(pts[0], w, h);
        if (!p) {
          return null;
        }
        steps.push({ op: "move", x: p.x, y: p.y });
        cur = p;
      } else if (cmd.localName === "lnTo") {
        const p = custPoint(pts[0], w, h);
        if (!p) {
          return null;
        }
        steps.push({ op: "line", x: p.x, y: p.y });
        cur = p;
      } else if (cmd.localName === "cubicBezTo") {
        const c1 = custPoint(pts[0], w, h);
        const c2 = custPoint(pts[1], w, h);
        const end = custPoint(pts[2], w, h);
        if (!c1 || !c2 || !end) {
          return null;
        }
        steps.push({ op: "cubic", x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y, x: end.x, y: end.y });
        cur = end;
      } else if (cmd.localName === "quadBezTo") {
        const c = custPoint(pts[0], w, h);
        const end = custPoint(pts[1], w, h);
        if (!c || !end) {
          return null;
        }
        steps.push({ op: "cubic", ...quadraticToCubic(cur, c, end) });
        cur = end;
        converted = true;
      } else if (cmd.localName === "arcTo") {
        const wR = Number(cmd.getAttribute("wR"));
        const hR = Number(cmd.getAttribute("hR"));
        const stAng = Number(cmd.getAttribute("stAng"));
        const swAng = Number(cmd.getAttribute("swAng"));
        if (![wR, hR, stAng, swAng].every(Number.isFinite)) {
          return null;
        }
        const rx = wR / w;
        const ry = hR / h;
        const start = (stAng / 60000) * (Math.PI / 180);
        const sweep = (swAng / 60000) * (Math.PI / 180);
        const cx = cur.x - rx * Math.cos(start);
        const cy = cur.y - ry * Math.sin(start);
        for (const seg of arcToCubics(cx, cy, rx, ry, start, sweep)) {
          steps.push({ op: "cubic", ...seg });
          cur = { x: seg.x, y: seg.y };
        }
        converted = true;
      } else if (cmd.localName === "close") {
        steps.push({ op: "close" });
        closed = true;
      } else {
        return null;
      }
    }
    if (steps.length > 0) {
      subpaths.push({ steps, closed });
    }
  }
  if (subpaths.length === 0) {
    return null;
  }
  const path: ShapePath = { subpaths };
  // Per-vertex smoothness is our own concept (SPEC-071) with no OOXML home,
  // round-tripped through an ignored de: attribute so point-edit mode
  // restores the right handles on reopen rather than inferring them.
  const smoothAttr = custGeom.getAttributeNS(DE_NAMESPACE, "smooth");
  if (smoothAttr) {
    path.smooth = smoothAttr.split("").map((c) => c === "1");
  }
  return { path, converted };
}

/** Serializes a `ShapePath` as `a:custGeom` on the shared 21600 grid
 * (AC-26). `a:gdLst`/`a:ahLst`/`a:cxnLst` are deliberately not written. */
export function custGeomXml(path: ShapePath): string {
  const g = PATH_GRID;
  const smoothAttr =
    path.smooth && path.smooth.length > 0
      ? ` xmlns:de="${DE_NAMESPACE}" de:smooth="${path.smooth.map((b) => (b ? "1" : "0")).join("")}"`
      : "";
  const pt = (x: number, y: number): string => `<a:pt x="${Math.round(x * g)}" y="${Math.round(y * g)}"/>`;
  let steps = "";
  for (const subpath of path.subpaths) {
    for (const step of subpath.steps) {
      if (step.op === "move") {
        steps += `<a:moveTo>${pt(step.x, step.y)}</a:moveTo>`;
      } else if (step.op === "line") {
        steps += `<a:lnTo>${pt(step.x, step.y)}</a:lnTo>`;
      } else if (step.op === "cubic") {
        steps += `<a:cubicBezTo>${pt(step.x1, step.y1)}${pt(step.x2, step.y2)}${pt(step.x, step.y)}</a:cubicBezTo>`;
      } else {
        steps += "<a:close/>";
      }
    }
  }
  return `<a:custGeom${smoothAttr}><a:pathLst><a:path w="${g}" h="${g}">${steps}</a:path></a:pathLst></a:custGeom>`;
}

/** The size + wrap a `wps:wsp` cannot supply for itself: for a top-level
 * shape they come from its `wp:inline`/`wp:anchor` container, for a group
 * member from the group's `chOff`/`chExt` mapping applied to the member's own
 * `a:xfrm` (see group.ts). */
export interface WspBox {
  widthCm: number;
  heightCm: number;
  wrap: Shape["wrap"];
}

/** Builds a Shape from a bare `wps:wsp` and an externally-resolved box —
 * everything the shape markup itself carries (geometry, custom path, fill,
 * outline, rotation, text), and nothing that lives on a container. Returns
 * null when the `wsp` has no `wps:spPr` to read.
 *
 * Split out of parseShapeDrawing so a `wpg:wgp` group member — which has no
 * container of its own — reads through exactly the same code as a top-level
 * shape, rather than a second parser drifting from this one. The caller adds
 * the container-derived fields: alt text, offsets, `relativeTo*`, z-order and
 * `groupOrdinal`. */
export function shapeFromWsp(wsp: Element, box: WspBox, ctx: ShapeTextContext): Shape | null {
  const spPr = firstChildElement(wsp, WPS_NAMESPACE, "spPr");
  if (!spPr) {
    return null;
  }
  const prstGeom = firstChildElement(spPr, A_NAMESPACE, "prstGeom");
  let geometry = PRST_TO_GEOMETRY[prstGeom?.getAttribute("prst") ?? ""];
  // SPEC-070 "Untangling the arrow preset": a `prst="line"` carrying a bare
  // `a:tailEnd` is the model's `arrow` (a line with an end arrowhead). The
  // legacy `straightConnector1` → `arrow` mapping stays in PRST_TO_GEOMETRY
  // for documents saved before this change. A tailEnd carrying explicit
  // size attributes is instead a SPEC-073 outline arrowhead on a plain
  // line, left to parseShapeStyle (the two writers never emit both forms).
  if (geometry === "line") {
    const lnEl = firstChildElement(spPr, A_NAMESPACE, "ln");
    const tailEnd = lnEl ? firstChildElement(lnEl, A_NAMESPACE, "tailEnd") : undefined;
    if (tailEnd && tailEnd.getAttribute("type") !== "none" && !arrowEndIsExplicit(lnEl, "tailEnd")) {
      geometry = "arrow";
    }
  }

  const xfrm = firstChildElement(spPr, A_NAMESPACE, "xfrm");
  const rotAttr = xfrm?.getAttribute("rot");
  const rotationDeg = rotAttr ? ooxmlRotToDeg(Number(rotAttr)) : 0;

  const shape: Shape = {
    type: "shape",
    geometry: geometry ?? "rect",
    widthCm: box.widthCm,
    heightCm: box.heightCm,
    wrap: box.wrap,
  };
  const custGeom = !geometry ? firstChildElement(spPr, A_NAMESPACE, "custGeom") : undefined;
  const parsedPath = custGeom ? parseCustGeom(custGeom) : null;
  if (parsedPath) {
    // A modelled path renders its real outline (AC-4); geometry stays the
    // "rect" bounding-box stand-in and nothing is preserved opaquely.
    shape.path = parsedPath.path;
    if (parsedPath.converted && custGeom) {
      shape.pathSource = { format: "docx", xml: serializeElement(custGeom) };
    }
  } else if (!geometry) {
    shape.unsupportedGeometryXml = serializeElement(prstGeom ?? custGeom ?? spPr);
  } else if (prstGeom) {
    const adjustments = parseAvLstAdjustments(prstGeom);
    if (adjustments.length > 0) {
      shape.adjustments = adjustments;
    }
  }
  if (rotationDeg) {
    shape.rotationDeg = rotationDeg;
  }
  const style = parseShapeStyle(spPr, ctx.resolveBlip);
  if (style.fill) {
    shape.fill = style.fill;
  }
  if (style.outline) {
    shape.outline = style.outline;
  }
  if (style.shadow) {
    shape.shadow = style.shadow;
  }
  if (style.preserved.length > 0) {
    shape.unsupportedStyleXml = style.preserved.join("");
  }
  const text = parseShapeText(wsp, ctx);
  if (text) {
    shape.text = text;
  }
  return shape;
}

/** Parses a `w:drawing` element into a Shape, or null when it is not a
 * single, recognisable `wps:wsp` this spec models — the caller's existing
 * SPEC-008 opaque-inline fallback then preserves it verbatim. */
export function parseShapeDrawing(
  drawingEl: Element,
  ctx: ShapeTextContext,
): ParsedShapeDrawing | null {
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
  const wsp = graphicData ? firstChildElement(graphicData, WPS_NAMESPACE, "wsp") : undefined;
  if (!wsp) {
    return null;
  }
  const docPr = firstChildElement(container, WP_NAMESPACE, "docPr");
  const altText = docPr?.getAttribute("descr") || undefined;
  const groupOrdinalAttr = docPr?.getAttributeNS(DE_NAMESPACE, "group");
  const groupOrdinal = groupOrdinalAttr ? Number(groupOrdinalAttr) : undefined;

  let wrap: Shape["wrap"] = "inline";
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

  const shape = shapeFromWsp(wsp, { widthCm: emuToCm(cx), heightCm: emuToCm(cy), wrap }, ctx);
  if (!shape) {
    return null;
  }
  if (altText) {
    shape.altText = altText;
  }
  if (groupOrdinal !== undefined && Number.isFinite(groupOrdinal)) {
    shape.groupOrdinal = groupOrdinal;
  }
  if (offsetXCm !== undefined) {
    shape.offsetXCm = offsetXCm;
  }
  if (offsetYCm !== undefined) {
    shape.offsetYCm = offsetYCm;
  }
  if (wrap !== "inline") {
    shape.relativeToH = "column";
    shape.relativeToV = "paragraph";
  }
  if (zOrder !== undefined) {
    shape.zOrder = zOrder;
  }
  return { shape };
}

/** A solid `ModelColor` as its DrawingML colour element (SPEC-077 AC-31): a
 * theme reference writes `a:schemeClr` with its transforms in canonical
 * order; a literal writes `a:srgbClr`. A reference combined with a below-1
 * alpha writes its resolved literal plus `a:alpha` — the AC-29 reader treats
 * a `schemeClr` with an alpha child as unmodelled and would flatten it to a
 * literal anyway, so resolving on write keeps the round trip consistent. */
export function solidClrXml(color: ModelColor, theme: DocumentTheme, alpha?: number): string {
  if (isThemeColor(color)) {
    if (alpha === undefined || alpha >= 1) {
      return schemeClrXmlFor(color);
    }
    return srgbClrXml(resolveColor(color, theme), alpha);
  }
  return srgbClrXml(color, alpha);
}

/** An `a:srgbClr`, with D5's 0..1 opacity written as `a:alpha` in
 * thousandths of a percent when below fully opaque. */
export function srgbClrXml(color: string, alpha?: number): string {
  const val = cssColorToOoxml(color) ?? "000000";
  if (alpha !== undefined && alpha < 1) {
    return `<a:srgbClr val="${val}"><a:alpha val="${Math.round(Math.max(0, alpha) * 100000)}"/></a:srgbClr>`;
  }
  return `<a:srgbClr val="${val}"/>`;
}

const MODEL_DASH_TO_OOXML: Record<ShapeOutline["dash"], string> = {
  solid: "",
  dash: "dash",
  dot: "sysDot",
  dashDot: "dashDot",
  longDash: "lgDash",
  longDashDot: "lgDashDot",
};

/** The `a:prstDash` for a model dash value — shared with connector.ts so the
 * two writers cannot drift (a connector's longDash/dashDot must round-trip
 * exactly like a shape's, AC-2/AC-32). */
export function dashXml(dash: ShapeOutline["dash"]): string {
  const val = MODEL_DASH_TO_OOXML[dash];
  return val ? `<a:prstDash val="${val}"/>` : "";
}

function arrowEndXml(tag: "headEnd" | "tailEnd", arrow: ShapeArrowhead): string {
  const size = OOXML_ARROW_SIZE[arrow.size];
  return `<a:${tag} type="${MODEL_ARROW_TO_OOXML[arrow.kind]}" w="${size}" len="${size}"/>`;
}

/** The outline `a:ln` (SPEC-073 AC-2/AC-32): width, cap/compound attributes,
 * colour+alpha, dash, join element and sized arrow ends. `legacyTail`
 * (SPEC-070 "Untangling the arrow preset") adds the bare
 * `<a:tailEnd type="triangle"/>` marking the model's `arrow` geometry —
 * suppressed when an explicit end arrowhead already writes a sized tail. */
function outlineXml(outline: ShapeOutline | undefined, theme: DocumentTheme, legacyTail = false): string {
  const bareTail = legacyTail ? '<a:tailEnd type="triangle"/>' : "";
  if (!outline) {
    return `<a:ln>${bareTail || "<a:noFill/>"}</a:ln>`;
  }
  const joinXml =
    outline.join === "round" ? "<a:round/>" : outline.join === "bevel" ? "<a:bevel/>" : "";
  const capAttr =
    outline.cap === "round" ? ' cap="rnd"' : outline.cap === "square" ? ' cap="sq"' : "";
  const cmpd =
    outline.compound === "double"
      ? "dbl"
      : outline.compound === "thickThin"
        ? "thickThin"
        : outline.compound === "thinThick"
          ? "thinThick"
          : "";
  const cmpdAttr = cmpd ? ` cmpd="${cmpd}"` : "";
  const head = outline.startArrow ? arrowEndXml("headEnd", outline.startArrow) : "";
  const tail = outline.endArrow ? arrowEndXml("tailEnd", outline.endArrow) : bareTail;
  return (
    `<a:ln w="${ptToEmu(outline.widthPt)}"${capAttr}${cmpdAttr}>` +
    `<a:solidFill>${solidClrXml(outline.color, theme, outline.alpha)}</a:solidFill>` +
    dashXml(outline.dash) +
    joinXml +
    head +
    tail +
    "</a:ln>"
  );
}

/** A `ShapeFill` as its spPr fill element (SPEC-073 AC-32). `relIdForMedia`
 * resolves a picture fill's mediaId to the relationship the media plan
 * registered — absent/unresolved falls back to `a:noFill` (the part is not
 * in the package, so a dangling r:embed would corrupt the file). */
function fillXml(
  fill: ShapeFill | undefined,
  theme: DocumentTheme,
  relIdForMedia?: (mediaId: string) => string | undefined,
): string {
  if (!fill || fill.kind === "none") {
    return "<a:noFill/>";
  }
  if (fill.kind === "solid") {
    return `<a:solidFill>${solidClrXml(fill.color, theme, fill.alpha)}</a:solidFill>`;
  }
  if (fill.kind === "gradient") {
    const stops = fill.gradient.stops
      .map(
        (stop) =>
          `<a:gs pos="${Math.round(Math.min(1, Math.max(0, stop.position)) * 100000)}">${srgbClrXml(stop.color, stop.alpha)}</a:gs>`,
      )
      .join("");
    const shape =
      fill.gradient.type === "radial"
        ? '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>'
        : `<a:lin ang="${Math.round((fill.gradient.angleDeg ?? 0) * 60000)}" scaled="1"/>`;
    return `<a:gradFill><a:gsLst>${stops}</a:gsLst>${shape}</a:gradFill>`;
  }
  if (fill.kind === "picture") {
    const relId = relIdForMedia?.(fill.mediaId);
    if (!relId) {
      return "<a:noFill/>";
    }
    const alphaXml =
      fill.alpha !== undefined && fill.alpha < 1
        ? `<a:alphaModFix amt="${Math.round(Math.max(0, fill.alpha) * 100000)}"/>`
        : "";
    const blip = alphaXml
      ? `<a:blip r:embed="${relId}" xmlns:r="${R_NAMESPACE}">${alphaXml}</a:blip>`
      : `<a:blip r:embed="${relId}" xmlns:r="${R_NAMESPACE}"/>`;
    const mode =
      fill.mode === "tile"
        ? '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>'
        : "<a:stretch><a:fillRect/></a:stretch>";
    return `<a:blipFill>${blip}${mode}</a:blipFill>`;
  }
  return (
    `<a:pattFill prst="${escapeXmlAttribute(fill.preset)}">` +
    `<a:fgClr>${srgbClrXml(fill.foreground)}</a:fgClr>` +
    `<a:bgClr>${srgbClrXml(fill.background)}</a:bgClr>` +
    "</a:pattFill>"
  );
}

/** The `a:outerShdw` for a modelled shadow (SPEC-073 AC-32). */
export function outerShdwXml(shadow: ShapeShadow): string {
  return (
    `<a:outerShdw blurRad="${ptToEmu(shadow.blurPt)}" dist="${ptToEmu(shadow.distancePt)}"` +
    ` dir="${Math.round(shadow.angleDeg * 60000)}" rotWithShape="0">` +
    srgbClrXml(shadow.color, shadow.alpha) +
    "</a:outerShdw>"
  );
}

/** CT_EffectList is an ordered sequence — blur, fillOverlay, glow, innerShdw,
 * outerShdw, prstShdw, reflection, softEdge — so a preserved effect must land
 * on the correct side of the modelled `a:outerShdw` (AC-36 "schema order"). */
const EFFECTS_BEFORE_OUTER_SHADOW = new Set(["blur", "fillOverlay", "glow", "innerShdw"]);

/** An empty split — the shape carried no `unsupportedStyleXml`. */
export function emptyPreservedStyle(): PreservedStyleSplit {
  return { fill: [], ln: [], effectsBefore: [], effectsAfter: [], other: [] };
}

export interface PreservedStyleSplit {
  fill: string[];
  ln: string[];
  /** Effect children that schema-precede `a:outerShdw` (blur/fillOverlay/glow/innerShdw). */
  effectsBefore: string[];
  /** `a:outerShdw` itself and everything after it (prstShdw/reflection/softEdge). */
  effectsAfter: string[];
  other: string[];
}

/** `unsupportedStyleXml`'s sibling elements, split by where the writer must
 * re-emit each: a preserved fill or `a:ln` replaces the modelled emission
 * slot; effect children merge into the one `a:effectLst` around the modelled
 * shadow's slot, keeping CT_EffectList's sequence valid; everything else
 * appends after it in stored order (AC-36 "schema order" — scene3d/sp3d/
 * extLst follow effectLst in CT_ShapeProperties). An unparseable string
 * degrades to the trailing slot rather than being dropped. */
export function splitPreservedStyle(xml: string): PreservedStyleSplit {
  const out = emptyPreservedStyle();
  let doc: Document;
  try {
    doc = parseXml(`<de:wrap xmlns:de="urn:docu-edit:shape">${xml}</de:wrap>`);
  } catch {
    out.other.push(xml);
    return out;
  }
  for (const child of Array.from(doc.documentElement.children)) {
    const serialized = serializeElement(child);
    if (child.namespaceURI !== A_NAMESPACE) {
      out.other.push(serialized);
    } else if (FILL_ELEMENT_NAMES.has(child.localName)) {
      out.fill.push(serialized);
    } else if (child.localName === "ln") {
      out.ln.push(serialized);
    } else if (EFFECT_CHILD_NAMES.has(child.localName)) {
      (EFFECTS_BEFORE_OUTER_SHADOW.has(child.localName) ? out.effectsBefore : out.effectsAfter).push(serialized);
    } else {
      out.other.push(serialized);
    }
  }
  return out;
}

/** Removes the named property slots from a `.docx` `unsupportedStyleXml`
 * string (AC-36 "exactly one source per property"): when an edit gives the
 * MODEL a value for fill/outline/shadow, the preserved (unparsed) element
 * occupying that slot must go, or a later save would resurrect the old
 * styling over the user's visible edit. Only the slot's own elements are
 * touched — a preserved glow stays when the shadow slot clears. Returns
 * undefined when nothing survives; an unparseable string returns unchanged
 * (nothing to remove surgically, and dropping it would lose content). */
export function clearPreservedDocxStyle(
  xml: string,
  slots: { fill?: boolean; outline?: boolean; shadow?: boolean },
): string | undefined {
  let doc: Document;
  try {
    doc = parseXml(`<de:wrap xmlns:de="${DE_NAMESPACE}">${xml}</de:wrap>`);
  } catch {
    return xml;
  }
  const kept: string[] = [];
  for (const child of Array.from(doc.documentElement.children)) {
    if (child.namespaceURI === A_NAMESPACE) {
      if (slots.fill && FILL_ELEMENT_NAMES.has(child.localName)) {
        continue;
      }
      if (slots.outline && child.localName === "ln") {
        continue;
      }
      if (slots.shadow && child.localName === "outerShdw") {
        continue;
      }
    }
    kept.push(serializeElement(child));
  }
  return kept.length > 0 ? kept.join("") : undefined;
}

/** Serialises a Shape's text frame's blocks into `wps:txbx/w:txbxContent`
 * paragraphs — injected by the caller (writer.ts) so this module stays
 * independent of paragraphXml's own dependency graph. */
export interface ShapeTextWriteContext {
  paragraphXml: (paragraph: Paragraph) => string;
  /** The document theme solid colours resolve/serialise against (SPEC-077). */
  theme: DocumentTheme;
  /** The relationship id the media plan registered for a mediaId (SPEC-073
   * picture fills) — absent for callers whose part carries no media. */
  relIdForMedia?: (mediaId: string) => string | undefined;
}

function anchorAttrOf(vAlign: "top" | "middle" | "bottom" | undefined): string {
  if (vAlign === "middle") {
    return "ctr";
  }
  if (vAlign === "bottom") {
    return "b";
  }
  return "t";
}

function txbxXml(shape: Shape, ctx: ShapeTextWriteContext): string {
  if (!shape.text) {
    return "";
  }
  const paragraphs = shape.text.blocks
    .map((block) => (block.type === "paragraph" ? ctx.paragraphXml(block) : ""))
    .join("");
  return (
    `<wps:txbx><w:txbxContent xmlns:w="${WORD_NAMESPACE}">${paragraphs || "<w:p/>"}</w:txbxContent></wps:txbx>` +
    `<wps:bodyPr wrap="square" anchor="${anchorAttrOf(shape.text.vAlign)}"/>`
  );
}

/** Serialises a Shape + the assigned docPrId into a `w:drawing` element
 * (wrapped by the caller in `<w:r>`, in `<w:p>` — mirrors drawing.ts's
 * drawingXml). Writes explicit `wp:positionH`/`wp:positionV` offsets (unlike
 * Image, which always writes 0) since Shape actually models a position. */
export function shapeDrawingXml(shape: Shape, docPrId: number, ctx: ShapeTextWriteContext): string {
  const cx = cmToEmu(shape.widthCm);
  const cy = cmToEmu(Math.max(0, shape.heightCm));
  const id = docPrId;
  const name = `Shape ${id}`;
  const descrAttr = shape.altText ? ` descr="${escapeXmlAttribute(shape.altText)}"` : "";
  const groupAttr = shape.groupOrdinal !== undefined ? ` de:group="${shape.groupOrdinal}"` : "";
  const docPr =
    `<wp:docPr id="${id}" name="${escapeXmlAttribute(name)}"${descrAttr}` +
    (groupAttr ? ` xmlns:de="${DE_NAMESPACE}"${groupAttr}` : "") +
    "/>";

  // SPEC-070: the model's `arrow` is "a line with an end arrowhead", not a
  // connector preset — it writes as `prst="line"` plus an `a:tailEnd`, the
  // same representation `.odt` already uses. `straightConnector1` is no longer
  // written (its read-only compatibility mapping stays in PRST_TO_GEOMETRY so
  // documents this app already saved still open as arrows).
  const isArrow = shape.geometry === "arrow" && !shape.unsupportedGeometryXml && !shape.path;
  // A path shape writes as `a:custGeom` (AC-26). When its path was reduced
  // from arcs/quadratics on read (`pathSource`) a same-format save re-emits
  // the verbatim source unchanged, losing nothing; otherwise the model's
  // move/line/cubic/close is written on the 21600 grid.
  const prst = shape.unsupportedGeometryXml || shape.path ? undefined : isArrow ? "line" : GEOMETRY_TO_PRST[shape.geometry];
  const geomXml = shape.path
    ? shape.pathSource?.format === "docx"
      ? shape.pathSource.xml
      : custGeomXml(shape.path)
    : (shape.unsupportedGeometryXml ?? `<a:prstGeom prst="${prst}">${avLstXml(shape)}</a:prstGeom>`);
  const rotAttr = shape.rotationDeg ? ` rot="${degToOoxmlRot(shape.rotationDeg)}"` : "";
  const isLine = LINE_GEOMETRIES.has(shape.geometry) && !shape.unsupportedGeometryXml && !shape.path;

  // SPEC-073 AC-36: modelled properties write from the model; preserved
  // (unparsed) ones re-emit verbatim in their schema slots — a preserved
  // fill/ln fills the modelled emission's slot only while the model holds no
  // value for it (an edit sets the model AND clears the preserved slot, see
  // clearPreservedShapeStyle; the model-first guard here covers any stale
  // leftover so the on-screen styling always wins), effect children join the
  // one effectLst around the modelled shadow's sequence position, the rest
  // follow.
  const preserved = shape.unsupportedStyleXml
    ? splitPreservedStyle(shape.unsupportedStyleXml)
    : emptyPreservedStyle();
  const fillPart = isLine
    ? preserved.fill.join("")
    : shape.fill
      ? fillXml(shape.fill, ctx.theme, ctx.relIdForMedia)
      : preserved.fill.length > 0
        ? preserved.fill.join("")
        : fillXml(undefined, ctx.theme, ctx.relIdForMedia);
  const lnPart = shape.outline
    ? outlineXml(shape.outline, ctx.theme, isArrow && !shape.outline.endArrow)
    : preserved.ln.length > 0
      ? preserved.ln.join("")
      : outlineXml(undefined, ctx.theme, isArrow);
  const effectPart =
    shape.shadow || preserved.effectsBefore.length > 0 || preserved.effectsAfter.length > 0
      ? `<a:effectLst>${preserved.effectsBefore.join("")}${shape.shadow ? outerShdwXml(shape.shadow) : ""}${preserved.effectsAfter.join("")}</a:effectLst>`
      : "";

  const spPr =
    "<wps:spPr>" +
    `<a:xfrm${rotAttr}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    geomXml +
    fillPart +
    lnPart +
    effectPart +
    preserved.other.join("") +
    "</wps:spPr>";

  const wsp =
    `<wps:wsp xmlns:wps="${WPS_NAMESPACE}">` +
    `<wps:cNvSpPr/>` +
    spPr +
    txbxXml(shape, ctx) +
    "</wps:wsp>";

  const graphic =
    `<a:graphic xmlns:a="${A_NAMESPACE}">` +
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    wsp +
    "</a:graphicData>" +
    "</a:graphic>";

  if (shape.wrap === "inline") {
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

  const behindDoc = shape.wrap === "behind" ? "1" : "0";
  let wrapXml: string;
  if (shape.wrap === "square") {
    wrapXml = '<wp:wrapSquare wrapText="bothSides"/>';
  } else if (shape.wrap === "topAndBottom") {
    wrapXml = "<wp:wrapTopAndBottom/>";
  } else {
    wrapXml = "<wp:wrapNone/>";
  }
  const offsetXEmu = cmToEmu(shape.offsetXCm ?? 0);
  const offsetYEmu = cmToEmu(shape.offsetYCm ?? 0);
  const relativeHeight = 251658240 + (shape.zOrder ?? 0);
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
