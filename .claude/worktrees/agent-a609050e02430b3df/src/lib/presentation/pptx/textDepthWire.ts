// The pptx wire mapping for SPEC-097's slide text depth — the normative source
// for how the list, spacing, text-body and hyperlink fields of the slide model
// ride `a:pPr`, `a:bodyPr`, `a:rPr@spc` and `a:hlinkClick`, and for the format
// of the two preservation slots (`SlideParagraph.pptxPPr`,
// `TextFrame/ShapeFrame.pptxBodyPr`) that carry everything the model does not.
//
// Preservation slot format (both slots): a serialised element
// `<a:pPr{attrs}>{children}</a:pPr>` / `<a:bodyPr{attrs}>{children}</a:bodyPr>`
// (self-closing when there are no children) holding ONLY the source's
// unmodelled attributes — source order, ` name="value"` with the value
// attribute-escaped, namespace declarations dropped — and its unmodelled
// children, each the element's own serialisation with the redundant DrawingML
// `xmlns:a` stripped from its start tag (the slide root always declares it),
// sorted into ECMA-376 slot order. Built deterministically from the DOM, so
// read → write → read reproduces the same string and write → read → write
// the same bytes. Absent when nothing is unmodelled. The writer splits a slot
// back into its slots by local name (`parseSlot` — a string scanner, never a
// DOM re-serialise) and re-emits each preserved child at its schema slot; a
// preserved spacing element or bullet REPLACES the modelled one at that slot.
//
// Reading is tolerant of unit form and attribute order. A modelled element
// whose value equals the model default (`a:lnSpc` 100%, `a:spcBef` 0pt, an
// explicit `anchor="t"`, `wrap="square"`, `a:noAutofit`, …) is something the
// model cannot express as "explicitly set", so it stays in the slot verbatim
// — the writer never emits a default itself, so its own output never lands
// there and a model written then re-read compares equal.
//
// Framework-free; imports only presentation siblings and src/lib/core.
import { bulletStyleForChar, type ListMarkerStyle } from "../../core/lists";
import { nextGraphemeBoundary } from "../../core/grapheme";
import { cmToEmu, emuToCm } from "../../core/units";
import { sanitizeHref } from "../../core/urls";
import { escapeXmlAttribute, resolvePartPath, serializeElement, type Relationship } from "../../core/xml";
import { DRAWINGML_NAMESPACE } from "../../core/drawingMl";
import type { SlideFrame, SlideParagraph, SlideRun, TextAnchor, TextBodyProps } from "../model";
import { effectiveHangingCm, effectiveIndentCm } from "../slideLists";
import { isLinkableKind } from "../slideLinks";
import { clampLetterSpacing, paragraphMarkerStyle } from "../text";
import { DEFAULT_INSETS_CM, MAX_TEXT_COLUMNS, quantizeScalePct } from "../textBody";

export const HYPERLINK_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
export const SLIDE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
export const IMAGE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
/** The `a:hlinkClick@action` of an in-deck slide jump. */
export const SLIDE_JUMP_ACTION = "ppaction://hlinksldjump";

const A_NS = DRAWINGML_NAMESPACE;
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XMLNS_NS = "http://www.w3.org/2000/xmlns/";

const r3 = (v: number): number => {
  const rounded = Math.round(v * 1000) / 1000;
  return rounded === 0 ? 0 : rounded; // never -0 (toEqual tells them apart)
};

// --- the model ↔ wire vocabularies ------------------------------------------

/** `a:buAutoNum@type` per numbered marker style (the `.`-suffixed family). */
const AUTONUM_TYPE: Partial<Record<ListMarkerStyle, string>> = {
  decimal: "arabicPeriod",
  lowerAlpha: "alphaLcPeriod",
  upperAlpha: "alphaUcPeriod",
  lowerRoman: "romanLcPeriod",
  upperRoman: "romanUcPeriod",
};
const STYLE_FOR_AUTONUM = new Map<string, ListMarkerStyle>(
  Object.entries(AUTONUM_TYPE).map(([style, type]) => [type as string, style as ListMarkerStyle]),
);

/** `a:buChar@char` per bullet marker style. */
const BULLET_CHAR: Partial<Record<ListMarkerStyle, string>> = { disc: "•", circle: "○", square: "▪" };

const ANCHOR_WIRE: Record<TextAnchor, string> = { top: "t", middle: "ctr", bottom: "b" };
const ANCHOR_MODEL: Record<string, TextAnchor> = { t: "top", ctr: "middle", b: "bottom" };

/** The pptx default insets in EMU (left/top/right/bottom). */
const DEFAULT_INSETS_EMU = {
  left: cmToEmu(DEFAULT_INSETS_CM.left),
  top: cmToEmu(DEFAULT_INSETS_CM.top),
  right: cmToEmu(DEFAULT_INSETS_CM.right),
  bottom: cmToEmu(DEFAULT_INSETS_CM.bottom),
};

// --- preservation slots -----------------------------------------------------

/** The ECMA-376 `CT_TextParagraphProperties` child slots, in schema order. */
type PPrSlot = "lnSpc" | "spcBef" | "spcAft" | "buProps" | "bullet" | "tail";
const PPR_SLOT_RANK: Record<PPrSlot, number> = { lnSpc: 0, spcBef: 1, spcAft: 2, buProps: 3, bullet: 4, tail: 5 };
const BULLET_PROPERTY_ELEMENTS = new Set(["buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont"]);
const BULLET_ELEMENTS = new Set(["buNone", "buAutoNum", "buChar", "buBlip"]);

function pPrSlotOf(localName: string): PPrSlot {
  if (localName === "lnSpc" || localName === "spcBef" || localName === "spcAft") return localName;
  if (BULLET_PROPERTY_ELEMENTS.has(localName)) return "buProps";
  if (BULLET_ELEMENTS.has(localName)) return "bullet";
  return "tail"; // a:tabLst / a:defRPr / a:extLst and anything unrecognised
}

/** The ECMA-376 `CT_TextBodyProperties` child slots, in schema order. */
type BodySlot = "warp" | "autofit" | "tail";
const BODY_SLOT_RANK: Record<BodySlot, number> = { warp: 0, autofit: 1, tail: 2 };
const AUTOFIT_ELEMENTS = new Set(["noAutofit", "normAutofit", "spAutoFit"]);

function bodySlotOf(localName: string): BodySlot {
  if (localName === "prstTxWarp") return "warp";
  if (AUTOFIT_ELEMENTS.has(localName)) return "autofit";
  return "tail"; // a:scene3d / a:sp3d / a:flatTx / a:extLst and anything unrecognised
}

/** The attributes a `a:pPr`/`a:bodyPr` slot serialiser emits from the model —
 * a preserved attribute of the same name is never emitted beside one. */
const PPR_MODELLED_ATTRS = new Set(["marL", "indent", "lvl", "algn"]);

/** A preserved child: its local name and verbatim serialisation. */
interface SlotChild {
  local: string;
  xml: string;
}

/** A preservation slot split back into its attributes and children. */
export interface ParsedSlot {
  /** `[name, " name=\"value\""]` pairs, in slot order. */
  attrs: Array<[string, string]>;
  children: SlotChild[];
}

const START_TAG = /^<([^\s/>]+)((?:\s+[^\s=/>]+\s*=\s*"[^"]*")*)\s*(\/?)>/;
const ATTR = /\s+([^\s=/>]+)\s*=\s*"([^"]*)"/g;
const TOKEN = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<(\/?)([^\s/>!?]+)(?:\s+[^\s=/>]+\s*=\s*"[^"]*")*\s*(\/?)>/g;

function localOf(qname: string): string {
  return qname.slice(qname.indexOf(":") + 1);
}

/** The top-level elements of an XML fragment, verbatim, by local name. */
function splitTopLevel(inner: string): SlotChild[] {
  const out: SlotChild[] = [];
  let depth = 0;
  let start = 0;
  let local = "";
  for (const match of inner.matchAll(TOKEN)) {
    const [token, closing, qname, selfClosing] = match;
    if (qname === undefined) {
      continue; // comment / PI / CDATA
    }
    const at = match.index ?? 0;
    if (closing === "/") {
      depth -= 1;
      if (depth === 0) out.push({ local, xml: inner.slice(start, at + token.length) });
      continue;
    }
    if (depth === 0) {
      start = at;
      local = localOf(qname);
    }
    if (selfClosing === "/") {
      if (depth === 0) out.push({ local, xml: inner.slice(start, at + token.length) });
    } else {
      depth += 1;
    }
  }
  return out;
}

/** Splits a preservation slot back into its attributes and children (an
 * absent/empty slot yields nothing). */
export function parseSlot(slot: string | undefined): ParsedSlot {
  const parsed: ParsedSlot = { attrs: [], children: [] };
  if (!slot) {
    return parsed;
  }
  const head = START_TAG.exec(slot);
  if (!head) {
    return parsed;
  }
  for (const attr of head[2].matchAll(ATTR)) {
    parsed.attrs.push([attr[1], ` ${attr[1]}="${attr[2]}"`]);
  }
  if (head[3] !== "/") {
    const close = slot.lastIndexOf("</");
    parsed.children = splitTopLevel(slot.slice(head[0].length, close >= head[0].length ? close : slot.length));
  }
  return parsed;
}

/** An element's own serialisation with the redundant DrawingML declaration
 * dropped from its start tag (it rides inside a slide whose root declares
 * `xmlns:a`), so the bytes are the same however many times it round-trips. */
function preservedChildXml(el: Element): string {
  const xml = serializeElement(el);
  const head = START_TAG.exec(xml);
  if (!head) {
    return xml;
  }
  const tag = head[0].replace(` xmlns:a="${A_NS}"`, "");
  return tag + xml.slice(head[0].length);
}

/** One preserved attribute, as the slot stores it. */
function attrXml(attr: Attr): string {
  return ` ${attr.name}="${escapeXmlAttribute(attr.value)}"`;
}

function isNamespaceDeclaration(attr: Attr): boolean {
  return attr.namespaceURI === XMLNS_NS || attr.name === "xmlns" || attr.name.startsWith("xmlns:");
}

/** Builds a slot string, or undefined when nothing was kept. Children sort
 * stably into schema slot order. */
function buildSlot(tag: string, attrs: string[], children: Array<{ rank: number; xml: string }>): string | undefined {
  if (attrs.length === 0 && children.length === 0) {
    return undefined;
  }
  const sorted = children
    .map((child, i) => ({ ...child, i }))
    .toSorted((a, b) => a.rank - b.rank || a.i - b.i)
    .map((child) => child.xml)
    .join("");
  return sorted ? `<${tag}${attrs.join("")}>${sorted}</${tag}>` : `<${tag}${attrs.join("")}/>`;
}

// --- numeric attribute forms --------------------------------------------------

/** An integer attribute, or undefined when absent/unparseable. */
function intAttr(el: Element, name: string): number | undefined {
  const raw = el.getAttribute(name);
  if (raw === null || raw.trim() === "") return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

/** A DrawingML percentage in 1000ths of a percent — the transitional integer
 * form (`"62500"`) or the strict string form (`"62.5%"`). */
function percentAttr(el: Element, name: string): number | undefined {
  const raw = el.getAttribute(name);
  if (raw === null || raw.trim() === "") return undefined;
  const trimmed = raw.trim();
  const value = trimmed.endsWith("%") ? Number(trimmed.slice(0, -1)) * 1000 : Number(trimmed);
  return Number.isFinite(value) ? value : undefined;
}

/** The only element child of `el`, by local name, or undefined when it has
 * none or several. */
function onlyChild(el: Element): Element | undefined {
  return el.children.length === 1 && el.children[0].namespaceURI === A_NS ? el.children[0] : undefined;
}

function isSingleGrapheme(text: string): boolean {
  return text.length > 0 && nextGraphemeBoundary(text, 0) === text.length;
}

// --- a:pPr (read) -------------------------------------------------------------

/** The paragraph fields an `a:pPr` carries beyond `align`/`level` (which the
 * reader resolves itself). */
export type ParagraphWireFields = Pick<
  SlideParagraph,
  | "bullet"
  | "bulletStyle"
  | "bulletChar"
  | "bulletImageId"
  | "startAt"
  | "indentCm"
  | "hangingCm"
  | "lineSpacing"
  | "lineSpacingPt"
  | "lineSpacingRule"
  | "spaceBeforePt"
  | "spaceAfterPt"
  | "pptxPPr"
>;

/** Reads an `a:pPr` into the paragraph's list/spacing fields and its
 * preservation slot. `defaultBullet` is the role default when the source names
 * no bullet ("disc" for a body placeholder); `level` is the already-read
 * `lvl`; `imagePart` resolves a `a:buBlip` `r:embed` to a `model.media` part
 * name (undefined when it does not resolve). */
export function readParagraphPPr(
  pPr: Element | undefined,
  defaultBullet: SlideParagraph["bullet"],
  level: number,
  imagePart: (relId: string) => string | undefined,
): ParagraphWireFields {
  const out: ParagraphWireFields = {};
  const kept: Array<{ rank: number; xml: string }> = [];
  const keep = (el: Element, slot: PPrSlot): void => {
    kept.push({ rank: PPR_SLOT_RANK[slot], xml: preservedChildXml(el) });
  };
  let bulletEl: Element | undefined;
  if (pPr) {
    for (const child of Array.from(pPr.children)) {
      const local = child.namespaceURI === A_NS ? child.localName : "";
      const slot = pPrSlotOf(local);
      if (local === "lnSpc") {
        if (!readLineSpacing(child, out)) keep(child, slot);
      } else if (local === "spcBef" || local === "spcAft") {
        const pt = spacingPoints(child);
        if (pt === undefined) keep(child, slot);
        else if (local === "spcBef") out.spaceBeforePt = pt;
        else out.spaceAfterPt = pt;
      } else if (slot === "bullet" && local !== "") {
        if (!bulletEl) bulletEl = child;
        else continue; // a second bullet element is malformed; the first wins
      } else {
        keep(child, slot);
      }
    }
  }

  // The bullet: the first bullet element, else the role default.
  if (bulletEl) {
    if (!readBullet(bulletEl, out, imagePart)) keep(bulletEl, "bullet");
  } else if (defaultBullet) {
    out.bullet = defaultBullet;
  }

  // marL/indent: a list item or an indented level always writes both, so a
  // value equal to that default reads back as absent; a plain level-0
  // paragraph writes them only when explicit, so it keeps whatever it has.
  const probe: SlideParagraph = { runs: [], bullet: out.bullet, level };
  const writerAlwaysEmits = out.bullet !== undefined || level > 0;
  const marL = pPr ? intAttr(pPr, "marL") : undefined;
  if (marL !== undefined && (!writerAlwaysEmits || marL !== cmToEmu(effectiveIndentCm(probe)))) {
    out.indentCm = r3(emuToCm(marL));
  }
  const indent = pPr ? intAttr(pPr, "indent") : undefined;
  if (indent !== undefined) {
    const defaultHanging = cmToEmu(effectiveHangingCm(probe));
    if (!writerAlwaysEmits || -indent !== defaultHanging) {
      out.hangingCm = r3(emuToCm(-indent));
    }
  }

  const attrs = pPr
    ? Array.from(pPr.attributes)
        .filter((attr) => !isNamespaceDeclaration(attr) && !PPR_MODELLED_ATTRS.has(attr.localName))
        .map(attrXml)
    : [];
  const slot = buildSlot("a:pPr", attrs, kept);
  if (slot) out.pptxPPr = slot;
  return out;
}

/** `a:lnSpc` → `lineSpacing` (multiple) or `lineSpacingPt` + "exact"; false
 * when the form is unmodelled (an explicit 100%, an out-of-range value, …). */
function readLineSpacing(lnSpc: Element, out: ParagraphWireFields): boolean {
  const child = onlyChild(lnSpc);
  if (!child || lnSpc.attributes.length > 0) return false;
  if (child.localName === "spcPct") {
    const val = percentAttr(child, "val");
    if (val === undefined || val === 100000) return false;
    const multiple = val / 100000;
    if (multiple < 0.25 || multiple > 5) return false;
    out.lineSpacing = multiple;
    return true;
  }
  if (child.localName === "spcPts") {
    const val = intAttr(child, "val");
    if (val === undefined) return false;
    const pt = val / 100;
    if (pt < 4 || pt > 1000) return false;
    out.lineSpacingPt = pt;
    out.lineSpacingRule = "exact"; // pptx line spacing in points IS exact
    return true;
  }
  return false;
}

/** `a:spcBef`/`a:spcAft` → points, or undefined when unmodelled (a
 * percentage form, an explicit 0, an out-of-range value). */
function spacingPoints(el: Element): number | undefined {
  const child = onlyChild(el);
  if (!child || child.localName !== "spcPts" || el.attributes.length > 0) return undefined;
  const val = intAttr(child, "val");
  if (val === undefined || val <= 0 || val > 50000) return undefined;
  return val / 100;
}

/** Models a bullet element; false when it (also) stays preserved — an
 * unrecognised numbering scheme, a multi-grapheme glyph, a picture that does
 * not resolve or carries more than a plain `r:embed`. */
function readBullet(
  el: Element,
  out: ParagraphWireFields,
  imagePart: (relId: string) => string | undefined,
): boolean {
  switch (el.localName) {
    case "buNone":
      return true;
    case "buChar": {
      out.bullet = "disc";
      const char = el.getAttribute("char") ?? "";
      const style = bulletStyleForChar(char);
      if (style && style !== "disc") out.bulletStyle = style;
      // A recognised glyph the writer would not reproduce byte-for-byte ("·",
      // "◦", a Symbol/Wingdings private-use bullet) also keeps its exact char,
      // which wins over the style on write — the family still reads as its
      // style, the source glyph survives.
      if (style && BULLET_CHAR[style] === char) return true;
      if (isSingleGrapheme(char)) {
        out.bulletChar = char;
        return true;
      }
      return false;
    }
    case "buAutoNum": {
      out.bullet = "number";
      const startAt = intAttr(el, "startAt");
      if (startAt !== undefined && startAt !== 1 && startAt >= 1 && startAt <= 32767) {
        out.startAt = Math.round(startAt);
      }
      const style = STYLE_FOR_AUTONUM.get(el.getAttribute("type") ?? "");
      if (!style) return false; // an unrecognised scheme renders decimal, kept verbatim
      if (style !== "decimal") out.bulletStyle = style;
      return true;
    }
    case "buBlip": {
      out.bullet = "disc";
      const blip = Array.from(el.getElementsByTagNameNS(A_NS, "blip"))[0];
      const embed = blip?.getAttributeNS(R_NS, "embed");
      const part = embed ? imagePart(embed) : undefined;
      if (part) out.bulletImageId = part;
      const plain =
        !!blip &&
        onlyChild(el) === blip &&
        blip.children.length === 0 &&
        Array.from(blip.attributes).every((attr) => isNamespaceDeclaration(attr) || (attr.namespaceURI === R_NS && attr.localName === "embed"));
      return !!part && plain;
    }
    default:
      return false;
  }
}

// --- a:rPr (read) -------------------------------------------------------------

/** `a:rPr@spc` (100ths of a point) → letter spacing in points, clamped to the
 * model's −10…10 / 0.1 grid; undefined when absent. */
export function readLetterSpacing(rPr: Element | undefined): number | undefined {
  const spc = rPr ? intAttr(rPr, "spc") : undefined;
  return spc === undefined ? undefined : clampLetterSpacing(spc / 100);
}

/** What a link read needs from its slide. */
export interface LinkReadContext {
  slideRels: Map<string, Relationship>;
  slidePartPath: string;
  /** Slide part path → 1-based position in the deck's slide order. */
  slideNumberOf?: Map<string, number>;
}

/** A link read off an `a:hlinkClick`. */
export type LinkFields = Pick<SlideRun, "linkHref" | "linkSlide" | "linkTooltip">;

/** Reads an `a:hlinkClick`: an external hyperlink relationship →
 * `linkHref` (through `sanitizeHref` — a dangerous URL is dropped), a
 * `ppaction://hlinksldjump` with a slide relationship → `linkSlide`, and the
 * tooltip. Anything else (custom shows, macros, media actions) yields {}. */
export function readHlinkClick(el: Element | undefined, ctx: LinkReadContext): LinkFields {
  if (!el) {
    return {};
  }
  const relId = el.getAttributeNS(R_NS, "id");
  const rel = relId ? ctx.slideRels.get(relId) : undefined;
  if (!rel) {
    return {};
  }
  const out: LinkFields = {};
  const action = el.getAttribute("action") ?? "";
  if (rel.type === HYPERLINK_RELATIONSHIP && action === "") {
    const href = sanitizeHref(rel.target);
    if (!href) return {};
    out.linkHref = href;
  } else if (rel.type === SLIDE_RELATIONSHIP && action === SLIDE_JUMP_ACTION) {
    const slide = ctx.slideNumberOf?.get(resolvePartPath(ctx.slidePartPath, rel.target));
    if (slide === undefined) return {};
    out.linkSlide = slide;
  } else {
    return {};
  }
  const tooltip = el.getAttribute("tooltip");
  if (tooltip) out.linkTooltip = tooltip;
  return out;
}

// --- a:bodyPr (read) ----------------------------------------------------------

/** Reads a frame's own `a:bodyPr` into `TextBodyProps` plus its preservation
 * slot. `kind` decides the anchor default ("top" for text, "middle" for a
 * shape). */
export function readBodyPr(
  bodyPr: Element | undefined,
  kind: "text" | "shape",
): { body?: TextBodyProps; pptxBodyPr?: string } {
  if (!bodyPr) {
    return {};
  }
  const body: TextBodyProps = {};
  const keptAttrs = new Set<string>();

  // Insets: all four when any differs from the EMU defaults; explicit
  // defaults stay preserved verbatim.
  const insetNames = { left: "lIns", top: "tIns", right: "rIns", bottom: "bIns" } as const;
  const insets = {
    left: intAttr(bodyPr, insetNames.left),
    top: intAttr(bodyPr, insetNames.top),
    right: intAttr(bodyPr, insetNames.right),
    bottom: intAttr(bodyPr, insetNames.bottom),
  };
  const sides = ["left", "top", "right", "bottom"] as const;
  if (sides.some((side) => insets[side] !== undefined && insets[side] !== DEFAULT_INSETS_EMU[side])) {
    body.insetsCm = {
      left: r3(emuToCm(insets.left ?? DEFAULT_INSETS_EMU.left)),
      top: r3(emuToCm(insets.top ?? DEFAULT_INSETS_EMU.top)),
      right: r3(emuToCm(insets.right ?? DEFAULT_INSETS_EMU.right)),
      bottom: r3(emuToCm(insets.bottom ?? DEFAULT_INSETS_EMU.bottom)),
    };
  } else {
    for (const side of sides) if (insets[side] !== undefined) keptAttrs.add(insetNames[side]);
  }

  const anchor = ANCHOR_MODEL[bodyPr.getAttribute("anchor") ?? ""];
  const anchorDefault: TextAnchor = kind === "shape" ? "middle" : "top";
  if (anchor && anchor !== anchorDefault) body.anchor = anchor;
  else if (bodyPr.hasAttribute("anchor")) keptAttrs.add("anchor");

  const numCol = intAttr(bodyPr, "numCol");
  if (numCol !== undefined && numCol > 1) {
    const spcCol = intAttr(bodyPr, "spcCol") ?? 0;
    body.columns = { count: Math.min(MAX_TEXT_COLUMNS, Math.round(numCol)), gapCm: r3(emuToCm(Math.max(0, spcCol))) };
  } else {
    if (bodyPr.hasAttribute("numCol")) keptAttrs.add("numCol");
    if (bodyPr.hasAttribute("spcCol")) keptAttrs.add("spcCol");
  }

  const wrap = bodyPr.getAttribute("wrap");
  if (wrap === "none") body.wrap = false;
  else if (wrap !== null) keptAttrs.add("wrap");

  const modelled = new Set(["lIns", "tIns", "rIns", "bIns", "anchor", "numCol", "spcCol", "wrap"]);
  const attrs = Array.from(bodyPr.attributes)
    .filter((attr) => !isNamespaceDeclaration(attr) && (!modelled.has(attr.localName) || keptAttrs.has(attr.localName)))
    .map(attrXml);

  const kept: Array<{ rank: number; xml: string }> = [];
  let autofitSeen = false;
  for (const child of Array.from(bodyPr.children)) {
    const local = child.namespaceURI === A_NS ? child.localName : "";
    const slot = bodySlotOf(local);
    if (slot === "autofit" && !autofitSeen) {
      autofitSeen = true;
      if (local === "normAutofit") {
        body.autofit = "shrink";
        const fontScale = percentAttr(child, "fontScale");
        if (fontScale !== undefined) body.fontScalePct = quantizeScalePct(Math.min(100, Math.max(10, fontScale / 1000)));
        const reduction = percentAttr(child, "lnSpcReduction");
        if (reduction !== undefined) body.lineScalePct = quantizeScalePct(Math.min(100, Math.max(80, 100 - reduction / 1000)));
        continue;
      }
      if (local === "spAutoFit") {
        body.autofit = "resizeShape";
        continue;
      }
      // An explicit a:noAutofit is "none" — kept so it still overrides the
      // layout's inherited autofit on re-save.
    }
    kept.push({ rank: BODY_SLOT_RANK[slot], xml: preservedChildXml(child) });
  }

  const result: { body?: TextBodyProps; pptxBodyPr?: string } = {};
  if (Object.keys(body).length > 0) result.body = body;
  const slot = buildSlot("a:bodyPr", attrs, kept);
  if (slot) result.pptxBodyPr = slot;
  return result;
}

// --- relationship lookup (write) ---------------------------------------------

/** What the serialiser resolves a picture bullet or a link against: the
 * slide's (carried-forward + freshly minted) relationships, its part path, and
 * the part name each deck slide is saved under (index = slide number − 1). */
export interface RelLookup {
  rels: Map<string, Relationship>;
  slidePartPath: string;
  slideParts: readonly string[];
}

/** The image relationship id resolving to `mediaId`, or undefined. */
export function imageRelId(lookup: Pick<RelLookup, "rels" | "slidePartPath">, mediaId: string): string | undefined {
  for (const rel of lookup.rels.values()) {
    if (rel.type === IMAGE_RELATIONSHIP && resolvePartPath(lookup.slidePartPath, rel.target) === mediaId) {
      return rel.id;
    }
  }
  return undefined;
}

/** The hyperlink relationship id targeting `href`, or undefined. */
export function hyperlinkRelId(lookup: RelLookup, href: string): string | undefined {
  for (const rel of lookup.rels.values()) {
    if (rel.type === HYPERLINK_RELATIONSHIP && rel.target === href) {
      return rel.id;
    }
  }
  return undefined;
}

/** The slide relationship id targeting deck slide `slideNumber`'s saved part,
 * or undefined (no such slide, or no relationship minted). */
export function slideRelId(lookup: RelLookup, slideNumber: number): string | undefined {
  const part = lookup.slideParts[slideNumber - 1];
  if (!part) {
    return undefined;
  }
  for (const rel of lookup.rels.values()) {
    if (rel.type === SLIDE_RELATIONSHIP && resolvePartPath(lookup.slidePartPath, rel.target) === part) {
      return rel.id;
    }
  }
  return undefined;
}

// --- a:hlinkClick (write) ------------------------------------------------------

/** The `a:hlinkClick` for a run or frame link, or "" when it carries none or
 * its relationship does not resolve. Declares `xmlns:r` itself so it is
 * self-contained wherever it lands. */
export function hlinkClickXml(link: LinkFields, lookup: RelLookup | undefined): string {
  if (!lookup) {
    return "";
  }
  let relId: string | undefined;
  let action = "";
  if (link.linkHref) {
    relId = hyperlinkRelId(lookup, link.linkHref);
  } else if (link.linkSlide) {
    relId = slideRelId(lookup, link.linkSlide);
    action = ` action="${SLIDE_JUMP_ACTION}"`;
  }
  if (!relId) {
    return "";
  }
  const tooltip = link.linkTooltip ? ` tooltip="${escapeXmlAttribute(link.linkTooltip)}"` : "";
  return `<a:hlinkClick xmlns:r="${R_NS}" r:id="${escapeXmlAttribute(relId)}"${action}${tooltip}/>`;
}

/** `a:rPr@spc` for a run's letter spacing, or "". */
export function letterSpacingAttr(run: SlideRun): string {
  return run.letterSpacingPt ? ` spc="${Math.round(run.letterSpacingPt * 100)}"` : "";
}

// --- a:pPr (write) ------------------------------------------------------------

const ALGN: Record<NonNullable<SlideParagraph["align"]>, string> = {
  left: "l",
  center: "ctr",
  right: "r",
  justify: "just",
};

/** The modelled bullet element of a paragraph. */
function bulletXml(para: SlideParagraph, lookup: RelLookup | undefined): string {
  if (!para.bullet) {
    // Explicit, so the reader's body-role "default to disc" rule cannot re-add one.
    return `<a:buNone/>`;
  }
  if (para.bulletImageId && lookup) {
    const embed = imageRelId(lookup, para.bulletImageId);
    if (embed) {
      return `<a:buBlip><a:blip r:embed="${escapeXmlAttribute(embed)}"/></a:buBlip>`;
    }
  }
  const style = paragraphMarkerStyle(para) ?? "disc";
  if (para.bullet === "number") {
    const startAt = para.startAt !== undefined && para.startAt !== 1 ? ` startAt="${para.startAt}"` : "";
    return `<a:buAutoNum type="${AUTONUM_TYPE[style] ?? "arabicPeriod"}"${startAt}/>`;
  }
  const char = para.bulletChar ?? BULLET_CHAR[style] ?? "•";
  return `<a:buChar char="${escapeXmlAttribute(char)}"/>`;
}

/** A paragraph's `a:pPr` (SPEC-097 wire mapping): `marL`/`indent` (EMU, from
 * the effective indents — always for a list item or an indented level, only
 * when explicit on a plain level-0 paragraph), `lvl`, `algn`, the preserved
 * attributes; then the children in `CT_TextParagraphProperties` order —
 * `a:lnSpc` → `a:spcBef` → `a:spcAft` → preserved bullet colour/size/font →
 * the bullet → preserved `a:tabLst`/`a:defRPr`/`a:extLst`. A preserved
 * spacing element or bullet replaces the modelled one at its slot. */
export function pPrXml(para: SlideParagraph, lookup: RelLookup | undefined): string {
  const preserved = parseSlot(para.pptxPPr);
  const level = para.level ?? 0;
  const always = para.bullet !== undefined || level > 0;
  const emitted = new Set<string>();
  let attrs = "";
  const attr = (name: string, value: string | number): void => {
    emitted.add(name);
    attrs += ` ${name}="${value}"`;
  };
  if (always || para.indentCm !== undefined) {
    attr("marL", cmToEmu(effectiveIndentCm(para)));
  }
  if (always || para.hangingCm !== undefined) {
    const hanging = cmToEmu(effectiveHangingCm(para));
    attr("indent", hanging === 0 ? 0 : -hanging);
  }
  if (para.level) attr("lvl", para.level);
  if (para.align && para.align !== "left") attr("algn", ALGN[para.align]);
  for (const [name, xml] of preserved.attrs) {
    if (!emitted.has(name)) attrs += xml;
  }

  const bySlot = new Map<PPrSlot, string>();
  for (const child of preserved.children) {
    const slot = pPrSlotOf(child.local);
    bySlot.set(slot, (bySlot.get(slot) ?? "") + child.xml);
  }
  const lnSpc =
    bySlot.get("lnSpc") ??
    (para.lineSpacingPt !== undefined
      ? `<a:lnSpc><a:spcPts val="${Math.round(para.lineSpacingPt * 100)}"/></a:lnSpc>`
      : para.lineSpacing !== undefined && para.lineSpacing !== 1
        ? `<a:lnSpc><a:spcPct val="${Math.round(para.lineSpacing * 100000)}"/></a:lnSpc>`
        : "");
  const spcBef =
    bySlot.get("spcBef") ?? (para.spaceBeforePt ? `<a:spcBef><a:spcPts val="${Math.round(para.spaceBeforePt * 100)}"/></a:spcBef>` : "");
  const spcAft =
    bySlot.get("spcAft") ?? (para.spaceAfterPt ? `<a:spcAft><a:spcPts val="${Math.round(para.spaceAfterPt * 100)}"/></a:spcAft>` : "");
  const bullet = bySlot.get("bullet") ?? bulletXml(para, lookup);
  return (
    `<a:pPr${attrs}>${lnSpc}${spcBef}${spcAft}${bySlot.get("buProps") ?? ""}${bullet}${bySlot.get("tail") ?? ""}</a:pPr>`
  );
}

// --- a:bodyPr (write) -----------------------------------------------------------

/** A frame's `a:bodyPr`: the bare `<a:bodyPr/>` when it has neither modelled
 * nor preserved body properties (every pre-SPEC-097 golden), else insets,
 * columns, anchor, `wrap="none"` and the preserved attributes, then the
 * children in `CT_TextBodyProperties` order — preserved `a:prstTxWarp` → the
 * one autofit child (modelled shrink/resize, else a preserved explicit
 * `a:noAutofit`) → preserved `a:scene3d`/`a:sp3d`/`a:flatTx`/`a:extLst`.
 * `fontScale`/`lnSpcReduction` are integer 1000ths of a percent — that
 * quantisation is SPEC-097's one sanctioned round-trip delta; each is written
 * exactly when the model stores the scale. */
export function bodyPrXml(body: TextBodyProps | undefined, pptxBodyPr: string | undefined): string {
  if (!body && !pptxBodyPr) {
    return `<a:bodyPr/>`;
  }
  const preserved = parseSlot(pptxBodyPr);
  const emitted = new Set<string>();
  let attrs = "";
  const attr = (name: string, value: string | number): void => {
    emitted.add(name);
    attrs += ` ${name}="${value}"`;
  };
  if (body?.insetsCm) {
    attr("lIns", cmToEmu(body.insetsCm.left));
    attr("tIns", cmToEmu(body.insetsCm.top));
    attr("rIns", cmToEmu(body.insetsCm.right));
    attr("bIns", cmToEmu(body.insetsCm.bottom));
  }
  if (body?.columns && body.columns.count > 1) {
    attr("numCol", body.columns.count);
    attr("spcCol", cmToEmu(body.columns.gapCm));
  }
  if (body?.anchor) attr("anchor", ANCHOR_WIRE[body.anchor]);
  if (body?.wrap === false) attr("wrap", "none");
  for (const [name, xml] of preserved.attrs) {
    if (!emitted.has(name)) attrs += xml;
  }

  const bySlot = new Map<BodySlot, string>();
  for (const child of preserved.children) {
    const slot = bodySlotOf(child.local);
    bySlot.set(slot, (bySlot.get(slot) ?? "") + child.xml);
  }
  let autofit = bySlot.get("autofit") ?? "";
  if (body?.autofit === "shrink") {
    const fontScale = body.fontScalePct !== undefined ? ` fontScale="${Math.round(body.fontScalePct * 1000)}"` : "";
    const reduction =
      body.lineScalePct !== undefined ? ` lnSpcReduction="${Math.round((100 - body.lineScalePct) * 1000)}"` : "";
    autofit = `<a:normAutofit${fontScale}${reduction}/>`;
  } else if (body?.autofit === "resizeShape") {
    autofit = `<a:spAutoFit/>`;
  }
  const children = `${bySlot.get("warp") ?? ""}${autofit}${bySlot.get("tail") ?? ""}`;
  return children ? `<a:bodyPr${attrs}>${children}</a:bodyPr>` : `<a:bodyPr${attrs}/>`;
}

// --- relationship needs (writer) -----------------------------------------------

/** The paragraphs a frame owns (text/shape bodies, table cells). */
function paragraphListsOf(frame: SlideFrame): SlideParagraph[][] {
  if (frame.kind === "text") return [frame.paragraphs];
  if (frame.kind === "shape") return frame.paragraphs ? [frame.paragraphs] : [];
  if (frame.kind === "table") return frame.rows.flatMap((row) => row.cells.map((cell) => cell.paragraphs));
  return [];
}

/** Every picture-bullet media part a frame's paragraphs reference, in order. */
export function bulletImageIdsOf(frame: SlideFrame): string[] {
  const ids: string[] = [];
  for (const paragraphs of paragraphListsOf(frame)) {
    for (const para of paragraphs) {
      if (para.bullet && para.bulletImageId) ids.push(para.bulletImageId);
    }
  }
  return ids;
}

/** The relationships a slide's model-serialised frames need, in first-use
 * order: image parts (image frames and picture bullets, groups descended),
 * external hrefs and in-deck slide numbers (linkable frames and their runs,
 * table-cell runs; only authored groups — a verbatim group keeps its own
 * relationships). */
export interface RelNeeds {
  images: string[];
  hrefs: string[];
  slides: number[];
}

export function collectRelNeeds(frames: SlideFrame[]): RelNeeds {
  const needs: RelNeeds = { images: [], hrefs: [], slides: [] };
  const image = new Set<string>();
  const href = new Set<string>();
  const slide = new Set<number>();
  const addLink = (item: LinkFields): void => {
    if (item.linkHref) {
      if (!href.has(item.linkHref)) {
        href.add(item.linkHref);
        needs.hrefs.push(item.linkHref);
      }
    } else if (item.linkSlide) {
      if (!slide.has(item.linkSlide)) {
        slide.add(item.linkSlide);
        needs.slides.push(item.linkSlide);
      }
    }
  };
  const addImage = (id: string): void => {
    if (!image.has(id)) {
      image.add(id);
      needs.images.push(id);
    }
  };
  const walk = (list: SlideFrame[], authored: boolean): void => {
    for (const frame of list) {
      if (frame.kind === "image") addImage(frame.mediaId);
      for (const id of bulletImageIdsOf(frame)) addImage(id);
      if (frame.kind === "group") {
        walk(frame.children, authored && frame.xml === undefined);
        continue;
      }
      if (!authored) continue;
      if (isLinkableKind(frame.kind)) addLink(frame);
      for (const paragraphs of paragraphListsOf(frame)) {
        for (const para of paragraphs) for (const run of para.runs) addLink(run);
      }
    }
  };
  walk(frames, true);
  return needs;
}
