// The frame → pptx XML serialiser (SPEC-081) — the normative source for what a
// rebuilt `.pptx` slide part contains. SPEC-080 decided WHEN a slide part
// rebuilds (a dirty or added/duplicated slide); this module decides WHAT that
// part holds: the frames in z-order with effective run/paragraph properties
// written explicitly (fidelity AC), the slide's `preserved` XML re-emitted in
// schema position, and foreign/group frames re-emitted verbatim.
//
// Framework-free; imports only presentation siblings and src/lib/core.
import { GEOMETRY_TO_PRST } from "../../core/drawingMl";
import { computeOwnerGrid, type BorderSide } from "../../core/tables";
import { cmToEmu } from "../../core/units";
import { escapeXmlAttribute, escapeXmlText, resolvePartPath, type Relationship } from "../../core/xml";
import {
  DEFAULT_TABLE_CELL_PADDING_CM,
  type ChartFrame,
  type ImageFrame,
  type PresentationModel,
  type ShapeFrame,
  type Slide,
  type SlideFieldKind,
  type SlideFieldPlaceholder,
  type SlideFill,
  type SlideFrame,
  type SlideParagraph,
  type SlideRun,
  type SlideTableCell,
  type TableFrame,
  type TextFrame,
} from "../model";
import { writePptxTransition } from "../transitions";
import { writePptxTiming } from "./timing";
import { DEFAULT_DATE_FORMAT, fieldText, isTitleSlide } from "../headerFooter";
import { FIELD_GUID, FORMAT_TO_PPTX_DATETIME, PH_IDX_OF_KIND, PH_NAME_OF_KIND, PH_TYPE_OF_KIND } from "./fields";

const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STANDARD_ROOT_ATTRS = ` xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"`;

const IMAGE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const CHART_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DEFAULT_SIZE_PT = 18;
const DEFAULT_COLOR = "#000000";

const ALGN: Record<NonNullable<SlideParagraph["align"]>, string> = {
  left: "l",
  center: "ctr",
  right: "r",
  justify: "just",
};

/** A model colour (`#rrggbb`) as a DrawingML `srgbClr@val` (uppercase, no #). */
function hex(color: string): string {
  return color.replace(/^#/, "").toUpperCase();
}

function capitalise(text: string): string {
  return text.length > 0 ? text[0].toUpperCase() + text.slice(1) : text;
}

/** Serialises a slide to its `.pptx` slide-part XML. `rels` is the slide's
 * (original, carried-forward) relationships, used to resolve an `ImageFrame`'s
 * media part back to its `r:embed` id; `slidePartPath` anchors that relative
 * resolution. */
export function serializeSlideXml(
  slide: Slide,
  rels: Map<string, Relationship>,
  slidePartPath: string,
  fields: ResolvedField[] = [],
): string {
  const rootAttrs = slide.preserved?.["@root"] ?? STANDARD_ROOT_ATTRS;
  let nextId = 2; // p:cNvPr id=1 is the group scaffold.
  // Map each rebuilt frame to the `p:cNvPr id` the shapes get (2,3,… in
  // z-order, minted only for the text/shape/image frames the serialiser
  // rebuilds — group/foreign frames re-emit verbatim and consume no id), so
  // the timing's `p:spTgt spid` references match (SPEC-086 wire ids).
  const spidFor = new Map<string, number>();
  for (const frame of slide.frames) {
    if (mintsCNvPr(frame)) {
      spidFor.set(frame.id, nextId++);
    }
  }
  nextId = 2;
  const shapes = slide.frames
    .map((frame) => frameXml(frame, () => nextId++, rels, slidePartPath))
    .join("");
  // SPEC-093: the deck field placeholders follow the frame shapes — each
  // rewritten from its preserved `p:sp` (only the cached `a:t` refreshed) or,
  // for a newly-enabled field, built canonical.
  const fieldSps = fields.map((rf) => fieldPlaceholderXml(rf, slide, () => nextId++)).join("");
  const bg = slide.preserved?.["p:bg"] ?? "";
  const clrMapOvr = slide.preserved?.["p:clrMapOvr"] ?? `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`;
  // A modelled transition wins over any preserved one (they are never both
  // present — the reader drops the source XML when it maps the transition);
  // otherwise re-emit the preserved (out-of-subset) transition verbatim.
  const transition = slide.transition ? writePptxTransition(slide.transition) : (slide.preserved?.["p:transition"] ?? "");
  // Modelled animations win over any preserved timing (never both — the reader
  // drops the source when it models them); a present (even empty) `animations`
  // is the modelled state, so preserved timing is never re-emitted alongside.
  const timing = slide.animations !== undefined
    ? writePptxTiming(slide.animations, (id) => spidFor.get(id))
    : (slide.preserved?.["p:timing"] ?? "");
  const extLst = slide.preserved?.["p:extLst"] ?? "";
  return (
    `${XML_DECL}<p:sld${rootAttrs}>` +
    `<p:cSld>${bg}<p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    shapes +
    fieldSps +
    `</p:spTree></p:cSld>` +
    clrMapOvr +
    transition +
    timing +
    extLst +
    `</p:sld>`
  );
}

// --- notes-slide part (SPEC-083) -------------------------------------------

const NOTES_ROOT_ATTRS = ` xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"`;

/** Serialises a slide's speaker notes to a `.pptx` notes-slide part
 * (`p:notes`): the fixed scaffold — a slide-image placeholder (`p:ph
 * type="sldImg"`, ignored on read) and the body placeholder (`p:ph type="body"
 * idx="1"`) carrying one `a:p` per notes entry (single run, no explicit rPr,
 * so formatting is flat). Geometry is the default `p:notesSz`
 * (6858000×9144000 EMU) with a body placeholder inset by ~1.32cm margins. */
export function serializeNotesSlideXml(notes: string[], nonBodyPlaceholders = ""): string {
  const sldImg =
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/>` +
    `<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>` +
    `<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="685800" y="685800"/><a:ext cx="5486400" cy="3086100"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`;
  const body =
    notes.length > 0
      ? notes.map((line) => (line.length > 0 ? `<a:p><a:r><a:t>${escapeXmlText(line)}</a:t></a:r></a:p>` : `<a:p/>`)).join("")
      : `<a:p/>`;
  const bodySp =
    `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/>` +
    `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="685800" y="3884613"/><a:ext cx="5486400" cy="4127500"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody></p:sp>`;
  return (
    `${XML_DECL}<p:notes${NOTES_ROOT_ATTRS}><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    sldImg +
    bodySp +
    nonBodyPlaceholders +
    `</p:spTree></p:cSld></p:notes>`
  );
}

/** Re-emits a source notes part's non-body field placeholders — `p:ph` of type
 * `hdr`/`ftr`/`dt`/`sldNum` (SPEC-093 AC-92/93), which the SPEC-083 rebuild used
 * to drop — refreshing a `slidenum`/`datetime*` field's cached `a:t` the same
 * way a slide's does (a plain hdr/ftr run stays verbatim). */
export function preserveNotesFieldPlaceholders(
  sourceNotesXml: string,
  slideNumberText: string,
  dateText: string,
): string {
  const out: string[] = [];
  for (const match of sourceNotesXml.matchAll(/<p:sp(?:\s[^>]*)?>[\s\S]*?<\/p:sp>/g)) {
    const sp = match[0];
    if (!/<p:ph\b[^>]*\btype="(?:hdr|ftr|dt|sldNum)"/.test(sp)) {
      continue;
    }
    const fld = /<a:fld\b[^>]*\btype="([^"]*)"/.exec(sp);
    const type = fld?.[1];
    if (type === "slidenum") {
      out.push(rewriteFieldCachedText(sp, slideNumberText));
    } else if (type?.startsWith("datetime")) {
      out.push(rewriteFieldCachedText(sp, dateText));
    } else {
      out.push(sp);
    }
  }
  return out.join("");
}

function frameXml(
  frame: SlideFrame,
  mintId: () => number,
  rels: Map<string, Relationship>,
  slidePartPath: string,
): string {
  switch (frame.kind) {
    case "text":
      return textFrameXml(frame, mintId());
    case "shape":
      return shapeFrameXml(frame, mintId());
    case "image":
      return imageFrameXml(frame, mintId(), rels, slidePartPath);
    case "table":
      return tableFrameXml(frame, mintId());
    case "chart":
      // An untouched chart re-emits its verbatim graphicFrame (byte-identical,
      // AC-18/AC-19); a touched one emits the canonical graphicFrame against
      // the chart part's relationship id (SPEC-092).
      return frame.xml ?? serializeChartFrame(frame, mintId(), chartRelIdFor(frame, rels, slidePartPath) ?? "");
    case "group":
    case "foreign":
      return frame.xml; // re-emitted verbatim, in z-order (AC-35)
    default:
      return "";
  }
}

/** Whether a frame mints a `p:cNvPr id` when serialised (so the animation
 * `spid` pre-pass reserves ids in the same order the real pass mints them). A
 * chart that re-emits verbatim `xml` consumes no id, exactly like group/
 * foreign. */
function mintsCNvPr(frame: SlideFrame): boolean {
  if (frame.kind === "chart") {
    return !frame.xml;
  }
  return frame.kind === "text" || frame.kind === "shape" || frame.kind === "image" || frame.kind === "table";
}

/** The relationship id whose chart part is `frame.sourcePart`, or undefined
 * (a chart created here whose rel the writer mints alongside its part). */
function chartRelIdFor(frame: ChartFrame, rels: Map<string, Relationship>, slidePartPath: string): string | undefined {
  if (!frame.sourcePart) {
    return undefined;
  }
  for (const rel of rels.values()) {
    if (rel.type === CHART_RELATIONSHIP && resolvePartPath(slidePartPath, rel.target) === frame.sourcePart) {
      return rel.id;
    }
  }
  return undefined;
}

/** The canonical `p:graphicFrame` for a touched chart (SPEC-092's Reference):
 * a `p:nvGraphicFramePr` scaffold consuming one minted `p:cNvPr id`, a
 * `p:xfrm` in EMU, and an `a:graphic/a:graphicData` holding `<c:chart r:id>`.
 * `relId` is the slide relationship the writer wrote the chart part under. */
export function serializeChartFrame(frame: ChartFrame, id: number, relId: string): string {
  const rot = frame.rotationDeg ? ` rot="${Math.round(frame.rotationDeg * 60000)}"` : "";
  return (
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id}"/>` +
    `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>` +
    `<p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm${rot}><a:off x="${cmToEmu(frame.xCm)}" y="${cmToEmu(frame.yCm)}"/>` +
    `<a:ext cx="${cmToEmu(frame.widthCm)}" cy="${cmToEmu(frame.heightCm)}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${C_NS}">` +
    `<c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="${escapeXmlAttribute(relId)}"/>` +
    `</a:graphicData></a:graphic></p:graphicFrame>`
  );
}

// --- xfrm / spPr -----------------------------------------------------------

function xfrmXml(frame: SlideFrame): string {
  const rot = frame.rotationDeg ? ` rot="${Math.round(frame.rotationDeg * 60000)}"` : "";
  return (
    `<a:xfrm${rot}><a:off x="${cmToEmu(frame.xCm)}" y="${cmToEmu(frame.yCm)}"/>` +
    `<a:ext cx="${cmToEmu(frame.widthCm)}" cy="${cmToEmu(frame.heightCm)}"/></a:xfrm>`
  );
}

function fillXml(fill: SlideFill | undefined): string {
  if (!fill) return "";
  if (fill.type === "none") return `<a:noFill/>`;
  return `<a:solidFill><a:srgbClr val="${hex(fill.color)}"/></a:solidFill>`;
}

/** A frame's `a:ln` outline, or "" for none. */
function strokeXml(stroke: { color: string; widthPt: number } | undefined): string {
  if (!stroke) return "";
  return `<a:ln w="${Math.round(stroke.widthPt * 12700)}"><a:solidFill><a:srgbClr val="${hex(stroke.color)}"/></a:solidFill></a:ln>`;
}

// --- text frame ------------------------------------------------------------

function textFrameXml(frame: TextFrame, id: number): string {
  const phAttrs =
    (frame.phType ? ` type="${escapeXmlAttribute(frame.phType)}"` : "") +
    (frame.phIdx !== undefined ? ` idx="${frame.phIdx}"` : "");
  const ph = frame.role ? `<p:ph${phAttrs}/>` : "";
  const name = frame.role ? capitalise(frame.role) : "TextBox";
  const nvSpPr =
    `<p:nvSpPr><p:cNvPr id="${id}" name="${escapeXmlAttribute(name)}"/>` +
    `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>`;
  const spPr = `<p:spPr>${xfrmXml(frame)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fillXml(frame.fill)}${strokeXml(frame.stroke)}</p:spPr>`;
  return `<p:sp>${nvSpPr}${spPr}${txBodyXml(frame.paragraphs)}</p:sp>`;
}

/** A `p:txBody` from a frame's paragraphs. An empty body emits a single empty
 * `a:p` (the from-scratch placeholder shape). */
function txBodyXml(paragraphs: SlideParagraph[]): string {
  const body = paragraphs.length > 0 ? paragraphs.map(paragraphXml).join("") : `<a:p/>`;
  return `<p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody>`;
}

function paragraphXml(para: SlideParagraph): string {
  const runs = para.runs.length > 0 ? para.runs.map(runXml).join("") : `<a:endParaRPr lang="en-US"/>`;
  return `<a:p>${pPrXml(para)}${runs}</a:p>`;
}

/** A paragraph's `a:pPr`: `algn` (omitted for left), `lvl` (omitted for 0),
 * and an explicit bullet element — `a:buNone` for none so the reader's
 * body-role "default to disc" rule cannot re-add one (fidelity). */
function pPrXml(para: SlideParagraph): string {
  const algn = para.align && para.align !== "left" ? ` algn="${ALGN[para.align]}"` : "";
  const lvl = para.level ? ` lvl="${para.level}"` : "";
  const bullet =
    para.bullet === "disc"
      ? `<a:buChar char="•"/>`
      : para.bullet === "number"
        ? `<a:buAutoNum type="arabicPeriod"/>`
        : `<a:buNone/>`;
  return `<a:pPr${algn}${lvl}>${bullet}</a:pPr>`;
}

/** A run's `a:r`. Every effective value is written explicitly (fidelity AC):
 * `sz` and `a:solidFill` always (the reader drops the 18pt / #000000 defaults
 * back to `undefined`), `b`/`i`/`u` when set, `a:latin` when a family is set. */
function runXml(run: SlideRun): string {
  const sz = ` sz="${Math.round((run.sizePt ?? DEFAULT_SIZE_PT) * 100)}"`;
  const b = run.bold ? ` b="1"` : "";
  const i = run.italic ? ` i="1"` : "";
  const u = run.underline ? ` u="sng"` : "";
  const color = `<a:solidFill><a:srgbClr val="${hex(run.color ?? DEFAULT_COLOR)}"/></a:solidFill>`;
  const latin = run.fontFamily ? `<a:latin typeface="${escapeXmlAttribute(run.fontFamily)}"/>` : "";
  return `<a:r><a:rPr lang="en-US"${sz}${b}${i}${u}>${color}${latin}</a:rPr><a:t>${escapeXmlText(run.text)}</a:t></a:r>`;
}

// --- deck field placeholders (SPEC-093) ------------------------------------

/** A resolved field ready to serialise: the entry, its freshly computed text,
 * and the `a:fld@type` to emit (`slidenum`/`datetime*`), or undefined for a
 * plain `a:r` (footer text and a fixed/plain date). */
export interface ResolvedField {
  field: SlideFieldPlaceholder;
  text: string;
  fldType?: string;
}

/** The `a:fld@type` a field emits, or undefined for a plain recomputed `a:r`
 * (SPEC-093 AC-67/68/69). A number is `slidenum`; a footer and a fixed/
 * per-slide-literal date are plain runs; an auto date maps its deck format to a
 * `datetime*` (or a plain `a:r` for the two formats with no faithful mapping);
 * an out-of-subset date keeps its verbatim wire type. */
export function fieldEmitType(model: PresentationModel, field: SlideFieldPlaceholder): string | undefined {
  if (field.kind === "slideNumber") {
    return "slidenum";
  }
  if (field.kind === "footer") {
    return undefined;
  }
  if (field.wireType) {
    return field.wireType;
  }
  if (model.headerFooter?.dateSource === "fixed" || field.text !== undefined) {
    return undefined;
  }
  return FORMAT_TO_PPTX_DATETIME[model.headerFooter?.dateFormat ?? DEFAULT_DATE_FORMAT];
}

/** The field placeholders to emit for a slide (SPEC-093): none on a suppressed
 * title slide (AC-26), otherwise every `Slide.fields` entry with its resolved
 * text + wire type. */
export function resolveSlideFields(model: PresentationModel, slideIndex: number, instant: Date): ResolvedField[] {
  const slide = model.slides[slideIndex];
  if (!slide?.fields || slide.fields.length === 0) {
    return [];
  }
  if (isTitleSlide(model, slideIndex) && model.headerFooter?.showOnTitleSlide === false) {
    return [];
  }
  return slide.fields.map((field) => ({
    field,
    text: fieldText(model, slideIndex, field.kind, instant),
    fldType: fieldEmitType(model, field),
  }));
}

/** One field placeholder `p:sp`: the preserved source `p:sp` with only its
 * cached `a:t` refreshed (AC-66), or — when the slide has no preserved copy
 * (a newly-enabled field) — the canonical `p:sp` (AC-70). */
function fieldPlaceholderXml(rf: ResolvedField, slide: Slide, mintId: () => number): string {
  const preserved = slide.preserved?.[`p:ph:${PH_TYPE_OF_KIND[rf.field.kind]}`];
  if (preserved) {
    return rewriteFieldCachedText(preserved, rf.text);
  }
  return canonicalFieldSp(rf, mintId());
}

/** Replaces the first `<a:t>…</a:t>` (or self-closing `<a:t/>`) with the fresh
 * value, leaving the rest of the preserved `p:sp` byte-for-byte. */
function rewriteFieldCachedText(spXml: string, text: string): string {
  const replacement = `<a:t>${escapeXmlText(text)}</a:t>`;
  if (/<a:t>[\s\S]*?<\/a:t>/.test(spXml)) {
    return spXml.replace(/<a:t>[\s\S]*?<\/a:t>/, () => replacement);
  }
  return spXml.replace(/<a:t\/>/, () => replacement);
}

/** The canonical field placeholder `p:sp` (SPEC-093 Reference). A number/auto
 * date is an `<a:fld>`; a footer/fixed date is a plain `<a:r>`. Exported for the
 * golden check. */
export function canonicalFieldSp(rf: ResolvedField, id: number): string {
  const kind: SlideFieldKind = rf.field.kind;
  const type = PH_TYPE_OF_KIND[kind];
  const idx = PH_IDX_OF_KIND[kind];
  const name = PH_NAME_OF_KIND[kind];
  const algn = rf.field.align === "center" ? "ctr" : rf.field.align === "right" ? "r" : "l";
  const xfrm =
    `<a:xfrm><a:off x="${cmToEmu(rf.field.xCm)}" y="${cmToEmu(rf.field.yCm)}"/>` +
    `<a:ext cx="${cmToEmu(rf.field.widthCm)}" cy="${cmToEmu(rf.field.heightCm)}"/></a:xfrm>`;
  const inner = rf.fldType
    ? `<a:fld id="${FIELD_GUID[kind]}" type="${rf.fldType}"><a:t>${escapeXmlText(rf.text)}</a:t></a:fld>`
    : `<a:r><a:rPr lang="en-US"/><a:t>${escapeXmlText(rf.text)}</a:t></a:r>`;
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeXmlAttribute(name)}"/>` +
    `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
    `<p:nvPr><p:ph type="${type}" sz="quarter" idx="${idx}"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr>${xfrm}</p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="${algn}"/>${inner}</a:p></p:txBody></p:sp>`
  );
}

// --- shape frame -----------------------------------------------------------

function shapeFrameXml(frame: ShapeFrame, id: number): string {
  const prst = GEOMETRY_TO_PRST[frame.geometry] ?? "rect";
  const nvSpPr =
    `<p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`;
  const spPr =
    `<p:spPr>${xfrmXml(frame)}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${fillXml(frame.fill)}${strokeXml(frame.stroke)}</p:spPr>`;
  const txBody = frame.paragraphs && frame.paragraphs.length > 0 ? txBodyXml(frame.paragraphs) : "";
  return `<p:sp>${nvSpPr}${spPr}${txBody}</p:sp>`;
}

// --- image frame -----------------------------------------------------------

function imageFrameXml(
  frame: ImageFrame,
  id: number,
  rels: Map<string, Relationship>,
  slidePartPath: string,
): string {
  const embed = embedIdFor(frame.mediaId, rels, slidePartPath);
  const blip = embed ? `<a:blip r:embed="${escapeXmlAttribute(embed)}"/>` : `<a:blip/>`;
  return (
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrmXml(frame)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${strokeXml(frame.stroke)}</p:spPr></p:pic>`
  );
}

/** The `r:embed` id whose image relationship resolves to `mediaId`, carried
 * from the slide's original rels (this spec never mints image rels). */
function embedIdFor(mediaId: string, rels: Map<string, Relationship>, slidePartPath: string): string | undefined {
  for (const rel of rels.values()) {
    if (rel.type === IMAGE_RELATIONSHIP && resolvePartPath(slidePartPath, rel.target) === mediaId) {
      return rel.id;
    }
  }
  return undefined;
}

// --- table frame (SPEC-091) -------------------------------------------------

/** The DrawingML table `a:graphicData@uri`. */
export const PPTX_TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";

const DASH_FOR: Partial<Record<BorderSide["style"], string>> = {
  dashed: "dash",
  dotted: "sysDot",
};

/** A cell border edge → `a:lnL/R/T/B`. `"none"` paints `a:noFill`; a set style
 * emits width, solid colour, compound for double, and a preset dash. */
function borderLnXml(tag: string, side: BorderSide | undefined): string {
  if (!side) {
    return "";
  }
  if (side.style === "none") {
    return `<${tag}><a:noFill/></${tag}>`;
  }
  const w = ` w="${Math.round(side.widthPt * 12700)}"`;
  const cmpd = side.style === "double" ? ` cmpd="dbl"` : "";
  const dash = DASH_FOR[side.style] ? `<a:prstDash val="${DASH_FOR[side.style]}"/>` : "";
  return `<${tag}${w}${cmpd}><a:solidFill><a:srgbClr val="${hex(side.color)}"/></a:solidFill>${dash}</${tag}>`;
}

/** A cell's `a:tcPr`: the preserved verbatim string when the cell has one
 * (AC-12), else regenerated from the modelled borders / padding / anchor /
 * fill. */
function tcPrXml(cell: SlideTableCell): string {
  if (cell.pptxTcPr) {
    return cell.pptxTcPr;
  }
  const padEmu = cmToEmu(cell.paddingCm ?? DEFAULT_TABLE_CELL_PADDING_CM);
  const mar = ` marL="${padEmu}" marR="${padEmu}" marT="${padEmu}" marB="${padEmu}"`;
  const anchor =
    cell.verticalAlign === "center" ? ` anchor="ctr"` : cell.verticalAlign === "bottom" ? ` anchor="b"` : "";
  const borders =
    borderLnXml("a:lnL", cell.borders?.left) +
    borderLnXml("a:lnR", cell.borders?.right) +
    borderLnXml("a:lnT", cell.borders?.top) +
    borderLnXml("a:lnB", cell.borders?.bottom);
  const fill = cell.fill ? fillXml(cell.fill) : "";
  return `<a:tcPr${mar}${anchor}>${borders}${fill}</a:tcPr>`;
}

/** A table cell's `a:txBody` (the DrawingML-namespaced counterpart of a shape's
 * `p:txBody`). An empty cell emits a single empty `a:p`. */
function cellTxBodyXml(paragraphs: SlideParagraph[]): string {
  const body = paragraphs.length > 0 ? paragraphs.map(paragraphXml).join("") : `<a:p/>`;
  return `<a:txBody><a:bodyPr/><a:lstStyle/>${body}</a:txBody>`;
}

/** The table's `a:tblPr`: preserved verbatim when present (AC-12), else the
 * modelled look flags plus an optional table style id. */
function tblPrXml(frame: TableFrame): string {
  if (frame.pptxTblPr) {
    return frame.pptxTblPr;
  }
  const firstRow = frame.firstRow ? ` firstRow="1"` : "";
  const bandRow = frame.bandRow ? ` bandRow="1"` : "";
  const styleId = frame.styleId ? `<a:tableStyleId>${escapeXmlText(frame.styleId)}</a:tableStyleId>` : "";
  return styleId ? `<a:tblPr${firstRow}${bandRow}>${styleId}</a:tblPr>` : `<a:tblPr${firstRow}${bandRow}/>`;
}

/** The graphicFrame's `p:nvGraphicFramePr`: the preserved verbatim string with
 * its `p:cNvPr` id re-pointed at the minted `id` (so an animation `p:spTgt`
 * resolves — AC-11e), else the canonical scaffold. */
function nvGraphicFramePrXml(frame: TableFrame, id: number): string {
  if (frame.pptxGraphicFramePr) {
    return frame.pptxGraphicFramePr.replace(/id="[^"]*"/, `id="${id}"`);
  }
  return (
    `<p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/>` +
    `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>`
  );
}

/** The graphicFrame transform (`p:xfrm`, not `a:xfrm`). */
function pXfrmXml(frame: TableFrame): string {
  const rot = frame.rotationDeg ? ` rot="${Math.round(frame.rotationDeg * 60000)}"` : "";
  return (
    `<p:xfrm${rot}><a:off x="${cmToEmu(frame.xCm)}" y="${cmToEmu(frame.yCm)}"/>` +
    `<a:ext cx="${cmToEmu(frame.widthCm)}" cy="${cmToEmu(frame.heightCm)}"/></p:xfrm>`
  );
}

/** Serialises a `TableFrame` to its `p:graphicFrame` XML with the given
 * `p:cNvPr` id — the test seam the golden spec pins (SPEC-091 AC-16). */
export function serializeTableFrame(frame: TableFrame, id: number): string {
  return tableFrameXml(frame, id);
}

function tableFrameXml(frame: TableFrame, id: number): string {
  const grid = computeOwnerGrid(frame);
  const gridCols = frame.columnWidthsCm.map((w) => `<a:gridCol w="${cmToEmu(w)}"/>`).join("");
  const rows = frame.rows
    .map((row, r) => {
      const h = ` h="${cmToEmu(row.heightCm ?? 0.9)}"`;
      const cells = row.cells
        .map((cell, c) => {
          const owner = grid[r][c];
          if (owner.originRow !== r || owner.originCol !== c) {
            const hMerge = c > owner.originCol ? ` hMerge="1"` : "";
            const vMerge = r > owner.originRow ? ` vMerge="1"` : "";
            return `<a:tc${hMerge}${vMerge}><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>`;
          }
          const gs = cell.colSpan && cell.colSpan > 1 ? ` gridSpan="${cell.colSpan}"` : "";
          const rs = cell.rowSpan && cell.rowSpan > 1 ? ` rowSpan="${cell.rowSpan}"` : "";
          return `<a:tc${gs}${rs}>${cellTxBodyXml(cell.paragraphs)}${tcPrXml(cell)}</a:tc>`;
        })
        .join("");
      return `<a:tr${h}>${cells}</a:tr>`;
    })
    .join("");
  const tbl = `<a:tbl>${tblPrXml(frame)}<a:tblGrid>${gridCols}</a:tblGrid>${rows}</a:tbl>`;
  return (
    `<p:graphicFrame>${nvGraphicFramePrXml(frame, id)}${pXfrmXml(frame)}` +
    `<a:graphic><a:graphicData uri="${PPTX_TABLE_URI}">${tbl}</a:graphicData></a:graphic></p:graphicFrame>`
  );
}
