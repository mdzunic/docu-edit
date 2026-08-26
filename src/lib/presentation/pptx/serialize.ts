// The frame → pptx XML serialiser (SPEC-081) — the normative source for what a
// rebuilt `.pptx` slide part contains. SPEC-080 decided WHEN a slide part
// rebuilds (a dirty or added/duplicated slide); this module decides WHAT that
// part holds: the frames in z-order with effective run/paragraph properties
// written explicitly (fidelity AC), the slide's `preserved` XML re-emitted in
// schema position, and foreign/group frames re-emitted verbatim.
//
// Framework-free; imports only presentation siblings and src/lib/core.
import { GEOMETRY_TO_PRST } from "../../core/drawingMl";
import { cmToEmu } from "../../core/units";
import { escapeXmlAttribute, escapeXmlText, resolvePartPath, type Relationship } from "../../core/xml";
import type {
  ImageFrame,
  ShapeFrame,
  Slide,
  SlideFill,
  SlideFrame,
  SlideParagraph,
  SlideRun,
  TextFrame,
} from "../model";

const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STANDARD_ROOT_ATTRS = ` xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"`;

const IMAGE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
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
): string {
  const rootAttrs = slide.preserved?.["@root"] ?? STANDARD_ROOT_ATTRS;
  let nextId = 2; // p:cNvPr id=1 is the group scaffold.
  const shapes = slide.frames
    .map((frame) => frameXml(frame, () => nextId++, rels, slidePartPath))
    .join("");
  const bg = slide.preserved?.["p:bg"] ?? "";
  const clrMapOvr = slide.preserved?.["p:clrMapOvr"] ?? `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`;
  const transition = slide.preserved?.["p:transition"] ?? "";
  const timing = slide.preserved?.["p:timing"] ?? "";
  const extLst = slide.preserved?.["p:extLst"] ?? "";
  return (
    `${XML_DECL}<p:sld${rootAttrs}>` +
    `<p:cSld>${bg}<p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    shapes +
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
export function serializeNotesSlideXml(notes: string[]): string {
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
    `</p:spTree></p:cSld></p:notes>`
  );
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
    case "group":
    case "foreign":
      return frame.xml; // re-emitted verbatim, in z-order (AC-35)
    default:
      return "";
  }
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
