// Reads the legacy VML picture markup a `w:pict` carries — the form every
// Word 97-2003 document keeps after conversion to `.docx`, and the one some
// LibreOffice and Google Docs exports still write:
//
//   <w:pict>
//     <v:shape type="#_x0000_t75" style="width:99.75pt;height:75pt">
//       <v:imagedata r:id="rId5" o:title="photo"/>
//     </v:shape>
//   </w:pict>
//
// VML predates DrawingML entirely: no `wp:extent`, no `a:blip` — the size
// lives in a CSS-ish `style` attribute and the media reference is `r:id`, not
// `r:embed`. `w:pict` has never been in MODELLED_RUN_CHILDREN, so a run
// holding one has always gone to the opaque-inline path, which round-trips
// but renders as nothing: an old document's pictures simply did not appear.
//
// Scope is pictures only. A `w:pict` holding a VML *shape* (`v:rect`,
// `v:oval`, a `v:shape` with a `path`), a horizontal rule (`o:hr="t"`) or a
// text box returns null and stays preserved exactly as before — VML's fill,
// stroke and geometry model is a second shape language, and half-reading it
// would paint shapes this module cannot describe.
import type { Image, ImageWrapMode } from "../documentModel";
import { contentTypeForPartName, registerPendingMedia } from "../core/media";
import { decodeImageDimensions } from "../core/imageDimensions";
import { firstChildElement, resolvePartPath, type Relationship } from "../core/xml";
import { IMAGE_RELATIONSHIP, type ParsedDrawing } from "./drawing";

const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const VML_NAMESPACE = "urn:schemas-microsoft-com:vml";
const WORD10_NAMESPACE = "urn:schemas-microsoft-com:office:word";
const OFFICE_NAMESPACE = "urn:schemas-microsoft-com:office:office";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** VML lengths are CSS lengths. Word writes `pt` for pictures; the other
 * absolute units appear in hand-authored and converted files. Percentages and
 * `auto` are relative to something this module cannot resolve, so they yield
 * undefined and the picture stays preserved. */
const CM_PER_UNIT: Record<string, number> = {
  pt: 2.54 / 72,
  pc: 2.54 / 6,
  in: 2.54,
  cm: 1,
  mm: 0.1,
  px: 2.54 / 96,
};

function vmlLengthToCm(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const match = /^\s*(-?[\d.]+)\s*([a-z%]*)\s*$/i.exec(value);
  if (!match) {
    return undefined;
  }
  const n = Number(match[1]);
  if (!Number.isFinite(n)) {
    return undefined;
  }
  // A bare number is VML's own coordinate unit, not a CSS length — not a size
  // this module can turn into centimetres.
  const factor = CM_PER_UNIT[match[2].toLowerCase()];
  return factor === undefined ? undefined : n * factor;
}

/** The `style` attribute parsed into its declarations, lower-cased keys. */
function parseStyle(style: string | null): Map<string, string> {
  const declarations = new Map<string, string>();
  if (!style) {
    return declarations;
  }
  for (const part of style.split(";")) {
    const colon = part.indexOf(":");
    if (colon <= 0) {
      continue;
    }
    declarations.set(part.slice(0, colon).trim().toLowerCase(), part.slice(colon + 1).trim());
  }
  return declarations;
}

/** The wrap mode a VML picture's `w10:wrap` and positioning describe.
 *
 * An absolutely positioned picture maps to `"infront"`/`"behind"`, which
 * render at the anchor origin — the same approximation this model already
 * makes for every floating DrawingML picture (renderFormatting.ts pins a
 * floating Image to `left: 0; top: 0`), so VML adds no new inaccuracy here.
 * Returns null for a wrap VML can express and the model cannot, which keeps
 * the picture preserved rather than flowed wrongly. */
function wrapModeOf(shapeEl: Element, style: Map<string, string>): ImageWrapMode | null {
  const wrapEl = firstChildElement(shapeEl, WORD10_NAMESPACE, "wrap");
  const wrapType = wrapEl?.getAttribute("type");
  if (wrapType === "square" || wrapType === "topAndBottom") {
    return wrapType === "square" ? "square" : "topAndBottom";
  }
  if (wrapType === "tight" || wrapType === "through") {
    // Contour wrapping — not one of the model's four modes.
    return null;
  }
  if (style.get("position") !== "absolute") {
    return "inline";
  }
  // `none`, or no w10:wrap at all on an absolutely positioned shape: it
  // floats over or under the text, chosen by z-index the way VML does.
  const zIndex = Number(style.get("z-index") ?? "0");
  return Number.isFinite(zIndex) && zIndex < 0 ? "behind" : "infront";
}

/** The single VML element inside a `w:pict` that could be a picture, or null
 * when the `w:pict` holds something else. `v:shape` is what Word writes for a
 * picture; `v:rect` also accepts a `v:imagedata`. A horizontal rule
 * (`o:hr="t"`) is a `v:rect` with no image at all — excluded here so it keeps
 * its verbatim preservation rather than resolving to a null picture later. */
function pictureShapeOf(pictEl: Element): Element | null {
  let shape: Element | null = null;
  for (const child of Array.from(pictEl.children)) {
    if (child.namespaceURI === VML_NAMESPACE && child.localName === "shapetype") {
      // The shared shape definition Word emits beside the picture; it carries
      // no picture of its own and is regenerated on save.
      continue;
    }
    if (child.namespaceURI !== VML_NAMESPACE) {
      return null;
    }
    if (child.localName !== "shape" && child.localName !== "rect") {
      return null;
    }
    if (shape) {
      return null; // Several shapes: not the single picture this models.
    }
    if (child.getAttributeNS(OFFICE_NAMESPACE, "hr") === "t") {
      return null;
    }
    shape = child;
  }
  return shape;
}

/** Parses a `w:pict` into a modelled Image + MediaPart, or null when it is
 * not the single embedded picture this module models — a VML shape, a text
 * box, a horizontal rule, a linked (not embedded) image, or a wrap mode the
 * model has no equivalent for. The caller's SPEC-008 opaque-inline fallback
 * then preserves it verbatim, exactly as it did for every `w:pict` before. */
export function parseVmlPicture(
  pictEl: Element,
  mainRels: Map<string, Relationship>,
  mainPartPath: string,
  parts: Map<string, Uint8Array>,
): ParsedDrawing | null {
  if (pictEl.namespaceURI !== WORD_NAMESPACE || pictEl.localName !== "pict") {
    return null;
  }
  const shapeEl = pictureShapeOf(pictEl);
  if (!shapeEl) {
    return null;
  }
  const imagedata = firstChildElement(shapeEl, VML_NAMESPACE, "imagedata");
  if (!imagedata) {
    return null;
  }
  // A `v:imagedata` references its media with `r:id`; `r:href` is a LINKED
  // image whose bytes live outside the package and which this model, which
  // only carries embedded media, cannot represent.
  const relId = imagedata.getAttributeNS(R_NAMESPACE, "id");
  if (!relId) {
    return null;
  }
  const rel = mainRels.get(relId);
  if (!rel || rel.type !== IMAGE_RELATIONSHIP) {
    return null;
  }
  const partPath = resolvePartPath(mainPartPath, rel.target);
  const bytes = parts.get(partPath);
  const contentType = contentTypeForPartName(partPath);
  if (!bytes || !contentType) {
    return null;
  }

  const style = parseStyle(shapeEl.getAttribute("style"));
  const widthCm = vmlLengthToCm(style.get("width"));
  const heightCm = vmlLengthToCm(style.get("height"));
  if (widthCm === undefined || heightCm === undefined || widthCm <= 0 || heightCm <= 0) {
    return null;
  }
  const wrap = wrapModeOf(shapeEl, style);
  if (!wrap) {
    return null;
  }

  // `o:title` is VML's alt text; `alt` on the shape is the older spelling
  // Word 2003 wrote, and converted documents carry either.
  const altText =
    imagedata.getAttributeNS(OFFICE_NAMESPACE, "title") || shapeEl.getAttribute("alt") || undefined;

  const natural = decodeImageDimensions(bytes, contentType);
  const image: Image = {
    type: "image",
    mediaId: partPath,
    naturalWidthPx: natural?.widthPx ?? Math.round((widthCm / 2.54) * 96),
    naturalHeightPx: natural?.heightPx ?? Math.round((heightCm / 2.54) * 96),
    widthCm,
    heightCm,
    wrap,
  };
  if (altText) {
    image.altText = altText;
  }

  // Deferred encode, exactly as drawing.ts's own picture path does it — the
  // raw bytes ride in the shared pending-media registry until a consumer
  // forces the encode through `ensureMediaEncoded`.
  const media: ParsedDrawing["media"] = { mediaId: partPath, contentType, data: "" };
  registerPendingMedia(media, bytes);
  return { image, media };
}
