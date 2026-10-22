// Reads/writes the DrawingML `w:drawing` markup SPEC-024 models: a single
// picture (`pic:pic`), either inline (`wp:inline`) or floating
// (`wp:anchor`). Anything else a `w:drawing` can hold — charts, shapes,
// textboxes, multiple graphics — is not modelled; `parseDrawing` returns
// null for those, and the caller (reader.ts's collectRuns) keeps its
// existing SPEC-008 opaque-inline fallback, exactly as it did before this
// spec. This module only ever narrows what is already preserved, never
// widens what gets dropped.
import type { Image, ImageWrapMode } from "../documentModel";
import { contentTypeForPartName, registerPendingMedia, type MediaPart } from "../core/media";
import { decodeImageDimensions } from "../core/imageDimensions";
import { escapeXmlAttribute, firstChildElement, resolvePartPath, type Relationship } from "../core/xml";
import { cmToEmu, emuToCm } from "../core/units";

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
// The same editor-only extension namespace shape.ts/chart.ts/connector.ts use
// to round-trip `groupOrdinal`. Image reached SPEC-051's grouping through
// DrawingPlacement but never carried the attribute, so a grouped picture lost
// its group on every .docx save; it is read and written here for the same
// reason the other three drawing kinds do it.
const DE_NAMESPACE = "urn:docu-edit:shape";

export const IMAGE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/** One parsed `w:drawing`: the modelled image (its `mediaId` is the source
 * package's media part path, e.g. "word/media/image1.png" — stable and
 * already unique within the package, so no id needs minting) plus that part
 * itself, ready to add to `DocumentModel.media`. */
export interface ParsedDrawing {
  image: Image;
  media: MediaPart;
}

/** `wp:anchor/@relativeHeight`'s origin — the same base shape.ts/chart.ts
 * use, so a zOrder means the same thing across every drawing kind. */
const RELATIVE_HEIGHT_BASE = 251658240;

/** A `wp:positionH`/`wp:positionV`'s `wp:posOffset`, in cm. Undefined when
 * the anchor positions relatively (`wp:align`) instead of by offset — the
 * model carries no alignment, so "absent" leaves the wrap mode's default
 * (DrawingPlacement's "never recentre" rule). */
function anchorOffsetCm(container: Element, positionName: "positionH" | "positionV"): number | undefined {
  const position = firstChildElement(container, WP_NAMESPACE, positionName);
  const posOffset = position ? firstChildElement(position, WP_NAMESPACE, "posOffset") : undefined;
  if (!posOffset?.textContent) {
    return undefined;
  }
  const emu = Number(posOffset.textContent);
  return Number.isFinite(emu) ? emuToCm(emu) : undefined;
}

function wrapModeOf(anchorOrInline: Element, isAnchor: boolean): ImageWrapMode | null {
  if (!isAnchor) {
    return "inline";
  }
  if (firstChildElement(anchorOrInline, WP_NAMESPACE, "wrapSquare")) {
    return "square";
  }
  const wrapNone = firstChildElement(anchorOrInline, WP_NAMESPACE, "wrapNone");
  if (wrapNone) {
    return anchorOrInline.getAttribute("behindDoc") === "1" ? "behind" : "infront";
  }
  // wrapTight/wrapThrough/wrapTopAndBottom/wrapPolygon — not one of SPEC-024's
  // four modes; the caller falls back to preserving the drawing opaquely.
  return null;
}

/** Parses a `w:drawing` element into a modelled Image + MediaPart, or null
 * when it is not a single, simply-wrapped picture this spec models (a
 * chart, shape, textbox, or an unsupported wrap mode) — the caller's
 * existing opaque-inline fallback then preserves it verbatim (SPEC-008). */
export function parseDrawing(
  drawingEl: Element,
  mainRels: Map<string, Relationship>,
  mainPartPath: string,
  parts: Map<string, Uint8Array>,
): ParsedDrawing | null {
  const children = Array.from(drawingEl.children).filter((el) => el.namespaceURI === WP_NAMESPACE);
  if (children.length !== 1) {
    return null;
  }
  const container = children[0];
  const isAnchor = container.localName === "anchor";
  if (!isAnchor && container.localName !== "inline") {
    return null;
  }
  const wrap = wrapModeOf(container, isAnchor);
  if (!wrap) {
    return null;
  }

  const extent = firstChildElement(container, WP_NAMESPACE, "extent");
  const cx = extent ? Number(extent.getAttribute("cx")) : NaN;
  const cy = extent ? Number(extent.getAttribute("cy")) : NaN;
  if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) {
    return null;
  }

  const graphic = firstChildElement(container, A_NAMESPACE, "graphic");
  const graphicData = graphic ? firstChildElement(graphic, A_NAMESPACE, "graphicData") : undefined;
  const picEl = graphicData ? firstChildElement(graphicData, PIC_NAMESPACE, "pic") : undefined;
  if (!picEl) {
    return null;
  }

  const docPr = firstChildElement(container, WP_NAMESPACE, "docPr");
  const altText = docPr?.getAttribute("descr") || undefined;
  const groupOrdinalAttr = docPr?.getAttributeNS(DE_NAMESPACE, "group");
  const groupOrdinal = groupOrdinalAttr ? Number(groupOrdinalAttr) : undefined;

  // Where the picture actually sits (SPEC-024 modelled only its wrap mode, so
  // every floating image opened at its anchor's origin regardless of what the
  // file said). Read exactly the way shape.ts reads a `wps:wsp`'s anchor, so
  // an Image and a Shape floating at the same spot land at the same spot.
  let offsetXCm: number | undefined;
  let offsetYCm: number | undefined;
  let zOrder: number | undefined;
  if (isAnchor) {
    offsetXCm = anchorOffsetCm(container, "positionH");
    offsetYCm = anchorOffsetCm(container, "positionV");
    const relativeHeight = container.getAttribute("relativeHeight");
    if (relativeHeight) {
      const parsed = Number(relativeHeight);
      if (Number.isFinite(parsed)) {
        zOrder = parsed - RELATIVE_HEIGHT_BASE;
      }
    }
  }

  const parsed = imageFromPic(picEl, { widthCm: emuToCm(cx), heightCm: emuToCm(cy), wrap }, {
    mainRels,
    mainPartPath,
    parts,
  });
  if (!parsed) {
    return null;
  }
  const image = parsed.image;
  if (altText) {
    image.altText = altText;
  }
  if (groupOrdinal !== undefined && Number.isFinite(groupOrdinal)) {
    image.groupOrdinal = groupOrdinal;
  }
  if (offsetXCm !== undefined) {
    image.offsetXCm = offsetXCm;
  }
  if (offsetYCm !== undefined) {
    image.offsetYCm = offsetYCm;
  }
  if (wrap !== "inline") {
    image.relativeToH = "column";
    image.relativeToV = "paragraph";
  }
  if (zOrder !== undefined) {
    image.zOrder = zOrder;
  }
  return parsed;
}

/** The size + wrap a `pic:pic` cannot supply for itself — the picture-shaped
 * twin of shape.ts's `WspBox`. */
export interface PicBox {
  widthCm: number;
  heightCm: number;
  wrap: ImageWrapMode;
}

/** The package lookups resolving an `a:blip/@r:embed` to real bytes. */
export interface PicMediaContext {
  mainRels: Map<string, Relationship>;
  mainPartPath: string;
  parts: Map<string, Uint8Array>;
}

/** A shape picture-fill's `a:blip/@r:embed` resolver (SPEC-073 AC-4): the
 * same relationship→part→MediaPart walk imageFromPic does, packaged for
 * shape.ts's styling parser. Deduplicates into `media` by part path, so a
 * picture fill and an Image of the same file share one MediaPart. */
export function blipMediaResolver(ctx: PicMediaContext, media: MediaPart[]): (embedId: string) => string | null {
  return (embedId) => {
    const rel = ctx.mainRels.get(embedId);
    if (!rel || rel.type !== IMAGE_RELATIONSHIP) {
      return null;
    }
    const partPath = resolvePartPath(ctx.mainPartPath, rel.target);
    const bytes = ctx.parts.get(partPath);
    const contentType = contentTypeForPartName(partPath);
    if (!bytes || !contentType) {
      return null;
    }
    if (!media.some((part) => part.mediaId === partPath)) {
      const part: MediaPart = { mediaId: partPath, contentType, data: "" };
      registerPendingMedia(part, bytes);
      media.push(part);
    }
    return partPath;
  };
}

/** Builds an Image + MediaPart from a bare `pic:pic` and an
 * externally-resolved box — everything the picture markup itself carries, and
 * nothing that lives on a container. Null when the blip does not resolve to
 * an image part this package actually holds.
 *
 * Split out of parseDrawing so a `wpg:wgp` group member picture — which has
 * no `wp:anchor` of its own — reads through the same code as a top-level one
 * (see group.ts). The caller adds the container-derived fields: alt text,
 * offsets, z-order and `groupOrdinal`. */
export function imageFromPic(picEl: Element, box: PicBox, ctx: PicMediaContext): ParsedDrawing | null {
  const blipFill = firstChildElement(picEl, PIC_NAMESPACE, "blipFill");
  const blip = blipFill ? firstChildElement(blipFill, A_NAMESPACE, "blip") : undefined;
  const relId = blip ? blip.getAttributeNS(R_NAMESPACE, "embed") : null;
  if (!relId) {
    return null;
  }
  const rel = ctx.mainRels.get(relId);
  if (!rel || rel.type !== IMAGE_RELATIONSHIP) {
    return null;
  }
  const partPath = resolvePartPath(ctx.mainPartPath, rel.target);
  const bytes = ctx.parts.get(partPath);
  const contentType = contentTypeForPartName(partPath);
  if (!bytes || !contentType) {
    return null;
  }

  const natural = decodeImageDimensions(bytes, contentType);
  const image: ParsedDrawing["image"] = {
    type: "image",
    mediaId: partPath,
    // Falls back to the displayed size when the bytes' own header can't be
    // read — "Reset size" then resets to what the source already showed,
    // the closest available approximation.
    naturalWidthPx: natural?.widthPx ?? Math.round((box.widthCm / 2.54) * 96),
    naturalHeightPx: natural?.heightPx ?? Math.round((box.heightCm / 2.54) * 96),
    widthCm: box.widthCm,
    heightCm: box.heightCm,
    wrap: box.wrap,
  };

  // `data` is filled in later — see parseDocx's encode pass: encoding here,
  // per drawing, put ~0.6s of the benchmark fixture's open inside the walk.
  // The raw bytes ride in the shared pending-media registry (core/media)
  // keyed by the part object, so every intermediate collector (body walk,
  // notes part, header/footer parts) forwards them without signature
  // changes, and every consumer that reads `data` can force the encode
  // through `ensureMediaEncoded`.
  const media: ParsedDrawing["media"] = { mediaId: partPath, contentType, data: "" };
  registerPendingMedia(media, bytes);
  return { image, media };
}

/** The `wp:wrap*` element for a wrap mode (the inverse of wrapModeOf), plus
 * `wp:anchor`'s own `behindDoc` attribute value — "square"/"behind"/
 * "infront" only; "inline" has no anchor at all (see drawingXml). */
function anchorWrapParts(wrap: ImageWrapMode): { behindDoc: string; wrapXml: string } {
  if (wrap === "square") {
    return { behindDoc: "0", wrapXml: '<wp:wrapSquare wrapText="bothSides"/>' };
  }
  if (wrap === "behind") {
    return { behindDoc: "1", wrapXml: "<wp:wrapNone/>" };
  }
  return { behindDoc: "0", wrapXml: "<wp:wrapNone/>" };
}

/** Serialises an Image + the relationship id its media was written under
 * into a `w:drawing` element (wrapped by the caller in `<w:r>`, itself
 * wrapped in `<w:p>` — SPEC-024's Image block has no direct OOXML body-level
 * form; a picture must live inside a run inside a paragraph). `"square"`
 * wraps text on both sides; `"behind"`/`"infront"` float without wrapping,
 * ordered by `behindDoc`; `"inline"` uses `wp:inline`, which has no
 * position/wrap of its own — it simply flows with the paragraph. `docPrId`
 * is the caller's own running per-document counter (document order), kept
 * out of this module so repeated/parallel writes of the same model stay
 * deterministic — no shared mutable state. */
export function drawingXml(image: Image, relId: string, docPrId: number): string {
  const cx = cmToEmu(image.widthCm);
  const cy = cmToEmu(image.heightCm);
  const id = docPrId;
  const name = `Picture ${id}`;
  const descrAttr = image.altText ? ` descr="${escapeXmlAttribute(image.altText)}"` : "";
  const groupAttr = image.groupOrdinal !== undefined ? ` de:group="${image.groupOrdinal}"` : "";
  const docPr =
    `<wp:docPr id="${id}" name="${escapeXmlAttribute(name)}"${descrAttr}` +
    (groupAttr ? ` xmlns:de="${DE_NAMESPACE}"${groupAttr}` : "") +
    "/>";
  const graphic =
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic xmlns:pic="${PIC_NAMESPACE}">` +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="${escapeXmlAttribute(name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    "<pic:blipFill>" +
    `<a:blip xmlns:r="${R_NAMESPACE}" r:embed="${relId}"/>` +
    "<a:stretch><a:fillRect/></a:stretch>" +
    "</pic:blipFill>" +
    "<pic:spPr>" +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    "</pic:spPr>" +
    "</pic:pic>" +
    "</a:graphicData>" +
    "</a:graphic>";

  if (image.wrap === "inline") {
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

  const { behindDoc, wrapXml } = anchorWrapParts(image.wrap);
  // The picture's own z-order when it has one, falling back to document order
  // (`id`) for a picture that predates image placement — the same shape
  // shape.ts/chart.ts write.
  const relativeHeight = RELATIVE_HEIGHT_BASE + (image.zOrder ?? id);
  const offsetX = cmToEmu(image.offsetXCm ?? 0);
  const offsetY = cmToEmu(image.offsetYCm ?? 0);
  return (
    `<w:drawing xmlns:wp="${WP_NAMESPACE}">` +
    '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"' +
    ` relativeHeight="${relativeHeight}" behindDoc="${behindDoc}" locked="0"` +
    ' layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/>' +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${offsetX}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${offsetY}</wp:posOffset></wp:positionV>` +
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
