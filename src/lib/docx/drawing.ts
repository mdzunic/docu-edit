// Reads/writes the DrawingML `w:drawing` markup SPEC-024 models: a single
// picture (`pic:pic`), either inline (`wp:inline`) or floating
// (`wp:anchor`). Anything else a `w:drawing` can hold — charts, shapes,
// textboxes, multiple graphics — is not modelled; `parseDrawing` returns
// null for those, and the caller (reader.ts's collectRuns) keeps its
// existing SPEC-008 opaque-inline fallback, exactly as it did before this
// spec. This module only ever narrows what is already preserved, never
// widens what gets dropped.
import type { Image, ImageWrapMode } from "../documentModel";
import type { MediaPart } from "../core/media";
import { decodeImageDimensions } from "../core/imageDimensions";
import { bytesToBase64Async } from "../opaque";
import { firstChildElement, resolvePartPath, type Relationship } from "../core/xml";
import { cmToEmu, emuToCm } from "../core/units";

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const IMAGE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

export const CONTENT_TYPE_BY_EXTENSION: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  svg: "image/svg+xml",
};

/** The file extension a media content type is written with, the inverse of
 * CONTENT_TYPE_BY_EXTENSION's lookup — "jpeg" (not "jpg") is the canonical
 * spelling this writer picks for JPEG parts. */
const EXTENSION_BY_CONTENT_TYPE: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpeg",
  "image/gif": "gif",
  "image/svg+xml": "svg",
};

export function contentTypeForPartName(partName: string): string | undefined {
  const dot = partName.lastIndexOf(".");
  if (dot < 0) {
    return undefined;
  }
  return CONTENT_TYPE_BY_EXTENSION[partName.slice(dot + 1).toLowerCase()];
}

export function extensionForContentType(contentType: string): string {
  return EXTENSION_BY_CONTENT_TYPE[contentType] ?? "png";
}

/** One parsed `w:drawing`: the modelled image (its `mediaId` is the source
 * package's media part path, e.g. "word/media/image1.png" — stable and
 * already unique within the package, so no id needs minting) plus that part
 * itself, ready to add to `DocumentModel.media`. */
export interface ParsedDrawing {
  image: Image;
  media: MediaPart;
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
  const blipFill = picEl ? firstChildElement(picEl, PIC_NAMESPACE, "blipFill") : undefined;
  const blip = blipFill ? firstChildElement(blipFill, A_NAMESPACE, "blip") : undefined;
  const relId = blip ? blip.getAttributeNS(R_NAMESPACE, "embed") : null;
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

  const docPr = firstChildElement(container, WP_NAMESPACE, "docPr");
  const altText = docPr?.getAttribute("descr") || undefined;

  const widthCm = emuToCm(cx);
  const heightCm = emuToCm(cy);
  const natural = decodeImageDimensions(bytes, contentType);

  const image: ParsedDrawing["image"] = {
    type: "image",
    mediaId: partPath,
    // Falls back to the displayed size when the bytes' own header can't be
    // read — "Reset size" then resets to what the source already showed,
    // the closest available approximation.
    naturalWidthPx: natural?.widthPx ?? Math.round((widthCm / 2.54) * 96),
    naturalHeightPx: natural?.heightPx ?? Math.round((heightCm / 2.54) * 96),
    widthCm,
    heightCm,
    wrap,
  };
  if (altText) {
    image.altText = altText;
  }

  // `data` is filled in by parseDocx's single deferred pass (see
  // fillPendingMediaData): encoding here, synchronously per drawing, put
  // ~0.6s of the benchmark fixture's open inside the walk. The raw bytes
  // ride in a WeakMap keyed by the part object so every intermediate
  // collector (body walk, notes part, header/footer parts) forwards them
  // without signature changes.
  const media: ParsedDrawing["media"] = { mediaId: partPath, contentType, data: "" };
  pendingMediaBytes.set(media, bytes);
  return { image, media };
}

// See parseDrawing's return: raw media bytes awaiting their one deferred
// base64 pass. WeakMap so a MediaPart dropped mid-parse costs nothing.
const pendingMediaBytes = new WeakMap<MediaPart, Uint8Array>();

/** Encodes a parsed-but-not-yet-encoded MediaPart's bytes into `data`
 * (no-op for parts that already carry data — an already-encoded part from
 * another path). parseDocx runs this over the final media list. */
export async function fillPendingMediaData(part: MediaPart): Promise<void> {
  const bytes = pendingMediaBytes.get(part);
  if (bytes === undefined) {
    return;
  }
  pendingMediaBytes.delete(part);
  part.data = await bytesToBase64Async(bytes);
}

function escapeXmlAttr(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
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
  const descrAttr = image.altText ? ` descr="${escapeXmlAttr(image.altText)}"` : "";
  const docPr = `<wp:docPr id="${id}" name="${escapeXmlAttr(name)}"${descrAttr}/>`;
  const graphic =
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic xmlns:pic="${PIC_NAMESPACE}">` +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="${escapeXmlAttr(name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
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
  return (
    `<w:drawing xmlns:wp="${WP_NAMESPACE}">` +
    '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"' +
    ` relativeHeight="${251658240 + id}" behindDoc="${behindDoc}" locked="0"` +
    ' layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
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
