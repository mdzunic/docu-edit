// The edited media document (SPEC-112): which part is "the current image",
// what size it is, and the three model transitions the image editor writes
// — save, save as a copy, flatten. Framework-free and store-free per the
// media layer's boundary, and pure: every transition returns a NEW model
// sharing the untouched source MediaPart by reference, so no edit path can
// mutate the source (asserted against a deep-frozen model).
//
// "The current image" is the saved render when the document has edits and
// the source otherwise — what the viewer shows, the library thumbnails, a
// download writes and a conversion places. An unedited document still
// serves its source bytes verbatim (SPEC-107's round-trip rule).
import { normaliseStack, resolveEditGeometry, type ImageEdit, type PixelSize } from "../core/imageEdits";
import { bytesToBase64 } from "../core/packageParts";
import type { MediaPart } from "../core/media";
import { MEDIA_CONTENT_TYPES, type MediaFormat, type MediaModel } from "./model";

/** Rendered bytes and what they are. */
export interface RenderedImage {
  bytes: Uint8Array;
  contentType: string;
  widthPx: number;
  heightPx: number;
}

/** Typed pass-through for a MediaModel arriving through a `.vue` prop
 * (runtime-typed `Object`; espree bars the annotation in `.vue` scripts). */
export function asMediaModel(value: unknown): MediaModel {
  return value as MediaModel;
}

/** The part with `mediaId`, if the model holds one. */
export function findMediaPart(model: MediaModel, mediaId: string | undefined): MediaPart | undefined {
  return mediaId === undefined ? undefined : model.media.find((part) => part.mediaId === mediaId);
}

/** The saved stack, normalised (empty for an unedited document). */
export function savedStack(model: MediaModel): ImageEdit[] {
  return normaliseStack(model.edits ?? []);
}

/** Whether the document carries a saved edit stack. */
export function hasEdits(model: MediaModel): boolean {
  return savedStack(model).length > 0;
}

/** Whether the image editor can open the document: any image (a PDF page
 * needs a renderer this product does not have — SPEC-107). */
export function isEditableImage(model: MediaModel): boolean {
  return model.format !== "pdf";
}

/** The part the stack applies to: an SVG's raster base once it has one,
 * else the source. */
export function editBasePart(model: MediaModel): MediaPart | undefined {
  return findMediaPart(model, model.baseMediaId) ?? findMediaPart(model, model.sourceMediaId);
}

/** The current image's pixel size: the saved stack's output size over the
 * base's natural size, or the natural size itself. Null when unknown. */
export function currentImageSize(model: MediaModel): PixelSize | null {
  if (!model.image) {
    return null;
  }
  const stack = savedStack(model);
  if (stack.length === 0) {
    return { widthPx: model.image.widthPx, heightPx: model.image.heightPx };
  }
  const geometry = resolveEditGeometry(stack, model.image);
  return { widthPx: geometry.widthPx, heightPx: geometry.heightPx };
}

/** The part holding the current image: the render of an edited document,
 * the source of an unedited one. Undefined for an edited document whose
 * render is missing (the caller renders it, or degrades). */
export function currentImagePart(model: MediaModel): MediaPart | undefined {
  if (hasEdits(model)) {
    return findMediaPart(model, model.renderMediaId);
  }
  return findMediaPart(model, model.sourceMediaId);
}

/** The content type a render of this document is written as: JPEG for a
 * JPEG source (nothing in the stack adds alpha — a straighten never shows
 * its empty corners, redactions are opaque, annotations draw over opaque
 * pixels), PNG for everything else. */
export function renderContentType(model: MediaModel): "image/png" | "image/jpeg" {
  return model.contentType === "image/jpeg" ? "image/jpeg" : "image/png";
}

/** The media format a content type names (PNG for anything unknown). */
export function formatOfContentType(contentType: string): MediaFormat {
  const entry = (Object.entries(MEDIA_CONTENT_TYPES) as Array<[MediaFormat, string]>).find(
    ([, type]) => type === contentType,
  );
  return entry ? entry[0] : "png";
}

function partOf(mediaId: string, image: { bytes: Uint8Array; contentType: string }): MediaPart {
  return { mediaId, contentType: image.contentType, data: bytesToBase64(image.bytes) };
}

/** Builds a model's edited state — its stack, render part and (for an SVG)
 * raster base part — onto `base`, which keeps every other field. */
function withStack(
  base: MediaModel,
  source: MediaModel,
  stack: ImageEdit[],
  render: RenderedImage | null,
  raster: RenderedImage | null,
  newId: () => string,
): MediaModel {
  const next: MediaModel = { ...base };
  delete next.edits;
  delete next.renderMediaId;
  delete next.baseMediaId;
  const sourcePart = findMediaPart(source, source.sourceMediaId);
  const media: MediaPart[] = sourcePart ? [sourcePart] : [];
  const normalised = normaliseStack(stack);
  if (normalised.length > 0) {
    const existingBase = findMediaPart(source, source.baseMediaId);
    if (raster) {
      next.baseMediaId = newId();
      media.push(partOf(next.baseMediaId, raster));
      next.image = { widthPx: raster.widthPx, heightPx: raster.heightPx };
    } else if (existingBase) {
      next.baseMediaId = existingBase.mediaId;
      media.push(existingBase);
    }
    next.edits = normalised;
    if (render) {
      next.renderMediaId = newId();
      media.push(partOf(next.renderMediaId, render));
    }
  }
  next.media = media;
  return next;
}

/** Save: the stack stored on the record and the render refreshed (a fresh
 * part id, so nothing keyed by the old one serves stale pixels); the source
 * part is carried over by reference, its id and bytes untouched. An empty
 * stack (a reverted image) drops the render and base, leaving the source
 * as the current image again. `raster` is an SVG's base, made on its first
 * edit. Returns a new model. */
export function withSavedStack(
  model: MediaModel,
  stack: ImageEdit[],
  render: RenderedImage | null,
  raster: RenderedImage | null = null,
  newId: () => string = () => crypto.randomUUID(),
): MediaModel {
  return withStack(model, model, stack, render, raster, newId);
}

/** Save as a copy: a NEW media document (fresh id, the given title)
 * carrying the same source bytes and the same stack, with its own render.
 * The original model is not touched. */
export function copiedWithStack(
  model: MediaModel,
  stack: ImageEdit[],
  render: RenderedImage | null,
  raster: RenderedImage | null,
  title: string,
  newId: () => string = () => crypto.randomUUID(),
): MediaModel {
  const copy: MediaModel = { ...model, id: newId(), title };
  const sourcePart = findMediaPart(model, model.sourceMediaId);
  const source: MediaModel = {
    ...model,
    media: [
      ...(sourcePart ? [{ ...sourcePart }] : []),
      ...model.media.filter((part) => part.mediaId === model.baseMediaId).map((part) => ({ ...part })),
    ],
  };
  return withStack(copy, source, stack, render, raster, newId);
}

/** Flatten and discard original: the rendered bytes REPLACE the source
 * part's (same `sourceMediaId`, new bytes), the stack, render and base are
 * gone, and the format, size and byte length describe the new pixels. The
 * original bytes are no longer anywhere in the returned model. */
export function withFlattened(model: MediaModel, rendered: RenderedImage): MediaModel {
  const next: MediaModel = { ...model };
  delete next.edits;
  delete next.renderMediaId;
  delete next.baseMediaId;
  next.media = [partOf(model.sourceMediaId, rendered)];
  next.contentType = rendered.contentType;
  next.format = formatOfContentType(rendered.contentType);
  next.byteLength = rendered.bytes.length;
  next.image = { widthPx: rendered.widthPx, heightPx: rendered.heightPx };
  return next;
}
