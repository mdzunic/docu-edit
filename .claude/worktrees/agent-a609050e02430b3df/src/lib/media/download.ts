// Download a media document — the media twin of lib/downloadDocument.ts and
// presentation/download.ts, kept framework-free. There is no writer: an
// unedited document's bytes that came in are the bytes that go out,
// byte-for-byte (round-trip fidelity is the project's first constraint, and
// for media it is trivial to honour absolutely). An edited image (SPEC-112)
// downloads its flattened render instead — so a redaction always leaves as
// pixels, never as a cover over a stored original.
import { ensureMediaEncoded, type MediaPart } from "../core/media";
import { base64ToBytes } from "../core/packageParts";
import { triggerBlobDownload } from "../core/triggerDownload";
import { currentImagePart, formatOfContentType, hasEdits } from "./editModel";
import { mediaFileExtension, type MediaModel } from "./model";

/** The source bytes of a media document, decoded from its stored part —
 * identical to the file that was opened. */
export function mediaSourceBytes(model: MediaModel): Uint8Array {
  const part = model.media.find((entry) => entry.mediaId === model.sourceMediaId);
  if (!part) {
    throw new Error(`Media document ${model.id} has no source part`);
  }
  ensureMediaEncoded(part);
  return base64ToBytes(part.data);
}

/** A part's bytes, encoding a still-deferred part first. */
export function mediaPartBytes(part: MediaPart): Uint8Array {
  ensureMediaEncoded(part);
  return base64ToBytes(part.data);
}

/** The file name a media document downloads under. */
export function mediaFileName(model: MediaModel): string {
  const title = model.title.trim();
  return `${title.length > 0 ? title : "Untitled"}.${mediaFileExtension(model.format)}`;
}

/** The current image's part — the render of an edited document, the source
 * of an unedited one. Throws for an edited document whose render is
 * missing: falling back to the source would hand out the pixels a
 * redaction covers. */
export function currentImagePartOrThrow(model: MediaModel): MediaPart {
  const part = currentImagePart(model);
  if (!part) {
    throw new Error(
      hasEdits(model)
        ? `Media document ${model.id} has edits but no saved render; open it and save again`
        : `Media document ${model.id} has no source part`,
    );
  }
  return part;
}

/** The current image's bytes and content type. */
export function currentImageBytes(model: MediaModel): { bytes: Uint8Array; contentType: string } {
  const part = currentImagePartOrThrow(model);
  return { bytes: mediaPartBytes(part), contentType: part.contentType };
}

/** The file name the current image downloads under: the source's name for
 * an unedited document, the render's format for an edited one. */
export function currentImageFileName(model: MediaModel): string {
  if (!hasEdits(model)) {
    return mediaFileName(model);
  }
  const part = currentImagePart(model);
  const format = part ? formatOfContentType(part.contentType) : model.format;
  return mediaFileName({ ...model, format });
}

/** Triggers a browser download of the current image: the verbatim source
 * bytes of an unedited document, the flattened render of an edited one. */
export function downloadMedia(model: MediaModel): void {
  const { bytes, contentType } = currentImageBytes(model);
  triggerBlobDownload(bytes, currentImageFileName(model), contentType);
}
