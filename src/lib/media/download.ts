// Download a media document — the media twin of lib/downloadDocument.ts and
// presentation/download.ts, kept framework-free. There is no writer: the
// bytes that came in are the bytes that go out, byte-for-byte (round-trip
// fidelity is the project's first constraint, and for media it is trivial to
// honour absolutely).
import { ensureMediaEncoded } from "../core/media";
import { base64ToBytes } from "../core/packageParts";
import { triggerBlobDownload } from "../core/triggerDownload";
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

/** The file name a media document downloads under. */
export function mediaFileName(model: MediaModel): string {
  const title = model.title.trim();
  return `${title.length > 0 ? title : "Untitled"}.${mediaFileExtension(model.format)}`;
}

/** Triggers a browser download of the verbatim source bytes. */
export function downloadMedia(model: MediaModel): void {
  triggerBlobDownload(mediaSourceBytes(model), mediaFileName(model), model.contentType);
}
