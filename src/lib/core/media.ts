// The product-neutral half of embedded media (SPEC-024): the media-part
// payload type, insert-time acceptance rules and the natural-size fitting
// math. Everything here is independent of the document model — a Slides
// product stores and validates the same media parts; only anchoring an
// Image block into the text flow is document-specific (lib/images.ts).
import { pxToCm } from "./units";

/** A binary image payload stored once and referenced by an image block's
 * `mediaId` — so an image used twice (AC-32) shares one entry, and deleting
 * one use does not by itself remove it (see lib/images.ts's
 * reference-counted pruning, AC-31). `mediaId` is deliberately not named
 * `id`: lib/equality.ts's IGNORED_KEYS drops any bare "id" key from
 * structural comparison, which would let a wrong mediaId reference pass a
 * round-trip test unnoticed. */
export interface MediaPart {
  mediaId: string;
  /** MIME type — "image/png", "image/jpeg", "image/gif", or "image/svg+xml"
   * for a `.odt` (never `.docx`: SVG is converted to PNG on insert, AC-2). */
  contentType: string;
  /** The image bytes, base64-encoded so the model stays JSON-serialisable
   * (see opaque.ts's OpaquePart, the same convention). */
  data: string;
}

/** A media part's bytes as a `data:` URI for the canvas `<img>` — resolved
 * only for rendering; AC-27 bars base64 data URIs from the saved XML, which
 * this never touches. "" for an unresolved/absent part. */
export function mediaDataUri(media: MediaPart | undefined): string {
  return media ? `data:${media.contentType};base64,${media.data}` : "";
}

/** Images larger than this are rejected at insertion time (AC-33). */
export const MAX_IMAGE_SIZE_BYTES = 10 * 1024 * 1024;

/** Shown when a rejected image is too large (AC-33, AC-34). */
export const IMAGE_TOO_LARGE_MESSAGE = "That image is too large (limit 10 MB)";

/** File types "Insert image" accepts, as both MIME types and an `<input
 * accept>` string (AC-1). */
export const ACCEPTED_IMAGE_MIME_TYPES = [
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/svg+xml",
] as const;

export const ACCEPTED_IMAGE_ACCEPT_ATTR = ACCEPTED_IMAGE_MIME_TYPES.join(",");

/** Whether a file's declared MIME type is one "Insert image" accepts. */
export function isAcceptedImageType(mimeType: string): boolean {
  return (ACCEPTED_IMAGE_MIME_TYPES as readonly string[]).includes(mimeType);
}

/** Whether a candidate image's byte length exceeds the 10 MB limit (AC-33). */
export function isImageTooLarge(byteLength: number): boolean {
  return byteLength > MAX_IMAGE_SIZE_BYTES;
}

/** The display size a freshly inserted image gets (AC-5, AC-6): its natural
 * size, scaled down to fit `textAreaCm` when it would otherwise overflow —
 * aspect ratio always preserved. */
export function fitImageSizeCm(
  naturalWidthPx: number,
  naturalHeightPx: number,
  textAreaCm: number,
): { widthCm: number; heightCm: number } {
  const naturalWidthCm = pxToCm(naturalWidthPx);
  const naturalHeightCm = pxToCm(naturalHeightPx);
  if (naturalWidthCm <= textAreaCm || naturalWidthCm <= 0) {
    return { widthCm: naturalWidthCm, heightCm: naturalHeightCm };
  }
  const scale = textAreaCm / naturalWidthCm;
  return { widthCm: textAreaCm, heightCm: naturalHeightCm * scale };
}

/** The smallest width/height a resize (handle drag or the options panel) can
 * leave an image — no AC specifies one, so this just keeps the image
 * visible/selectable rather than collapsing to nothing. */
export const MIN_IMAGE_SIZE_CM = 0.2;

/** Extension → MIME type for the image formats the packages carry. Shared
 * by both format readers (typing the bytes behind a `word/media/*` or
 * `Pictures/*` part path) and docx/writer.ts's `[Content_Types].xml`
 * Default entries. */
export const CONTENT_TYPE_BY_EXTENSION: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  svg: "image/svg+xml",
};

/** The file extension a media content type is written with, the inverse of
 * CONTENT_TYPE_BY_EXTENSION's lookup — "jpeg" (not "jpg") is the canonical
 * spelling the writers pick for JPEG parts. */
const EXTENSION_BY_CONTENT_TYPE: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpeg",
  "image/gif": "gif",
  "image/svg+xml": "svg",
};

/** The MIME type a media part path's extension implies (case-insensitive),
 * or undefined for an extension no accepted image format uses. */
export function contentTypeForPartName(partName: string): string | undefined {
  const dot = partName.lastIndexOf(".");
  if (dot < 0) {
    return undefined;
  }
  return CONTENT_TYPE_BY_EXTENSION[partName.slice(dot + 1).toLowerCase()];
}

/** The extension a media part is written under for its content type —
 * falls back to "png" for a type outside the accepted set. */
export function extensionForContentType(contentType: string): string {
  return EXTENSION_BY_CONTENT_TYPE[contentType] ?? "png";
}
