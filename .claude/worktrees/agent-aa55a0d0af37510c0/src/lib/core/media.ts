// The product-neutral half of embedded media (SPEC-024): the media-part
// payload type, insert-time acceptance rules and the natural-size fitting
// math. Everything here is independent of the document model — a Slides
// product stores and validates the same media parts; only anchoring an
// Image block into the text flow is document-specific (lib/images.ts).
import { pxToCm } from "./units";
import { bytesToBase64 } from "./packageParts";
import { yieldToMainThread } from "./idleScheduler";

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

// Raw (not-yet-base64-encoded) bytes for a media part whose encode was
// deferred off the open path (SPEC-047): on a large `.docx` the base64 pass
// over ~50 MB of images is ~1s of main-thread work that would otherwise sit
// between choosing the file and the first page painting. When a reader
// defers, it registers the raw bytes here keyed by the part object and
// leaves `data` empty; every consumer that reads `data` in earnest
// (rendering, saving, persisting, thumbnailing, version hashing) calls
// `ensureMediaEncoded`/`flushPendingMedia` first, so the encode happens
// lazily, per part, at the moment the bytes are actually needed rather than
// all at once during open. A WeakMap so a part dropped before it is ever
// read costs nothing to forget.
const pendingMediaBytes = new WeakMap<MediaPart, Uint8Array>();

/** Records a media part's raw bytes as pending a deferred base64 encode
 * (SPEC-047). The part's `data` stays "" until a consumer forces the encode
 * through `ensureMediaEncoded`. */
export function registerPendingMedia(part: MediaPart, bytes: Uint8Array): void {
  pendingMediaBytes.set(part, bytes);
}

/** The pending (not-yet-encoded) byte length for a media part, or 0 when it
 * already carries its `data` — used to size the background encode's chunks
 * against the AC-16 main-thread-block budget. */
export function pendingMediaByteLength(part: MediaPart): number {
  return pendingMediaBytes.get(part)?.length ?? 0;
}

/** Encodes a still-deferred media part's bytes into `data` (a no-op once it
 * has been encoded, or for a part that never deferred). Idempotent and
 * synchronous: the one `btoa` for a single ~100 KB image is a short task,
 * unlike the whole-document pass this exists to break up. */
export function ensureMediaEncoded(part: MediaPart | undefined): void {
  if (!part) {
    return;
  }
  const bytes = pendingMediaBytes.get(part);
  if (bytes === undefined) {
    return;
  }
  pendingMediaBytes.delete(part);
  part.data = bytesToBase64(bytes);
}

/** Force-encodes every still-deferred part in a media list, synchronously.
 * Called by the writers, persist and version hashing right before they read
 * `data` in bulk, so a document saved or snapshotted before the background
 * pass (below) has finished never loses an image (round-trip fidelity is the
 * project's first constraint). */
export function flushPendingMedia(parts: readonly MediaPart[] | undefined): void {
  for (const part of parts ?? []) {
    ensureMediaEncoded(part);
  }
}

// One deferred encode chunk: ~9ms of `btoa` on the factory container at the
// documented ~18ms/MB. Deliberately small — this pass runs after first paint,
// overlapping the settle/scroll/typing phases, so each chunk is kept well
// under those phases' main-thread-block and typing-latency budgets (AC-15,
// AC-16) with room for a loaded host; the MessageChannel yield between chunks
// is near-free, so the extra chunks cost little.
const BACKGROUND_ENCODE_CHUNK_BYTES = 512 * 1024;

/** Encodes all still-deferred media in byte-bounded chunks with a real
 * macrotask yield between them (SPEC-047 AC-16), so the deferred pass reads
 * as many short tasks rather than one long block. The interactive open path
 * kicks this off after first paint; it races harmlessly with the per-part
 * force-encodes above (both re-check WeakMap membership with no await between
 * the read and the encode, so a part is never encoded twice or lost). */
export async function encodePendingMediaInBackground(
  parts: readonly MediaPart[] | undefined,
  chunkBytes: number = BACKGROUND_ENCODE_CHUNK_BYTES,
): Promise<void> {
  const list = parts ?? [];
  let pending = 0;
  for (let i = 0; i < list.length; i++) {
    const byteLength = pendingMediaByteLength(list[i]);
    if (byteLength === 0) {
      continue;
    }
    ensureMediaEncoded(list[i]);
    pending += byteLength;
    if (pending >= chunkBytes && i < list.length - 1) {
      pending = 0;
      await yieldToMainThread();
    }
  }
}

/** A media part's bytes as a `data:` URI for the canvas `<img>` — resolved
 * only for rendering; AC-27 bars base64 data URIs from the saved XML, which
 * this never touches. "" for an unresolved/absent part. Force-encodes a
 * still-deferred part on demand (SPEC-047), so a page's images paint the
 * moment it is realised without waiting for the whole-document encode. */
export function mediaDataUri(media: MediaPart | undefined): string {
  if (!media) {
    return "";
  }
  ensureMediaEncoded(media);
  return `data:${media.contentType};base64,${media.data}`;
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

/** The image MIME types the LIBRARY holds as media documents (SPEC-107):
 * SPEC-024's four insert-into-a-document types plus WebP. Deliberately a
 * separate list — ACCEPTED_IMAGE_MIME_TYPES stays the insert list SPEC-024
 * owns, and this spec does not change it. */
export const LIBRARY_IMAGE_MIME_TYPES = [
  ...ACCEPTED_IMAGE_MIME_TYPES,
  "image/webp",
] as const;

/** Whether a file's declared MIME type is one "Insert image" accepts. */
export function isAcceptedImageType(mimeType: string): boolean {
  return (ACCEPTED_IMAGE_MIME_TYPES as readonly string[]).includes(mimeType);
}

// --- audio/video media (SPEC-095) ------------------------------------------

/** The audio/video container MIME types "Insert media" accepts (SPEC-095),
 * as both a type list and an `<input accept>` string. Deliberately the
 * broadly-supported containers: we play exactly what a plain
 * `<video>`/`<audio>` element plays and never transcode, so the browser's own
 * `canPlayType` (core/mediaPlayback.ts) still decides playability per codec at
 * render time — a container we accept but the browser cannot decode degrades
 * to a labelled poster, never an error. */
export const ACCEPTED_MEDIA_MIME_TYPES = [
  "video/mp4",
  "video/webm",
  "video/ogg",
  "audio/mpeg",
  "audio/mp4",
  "audio/ogg",
  "audio/wav",
] as const;

export const MEDIA_ACCEPT_ATTR = ACCEPTED_MEDIA_MIME_TYPES.join(",");

/** Whether a file's declared MIME type is one "Insert media" accepts. */
export function isAcceptedMediaType(mimeType: string): boolean {
  return (ACCEPTED_MEDIA_MIME_TYPES as readonly string[]).includes(mimeType);
}

/** Media inserted over this is rejected (SPEC-095): a clip persists whole in a
 * single IndexedDB record, so the ceiling keeps one insert from bloating the
 * store, and sits well under `openDocument`'s 50 MB whole-file limit — an
 * over-ceiling clip could therefore only ever arrive already inside an opened
 * deck (where it is preserved with `oversize: true`, never rejected). */
export const MAX_MEDIA_SIZE_BYTES = 25 * 1024 * 1024;

/** Shown when a rejected clip is too large (SPEC-095). */
export const MEDIA_TOO_LARGE_MESSAGE = "That media file is too large (limit 25 MB)";

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
  // WebP (a picker-accepted image, SPEC-082) and the audio/video containers
  // SPEC-095 carries — so a media part path types correctly and never gets a
  // `.png` extension. `.ogg` is the audio spelling (video is rare and shares
  // the extension); the video↔audio split for a clip is decided by its
  // relationship type (pptx) / draw:mime-type (odp), not the extension.
  webp: "image/webp",
  mp4: "video/mp4",
  webm: "video/webm",
  ogg: "audio/ogg",
  m4a: "audio/mp4",
  mp3: "audio/mpeg",
  wav: "audio/wav",
};

/** The file extension a media content type is written with, the inverse of
 * CONTENT_TYPE_BY_EXTENSION's lookup — "jpeg" (not "jpg") is the canonical
 * spelling the writers pick for JPEG parts. */
const EXTENSION_BY_CONTENT_TYPE: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpeg",
  "image/gif": "gif",
  "image/svg+xml": "svg",
  // WebP (SPEC-082) and the audio/video containers SPEC-095 carries.
  "image/webp": "webp",
  "video/mp4": "mp4",
  "video/webm": "webm",
  "video/ogg": "ogg",
  "audio/mpeg": "mp3",
  "audio/mp4": "m4a",
  "audio/ogg": "ogg",
  "audio/wav": "wav",
};

/** The MIME type a media part path's extension implies (case-insensitive),
 * or undefined for an extension no accepted image/media format uses. */
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
