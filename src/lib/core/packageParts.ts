// Opaque package parts (SPEC-008), product-neutral: every package format
// this repo opens is a ZIP of named parts, and every product's model keeps
// the parts it does not model as verbatim bytes so a same-format save can
// write them back unchanged. The base64 codecs live here too — models stay
// JSON-serialisable, so part bytes are not held as raw Uint8Arrays.
import { yieldToMainThread } from "./idleScheduler";

/** A package part the reader does not model (SPEC-008) — `customXml/`,
 * charts, themes, embedded objects, images, `word/webSettings.xml`, … —
 * retained for the lifetime of the document and written back byte-identical
 * on a save to the same format. */
export interface OpaquePart {
  /** Path within the source package, e.g. "word/charts/chart1.xml". */
  name: string;
  /** The part's original (decompressed) bytes, base64-encoded so the model
   * stays JSON-serialisable (see the header comment). */
  data: string;
}

// Chunked so String.fromCharCode never exceeds the argument-count limit.
const BASE64_CHUNK = 0x8000;

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += BASE64_CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(i, i + BASE64_CHUNK));
  }
  return btoa(binary);
}

/** The async twin of `bytesToBase64` (SPEC-047 AC-3/AC-16): encodes through
 * `FileReader.readAsDataURL`, the browser's native base64 encoder; the output
 * is byte-identical standard base64. The native encode itself is quick, but
 * every call costs one async round trip, and over the hundreds of small parts
 * a large package carries that scheduling dominates: on the 46MB benchmark
 * fixture a per-part FileReader pass measured ~2x the wall clock of a
 * synchronous `bytesToBase64` pass, so the docx reader defers its encodes and
 * runs them synchronously in byte-bounded chunks instead — see
 * `ensureMediaEncoded` (core/media.ts) and the deferred-encode pass in
 * docx/reader.ts. Suited to keeping an occasional large encode off the main
 * thread, not to per-part loops; of the readers only pptx/odp still call it
 * that way. Falls back to the sync path where FileReader/Blob are missing. */
export function bytesToBase64Async(bytes: Uint8Array): Promise<string> {
  if (typeof FileReader !== "function" || typeof Blob !== "function") {
    return Promise.resolve(bytesToBase64(bytes));
  }
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => {
      const url = reader.result as string;
      resolve(url.slice(url.indexOf(",") + 1));
    });
    reader.addEventListener("error", () => reject(reader.error ?? new Error("base64 encode failed")));
    reader.readAsDataURL(new Blob([bytes as BlobPart]));
  });
}

export function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/** A list of opaque parts as name -> decoded bytes, in list order. */
export function decodedPartBytes(opaque: readonly OpaquePart[] | undefined): Map<string, Uint8Array> {
  const parts = new Map<string, Uint8Array>();
  for (const part of opaque ?? []) {
    parts.set(part.name, base64ToBytes(part.data));
  }
  return parts;
}

/** The decoded byte length of a well-formed base64 string (standard
 * padding, length a multiple of 4 — true of everything this repo encodes
 * with `bytesToBase64`/`btoa`): used to preallocate `base64ToBytesChunked`'s
 * output once, rather than growing it. */
function base64DecodedLength(base64: string): number {
  let padding = 0;
  if (base64.endsWith("==")) {
    padding = 2;
  } else if (base64.endsWith("=")) {
    padding = 1;
  }
  return Math.floor(base64.length / 4) * 3 - padding;
}

/** The async, yielding twin of `base64ToBytes` (SPEC-137 AC-46): decodes at
 * most `maxSliceChars` of base64 per `atob` call — rounded down to a
 * multiple of 4 so a slice boundary never splits a base64 quartet (padding,
 * when present, only ever falls in the base64 string's own final quartet,
 * so every slice but the one containing it decodes cleanly on its own) —
 * yielding to the main thread (`yieldToMainThread`) between slices, and
 * writing each slice's bytes into one preallocated output array. Decodes to
 * exactly the same bytes as `base64ToBytes(base64)`; a part small enough to
 * fit in one slice decodes synchronously, with no yield at all. */
export async function base64ToBytesChunked(base64: string, maxSliceChars = 4 * 1024 * 1024): Promise<Uint8Array> {
  const sliceChars = Math.max(4, Math.floor(maxSliceChars / 4) * 4);
  if (base64.length <= sliceChars) {
    return base64ToBytes(base64);
  }
  const out = new Uint8Array(base64DecodedLength(base64));
  let offset = 0;
  for (let start = 0; start < base64.length; start += sliceChars) {
    const end = Math.min(start + sliceChars, base64.length);
    const binary = atob(base64.slice(start, end));
    for (let i = 0; i < binary.length; i++) {
      out[offset + i] = binary.charCodeAt(i);
    }
    offset += binary.length;
    if (end < base64.length) {
      await yieldToMainThread();
    }
  }
  return out;
}

/** The size, in base64 characters, at or above which `decodedPartBytesAsync`
 * decodes a part through `base64ToBytesChunked` rather than the synchronous
 * `base64ToBytes`: small parts (a few KB of XML, an icon) cost nothing to
 * decode in one `atob` call; a retained worksheet part's base64 can run to
 * tens of MB on a large workbook (SPEC-137's benchmark fixture), long enough
 * that slicing it is worth the (tiny) overhead of the slices themselves. */
const CHUNKED_DECODE_THRESHOLD_CHARS = 1024 * 1024;

/** The async twin of `decodedPartBytes` (SPEC-137's save path): the same
 * map, but a part whose base64 is at least `CHUNKED_DECODE_THRESHOLD_CHARS`
 * long is decoded through `base64ToBytesChunked`, yielding mid-decode,
 * instead of one synchronous `atob` over the whole part. */
export async function decodedPartBytesAsync(opaque: readonly OpaquePart[] | undefined): Promise<Map<string, Uint8Array>> {
  const parts = new Map<string, Uint8Array>();
  for (const part of opaque ?? []) {
    const bytes = part.data.length >= CHUNKED_DECODE_THRESHOLD_CHARS ? await base64ToBytesChunked(part.data) : base64ToBytes(part.data);
    parts.set(part.name, bytes);
  }
  return parts;
}
