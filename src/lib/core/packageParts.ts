// Opaque package parts (SPEC-008), product-neutral: every package format
// this repo opens is a ZIP of named parts, and every product's model keeps
// the parts it does not model as verbatim bytes so a same-format save can
// write them back unchanged. The base64 codecs live here too — models stay
// JSON-serialisable, so part bytes are not held as raw Uint8Arrays.

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
 * `fillPendingMediaDataSync` (docx/drawing.ts) and the deferred-encode pass in
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
