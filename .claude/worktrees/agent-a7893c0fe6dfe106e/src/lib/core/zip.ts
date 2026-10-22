// A minimal OPC/ZIP reader: enough to list and inflate the entries of a
// `.docx` package. No archive-writing, no ZIP64, no encryption — `.docx`
// packages are always small, single-disk, unencrypted ZIPs. Uses the
// platform's DecompressionStream for DEFLATE rather than a bundled inflate
// implementation, per this project's no-new-dependencies constraint.

import { yieldToMainThread } from "./idleScheduler";

export class ZipError extends Error {}

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;

const EOCD_MIN_SIZE = 22;
const MAX_COMMENT_LENGTH = 0xffff;

const COMPRESSION_STORE = 0;
const COMPRESSION_DEFLATE = 8;

// --- decompression budget ---------------------------------------------------
// DEFLATE reaches ~1000:1 on repetitive input, so the caller's limit on the
// *compressed* file (openDocument.ts's MAX_FILE_SIZE_BYTES) bounds nothing
// here: a 1MB archive of zeros inflates to 1GB and takes the tab with it.
// unzip() therefore inflates against a byte budget, enforced twice — once
// against the sizes the central directory declares (cheap, and rejects an
// honest bomb before a single byte is inflated) and again against the bytes
// actually produced, since a hostile archive is free to under-declare.
//
// The budget scales with the archive so a small package can still hold a
// legitimately text-heavy document, with a floor and a ceiling around it.
// For calibration: SPEC-047's 44MB benchmark package (507 parts, 500 images)
// inflates to 48MB — a ratio of 1.09, because a large package is media and
// media does not compress. Real packages sit nowhere near these numbers.
const INFLATED_BYTES_FLOOR = 64 * 1024 * 1024;
const INFLATED_BYTES_CEILING = 256 * 1024 * 1024;
const INFLATION_RATIO_LIMIT = 8;

/** The most an archive of `archiveBytes` may inflate to in total. */
function inflatedByteBudget(archiveBytes: number): number {
  return Math.min(
    INFLATED_BYTES_CEILING,
    Math.max(INFLATED_BYTES_FLOOR, archiveBytes * INFLATION_RATIO_LIMIT),
  );
}

interface CentralDirectoryEntry {
  fileName: string;
  compressionMethod: number;
  compressedSize: number;
  /** The size the central directory CLAIMS this entry inflates to. Used for
   * the cheap up-front rejection only — never trusted as a buffer size. */
  declaredSize: number;
  localHeaderOffset: number;
}

function findEndOfCentralDirectory(view: DataView): number {
  const maxSearch = Math.min(
    view.byteLength,
    EOCD_MIN_SIZE + MAX_COMMENT_LENGTH,
  );
  const searchStart = view.byteLength - maxSearch;
  for (let offset = view.byteLength - EOCD_MIN_SIZE; offset >= searchStart; offset--) {
    if (view.getUint32(offset, true) === EOCD_SIGNATURE) {
      return offset;
    }
  }
  throw new ZipError("Not a ZIP archive: end of central directory not found");
}

function readCentralDirectory(
  view: DataView,
  bytes: Uint8Array,
): CentralDirectoryEntry[] {
  const eocdOffset = findEndOfCentralDirectory(view);
  const totalEntries = view.getUint16(eocdOffset + 10, true);
  const centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);

  const decoder = new TextDecoder("utf-8");
  const entries: CentralDirectoryEntry[] = [];
  let offset = centralDirectoryOffset;

  for (let i = 0; i < totalEntries; i++) {
    if (offset + 46 > view.byteLength || view.getUint32(offset, true) !== CENTRAL_DIRECTORY_SIGNATURE) {
      throw new ZipError("Not a ZIP archive: malformed central directory");
    }
    const compressionMethod = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const declaredSize = view.getUint32(offset + 24, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraFieldLength = view.getUint16(offset + 30, true);
    const fileCommentLength = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);

    const nameStart = offset + 46;
    const fileName = decoder.decode(
      bytes.subarray(nameStart, nameStart + fileNameLength),
    );

    entries.push({ fileName, compressionMethod, compressedSize, declaredSize, localHeaderOffset });
    offset = nameStart + fileNameLength + extraFieldLength + fileCommentLength;
  }

  return entries;
}

/** Inflates one entry, giving up as soon as it has produced more than
 * `budget` bytes — the enforcement that does not trust the archive's own
 * declared sizes. The reader is cancelled on the way out so a bomb's
 * DecompressionStream stops rather than running on in the background. */
async function inflateRaw(data: Uint8Array, budget: number, fileName: string): Promise<Uint8Array> {
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(data);
      controller.close();
    },
  }).pipeThrough(new DecompressionStream("deflate-raw"));

  const chunks: Uint8Array[] = [];
  const reader = stream.getReader();
  let total = 0;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) {
        break;
      }
      if (value) {
        total += value.length;
        if (total > budget) {
          throw new ZipError(
            `Refusing to inflate ${fileName}: the archive expands well beyond its declared size`,
          );
        }
        chunks.push(value);
      }
    }
  } finally {
    await reader.cancel().catch(() => {});
  }

  const result = new Uint8Array(total);
  let position = 0;
  for (const chunk of chunks) {
    result.set(chunk, position);
    position += chunk.length;
  }
  return result;
}

function readEntryData(
  view: DataView,
  bytes: Uint8Array,
  entry: CentralDirectoryEntry,
): Uint8Array {
  const offset = entry.localHeaderOffset;
  if (offset + 30 > view.byteLength || view.getUint32(offset, true) !== LOCAL_FILE_HEADER_SIGNATURE) {
    throw new ZipError(`Not a ZIP archive: malformed local file header for ${entry.fileName}`);
  }
  const fileNameLength = view.getUint16(offset + 26, true);
  const extraFieldLength = view.getUint16(offset + 28, true);
  const dataStart = offset + 30 + fileNameLength + extraFieldLength;
  return bytes.subarray(dataStart, dataStart + entry.compressedSize);
}

/** Unzips a `.docx`/`.odt` (OPC) package into its decompressed parts, keyed
 * by their path within the archive (e.g. "word/document.xml"). Throws
 * `ZipError` if `data` is not a well-formed ZIP archive, or if inflating it
 * would exceed the decompression budget (see INFLATED_BYTES_FLOOR). */
export async function unzip(data: ArrayBuffer | Uint8Array): Promise<Map<string, Uint8Array>> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (bytes.byteLength < EOCD_MIN_SIZE) {
    throw new ZipError("Not a ZIP archive: file too small");
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  const entries = readCentralDirectory(view, bytes);

  // Pass one of the decompression budget (see INFLATED_BYTES_FLOOR above):
  // reject on what the central directory itself claims, before inflating
  // anything. A bomb built by a normal ZIP tool declares its real size and
  // dies here; one that lies is caught by the running total below.
  const budget = inflatedByteBudget(bytes.byteLength);
  let declaredTotal = 0;
  for (const entry of entries) {
    declaredTotal += entry.declaredSize;
    if (declaredTotal > budget) {
      throw new ZipError(
        "Refusing to unzip: the archive declares more decompressed content than we will open",
      );
    }
  }

  const parts = new Map<string, Uint8Array>();
  let inflatedTotal = 0;

  // `await inflateRaw(...)` resolves through microtasks only, so a package
  // with hundreds of entries (SPEC-047's fixture has 500 images) inflates
  // as ONE fused main-thread task without the periodic real yield below —
  // the largest single block of the whole open phase (AC-16).
  let sliceStart = performance.now();
  for (const entry of entries) {
    if (entry.fileName.endsWith("/")) {
      continue;
    }
    if (performance.now() - sliceStart > 8) {
      await yieldToMainThread();
      sliceStart = performance.now();
    }
    const compressed = readEntryData(view, bytes, entry);
    let content: Uint8Array;
    if (entry.compressionMethod === COMPRESSION_STORE) {
      content = compressed;
    } else if (entry.compressionMethod === COMPRESSION_DEFLATE) {
      // Pass two: what is left of the budget after every entry inflated so
      // far, so the ceiling holds across the package and not just per entry.
      content = await inflateRaw(compressed, budget - inflatedTotal, entry.fileName);
    } else {
      throw new ZipError(
        `Unsupported ZIP compression method ${entry.compressionMethod} for ${entry.fileName}`,
      );
    }
    inflatedTotal += content.length;
    parts.set(entry.fileName, content);
  }

  return parts;
}
