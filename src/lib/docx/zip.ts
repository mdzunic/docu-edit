// A minimal OPC/ZIP reader: enough to list and inflate the entries of a
// `.docx` package. No archive-writing, no ZIP64, no encryption — `.docx`
// packages are always small, single-disk, unencrypted ZIPs. Uses the
// platform's DecompressionStream for DEFLATE rather than a bundled inflate
// implementation, per this project's no-new-dependencies constraint.

export class ZipError extends Error {}

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;

const EOCD_MIN_SIZE = 22;
const MAX_COMMENT_LENGTH = 0xffff;

const COMPRESSION_STORE = 0;
const COMPRESSION_DEFLATE = 8;

interface CentralDirectoryEntry {
  fileName: string;
  compressionMethod: number;
  compressedSize: number;
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
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraFieldLength = view.getUint16(offset + 30, true);
    const fileCommentLength = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);

    const nameStart = offset + 46;
    const fileName = decoder.decode(
      bytes.subarray(nameStart, nameStart + fileNameLength),
    );

    entries.push({ fileName, compressionMethod, compressedSize, localHeaderOffset });
    offset = nameStart + fileNameLength + extraFieldLength + fileCommentLength;
  }

  return entries;
}

async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(data);
      controller.close();
    },
  }).pipeThrough(new DecompressionStream("deflate-raw"));

  const chunks: Uint8Array[] = [];
  const reader = stream.getReader();
  for (;;) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    if (value) {
      chunks.push(value);
    }
  }

  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
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
 * `ZipError` if `data` is not a well-formed ZIP archive. */
export async function unzip(data: ArrayBuffer | Uint8Array): Promise<Map<string, Uint8Array>> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (bytes.byteLength < EOCD_MIN_SIZE) {
    throw new ZipError("Not a ZIP archive: file too small");
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  const entries = readCentralDirectory(view, bytes);
  const parts = new Map<string, Uint8Array>();

  for (const entry of entries) {
    if (entry.fileName.endsWith("/")) {
      continue;
    }
    const compressed = readEntryData(view, bytes, entry);
    let content: Uint8Array;
    if (entry.compressionMethod === COMPRESSION_STORE) {
      content = compressed;
    } else if (entry.compressionMethod === COMPRESSION_DEFLATE) {
      content = await inflateRaw(compressed);
    } else {
      throw new ZipError(
        `Unsupported ZIP compression method ${entry.compressionMethod} for ${entry.fileName}`,
      );
    }
    parts.set(entry.fileName, content);
  }

  return parts;
}
