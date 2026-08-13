// A minimal OPC/ZIP writer: the counterpart to zip.ts's reader, enough to
// assemble a `.docx` or `.odt` package. Entries are DEFLATE-compressed via
// the platform's CompressionStream (per this project's no-new-dependencies
// constraint) unless marked `store`, sizes and CRC-32 are written into the
// local headers (no data descriptors), and there is no ZIP64 or encryption —
// document packages are always small, single-disk, unencrypted ZIPs.
import { yieldToMainThread } from "../idleScheduler";

const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const EOCD_SIGNATURE = 0x06054b50;

const VERSION_NEEDED = 20;
const COMPRESSION_STORE = 0;
const COMPRESSION_DEFLATE = 8;

/** A file to be added to the archive, in package order. */
export interface ZipEntryInput {
  /** Path within the archive, e.g. "word/document.xml". */
  name: string;
  data: Uint8Array;
  /** Write the entry uncompressed (ZIP STORE). ODF packages require this
   * for the `mimetype` entry so format sniffers can read it from the raw
   * bytes; everything else defaults to DEFLATE. */
  store?: boolean;
}

// Standard CRC-32 (IEEE 802.3, the polynomial ZIP mandates), table-driven.
const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32Range(crc: number, bytes: Uint8Array, start: number, end: number): number {
  for (let i = start; i < end; i++) {
    crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return crc;
}

export function crc32(bytes: Uint8Array): number {
  return (crc32Range(0xffffffff, bytes, 0, bytes.length) ^ 0xffffffff) >>> 0;
}

/** How long the entries loop below may run before it must yield (SPEC-047
 * AC-15's ≤100ms main-thread-block-during-save ceiling): a large fixture's
 * per-entry CRC-32 pass has no `await` of its own (unlike `deflateRaw`,
 * which already awaits a stream read per entry — but that await resolves
 * as a microtask, not a real yield back to the browser's input/render
 * queue, so a few hundred entries in a row still reads as one long task). */
const SAVE_YIELD_BUDGET_MS = 4;

// A single entry can itself be the long task: `word/document.xml` for a
// 300-page fixture is several MB on its own, and CRC-32's tight per-byte
// loop has nowhere to yield mid-entry — the entries-loop check above only
// ever fires *between* entries. Chunked in ~1MB slices with the same
// time-budgeted yield, so one huge entry reads as several short tasks
// instead of the one multi-hundred-ms task profiling this fixture's save
// actually found (crc32 was the single largest self-time contributor).
const CRC_CHUNK_BYTES = 1_000_000;

async function crc32Chunked(bytes: Uint8Array): Promise<number> {
  let crc = 0xffffffff;
  let offset = 0;
  let lastYield = performance.now();
  while (offset < bytes.length) {
    const end = Math.min(bytes.length, offset + CRC_CHUNK_BYTES);
    crc = crc32Range(crc, bytes, offset, end);
    offset = end;
    if (offset < bytes.length && performance.now() - lastYield > SAVE_YIELD_BUDGET_MS) {
      await yieldToMainThread();
      lastYield = performance.now();
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

async function deflateRaw(data: Uint8Array): Promise<Uint8Array> {
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(data);
      controller.close();
    },
  }).pipeThrough(new CompressionStream("deflate-raw"));

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

/** Append-only little-endian byte buffer the record builders write into. */
class ByteWriter {
  private chunks: Uint8Array[] = [];
  length = 0;

  u16(value: number): void {
    const bytes = new Uint8Array(2);
    new DataView(bytes.buffer).setUint16(0, value, true);
    this.bytes(bytes);
  }

  u32(value: number): void {
    const bytes = new Uint8Array(4);
    new DataView(bytes.buffer).setUint32(0, value >>> 0, true);
    this.bytes(bytes);
  }

  bytes(data: Uint8Array): void {
    this.chunks.push(data);
    this.length += data.length;
  }

  toUint8Array(): Uint8Array {
    const result = new Uint8Array(this.length);
    let position = 0;
    for (const chunk of this.chunks) {
      result.set(chunk, position);
      position += chunk.length;
    }
    return result;
  }
}

interface PreparedEntry {
  nameBytes: Uint8Array;
  compressed: Uint8Array;
  compressionMethod: number;
  crc: number;
  uncompressedSize: number;
  localHeaderOffset: number;
}

/** Assembles entries into a ZIP archive, preserving the given order (the
 * writer puts "[Content_Types].xml" first by convention). Entry data is
 * DEFLATE-compressed; last-modified timestamps are left at the DOS epoch —
 * nothing in OPC reads them, and a fixed value keeps output deterministic. */
export async function zip(entries: ZipEntryInput[]): Promise<Uint8Array> {
  const encoder = new TextEncoder();
  const out = new ByteWriter();
  const prepared: PreparedEntry[] = [];
  let lastYield = performance.now();

  for (const entry of entries) {
    if (performance.now() - lastYield > SAVE_YIELD_BUDGET_MS) {
      await yieldToMainThread();
      lastYield = performance.now();
    }
    const nameBytes = encoder.encode(entry.name);
    const compressed = entry.store ? entry.data : await deflateRaw(entry.data);
    const record: PreparedEntry = {
      nameBytes,
      compressed,
      compressionMethod: entry.store ? COMPRESSION_STORE : COMPRESSION_DEFLATE,
      crc: await crc32Chunked(entry.data),
      uncompressedSize: entry.data.length,
      localHeaderOffset: out.length,
    };
    prepared.push(record);

    out.u32(LOCAL_FILE_HEADER_SIGNATURE);
    out.u16(VERSION_NEEDED);
    out.u16(0); // general-purpose bit flag: no data descriptor, no UTF-8 flag
    out.u16(record.compressionMethod);
    out.u16(0); // last-modified time
    out.u16(0); // last-modified date
    out.u32(record.crc);
    out.u32(compressed.length);
    out.u32(record.uncompressedSize);
    out.u16(nameBytes.length);
    out.u16(0); // extra field length
    out.bytes(nameBytes);
    out.bytes(compressed);
  }

  const centralDirectoryOffset = out.length;
  for (const record of prepared) {
    out.u32(CENTRAL_DIRECTORY_SIGNATURE);
    out.u16(VERSION_NEEDED); // version made by
    out.u16(VERSION_NEEDED); // version needed to extract
    out.u16(0); // general-purpose bit flag
    out.u16(record.compressionMethod);
    out.u16(0); // last-modified time
    out.u16(0); // last-modified date
    out.u32(record.crc);
    out.u32(record.compressed.length);
    out.u32(record.uncompressedSize);
    out.u16(record.nameBytes.length);
    out.u16(0); // extra field length
    out.u16(0); // file comment length
    out.u16(0); // disk number start
    out.u16(0); // internal file attributes
    out.u32(0); // external file attributes
    out.u32(record.localHeaderOffset);
    out.bytes(record.nameBytes);
  }
  const centralDirectorySize = out.length - centralDirectoryOffset;

  out.u32(EOCD_SIGNATURE);
  out.u16(0); // disk number
  out.u16(0); // disk with central directory
  out.u16(prepared.length); // entries on this disk
  out.u16(prepared.length); // total entries
  out.u32(centralDirectorySize);
  out.u32(centralDirectoryOffset);
  out.u16(0); // comment length

  return out.toUint8Array();
}
