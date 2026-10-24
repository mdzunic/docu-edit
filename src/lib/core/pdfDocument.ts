// A STRUCTURAL PDF reader (SPEC-107): turns bytes into an object graph and a
// page list, and stops there. No content stream is interpreted and no glyph
// is decoded — rasterising a PDF is explicitly out of this product's scope,
// so everything here is what can be learned about a file without rendering
// it: header version, the cross-reference table (classic tables with /Prev
// chains, hybrid /XRefStm, and PDF 1.5 xref streams), object streams, the
// page tree with attribute inheritance, and the /Info dictionary.
//
// Async because inflating an xref or object stream goes through the
// platform's DecompressionStream — the same trade core/zip.ts makes.
//
// Bounds are part of the contract, not defensive extras: a cyclic /Prev
// chain, a self-referential page tree or a decompression bomb terminates
// with an error rather than a hang, and each bound has a test that trips it
// (pdfDocument.spec.ts).
//
// Generations: the object map is keyed by object number alone. The
// generation an xref entry gives is honoured when the object is read (its
// `N G obj` header must carry it, and the newest xref section's entry wins),
// but it is never a second map key. A file that genuinely needs
// generation-keyed lookup — a reference whose generation differs from the
// live object's — is refused with a PdfParseError rather than half-read.
//
// The file's second half is the WRITER (SPEC-111): assemblePdf copies pages
// of parsed documents into one new file — see "the writer half" below.
import { yieldToMainThread } from "./idleScheduler";
import { PDF_PRODUCER } from "./pdfBuild";

/** Hard ceilings on what a parse will touch. */
export const MAX_PDF_OBJECTS = 500_000;
export const MAX_XREF_CHAIN = 64;
export const MAX_RESOLVE_DEPTH = 64;
export const MAX_INFLATED_STREAM_BYTES = 256 * 1024 * 1024;

/** A structural parse failure, carrying the byte offset it happened at. */
export class PdfParseError extends Error {
  readonly offset: number;

  constructor(message: string, offset: number) {
    super(message);
    this.offset = offset;
  }
}

/** An encrypted file — refused, never guessed at; decrypting is out of
 * scope. */
export class PdfEncryptedError extends PdfParseError {
  constructor(offset: number) {
    super("This PDF is password-protected", offset);
  }
}

// Two more refusals that are final: neither is an unusable xref, so the
// one reconstruction retry would only meet them again.
/** A reference whose generation is not the live object's (see the header). */
class PdfGenerationError extends PdfParseError {}
/** A stream that inflates past MAX_INFLATED_STREAM_BYTES. */
class PdfInflateLimitError extends PdfParseError {}

/** One parsed PDF value. Strings stay raw bytes (their encoding is
 * context-dependent — see `pdfTextOf`); a stream keeps its dictionary plus
 * the raw (still-encoded) data, decoded on demand by `streamData`. */
export type PdfValue =
  | { t: "null" }
  | { t: "bool"; v: boolean }
  | { t: "num"; v: number }
  | { t: "str"; v: Uint8Array }
  | { t: "name"; v: string }
  | { t: "array"; v: PdfValue[] }
  | { t: "dict"; v: Map<string, PdfValue> }
  | { t: "stream"; dict: Map<string, PdfValue>; raw: Uint8Array }
  | { t: "ref"; num: number; gen: number };

const PDF_NULL: PdfValue = { t: "null" };

/** Where each loaded object's `N G obj` header sits in the file, per parsed
 * document — kept beside the pinned PdfDocument shape (not on it) so errors
 * raised after the parse (resolve, streamData) still report a real byte
 * offset. Object-stream members map to their container's offset. */
const objectOffsets = new WeakMap<PdfDocument, Map<number, number>>();

function offsetOfObject(doc: PdfDocument, num: number): number {
  return objectOffsets.get(doc)?.get(num) ?? 0;
}

/** The file offset of a stream's raw data, when it is a view into the
 * document's own bytes (every stream `parsePdf` reads is); 0 otherwise. */
function offsetOfStream(doc: PdfDocument, raw: Uint8Array): number {
  const start = raw.byteOffset - doc.bytes.byteOffset;
  return raw.buffer === doc.bytes.buffer && start >= 0 && start <= doc.bytes.length ? start : 0;
}

/** One leaf of the page tree, with its inherited attributes resolved. */
export interface PdfPageRef {
  ref: { num: number; gen: number };
  /** The effective CropBox (fallback MediaBox, fallback US Letter),
   * normalised so width/height are positive. */
  widthPt: number;
  heightPt: number;
  rotate: 0 | 90 | 180 | 270;
  mediaBoxPt: [number, number, number, number];
  cropBoxPt?: [number, number, number, number];
  /** Resources/MediaBox/CropBox/Rotate as resolved through inheritance. */
  inherited: Map<string, PdfValue>;
}

export interface PdfDocument {
  /** Header version, e.g. "1.7". */
  version: string;
  /** The source bytes, retained for byte-preserving copy. */
  bytes: Uint8Array;
  /** By object number. Real-world non-incremental PDFs use generation 0 for
   * every live object; a non-zero generation in an xref entry is honoured
   * when reading (the newest xref section wins) but is never a second map
   * key. */
  objects: Map<number, PdfValue>;
  trailer: Map<string, PdfValue>;
  pages: PdfPageRef[];
  info?: Map<string, PdfValue>;
  linearised: boolean;
  /** True when the xref was unusable and the object table was rebuilt by
   * scanning the whole buffer for `N G obj` headers. */
  reconstructed: boolean;
}

// --- byte-level helpers ------------------------------------------------------

interface Cursor {
  bytes: Uint8Array;
  pos: number;
}

function isWhitespace(byte: number): boolean {
  return byte === 0 || byte === 9 || byte === 10 || byte === 12 || byte === 13 || byte === 32;
}

function isDelimiter(byte: number): boolean {
  return (
    byte === 0x28 || byte === 0x29 || byte === 0x3c || byte === 0x3e || byte === 0x5b ||
    byte === 0x5d || byte === 0x7b || byte === 0x7d || byte === 0x2f || byte === 0x25
  );
}

function isRegular(byte: number | undefined): boolean {
  return byte !== undefined && !isWhitespace(byte) && !isDelimiter(byte);
}

function isDigit(byte: number | undefined): boolean {
  return byte !== undefined && byte >= 0x30 && byte <= 0x39;
}

/** Skips whitespace and `%` comments (which run to end of line). */
function skipWs(cur: Cursor): void {
  const { bytes } = cur;
  for (;;) {
    while (cur.pos < bytes.length && isWhitespace(bytes[cur.pos])) {
      cur.pos += 1;
    }
    if (bytes[cur.pos] === 0x25) {
      while (cur.pos < bytes.length && bytes[cur.pos] !== 10 && bytes[cur.pos] !== 13) {
        cur.pos += 1;
      }
      continue;
    }
    return;
  }
}

/** Whether the keyword starts at `pos` and ends on a token boundary. */
function keywordAt(bytes: Uint8Array, pos: number, keyword: string): boolean {
  for (let i = 0; i < keyword.length; i += 1) {
    if (bytes[pos + i] !== keyword.charCodeAt(i)) {
      return false;
    }
  }
  return !isRegular(bytes[pos + keyword.length]);
}

function expectKeyword(cur: Cursor, keyword: string): void {
  skipWs(cur);
  if (!keywordAt(cur.bytes, cur.pos, keyword)) {
    throw new PdfParseError(`expected "${keyword}"`, cur.pos);
  }
  cur.pos += keyword.length;
}

/** Parses a plain (unsigned) integer token; null when none is present. */
function tryParseUint(cur: Cursor): number | null {
  skipWs(cur);
  const start = cur.pos;
  let value = 0;
  while (isDigit(cur.bytes[cur.pos])) {
    value = value * 10 + (cur.bytes[cur.pos] - 0x30);
    cur.pos += 1;
    if (cur.pos - start > 10) {
      throw new PdfParseError("integer too long", start);
    }
  }
  return cur.pos > start ? value : null;
}

function parseUint(cur: Cursor): number {
  const value = tryParseUint(cur);
  if (value === null) {
    throw new PdfParseError("expected an integer", cur.pos);
  }
  return value;
}

const asciiDecoder = new TextDecoder("latin1");

function latin1(bytes: Uint8Array): string {
  return asciiDecoder.decode(bytes);
}

// --- value parsing -----------------------------------------------------------

const MAX_NESTING = 256;

function parseName(cur: Cursor): PdfValue {
  cur.pos += 1; // "/"
  let name = "";
  const { bytes } = cur;
  while (isRegular(bytes[cur.pos])) {
    let byte = bytes[cur.pos];
    if (byte === 0x23 && cur.pos + 2 < bytes.length) {
      const hex = latin1(bytes.subarray(cur.pos + 1, cur.pos + 3));
      const code = Number.parseInt(hex, 16);
      if (Number.isFinite(code)) {
        byte = code;
        cur.pos += 2;
      }
    }
    name += String.fromCharCode(byte);
    cur.pos += 1;
    if (name.length > 512) {
      throw new PdfParseError("name too long", cur.pos);
    }
  }
  return { t: "name", v: name };
}

function parseLiteralString(cur: Cursor): PdfValue {
  const start = cur.pos;
  cur.pos += 1; // "("
  const out: number[] = [];
  let depth = 1;
  const { bytes } = cur;
  while (cur.pos < bytes.length) {
    const byte = bytes[cur.pos];
    cur.pos += 1;
    if (byte === 0x5c) {
      const escaped = bytes[cur.pos];
      cur.pos += 1;
      switch (escaped) {
        case 0x6e: out.push(10); break; // \n
        case 0x72: out.push(13); break; // \r
        case 0x74: out.push(9); break; // \t
        case 0x62: out.push(8); break; // \b
        case 0x66: out.push(12); break; // \f
        case 0x28: out.push(0x28); break;
        case 0x29: out.push(0x29); break;
        case 0x5c: out.push(0x5c); break;
        case 13: // line continuation (\<CR>[<LF>])
          if (bytes[cur.pos] === 10) {
            cur.pos += 1;
          }
          break;
        case 10:
          break;
        default:
          if (isDigit(escaped)) {
            let code = escaped - 0x30;
            for (let i = 0; i < 2 && isDigit(bytes[cur.pos]); i += 1) {
              code = code * 8 + (bytes[cur.pos] - 0x30);
              cur.pos += 1;
            }
            out.push(code & 0xff);
          } else if (escaped !== undefined) {
            out.push(escaped);
          }
      }
      continue;
    }
    if (byte === 0x28) {
      depth += 1;
    } else if (byte === 0x29) {
      depth -= 1;
      if (depth === 0) {
        return { t: "str", v: new Uint8Array(out) };
      }
    }
    out.push(byte);
  }
  throw new PdfParseError("unterminated string", start);
}

function parseHexString(cur: Cursor): PdfValue {
  const start = cur.pos;
  cur.pos += 1; // "<"
  let hex = "";
  const { bytes } = cur;
  while (cur.pos < bytes.length && bytes[cur.pos] !== 0x3e) {
    const byte = bytes[cur.pos];
    if (!isWhitespace(byte)) {
      hex += String.fromCharCode(byte);
    }
    cur.pos += 1;
  }
  if (bytes[cur.pos] !== 0x3e) {
    throw new PdfParseError("unterminated hex string", start);
  }
  cur.pos += 1;
  if (hex.length % 2 === 1) {
    hex += "0";
  }
  const out = new Uint8Array(hex.length / 2);
  for (let i = 0; i < out.length; i += 1) {
    const code = Number.parseInt(hex.slice(i * 2, i * 2 + 2), 16);
    if (Number.isNaN(code)) {
      throw new PdfParseError("bad hex string", start);
    }
    out[i] = code;
  }
  return { t: "str", v: out };
}

function parseNumberOrRef(cur: Cursor): PdfValue {
  const start = cur.pos;
  const { bytes } = cur;
  let text = "";
  while (isRegular(bytes[cur.pos]) && /[\d+.eE-]/.test(String.fromCharCode(bytes[cur.pos]))) {
    text += String.fromCharCode(bytes[cur.pos]);
    cur.pos += 1;
    if (text.length > 40) {
      throw new PdfParseError("number too long", start);
    }
  }
  const value = Number(text);
  if (text.length === 0 || !Number.isFinite(value)) {
    throw new PdfParseError(`bad number "${text}"`, start);
  }
  // "N G R" lookahead: two non-negative integers followed by the R keyword
  // form an indirect reference.
  if (Number.isInteger(value) && value >= 0 && /^\d+$/.test(text)) {
    const save = cur.pos;
    const gen = tryParseUint(cur);
    if (gen !== null) {
      skipWs(cur);
      if (keywordAt(bytes, cur.pos, "R")) {
        cur.pos += 1;
        return { t: "ref", num: value, gen };
      }
    }
    cur.pos = save;
  }
  return { t: "num", v: value };
}

function parseValue(cur: Cursor, depth: number): PdfValue {
  if (depth > MAX_NESTING) {
    throw new PdfParseError("value nesting too deep", cur.pos);
  }
  skipWs(cur);
  const { bytes } = cur;
  const byte = bytes[cur.pos];
  if (byte === undefined) {
    throw new PdfParseError("unexpected end of file", cur.pos);
  }
  if (byte === 0x3c) {
    if (bytes[cur.pos + 1] === 0x3c) {
      cur.pos += 2;
      const dict = new Map<string, PdfValue>();
      for (;;) {
        skipWs(cur);
        if (bytes[cur.pos] === 0x3e && bytes[cur.pos + 1] === 0x3e) {
          cur.pos += 2;
          return { t: "dict", v: dict };
        }
        if (bytes[cur.pos] !== 0x2f) {
          throw new PdfParseError("expected a name key", cur.pos);
        }
        const key = parseName(cur);
        dict.set((key as { v: string }).v, parseValue(cur, depth + 1));
      }
    }
    return parseHexString(cur);
  }
  if (byte === 0x28) {
    return parseLiteralString(cur);
  }
  if (byte === 0x2f) {
    return parseName(cur);
  }
  if (byte === 0x5b) {
    cur.pos += 1;
    const items: PdfValue[] = [];
    for (;;) {
      skipWs(cur);
      if (bytes[cur.pos] === 0x5d) {
        cur.pos += 1;
        return { t: "array", v: items };
      }
      items.push(parseValue(cur, depth + 1));
      if (items.length > MAX_PDF_OBJECTS) {
        throw new PdfParseError("array too long", cur.pos);
      }
    }
  }
  if (keywordAt(bytes, cur.pos, "true")) {
    cur.pos += 4;
    return { t: "bool", v: true };
  }
  if (keywordAt(bytes, cur.pos, "false")) {
    cur.pos += 5;
    return { t: "bool", v: false };
  }
  if (keywordAt(bytes, cur.pos, "null")) {
    cur.pos += 4;
    return PDF_NULL;
  }
  return parseNumberOrRef(cur);
}

// --- indirect objects and streams --------------------------------------------

/** Finds "endstream" from `from`, returning the data end (with the EOL
 * before the keyword trimmed) — the bound every real reader falls back to
 * when /Length is wrong or unresolvable. */
function scanToEndstream(bytes: Uint8Array, from: number): number {
  for (let i = from; i + 9 <= bytes.length; i += 1) {
    if (bytes[i] === 0x65 && keywordAt(bytes, i, "endstream")) {
      let end = i;
      if (end > from && bytes[end - 1] === 10) {
        end -= 1;
      }
      if (end > from && bytes[end - 1] === 13) {
        end -= 1;
      }
      return end;
    }
  }
  throw new PdfParseError("unterminated stream", from);
}

type LengthResolver = (ref: { num: number; gen: number }) => number | null;

interface IndirectObject {
  num: number;
  gen: number;
  value: PdfValue;
  end: number;
}

/** Parses one `N G obj … endobj` at `offset`. `resolveLength` supplies the
 * value of an indirect /Length; a wrong or unresolvable length falls back to
 * scanning for `endstream`. */
function parseObjectAt(bytes: Uint8Array, offset: number, resolveLength: LengthResolver): IndirectObject {
  const cur: Cursor = { bytes, pos: offset };
  const num = parseUint(cur);
  const gen = parseUint(cur);
  expectKeyword(cur, "obj");
  let value = parseValue(cur, 0);
  skipWs(cur);
  if (value.t === "dict" && keywordAt(bytes, cur.pos, "stream")) {
    cur.pos += 6; // "stream"
    if (bytes[cur.pos] === 13) {
      cur.pos += 1;
    }
    if (bytes[cur.pos] === 10) {
      cur.pos += 1;
    }
    const start = cur.pos;
    const lengthValue = value.v.get("Length");
    let length: number | null = null;
    if (lengthValue?.t === "num" && Number.isInteger(lengthValue.v) && lengthValue.v >= 0) {
      length = lengthValue.v;
    } else if (lengthValue?.t === "ref") {
      length = resolveLength(lengthValue);
    }
    let end: number;
    if (length !== null && start + length <= bytes.length) {
      const check: Cursor = { bytes, pos: start + length };
      skipWs(check);
      if (keywordAt(bytes, check.pos, "endstream")) {
        end = start + length;
        cur.pos = check.pos + 9;
      } else {
        end = scanToEndstream(bytes, start);
        cur.pos = end;
        expectKeyword(cur, "endstream");
      }
    } else {
      end = scanToEndstream(bytes, start);
      cur.pos = end;
      expectKeyword(cur, "endstream");
    }
    value = { t: "stream", dict: value.v, raw: bytes.subarray(start, end) };
    skipWs(cur);
  }
  if (keywordAt(bytes, cur.pos, "endobj")) {
    cur.pos += 6;
  }
  return { num, gen, value, end: cur.pos };
}

// --- inflation and stream decoding -------------------------------------------

async function inflate(
  data: Uint8Array,
  offset: number,
  maxBytes: number = MAX_INFLATED_STREAM_BYTES,
): Promise<Uint8Array> {
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(data);
      controller.close();
    },
  }).pipeThrough(new DecompressionStream("deflate"));
  const chunks: Uint8Array[] = [];
  let total = 0;
  const reader = stream.getReader();
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) {
        break;
      }
      if (value) {
        total += value.length;
        if (total > maxBytes) {
          throw new PdfInflateLimitError("inflated stream exceeds the size limit", offset);
        }
        chunks.push(value);
      }
    }
  } catch (error) {
    if (error instanceof PdfParseError) {
      await reader.cancel().catch(() => {});
      throw error;
    }
    throw new PdfParseError(`bad FlateDecode stream: ${String(error)}`, offset);
  }
  const result = new Uint8Array(total);
  let position = 0;
  for (const chunk of chunks) {
    result.set(chunk, position);
    position += chunk.length;
  }
  return result;
}

/** The largest /Columns a predictor may declare — far past any real image
 * or xref row, small enough that no row buffer can be a hostile allocation. */
const MAX_PREDICTOR_COLUMNS = 1 << 24;

/** Undoes TIFF Predictor 2 (horizontal differencing) in place, one row of
 * `columns` × `colors` samples at a time, at any bit depth a PDF allows. */
function undoTiffPredictor(out: Uint8Array, columns: number, colors: number, bitsPerComponent: number): void {
  const rowBytes = Math.ceil((columns * colors * bitsPerComponent) / 8);
  for (let base = 0; base < out.length; base += rowBytes) {
    const end = Math.min(base + rowBytes, out.length);
    if (bitsPerComponent === 8) {
      for (let i = base + colors; i < end; i += 1) {
        out[i] = (out[i] + out[i - colors]) & 0xff;
      }
    } else if (bitsPerComponent === 16) {
      const step = colors * 2;
      for (let i = base + step; i + 1 < end; i += 2) {
        const sum = ((out[i] << 8) | out[i + 1]) + ((out[i - step] << 8) | out[i - step + 1]);
        out[i] = (sum >> 8) & 0xff;
        out[i + 1] = sum & 0xff;
      }
    } else {
      // 1, 2 or 4 bits: samples packed big-endian within each byte.
      const mask = (1 << bitsPerComponent) - 1;
      const samples = Math.floor(((end - base) * 8) / bitsPerComponent);
      const read = (s: number): number => {
        const bit = s * bitsPerComponent;
        return (out[base + (bit >> 3)] >> (8 - bitsPerComponent - (bit & 7))) & mask;
      };
      const write = (s: number, value: number): void => {
        const bit = s * bitsPerComponent;
        const shift = 8 - bitsPerComponent - (bit & 7);
        const at = base + (bit >> 3);
        out[at] = (out[at] & ~(mask << shift)) | ((value & mask) << shift);
      };
      for (let s = colors; s < Math.min(samples, columns * colors); s += 1) {
        write(s, read(s) + read(s - colors));
      }
    }
  }
}

/** Undoes a /Predictor pre-transform. 10–15 are the PNG predictors (what
 * xref streams actually use); 2 is TIFF horizontal differencing. Parameters
 * outside what a PDF can declare are refused, never guessed at. */
function applyPredictor(
  data: Uint8Array,
  predictor: number,
  columns: number,
  colors: number,
  bitsPerComponent: number,
  offset: number,
): Uint8Array {
  if (predictor <= 1) {
    return data;
  }
  if (
    !Number.isInteger(columns) || columns < 1 || columns > MAX_PREDICTOR_COLUMNS ||
    !Number.isInteger(colors) || colors < 1 || colors > 32 ||
    ![1, 2, 4, 8, 16].includes(bitsPerComponent)
  ) {
    throw new PdfParseError("bad /DecodeParms for a predictor", offset);
  }
  if (predictor === 2) {
    const out = data.slice();
    undoTiffPredictor(out, columns, colors, bitsPerComponent);
    return out;
  }
  if (predictor < 10 || predictor > 15) {
    throw new PdfParseError(`unsupported predictor ${predictor}`, offset);
  }
  const bpp = Math.max(1, Math.ceil((colors * bitsPerComponent) / 8));
  const rowBytes = Math.ceil((columns * colors * bitsPerComponent) / 8);
  const rows = Math.floor(data.length / (rowBytes + 1));
  if (rows === 0) {
    return new Uint8Array(0);
  }
  const out = new Uint8Array(rows * rowBytes);
  const prior = new Uint8Array(rowBytes);
  for (let row = 0; row < rows; row += 1) {
    const filter = data[row * (rowBytes + 1)];
    const src = data.subarray(row * (rowBytes + 1) + 1, (row + 1) * (rowBytes + 1));
    const dst = out.subarray(row * rowBytes, (row + 1) * rowBytes);
    for (let i = 0; i < rowBytes; i += 1) {
      const raw = src[i];
      const left = i >= bpp ? dst[i - bpp] : 0;
      const up = prior[i];
      const upLeft = i >= bpp ? prior[i - bpp] : 0;
      let byte: number;
      switch (filter) {
        case 0: byte = raw; break;
        case 1: byte = raw + left; break;
        case 2: byte = raw + up; break;
        case 3: byte = raw + Math.floor((left + up) / 2); break;
        case 4: {
          const p = left + up - upLeft;
          const pa = Math.abs(p - left);
          const pb = Math.abs(p - up);
          const pc = Math.abs(p - upLeft);
          byte = raw + (pa <= pb && pa <= pc ? left : pb <= pc ? up : upLeft);
          break;
        }
        default:
          throw new PdfParseError(`bad PNG predictor filter ${filter}`, offset);
      }
      dst[i] = byte & 0xff;
    }
    prior.set(dst);
  }
  return out;
}

function decodeAsciiHex(data: Uint8Array, offset: number): Uint8Array {
  let hex = "";
  for (const byte of data) {
    if (byte === 0x3e) {
      break;
    }
    if (!isWhitespace(byte)) {
      hex += String.fromCharCode(byte);
    }
  }
  if (hex.length % 2 === 1) {
    hex += "0";
  }
  const out = new Uint8Array(hex.length / 2);
  for (let i = 0; i < out.length; i += 1) {
    const code = Number.parseInt(hex.slice(i * 2, i * 2 + 2), 16);
    if (Number.isNaN(code)) {
      throw new PdfParseError("bad ASCIIHexDecode data", offset);
    }
    out[i] = code;
  }
  return out;
}

function decodeAscii85(data: Uint8Array, offset: number): Uint8Array {
  const out: number[] = [];
  let group: number[] = [];
  for (let i = 0; i < data.length; i += 1) {
    const byte = data[i];
    if (isWhitespace(byte)) {
      continue;
    }
    if (byte === 0x7e) {
      break; // "~>"
    }
    if (byte === 0x7a && group.length === 0) {
      out.push(0, 0, 0, 0);
      continue;
    }
    if (byte < 0x21 || byte > 0x75) {
      throw new PdfParseError("bad ASCII85Decode data", offset);
    }
    group.push(byte - 0x21);
    if (group.length === 5) {
      let value = 0;
      for (const digit of group) {
        value = value * 85 + digit;
      }
      out.push((value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff);
      group = [];
    }
  }
  if (group.length > 0) {
    const missing = 5 - group.length;
    for (let i = 0; i < missing; i += 1) {
      group.push(84);
    }
    let value = 0;
    for (const digit of group) {
      value = value * 85 + digit;
    }
    const bytes = [(value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff];
    out.push(...bytes.slice(0, 4 - missing));
  }
  return new Uint8Array(out);
}

/** Image-compression filters this layer never decodes: nothing here needs
 * their samples, and a later spec copies them byte-for-byte. */
const RAW_FILTERS = new Set(["DCTDecode", "JPXDecode", "CCITTFaxDecode", "JBIG2Decode", "RunLengthDecode"]);

interface DecodeParms {
  predictor: number;
  columns: number;
  colors: number;
  bitsPerComponent: number;
}

function readDecodeParms(doc: PdfDocument, value: PdfValue | undefined): DecodeParms {
  const parms: DecodeParms = { predictor: 1, columns: 1, colors: 1, bitsPerComponent: 8 };
  const resolved = value === undefined ? PDF_NULL : resolve(doc, value);
  if (resolved.t !== "dict") {
    return parms;
  }
  const numOf = (key: string, fallback: number): number => {
    const entry = resolve(doc, resolved.v.get(key) ?? PDF_NULL);
    return entry.t === "num" ? entry.v : fallback;
  };
  parms.predictor = numOf("Predictor", 1);
  parms.columns = numOf("Columns", 1);
  parms.colors = numOf("Colors", 1);
  parms.bitsPerComponent = numOf("BitsPerComponent", 8);
  return parms;
}

/** Decodes a stream's data: /FlateDecode (with predictors), /ASCIIHexDecode
 * and /ASCII85Decode are decoded; the image-compression filters are kept raw
 * (their samples are never needed here); an unknown filter is a
 * PdfParseError. */
export async function streamData(
  doc: PdfDocument,
  stream: { dict: Map<string, PdfValue>; raw: Uint8Array },
  maxBytes: number = MAX_INFLATED_STREAM_BYTES,
): Promise<Uint8Array> {
  const offset = offsetOfStream(doc, stream.raw);
  const filterValue = resolve(doc, stream.dict.get("Filter") ?? PDF_NULL);
  const parmsValue = stream.dict.get("DecodeParms") ?? stream.dict.get("DP");
  const filters: string[] = [];
  const parmsList: PdfValue[] = [];
  if (filterValue.t === "name") {
    filters.push(filterValue.v);
    parmsList.push(parmsValue ?? PDF_NULL);
  } else if (filterValue.t === "array") {
    const resolvedParms = parmsValue === undefined ? PDF_NULL : resolve(doc, parmsValue);
    for (let i = 0; i < filterValue.v.length; i += 1) {
      const entry = resolve(doc, filterValue.v[i]);
      if (entry.t !== "name") {
        throw new PdfParseError("bad /Filter entry", offset);
      }
      filters.push(entry.v);
      parmsList.push(resolvedParms.t === "array" ? resolvedParms.v[i] ?? PDF_NULL : resolvedParms);
    }
  }
  let data = stream.raw;
  for (let i = 0; i < filters.length; i += 1) {
    const filter = filters[i];
    if (RAW_FILTERS.has(filter)) {
      return data;
    }
    if (filter === "FlateDecode") {
      const parms = readDecodeParms(doc, parmsList[i]);
      data = applyPredictor(
        await inflate(data, offset, maxBytes),
        parms.predictor,
        parms.columns,
        parms.colors,
        parms.bitsPerComponent,
        offset,
      );
    } else if (filter === "ASCIIHexDecode") {
      data = decodeAsciiHex(data, offset);
    } else if (filter === "ASCII85Decode") {
      data = decodeAscii85(data, offset);
    } else {
      throw new PdfParseError(`unsupported stream filter /${filter}`, offset);
    }
  }
  return data;
}

// --- reference resolution ----------------------------------------------------

/** Follows references to a direct value, depth-capped so reference cycles
 * terminate. A dangling reference resolves to null (as the PDF spec says). */
export function resolve(doc: PdfDocument, value: PdfValue): PdfValue {
  let current = value;
  for (let depth = 0; current.t === "ref"; depth += 1) {
    if (depth >= MAX_RESOLVE_DEPTH) {
      throw new PdfParseError(
        "reference chain too deep",
        value.t === "ref" ? offsetOfObject(doc, value.num) : 0,
      );
    }
    current = doc.objects.get(current.num) ?? PDF_NULL;
  }
  return current;
}

// --- cross-reference reading -------------------------------------------------

type XrefEntry =
  | { type: 0 }
  | { type: 1; offset: number; gen: number }
  | { type: 2; container: number; index: number };

interface XrefSection {
  entries: Map<number, XrefEntry>;
  trailer: Map<string, PdfValue>;
  prev: number | null;
}

function addEntry(entries: Map<number, XrefEntry>, num: number, entry: XrefEntry): void {
  if (!entries.has(num)) {
    entries.set(num, entry);
  }
}

/** A non-negative integer — what /Size, /W and /Index entries must be. */
function isCount(value: PdfValue): value is { t: "num"; v: number } {
  return value.t === "num" && Number.isInteger(value.v) && value.v >= 0;
}

/** Parses an xref STREAM object at `offset` (also used for hybrid /XRefStm
 * targets). The stream's own /Length must be direct (real xref streams'
 * always is; an indirect one falls back to the endstream scan). */
async function readXrefStream(bytes: Uint8Array, offset: number): Promise<XrefSection> {
  const object = parseObjectAt(bytes, offset, () => null);
  if (object.value.t !== "stream") {
    throw new PdfParseError("expected an xref stream", offset);
  }
  const dict = object.value.dict;
  // An xref stream's dictionary entries are all direct (ISO 32000 7.5.8.2),
  // so the full filter chain decodes it against a bare document — the real
  // one does not exist yet at xref-read time.
  const data = await streamData(buildDocument(bytes, "", false), object.value);
  const wValue = dict.get("W");
  const sizeValue = dict.get("Size");
  if (wValue?.t !== "array" || sizeValue === undefined || !isCount(sizeValue)) {
    throw new PdfParseError("bad xref stream dictionary", offset);
  }
  // Each field is at most 8 bytes wide (a 64-bit offset); anything else is a
  // hostile or broken /W.
  const widths = wValue.v.map((entry) => (isCount(entry) && entry.v <= 8 ? entry.v : -1));
  const rowBytes = widths.reduce((sum, w) => sum + w, 0);
  if (widths.length < 3 || widths.includes(-1) || rowBytes <= 0) {
    throw new PdfParseError("bad xref stream /W", offset);
  }
  const indexValue = dict.get("Index");
  const index: number[] = [];
  if (indexValue?.t === "array") {
    for (const entry of indexValue.v) {
      if (!isCount(entry)) {
        throw new PdfParseError("bad xref stream /Index", offset);
      }
      index.push(entry.v);
    }
  } else {
    index.push(0, sizeValue.v);
  }
  const entries = new Map<number, XrefEntry>();
  let row = 0;
  const readField = (rowStart: number, fieldIndex: number): number => {
    let start = rowStart;
    for (let i = 0; i < fieldIndex; i += 1) {
      start += widths[i];
    }
    let value = 0;
    for (let i = 0; i < widths[fieldIndex]; i += 1) {
      value = value * 256 + data[start + i];
    }
    return value;
  };
  for (let pair = 0; pair + 1 < index.length; pair += 2) {
    const first = index[pair];
    const count = index[pair + 1];
    if (count < 0 || count > MAX_PDF_OBJECTS) {
      throw new PdfParseError("xref stream subsection too large", offset);
    }
    for (let i = 0; i < count; i += 1, row += 1) {
      const rowStart = row * rowBytes;
      if (rowStart + rowBytes > data.length) {
        break; // truncated data: keep what was read
      }
      const type = widths[0] === 0 ? 1 : readField(rowStart, 0);
      const num = first + i;
      if (type === 1) {
        addEntry(entries, num, { type: 1, offset: readField(rowStart, 1), gen: readField(rowStart, 2) });
      } else if (type === 2) {
        addEntry(entries, num, { type: 2, container: readField(rowStart, 1), index: readField(rowStart, 2) });
      } else {
        addEntry(entries, num, { type: 0 });
      }
      if (entries.size > MAX_PDF_OBJECTS) {
        throw new PdfParseError("too many xref entries", offset);
      }
    }
  }
  const prevValue = dict.get("Prev");
  return {
    entries,
    trailer: dict,
    prev: prevValue?.t === "num" ? prevValue.v : null,
  };
}

/** Parses a CLASSIC xref table (plus its trailer) at `offset`, merging a
 * hybrid trailer's /XRefStm entries underneath the table's own. */
async function readXrefTable(bytes: Uint8Array, offset: number): Promise<XrefSection> {
  const cur: Cursor = { bytes, pos: offset };
  expectKeyword(cur, "xref");
  const entries = new Map<number, XrefEntry>();
  let total = 0;
  for (;;) {
    skipWs(cur);
    if (keywordAt(bytes, cur.pos, "trailer")) {
      cur.pos += 7;
      break;
    }
    const first = parseUint(cur);
    const count = parseUint(cur);
    if (count < 0 || count > MAX_PDF_OBJECTS) {
      throw new PdfParseError("xref subsection too large", cur.pos);
    }
    total += count;
    if (total > MAX_PDF_OBJECTS) {
      throw new PdfParseError("too many xref entries", cur.pos);
    }
    for (let i = 0; i < count; i += 1) {
      const entryOffset = parseUint(cur);
      const gen = parseUint(cur);
      skipWs(cur);
      const kind = bytes[cur.pos];
      cur.pos += 1;
      if (kind === 0x6e) {
        addEntry(entries, first + i, { type: 1, offset: entryOffset, gen });
      } else if (kind === 0x66) {
        addEntry(entries, first + i, { type: 0 });
      } else {
        throw new PdfParseError("bad xref entry", cur.pos - 1);
      }
    }
  }
  const trailerValue = parseValue(cur, 0);
  if (trailerValue.t !== "dict") {
    throw new PdfParseError("bad trailer", cur.pos);
  }
  const trailer = trailerValue.v;
  // Hybrid file: the classic table marks object-stream members free; the
  // /XRefStm stream carries their real (type-2) entries. The table's own
  // live entries stay authoritative, so the stream only fills gaps and
  // replaces free markers.
  const xrefStm = trailer.get("XRefStm");
  if (xrefStm?.t === "num") {
    try {
      const streamSection = await readXrefStream(bytes, xrefStm.v);
      for (const [num, entry] of streamSection.entries) {
        const existing = entries.get(num);
        if (existing === undefined || existing.type === 0) {
          entries.set(num, entry);
        }
      }
    } catch (error) {
      if (error instanceof PdfEncryptedError) {
        throw error;
      }
      // A broken hybrid stream leaves the classic table standing.
    }
  }
  const prevValue = trailer.get("Prev");
  return { entries, trailer, prev: prevValue?.t === "num" ? prevValue.v : null };
}

/** Follows the whole xref chain from `startOffset` (newest section first),
 * newest entry winning per object number, bounded by MAX_XREF_CHAIN and a
 * visited set so a cyclic /Prev terminates. */
async function readXrefChain(
  bytes: Uint8Array,
  startOffset: number,
): Promise<{ entries: Map<number, XrefEntry>; trailer: Map<string, PdfValue> }> {
  const entries = new Map<number, XrefEntry>();
  const trailer = new Map<string, PdfValue>();
  const visited = new Set<number>();
  let offset: number | null = startOffset;
  let sections = 0;
  while (offset !== null) {
    if (visited.has(offset)) {
      throw new PdfParseError("cyclic xref /Prev chain", offset);
    }
    visited.add(offset);
    sections += 1;
    if (sections > MAX_XREF_CHAIN) {
      throw new PdfParseError("xref chain too long", offset);
    }
    if (offset < 0 || offset >= bytes.length) {
      throw new PdfParseError("xref offset out of range", offset);
    }
    const probe: Cursor = { bytes, pos: offset };
    skipWs(probe);
    const section: XrefSection = keywordAt(bytes, probe.pos, "xref")
      ? await readXrefTable(bytes, probe.pos)
      : await readXrefStream(bytes, probe.pos);
    for (const [num, entry] of section.entries) {
      addEntry(entries, num, entry);
      if (entries.size > MAX_PDF_OBJECTS) {
        throw new PdfParseError("too many xref entries", offset);
      }
    }
    for (const [key, value] of section.trailer) {
      if (!trailer.has(key)) {
        trailer.set(key, value);
      }
    }
    offset = section.prev;
  }
  return { entries, trailer };
}

// --- the page tree -----------------------------------------------------------

/** The attributes a page inherits from its ancestors (ISO 32000 7.7.3.4). */
const INHERITABLE_KEYS = ["Resources", "MediaBox", "CropBox", "Rotate"] as const;

function normaliseBox(doc: PdfDocument, value: PdfValue | undefined): [number, number, number, number] | null {
  const resolved = value === undefined ? PDF_NULL : resolve(doc, value);
  if (resolved.t !== "array" || resolved.v.length !== 4) {
    return null;
  }
  const nums: number[] = [];
  for (const entry of resolved.v) {
    const item = resolve(doc, entry);
    if (item.t !== "num" || !Number.isFinite(item.v)) {
      return null;
    }
    nums.push(item.v);
  }
  const [x1, y1, x2, y2] = nums;
  return [Math.min(x1, x2), Math.min(y1, y2), Math.max(x1, x2), Math.max(y1, y2)];
}

function normaliseRotate(doc: PdfDocument, value: PdfValue | undefined): 0 | 90 | 180 | 270 {
  const resolved = value === undefined ? PDF_NULL : resolve(doc, value);
  if (resolved.t !== "num" || !Number.isFinite(resolved.v)) {
    return 0;
  }
  const wrapped = ((Math.round(resolved.v) % 360) + 360) % 360;
  return wrapped === 90 || wrapped === 180 || wrapped === 270 ? wrapped : 0;
}

/** US Letter — the fallback when a malformed page tree carries no readable
 * MediaBox anywhere. */
const DEFAULT_MEDIA_BOX: [number, number, number, number] = [0, 0, 612, 792];

function walkPageTree(
  doc: PdfDocument,
  nodeRef: { num: number; gen: number },
  inherited: Map<string, PdfValue>,
  visited: Set<number>,
  depth: number,
): void {
  const at = offsetOfObject(doc, nodeRef.num);
  if (depth > MAX_RESOLVE_DEPTH) {
    throw new PdfParseError("page tree too deep", at);
  }
  if (visited.has(nodeRef.num)) {
    throw new PdfParseError("cyclic page tree", at);
  }
  visited.add(nodeRef.num);
  const node = resolve(doc, { t: "ref", num: nodeRef.num, gen: nodeRef.gen });
  if (node.t !== "dict") {
    throw new PdfParseError(`page tree node ${nodeRef.num} is not a dictionary`, at);
  }
  const own = new Map(inherited);
  for (const key of INHERITABLE_KEYS) {
    const value = node.v.get(key);
    if (value !== undefined) {
      own.set(key, resolve(doc, value));
    }
  }
  const type = resolve(doc, node.v.get("Type") ?? PDF_NULL);
  const kids = resolve(doc, node.v.get("Kids") ?? PDF_NULL);
  if ((type.t === "name" && type.v === "Pages") || kids.t === "array") {
    if (kids.t !== "array") {
      throw new PdfParseError(`pages node ${nodeRef.num} has no /Kids`, at);
    }
    for (const kid of kids.v) {
      if (kid.t !== "ref") {
        throw new PdfParseError("page tree kid is not a reference", at);
      }
      walkPageTree(doc, kid, own, visited, depth + 1);
      if (doc.pages.length > MAX_PDF_OBJECTS) {
        throw new PdfParseError("too many pages", at);
      }
    }
    return;
  }
  const mediaBox = normaliseBox(doc, own.get("MediaBox")) ?? DEFAULT_MEDIA_BOX;
  let cropBox = normaliseBox(doc, own.get("CropBox"));
  if (cropBox !== null) {
    // The CropBox is clipped to the MediaBox (ISO 32000): a crop larger than
    // the page crops nothing.
    cropBox = [
      Math.max(cropBox[0], mediaBox[0]),
      Math.max(cropBox[1], mediaBox[1]),
      Math.min(cropBox[2], mediaBox[2]),
      Math.min(cropBox[3], mediaBox[3]),
    ];
    if (!(cropBox[2] > cropBox[0]) || !(cropBox[3] > cropBox[1])) {
      cropBox = null;
    }
  }
  const effective = cropBox ?? mediaBox;
  doc.pages.push({
    ref: { num: nodeRef.num, gen: nodeRef.gen },
    widthPt: effective[2] - effective[0],
    heightPt: effective[3] - effective[1],
    rotate: normaliseRotate(doc, own.get("Rotate")),
    mediaBoxPt: mediaBox,
    cropBoxPt: cropBox ?? undefined,
    inherited: own,
  });
}

// --- object loading ----------------------------------------------------------

/** What a parse learned about each live object: its generation (from the
 * xref entry, or from the header of the definition a rebuilt table kept)
 * and the byte offset of its `N G obj` header (an object-stream member
 * reports its container's). */
interface LiveObjects {
  generations: Map<number, number>;
  offsets: Map<number, number>;
}

function isName(value: PdfValue | undefined, name: string): boolean {
  return value?.t === "name" && value.v === name;
}

/** Parses an object stream's members into `doc.objects` — the step the
 * structured and reconstruction paths share. `accept` says which member
 * numbers to take. */
function expandObjectStream(
  doc: PdfDocument,
  live: LiveObjects,
  containerNum: number,
  container: { dict: Map<string, PdfValue> },
  data: Uint8Array,
  accept: (num: number) => boolean,
): void {
  const at = live.offsets.get(containerNum) ?? 0;
  const firstValue = container.dict.get("First");
  const nValue = container.dict.get("N");
  const first = firstValue?.t === "num" ? firstValue.v : 0;
  const count = nValue?.t === "num" ? nValue.v : 0;
  if (count < 0 || count > MAX_PDF_OBJECTS) {
    throw new PdfParseError("object stream too large", at);
  }
  const header: Cursor = { bytes: data, pos: 0 };
  for (let i = 0; i < count; i += 1) {
    const num = tryParseUint(header);
    const memberOffset = tryParseUint(header);
    if (num === null || memberOffset === null) {
      break;
    }
    if (!accept(num) || first + memberOffset >= data.length) {
      continue;
    }
    try {
      doc.objects.set(num, parseValue({ bytes: data, pos: first + memberOffset }, 0));
      // A compressed object's generation is always 0 (ISO 32000 7.5.7).
      live.generations.set(num, 0);
      live.offsets.set(num, at);
    } catch {
      // Tolerated, as for top-level objects.
    }
  }
}

/** Loads every object the xref lists into `doc.objects`. A type-1 entry is
 * read at its offset, and its header must carry the entry's object number
 * AND generation — the generation is honoured when reading, never a second
 * map key. A broken entry is skipped (tolerantly: the page walk decides
 * whether the file is still usable). Type-2 entries are then expanded out
 * of their object streams. */
async function loadObjects(doc: PdfDocument, entries: Map<number, XrefEntry>, live: LiveObjects): Promise<void> {
  const resolveLength: LengthResolver = (ref) => {
    const entry = entries.get(ref.num);
    if (entry === undefined || entry.type !== 1) {
      return null;
    }
    try {
      const object = parseObjectAt(doc.bytes, entry.offset, () => null);
      return object.value.t === "num" && Number.isInteger(object.value.v) && object.value.v >= 0
        ? object.value.v
        : null;
    } catch {
      return null;
    }
  };
  const containers = new Set<number>();
  for (const [num, entry] of entries) {
    if (entry.type === 1) {
      try {
        const object = parseObjectAt(doc.bytes, entry.offset, resolveLength);
        if (object.num === num && object.gen === entry.gen) {
          doc.objects.set(num, object.value);
          live.generations.set(num, entry.gen);
          live.offsets.set(num, entry.offset);
        }
      } catch {
        // Tolerated: the page walk fails loudly if this object mattered.
      }
    } else if (entry.type === 2) {
      containers.add(entry.container);
    }
  }
  for (const containerNum of containers) {
    const container = doc.objects.get(containerNum);
    if (container === undefined || container.t !== "stream") {
      continue;
    }
    let data: Uint8Array;
    try {
      data = await streamData(doc, container);
    } catch (error) {
      if (error instanceof PdfInflateLimitError) {
        throw error;
      }
      continue;
    }
    expandObjectStream(doc, live, containerNum, container, data, (num) => {
      const entry = entries.get(num);
      return entry?.type === 2 && entry.container === containerNum;
    });
  }
}

/** Refuses a file that genuinely needs generation-keyed lookup: every
 * reference to a live object must name that object's live generation. (A
 * reference to an absent or free object number is merely dangling — it
 * resolves to null, as ISO 32000 says — and needs no generation at all.) */
function checkGenerations(doc: PdfDocument, live: LiveObjects, trailerOffset: number): void {
  const check = (root: PdfValue, at: number): void => {
    const stack: PdfValue[] = [root];
    while (stack.length > 0) {
      const value = stack.pop() as PdfValue;
      if (value.t === "ref") {
        const generation = live.generations.get(value.num);
        if (generation !== undefined && generation !== value.gen) {
          throw new PdfGenerationError(
            `object ${value.num} is referenced as generation ${value.gen}, but its live ` +
              `definition is generation ${generation} — generation-keyed lookup is not supported`,
            at,
          );
        }
      } else if (value.t === "array") {
        for (const item of value.v) {
          stack.push(item);
        }
      } else if (value.t === "dict" || value.t === "stream") {
        for (const item of (value.t === "dict" ? value.v : value.dict).values()) {
          stack.push(item);
        }
      }
    }
  };
  for (const [num, value] of doc.objects) {
    check(value, live.offsets.get(num) ?? 0);
  }
  check({ t: "dict", v: doc.trailer }, trailerOffset);
}

// --- header, startxref, assembly ---------------------------------------------

const PDF_HEADER_SCAN_BYTES = 1024;

function readHeader(bytes: Uint8Array): { version: string; offset: number } {
  const limit = Math.min(bytes.length, PDF_HEADER_SCAN_BYTES);
  const head = latin1(bytes.subarray(0, limit));
  const match = /%PDF-(\d+\.\d+)/.exec(head);
  if (!match || match.index === undefined) {
    throw new PdfParseError("not a PDF file (no %PDF- header)", 0);
  }
  return { version: match[1], offset: match.index };
}

function findStartXref(bytes: Uint8Array): number {
  const windowSize = Math.min(bytes.length, 2048);
  const tail = latin1(bytes.subarray(bytes.length - windowSize));
  const at = tail.lastIndexOf("startxref");
  if (at < 0) {
    throw new PdfParseError("no startxref", Math.max(0, bytes.length - windowSize));
  }
  const cur: Cursor = { bytes, pos: bytes.length - windowSize + at + 9 };
  const offset = tryParseUint(cur);
  if (offset === null) {
    throw new PdfParseError("bad startxref offset", cur.pos);
  }
  return offset;
}

function isLinearised(bytes: Uint8Array): boolean {
  return latin1(bytes.subarray(0, Math.min(bytes.length, 2048))).includes("/Linearized");
}

function buildDocument(bytes: Uint8Array, version: string, reconstructed: boolean): PdfDocument {
  return {
    version,
    bytes,
    objects: new Map(),
    trailer: new Map(),
    pages: [],
    info: undefined,
    linearised: isLinearised(bytes),
    reconstructed,
  };
}

/** Resolves the catalog, walks the page tree and picks up /Info — the shared
 * tail of the structured and reconstruction paths. `trailerOffset` is where
 * the trailer was found, the offset reported when there is no catalog. */
function finishDocument(doc: PdfDocument, trailerOffset: number): void {
  const rootValue = doc.trailer.get("Root");
  const root = rootValue === undefined ? PDF_NULL : resolve(doc, rootValue);
  if (root.t !== "dict") {
    throw new PdfParseError("no document catalog", trailerOffset);
  }
  const pagesRef = root.v.get("Pages");
  if (pagesRef?.t !== "ref") {
    throw new PdfParseError(
      "catalog has no /Pages",
      rootValue?.t === "ref" ? offsetOfObject(doc, rootValue.num) : trailerOffset,
    );
  }
  walkPageTree(doc, pagesRef, new Map(), new Set(), 0);
  if (doc.pages.length === 0) {
    throw new PdfParseError("document has no pages", offsetOfObject(doc, pagesRef.num));
  }
  const infoValue = doc.trailer.get("Info");
  if (infoValue !== undefined) {
    const info = resolve(doc, infoValue);
    if (info.t === "dict") {
      doc.info = info.v;
    }
  }
}

function checkEncrypted(trailer: Map<string, PdfValue>, offset: number): void {
  if (trailer.has("Encrypt")) {
    throw new PdfEncryptedError(offset);
  }
}

async function parseWithXref(bytes: Uint8Array, version: string, headerOffset: number): Promise<PdfDocument> {
  let startOffset = findStartXref(bytes);
  // Offsets in a file with junk before the header are usually relative to
  // the header; probe both interpretations.
  const probe: Cursor = { bytes, pos: startOffset };
  skipWs(probe);
  if (
    headerOffset > 0 &&
    !keywordAt(bytes, probe.pos, "xref") &&
    !isDigit(bytes[probe.pos])
  ) {
    startOffset += headerOffset;
  }
  const { entries, trailer } = await readXrefChain(bytes, startOffset);
  checkEncrypted(trailer, startOffset);
  const doc = buildDocument(bytes, version, false);
  doc.trailer = trailer;
  const live: LiveObjects = { generations: new Map(), offsets: new Map() };
  objectOffsets.set(doc, live.offsets);
  await loadObjects(doc, entries, live);
  checkGenerations(doc, live, startOffset);
  finishDocument(doc, startOffset);
  return doc;
}

/** The reconstruction path: scan the whole buffer for `N G obj` headers,
 * take the LAST definition of each object number, find a trailer (or the
 * /Type /Catalog itself) and rebuild. A meaningful share of real-world PDFs
 * have a broken xref; refusing them outright would call files corrupt that
 * every other reader opens. */
async function parseReconstructed(bytes: Uint8Array, version: string): Promise<PdfDocument> {
  const text = latin1(bytes);
  const doc = buildDocument(bytes, version, true);
  const live: LiveObjects = { generations: new Map(), offsets: new Map() };
  objectOffsets.set(doc, live.offsets);
  const headers = new Map<number, number>();
  const headerPattern = /(?:^|[^0-9])(\d{1,8})[ \t\r\n]+(\d{1,5})[ \t\r\n]+obj\b/g;
  for (let match = headerPattern.exec(text); match !== null; match = headerPattern.exec(text)) {
    const num = Number(match[1]);
    const start = match.index + match[0].indexOf(match[1]);
    headers.set(num, start);
    if (headers.size > MAX_PDF_OBJECTS) {
      throw new PdfParseError("too many objects", start);
    }
    // Continue scanning from just past the header, not past the whole
    // object, so nested "N G obj"-looking bytes inside streams are rare but
    // harmless (last definition wins anyway).
    headerPattern.lastIndex = match.index + match[0].length;
  }
  if (headers.size === 0) {
    throw new PdfParseError("no objects found", 0);
  }
  const resolveLength: LengthResolver = (ref) => {
    const at = headers.get(ref.num);
    if (at === undefined) {
      return null;
    }
    try {
      const object = parseObjectAt(bytes, at, () => null);
      return object.value.t === "num" ? object.value.v : null;
    } catch {
      return null;
    }
  };
  for (const [num, offset] of headers) {
    try {
      const object = parseObjectAt(bytes, offset, resolveLength);
      doc.objects.set(num, object.value);
      live.generations.set(num, object.gen);
      live.offsets.set(num, offset);
    } catch {
      // Skipped: the page walk fails loudly if this object mattered.
    }
  }
  // Expand object streams so members hidden inside them are reachable too
  // (a top-level definition beats a compressed one).
  for (const [num, value] of Array.from(doc.objects)) {
    if (value.t !== "stream" || !isName(value.dict.get("Type"), "ObjStm")) {
      continue;
    }
    let data: Uint8Array;
    try {
      data = await streamData(doc, value);
    } catch (error) {
      if (error instanceof PdfInflateLimitError) {
        throw error;
      }
      continue;
    }
    expandObjectStream(doc, live, num, value, data, (member) => !doc.objects.has(member));
  }
  // Every trailer the file carries, in file order: each classic `trailer`
  // dictionary and each cross-reference stream's dictionary (a PDF 1.5 file
  // has no `trailer` keyword — its xref stream carries /Root, /Info and
  // /Encrypt).
  const trailers: Array<{ at: number; dict: Map<string, PdfValue> }> = [];
  for (let at = text.indexOf("trailer"); at >= 0; at = text.indexOf("trailer", at + 7)) {
    try {
      const value = parseValue({ bytes, pos: at + 7 }, 0);
      if (value.t === "dict") {
        trailers.push({ at, dict: value.v });
      }
    } catch {
      // Not a trailer after all.
    }
  }
  for (const [num, value] of doc.objects) {
    if (value.t === "stream" && isName(value.dict.get("Type"), "XRef")) {
      trailers.push({ at: live.offsets.get(num) ?? 0, dict: value.dict });
    }
  }
  trailers.sort((a, b) => a.at - b.at);
  // An encrypted file is refused whichever of its trailers says so.
  for (const trailer of trailers) {
    checkEncrypted(trailer.dict, trailer.at);
  }
  // The newest trailer whose /Root resolves to a catalog supplies /Root and
  // /Info; with none, the /Type /Catalog itself is found by scanning (the
  // last one in the file, as with every other object).
  const isCatalog = (value: PdfValue | undefined): boolean => {
    const resolved = value === undefined ? PDF_NULL : resolve(doc, value);
    return resolved.t === "dict" && resolved.v.get("Pages")?.t === "ref";
  };
  const newestFirst = trailers.toReversed();
  const chosen = newestFirst.find((trailer) => isCatalog(trailer.dict.get("Root")));
  let trailerAt = chosen?.at ?? 0;
  if (chosen) {
    doc.trailer = chosen.dict;
  } else {
    let catalog: { num: number; at: number } | null = null;
    for (const [num, value] of doc.objects) {
      const at = live.offsets.get(num) ?? 0;
      if (value.t === "dict" && isName(value.v.get("Type"), "Catalog") && (catalog === null || at > catalog.at)) {
        catalog = { num, at };
      }
    }
    if (catalog !== null) {
      doc.trailer = new Map([["Root", { t: "ref", num: catalog.num, gen: live.generations.get(catalog.num) ?? 0 }]]);
      const info = newestFirst.find((trailer) => trailer.dict.has("Info"))?.dict.get("Info");
      if (info !== undefined) {
        doc.trailer.set("Info", info);
      }
      trailerAt = catalog.at;
    }
  }
  checkGenerations(doc, live, trailerAt);
  finishDocument(doc, trailerAt);
  return doc;
}

/** Parses a PDF's structure: header version, xref (both flavours, /Prev
 * chains, hybrid /XRefStm), objects (object streams included), the page tree
 * with inheritance, and /Info. An unusable xref is retried ONCE by scanning
 * the whole buffer (`reconstructed: true`); an encrypted file throws
 * PdfEncryptedError; anything still unreadable throws PdfParseError with the
 * byte offset. */
export async function parsePdf(bytes: Uint8Array): Promise<PdfDocument> {
  const header = readHeader(bytes);
  try {
    return await parseWithXref(bytes, header.version, header.offset);
  } catch (error) {
    // Refusals are final: encryption, a generation conflict and a
    // decompression bomb are not an unusable xref.
    if (
      error instanceof PdfEncryptedError ||
      error instanceof PdfGenerationError ||
      error instanceof PdfInflateLimitError
    ) {
      throw error;
    }
    return parseReconstructed(bytes, header.version);
  }
}

// --- text helpers ------------------------------------------------------------

/** A PDF string or name as display text: UTF-16BE with BOM decoded as such,
 * anything else as PDFDocEncoding (≈ latin1 for the printable range). "" for
 * absent/other values. */
export function pdfTextOf(value: PdfValue | undefined): string {
  if (value === undefined) {
    return "";
  }
  if (value.t === "name") {
    return value.v;
  }
  if (value.t !== "str") {
    return "";
  }
  const bytes = value.v;
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    let text = "";
    for (let i = 2; i + 1 < bytes.length; i += 2) {
      text += String.fromCharCode((bytes[i] << 8) | bytes[i + 1]);
    }
    return text;
  }
  return latin1(bytes);
}

/** A PDF date string ("D:YYYYMMDDHHmmSS…") as "YYYY-MM-DD HH:mm", degrading
 * to the date alone (or null) as fields run out. */
export function parsePdfDate(text: string): string | null {
  const match = /^D:(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?/.exec(text.trim());
  if (!match) {
    return null;
  }
  const [, year, month, day, hour, minute] = match;
  let out = year;
  if (month) {
    out += `-${month}`;
  }
  if (day) {
    out += `-${day}`;
  }
  if (hour && minute) {
    out += ` ${hour}:${minute}`;
  }
  return out;
}

// --- the writer half: merging pages into a new PDF (SPEC-111) ----------------
//
// assemblePdf writes ONE new file holding exactly the selected pages of one or
// more parsed sources, in the plan's order. No page is interpreted and no
// stream is re-encoded: a copied stream keeps its /Filter, /DecodeParms and
// raw bytes verbatim (only /Length is restated, as a direct number), which is
// what makes the merge safe on files this product cannot render. The copy:
//
// 1. Reachability. From each selected page's dictionary (and its kept
//    annotations) every reference is walked breadth-first (pdfReachable) —
//    never through /Parent, and never into the catalog, the page tree or
//    another page. The visited set is per source, so an object two selected
//    pages of one file share is copied once; the same object in two files is
//    copied twice (no content dedupe across sources: comparing decoded
//    streams would defeat the verbatim copy).
// 2. Renumbering. Each reached object takes the next output number, from 1,
//    and every reference inside a copied value is rewritten through its
//    source's map (rewritePdfRefs). A reference to a page becomes the output
//    page it became (its first appearance), or null when it is not output.
// 3. Inheritance. /Resources, /MediaBox, /CropBox and /Rotate are written
//    onto the copied page itself, from PdfPageRef.inherited, so it no longer
//    leans on a parent chain that is not copied; rotateBy is added here.
// 4. Structure. A new /Pages, /Root and /Info (from the plan, never from a
//    source, so two sources cannot fight over the title), the sources'
//    /OCProperties and /Dests merged, and one outline entry per source.
// 5. Emission. Objects in number order, then a classic xref table and
//    trailer under a PDF 1.7 header: no object streams, no compression pass,
//    no linearisation — larger than a producer-optimised file, and every
//    reader opens it.
//
// Annotations are page-owned: each output page gets its own copy of the
// annotation dictionaries on its /Annots (ISO 32000 lets an annotation sit on
// one page only — a duplicated page needs its own), while everything they
// reach (appearance streams, actions) is shared like any other object.
//
// What cannot be carried is REPORTED, never silently lost (PdfDropRecord —
// SPEC-008's DropRecord shape over its own closed union, as SPEC-106's
// DeckDropRecord is; the text product's DropKind would pull documentModel.ts
// across a product boundary). Deterministic: the same plan writes the same
// bytes — no clock unless `plan.modified`, no randomness, insertion-ordered
// maps. Sources are only ever read; nothing in a PdfDocument is mutated.

/** A merge plan's ceilings, refused up front with a message naming them. */
export const MAX_MERGE_PAGES = 2000;
export const MAX_MERGE_SOURCE_BYTES = 500 * 1024 * 1024;
/** Copying yields to the main thread every this many pages. */
export const MERGE_YIELD_PAGES = 8;

/** What a merge could not carry (the closed union of SPEC-111). */
export type PdfDropKind =
  | "acroForm"
  | "structureTree"
  | "javascript"
  | "openAction"
  | "embeddedFile"
  | "namedDestination"
  | "outlineEntry"
  | "annotation"
  | "pageLabel"
  | "encryptedSource"
  | "unreadableSource";

/** Every PdfDropKind, in the union's order (a report's grouping order). */
export const PDF_DROP_KINDS: readonly PdfDropKind[] = [
  "acroForm",
  "structureTree",
  "javascript",
  "openAction",
  "embeddedFile",
  "namedDestination",
  "outlineEntry",
  "annotation",
  "pageLabel",
  "encryptedSource",
  "unreadableSource",
];

export interface PdfDropRecord {
  kind: PdfDropKind;
  /** The plan's key for the source it came from. */
  source: string;
  /** The 1-based page of that source it sat on (an annotation) or pointed
   * at (a bookmark, a named destination), where it had one. */
  page?: number;
  reason: string;
}

export interface PdfPageSelection {
  /** A parsed source, by the key the plan uses for reporting. */
  source: string;
  doc: PdfDocument;
  /** 0-based index into doc.pages. */
  pageIndex: number;
  /** Added to the page's own /Rotate, normalised into 0|90|180|270. */
  rotateBy?: 0 | 90 | 180 | 270;
}

export interface PdfAssemblyPlan {
  /** Output order. */
  pages: PdfPageSelection[];
  title: string;
  author?: string;
  subject?: string;
  /** One top-level outline entry per source, in first-appearance order. */
  outline: boolean;
  /** /Info's dates. Absent, none is written — the bytes stay a function of
   * the plan alone. */
  modified?: Date;
}

export interface PdfAssemblyResult {
  bytes: Uint8Array;
  dropped: PdfDropRecord[];
}

/** Optional hooks beside the plan: progress, cancellation at every page
 * boundary, and the yield taken every MERGE_YIELD_PAGES pages. */
export interface PdfAssemblyHooks {
  onProgress?: (done: number, total: number) => void;
  /** Checked at every page boundary; true stops the copy with nothing
   * produced (PdfAssemblyCancelled). */
  cancelled?: () => boolean;
  /** Awaited every MERGE_YIELD_PAGES pages; yieldToMainThread by default. */
  between?: () => Promise<void>;
}

/** A plan refused before anything is copied: an encrypted or unreadable
 * source (one record per file, each named), or a bound. Nothing partial is
 * ever produced. */
export class PdfAssemblyRefusal extends Error {
  /** One record per refused source; empty when a bound refused the plan. */
  readonly dropped: PdfDropRecord[];

  constructor(message: string, dropped: PdfDropRecord[] = [], cause?: unknown) {
    super(message, cause === undefined ? undefined : { cause });
    this.dropped = dropped;
  }
}

/** A copy stopped at a page boundary by its `cancelled` hook. */
export class PdfAssemblyCancelled extends Error {
  constructor() {
    super("The merge was cancelled.");
  }
}

/** Degrees as a page /Rotate: the nearest multiple of 90, in 0..270 (so
 * -90 is 270 and 450 is 90). */
export function normaliseRotation(degrees: number): 0 | 90 | 180 | 270 {
  if (!Number.isFinite(degrees)) {
    return 0;
  }
  const quarter = ((Math.round(degrees / 90) % 4) + 4) % 4;
  return (quarter * 90) as 0 | 90 | 180 | 270;
}

function megabytes(bytes: number): string {
  return `${Math.ceil(bytes / (1024 * 1024))} MB`;
}

/** Why a plan of `pageCount` pages over sources totalling `sourceBytes`
 * breaks a merge bound — naming the limit — or null. */
export function mergeLimitRefusal(pageCount: number, sourceBytes: number): string | null {
  if (pageCount > MAX_MERGE_PAGES) {
    return (
      `A merged PDF can have at most ${MAX_MERGE_PAGES} pages; this one would have ${pageCount}. ` +
      `Remove ${pageCount - MAX_MERGE_PAGES} and try again.`
    );
  }
  if (sourceBytes > MAX_MERGE_SOURCE_BYTES) {
    return (
      `The PDFs in one merge can add up to at most ${megabytes(MAX_MERGE_SOURCE_BYTES)}; these add up to ` +
      `${megabytes(sourceBytes)}. Remove a source and try again.`
    );
  }
  return null;
}

/** The record refusing one source: password-protected (SPEC-107's
 * PdfEncryptedError) or otherwise unreadable. */
export function sourceRefusalRecord(source: string, cause: unknown): PdfDropRecord {
  if (cause instanceof PdfEncryptedError) {
    return { kind: "encryptedSource", source, reason: `“${source}” is password-protected, so it cannot be merged.` };
  }
  const detail = cause instanceof Error && cause.message ? ` (${cause.message})` : "";
  return { kind: "unreadableSource", source, reason: `“${source}” could not be read as a PDF${detail}, so it cannot be merged.` };
}

function refusalMessage(records: readonly PdfDropRecord[]): string {
  const advice =
    records.length === 1
      ? "Remove it from the plan or replace it, then try again."
      : "Remove them from the plan or replace them, then try again.";
  return `${records.map((record) => record.reason).join(" ")} ${advice}`;
}

export interface PdfSourceInput {
  /** The plan's key for this source. */
  source: string;
  bytes: Uint8Array;
}

/** Parses every source of a plan, refusing the WHOLE plan when any of them
 * is encrypted or unreadable: every offending file is named (one record
 * each), no partial set of sources is returned, and none is silently
 * dropped. `parse` is SPEC-107's reader unless a caller caches it. */
export async function openPdfSources(
  inputs: readonly PdfSourceInput[],
  parse: (bytes: Uint8Array) => Promise<PdfDocument> = parsePdf,
): Promise<PdfDocument[]> {
  const docs: PdfDocument[] = [];
  const refused: PdfDropRecord[] = [];
  let firstCause: unknown;
  for (const input of inputs) {
    try {
      docs.push(await parse(input.bytes));
    } catch (cause) {
      refused.push(sourceRefusalRecord(input.source, cause));
      firstCause ??= cause;
    }
  }
  if (refused.length > 0) {
    throw new PdfAssemblyRefusal(refusalMessage(refused), refused, firstCause);
  }
  return docs;
}

// --- reachability and reference rewriting -------------------------------------

/** Object numbers reachable from `roots` by following indirect references
 * breadth-first, each once, in the order first reached — the copy set of a
 * merge. A reference under a /Parent key is never followed (the page tree is
 * rebuilt; an annotation's parent is copied on its own account or belongs to
 * a structure the merge drops), nor a stream's /Length (restated directly),
 * nor a dangling one, nor one `skip` turns away (the catalog, the page tree,
 * pages, what is already copied) — and nothing past a skipped object. */
export function pdfReachable(
  doc: PdfDocument,
  roots: readonly PdfValue[],
  skip: (num: number) => boolean = () => false,
): number[] {
  const order: number[] = [];
  const seen = new Set<number>();
  const scan = (value: PdfValue, key: string | null): void => {
    switch (value.t) {
      case "ref":
        if (key !== "Parent" && !seen.has(value.num) && doc.objects.has(value.num) && !skip(value.num)) {
          seen.add(value.num);
          order.push(value.num);
        }
        return;
      case "array":
        for (const item of value.v) {
          scan(item, null);
        }
        return;
      case "dict":
        for (const [entryKey, item] of value.v) {
          scan(item, entryKey);
        }
        return;
      case "stream":
        for (const [entryKey, item] of value.dict) {
          if (entryKey !== "Length") {
            scan(item, entryKey);
          }
        }
        return;
      default:
    }
  };
  for (const root of roots) {
    scan(root, null);
  }
  for (let head = 0; head < order.length; head += 1) {
    scan(doc.objects.get(order[head]) ?? PDF_NULL, null);
    if (order.length > MAX_PDF_OBJECTS) {
      throw new PdfParseError("too many reachable objects", offsetOfObject(doc, order[head]));
    }
  }
  return order;
}

/** A deep copy of `value` with every indirect reference replaced by what
 * `map` returns for it — given the reference and the dictionary key it sits
 * under (null in an array or at the top). `undefined` drops a dictionary
 * entry (an array item becomes null). A stream's raw data is SHARED, never
 * copied or re-encoded — its /Filter and /DecodeParms come across like any
 * entry — and its /Length is restated as the raw length. `value` is never
 * mutated. */
export function rewritePdfRefs(
  value: PdfValue,
  map: (ref: { num: number; gen: number }, key: string | null) => PdfValue | undefined,
): PdfValue {
  const copyDict = (dict: Map<string, PdfValue>, restated?: string): Map<string, PdfValue> => {
    const out = new Map<string, PdfValue>();
    for (const [key, item] of dict) {
      if (key === restated) {
        // Holds the key's place; the caller sets its value.
        out.set(key, PDF_NULL);
        continue;
      }
      const copied = copy(item, key);
      if (copied !== undefined) {
        out.set(key, copied);
      }
    }
    return out;
  };
  const copy = (item: PdfValue, key: string | null): PdfValue | undefined => {
    switch (item.t) {
      case "ref":
        return map(item, key);
      case "array":
        return { t: "array", v: item.v.map((entry) => copy(entry, null) ?? PDF_NULL) };
      case "dict":
        return { t: "dict", v: copyDict(item.v) };
      case "stream": {
        const dict = copyDict(item.dict, "Length");
        dict.set("Length", { t: "num", v: item.raw.length });
        return { t: "stream", dict, raw: item.raw };
      }
      default:
        return item;
    }
  };
  return copy(value, null) ?? PDF_NULL;
}

// --- serialisation ---------------------------------------------------------------

/** Text as bytes, one byte per char (every string the writer builds is
 * ASCII or already one char per byte). */
function latin1Bytes(text: string): Uint8Array {
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i += 1) {
    out[i] = text.charCodeAt(i) & 0xff;
  }
  return out;
}

/** Bytes as a string, one char per byte — exact for every byte value, which
 * the reader's windows-1252 `latin1` is not (0x80–0x9f). */
function byteString(bytes: Uint8Array): string {
  let text = "";
  for (let i = 0; i < bytes.length; i += 8192) {
    text += String.fromCharCode(...bytes.subarray(i, i + 8192));
  }
  return text;
}

/** A number in PDF syntax, which has no exponent notation. */
function formatPdfNumber(value: number): string {
  if (!Number.isFinite(value) || value === 0) {
    return "0";
  }
  const text = String(value);
  if (!/e/i.test(text)) {
    return text;
  }
  // No real PDF value is this large; clamp rather than write an exponent.
  if (Math.abs(value) >= 1e21) {
    return value > 0 ? "100000000000000000000" : "-100000000000000000000";
  }
  const fixed = value.toFixed(10).replace(/\.?0+$/, "");
  return fixed === "-0" ? "0" : fixed;
}

function hex2(byte: number): string {
  return byte.toString(16).toUpperCase().padStart(2, "0");
}

/** A name in PDF syntax: anything but a regular printable character (and
 * the `#` escape itself) is written `#xx`. */
function formatPdfName(name: string): string {
  let out = "/";
  for (let i = 0; i < name.length; i += 1) {
    const code = name.charCodeAt(i) & 0xff;
    out += code > 0x20 && code < 0x7f && code !== 0x23 && !isDelimiter(code) ? name[i] : `#${hex2(code)}`;
  }
  return out;
}

/** A string in PDF syntax: literal when every byte is printable ASCII (with
 * `\`, `(` and `)` escaped), hex otherwise — exact for any bytes. */
function formatPdfString(bytes: Uint8Array): string {
  let printable = true;
  for (const byte of bytes) {
    if (byte < 0x20 || byte > 0x7e) {
      printable = false;
      break;
    }
  }
  if (printable) {
    let out = "(";
    for (const byte of bytes) {
      const char = String.fromCharCode(byte);
      out += byte === 0x28 || byte === 0x29 || byte === 0x5c ? `\\${char}` : char;
    }
    return `${out})`;
  }
  let out = "<";
  for (const byte of bytes) {
    out += hex2(byte);
  }
  return `${out}>`;
}

/** One value in PDF syntax (a stream writes its dictionary; the emitter
 * writes its data). */
function formatPdfValue(value: PdfValue): string {
  switch (value.t) {
    case "null":
      return "null";
    case "bool":
      return value.v ? "true" : "false";
    case "num":
      return formatPdfNumber(value.v);
    case "str":
      return formatPdfString(value.v);
    case "name":
      return formatPdfName(value.v);
    case "ref":
      return `${value.num} ${value.gen} R`;
    case "array":
      return `[${value.v.map(formatPdfValue).join(" ")}]`;
    case "dict":
      return `<<${Array.from(value.v, ([key, item]) => ` ${formatPdfName(key)} ${formatPdfValue(item)}`).join("")} >>`;
    case "stream":
      return formatPdfValue({ t: "dict", v: value.dict });
  }
}

/** The file: header, objects 1..n in number order, a classic xref table and
 * the trailer. Every offset is a byte offset — all syntax is ASCII. */
function emitPdf(objects: readonly PdfValue[], trailer: Map<string, PdfValue>): Uint8Array {
  const chunks: Uint8Array[] = [];
  let length = 0;
  const push = (data: string | Uint8Array): void => {
    const bytes = typeof data === "string" ? latin1Bytes(data) : data;
    chunks.push(bytes);
    length += bytes.length;
  };
  // The comment line of four high bytes marks the file as binary.
  push("%PDF-1.7\n%âãÏÓ\n");
  const offsets: number[] = [];
  for (let num = 1; num < objects.length; num += 1) {
    offsets.push(length);
    const value = objects[num];
    if (value.t === "stream") {
      push(`${num} 0 obj\n${formatPdfValue(value)}\nstream\n`);
      push(value.raw);
      push("\nendstream\nendobj\n");
    } else {
      push(`${num} 0 obj\n${formatPdfValue(value)}\nendobj\n`);
    }
  }
  const xrefAt = length;
  const rows = [`xref\n0 ${objects.length}\n0000000000 65535 f \n`];
  for (const offset of offsets) {
    rows.push(`${String(offset).padStart(10, "0")} 00000 n \n`);
  }
  push(rows.join(""));
  push(`trailer\n${formatPdfValue({ t: "dict", v: trailer })}\nstartxref\n${xrefAt}\n%%EOF\n`);
  const out = new Uint8Array(length);
  let cursor = 0;
  for (const chunk of chunks) {
    out.set(chunk, cursor);
    cursor += chunk.length;
  }
  return out;
}

// --- the copier --------------------------------------------------------------------

const outRef = (num: number): PdfValue => ({ t: "ref", num, gen: 0 });
const nameValue = (v: string): PdfValue => ({ t: "name", v });
const arrayValue = (v: PdfValue[]): PdfValue => ({ t: "array", v });
const dictValue = (entries: Array<[string, PdfValue]>): PdfValue => ({ t: "dict", v: new Map(entries) });

/** A text string (ISO 32000 7.9.2.2): printable ASCII as is, anything else
 * UTF-16BE behind its byte-order mark. */
function textStringValue(text: string): PdfValue {
  if (/^[\x20-\x7e]*$/.test(text)) {
    return { t: "str", v: latin1Bytes(text) };
  }
  const out = new Uint8Array(2 + text.length * 2);
  out[0] = 0xfe;
  out[1] = 0xff;
  for (let i = 0; i < text.length; i += 1) {
    const code = text.charCodeAt(i);
    out[2 + i * 2] = code >> 8;
    out[3 + i * 2] = code & 0xff;
  }
  return { t: "str", v: out };
}

/** A date string (ISO 32000 7.9.4), in UTC. */
function dateStringValue(date: Date): PdfValue {
  const pad = (value: number, width = 2): string => String(value).padStart(width, "0");
  return {
    t: "str",
    v: latin1Bytes(
      `D:${pad(date.getUTCFullYear(), 4)}${pad(date.getUTCMonth() + 1)}${pad(date.getUTCDate())}` +
        `${pad(date.getUTCHours())}${pad(date.getUTCMinutes())}${pad(date.getUTCSeconds())}Z`,
    ),
  };
}

/** A named destination of a source. */
interface NamedDest {
  /** The destination as written: an array, or a dictionary holding /D. */
  value: PdfValue;
  /** The source page it points at; null when that cannot be read. */
  pageIndex: number | null;
  /** Its name in the output, when it is kept. */
  outName?: string;
}

/** Everything the copier knows about one source. */
interface MergeSource {
  key: string;
  doc: PdfDocument;
  catalog: Map<string, PdfValue>;
  /** Never copied: the catalog and the page tree's interior nodes. */
  stop: Set<number>;
  /** Page object number → page index. */
  pageIndexByNum: Map<number, number>;
  /** Page index → the output page it first became. */
  firstOutPage: Map<number, number>;
  /** Source object number → output number: the per-source visited set. */
  shared: Map<number, number>;
  /** "n:<name>" for the catalog's /Dests dictionary, "s:<bytes>" for the
   * /Dests name tree — the two namespaces a destination name can live in. */
  named: Map<string, NamedDest>;
}

/** A page tree's interior nodes, walked from its root with a visited set. */
function pageTreeNodes(doc: PdfDocument, root: PdfValue | undefined, into: Set<number>): void {
  const stack = root?.t === "ref" ? [root.num] : [];
  const seen = new Set<number>();
  while (stack.length > 0) {
    const num = stack.pop() as number;
    if (seen.has(num)) {
      continue;
    }
    seen.add(num);
    const node = doc.objects.get(num);
    const kids = node?.t === "dict" ? resolve(doc, node.v.get("Kids") ?? PDF_NULL) : PDF_NULL;
    if (kids.t !== "array") {
      continue; // a leaf: a page
    }
    into.add(num);
    for (const kid of kids.v) {
      if (kid.t === "ref") {
        stack.push(kid.num);
      }
    }
  }
}

function mergeSource(key: string, doc: PdfDocument): MergeSource {
  const rootValue = doc.trailer.get("Root");
  const root = rootValue === undefined ? PDF_NULL : resolve(doc, rootValue);
  const catalog = root.t === "dict" ? root.v : new Map<string, PdfValue>();
  const stop = new Set<number>();
  if (rootValue?.t === "ref") {
    stop.add(rootValue.num);
  }
  pageTreeNodes(doc, catalog.get("Pages"), stop);
  return {
    key,
    doc,
    catalog,
    stop,
    pageIndexByNum: new Map(doc.pages.map((page, index) => [page.ref.num, index])),
    firstOutPage: new Map(),
    shared: new Map(),
    named: new Map(),
  };
}

/** A name tree's entries — keys one char per byte — walked with a visited
 * set under the resolution depth cap. */
function nameTreeEntries(doc: PdfDocument, root: PdfValue | undefined): Array<[string, PdfValue]> {
  const out: Array<[string, PdfValue]> = [];
  const seen = new Set<number>();
  const visit = (value: PdfValue, depth: number): void => {
    if (depth > MAX_RESOLVE_DEPTH || (value.t === "ref" && seen.has(value.num))) {
      return;
    }
    if (value.t === "ref") {
      seen.add(value.num);
    }
    const node = resolve(doc, value);
    if (node.t !== "dict") {
      return;
    }
    const names = resolve(doc, node.v.get("Names") ?? PDF_NULL);
    if (names.t === "array") {
      for (let i = 0; i + 1 < names.v.length && out.length < MAX_PDF_OBJECTS; i += 2) {
        const key = resolve(doc, names.v[i]);
        if (key.t === "str") {
          out.push([byteString(key.v), names.v[i + 1]]);
        }
      }
    }
    const kids = resolve(doc, node.v.get("Kids") ?? PDF_NULL);
    if (kids.t === "array") {
      for (const kid of kids.v) {
        visit(kid, depth + 1);
      }
    }
  };
  if (root !== undefined) {
    visit(root, 0);
  }
  return out;
}

/** The source page an explicit destination points at — an array whose first
 * element is a page reference, or a dictionary holding one as /D — or null. */
function destPageIndex(src: MergeSource, dest: PdfValue, depth = 0): number | null {
  const value = resolve(src.doc, dest);
  if (value.t === "array") {
    const first = value.v[0];
    return first?.t === "ref" ? src.pageIndexByNum.get(first.num) ?? null : null;
  }
  if (value.t === "dict" && depth === 0) {
    return destPageIndex(src, value.v.get("D") ?? PDF_NULL, 1);
  }
  return null;
}

/** What becomes of a destination in the output: kept (as a string `rename`
 * when its name changes or was a PDF 1.1 name object), or dropped — saying
 * where it pointed. */
type DestVerdict = { keep: true; rename?: string } | { keep: false; page?: number; where: string };

function judgeDest(src: MergeSource, dest: PdfValue): DestVerdict {
  const value = resolve(src.doc, dest);
  if (value.t === "name" || value.t === "str") {
    const text = value.t === "name" ? value.v : byteString(value.v);
    const label = value.t === "name" ? value.v : pdfTextOf(value);
    const entry = src.named.get(`${value.t === "name" ? "n" : "s"}:${text}`);
    if (!entry) {
      return { keep: false, where: `the destination “${label}”, which does not exist` };
    }
    if (entry.outName === undefined) {
      return entry.pageIndex === null
        ? { keep: false, where: `the destination “${label}”, which could not be read` }
        : { keep: false, page: entry.pageIndex + 1, where: `page ${entry.pageIndex + 1}, which is not in the merged PDF` };
    }
    return value.t === "name" || entry.outName !== text ? { keep: true, rename: entry.outName } : { keep: true };
  }
  const pageIndex = destPageIndex(src, value);
  if (pageIndex === null) {
    return { keep: false, where: "a destination that could not be read" };
  }
  if (!src.firstOutPage.has(pageIndex)) {
    return { keep: false, page: pageIndex + 1, where: `page ${pageIndex + 1}, which is not in the merged PDF` };
  }
  return { keep: true };
}

/** The internal destination an annotation or bookmark jumps to: its /Dest,
 * or the /D of its /GoTo action. */
function internalDest(
  doc: PdfDocument,
  dict: Map<string, PdfValue>,
): { dest: PdfValue; action?: Map<string, PdfValue> } | null {
  const dest = dict.get("Dest");
  if (dest !== undefined) {
    return { dest };
  }
  const action = resolve(doc, dict.get("A") ?? PDF_NULL);
  if (action.t === "dict" && isName(resolve(doc, action.v.get("S") ?? PDF_NULL), "GoTo")) {
    return { dest: action.v.get("D") ?? PDF_NULL, action: action.v };
  }
  return null;
}

/** `dict` with its destination renamed to the string `name` — on a copy, an
 * action dictionary inlined so a shared one is never edited. */
function renameDest(dict: Map<string, PdfValue>, target: { action?: Map<string, PdfValue> }, name: string): void {
  const renamed: PdfValue = { t: "str", v: latin1Bytes(name) };
  if (target.action) {
    dict.set("A", { t: "dict", v: new Map(target.action).set("D", renamed) });
  } else {
    dict.set("Dest", renamed);
  }
}

/** Reads a source's named destinations and decides each one's output name:
 * kept under its own name, or under a source-qualified one on a collision
 * ("<source>:<name>", then numbered), when its page is in the output; else
 * dropped with a `namedDestination` record. */
function nameDestinations(src: MergeSource, taken: Set<string>, dropped: PdfDropRecord[]): void {
  const { doc, catalog } = src;
  const dests = resolve(doc, catalog.get("Dests") ?? PDF_NULL);
  if (dests.t === "dict") {
    for (const [name, value] of dests.v) {
      src.named.set(`n:${name}`, { value, pageIndex: destPageIndex(src, value) });
    }
  }
  const names = resolve(doc, catalog.get("Names") ?? PDF_NULL);
  if (names.t === "dict") {
    for (const [key, value] of nameTreeEntries(doc, names.v.get("Dests"))) {
      src.named.set(`s:${key}`, { value, pageIndex: destPageIndex(src, value) });
    }
  }
  const prefix = byteString(new TextEncoder().encode(`${src.key}:`));
  for (const [key, entry] of src.named) {
    const text = key.slice(2);
    const label = key.startsWith("s:") ? pdfTextOf({ t: "str", v: latin1Bytes(text) }) : text;
    if (entry.pageIndex === null || !src.firstOutPage.has(entry.pageIndex)) {
      dropped.push({
        kind: "namedDestination",
        source: src.key,
        ...(entry.pageIndex === null ? {} : { page: entry.pageIndex + 1 }),
        reason:
          entry.pageIndex === null
            ? `The named destination “${label}” could not be read, so it was removed.`
            : `The named destination “${label}” pointed at page ${entry.pageIndex + 1}, which is not in the merged PDF, so it was removed.`,
      });
      continue;
    }
    let outName = text;
    for (let n = 1; taken.has(outName); n += 1) {
      outName = n === 1 ? `${prefix}${text}` : `${prefix}${text} (${n})`;
    }
    taken.add(outName);
    entry.outName = outName;
  }
}

/** One bookmark of the output, before numbering. `src` is the source whose
 * values `link` and `extra` hold (null: already output values). */
interface OutlineNode {
  title: PdfValue;
  link: Array<[string, PdfValue]>;
  extra: Array<[string, PdfValue]>;
  open: boolean;
  children: OutlineNode[];
  src: MergeSource | null;
}

/** A source's outline tree, filtered: an entry pointing at a page that is
 * not in the output is removed (with an `outlineEntry` record) and its
 * surviving descendants take its place; a renamed destination follows its
 * new name. Walked with a visited set and the resolution depth cap. */
function readOutline(src: MergeSource, dropped: PdfDropRecord[]): OutlineNode[] {
  const { doc } = src;
  const root = resolve(doc, src.catalog.get("Outlines") ?? PDF_NULL);
  if (root.t !== "dict") {
    return [];
  }
  const seen = new Set<number>();
  const readList = (first: PdfValue | undefined, depth: number): OutlineNode[] => {
    const nodes: OutlineNode[] = [];
    let cursor = first;
    while (cursor?.t === "ref" && !seen.has(cursor.num) && depth <= MAX_RESOLVE_DEPTH && seen.size < MAX_PDF_OBJECTS) {
      seen.add(cursor.num);
      const item = resolve(doc, cursor);
      if (item.t !== "dict") {
        break;
      }
      cursor = item.v.get("Next");
      const children = readList(item.v.get("First"), depth + 1);
      const title = resolve(doc, item.v.get("Title") ?? PDF_NULL);
      const count = resolve(doc, item.v.get("Count") ?? PDF_NULL);
      const node: OutlineNode = {
        title: title.t === "str" ? title : { t: "str", v: new Uint8Array(0) },
        link: [],
        extra: [],
        open: count.t === "num" && count.v > 0,
        children,
        src,
      };
      for (const key of ["C", "F"]) {
        const value = item.v.get(key);
        if (value !== undefined) {
          node.extra.push([key, value]);
        }
      }
      const target = internalDest(doc, item.v);
      if (target) {
        const verdict = judgeDest(src, target.dest);
        if (!verdict.keep) {
          dropped.push({
            kind: "outlineEntry",
            source: src.key,
            ...(verdict.page === undefined ? {} : { page: verdict.page }),
            reason: `The bookmark “${pdfTextOf(node.title)}” pointed at ${verdict.where}, so it was removed.`,
          });
          nodes.push(...children);
          continue;
        }
        const link = new Map<string, PdfValue>(target.action ? [["A", item.v.get("A") ?? PDF_NULL]] : [["Dest", target.dest]]);
        if (verdict.rename !== undefined) {
          renameDest(link, target, verdict.rename);
        }
        node.link = Array.from(link);
      } else if (item.v.has("A")) {
        // Any other action (a web link, a named action) comes across as is.
        node.link = [["A", item.v.get("A") ?? PDF_NULL]];
      }
      nodes.push(node);
    }
    return nodes;
  };
  return readList(root.v.get("First"), 0);
}

/** How many descendants of an outline node show when it is open. */
function openCount(node: OutlineNode): number {
  return node.children.reduce((sum, child) => sum + 1 + (child.open ? openCount(child) : 0), 0);
}

/** One record per catalog-level construct a merge does not carry. */
function catalogDrops(src: MergeSource, pageCount: number, dropped: PdfDropRecord[]): void {
  const { doc, catalog, key } = src;
  const record = (kind: PdfDropKind, reason: string): void => {
    dropped.push({ kind, source: key, reason });
  };
  if (catalog.has("AcroForm")) {
    const form = resolve(doc, catalog.get("AcroForm") ?? PDF_NULL);
    const sigFlags = form.t === "dict" ? resolve(doc, form.v.get("SigFlags") ?? PDF_NULL) : PDF_NULL;
    record(
      "acroForm",
      "Its form fields are no longer fillable: what was filled in still shows, but the form itself is not carried into a merged PDF." +
        (sigFlags.t === "num" && sigFlags.v > 0 ? " Its digital signature does not survive a merge and was not carried." : ""),
    );
  }
  if (catalog.has("StructTreeRoot")) {
    record("structureTree", "Its structure tree (the tags that make it accessible) was not carried, so the merged PDF is not a tagged PDF.");
  }
  const names = resolve(doc, catalog.get("Names") ?? PDF_NULL);
  if (names.t === "dict" && names.v.has("JavaScript")) {
    record("javascript", "Its document scripts (JavaScript) were removed.");
  }
  // A catalog's /AA holds JavaScript actions only (ISO 32000 table 197).
  if (catalog.has("AA")) {
    record("javascript", "Its document actions (scripts run on closing, saving or printing) were removed.");
  }
  if (catalog.has("OpenAction")) {
    record("openAction", "What it did on opening (its open action) was removed.");
  }
  if (names.t === "dict" && names.v.has("EmbeddedFiles")) {
    const files = nameTreeEntries(doc, names.v.get("EmbeddedFiles"));
    if (files.length === 0) {
      record("embeddedFile", "Its attached files were not carried.");
    }
    for (const [name] of files) {
      record("embeddedFile", `The attached file “${pdfTextOf({ t: "str", v: latin1Bytes(name) })}” was not carried.`);
    }
  }
  if (catalog.has("PageLabels")) {
    record("pageLabel", `Its page labels were removed; the merged PDF numbers its pages 1 to ${pageCount}.`);
  }
}

/** The value a page's inheritable attribute takes, as written: the page's
 * own entry, else the nearest ancestor's — a reference kept a reference, so
 * a /Resources the whole tree shares is still copied once. What the reader
 * resolved (PdfPageRef.inherited) is the authority: where a malformed
 * /Parent chain disagrees with it, its resolved value is written instead. */
function inheritedValue(doc: PdfDocument, page: PdfPageRef, key: string): PdfValue | undefined {
  const resolved = page.inherited.get(key);
  if (resolved === undefined) {
    return undefined;
  }
  let node = resolve(doc, { t: "ref", num: page.ref.num, gen: page.ref.gen });
  for (let depth = 0; node.t === "dict" && depth <= MAX_RESOLVE_DEPTH; depth += 1) {
    const raw = node.v.get(key);
    if (raw !== undefined) {
      return resolve(doc, raw) === resolved ? raw : resolved;
    }
    node = resolve(doc, node.v.get("Parent") ?? PDF_NULL);
  }
  return resolved;
}

const CATALOG_NUM = 1;
const PAGES_NUM = 2;

/** Refuses a plan before anything is copied: an encrypted source, a
 * selection naming a page its source does not have, an empty plan, or a
 * bound (MAX_MERGE_PAGES, MAX_MERGE_SOURCE_BYTES). */
function refusePlan(plan: PdfAssemblyPlan, sources: Map<PdfDocument, MergeSource>): void {
  const refused: PdfDropRecord[] = [];
  for (const src of sources.values()) {
    if (src.doc.trailer.has("Encrypt")) {
      refused.push(sourceRefusalRecord(src.key, new PdfEncryptedError(0)));
    }
  }
  for (const selection of plan.pages) {
    const page = Number.isInteger(selection.pageIndex) ? selection.doc.pages[selection.pageIndex] : undefined;
    const dict = page ? resolve(selection.doc, { t: "ref", num: page.ref.num, gen: page.ref.gen }) : PDF_NULL;
    if (dict.t !== "dict" && !refused.some((record) => record.source === selection.source)) {
      refused.push(
        sourceRefusalRecord(selection.source, new Error(`it has no page ${selection.pageIndex + 1}`)),
      );
    }
  }
  if (refused.length > 0) {
    throw new PdfAssemblyRefusal(refusalMessage(refused), refused);
  }
  if (plan.pages.length === 0) {
    throw new PdfAssemblyRefusal("There are no pages to merge.");
  }
  let sourceBytes = 0;
  for (const src of sources.values()) {
    sourceBytes += src.doc.bytes.byteLength;
  }
  const limit = mergeLimitRefusal(plan.pages.length, sourceBytes);
  if (limit) {
    throw new PdfAssemblyRefusal(limit);
  }
}

/** Writes the plan's pages into one new PDF (see the section header). Throws
 * PdfAssemblyRefusal before copying anything when the plan is refused, and
 * PdfAssemblyCancelled when `hooks.cancelled` stops it at a page boundary;
 * neither leaves anything behind. Never mutates the plan or its sources. */
export async function assemblePdf(plan: PdfAssemblyPlan, hooks: PdfAssemblyHooks = {}): Promise<PdfAssemblyResult> {
  // Sources by identity, in first-appearance order.
  const sources = new Map<PdfDocument, MergeSource>();
  for (const selection of plan.pages) {
    if (!sources.has(selection.doc)) {
      sources.set(selection.doc, mergeSource(selection.source, selection.doc));
    }
  }
  refusePlan(plan, sources);
  const between = hooks.between ?? yieldToMainThread;
  const dropped: PdfDropRecord[] = [];
  const total = plan.pages.length;

  // [0] is the free head of the xref; 1 is the catalog, 2 the page tree,
  // then one number per output page — known up front, so a link can point
  // at a page copied after it.
  const objects: PdfValue[] = [PDF_NULL, PDF_NULL, PDF_NULL];
  const allocate = (): number => objects.push(PDF_NULL) - 1;
  const pageOut = plan.pages.map(() => allocate());
  plan.pages.forEach((selection, index) => {
    const src = sources.get(selection.doc) as MergeSource;
    if (!src.firstOutPage.has(selection.pageIndex)) {
      src.firstOutPage.set(selection.pageIndex, pageOut[index]);
    }
  });
  const taken = new Set<string>();
  for (const src of sources.values()) {
    nameDestinations(src, taken, dropped);
  }

  const NO_OWNED = new Map<number, number>();
  /** Rewrites a value of `src` into output terms. `owned` holds the
   * page-owned copies (annotations) of the page being copied, `page` that
   * page's source number and output number. */
  const rewriterFor =
    (src: MergeSource, owned: Map<number, number>, page?: { num: number; out: number }) =>
    (value: PdfValue): PdfValue =>
      rewritePdfRefs(value, (ref, key) => {
        if (key === "Parent") {
          const parent = owned.get(ref.num) ?? src.shared.get(ref.num);
          return parent === undefined ? undefined : outRef(parent);
        }
        if (page && ref.num === page.num) {
          return outRef(page.out);
        }
        const own = owned.get(ref.num);
        if (own !== undefined) {
          return outRef(own);
        }
        const pageIndex = src.pageIndexByNum.get(ref.num);
        if (pageIndex !== undefined) {
          const out = src.firstOutPage.get(pageIndex);
          return out === undefined ? PDF_NULL : outRef(out);
        }
        const out = src.shared.get(ref.num);
        return out === undefined ? PDF_NULL : outRef(out);
      });
  /** Copies (numbers, then rewrites) everything `roots` reach that `src`
   * has not copied yet. */
  const copyReachable = (src: MergeSource, roots: PdfValue[], owned = NO_OWNED, page?: { num: number; out: number }): void => {
    const fresh = pdfReachable(
      src.doc,
      roots,
      (num) => src.shared.has(num) || owned.has(num) || src.pageIndexByNum.has(num) || src.stop.has(num),
    );
    for (const num of fresh) {
      src.shared.set(num, allocate());
    }
    const rewrite = rewriterFor(src, owned, page);
    for (const num of fresh) {
      objects[src.shared.get(num) as number] = rewrite(src.doc.objects.get(num) ?? PDF_NULL);
    }
  };
  /** A source value in output terms, copying what it reaches. */
  const copyValue = (src: MergeSource, value: PdfValue): PdfValue => {
    copyReachable(src, [value]);
    return rewriterFor(src, NO_OWNED)(value);
  };

  hooks.onProgress?.(0, total);
  for (const [index, selection] of plan.pages.entries()) {
    if (index > 0 && index % MERGE_YIELD_PAGES === 0) {
      await between();
    }
    if (hooks.cancelled?.()) {
      throw new PdfAssemblyCancelled();
    }
    const src = sources.get(selection.doc) as MergeSource;
    const { doc } = src;
    const page = doc.pages[selection.pageIndex];
    const source = resolve(doc, { t: "ref", num: page.ref.num, gen: page.ref.gen }) as { t: "dict"; v: Map<string, PdfValue> };
    const pageNumber = selection.pageIndex + 1;

    // The page dictionary, materialised: inherited attributes onto the page
    // itself, the rotation added, the parent left to the new tree.
    const working = new Map(source.v);
    working.delete("Parent");
    for (const key of ["Resources", "MediaBox", "CropBox"]) {
      const value = inheritedValue(doc, page, key);
      if (value !== undefined) {
        working.set(key, value);
      }
    }
    if (!working.has("Resources")) {
      working.set("Resources", dictValue([]));
    }
    if (!working.has("MediaBox")) {
      // What the reader measured the page by when no box was readable.
      working.set("MediaBox", arrayValue(DEFAULT_MEDIA_BOX.map((v) => ({ t: "num", v }))));
    }
    working.set("Rotate", { t: "num", v: normaliseRotation(page.rotate + (selection.rotateBy ?? 0)) });

    // Its annotations: page-owned copies, an internal link remapped (or, to
    // a page not in the output, made inert), a signature's bytes left out.
    const owned = new Map<number, number>();
    const ownedValues: Array<[number, PdfValue]> = [];
    const annots: PdfValue[] = [];
    const annotsValue = resolve(doc, source.v.get("Annots") ?? PDF_NULL);
    for (const entry of annotsValue.t === "array" ? annotsValue.v : []) {
      const annotValue = resolve(doc, entry);
      if (annotValue.t !== "dict") {
        continue; // a broken entry carries nothing
      }
      const annot = new Map(annotValue.v);
      const subtype = resolve(doc, annot.get("Subtype") ?? PDF_NULL);
      const target = internalDest(doc, annotValue.v);
      if (target) {
        const verdict = judgeDest(src, target.dest);
        if (!verdict.keep) {
          const isLink = isName(subtype, "Link");
          dropped.push({
            kind: "annotation",
            source: src.key,
            page: pageNumber,
            reason: isLink
              ? `A link on page ${pageNumber} pointed at ${verdict.where}, so it was removed.`
              : `A ${subtype.t === "name" ? subtype.v : "page"} annotation on page ${pageNumber} pointed at ${verdict.where}, so its action was removed.`,
          });
          if (isLink) {
            continue;
          }
          annot.delete("Dest");
          annot.delete("A");
        } else if (verdict.rename !== undefined) {
          renameDest(annot, target, verdict.rename);
        }
      }
      // A filled form keeps its look (/AP comes across); a signature's bytes
      // do not — a merge invalidates the signature anyway.
      if (isName(subtype, "Widget") && isName(resolve(doc, annot.get("FT") ?? PDF_NULL), "Sig")) {
        annot.delete("V");
      }
      if (entry.t === "ref") {
        if (!owned.has(entry.num)) {
          owned.set(entry.num, allocate());
          ownedValues.push([entry.num, { t: "dict", v: annot }]);
        }
        annots.push(entry);
      } else {
        annots.push({ t: "dict", v: annot });
      }
    }
    if (annots.length > 0) {
      working.set("Annots", arrayValue(annots));
    } else {
      working.delete("Annots");
    }

    const scope = { num: page.ref.num, out: pageOut[index] };
    const pageValue: PdfValue = { t: "dict", v: working };
    copyReachable(src, [pageValue, ...ownedValues.map(([, value]) => value)], owned, scope);
    const rewrite = rewriterFor(src, owned, scope);
    for (const [num, value] of ownedValues) {
      objects[owned.get(num) as number] = rewrite(value);
    }
    const copied = rewrite(pageValue) as { t: "dict"; v: Map<string, PdfValue> };
    const pageDict = new Map<string, PdfValue>([
      ["Type", nameValue("Page")],
      ["Parent", outRef(PAGES_NUM)],
    ]);
    for (const [key, value] of copied.v) {
      if (key !== "Type" && key !== "Parent") {
        pageDict.set(key, value);
      }
    }
    objects[pageOut[index]] = { t: "dict", v: pageDict };
    hooks.onProgress?.(index + 1, total);
  }
  if (hooks.cancelled?.()) {
    throw new PdfAssemblyCancelled();
  }

  const catalog = new Map<string, PdfValue>([
    ["Type", nameValue("Catalog")],
    ["Pages", outRef(PAGES_NUM)],
  ]);

  // Outlines: one top-level entry per source, in first-appearance order,
  // pointing at its first output page, its own tree (filtered) beneath.
  if (plan.outline) {
    const tops: OutlineNode[] = [];
    for (const src of sources.values()) {
      const firstPage = Math.min(...src.firstOutPage.values());
      tops.push({
        title: textStringValue(src.key),
        link: [["Dest", arrayValue([outRef(firstPage), nameValue("Fit")])]],
        extra: [],
        open: true,
        children: readOutline(src, dropped),
        src: null,
      });
    }
    const root = allocate();
    const writeItems = (nodes: OutlineNode[], parent: number): { first: number; last: number } => {
      const nums = nodes.map(() => allocate());
      nodes.forEach((node, i) => {
        const dict = new Map<string, PdfValue>([
          ["Title", node.title],
          ["Parent", outRef(parent)],
        ]);
        if (i > 0) {
          dict.set("Prev", outRef(nums[i - 1]));
        }
        if (i + 1 < nums.length) {
          dict.set("Next", outRef(nums[i + 1]));
        }
        if (node.children.length > 0) {
          const kids = writeItems(node.children, nums[i]);
          dict.set("First", outRef(kids.first));
          dict.set("Last", outRef(kids.last));
          // Positive: open, showing this many; negative: closed.
          dict.set("Count", { t: "num", v: node.open ? openCount(node) : -openCount(node) });
        }
        for (const [key, value] of [...node.link, ...node.extra]) {
          dict.set(key, node.src ? copyValue(node.src, value) : value);
        }
        objects[nums[i]] = { t: "dict", v: dict };
      });
      return { first: nums[0], last: nums[nums.length - 1] };
    };
    const top = writeItems(tops, root);
    const visible = tops.reduce((sum, node) => sum + 1 + (node.open ? openCount(node) : 0), 0);
    objects[root] = dictValue([
      ["Type", nameValue("Outlines")],
      ["First", outRef(top.first)],
      ["Last", outRef(top.last)],
      ["Count", { t: "num", v: visible }],
    ]);
    catalog.set("Outlines", outRef(root));
  } else {
    for (const src of sources.values()) {
      const outlines = resolve(src.doc, src.catalog.get("Outlines") ?? PDF_NULL);
      if (outlines.t === "dict" && outlines.v.has("First")) {
        dropped.push({ kind: "outlineEntry", source: src.key, reason: "Its bookmarks were left out: this merge has no outline." });
      }
    }
  }

  // The kept named destinations, as one name tree (keys in byte order).
  const destEntries: Array<[string, PdfValue]> = [];
  for (const src of sources.values()) {
    for (const entry of src.named.values()) {
      if (entry.outName !== undefined) {
        destEntries.push([entry.outName, copyValue(src, entry.value)]);
      }
    }
  }
  if (destEntries.length > 0) {
    destEntries.sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));
    const names = destEntries.flatMap(([name, value]): PdfValue[] => [{ t: "str", v: latin1Bytes(name) }, value]);
    catalog.set("Names", dictValue([["Dests", dictValue([["Names", arrayValue(names)]])]]));
  }

  // Optional content: /OCGs concatenated, /D's ON and OFF unioned (a source
  // whose base state is OFF contributes every group it does not turn ON),
  // its /Order, /RBGroups and /Locked concatenated.
  const groups: PdfValue[] = [];
  const config = { ON: [] as PdfValue[], OFF: [] as PdfValue[], Order: [] as PdfValue[], RBGroups: [] as PdfValue[], Locked: [] as PdfValue[] };
  for (const src of sources.values()) {
    const { doc } = src;
    const props = resolve(doc, src.catalog.get("OCProperties") ?? PDF_NULL);
    if (props.t !== "dict") {
      continue;
    }
    const listOf = (holder: PdfValue, key: string): PdfValue[] => {
      const value = holder.t === "dict" ? resolve(doc, holder.v.get(key) ?? PDF_NULL) : PDF_NULL;
      return value.t === "array" ? value.v : [];
    };
    const defaults = resolve(doc, props.v.get("D") ?? PDF_NULL);
    const sourceGroups = listOf(props, "OCGs");
    groups.push(...sourceGroups.map((group) => copyValue(src, group)));
    const on = listOf(defaults, "ON");
    config.ON.push(...on.map((group) => copyValue(src, group)));
    config.OFF.push(...listOf(defaults, "OFF").map((group) => copyValue(src, group)));
    if (defaults.t === "dict" && isName(resolve(doc, defaults.v.get("BaseState") ?? PDF_NULL), "OFF")) {
      const turnedOn = new Set(on.filter((group) => group.t === "ref").map((group) => (group as { num: number }).num));
      config.OFF.push(
        ...sourceGroups
          .filter((group) => group.t === "ref" && !turnedOn.has(group.num))
          .map((group) => copyValue(src, group)),
      );
    }
    for (const key of ["Order", "RBGroups", "Locked"] as const) {
      config[key].push(...listOf(defaults, key).map((value) => copyValue(src, value)));
    }
  }
  if (groups.length > 0) {
    /** References in order, each once (a union). */
    const union = (values: PdfValue[]): PdfValue[] => {
      const seen = new Set<number>();
      return values.filter((value) => value.t === "ref" && !seen.has(value.num) && seen.add(value.num));
    };
    const defaults: Array<[string, PdfValue]> = [];
    for (const [key, values] of Object.entries(config)) {
      const list = key === "ON" || key === "OFF" ? union(values) : values;
      if (list.length > 0) {
        defaults.push([key, arrayValue(list)]);
      }
    }
    catalog.set("OCProperties", dictValue([["OCGs", arrayValue(union(groups))], ["D", dictValue(defaults)]]));
  }

  for (const src of sources.values()) {
    catalogDrops(src, total, dropped);
  }

  objects[CATALOG_NUM] = { t: "dict", v: catalog };
  objects[PAGES_NUM] = dictValue([
    ["Type", nameValue("Pages")],
    ["Kids", arrayValue(pageOut.map(outRef))],
    ["Count", { t: "num", v: total }],
  ]);
  const info = new Map<string, PdfValue>([["Title", textStringValue(plan.title)]]);
  const author = plan.author?.trim();
  const subject = plan.subject?.trim();
  if (author) {
    info.set("Author", textStringValue(author));
  }
  if (subject) {
    info.set("Subject", textStringValue(subject));
  }
  info.set("Creator", textStringValue(PDF_PRODUCER));
  info.set("Producer", textStringValue(PDF_PRODUCER));
  if (plan.modified) {
    info.set("CreationDate", dateStringValue(plan.modified));
    info.set("ModDate", dateStringValue(plan.modified));
  }
  const infoNum = allocate();
  objects[infoNum] = { t: "dict", v: info };
  const trailer = new Map<string, PdfValue>([
    ["Size", { t: "num", v: objects.length }],
    ["Root", outRef(CATALOG_NUM)],
    ["Info", outRef(infoNum)],
  ]);
  return { bytes: emitPdf(objects, trailer), dropped };
}

/** The page geometry a plan promises — each page's size as its source
 * measures it, and its /Rotate plus rotateBy. */
export function plannedPdfPages(
  plan: Pick<PdfAssemblyPlan, "pages">,
): Array<{ widthPt: number; heightPt: number; rotate: 0 | 90 | 180 | 270 }> {
  return plan.pages.map((selection) => {
    const page = selection.doc.pages[selection.pageIndex];
    return {
      widthPt: page?.widthPt ?? 0,
      heightPt: page?.heightPt ?? 0,
      rotate: normaliseRotation((page?.rotate ?? 0) + (selection.rotateBy ?? 0)),
    };
  });
}

/** The self-check every produced file is held to before it is kept: its
 * bytes parse back, to the planned page count, each page at its planned size
 * and rotation. Throws a user-facing Error otherwise — nothing built from
 * such bytes may be stored or downloaded. */
export async function checkAssembledPdf(
  bytes: Uint8Array,
  expected: ReturnType<typeof plannedPdfPages>,
  parse: (bytes: Uint8Array) => Promise<PdfDocument> = parsePdf,
): Promise<PdfDocument> {
  let doc: PdfDocument;
  try {
    doc = await parse(bytes);
  } catch (cause) {
    const reason = cause instanceof Error && cause.message ? cause.message : String(cause);
    throw new Error(`the merged PDF does not read back (${reason}), so nothing was kept.`, { cause });
  }
  if (doc.pages.length !== expected.length) {
    throw new Error(`the merged PDF has ${doc.pages.length} pages instead of ${expected.length}, so nothing was kept.`);
  }
  const describe = (page: { widthPt: number; heightPt: number; rotate: number }): string =>
    `${Math.round(page.widthPt)} × ${Math.round(page.heightPt)} pt, rotated ${page.rotate}°`;
  expected.forEach((want, index) => {
    const got = doc.pages[index];
    if (Math.abs(got.widthPt - want.widthPt) > 0.01 || Math.abs(got.heightPt - want.heightPt) > 0.01 || got.rotate !== want.rotate) {
      throw new Error(`page ${index + 1} of the merged PDF is ${describe(got)} instead of ${describe(want)}, so nothing was kept.`);
    }
  });
  return doc;
}
