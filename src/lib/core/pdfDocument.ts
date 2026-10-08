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
  const isCount = (value: PdfValue): value is { t: "num"; v: number } =>
    value.t === "num" && Number.isInteger(value.v) && value.v >= 0;
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
