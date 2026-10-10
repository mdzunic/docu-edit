// Reads an image's intrinsic pixel dimensions straight from its encoded
// bytes (SPEC-024's "natural size", AC-21) — no `<img>`/`Image()` decode, so
// this works identically in the browser and under Vitest (jsdom does not
// implement image decoding). Supports the four formats "Insert image"
// accepts (PNG, JPEG, GIF, SVG) plus WebP for the library's media kind
// (SPEC-107); returns null for anything else or malformed data, which
// callers treat as "unknown", not a hard failure.
import { parseXml, XmlParseError } from "./xml";

export interface ImageDimensions {
  widthPx: number;
  heightPx: number;
}

function readUint32BE(bytes: Uint8Array, offset: number): number {
  return (bytes[offset] << 24 | bytes[offset + 1] << 16 | bytes[offset + 2] << 8 | bytes[offset + 3]) >>> 0;
}

function readUint16BE(bytes: Uint8Array, offset: number): number {
  return (bytes[offset] << 8) | bytes[offset + 1];
}

function readUint16LE(bytes: Uint8Array, offset: number): number {
  return (bytes[offset + 1] << 8) | bytes[offset];
}

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

function decodePng(bytes: Uint8Array): ImageDimensions | null {
  if (bytes.length < 24 || !PNG_SIGNATURE.every((byte, i) => bytes[i] === byte)) {
    return null;
  }
  // IHDR is always the first chunk, immediately after the signature: 4-byte
  // length, 4-byte "IHDR" type, then width/height as big-endian uint32s.
  const width = readUint32BE(bytes, 16);
  const height = readUint32BE(bytes, 20);
  return width > 0 && height > 0 ? { widthPx: width, heightPx: height } : null;
}

const GIF_SIGNATURES = ["GIF87a", "GIF89a"];

function decodeGif(bytes: Uint8Array): ImageDimensions | null {
  if (bytes.length < 10) {
    return null;
  }
  const header = new TextDecoder("ascii").decode(bytes.subarray(0, 6));
  if (!GIF_SIGNATURES.includes(header)) {
    return null;
  }
  // The logical screen descriptor follows the 6-byte signature: width/height
  // as little-endian uint16s.
  const width = readUint16LE(bytes, 6);
  const height = readUint16LE(bytes, 8);
  return width > 0 && height > 0 ? { widthPx: width, heightPx: height } : null;
}

// JPEG markers that carry a frame's dimensions (baseline/progressive/etc.
// start-of-frame variants) — the ones a JFIF/EXIF encoder actually emits.
const SOF_MARKERS = new Set([
  0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf,
]);

function decodeJpeg(bytes: Uint8Array): ImageDimensions | null {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) {
    return null;
  }
  let offset = 2;
  while (offset + 4 <= bytes.length) {
    if (bytes[offset] !== 0xff) {
      offset++;
      continue;
    }
    const marker = bytes[offset + 1];
    if (marker === 0xd8 || marker === 0xd9) {
      offset += 2;
      continue;
    }
    // Markers with no payload (0xD0-0xD7 restart markers, 0x01 TEM) carry no
    // length field.
    if (marker >= 0xd0 && marker <= 0xd7) {
      offset += 2;
      continue;
    }
    const segmentLength = readUint16BE(bytes, offset + 2);
    if (SOF_MARKERS.has(marker) && offset + 9 <= bytes.length) {
      const height = readUint16BE(bytes, offset + 5);
      const width = readUint16BE(bytes, offset + 7);
      return width > 0 && height > 0 ? { widthPx: width, heightPx: height } : null;
    }
    offset += 2 + segmentLength;
  }
  return null;
}

function readUint24LE(bytes: Uint8Array, offset: number): number {
  return bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16);
}

/** WebP (SPEC-107): a RIFF container with a `WEBP` form type whose first
 * chunk carries the dimensions in one of three layouts — `VP8 ` (lossy),
 * `VP8L` (lossless) or `VP8X` (extended). Null for anything malformed. */
function decodeWebp(bytes: Uint8Array): ImageDimensions | null {
  const ascii = (offset: number, length: number): string =>
    new TextDecoder("ascii").decode(bytes.subarray(offset, offset + length));
  if (bytes.length < 25 || ascii(0, 4) !== "RIFF" || ascii(8, 4) !== "WEBP") {
    return null;
  }
  const chunk = ascii(12, 4);
  if (chunk === "VP8 ") {
    // Lossy: a 3-byte frame tag, then the 9D 01 2A start code, then width
    // and height as 14-bit little-endian values.
    if (bytes.length < 30 || bytes[23] !== 0x9d || bytes[24] !== 0x01 || bytes[25] !== 0x2a) {
      return null;
    }
    const width = readUint16LE(bytes, 26) & 0x3fff;
    const height = readUint16LE(bytes, 28) & 0x3fff;
    return width > 0 && height > 0 ? { widthPx: width, heightPx: height } : null;
  }
  if (chunk === "VP8L") {
    // Lossless: a 0x2F signature byte, then width-1 and height-1 as
    // consecutive 14-bit fields of a little-endian 32-bit read, an alpha
    // bit, and a 3-bit version that must be 0.
    if (bytes[20] !== 0x2f || bytes[24] >> 5 !== 0) {
      return null;
    }
    const packed =
      bytes[21] | (bytes[22] << 8) | (bytes[23] << 16) | (bytes[24] << 24);
    const width = (packed & 0x3fff) + 1;
    const height = ((packed >> 14) & 0x3fff) + 1;
    return { widthPx: width, heightPx: height };
  }
  if (chunk === "VP8X") {
    // Extended: flags + reserved, then canvas width-1 and height-1 as
    // 24-bit little-endian values.
    if (bytes.length < 30) {
      return null;
    }
    const width = readUint24LE(bytes, 24) + 1;
    const height = readUint24LE(bytes, 27) + 1;
    return width > 0 && height > 0 ? { widthPx: width, heightPx: height } : null;
  }
  return null;
}

const PX_PER_INCH = 96;
const CM_PER_INCH = 2.54;

// SVG length units expressed as pixels at the SVG spec's fixed 96 DPI (the
// same reference used everywhere else in this module).
const SVG_UNIT_TO_PX: Record<string, number> = {
  px: 1,
  in: PX_PER_INCH,
  cm: PX_PER_INCH / CM_PER_INCH,
  mm: PX_PER_INCH / CM_PER_INCH / 10,
  pt: PX_PER_INCH / 72,
  pc: PX_PER_INCH / 72 * 12,
};

function parseSvgLength(value: string | null): number | undefined {
  if (!value) {
    return undefined;
  }
  const match = /^(-?\d+(?:\.\d+)?)(px|in|cm|mm|pt|pc)?$/.exec(value.trim());
  if (!match) {
    return undefined;
  }
  const unit = match[2] ?? "px";
  return Number(match[1]) * (SVG_UNIT_TO_PX[unit] ?? 1);
}

function decodeSvg(bytes: Uint8Array): ImageDimensions | null {
  let text: string;
  try {
    text = new TextDecoder("utf-8").decode(bytes);
  } catch {
    return null;
  }
  let doc: Document;
  try {
    doc = parseXml(text);
  } catch (error) {
    if (error instanceof XmlParseError) {
      return null;
    }
    throw error;
  }
  const root = doc.documentElement;
  if (root.localName !== "svg") {
    return null;
  }
  const width = parseSvgLength(root.getAttribute("width"));
  const height = parseSvgLength(root.getAttribute("height"));
  if (width && height) {
    return { widthPx: width, heightPx: height };
  }
  // No explicit width/height: fall back to the viewBox's own units (SVG's
  // own default is "1 user unit = 1px" absent other sizing).
  const viewBox = root.getAttribute("viewBox");
  if (!viewBox) {
    return null;
  }
  const parts = viewBox.trim().split(/[\s,]+/).map(Number);
  if (parts.length !== 4 || parts.some((n) => !Number.isFinite(n))) {
    return null;
  }
  const [, , vbWidth, vbHeight] = parts;
  return vbWidth > 0 && vbHeight > 0 ? { widthPx: vbWidth, heightPx: vbHeight } : null;
}

/** Decodes an image's intrinsic pixel dimensions from its raw bytes, or null
 * when the format isn't recognised/supported or the data is malformed. Format
 * is sniffed from the bytes themselves (PNG/GIF/WebP signatures, JPEG's SOI
 * marker) except SVG, identified by `contentType` since it is plain XML with
 * no fixed magic bytes. */
export function decodeImageDimensions(bytes: Uint8Array, contentType: string): ImageDimensions | null {
  if (contentType === "image/svg+xml") {
    return decodeSvg(bytes);
  }
  return decodePng(bytes) ?? decodeGif(bytes) ?? decodeJpeg(bytes) ?? decodeWebp(bytes);
}

/** How many frames (image descriptors) a GIF holds, walking its block
 * structure rather than scanning for the 0x2C byte — the image editor
 * (SPEC-112) edits an animated GIF's first frame only and says so before it
 * starts. 0 for anything that is not a well-formed GIF; a truncated file
 * counts the frames seen before the cut. */
export function gifFrameCount(bytes: Uint8Array): number {
  if (decodeGif(bytes) === null || bytes.length < 13) {
    return 0;
  }
  let offset = 13;
  const packed = bytes[10];
  if (packed & 0x80) {
    offset += 3 * (1 << ((packed & 0x07) + 1));
  }
  const skipSubBlocks = (from: number): number => {
    let at = from;
    while (at < bytes.length && bytes[at] !== 0) {
      at += bytes[at] + 1;
    }
    return at + 1;
  };
  let frames = 0;
  while (offset < bytes.length) {
    const introducer = bytes[offset];
    if (introducer === 0x2c) {
      frames += 1;
      if (offset + 10 > bytes.length) {
        break;
      }
      const local = bytes[offset + 9];
      offset += 10;
      if (local & 0x80) {
        offset += 3 * (1 << ((local & 0x07) + 1));
      }
      // The LZW minimum code size, then the data sub-blocks.
      offset = skipSubBlocks(offset + 1);
    } else if (introducer === 0x21) {
      offset = skipSubBlocks(offset + 2);
    } else {
      break;
    }
  }
  return frames;
}
