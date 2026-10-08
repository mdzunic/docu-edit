// DOM-free builders for the SPEC-107 media fixtures — the same pattern as
// presentation/fixtures/build.ts: the checked-in e2e fixtures are generated
// from here by scripts/build-media-fixtures.ts, and the unit specs build the
// same bytes, so the two can never drift apart. Deflate goes through the
// platform's CompressionStream (zlib format, what PDF's /FlateDecode and
// PNG's IDAT both carry), so this module stays importable from core-only
// territory and runs identically under Node and the browser.
//
// spec107-pages.pdf is NOT built here: the fixture script writes it with the
// app's own buildPdf (src/lib/pdf/buildPdf.ts), which this module cannot
// import — the media layer imports only its siblings and core. The unit
// specs read that checked-in file directly.
//
// Production notes (AC: fixtures are checked in "with a note recording how
// each was produced"; scripts/build-media-fixtures.ts restates them):
// - spec107-shot.png     — buildPngFixtureBytes: a real 64×40 PNG written
//                          chunk by chunk (IHDR/IDAT/IEND with CRCs), RGB
//                          gradient scanlines deflated via CompressionStream.
// - spec107-photo.jpg    — buildJpegFixtureBytes: a known-good minimal
//                          baseline JPEG (1×1), embedded as base64 — this
//                          repo has no JPEG encoder and needs none; the
//                          fixture exists to exercise detection and
//                          byte-fidelity, not decoding.
// - spec107-anim.gif     — buildGifFixtureBytes: a real 2-frame animated
//                          GIF89a (8×8), LZW-coded with the classic
//                          "clear-code per symbol" trick so no compressor is
//                          needed.
// - spec107-mark.svg     — buildSvgFixtureBytes: plain XML with a leading
//                          comment and doctype, exercising the structural
//                          sniff.
// - spec107-shot.webp    — buildWebpFixtureBytes: a RIFF/WEBP container with
//                          a correct VP8L size header. The lossless payload
//                          is minimal, not guaranteed browser-decodable —
//                          the fixture exercises detection, dimensions and
//                          byte fidelity (nothing in this spec rasterises
//                          it).
// - spec107-xrefstream.pdf — buildXrefStreamPdfBytes: a PDF 1.5 file whose
//                          catalog and /Info live in an object stream and
//                          whose xref is a cross-reference STREAM with PNG
//                          predictor (/Predictor 12, /W [1 2 1]).
// - spec107-encrypted.pdf — buildEncryptedPdfBytes: a classic-xref file
//                          whose trailer carries /Encrypt (a standard
//                          security-handler stub); readers must refuse it
//                          without touching the objects.
import { bytesToBase64, base64ToBytes } from "../../core/packageParts";

export { base64ToBytes, bytesToBase64 };

function ascii(text: string): Uint8Array {
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i += 1) {
    out[i] = text.charCodeAt(i) & 0xff;
  }
  return out;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const out = new Uint8Array(total);
  let cursor = 0;
  for (const part of parts) {
    out.set(part, cursor);
    cursor += part.length;
  }
  return out;
}

/** zlib-format deflate via the platform's CompressionStream. */
export async function deflate(data: Uint8Array): Promise<Uint8Array> {
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(data);
      controller.close();
    },
  }).pipeThrough(new CompressionStream("deflate"));
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
  return concat(chunks);
}

// --- PNG ---------------------------------------------------------------------

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c;
  }
  return table;
})();

function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function u32be(value: number): Uint8Array {
  return new Uint8Array([(value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff]);
}

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const body = concat([ascii(type), data]);
  return concat([u32be(data.length), body, u32be(crc32(body))]);
}

export const PNG_FIXTURE_WIDTH = 64;
export const PNG_FIXTURE_HEIGHT = 40;

/** A real, decodable 64×40 RGB PNG with a diagonal gradient — big enough
 * that a downscaled library thumbnail is visibly an image. */
export async function buildPngFixtureBytes(): Promise<Uint8Array> {
  const width = PNG_FIXTURE_WIDTH;
  const height = PNG_FIXTURE_HEIGHT;
  const raw = new Uint8Array(height * (1 + width * 3));
  for (let y = 0; y < height; y += 1) {
    const row = y * (1 + width * 3);
    raw[row] = 0; // filter: None
    for (let x = 0; x < width; x += 1) {
      const px = row + 1 + x * 3;
      raw[px] = Math.round((x / (width - 1)) * 255);
      raw[px + 1] = Math.round((y / (height - 1)) * 255);
      raw[px + 2] = 96;
    }
  }
  const ihdr = concat([u32be(width), u32be(height), new Uint8Array([8, 2, 0, 0, 0])]);
  return concat([
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", await deflate(raw)),
    pngChunk("IEND", new Uint8Array(0)),
  ]);
}

// --- JPEG --------------------------------------------------------------------

// A minimal known-good baseline JPEG (1×1) — see the production note above.
const JPEG_1X1_BASE64 =
  "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0a" +
  "HBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIy" +
  "MjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAABAAEDASIA" +
  "AhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQA" +
  "AAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3" +
  "ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWm" +
  "p6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/9oACAEB" +
  "AAA/AL+B/9k=";

export function buildJpegFixtureBytes(): Uint8Array {
  return base64ToBytes(JPEG_1X1_BASE64);
}

// --- GIF ---------------------------------------------------------------------

/** LZW-codes one frame's pixels the "no compression" way: a clear code
 * before every pixel code keeps the dictionary at its initial size, so no
 * real LZW state is needed and the output is still valid GIF data. */
function gifPixelData(minCodeSize: number, pixels: number[]): Uint8Array {
  const clearCode = 1 << minCodeSize;
  const endCode = clearCode + 1;
  const codeSize = minCodeSize + 1;
  const bits: number[] = [];
  const pushCode = (code: number) => {
    for (let i = 0; i < codeSize; i += 1) {
      bits.push((code >> i) & 1);
    }
  };
  pushCode(clearCode);
  for (const pixel of pixels) {
    pushCode(pixel);
    pushCode(clearCode);
  }
  pushCode(endCode);
  const bytes: number[] = [];
  for (let i = 0; i < bits.length; i += 8) {
    let byte = 0;
    for (let b = 0; b < 8 && i + b < bits.length; b += 1) {
      byte |= bits[i + b] << b;
    }
    bytes.push(byte);
  }
  // Sub-blocks: length prefix (≤255) per block, then a 0 terminator.
  const out: number[] = [minCodeSize];
  for (let i = 0; i < bytes.length; i += 255) {
    const block = bytes.slice(i, i + 255);
    out.push(block.length, ...block);
  }
  out.push(0);
  return new Uint8Array(out);
}

export const GIF_FIXTURE_SIZE = 8;

/** A real 2-frame animated GIF89a (8×8), looping, 4 colours. */
export function buildGifFixtureBytes(): Uint8Array {
  const size = GIF_FIXTURE_SIZE;
  const parts: Uint8Array[] = [];
  parts.push(ascii("GIF89a"));
  // Logical screen: 8×8, global colour table of 4 entries.
  parts.push(new Uint8Array([size, 0, size, 0, 0x91, 0, 0]));
  parts.push(new Uint8Array([
    0x1f, 0x77, 0xb4, // blue
    0xff, 0x7f, 0x0e, // orange
    0x2c, 0xa0, 0x2c, // green
    0xff, 0xff, 0xff, // white
  ]));
  // NETSCAPE2.0 loop-forever extension.
  parts.push(new Uint8Array([0x21, 0xff, 0x0b]));
  parts.push(ascii("NETSCAPE2.0"));
  parts.push(new Uint8Array([3, 1, 0, 0, 0]));
  const frame = (colorA: number, colorB: number): Uint8Array => {
    const pixels: number[] = [];
    for (let y = 0; y < size; y += 1) {
      for (let x = 0; x < size; x += 1) {
        pixels.push((x + y) % 2 === 0 ? colorA : colorB);
      }
    }
    return concat([
      // Graphic control: 30 cs delay.
      new Uint8Array([0x21, 0xf9, 4, 0, 30, 0, 0, 0]),
      // Image descriptor: full frame, no local colour table.
      new Uint8Array([0x2c, 0, 0, 0, 0, size, 0, size, 0, 0]),
      gifPixelData(2, pixels),
    ]);
  };
  parts.push(frame(0, 1));
  parts.push(frame(2, 3));
  parts.push(new Uint8Array([0x3b])); // trailer
  return concat(parts);
}

// --- SVG ---------------------------------------------------------------------

export const SVG_FIXTURE_TEXT = `<!-- SPEC-107 fixture: a mark with a doctype and this comment before the root -->
<!DOCTYPE svg>
<svg xmlns="http://www.w3.org/2000/svg" width="96" height="72" viewBox="0 0 96 72">
  <rect x="4" y="4" width="88" height="64" rx="8" fill="#1f77b4" />
  <circle cx="48" cy="36" r="20" fill="#ff7f0e" />
</svg>
`;

export function buildSvgFixtureBytes(): Uint8Array {
  return new TextEncoder().encode(SVG_FIXTURE_TEXT);
}

// --- WebP --------------------------------------------------------------------

export const WEBP_FIXTURE_WIDTH = 48;
export const WEBP_FIXTURE_HEIGHT = 30;

function u32le(value: number): Uint8Array {
  return new Uint8Array([value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff]);
}

/** A RIFF/WEBP container with a correct VP8L size header (48×30). See the
 * production note: structural, not guaranteed browser-decodable. */
export function buildWebpFixtureBytes(): Uint8Array {
  const packed =
    (WEBP_FIXTURE_WIDTH - 1) | ((WEBP_FIXTURE_HEIGHT - 1) << 14); // + alpha 0, version 0
  const payload = concat([
    new Uint8Array([0x2f]),
    u32le(packed),
    // Minimal filler bitstream bytes.
    new Uint8Array([0x00, 0x00, 0x00, 0x00, 0x00, 0x00]),
  ]);
  const chunk = concat([
    ascii("VP8L"),
    u32le(payload.length),
    payload,
    payload.length % 2 === 1 ? new Uint8Array([0]) : new Uint8Array(0),
  ]);
  return concat([ascii("RIFF"), u32le(4 + chunk.length), ascii("WEBP"), chunk]);
}

// --- PDF ---------------------------------------------------------------------

/** Assembles numbered objects into a classic-xref PDF file. Each entry in
 * `objects` is the object BODY (without `N 0 obj`/`endobj`); object numbers
 * are 1-based positions in the array. */
export function assembleClassicPdf(
  version: string,
  objects: Array<string | Uint8Array>,
  trailerExtra: string,
): Uint8Array {
  const parts: Uint8Array[] = [];
  let position = 0;
  const push = (data: string | Uint8Array): void => {
    const bytes = typeof data === "string" ? ascii(data) : data;
    parts.push(bytes);
    position += bytes.length;
  };
  push(`%PDF-${version}\n%âãÏÓ\n`);
  const offsets: number[] = [];
  objects.forEach((body, index) => {
    offsets.push(position);
    push(`${index + 1} 0 obj\n`);
    push(body);
    push("\nendobj\n");
  });
  const xrefStart = position;
  let xref = `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets) {
    xref += `${String(offset).padStart(10, "0")} 00000 n \n`;
  }
  xref += `trailer\n<< /Size ${objects.length + 1} ${trailerExtra} >>\n`;
  xref += `startxref\n${xrefStart}\n%%EOF\n`;
  push(xref);
  return concat(parts);
}

/** spec107-xrefstream.pdf: PDF 1.5 — catalog and /Info live in an object
 * stream, and the xref is a cross-reference STREAM using the PNG Up
 * predictor (/Predictor 12, /W [1 2 1]). */
export async function buildXrefStreamPdfBytes(): Promise<Uint8Array> {
  const parts: Uint8Array[] = [];
  let position = 0;
  const push = (data: string | Uint8Array): void => {
    const bytes = typeof data === "string" ? ascii(data) : data;
    parts.push(bytes);
    position += bytes.length;
  };
  push("%PDF-1.5\n%âãÏÓ\n");

  // Objects 1 (catalog) and 5 (info) live inside object stream 4.
  const contained = [
    { num: 1, body: "<< /Type /Catalog /Pages 2 0 R >>" },
    { num: 5, body: "<< /Title (SPEC-107 xref stream) /Producer (docu-edit) >>" },
  ];
  let header = "";
  let bodies = "";
  for (const entry of contained) {
    header += `${entry.num} ${bodies.length} `;
    bodies += `${entry.body}\n`;
  }
  const objStmContent = ascii(header + bodies);
  const objStmData = await deflate(objStmContent);

  const offsets = new Map<number, number>();
  offsets.set(2, position);
  push("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 612 792] >>\nendobj\n");
  offsets.set(3, position);
  push("3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 400] /Rotate 180 >>\nendobj\n");
  offsets.set(4, position);
  push(
    `4 0 obj\n<< /Type /ObjStm /N ${contained.length} /First ${header.length} ` +
      `/Filter /FlateDecode /Length ${objStmData.length} >>\nstream\n`,
  );
  push(objStmData);
  push("\nendstream\nendobj\n");

  // The xref stream itself is object 6, at the offset startxref points to.
  const xrefStreamOffset = position;
  offsets.set(6, xrefStreamOffset);
  // Rows: [type(1) field2(2) field3(1)] for objects 0..6.
  const rows: number[][] = [
    [0, 0, 0], // 0: free
    [2, 4, 0], // 1: in object stream 4, index 0
    [1, offsets.get(2) ?? 0, 0],
    [1, offsets.get(3) ?? 0, 0],
    [1, offsets.get(4) ?? 0, 0],
    [2, 4, 1], // 5: in object stream 4, index 1
    [1, xrefStreamOffset, 0],
  ];
  const rowBytes = 4; // W [1 2 1]
  const unfiltered = new Uint8Array(rows.length * rowBytes);
  rows.forEach((row, i) => {
    unfiltered[i * rowBytes] = row[0];
    unfiltered[i * rowBytes + 1] = (row[1] >>> 8) & 0xff;
    unfiltered[i * rowBytes + 2] = row[1] & 0xff;
    unfiltered[i * rowBytes + 3] = row[2];
  });
  // PNG Up filter (type 2) per row, matching /Predictor 12.
  const filtered = new Uint8Array(rows.length * (rowBytes + 1));
  const prior = new Uint8Array(rowBytes);
  rows.forEach((_, i) => {
    filtered[i * (rowBytes + 1)] = 2;
    for (let b = 0; b < rowBytes; b += 1) {
      const raw = unfiltered[i * rowBytes + b];
      filtered[i * (rowBytes + 1) + 1 + b] = (raw - prior[b]) & 0xff;
      prior[b] = raw;
    }
  });
  const xrefData = await deflate(filtered);
  push(
    `6 0 obj\n<< /Type /XRef /Size 7 /W [1 2 1] /Index [0 7] /Root 1 0 R /Info 5 0 R ` +
      `/Filter /FlateDecode /DecodeParms << /Predictor 12 /Columns 4 >> /Length ${xrefData.length} >>\nstream\n`,
  );
  push(xrefData);
  push("\nendstream\nendobj\n");
  push(`startxref\n${xrefStreamOffset}\n%%EOF\n`);
  return concat(parts);
}

/** spec107-encrypted.pdf: a classic-xref file whose trailer names an
 * /Encrypt dictionary — refused with PdfEncryptedError, never parsed. */
export function buildEncryptedPdfBytes(): Uint8Array {
  return assembleClassicPdf(
    "1.7",
    [
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 612 792] >>",
      "<< /Type /Page /Parent 2 0 R >>",
      "<< /Filter /Standard /V 1 /R 2 /O <68656c6c6f> /U <776f726c64> /P -44 >>",
    ],
    "/Root 1 0 R /Encrypt 4 0 R",
  );
}
