// DOM-free builders for the SPEC-107 media fixtures — the same pattern as
// presentation/fixtures/build.ts: the checked-in e2e fixtures are generated
// from here by scripts/build-media-fixtures.ts, and the unit specs build the
// same bytes, so the two can never drift apart. Deflate goes through the
// platform's CompressionStream (zlib format, what PDF's /FlateDecode and
// PNG's IDAT both carry), so this module stays importable from core-only
// territory and runs identically under Node and the browser.
//
// spec107-pages.pdf is NOT built here: the fixture script writes it with the
// app's own buildPdf (src/lib/core/pdfBuild.ts since SPEC-104). The unit
// specs read that checked-in file directly.
//
// Production notes (AC: fixtures are checked in "with a note recording how
// each was produced"; scripts/build-media-fixtures.ts restates them):
// - spec107-shot.png     — buildPngFixtureBytes: a real 64×40 PNG written
//                          chunk by chunk (IHDR/IDAT/IEND with CRCs), RGB
//                          gradient scanlines deflated via CompressionStream.
// - spec107-photo.jpg    — buildJpegFixtureBytes: a real 160 × 120
//                          baseline JPEG of four flat colour quadrants,
//                          embedded as base64. This repo has no JPEG encoder:
//                          the bytes were produced once by Chromium's own
//                          canvas encoder (fillRect per quadrant, then
//                          toDataURL("image/jpeg", 0.9)). SPEC-107 used it
//                          for detection and byte fidelity at 1 × 1; SPEC-112
//                          needs an image a 16:9 crop and a rotation can be
//                          seen on, and pixels a test can sample.
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
// - spec111-forms.pdf    — buildFormsPdfBytes (SPEC-111): a filled form —
//                          widgets with appearance streams, a signature,
//                          document JavaScript, an open action, an attached
//                          file, page labels, optional content — over a
//                          two-level page tree that supplies inherited
//                          attributes. Its builder's comment lists it all.
// - spec111-outlines.pdf — buildOutlinesPdfBytes (SPEC-111): a tagged
//                          two-page PDF with bookmarks, named destinations,
//                          and internal and web links.
// - spec114-scanned.pdf  — buildScannedPdfBytes (SPEC-114): one page that
//                          paints an image XObject and shows no text — a
//                          scan with no text layer.
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

// A real 160 × 120 baseline JPEG — see the production note above. Four flat
// quadrants (top-left #c83c3c, top-right #3ca03c, bottom-left #3c50c8,
// bottom-right #808080) so a sampled pixel is a known colour after decoding.
const JPEG_PHOTO_BASE64 =
  "/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdC" +
  "IFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAA" +
  "AADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlk" +
  "ZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAA" +
  "ABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAA" +
  "AAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAA" +
  "AABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEA" +
  "AAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAA" +
  "ACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAMCAgMCAgMDAwMEAwME" +
  "BQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQME" +
  "BAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQU" +
  "FBQUFBT/wAARCAB4AKADASIAAhEBAxEB/8QAGAABAAMBAAAAAAAAAAAAAAAAAAQFBwb/xAAYEAEA" +
  "AwEAAAAAAAAAAAAAAAAAFWOi4f/EABkBAQADAQEAAAAAAAAAAAAAAAAFBggJB//EACIRAQAAAwgD" +
  "AAAAAAAAAAAAAAACBhUBBBZUoaLR0gVSU//aAAwDAQACEQMRAD8AzoB5+12AAAAAAAA60BDuagAA" +
  "AAAAADkgEw6VgAAAAAAAOtAQ7moAAAAAAAA5IBMOlYAAAAAAADrQEO5qAAAAAAAAOSFZM064TNOu" +
  "PUsDzBlt8HZu/HUvZnZH1WYrJmnXCZp1wwPMGW3wdjHUvZnZH1WYrJmnXCZp1wwPMGW3wdjHUvZn" +
  "ZH1WYrJmnXCZp1wwPMGW3wdjHUvZnZH1WYrJmnXCZp1wwPMGW3wdjHUvZnZH1aAJkfZkj7MqHSr5" +
  "6a2csWUTyHz1h5QxMj7MkfZkpV89NbOSieQ+esPKGJkfZkj7MlKvnprZyUTyHz1h5QxMj7MkfZkp" +
  "V89NbOSieQ+esPKGJkfZkj7MlKvnprZyUTyHz1h5ZAA3UnwAAAAAAAG1gMrrqAAAAAAAAxQBqhSg" +
  "AAAAAAAG1gMrrqAAAAAAAAxQBqhSgAAAAAAAG1gMrrqAAAAAAAAxQBqhSgAAAAAAAG1gMrrqAAAA" +
  "AAAA/9k=";

export function buildJpegFixtureBytes(): Uint8Array {
  return base64ToBytes(JPEG_PHOTO_BASE64);
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

// --- SPEC-111 merge fixtures ---------------------------------------------------

/** A stream object's body: `dict` (without /Length, which is added) then
 * the data, verbatim. */
function streamBody(dict: string, data: Uint8Array): Uint8Array {
  return concat([ascii(`<< ${dict} /Length ${data.length} >>\nstream\n`), data, ascii("\nendstream")]);
}

/** spec114-scanned.pdf (SPEC-114): what a scanner hands over — one US
 * Letter page that paints a single image XObject (a 16 × 16 grey gradient,
 * raw 8-bit samples, no filter) and has no text operators at all, no font
 * and no text layer. Library search must list it under "Not searched" as a
 * scanned image, never silently skip it. Synchronous and filter-free, so the
 * bytes are identical on every run. */
export function buildScannedPdfBytes(): Uint8Array {
  const side = 16;
  const samples = Uint8Array.from({ length: side * side }, (_, i) => Math.round(((i % side) + Math.floor(i / side)) * (255 / (2 * (side - 1)))));
  const objects: Array<string | Uint8Array> = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>",
    streamBody("", ascii("q 432 0 0 576 90 108 cm /Im1 Do Q\n")),
    streamBody(`/Type /XObject /Subtype /Image /Width ${side} /Height ${side} /ColorSpace /DeviceGray /BitsPerComponent 8`, samples),
    "<< /Title (Scanned receipt) /Producer (a scanner) >>",
  ];
  return assembleClassicPdf("1.4", objects, "/Root 1 0 R /Info 6 0 R");
}

/** spec111-forms.pdf: everything a merge must carry, or report, about a
 * filled form — two pages under a two-level page tree whose intermediate
 * node supplies /Resources (by reference), /CropBox and /Rotate 90 while the
 * root supplies /MediaBox (page 2 overrides /MediaBox and /Rotate); page 1
 * carries /BleedBox, /TrimBox and /ArtBox, a deflated content stream drawing
 * a DCT image (the JPEG fixture's bytes, raw), a form XObject in optional
 * content group "Layer two", a shading, a shading pattern, an /Indexed
 * colour space and an ExtGState; and three widgets with appearance streams —
 * a filled text field, a checkbox under a parent field, and a signature
 * field whose /V holds the signature bytes (<0123456789ABCDEF>). The catalog
 * has an /AcroForm (/SigFlags 3), document JavaScript, a JavaScript
 * /OpenAction, an attached file "data.csv", /PageLabels and /OCProperties
 * ("Layer one" ON, "Layer two" OFF). */
export async function buildFormsPdfBytes(): Promise<Uint8Array> {
  const jpeg = buildJpegFixtureBytes();
  const page1 = await deflate(
    ascii(
      "0.12 0.47 0.71 RG 2 w 36 36 540 720 re S\n" +
        "BT /F1 18 Tf 72 700 Td (SPEC-111 form) Tj ET\n" +
        "q 160 0 0 120 380 600 cm /Im1 Do Q\n" +
        "/OC /oc1 BDC q /GS1 gs 1 0 0 rg 72 640 200 20 re f Q EMC\n" +
        "q 380 520 m 540 520 l 540 580 l 380 580 l h W n /Sh1 sh Q\n" +
        "q /Pattern cs /P1 scn 380 440 160 60 re f Q\n" +
        "q /CS1 cs 1 sc 380 400 160 20 re f Q\n" +
        "q 1 0 0 1 380 340 cm /Fm1 Do Q\n",
    ),
  );
  const page2 = ascii("0.12 0.47 0.71 RG 2 w 36 36 428 428 re S\nBT /F1 18 Tf 72 400 Td (Page two) Tj ET\n");
  const objects: Array<string | Uint8Array> = [
    // 1 catalog
    "<< /Type /Catalog /Pages 2 0 R /AcroForm 20 0 R /Names << /JavaScript 23 0 R /EmbeddedFiles 25 0 R >> " +
      "/OpenAction 28 0 R /PageLabels << /Nums [0 << /S /r >>] >> " +
      "/OCProperties << /OCGs [12 0 R 13 0 R] /D << /Order [12 0 R 13 0 R] /ON [12 0 R] /OFF [13 0 R] >> >> >>",
    // 2 the page tree's root: /MediaBox
    "<< /Type /Pages /Kids [3 0 R] /Count 2 /MediaBox [0 0 612 792] >>",
    // 3 an intermediate node: /Resources (by reference), /CropBox, /Rotate
    "<< /Type /Pages /Parent 2 0 R /Kids [4 0 R 5 0 R] /Count 2 /Resources 6 0 R /CropBox [10 10 602 782] /Rotate 90 >>",
    // 4 page 1
    "<< /Type /Page /Parent 3 0 R /Contents 7 0 R /Annots [16 0 R 17 0 R 18 0 R] " +
      "/BleedBox [5 5 607 787] /TrimBox [15 15 597 777] /ArtBox [20 20 592 772] >>",
    // 5 page 2: its own /MediaBox and /Rotate
    "<< /Type /Page /Parent 3 0 R /MediaBox [0 0 500 500] /Rotate 0 /Contents 8 0 R >>",
    // 6 the shared resources
    "<< /Font << /F1 9 0 R >> /XObject << /Im1 10 0 R /Fm1 11 0 R >> /Properties << /oc1 12 0 R /oc2 13 0 R >> " +
      "/ColorSpace << /CS1 14 0 R >> /Shading << /Sh1 15 0 R >> /Pattern << /P1 29 0 R >> /ExtGState << /GS1 << /CA 0.5 /ca 0.5 >> >> >>",
    // 7, 8 content
    streamBody("/Filter /FlateDecode", page1),
    streamBody("", page2),
    // 9 a standard font
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
    // 10 a DCT image: the JPEG fixture's own bytes
    streamBody("/Type /XObject /Subtype /Image /Width 160 /Height 120 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode", jpeg),
    // 11 a form XObject in "Layer two"
    streamBody(
      "/Type /XObject /Subtype /Form /BBox [0 0 160 40] /Resources << /Font << /F1 9 0 R >> >> /OC 13 0 R",
      ascii("BT /F1 12 Tf 4 14 Td (Layer two) Tj ET"),
    ),
    // 12, 13 optional content groups
    "<< /Type /OCG /Name (Layer one) >>",
    "<< /Type /OCG /Name (Layer two) >>",
    // 14 an indexed colour space
    "[/Indexed /DeviceRGB 1 <FF00000000FF>]",
    // 15 an axial shading over a function
    "<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [380 0 540 0] /Function 30 0 R /Extend [true true] >>",
    // 16 a filled text field (field and widget in one dictionary)
    "<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /V (Ada Lovelace) /DA (/Helv 12 Tf 0 g) " +
      "/Rect [72 560 300 584] /F 4 /P 4 0 R /AP << /N 19 0 R >> >>",
    // 17 a checkbox widget under its parent field
    "<< /Type /Annot /Subtype /Widget /Parent 21 0 R /Rect [72 520 90 538] /F 4 /P 4 0 R /AS /Yes " +
      "/AP << /N << /Yes 31 0 R /Off 32 0 R >> >> >>",
    // 18 a signed signature field
    "<< /Type /Annot /Subtype /Widget /FT /Sig /T (signature) /V 22 0 R /Rect [72 440 300 480] /F 4 /P 4 0 R /AP << /N 33 0 R >> >>",
    // 19 the text field's appearance
    streamBody(
      "/Type /XObject /Subtype /Form /BBox [0 0 228 24] /Resources << /Font << /Helv 9 0 R >> >>",
      ascii("/Tx BMC BT /Helv 12 Tf 0 g 2 7 Td (Ada Lovelace) Tj ET EMC"),
    ),
    // 20 the form
    "<< /Fields [16 0 R 21 0 R 18 0 R] /SigFlags 3 /DA (/Helv 0 Tf 0 g) /DR << /Font << /Helv 9 0 R >> >> >>",
    // 21 the checkbox's parent field
    "<< /FT /Btn /T (agree) /V /Yes /Kids [17 0 R] >>",
    // 22 the signature value: its bytes must not travel
    "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [0 0 0 0] /Contents <0123456789ABCDEF> >>",
    // 23, 24 document JavaScript
    "<< /Names [(init) 24 0 R] >>",
    "<< /S /JavaScript /JS (app.alert\\('SPEC-111'\\);) >>",
    // 25, 26, 27 an attached file
    "<< /Names [(data.csv) 26 0 R] >>",
    "<< /Type /Filespec /F (data.csv) /UF (data.csv) /EF << /F 27 0 R >> >>",
    streamBody("/Type /EmbeddedFile /Subtype /text#2Fcsv", ascii("a,b\n1,2\n")),
    // 28 the open action
    "<< /S /JavaScript /JS (this.print\\(\\);) >>",
    // 29 a shading pattern
    "<< /Type /Pattern /PatternType 2 /Shading 15 0 R >>",
    // 30 the shading's function
    "<< /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >>",
    // 31, 32 the checkbox's appearances
    streamBody("/Type /XObject /Subtype /Form /BBox [0 0 18 18]", ascii("0 0 1 rg 3 3 12 12 re f")),
    streamBody("/Type /XObject /Subtype /Form /BBox [0 0 18 18]", ascii("")),
    // 33 the signature's appearance
    streamBody(
      "/Type /XObject /Subtype /Form /BBox [0 0 228 40] /Resources << /Font << /F1 9 0 R >> >>",
      ascii("BT /F1 10 Tf 4 16 Td (Signed: A. Lovelace) Tj ET"),
    ),
    // 34 info
    "<< /Title (SPEC-111 forms) /Producer (docu-edit) >>",
  ];
  return assembleClassicPdf("1.7", objects, "/Root 1 0 R /Info 34 0 R");
}

/** spec111-outlines.pdf: a tagged two-page PDF with bookmarks and links —
 * an A5 portrait page and a landscape 595 × 420 pt page, neither rotated,
 * sharing /Resources inherited from the page tree. The catalog has a
 * /StructTreeRoot (with /MarkInfo), an outline ("Chapter 1" → page 1 with
 * child "Section 1.1" → the named destination (chapter1); "Chapter 2" → page
 * 2 through a /GoTo action) and a /Dests name tree ((chapter1) → page 1,
 * (chapter2) → page 2). Page 1 links to page 2 by an explicit destination,
 * to (chapter2) by a /GoTo action, and to https://example.com/spec111 by a
 * /URI action; page 2 links back through /Dest (chapter1). */
export async function buildOutlinesPdfBytes(): Promise<Uint8Array> {
  const page1 = await deflate(
    ascii(
      "0.12 0.47 0.71 RG 2 w 36 36 348 523 re S\n" +
        "/H1 << /MCID 0 >> BDC BT /F1 24 Tf 72 500 Td (Chapter 1) Tj ET EMC\n" +
        "BT /F1 12 Tf 72 405 Td (Go to page 2) Tj ET\n" +
        "BT /F1 12 Tf 72 375 Td (Go to chapter 2) Tj ET\n" +
        "BT /F1 12 Tf 72 345 Td (example.com) Tj ET\n",
    ),
  );
  const page2 = ascii(
    "1 0.5 0.05 RG 2 w 36 36 523 348 re S\n" +
      "/H1 << /MCID 0 >> BDC BT /F1 24 Tf 72 330 Td (Chapter 2) Tj ET EMC\n" +
      "BT /F1 12 Tf 72 255 Td (Back to chapter 1) Tj ET\n",
  );
  const objects: Array<string | Uint8Array> = [
    // 1 catalog
    "<< /Type /Catalog /Pages 2 0 R /Outlines 13 0 R /Names << /Dests 17 0 R >> /StructTreeRoot 18 0 R " +
      "/MarkInfo << /Marked true >> /Lang (en) /PageMode /UseOutlines >>",
    // 2 the page tree: /Resources by reference
    "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 /Resources 5 0 R >>",
    // 3, 4 the pages
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 420 595] /Contents 7 0 R /Annots [9 0 R 10 0 R 11 0 R] /StructParents 0 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 420] /Contents 8 0 R /Annots [12 0 R] /StructParents 1 >>",
    // 5, 6 resources and a standard font
    "<< /Font << /F1 6 0 R >> >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
    // 7, 8 content
    streamBody("/Filter /FlateDecode", page1),
    streamBody("", page2),
    // 9 page 1 → page 2, an explicit destination
    "<< /Type /Annot /Subtype /Link /Rect [68 400 300 420] /Border [0 0 0] /P 3 0 R /Dest [4 0 R /XYZ 0 420 null] >>",
    // 10 page 1 → (chapter2), a /GoTo action
    "<< /Type /Annot /Subtype /Link /Rect [68 370 300 390] /Border [0 0 0] /P 3 0 R /A << /S /GoTo /D (chapter2) >> >>",
    // 11 page 1 → the web
    "<< /Type /Annot /Subtype /Link /Rect [68 340 300 360] /Border [0 0 0] /A << /S /URI /URI (https://example.com/spec111) >> >>",
    // 12 page 2 → (chapter1)
    "<< /Type /Annot /Subtype /Link /Rect [68 250 300 270] /Border [0 0 0] /P 4 0 R /Dest (chapter1) >>",
    // 13–16 the outline
    "<< /Type /Outlines /First 14 0 R /Last 16 0 R /Count 3 >>",
    "<< /Title (Chapter 1) /Parent 13 0 R /Next 16 0 R /First 15 0 R /Last 15 0 R /Count 1 /Dest [3 0 R /Fit] >>",
    "<< /Title (Section 1.1) /Parent 14 0 R /Dest (chapter1) >>",
    "<< /Title (Chapter 2) /Parent 13 0 R /Prev 14 0 R /A << /S /GoTo /D [4 0 R /Fit] >> >>",
    // 17 the named destinations
    "<< /Names [(chapter1) [3 0 R /XYZ 0 595 null] (chapter2) [4 0 R /XYZ 0 420 null]] >>",
    // 18–22 the structure tree
    "<< /Type /StructTreeRoot /K 19 0 R /ParentTree 22 0 R /ParentTreeNextKey 2 >>",
    "<< /Type /StructElem /S /Document /P 18 0 R /K [20 0 R 21 0 R] >>",
    "<< /Type /StructElem /S /H1 /P 19 0 R /Pg 3 0 R /K 0 >>",
    "<< /Type /StructElem /S /H1 /P 19 0 R /Pg 4 0 R /K 0 >>",
    "<< /Nums [0 [20 0 R] 1 [21 0 R]] >>",
    // 23 info
    "<< /Title (SPEC-111 outlines) /Producer (docu-edit) >>",
  ];
  return assembleClassicPdf("1.7", objects, "/Root 1 0 R /Info 23 0 R");
}
