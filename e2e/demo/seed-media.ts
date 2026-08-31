import { deflateSync } from "node:zlib";
import { writeFileSync } from "node:fs";

// Seed-file generators for the library tour (library.tour.ts). The repo's
// SPEC-107 media fixtures are deliberately minimal — a 1×1 test pixel, a
// 726-byte structural PDF — which is exactly right for parsers and exactly
// wrong for a product demo, where the cards would read "1 × 1 px" over a
// blank preview. These build presentable stand-ins at seed time instead:
// a real multi-page text PDF and a real-sized poster PNG, dependency-free.

// --- PNG -------------------------------------------------------------------

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const chunk = new Uint8Array(12 + data.length);
  const view = new DataView(chunk.buffer);
  view.setUint32(0, data.length);
  for (let i = 0; i < 4; i += 1) {
    chunk[4 + i] = type.charCodeAt(i);
  }
  chunk.set(data, 8);
  view.setUint32(8 + data.length, crc32(chunk.subarray(4, 8 + data.length)));
  return chunk;
}

/** Writes an 800×600 RGB poster — a dusk gradient with a low sun — so the
 * library's image card shows real dimensions over a real preview. */
export function writeDemoPoster(path: string): void {
  const width = 800;
  const height = 600;
  const raw = new Uint8Array(height * (1 + width * 3));
  const sun = { x: 560, y: 380, r: 110 };
  for (let y = 0; y < height; y += 1) {
    const row = y * (1 + width * 3);
    raw[row] = 0; // filter: none
    const t = y / (height - 1);
    // Dusk sky: deep indigo into warm coral.
    const base = {
      r: 27 + (226 - 27) * t,
      g: 42 + (109 - 42) * t,
      b: 107 + (90 - 107) * t,
    };
    for (let x = 0; x < width; x += 1) {
      let { r, g, b } = base;
      const d = Math.hypot(x - sun.x, y - sun.y);
      if (d < sun.r) {
        const glow = 1 - d / sun.r;
        r += (255 - r) * (0.35 + 0.65 * glow);
        g += (240 - g) * (0.3 + 0.6 * glow);
        b += (200 - b) * 0.35 * glow;
      }
      const at = row + 1 + x * 3;
      raw[at] = Math.min(255, Math.round(r));
      raw[at + 1] = Math.min(255, Math.round(g));
      raw[at + 2] = Math.min(255, Math.round(b));
    }
  }
  const ihdr = new Uint8Array(13);
  const ihdrView = new DataView(ihdr.buffer);
  ihdrView.setUint32(0, width);
  ihdrView.setUint32(4, height);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // colour type: truecolour
  const png = Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", deflateSync(raw)),
    pngChunk("IEND", new Uint8Array(0)),
  ]);
  writeFileSync(path, png);
}

// --- PDF -------------------------------------------------------------------

interface PdfLine {
  x: number;
  y: number;
  size: number;
  bold?: boolean;
  text: string;
}

function escapePdfText(text: string): string {
  return text.replace(/[\\()]/g, (char) => `\\${char}`);
}

function pageStream(lines: PdfLine[]): string {
  return lines
    .map(
      (line) =>
        `BT /${line.bold ? "F2" : "F1"} ${line.size} Tf ${line.x} ${line.y} Td (${escapePdfText(line.text)}) Tj ET`,
    )
    .join("\n");
}

/** A paragraph block laid out as successive lines from `top` downward. */
function block(top: number, size: number, texts: string[], bold = false): PdfLine[] {
  return texts.map((text, index) => ({ x: 72, y: top - index * (size + 5), size, bold, text }));
}

/** Writes a classic-xref, three-page A4 PDF with real Helvetica text, so the
 * PDF card gets an honest page count and a legible rasterised preview. */
export function writeDemoPdf(path: string): void {
  const pages: PdfLine[][] = [
    [
      ...block(760, 22, ["SERVICES AGREEMENT"], true),
      ...block(726, 12, ["Between Acme Publishing and Nordwind Studio"]),
      ...block(672, 14, ["1. Scope of work"], true),
      ...block(646, 11, [
        "The studio will design, edit and deliver the client's quarterly",
        "publications, including layout, typesetting and final export to",
        "print-ready documents in the client's house formats.",
      ]),
      ...block(560, 14, ["2. Deliverables"], true),
      ...block(534, 11, [
        "Four issues per year, each delivered as an editable document",
        "and an approved final render, with all source files retained",
        "by the client in the client's own archive.",
      ]),
    ],
    [
      ...block(760, 14, ["3. Schedule"], true),
      ...block(734, 11, [
        "Drafts are due three weeks before each publication date.",
        "Review rounds close one week before print.",
      ]),
      ...block(660, 14, ["4. Fees and payment"], true),
      ...block(634, 11, [
        "Fees are invoiced per issue, payable within thirty days.",
        "Late revisions beyond the second round are billed hourly.",
      ]),
    ],
    [
      ...block(760, 14, ["5. Term and signatures"], true),
      ...block(734, 11, [
        "This agreement runs for one year and renews unless either",
        "party gives notice sixty days before the renewal date.",
      ]),
      ...block(640, 11, ["Signed,"]),
      ...block(600, 11, ["_____________________          _____________________"]),
      ...block(580, 11, ["Acme Publishing                Nordwind Studio"]),
    ],
  ];

  const objects: string[] = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R] /Count 3 >>",
    ...pages.map(
      (_, index) =>
        `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 6 0 R /F2 7 0 R >> >> /Contents ${
          8 + index
        } 0 R >>`,
    ),
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>",
    ...pages.map((lines) => {
      const stream = pageStream(lines);
      return `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`;
    }),
  ];

  let body = "%PDF-1.4\n";
  const offsets: number[] = [];
  objects.forEach((object, index) => {
    offsets.push(body.length);
    body += `${index + 1} 0 obj\n${object}\nendobj\n`;
  });
  const xrefStart = body.length;
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets) {
    body += `${String(offset).padStart(10, "0")} 00000 n \n`;
  }
  body += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF\n`;
  writeFileSync(path, Buffer.from(body, "latin1"));
}
