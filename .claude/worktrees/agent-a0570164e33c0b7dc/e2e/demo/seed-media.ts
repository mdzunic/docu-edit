import { writeFileSync } from "node:fs";

// Seed-file generator for the library tour (library.tour.ts). The repo's
// SPEC-107 PDF fixtures are deliberately minimal — 726 bytes of empty
// pages — which is exactly right for parsers and exactly wrong for a
// product demo, where the card would raster a blank preview. This builds a
// presentable stand-in at seed time instead, dependency-free. (The tour's
// image seed is a real photograph, e2e/demo/golden-gate.jpg.)

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
