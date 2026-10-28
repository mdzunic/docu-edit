// Regenerates the checked-in SPEC-107 media e2e fixtures:
//   npx vite-node scripts/build-media-fixtures.ts
//
// How each file is produced (the production note SPEC-107 asks for):
// - spec107-pages.pdf — the app's OWN PDF writer, buildPdf, so the file is
//   reproducible from this script alone. The writer is imported from
//   src/lib/core/pdfBuild.ts (moved there by SPEC-104). Three pages of mixed sizes — US Letter portrait,
//   A4 with /Rotate 90, and a 720 × 540 pt landscape page — each drawn as
//   vector paths only (a text item would need a font embedded), plus an
//   /Info dictionary with a title, author, producer and creation date.
// - spec107-xrefstream.pdf — written byte by byte in
//   src/lib/media/fixtures/build.ts (buildXrefStreamPdfBytes): a PDF 1.5 file
//   whose catalog and /Info live in a deflated object stream and whose
//   cross-reference table is a STREAM using the PNG Up predictor
//   (/Predictor 12, /W [1 2 1]). buildPdf only writes classic xref tables,
//   hence the hand-built file.
// - spec107-encrypted.pdf — written byte by byte in the same module
//   (buildEncryptedPdfBytes): a classic-xref file whose trailer names an
//   /Encrypt dictionary (a Standard security handler stub, /V 1 /R 2). Its
//   strings are not actually encrypted — readers must refuse the file on the
//   /Encrypt entry alone, without touching the objects.
// - spec107-photo.jpg, -shot.png, -mark.svg, -anim.gif, -shot.webp — the
//   image builders in the same module; their header comment records how each
//   is made (a real 64×40 PNG, a real 160×120 baseline JPEG, an SVG behind a
//   comment and doctype, a 2-frame animated GIF89a, a VP8L WebP container).
// - spec111-forms.pdf, spec111-outlines.pdf (SPEC-111's merge fixtures) —
//   written byte by byte in the same module (buildFormsPdfBytes,
//   buildOutlinesPdfBytes): a filled form with widgets, a signature, document
//   JavaScript, an open action, an attached file, page labels, optional
//   content and inherited page attributes; and a tagged two-page PDF with
//   bookmarks, named destinations and internal and web links. Their header
//   comments list every construct each one carries.
// - spec114-scanned.pdf (SPEC-114's "Not searched" fixture) — written byte by
//   byte in the same module (buildScannedPdfBytes): one page painting an
//   image XObject, with no text operators and no font.
// The unit specs read these same checked-in bytes, so unit and e2e
// assertions cannot drift apart.
import { writeFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { buildPdf, type PdfPageInput, type PdfPathSubpath } from "../src/lib/core/pdfBuild";
import {
  buildEncryptedPdfBytes,
  buildFormsPdfBytes,
  buildGifFixtureBytes,
  buildJpegFixtureBytes,
  buildOutlinesPdfBytes,
  buildPngFixtureBytes,
  buildScannedPdfBytes,
  buildSvgFixtureBytes,
  buildWebpFixtureBytes,
  buildXrefStreamPdfBytes,
} from "../src/lib/media/fixtures/build";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "e2e", "fixtures");
mkdirSync(fixturesDir, { recursive: true });

function rect(xPt: number, yPt: number, widthPt: number, heightPt: number): PdfPathSubpath {
  return {
    start: { xPt, yPt },
    segments: [
      { toXPt: xPt + widthPt, toYPt: yPt },
      { toXPt: xPt + widthPt, toYPt: yPt + heightPt },
      { toXPt: xPt, toYPt: yPt + heightPt },
    ],
    closed: true,
  };
}

/** One fixture page: a stroked frame inset 36 pt, and one filled square per
 * page number along its top edge, so each page is recognisable in the
 * browser's viewer. */
function fixturePage(widthPt: number, heightPt: number, pageNumber: number, rotate: 0 | 90 = 0): PdfPageInput {
  const inset = 36;
  const squares = Array.from({ length: pageNumber }, (_, i) => rect(inset + 16 + i * 40, inset + 16, 28, 28));
  return {
    widthPt,
    heightPt,
    rotate,
    items: [],
    links: [],
    headings: [],
    images: [],
    paths: [
      { subpaths: [rect(inset, inset, widthPt - inset * 2, heightPt - inset * 2)], stroke: { color: [0.12, 0.47, 0.71], widthPt: 2 } },
      { subpaths: squares, fill: [1, 0.5, 0.05] },
    ],
  };
}

const pagesPdf = await buildPdf({
  title: "SPEC-107 pages",
  info: { author: "docu-edit factory", creationDate: new Date(Date.UTC(2026, 7, 30, 12, 0, 0)) },
  pages: [fixturePage(612, 792, 1), fixturePage(595.28, 841.89, 2, 90), fixturePage(720, 540, 3)],
  fonts: new Map(),
});

const files: Array<[string, Uint8Array]> = [
  ["spec107-photo.jpg", buildJpegFixtureBytes()],
  ["spec107-shot.png", await buildPngFixtureBytes()],
  ["spec107-mark.svg", buildSvgFixtureBytes()],
  ["spec107-anim.gif", buildGifFixtureBytes()],
  ["spec107-shot.webp", buildWebpFixtureBytes()],
  ["spec107-pages.pdf", pagesPdf],
  ["spec107-xrefstream.pdf", await buildXrefStreamPdfBytes()],
  ["spec107-encrypted.pdf", buildEncryptedPdfBytes()],
  ["spec111-forms.pdf", await buildFormsPdfBytes()],
  ["spec111-outlines.pdf", await buildOutlinesPdfBytes()],
  ["spec114-scanned.pdf", buildScannedPdfBytes()],
];

for (const [name, bytes] of files) {
  writeFileSync(join(fixturesDir, name), bytes);
}
console.log(files.map(([name, bytes]) => `${name} (${bytes.byteLength}B)`).join(", "));
// The core idle scheduler's keep-alive timer would otherwise hold the
// process open after the work is done (a known trap for Node callers).
process.exit(0);
