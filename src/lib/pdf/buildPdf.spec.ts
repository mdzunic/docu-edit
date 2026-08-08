import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { buildPdf, pdfFileName, type PdfBuildInput, type PdfPageInput } from "./buildPdf";
import type { FaceKey } from "./fontMap";
import { parseSfnt, type SfntFont } from "./sfnt";

function loadFace(file: string): SfntFont {
  return parseSfnt(
    new Uint8Array(readFileSync(resolve(__dirname, "../../../public/fonts", file))),
  );
}

const sans = loadFace("LiberationSans-Regular.ttf");
const sansBold = loadFace("LiberationSans-Bold.ttf");

function fonts(map: Partial<Record<FaceKey, SfntFont>> = { "sans-regular": sans }) {
  return new Map(Object.entries(map)) as Map<FaceKey, SfntFont>;
}

function page(overrides: Partial<PdfPageInput> = {}): PdfPageInput {
  return {
    widthPt: 595.28,
    heightPt: 841.89,
    items: [
      {
        text: "Hello world",
        xPt: 72,
        yPt: 90,
        sizePt: 11,
        face: "sans-regular",
        color: [0, 0, 0],
      },
    ],
    links: [],
    headings: [],
    ...overrides,
  };
}

/** The file as a latin-1 string, for structural assertions. */
function latin1(bytes: Uint8Array): string {
  let text = "";
  for (const byte of bytes) {
    text += String.fromCharCode(byte);
  }
  return text;
}

function build(input: Partial<PdfBuildInput> = {}): string {
  return latin1(
    buildPdf({ title: "Test", pages: [page()], fonts: fonts(), ...input }),
  );
}

describe("buildPdf", () => {
  it("emits a v1.7 header, catalog, xref and EOF", () => {
    const pdf = build();
    expect(pdf.startsWith("%PDF-1.7\n")).toBe(true);
    expect(pdf).toContain("/Type /Catalog");
    expect(pdf).toContain("xref");
    expect(pdf.trimEnd().endsWith("%%EOF")).toBe(true);
  });

  it("records correct xref byte offsets for every object", () => {
    const bytes = buildPdf({ title: "T", pages: [page()], fonts: fonts() });
    const pdf = latin1(bytes);
    const xrefStart = Number(/startxref\n(\d+)\n/.exec(pdf)?.[1]);
    expect(pdf.slice(xrefStart, xrefStart + 4)).toBe("xref");
    const table = pdf.slice(xrefStart);
    const size = Number(/xref\n0 (\d+)\n/.exec(table)?.[1]);
    const entries = table.match(/\d{10} \d{5} [nf] /g) ?? [];
    expect(entries.length).toBe(size);
    entries.slice(1).forEach((entry, index) => {
      const offset = Number(entry.slice(0, 10));
      expect(pdf.slice(offset, offset + String(index + 1).length + 6)).toBe(
        `${index + 1} 0 obj`,
      );
    });
  });

  it("writes text as Tj operators against an embedded CIDFontType2 face", () => {
    const pdf = build();
    expect(pdf).toContain(" Tf ");
    expect(pdf).toContain("> Tj ET");
    expect(pdf).toContain("/Subtype /CIDFontType2");
    expect(pdf).toContain("/Encoding /Identity-H");
    expect(pdf).toContain("/FontFile2");
    expect(pdf).toContain("/ToUnicode");
    expect(pdf).toContain("/BaseFont /LiberationSans");
    // The embedded program is the verbatim font file (sfnt version tag).
    expect(pdf).toContain("/Length1 " + sans.bytes.length);
  });

  it("encodes text as the face's glyph ids with a reverse ToUnicode map", () => {
    const pdf = build({
      pages: [page({ items: [{ text: "A", xPt: 0, yPt: 10, sizePt: 12, face: "sans-regular", color: [0, 0, 0] }] })],
    });
    const glyphHex = sans.glyphFor(0x41).toString(16).padStart(4, "0");
    expect(pdf).toContain(`<${glyphHex}> Tj`);
    expect(pdf).toContain(`<${glyphHex}> <0041>`);
  });

  it("only embeds faces the items use, with per-glyph widths", () => {
    const pdf = build({
      pages: [
        page({
          items: [
            { text: "a", xPt: 0, yPt: 10, sizePt: 12, face: "sans-regular", color: [0, 0, 0] },
            { text: "b", xPt: 0, yPt: 30, sizePt: 12, face: "sans-bold", color: [0, 0, 0] },
          ],
        }),
      ],
      fonts: fonts({ "sans-regular": sans, "sans-bold": sansBold, "serif-regular": sans }),
    });
    expect(pdf).toContain("/BaseFont /LiberationSans-Bold");
    expect((pdf.match(/\/FontFile2/g) ?? []).length).toBe(2);
    expect(pdf).toContain("/W [ ");
  });

  it("substitutes a loaded face for one that is missing", () => {
    const pdf = build({
      pages: [page({ items: [{ text: "x", xPt: 0, yPt: 10, sizePt: 12, face: "mono-bold", color: [0, 0, 0] }] })],
    });
    expect(pdf).toContain("/BaseFont /LiberationSans");
  });

  it("throws when text exists but no fonts were provided", () => {
    expect(() => build({ fonts: new Map() })).toThrow(/no fonts/);
  });

  it("writes the title into Info, empty when the document has none", () => {
    expect(build({ title: "My report" })).toContain("/Title (My report)");
    expect(build({ title: "" })).toContain("/Title ()");
  });

  it("writes non-ASCII titles as UTF-16BE with a BOM", () => {
    expect(build({ title: "Café" })).toContain("/Title <FEFF00430061006600E9>");
  });

  it("builds a nested outline tree from headings in document order", () => {
    const pdf = build({
      pages: [
        page({
          headings: [
            { level: 1, title: "One", yPt: 50 },
            { level: 2, title: "One point one", yPt: 200 },
            { level: 3, title: "Deep", yPt: 300 },
            { level: 1, title: "Two", yPt: 500 },
          ],
        }),
      ],
    });
    expect(pdf).toContain("/Type /Outlines");
    expect(pdf).toContain("/PageMode /UseOutlines");
    expect(pdf).toContain("/Title (One)");
    expect(pdf).toContain("/Title (Two)");
    // "One" has two descendants; the root counts all four entries.
    expect(pdf).toContain("/Count 2");
    expect(pdf).toContain("/Count 4");
    // Destinations are y-flipped from the top-anchored input.
    expect(pdf).toContain(`/XYZ 0 ${841.89 - 50} 0`);
  });

  it("omits outlines (and UseOutlines) when there are no headings", () => {
    const pdf = build();
    expect(pdf).not.toContain("/Outlines");
    expect(pdf).not.toContain("/UseOutlines");
  });

  it("writes external links as URI annotations with escaped targets", () => {
    const pdf = build({
      pages: [
        page({
          links: [
            { xPt: 10, yPt: 20, widthPt: 100, heightPt: 12, href: "https://example.com/a(b)" },
          ],
        }),
      ],
    });
    expect(pdf).toContain("/Subtype /Link");
    expect(pdf).toContain("/S /URI /URI (https://example.com/a\\(b\\))");
    expect(pdf).toContain("/Rect [10 ");
  });

  it("writes internal links as destinations to the target page", () => {
    const pdf = build({
      pages: [
        page({ links: [{ xPt: 10, yPt: 20, widthPt: 50, heightPt: 12, destPage: 1 }] }),
        page({ items: [] }),
      ],
    });
    expect(pdf).toContain("/Dest [");
    expect(pdf).not.toContain("/S /URI");
  });

  it("drops links whose internal target page does not exist", () => {
    const pdf = build({
      pages: [page({ links: [{ xPt: 0, yPt: 0, widthPt: 10, heightPt: 10, destPage: 9 }] })],
    });
    expect(pdf).not.toContain("/Subtype /Link");
  });

  it("builds a zero-page document for an empty model", () => {
    const pdf = build({ pages: [], fonts: new Map() });
    expect(pdf).toContain("/Kids [] /Count 0");
    expect(pdf).toContain("/Title (Test)");
    expect(pdf).not.toContain("/Font");
  });

  it("uses per-page media boxes from each page's own size", () => {
    const pdf = build({
      pages: [page(), page({ widthPt: 612, heightPt: 792, items: [] })],
    });
    expect(pdf).toContain("/MediaBox [0 0 595.28 841.89]");
    expect(pdf).toContain("/MediaBox [0 0 612 792]");
    expect(pdf).toContain("/Count 2");
  });
});

describe("pdfFileName", () => {
  it("appends .pdf to the trimmed title, falling back when blank", () => {
    expect(pdfFileName("Report ", "Untitled document")).toBe("Report.pdf");
    expect(pdfFileName("  ", "Untitled document")).toBe("Untitled document.pdf");
  });
});
