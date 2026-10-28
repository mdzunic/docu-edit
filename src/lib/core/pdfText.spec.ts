import { readFileSync } from "node:fs";
import { resolve as resolvePath } from "node:path";
import { describe, expect, it } from "vitest";
import { ASSEMBLY_PLAN_DEFAULTS } from "./assemblyLayout";
import { COVERAGE_REASONS } from "./extractedText";
import { buildPdf, type PdfImageData, type PdfPageInput, type PdfTextItem } from "./pdfBuild";
import type { FaceKey } from "./pdfFonts";
import { parseSfnt, type SfntFont } from "./sfnt";
import { resolveSheet, SHEET_DEFAULTS } from "./sheetLayout";
import {
  BASE_ENCODINGS,
  extractPdfText,
  glyphNameToText,
  MAX_PDF_TEXT_PAGES,
  parseToUnicode,
  WORD_GAP_THOUSANDTHS,
} from "./pdfText";
import { assembleClassicPdf, buildJpegFixtureBytes, buildScannedPdfBytes, bytesToBase64 } from "../media/fixtures/build";
import { buildImagePdf, planPdfPages } from "../media/imagesToPdf";

const ascii = (text: string): Uint8Array => Uint8Array.from(text, (char) => char.charCodeAt(0) & 0xff);

function stream(dict: string, data: string): Uint8Array {
  const bytes = ascii(data);
  return Uint8Array.from([...ascii(`<< ${dict} /Length ${bytes.length} >>\nstream\n`), ...bytes, ...ascii("\nendstream")]);
}

/** A PDF whose pages each have their own content and share /F1 (`font`,
 * plus any further objects, numbered from 4 + pages × 2). */
function pdf(contents: string[], font: string, extra: Array<string | Uint8Array> = []): Uint8Array {
  const n = contents.length;
  const kids = contents.map((_, i) => `${3 + i * 2} 0 R`).join(" ");
  const fontRef = 3 + n * 2;
  const objects: Array<string | Uint8Array> = ["<< /Type /Catalog /Pages 2 0 R >>", `<< /Type /Pages /Kids [${kids}] /Count ${n} >>`];
  contents.forEach((content, i) => {
    objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 ${fontRef} 0 R >> >> /Contents ${4 + i * 2} 0 R >>`);
    objects.push(stream("", content));
  });
  objects.push(font, ...extra);
  return assembleClassicPdf("1.7", objects, "/Root 1 0 R");
}

const HELVETICA = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>";

describe("decoding through /ToUnicode", () => {
  const cmap = [
    "/CIDInit /ProcSet findresource begin 12 dict begin begincmap",
    "1 begincodespacerange <0000> <FFFF> endcodespacerange",
    "4 beginbfchar <0001> <0048> <0002> <0065> <0006> <006F> <0007> <00660069> endbfchar",
    "2 beginbfrange <0003> <0004> <006C> <0010> <0011> [<0057> <0072>] endbfrange",
    "endcmap CMapName currentdict /CMap defineresource pop end end",
  ].join("\n");

  it("parses bfchar, both bfrange forms and the codespace", () => {
    const map = parseToUnicode(ascii(cmap));
    expect(map.codespaces).toEqual([{ bytes: 2, lo: 0, hi: 0xffff }]);
    expect([1, 2, 3, 4, 6, 7, 0x10, 0x11].map((code) => map.map.get(code))).toEqual(["H", "e", "l", "m", "o", "fi", "W", "r"]);
  });

  it("decodes a Type0 font's two-byte codes through its map", async () => {
    const file = pdf(
      ["BT /F1 12 Tf 72 700 Td <00010002000300030006> Tj 0 -14 Td <0007000300020010> Tj ET"],
      "<< /Type /Font /Subtype /Type0 /BaseFont /Test /Encoding /Identity-H /DescendantFonts [6 0 R] /ToUnicode 7 0 R >>",
      ["<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Test /DW 500 >>", stream("", cmap)],
    );
    expect(await extractPdfText(file)).toMatchObject({ text: "Hello\nfileW", coverage: "full" });
  });
});

/** A Type 1 font over `base` with /Differences: 65 B, 66 é (by uni name),
 * 70 é, 71 the ffi ligature (by its parts), 72 a small-cap a. */
function differencesFont(base: string): string {
  return `<< /Type /Font /Subtype /Type1 /BaseFont /Times-Roman /Encoding << /Type /Encoding /BaseEncoding /${base} /Differences [65 /B /uni00E9 70 /eacute /f_f_i /a.sc] >> >>`;
}

describe("/Differences over each base encoding", () => {
  const font = differencesFont;

  it("states the three tables", () => {
    for (const table of Object.values(BASE_ENCODINGS)) {
      expect(table).toHaveLength(256);
      expect(table[0x41]).toBe("A");
    }
    expect([glyphNameToText("quoteright"), glyphNameToText("u1F600"), glyphNameToText("g42")]).toEqual(["’", "\u{1F600}", null]);
  });

  it("over WinAnsiEncoding", async () => {
    // 0x92 is ’ and 0x80 € under code page 1252.
    const file = pdf(["BT /F1 12 Tf 72 700 Td (\\101\\102\\106\\107\\110 \\222\\200) Tj ET"], font("WinAnsiEncoding"));
    expect((await extractPdfText(file)).text).toBe("Bééffia ’€");
  });

  it("over MacRomanEncoding", async () => {
    // 0xD5 is ’ and 0x8E é in Mac Roman.
    const file = pdf(["BT /F1 12 Tf 72 700 Td (\\101\\102\\325\\216) Tj ET"], font("MacRomanEncoding"));
    expect((await extractPdfText(file)).text).toBe("Bé’é");
  });

  it("over StandardEncoding", async () => {
    // 0x27 and 0x60 are the typographic quotes, 0xAE the fi ligature.
    const file = pdf(["BT /F1 12 Tf 72 700 Td (\\101\\047\\140\\256) Tj ET"], font("StandardEncoding"));
    expect((await extractPdfText(file)).text).toBe("B’‘ﬁ");
  });

  it("reads a plain /Encoding name, and StandardEncoding when there is none", async () => {
    expect((await extractPdfText(pdf(["BT /F1 12 Tf 72 700 Td (caf\\351) Tj ET"], HELVETICA))).text).toBe("café");
    const bare = "<< /Type /Font /Subtype /Type1 /BaseFont /Times-Roman >>";
    expect((await extractPdfText(pdf(["BT /F1 12 Tf 72 700 Td (it\\047s) Tj ET"], bare))).text).toBe("it’s");
  });
});

describe("a Type0/CID font with no /ToUnicode is unextractable for that page", () => {
  const cidFont = "<< /Type /Font /Subtype /Type0 /BaseFont /Test /Encoding /Identity-H /DescendantFonts [6 0 R] >>";
  const descendant = "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Test >>";

  it("refuses the whole document when every page is such text", async () => {
    const file = pdf(["BT /F1 12 Tf 72 700 Td <00240025> Tj ET"], cidFont, [descendant]);
    expect(await extractPdfText(file)).toEqual({ text: "", coverage: "none", reason: COVERAGE_REASONS.noCharMap, pageCount: 1 });
  });

  it("keeps the readable pages and says which were lost", async () => {
    // Page 1 shows the unmapped font; page 2 a WinAnsi font of its own.
    const objects: Array<string | Uint8Array> = [
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 0 R 5 0 R] /Count 2 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 7 0 R >> >> /Contents 4 0 R >>",
      stream("", "BT /F1 12 Tf 72 700 Td <00240025> Tj ET"),
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F2 9 0 R >> >> /Contents 6 0 R >>",
      stream("", "BT /F2 12 Tf 72 700 Td (Readable text) Tj ET"),
      cidFont.replace("6 0 R", "8 0 R"),
      descendant,
      HELVETICA,
    ];
    const result = await extractPdfText(assembleClassicPdf("1.7", objects, "/Root 1 0 R"));
    expect(result.text).toBe("Readable text");
    expect(result.coverage).toBe("partial");
    expect(result.reason).toBe(`${COVERAGE_REASONS.noCharMap} on 1 of 2 pages`);
  });

  it("reads a Uni…-UCS2 CMap's codes as the UTF-16 they are", async () => {
    const font = cidFont.replace("/Identity-H", "/UniJIS-UCS2-H");
    const file = pdf(["BT /F1 12 Tf 72 700 Td <67714EAC> Tj ET"], font, [descendant]);
    expect(await extractPdfText(file)).toMatchObject({ text: "東京", coverage: "full" });
  });
});

describe("word gaps and line breaks", () => {
  it(`inserts a space where a TJ offset exceeds ${WORD_GAP_THOUSANDTHS} thousandths`, async () => {
    const file = pdf(["BT /F1 12 Tf 72 700 Td [(Rotter) -150 (dam) -250 (figures) 120 (!)] TJ ET"], HELVETICA);
    expect((await extractPdfText(file)).text).toBe("Rotterdam figures!");
  });

  it("breaks the line where the text-line matrix moves down, not for a raised figure", async () => {
    const content = [
      "BT /F1 12 Tf 14 TL 72 700 Td (first line) Tj T* (second line) Tj",
      // 15 glyphs at the no-/Widths half-em: the figure starts at the pen, 4 pt up.
      "0 -14 Td (third ) Tj (same line) Tj 90 4 Td (2) Tj",
      "(fourth) ' ET",
    ].join("\n");
    expect((await extractPdfText(pdf([content], HELVETICA))).text).toBe("first line\nsecond line\nthird same line2\nfourth");
  });

  it("puts a space between text resumed further along the same line", async () => {
    const content = "BT /F1 10 Tf 72 600 Td (Cell A) Tj ET BT /F1 10 Tf 300 600 Td (Cell B) Tj ET";
    expect((await extractPdfText(pdf([content], HELVETICA))).text).toBe("Cell A Cell B");
  });

  it("reads text inside a form XObject, and skips inline image data", async () => {
    const objects: Array<string | Uint8Array> = [
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> /XObject << /Fm1 6 0 R >> >> /Contents 4 0 R >>",
      stream("", "BI /W 2 /H 1 /BPC 8 /CS /G ID ÿ) EI\nBT /F1 12 Tf 72 700 Td (Before) Tj ET q 1 0 0 1 0 -100 cm /Fm1 Do Q"),
      HELVETICA,
      stream("/Type /XObject /Subtype /Form /BBox [0 0 612 792]", "BT /F1 12 Tf 72 500 Td (Inside the form) Tj ET"),
    ];
    expect((await extractPdfText(assembleClassicPdf("1.7", objects, "/Root 1 0 R"))).text).toBe("Before\nInside the form");
  });
});

describe("coverage is reported, never swallowed", () => {
  it("calls the committed scanned fixture a scan with no text layer — and never skips it", async () => {
    const bytes = new Uint8Array(readFileSync(resolvePath(__dirname, "../../../e2e/fixtures/spec114-scanned.pdf")));
    expect(Array.from(bytes)).toEqual(Array.from(buildScannedPdfBytes()));
    expect(await extractPdfText(bytes)).toEqual({ text: "", coverage: "none", reason: COVERAGE_REASONS.scanned, pageCount: 1 });
  });

  it("refuses an unreadable or encrypted file with its reason", async () => {
    expect(await extractPdfText(ascii("not a pdf"))).toMatchObject({ coverage: "none", reason: COVERAGE_REASONS.unreadable });
    const encrypted = new Uint8Array(readFileSync(resolvePath(__dirname, "../../../e2e/fixtures/spec107-encrypted.pdf")));
    expect(await extractPdfText(encrypted)).toMatchObject({ coverage: "none", reason: COVERAGE_REASONS.encrypted });
  });

  it(`reads at most MAX_PDF_TEXT_PAGES (${MAX_PDF_TEXT_PAGES}) pages, saying so`, async () => {
    const file = pdf(["BT /F1 12 Tf 72 700 Td (one) Tj ET", "BT /F1 12 Tf 72 700 Td (two) Tj ET", "BT /F1 12 Tf 72 700 Td (three) Tj ET"], HELVETICA);
    const result = await extractPdfText(file, { maxPages: 2 });
    expect(result).toMatchObject({ text: "one\ntwo", coverage: "partial", pageCount: 3 });
    expect(result.reason).toContain(COVERAGE_REASONS.sizeCeiling);
  });

  it("reads a text PDF from the e2e fixtures fully (Helvetica, WinAnsi, no /Widths)", async () => {
    const bytes = new Uint8Array(readFileSync(resolvePath(__dirname, "../../../e2e/fixtures/spec111-outlines.pdf")));
    const result = await extractPdfText(bytes);
    expect(result.coverage).toBe("full");
    expect(result.text).toBe("Chapter 1\nGo to page 2\nGo to chapter 2\nexample.com\nChapter 2\nBack to chapter 1");
  });
});

// --- this product's own PDFs -----------------------------------------------------

const sans: SfntFont = parseSfnt(new Uint8Array(readFileSync(resolvePath(__dirname, "../../../public/fonts/LiberationSans-Regular.ttf"))));
const fonts = (): Map<FaceKey, SfntFont> => new Map([["sans-regular", sans]]);

function textItem(text: string, xPt: number, yPt: number): PdfTextItem {
  return { text, xPt, yPt, sizePt: 12, face: "sans-regular", color: [0, 0, 0] };
}

function page(items: PdfTextItem[], extra: Partial<PdfPageInput> = {}): PdfPageInput {
  return { widthPt: 612, heightPt: 792, items, links: [], headings: [], images: [], ...extra };
}

describe("every PDF this product writes is searchable — or says it has no text", () => {
  it("re-extracts the strings a buildPdf output (SPEC-034's writer) put in, as full", async () => {
    const lines = ["Rotterdam harbour figures", "Second line: café, naïve, 3.14 €", "Third — “quoted” and don’t"];
    const bytes = await buildPdf({
      title: "Round trip",
      fonts: fonts(),
      pages: [
        page(lines.map((line, i) => textItem(line, 72, 100 + i * 16))),
        page([textItem("Hello ", 72, 100), textItem("world", 106, 100), textItem("Name", 72, 140), textItem("Value", 300, 140)]),
      ],
    });
    const result = await extractPdfText(bytes);
    expect(result.coverage).toBe("full");
    expect(result.text).toBe([...lines, "Hello world", "Name Value"].join("\n"));
  });

  it("calls an image-only PDF it writes (SPEC-108's conversion) a scan: none", async () => {
    // ConvertToPdfDialog's page: a background rect and one placed image, no text.
    const image: PdfImageData = { widthPx: 1, heightPx: 1, filter: "FlateDecode", colorSpace: "DeviceGray", data: new Uint8Array([0x78, 0x9c, 0x63, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01]) };
    const bytes = await buildPdf({
      title: "Photo",
      fonts: new Map(),
      images: new Map([["photo", image]]),
      pages: [page([], { rects: [{ xPt: 0, yPt: 0, widthPt: 612, heightPt: 792, fill: "#ffffff" }], images: [{ src: "photo", xPt: 36, yPt: 36, widthPt: 540, heightPt: 720 }] })],
    });
    expect(await extractPdfText(bytes)).toMatchObject({ text: "", coverage: "none", reason: COVERAGE_REASONS.scanned });
  });

  it("calls a SPEC-110 image assembly none, page numbers and all", async () => {
    const jpeg = buildJpegFixtureBytes();
    const items = [
      { id: "a", title: "a", widthPx: 160, heightPx: 120 },
      { id: "b", title: "b", widthPx: 160, heightPx: 120 },
    ];
    const geometry = resolveSheet(SHEET_DEFAULTS);
    const pages = planPdfPages(items, { ...ASSEMBLY_PLAN_DEFAULTS, pageMode: "one-per-page", nUp: 4 }, geometry);
    const part = { mediaId: "m", contentType: "image/jpeg", data: bytesToBase64(jpeg) };
    for (const pageNumbers of [false, true]) {
      const bytes = await buildImagePdf(pages, new Map([["a", part], ["b", part]]), {
        title: "Receipts",
        pageNumbers,
        modified: new Date(Date.UTC(2026, 9, 28)),
        items,
        sheet: { geometry, background: "#ffffff", dpi: 150 },
        fonts: fonts(),
        steps: { between: async () => {} },
      });
      expect(await extractPdfText(bytes)).toMatchObject({ text: "", coverage: "none", reason: COVERAGE_REASONS.scanned });
    }
  });
});
