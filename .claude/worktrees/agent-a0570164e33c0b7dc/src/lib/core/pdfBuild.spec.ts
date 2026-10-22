import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import {
  buildPdf,
  pdfFileName,
  type PdfBuildInput,
  type PdfImageData,
  type PdfImageItem,
  type PdfPageInput,
  type PdfPathItem,
} from "./pdfBuild";
import { equalBytes } from "./equality";
import type { FaceKey } from "./pdfFonts";
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
    images: [],
    ...overrides,
  };
}

function imageItem(overrides: Partial<PdfImageItem> = {}): PdfImageItem {
  return { src: "img", xPt: 72, yPt: 100, widthPt: 100, heightPt: 50, ...overrides };
}

/** Decoded-image stand-ins keyed by src; the bytes never have to be a real
 * JPEG because the builder only ever passes them through. */
function imageMap(srcs: string[] = ["img"]): Map<string, PdfImageData> {
  return new Map(
    srcs.map((src) => [
      src,
      {
        widthPx: 8,
        heightPx: 4,
        filter: "DCTDecode" as const,
        colorSpace: "DeviceRGB" as const,
        data: new Uint8Array([0xff, 0xd8, 0xff, 0xd9]),
      },
    ]),
  );
}

/** The file as a latin-1 string, for structural assertions. */
function latin1(bytes: Uint8Array): string {
  let text = "";
  for (const byte of bytes) {
    text += String.fromCharCode(byte);
  }
  return text;
}

/** The file as latin-1, with every /FlateDecode stream expanded in place.
 * Content streams and font programs are compressed in the real output, so
 * the structural assertions below would otherwise be reading zlib bytes —
 * and round-tripping each one here doubles as proof that what the writer
 * deflated is valid zlib a reader can actually get back. */
async function inflateStreams(bytes: Uint8Array): Promise<string> {
  const raw = latin1(bytes);
  const header = /<<([^>]*)\/Length (\d+) >>\nstream\n/g;
  let out = "";
  let cursor = 0;
  for (let match = header.exec(raw); match !== null; match = header.exec(raw)) {
    const start = match.index + match[0].length;
    const length = Number(match[2]);
    out += raw.slice(cursor, start);
    const body = bytes.subarray(start, start + length);
    let expanded: string | null = null;
    if (match[1].includes("/Filter /FlateDecode")) {
      try {
        const stream = new ReadableStream<Uint8Array>({
          start(controller) {
            controller.enqueue(body);
            controller.close();
          },
        }).pipeThrough(new DecompressionStream("deflate"));
        expanded = latin1(new Uint8Array(await new Response(stream).arrayBuffer()));
      } catch {
        // Not really zlib — the image tests hand the writer synthetic bytes
        // that it passes through untouched. Read them as they are.
        expanded = null;
      }
    }
    out += expanded ?? latin1(body);
    cursor = start + length;
  }
  return out + raw.slice(cursor);
}

async function build(input: Partial<PdfBuildInput> = {}): Promise<string> {
  return inflateStreams(
    await buildPdf({ title: "Test", pages: [page()], fonts: fonts(), ...input }),
  );
}

describe("buildPdf", () => {
  it("emits a v1.7 header, catalog, xref and EOF", async () => {
    const pdf = await build();
    expect(pdf.startsWith("%PDF-1.7\n")).toBe(true);
    expect(pdf).toContain("/Type /Catalog");
    expect(pdf).toContain("xref");
    expect(pdf.trimEnd().endsWith("%%EOF")).toBe(true);
  });

  it("records correct xref byte offsets for every object", async () => {
    const bytes = await buildPdf({ title: "T", pages: [page()], fonts: fonts() });
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

  // The embedded font programs dwarf everything else in a short document —
  // four raw Liberation faces are 1.6 MB, and the Carlito set most .docx
  // files resolve to is 2.7 MB — so an uncompressed one-page export was 98%
  // font bytes. Deflating the font and content streams roughly halves a
  // one-page file and takes a long one to a sixth of its former size.
  it("compresses the font programs and page content streams", async () => {
    const bytes = await buildPdf({ title: "T", pages: [page()], fonts: fonts() });
    const raw = latin1(bytes);
    // The uncompressed markers are only visible once the streams are expanded.
    expect(raw).not.toContain(" Tj ET");
    expect(await inflateStreams(bytes)).toContain(" Tj ET");
    expect(raw).toContain("/Filter /FlateDecode");
    // /Length1 stays the UNCOMPRESSED sfnt length, which is what a reader
    // needs to reconstruct the font; /Length is the deflated stream.
    const fontFile = /\/Length1 (\d+) \/Filter \/FlateDecode \/Length (\d+) >>/.exec(raw);
    expect(fontFile, "the /FontFile2 stream is deflated").not.toBeNull();
    expect(Number(fontFile![2])).toBeLessThan(Number(fontFile![1]));
  });

  it("writes text as Tj operators against an embedded CIDFontType2 face", async () => {
    const pdf = await build();
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

  it("encodes text as the face's glyph ids with a reverse ToUnicode map", async () => {
    const pdf = await build({
      pages: [page({ items: [{ text: "A", xPt: 0, yPt: 10, sizePt: 12, face: "sans-regular", color: [0, 0, 0] }] })],
    });
    const glyphHex = sans.glyphFor(0x41).toString(16).padStart(4, "0");
    expect(pdf).toContain(`<${glyphHex}> Tj`);
    expect(pdf).toContain(`<${glyphHex}> <0041>`);
  });

  it("only embeds faces the items use, with per-glyph widths", async () => {
    const pdf = await build({
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

  it("substitutes a loaded face for one that is missing", async () => {
    const pdf = await build({
      pages: [page({ items: [{ text: "x", xPt: 0, yPt: 10, sizePt: 12, face: "mono-bold", color: [0, 0, 0] }] })],
    });
    expect(pdf).toContain("/BaseFont /LiberationSans");
  });

  it("throws when text exists but no fonts were provided", async () => {
    await expect(build({ fonts: new Map() })).rejects.toThrow(/no fonts/);
  });

  it("writes the title into Info, empty when the document has none", async () => {
    expect(await build({ title: "My report" })).toContain("/Title (My report)");
    expect(await build({ title: "" })).toContain("/Title ()");
  });

  it("writes non-ASCII titles as UTF-16BE with a BOM", async () => {
    expect(await build({ title: "Café" })).toContain("/Title <FEFF00430061006600E9>");
  });

  it("writes the optional author and dates into Info, and nothing when absent", async () => {
    const pdf = await build({
      title: "Pages",
      info: {
        author: "Ana",
        creationDate: new Date(Date.UTC(2026, 7, 30, 12, 0, 5)),
        modDate: new Date(Date.UTC(2026, 9, 1, 9, 30, 0)),
      },
    });
    expect(pdf).toContain(
      "<< /Title (Pages) /Author (Ana) /Producer (docu-edit) /CreationDate (D:20260830120005Z) /ModDate (D:20261001093000Z) >>",
    );
    expect(await build({ title: "Pages" })).toContain("<< /Title (Pages) /Producer (docu-edit) >>");
  });

  it("writes a page's /Rotate only when it is turned (SPEC-107)", async () => {
    const pdf = await build({
      pages: [page({ items: [] }), page({ items: [], rotate: 90 }), page({ items: [], rotate: 0 })],
      fonts: new Map(),
    });
    expect(pdf.split("/Rotate").length - 1).toBe(1);
    expect(pdf).toContain("/MediaBox [0 0 595.28 841.89] /Rotate 90 /Resources");
  });

  it("builds a nested outline tree from headings in document order", async () => {
    const pdf = await build({
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

  it("omits outlines (and UseOutlines) when there are no headings", async () => {
    const pdf = await build();
    expect(pdf).not.toContain("/Outlines");
    expect(pdf).not.toContain("/UseOutlines");
  });

  it("writes external links as URI annotations with escaped targets", async () => {
    const pdf = await build({
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

  it("writes internal links as destinations to the target page", async () => {
    const pdf = await build({
      pages: [
        page({ links: [{ xPt: 10, yPt: 20, widthPt: 50, heightPt: 12, destPage: 1 }] }),
        page({ items: [] }),
      ],
    });
    expect(pdf).toContain("/Dest [");
    expect(pdf).not.toContain("/S /URI");
  });

  it("drops links whose internal target page does not exist", async () => {
    const pdf = await build({
      pages: [page({ links: [{ xPt: 0, yPt: 0, widthPt: 10, heightPt: 10, destPage: 9 }] })],
    });
    expect(pdf).not.toContain("/Subtype /Link");
  });

  it("builds a zero-page document for an empty model", async () => {
    const pdf = await build({ pages: [], fonts: new Map() });
    expect(pdf).toContain("/Kids [] /Count 0");
    expect(pdf).toContain("/Title (Test)");
    expect(pdf).not.toContain("/Font");
  });

  it("uses per-page media boxes from each page's own size", async () => {
    const pdf = await build({
      pages: [page(), page({ widthPt: 612, heightPt: 792, items: [] })],
    });
    expect(pdf).toContain("/MediaBox [0 0 595.28 841.89]");
    expect(pdf).toContain("/MediaBox [0 0 612 792]");
    expect(pdf).toContain("/Count 2");
  });

  it("embeds a placed image as an XObject and draws it at the placed box", async () => {
    const pdf = await build({
      pages: [page({ images: [imageItem()] })],
      images: imageMap(),
    });
    expect(pdf).toContain("/Subtype /Image /Width 8 /Height 4");
    expect(pdf).toContain("/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode");
    expect(pdf).toContain("/XObject << /Im1");
    // 100×50pt box at (72, 100) from the page top on an 841.89pt page:
    // PDF's origin is bottom-left, so y = 841.89 − 100 − 50.
    expect(pdf).toContain("q 100 0 0 50 72 691.89 cm /Im1 Do Q");
  });

  it("paints images under the text, and in-front images over it", async () => {
    const pdf = await build({
      pages: [
        page({
          images: [imageItem({ src: "under" }), imageItem({ src: "over", onTop: true })],
        }),
      ],
      images: imageMap(["under", "over"]),
    });
    const under = pdf.indexOf("/Im1 Do");
    const over = pdf.indexOf("/Im2 Do");
    const text = pdf.indexOf(" Tj ");
    expect(under).toBeGreaterThan(-1);
    expect(under).toBeLessThan(text);
    expect(text).toBeLessThan(over);
  });

  it("embeds one XObject for a source placed on several pages", async () => {
    const pdf = await build({
      pages: [page({ images: [imageItem()] }), page({ items: [], images: [imageItem()] })],
      images: imageMap(),
    });
    expect(pdf.split("/Subtype /Image").length - 1).toBe(1);
    expect(pdf.split("/Im1 Do").length - 1).toBe(2);
  });

  it("writes a transparent image's alpha as an /SMask", async () => {
    const pdf = await build({
      pages: [page({ images: [imageItem()] })],
      images: new Map([
        [
          "img",
          {
            widthPx: 8,
            heightPx: 4,
            filter: "FlateDecode" as const,
            colorSpace: "DeviceRGB" as const,
            data: new Uint8Array([1, 2, 3]),
            alpha: new Uint8Array([9, 9]),
          },
        ],
      ]),
    });
    expect(pdf).toContain("/SMask");
    expect(pdf).toContain("/ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode");
  });

  it("skips a placed image that has no decoded data, keeping the rest", async () => {
    const pdf = await build({ pages: [page({ images: [imageItem()] })], images: new Map() });
    expect(pdf).not.toContain("/Subtype /Image");
    expect(pdf).not.toContain("/XObject");
    expect(pdf).toContain(" Tj ");
  });
});

describe("filled rects (SPEC-108)", () => {
  it("fills a #rrggbb rect beneath the images, flipping y into PDF space", async () => {
    const pdf = await build({
      pages: [
        page({
          items: [],
          rects: [{ xPt: 0, yPt: 0, widthPt: 595.28, heightPt: 841.89, fill: "#ff8000" }],
          images: [imageItem()],
        }),
      ],
      images: imageMap(),
    });
    const fill = pdf.indexOf("q 1 0.5 0 rg 0 0 595.28 841.89 re f Q");
    expect(fill).toBeGreaterThan(-1);
    expect(fill).toBeLessThan(pdf.indexOf("/Im1 Do"));
  });

  it("measures a rect from the page top, like an image", async () => {
    const pdf = await build({
      pages: [page({ items: [], rects: [{ xPt: 10, yPt: 100, widthPt: 20, heightPt: 50, fill: "#000000" }] })],
    });
    expect(pdf).toContain("q 0 0 0 rg 10 691.89 20 50 re f Q");
  });

  it("paints nothing for a fill that is not #rrggbb, or for a page with no rects", async () => {
    const pdf = await build({
      pages: [page({ items: [], rects: [{ xPt: 0, yPt: 0, widthPt: 5, heightPt: 5, fill: "transparent" }] }), page({ items: [] })],
    });
    expect(pdf).not.toContain(" re f");
  });
});

describe("stroked and on-top rects (SPEC-104)", () => {
  it("fills an rgb-triple rect exactly as the matching #rrggbb one", async () => {
    const pdf = await build({
      pages: [page({ items: [], rects: [{ xPt: 10, yPt: 100, widthPt: 20, heightPt: 50, fill: [1, 0.5, 0] }] })],
    });
    expect(pdf).toContain("q 1 0.5 0 rg 10 691.89 20 50 re f Q");
  });

  it("strokes a rect with its colour and width, defaulting to 0.75 pt", async () => {
    const pdf = await build({
      pages: [
        page({
          items: [],
          rects: [
            { xPt: 10, yPt: 100, widthPt: 20, heightPt: 50, stroke: [0, 0, 0] },
            { xPt: 0, yPt: 0, widthPt: 5, heightPt: 5, stroke: [0.5, 0.5, 0.5], strokeWidthPt: 0.5 },
          ],
        }),
      ],
    });
    expect(pdf).toContain("q 0 0 0 RG 0.75 w 10 691.89 20 50 re S Q");
    expect(pdf).toContain("q 0.5 0.5 0.5 RG 0.5 w 0 836.89 5 5 re S Q");
    expect(pdf).not.toContain(" re f");
  });

  it("fills then strokes a rect carrying both", async () => {
    const pdf = await build({
      pages: [page({ items: [], rects: [{ xPt: 0, yPt: 0, widthPt: 10, heightPt: 10, fill: "#ffffff", stroke: [0, 0, 0] }] })],
    });
    const fill = pdf.indexOf("q 1 1 1 rg 0 831.89 10 10 re f Q");
    const stroke = pdf.indexOf("q 0 0 0 RG 0.75 w 0 831.89 10 10 re S Q");
    expect(fill).toBeGreaterThan(-1);
    expect(stroke).toBeGreaterThan(fill);
  });

  it("paints a rect before the text run, and an onTop rect after it", async () => {
    const pdf = await build({
      pages: [
        page({
          rects: [
            { xPt: 0, yPt: 0, widthPt: 10, heightPt: 10, fill: "#ff0000" },
            { xPt: 0, yPt: 0, widthPt: 10, heightPt: 10, fill: "#0000ff", onTop: true },
          ],
        }),
      ],
    });
    const under = pdf.indexOf("q 1 0 0 rg");
    const text = pdf.indexOf(" Tj ");
    const over = pdf.indexOf("q 0 0 1 rg");
    expect(under).toBeGreaterThan(-1);
    expect(text).toBeGreaterThan(under);
    expect(over).toBeGreaterThan(text);
  });
});

describe("determinism (SPEC-104 AC-40)", () => {
  it("builds byte-identical files from the same input — the writer stamps no date", async () => {
    const input: PdfBuildInput = {
      title: "Deck",
      pages: [
        page({
          rects: [{ xPt: 0, yPt: 0, widthPt: 20, heightPt: 20, fill: [0.2, 0.4, 0.6], stroke: [0, 0, 0] }],
          headings: [{ level: 1, title: "Slide 1", yPt: 0 }],
        }),
      ],
      fonts: fonts(),
    };
    const first = await buildPdf(input);
    const second = await buildPdf(input);
    expect(equalBytes(first, second)).toBe(true);
    expect(new TextDecoder("latin1").decode(first)).not.toContain("/CreationDate");
  });
});

describe("pdfFileName", () => {
  it("appends .pdf to the trimmed title, falling back when blank", async () => {
    expect(pdfFileName("Report ", "Untitled document")).toBe("Report.pdf");
    expect(pdfFileName("  ", "Untitled document")).toBe("Untitled document.pdf");
  });
});

function pathItem(overrides: Partial<PdfPathItem> = {}): PdfPathItem {
  return {
    subpaths: [
      {
        start: { xPt: 10, yPt: 20 },
        segments: [{ toXPt: 110, toYPt: 20 }],
      },
    ],
    stroke: { color: [1, 0, 0], widthPt: 2 },
    ...overrides,
  };
}

// SPEC-070: connectors are the one drawing the page paints as SVG, so they
// reach the PDF as real vector paths rather than a rasterized bitmap.
describe("vector paths", () => {
  it("strokes a line, flipping y into PDF space", async () => {
    const pdf = await build({ pages: [page({ paths: [pathItem()] })] });
    // 841.89 - 20 = 821.89
    expect(pdf).toContain("1 0 0 RG 2 w [] 0 d 0 J 0 j 10 821.89 m");
    expect(pdf).toContain("110 821.89 l");
    expect(pdf).toContain("S Q");
  });

  it("fills a closed sub-path (an arrowhead) with `f`, not `S`", async () => {
    const pdf = await build({
      pages: [
        page({
          paths: [
            pathItem({
              stroke: undefined,
              fill: [0, 0, 1],
              subpaths: [
                {
                  start: { xPt: 10, yPt: 10 },
                  segments: [
                    { toXPt: 0, toYPt: 4 },
                    { toXPt: 0, toYPt: 16 },
                  ],
                  closed: true,
                },
              ],
            }),
          ],
        }),
      ],
    });
    expect(pdf).toContain("0 0 1 rg 10 831.89 m");
    expect(pdf).toContain("h\nf Q");
    expect(pdf).not.toContain("S Q");
  });

  it("writes a cubic segment with the `c` operator", async () => {
    const pdf = await build({
      pages: [
        page({
          paths: [
            pathItem({
              subpaths: [
                {
                  start: { xPt: 0, yPt: 0 },
                  segments: [
                    { toXPt: 100, toYPt: 0, c1XPt: 30, c1YPt: 40, c2XPt: 70, c2YPt: 40 },
                  ],
                },
              ],
            }),
          ],
        }),
      ],
    });
    expect(pdf).toContain("30 801.89 70 801.89 100 841.89 c");
  });

  it("carries the cap and join styles the painted element used", async () => {
    const pdf = await build({
      pages: [
        page({
          paths: [pathItem({ stroke: { color: [0, 0, 0], widthPt: 1, cap: 1, join: 2 } })],
        }),
      ],
    });
    expect(pdf).toContain("1 J 2 j");
  });

  it("emits the dash pattern the stroke carries", async () => {
    const pdf = await build({
      pages: [page({ paths: [pathItem({ stroke: { color: [0, 0, 0], widthPt: 1, dashPt: [4.5, 3] } })] })],
    });
    expect(pdf).toContain("[4.5 3] 0 d");
  });

  it("paints behind-text paths before the text and in-front ones after it", async () => {
    const pdf = await build({
      pages: [
        page({
          paths: [
            pathItem({ stroke: { color: [1, 0, 0], widthPt: 2 }, onTop: false }),
            pathItem({ stroke: { color: [0, 1, 0], widthPt: 2 }, onTop: true }),
          ],
        }),
      ],
    });
    const behind = pdf.indexOf("1 0 0 RG");
    const text = pdf.indexOf(" Tj ET");
    const front = pdf.indexOf("0 1 0 RG");
    expect(behind).toBeGreaterThan(-1);
    expect(behind).toBeLessThan(text);
    expect(text).toBeLessThan(front);
  });

  it("skips a path with neither stroke nor fill, and one with no geometry", async () => {
    const pdf = await build({
      pages: [
        page({
          paths: [
            pathItem({ stroke: undefined, fill: undefined }),
            pathItem({ subpaths: [] }),
          ],
        }),
      ],
    });
    expect(pdf).not.toContain(" m\n");
  });

  it("draws nothing extra for a page that carries no `paths` at all", async () => {
    const pdf = await build({ pages: [page()] });
    expect(pdf).not.toContain(" RG ");
    expect(pdf).toContain(" Tj ET");
  });
});
