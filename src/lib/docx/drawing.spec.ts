import { describe, it, expect } from "vitest";
import { IMAGE_RELATIONSHIP, drawingXml, parseDrawing } from "./drawing";
import { parseXml } from "../core/xml";
import type { Image } from "../documentModel";

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// `<w:drawing>` uses the unprefixed "w:" namespace, which — like every other
// WordprocessingML element — is declared once at word/document.xml's root in
// a real package, not redeclared on the drawing itself (drawingXml only
// declares the wp: prefix it introduces). A standalone parse needs a wrapper
// that supplies xmlns:w, mirroring the odf/drawing.spec.ts fixtures' own
// wrapper for the text:/style: namespaces drawFrameXml likewise omits.
function drawingElement(xml: string): Element {
  const wrapped = `<root xmlns:w="${WORD_NS}">${xml}</root>`;
  return parseXml(wrapped).documentElement.children[0];
}

/** A minimal 24-byte PNG (see imageDimensions.spec.ts's own builder) so
 * decodeImageDimensions inside parseDrawing has something real to read. */
function buildPngBytes(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(24);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  bytes.set([0x49, 0x48, 0x44, 0x52], 12);
  bytes[16] = (width >>> 24) & 0xff;
  bytes[17] = (width >>> 16) & 0xff;
  bytes[18] = (width >>> 8) & 0xff;
  bytes[19] = width & 0xff;
  bytes[20] = (height >>> 24) & 0xff;
  bytes[21] = (height >>> 16) & 0xff;
  bytes[22] = (height >>> 8) & 0xff;
  bytes[23] = height & 0xff;
  return bytes;
}

const PNG_BYTES = buildPngBytes(400, 200);
const MAIN_PART_PATH = "word/document.xml";

function relsMap(entries: Array<[string, { type: string; target: string }]>): Map<string, { id: string; type: string; target: string }> {
  const map = new Map<string, { id: string; type: string; target: string }>();
  for (const [id, { type, target }] of entries) {
    map.set(id, { id, type, target });
  }
  return map;
}

function partsMap(entries: Array<[string, Uint8Array]>): Map<string, Uint8Array> {
  return new Map(entries);
}

const IMAGE_REL = (target: string) => ({ type: IMAGE_RELATIONSHIP, target });

function inlinePictureXml(opts: { cx?: number; cy?: number; relId?: string; descr?: string } = {}): string {
  const cx = opts.cx ?? 3810000;
  const cy = opts.cy ?? 1905000;
  const descrAttr = opts.descr ? ` descr="${opts.descr}"` : "";
  return (
    `<w:drawing xmlns:wp="${WP_NS}">` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="1" name="Picture 1"${descrAttr}/>` +
    `<a:graphic xmlns:a="${A_NS}">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="${PIC_NS}">` +
    `<pic:blipFill><a:blip xmlns:r="${R_NS}" r:embed="${opts.relId ?? "rId1"}"/></pic:blipFill>` +
    `</pic:pic>` +
    `</a:graphicData>` +
    `</a:graphic>` +
    `</wp:inline>` +
    `</w:drawing>`
  );
}

function anchorPictureXml(opts: {
  wrapXml: string;
  behindDoc?: string;
  cx?: number;
  cy?: number;
  relId?: string;
}): string {
  const cx = opts.cx ?? 3810000;
  const cy = opts.cy ?? 1905000;
  return (
    `<w:drawing xmlns:wp="${WP_NS}">` +
    `<wp:anchor behindDoc="${opts.behindDoc ?? "0"}" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    opts.wrapXml +
    `<wp:docPr id="1" name="Picture 1"/>` +
    `<a:graphic xmlns:a="${A_NS}">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="${PIC_NS}">` +
    `<pic:blipFill><a:blip xmlns:r="${R_NS}" r:embed="${opts.relId ?? "rId1"}"/></pic:blipFill>` +
    `</pic:pic>` +
    `</a:graphicData>` +
    `</a:graphic>` +
    `</wp:anchor>` +
    `</w:drawing>`
  );
}

describe("parseDrawing — inline picture (AC-4, AC-7)", () => {
  it("parses extent, resolves the blip's relationship and reads alt text from docPr descr", () => {
    const el = drawingElement(inlinePictureXml({ descr: "A red circle" }));
    const rels = relsMap([["rId1", IMAGE_REL("media/image1.png")]]);
    const parts = partsMap([["word/media/image1.png", PNG_BYTES]]);
    const result = parseDrawing(el, rels, MAIN_PART_PATH, parts);
    expect(result).not.toBeNull();
    expect(result!.image.type).toBe("image");
    expect(result!.image.wrap).toBe("inline");
    expect(result!.image.mediaId).toBe("word/media/image1.png");
    expect(result!.image.altText).toBe("A red circle");
    expect(result!.image.widthCm).toBeCloseTo(3810000 / 360000, 6);
    expect(result!.image.heightCm).toBeCloseTo(1905000 / 360000, 6);
    // Natural size comes from decoding the actual PNG bytes.
    expect(result!.image.naturalWidthPx).toBe(400);
    expect(result!.image.naturalHeightPx).toBe(200);
    expect(result!.media).toEqual({
      mediaId: "word/media/image1.png",
      contentType: "image/png",
      data: expect.any(String),
    });
  });

  it("has no altText field when docPr carries no descr", () => {
    const el = drawingElement(inlinePictureXml());
    const rels = relsMap([["rId1", IMAGE_REL("media/image1.png")]]);
    const parts = partsMap([["word/media/image1.png", PNG_BYTES]]);
    const result = parseDrawing(el, rels, MAIN_PART_PATH, parts)!;
    expect(result.image.altText).toBeUndefined();
  });
});

describe("parseDrawing — floating wrap modes (AC-13, AC-14)", () => {
  const rels = relsMap([["rId1", IMAGE_REL("media/image1.png")]]);
  const parts = partsMap([["word/media/image1.png", PNG_BYTES]]);

  it("wrapSquare maps to 'square'", () => {
    const el = drawingElement(anchorPictureXml({ wrapXml: '<wp:wrapSquare wrapText="bothSides"/>' }));
    expect(parseDrawing(el, rels, MAIN_PART_PATH, parts)!.image.wrap).toBe("square");
  });

  it("wrapNone with behindDoc=1 maps to 'behind'", () => {
    const el = drawingElement(anchorPictureXml({ wrapXml: "<wp:wrapNone/>", behindDoc: "1" }));
    expect(parseDrawing(el, rels, MAIN_PART_PATH, parts)!.image.wrap).toBe("behind");
  });

  it("wrapNone with behindDoc=0 maps to 'infront'", () => {
    const el = drawingElement(anchorPictureXml({ wrapXml: "<wp:wrapNone/>", behindDoc: "0" }));
    expect(parseDrawing(el, rels, MAIN_PART_PATH, parts)!.image.wrap).toBe("infront");
  });

  it("returns null for an unsupported wrap mode (wrapTight)", () => {
    const el = drawingElement(anchorPictureXml({ wrapXml: '<wp:wrapTight wrapText="bothSides"/>' }));
    expect(parseDrawing(el, rels, MAIN_PART_PATH, parts)).toBeNull();
  });
});

describe("parseDrawing — null cases (SPEC-008 opaque fallback)", () => {
  const rels = relsMap([["rId1", IMAGE_REL("media/image1.png")]]);
  const parts = partsMap([["word/media/image1.png", PNG_BYTES]]);

  it("returns null for a non-picture graphicData (e.g. a chart)", () => {
    const xml =
      `<w:drawing xmlns:wp="${WP_NS}">` +
      `<wp:inline><wp:extent cx="3810000" cy="1905000"/><wp:docPr id="1" name="Chart 1"/>` +
      `<a:graphic xmlns:a="${A_NS}">` +
      `<a:graphicData uri="${CHART_NS}"><c:chart xmlns:c="${CHART_NS}" r:id="rId1" xmlns:r="${R_NS}"/></a:graphicData>` +
      `</a:graphic></wp:inline></w:drawing>`;
    expect(parseDrawing(drawingElement(xml), rels, MAIN_PART_PATH, parts)).toBeNull();
  });

  it("returns null when the drawing holds more than one wp: child", () => {
    const xml =
      `<w:drawing xmlns:wp="${WP_NS}">` +
      inlinePictureXml().replace(/^<w:drawing[^>]*>/, "").replace(/<\/w:drawing>$/, "") +
      inlinePictureXml().replace(/^<w:drawing[^>]*>/, "").replace(/<\/w:drawing>$/, "") +
      `</w:drawing>`;
    expect(parseDrawing(drawingElement(xml), rels, MAIN_PART_PATH, parts)).toBeNull();
  });

  it("returns null with no extent", () => {
    const xml =
      `<w:drawing xmlns:wp="${WP_NS}"><wp:inline><wp:docPr id="1" name="Picture 1"/>` +
      `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="x"><pic:pic xmlns:pic="${PIC_NS}">` +
      `<pic:blipFill><a:blip xmlns:r="${R_NS}" r:embed="rId1"/></pic:blipFill></pic:pic>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing>`;
    expect(parseDrawing(drawingElement(xml), rels, MAIN_PART_PATH, parts)).toBeNull();
  });

  it("returns null for a missing relationship id", () => {
    const el = drawingElement(inlinePictureXml({ relId: "rIdMissing" }));
    expect(parseDrawing(el, rels, MAIN_PART_PATH, parts)).toBeNull();
  });

  it("returns null when the relationship exists but has the wrong type", () => {
    const wrongTypeRels = relsMap([
      ["rId1", { type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", target: "media/image1.png" }],
    ]);
    const el = drawingElement(inlinePictureXml());
    expect(parseDrawing(el, wrongTypeRels, MAIN_PART_PATH, parts)).toBeNull();
  });

  it("returns null when the relationship resolves to a part with no bytes", () => {
    const el = drawingElement(inlinePictureXml());
    expect(parseDrawing(el, rels, MAIN_PART_PATH, partsMap([]))).toBeNull();
  });
});

describe("drawingXml round-trips through parseDrawing", () => {
  const rels = relsMap([["rId9", IMAGE_REL("media/image7.png")]]);
  const parts = partsMap([["word/media/image7.png", PNG_BYTES]]);

  function roundTrip(image: Image) {
    const xml = drawingXml(image, "rId9", 5);
    const parsed = parseDrawing(drawingElement(xml), rels, MAIN_PART_PATH, parts);
    return parsed;
  }

  it("round-trips an inline image", () => {
    const source: Image = {
      type: "image",
      mediaId: "word/media/image7.png",
      naturalWidthPx: 400,
      naturalHeightPx: 200,
      widthCm: 5,
      heightCm: 2.5,
      wrap: "inline",
      altText: "A diagram",
    };
    const parsed = roundTrip(source);
    expect(parsed).not.toBeNull();
    expect(parsed!.image.wrap).toBe("inline");
    expect(parsed!.image.widthCm).toBeCloseTo(5, 6);
    expect(parsed!.image.heightCm).toBeCloseTo(2.5, 6);
    expect(parsed!.image.altText).toBe("A diagram");
    expect(parsed!.image.mediaId).toBe("word/media/image7.png");
  });

  it("round-trips a floating (behind) image", () => {
    const source: Image = {
      type: "image",
      mediaId: "word/media/image7.png",
      naturalWidthPx: 400,
      naturalHeightPx: 200,
      widthCm: 8,
      heightCm: 4,
      wrap: "behind",
    };
    const parsed = roundTrip(source);
    expect(parsed).not.toBeNull();
    expect(parsed!.image.wrap).toBe("behind");
    expect(parsed!.image.widthCm).toBeCloseTo(8, 6);
    expect(parsed!.image.heightCm).toBeCloseTo(4, 6);
  });

  it("round-trips a floating (square) image", () => {
    const source: Image = {
      type: "image",
      mediaId: "word/media/image7.png",
      naturalWidthPx: 400,
      naturalHeightPx: 200,
      widthCm: 8,
      heightCm: 4,
      wrap: "square",
    };
    const parsed = roundTrip(source);
    expect(parsed!.image.wrap).toBe("square");
  });

  it("round-trips a floating (infront) image", () => {
    const source: Image = {
      type: "image",
      mediaId: "word/media/image7.png",
      naturalWidthPx: 400,
      naturalHeightPx: 200,
      widthCm: 8,
      heightCm: 4,
      wrap: "infront",
    };
    const parsed = roundTrip(source);
    expect(parsed!.image.wrap).toBe("infront");
  });
});

// The contentTypeForPartName / extensionForContentType tests moved to
// core/media.spec.ts with the helpers themselves (both formats share them).
