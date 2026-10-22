// Image placement: a floating picture's position, stacking order and text
// room, across both formats and the renderer.
//
// SPEC-024 modelled only an image's wrap mode. Neither reader read
// `wp:positionH`/`svg:x`, neither writer wrote them, and imageDisplayStyle
// pinned every floating image to `left: 0; top: 0` — so a picture Word had
// put halfway down the page opened at its anchor's origin, and the model's
// own commands (moveGroup, which does set an image's offsets) moved nothing
// visible. Image now carries the same DrawingPlacement fields Shape, Chart
// and Connector do, and renders through the same rules.
import { describe, it, expect } from "vitest";
import { drawingWrapSpacerStyle, imageDisplayStyle, shapeDisplayStyle } from "./renderFormatting";
import { parseDrawing, drawingXml } from "./docx/drawing";
import { parseDrawFrame, drawFrameXml, collectGraphicWrapStyles } from "./odf/drawing";
import { parseXml } from "./core/xml";
import type { Image, ImageWrapMode, Shape } from "./documentModel";

const PNG_BYTES = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x0a, 0x00, 0x00, 0x00, 0x14, 0x08, 0x02, 0x00, 0x00, 0x00,
]);

function image(overrides: Partial<Image> = {}): Image {
  return {
    type: "image",
    mediaId: "word/media/image1.png",
    naturalWidthPx: 10,
    naturalHeightPx: 20,
    widthCm: 3,
    heightCm: 2,
    wrap: "infront",
    ...overrides,
  };
}

describe("imageDisplayStyle — the same rules shapeDisplayStyle uses", () => {
  it("positions a floating image at its own offsets", () => {
    const style = imageDisplayStyle(image({ offsetXCm: 4.5, offsetYCm: 1.25 }));
    expect(style.position).toBe("absolute");
    expect(style.left).toBe("4.5cm");
    expect(style.top).toBe("1.25cm");
  });

  it("treats absent offsets as the anchor's own origin, never a recentre", () => {
    const style = imageDisplayStyle(image());
    expect(style.left).toBe("0cm");
    expect(style.top).toBe("0cm");
  });

  it("ignores offsets for an inline image, which sits in the flow", () => {
    const style = imageDisplayStyle(image({ wrap: "inline", offsetXCm: 4, offsetYCm: 4 }));
    expect(style.position).toBeUndefined();
    expect(style.left).toBeUndefined();
    expect(style.width).toBe("3cm");
  });

  it("stacks behind under the text and in-front above it, ordered by zOrder", () => {
    const behind = imageDisplayStyle(image({ wrap: "behind", zOrder: 3 }));
    const front = imageDisplayStyle(image({ wrap: "infront", zOrder: 3 }));
    expect(Number(behind.zIndex)).toBeLessThan(0);
    expect(Number(front.zIndex)).toBeGreaterThan(0);
    // Two far-apart tiers: a "behind" image never climbs above the text no
    // matter how high its zOrder goes.
    expect(Number(imageDisplayStyle(image({ wrap: "behind", zOrder: 9999 })).zIndex)).toBeLessThan(0);
  });

  it("puts an image and a shape at the same offsets in the same place", () => {
    // The point of the change: a picture and a shape the source file placed
    // together must land together.
    const shape: Shape = {
      type: "shape",
      geometry: "rect",
      widthCm: 3,
      heightCm: 2,
      wrap: "infront",
      offsetXCm: 4.5,
      offsetYCm: 1.25,
      zOrder: 2,
    };
    const imageStyle = imageDisplayStyle(image({ offsetXCm: 4.5, offsetYCm: 1.25, zOrder: 2 }));
    const shapeStyle = shapeDisplayStyle(shape);
    expect(imageStyle.left).toBe(shapeStyle.left);
    expect(imageStyle.top).toBe(shapeStyle.top);
    expect(imageStyle.zIndex).toBe(shapeStyle.zIndex);
  });

  it("reserves text room through the shared wrap spacer, honouring the offset", () => {
    // `square` used to be a bare `float: left`, which cannot sit at an
    // offset; the spacer is the same in-flow float every drawing kind uses.
    const spacer = drawingWrapSpacerStyle(image({ wrap: "square", offsetXCm: 1, offsetYCm: 2 }), 16);
    expect(spacer).not.toBeNull();
    expect(spacer!.float).toBe("left");
    expect(spacer!.shapeOutside).toContain("2cm");
    expect(drawingWrapSpacerStyle(image({ wrap: "inline" }), 16)).toBeNull();
    expect(drawingWrapSpacerStyle(image({ wrap: "infront" }), 16)).toBeNull();
  });
});

// --- .docx wire ------------------------------------------------------------

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const MAIN_PART_PATH = "word/document.xml";

function drawingElement(xml: string): Element {
  return parseXml(`<root xmlns:w="${WORD_NS}">${xml}</root>`).documentElement.children[0];
}

const docxRels = new Map([["rId5", { id: "rId5", type: IMAGE_REL_TYPE, target: "media/image1.png" }]]);
const docxParts = new Map([["word/media/image1.png", PNG_BYTES]]);

function readDocx(xml: string): Image {
  const parsed = parseDrawing(drawingElement(xml), docxRels, MAIN_PART_PATH, docxParts);
  expect(parsed).not.toBeNull();
  return parsed!.image;
}

/** A floating `w:drawing` picture at the given EMU offsets. */
function docxAnchor(options: { xEmu: number; yEmu: number; relativeHeight?: number }): string {
  const { xEmu, yEmu, relativeHeight = 251658243 } = options;
  return (
    '<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
    '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"' +
    ` relativeHeight="${relativeHeight}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    '<wp:simplePos x="0" y="0"/>' +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${xEmu}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${yEmu}</wp:posOffset></wp:positionV>` +
    '<wp:extent cx="1080000" cy="720000"/><wp:wrapNone/>' +
    '<wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr/>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId5"/>' +
    "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1080000" cy="720000"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    "</a:graphicData></a:graphic></wp:anchor></w:drawing>"
  );
}

describe("docx — a floating picture's position on the wire", () => {
  it("reads wp:positionH/positionV posOffset into offsets", () => {
    const read = readDocx(docxAnchor({ xEmu: 1620000, yEmu: 450000 }));
    expect(read.offsetXCm).toBeCloseTo(4.5, 6);
    expect(read.offsetYCm).toBeCloseTo(1.25, 6);
    expect(read.relativeToH).toBe("column");
    expect(read.relativeToV).toBe("paragraph");
  });

  it("reads relativeHeight into zOrder against the shared base", () => {
    expect(readDocx(docxAnchor({ xEmu: 0, yEmu: 0, relativeHeight: 251658247 })).zOrder).toBe(7);
  });

  it("leaves offsets absent for a relatively positioned anchor (wp:align, not posOffset)", () => {
    // The model carries no alignment, and "absent" means the wrap mode's
    // default — never a recentre.
    const aligned = docxAnchor({ xEmu: 0, yEmu: 0 })
      .replace("<wp:posOffset>0</wp:posOffset>", "<wp:align>center</wp:align>")
      .replace("<wp:posOffset>0</wp:posOffset>", "<wp:align>top</wp:align>");
    const read = readDocx(aligned);
    expect(read.offsetXCm).toBeUndefined();
    expect(read.offsetYCm).toBeUndefined();
  });

  it("round-trips offsets and zOrder through drawingXml", () => {
    const source = image({ wrap: "infront", offsetXCm: 4.5, offsetYCm: 1.25, zOrder: 7 });
    const reread = readDocx(drawingXml(source, "rId5", 1));
    expect(reread.offsetXCm).toBeCloseTo(4.5, 6);
    expect(reread.offsetYCm).toBeCloseTo(1.25, 6);
    expect(reread.zOrder).toBe(7);
  });

  it("writes no offsets on an inline picture, which has no anchor at all", () => {
    const xml = drawingXml(image({ wrap: "inline", offsetXCm: 4, offsetYCm: 4 }), "rId5", 1);
    expect(xml).not.toContain("posOffset");
    expect(readDocx(xml).offsetXCm).toBeUndefined();
  });
});

// --- .odt wire -------------------------------------------------------------

const ODF_NS =
  ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
  ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
  ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
  ' xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"' +
  ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' +
  ' xmlns:xlink="http://www.w3.org/1999/xlink"';

const odtParts = new Map([["Pictures/image1.png", PNG_BYTES]]);

/** Parses one `draw:frame` in a document that declares the wrap style it
 * names, so collectGraphicWrapStyles resolves it the way the reader does. */
function readOdf(frameXml: string, wrap: ImageWrapMode = "infront"): Image {
  const wrapProps =
    wrap === "square"
      ? 'style:wrap="parallel" style:number-wrapped-paragraphs="no-limit"'
      : 'style:wrap="run-through" style:run-through="foreground"';
  const doc = parseXml(
    `<office:document-content${ODF_NS}><office:automatic-styles>` +
      `<style:style style:name="gr1" style:family="graphic"><style:graphic-properties ${wrapProps}/></style:style>` +
      `</office:automatic-styles><office:body><office:text><text:p>${frameXml}</text:p></office:text></office:body>` +
      "</office:document-content>",
  );
  const frame = doc.getElementsByTagNameNS("urn:oasis:names:tc:opendocument:xmlns:drawing:1.0", "frame")[0];
  const parsed = parseDrawFrame(frame, odtParts, collectGraphicWrapStyles(doc, undefined));
  expect(parsed).not.toBeNull();
  return parsed!.image;
}

function odfFrame(attrs: string): string {
  return (
    `<draw:frame text:anchor-type="paragraph" draw:style-name="gr1"${attrs}` +
    ' svg:width="3cm" svg:height="2cm">' +
    '<draw:image xlink:href="Pictures/image1.png"/></draw:frame>'
  );
}

describe("odt — a floating picture's position on the wire", () => {
  it("reads svg:x/svg:y into offsets and draw:z-index into zOrder", () => {
    const read = readOdf(odfFrame(' draw:z-index="4" svg:x="4.5cm" svg:y="1.25cm"'));
    expect(read.offsetXCm).toBeCloseTo(4.5, 6);
    expect(read.offsetYCm).toBeCloseTo(1.25, 6);
    expect(read.zOrder).toBe(4);
    expect(read.relativeToH).toBe("column");
  });

  it("ignores position on an as-char frame, which sits in the flow", () => {
    const inlineFrame =
      '<draw:frame text:anchor-type="as-char" svg:x="4cm" svg:y="4cm" svg:width="3cm" svg:height="2cm">' +
      '<draw:image xlink:href="Pictures/image1.png"/></draw:frame>';
    const read = readOdf(inlineFrame);
    expect(read.wrap).toBe("inline");
    expect(read.offsetXCm).toBeUndefined();
  });

  it("round-trips offsets and z-index through drawFrameXml", () => {
    const source = image({ wrap: "infront", offsetXCm: 4.5, offsetYCm: 1.25, zOrder: 4 });
    const reread = readOdf(drawFrameXml(source, "Pictures/image1.png", "gr1"));
    expect(reread.offsetXCm).toBeCloseTo(4.5, 6);
    expect(reread.offsetYCm).toBeCloseTo(1.25, 6);
    expect(reread.zOrder).toBe(4);
  });

  it("writes no position on an inline frame", () => {
    const xml = drawFrameXml(image({ wrap: "inline", offsetXCm: 4, offsetYCm: 4 }), "Pictures/image1.png", undefined);
    expect(xml).not.toContain("svg:x");
    expect(xml).not.toContain("draw:z-index");
  });
});
