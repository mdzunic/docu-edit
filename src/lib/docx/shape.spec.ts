// SPEC-051: the DrawingML shape reader/writer. Mirrors docx/drawing.spec.ts's
// own fixture conventions (a `drawingElement` wrapper supplying the ambient
// `xmlns:w` a real word/document.xml would already declare).
import { describe, it, expect } from "vitest";
import { parseShapeDrawing, shapeDrawingXml, type ShapeTextContext, type ShapeTextWriteContext } from "./shape";
import { parseXml } from "./xml";
import { canonicalizeXml } from "../equality";
import { SHAPE_CATALOGUE, type Paragraph, type Shape } from "../documentModel";

const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";

// See docx/drawing.spec.ts's own drawingElement() for why this wrapper is
// needed: `<w:drawing>` uses the unprefixed "w:" namespace, normally
// declared once at word/document.xml's root, not redeclared on the drawing
// itself.
function drawingElement(xml: string): Element {
  const wrapped = `<root xmlns:w="${WORD_NS}">${xml}</root>`;
  return parseXml(wrapped).documentElement.children[0];
}

/** Minimal paragraph parse/write stubs — exercising shape.ts's own txbx
 * plumbing, not the full paragraph parser (out of scope for this file). A
 * paragraph round-trips as a single run's plain text. */
function escapeXmlText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

const readCtx: ShapeTextContext = {
  parseParagraph: (pEl: Element): Paragraph => {
    const texts = Array.from(pEl.getElementsByTagNameNS(WORD_NS, "t")).map((t) => t.textContent ?? "");
    const text = texts.join("");
    return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
  },
};

const writeCtx: ShapeTextWriteContext = {
  paragraphXml: (paragraph: Paragraph): string => {
    const text = paragraph.runs.map((r) => r.text).join("");
    return `<w:p><w:r><w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r></w:p>`;
  },
};

function baseShape(overrides: Partial<Shape> = {}): Shape {
  return {
    type: "shape",
    geometry: "rect",
    widthCm: 5,
    heightCm: 3,
    wrap: "inline",
    ...overrides,
  };
}

function roundTrip(shape: Shape, docPrId = 7) {
  const xml = shapeDrawingXml(shape, docPrId, writeCtx);
  return parseShapeDrawing(drawingElement(xml), readCtx);
}

// --- parseShapeDrawing — null cases (SPEC-008 opaque fallback) -----------------

describe("parseShapeDrawing — null cases", () => {
  function inlineShapeXml(prst = "rect"): string {
    return (
      `<w:drawing xmlns:wp="${WP_NS}">` +
      `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="1800000" cy="900000"/>` +
      `<wp:docPr id="1" name="Shape 1"/>` +
      `<a:graphic xmlns:a="${A_NS}">` +
      `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
      `<wps:wsp xmlns:wps="${WPS_NS}">` +
      `<wps:cNvSpPr/>` +
      `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="900000"/></a:xfrm>` +
      `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>` +
      `<a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>` +
      `</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`
    );
  }

  it("returns null when the drawing holds more than one wp: child", () => {
    const inner = inlineShapeXml().replace(/^<w:drawing[^>]*>/, "").replace(/<\/w:drawing>$/, "");
    const xml = `<w:drawing xmlns:wp="${WP_NS}">${inner}${inner}</w:drawing>`;
    expect(parseShapeDrawing(drawingElement(xml), readCtx)).toBeNull();
  });

  it("returns null for a container that is neither wp:inline nor wp:anchor", () => {
    const xml = `<w:drawing xmlns:wp="${WP_NS}"><wp:cNvGraphicFramePr/></w:drawing>`;
    expect(parseShapeDrawing(drawingElement(xml), readCtx)).toBeNull();
  });

  it("returns null with no extent", () => {
    const xml =
      `<w:drawing xmlns:wp="${WP_NS}"><wp:inline><wp:docPr id="1" name="Shape 1"/>` +
      `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="x">` +
      `<wps:wsp xmlns:wps="${WPS_NS}"><wps:spPr><a:prstGeom prst="rect"/></wps:spPr></wps:wsp>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing>`;
    expect(parseShapeDrawing(drawingElement(xml), readCtx)).toBeNull();
  });

  it("returns null for a negative extent", () => {
    const xml =
      `<w:drawing xmlns:wp="${WP_NS}"><wp:inline><wp:extent cx="-1" cy="900000"/><wp:docPr id="1" name="Shape 1"/>` +
      `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="x">` +
      `<wps:wsp xmlns:wps="${WPS_NS}"><wps:spPr><a:prstGeom prst="rect"/></wps:spPr></wps:wsp>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing>`;
    expect(parseShapeDrawing(drawingElement(xml), readCtx)).toBeNull();
  });

  it("returns null for a non-shape graphicData (e.g. a picture)", () => {
    const xml =
      `<w:drawing xmlns:wp="${WP_NS}"><wp:inline><wp:extent cx="1800000" cy="900000"/><wp:docPr id="1" name="Picture 1"/>` +
      `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="picture">` +
      `<pic:pic xmlns:pic="${PIC_NS}"/>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing>`;
    expect(parseShapeDrawing(drawingElement(xml), readCtx)).toBeNull();
  });

  it("returns null when spPr is missing", () => {
    const xml =
      `<w:drawing xmlns:wp="${WP_NS}"><wp:inline><wp:extent cx="1800000" cy="900000"/><wp:docPr id="1" name="Shape 1"/>` +
      `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="x">` +
      `<wps:wsp xmlns:wps="${WPS_NS}"><wps:cNvSpPr/></wps:wsp>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing>`;
    expect(parseShapeDrawing(drawingElement(xml), readCtx)).toBeNull();
  });

  it("returns null for an anchor with none of wrapSquare/wrapTopAndBottom/wrapNone", () => {
    const xml =
      `<w:drawing xmlns:wp="${WP_NS}"><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:extent cx="1800000" cy="900000"/><wp:wrapTight wrapText="bothSides"/><wp:docPr id="1" name="Shape 1"/>` +
      `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="x">` +
      `<wps:wsp xmlns:wps="${WPS_NS}"><wps:spPr><a:prstGeom prst="rect"/></wps:spPr></wps:wsp>` +
      `</a:graphicData></a:graphic></wp:anchor></w:drawing>`;
    expect(parseShapeDrawing(drawingElement(xml), readCtx)).toBeNull();
  });
});

// --- parseShapeDrawing — recognized shapes --------------------------------------

describe("parseShapeDrawing — geometry mapping", () => {
  it("maps every catalogue geometry's prst back to itself, round-tripped through the writer", () => {
    for (const entry of SHAPE_CATALOGUE) {
      const source = baseShape({ geometry: entry.geometry, wrap: "inline" });
      const parsed = roundTrip(source);
      expect(parsed, entry.geometry).not.toBeNull();
      expect(parsed!.shape.geometry).toBe(entry.geometry);
      expect(parsed!.shape.unsupportedGeometryXml).toBeUndefined();
    }
  });

  it("an unrecognised prst falls back to 'rect' and captures the verbatim prstGeom as unsupportedGeometryXml (AC-60)", () => {
    const xml =
      `<w:drawing xmlns:wp="${WP_NS}"><wp:inline><wp:extent cx="1800000" cy="900000"/><wp:docPr id="1" name="Shape 1"/>` +
      `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="x">` +
      `<wps:wsp xmlns:wps="${WPS_NS}"><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="900000"/></a:xfrm>` +
      `<a:prstGeom prst="cube"><a:avLst/></a:prstGeom></wps:spPr></wps:wsp>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing>`;
    const result = parseShapeDrawing(drawingElement(xml), readCtx);
    expect(result).not.toBeNull();
    expect(result!.shape.geometry).toBe("rect");
    expect(result!.shape.unsupportedGeometryXml).toBeDefined();
    expect(result!.shape.unsupportedGeometryXml).toContain('prst="cube"');
  });
});

// --- wrap modes and offsets ------------------------------------------------------

describe("parseShapeDrawing — wrap modes and offsets", () => {
  function anchorXml(opts: { wrapXml: string; behindDoc?: string; relativeHeight?: string; offsetX?: string; offsetY?: string }): string {
    const posH = opts.offsetX !== undefined ? `<wp:positionH relativeFrom="column"><wp:posOffset>${opts.offsetX}</wp:posOffset></wp:positionH>` : "";
    const posV = opts.offsetY !== undefined ? `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${opts.offsetY}</wp:posOffset></wp:positionV>` : "";
    return (
      `<w:drawing xmlns:wp="${WP_NS}"><wp:anchor behindDoc="${opts.behindDoc ?? "0"}" simplePos="0" locked="0"` +
      ` layoutInCell="1" allowOverlap="1"${opts.relativeHeight ? ` relativeHeight="${opts.relativeHeight}"` : ""}>` +
      posH +
      posV +
      `<wp:extent cx="1800000" cy="900000"/>${opts.wrapXml}<wp:docPr id="1" name="Shape 1"/>` +
      `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="x">` +
      `<wps:wsp xmlns:wps="${WPS_NS}"><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr></wps:wsp>` +
      `</a:graphicData></a:graphic></wp:anchor></w:drawing>`
    );
  }

  it("wrapSquare maps to 'square'", () => {
    const el = drawingElement(anchorXml({ wrapXml: '<wp:wrapSquare wrapText="bothSides"/>' }));
    expect(parseShapeDrawing(el, readCtx)!.shape.wrap).toBe("square");
  });

  it("wrapTopAndBottom maps to 'topAndBottom'", () => {
    const el = drawingElement(anchorXml({ wrapXml: "<wp:wrapTopAndBottom/>" }));
    expect(parseShapeDrawing(el, readCtx)!.shape.wrap).toBe("topAndBottom");
  });

  it("wrapNone with behindDoc=1 maps to 'behind'", () => {
    const el = drawingElement(anchorXml({ wrapXml: "<wp:wrapNone/>", behindDoc: "1" }));
    expect(parseShapeDrawing(el, readCtx)!.shape.wrap).toBe("behind");
  });

  it("wrapNone with behindDoc=0 maps to 'infront'", () => {
    const el = drawingElement(anchorXml({ wrapXml: "<wp:wrapNone/>", behindDoc: "0" }));
    expect(parseShapeDrawing(el, readCtx)!.shape.wrap).toBe("infront");
  });

  it("reads positionH/positionV posOffset into offsetXCm/offsetYCm", () => {
    const el = drawingElement(anchorXml({ wrapXml: '<wp:wrapSquare wrapText="bothSides"/>', offsetX: "360000", offsetY: "720000" }));
    const shape = parseShapeDrawing(el, readCtx)!.shape;
    expect(shape.offsetXCm).toBeCloseTo(1, 6);
    expect(shape.offsetYCm).toBeCloseTo(2, 6);
  });

  it("derives zOrder from relativeHeight (offset by the base 251658240)", () => {
    const el = drawingElement(anchorXml({ wrapXml: '<wp:wrapSquare wrapText="bothSides"/>', relativeHeight: "251658245" }));
    expect(parseShapeDrawing(el, readCtx)!.shape.zOrder).toBe(5);
  });

  it("an inline shape has no offsets or zOrder", () => {
    const shape = roundTrip(baseShape({ wrap: "inline" }))!.shape;
    expect(shape.offsetXCm).toBeUndefined();
    expect(shape.offsetYCm).toBeUndefined();
    expect(shape.zOrder).toBeUndefined();
  });
});

// --- fill / outline ---------------------------------------------------------------

describe("parseShapeDrawing — fill and outline", () => {
  function shapeWithSpPrXml(spPrInner: string): string {
    return (
      `<w:drawing xmlns:wp="${WP_NS}"><wp:inline><wp:extent cx="1800000" cy="900000"/><wp:docPr id="1" name="Shape 1"/>` +
      `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="x">` +
      `<wps:wsp xmlns:wps="${WPS_NS}"><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${spPrInner}</wps:spPr></wps:wsp>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing>`
    );
  }

  it("parses a solid fill colour", () => {
    const xml = shapeWithSpPrXml('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>');
    expect(parseShapeDrawing(drawingElement(xml), readCtx)!.shape.fill).toBe("#ff0000");
  });

  it("no solidFill means no fill", () => {
    const xml = shapeWithSpPrXml("<a:noFill/>");
    expect(parseShapeDrawing(drawingElement(xml), readCtx)!.shape.fill).toBeUndefined();
  });

  it("parses outline colour, width and solid/dash/dot", () => {
    const solidXml = shapeWithSpPrXml('<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>');
    const solid = parseShapeDrawing(drawingElement(solidXml), readCtx)!.shape.outline!;
    expect(solid).toEqual({ color: "#000000", widthPt: 1, dash: "solid" });

    const dashXml = shapeWithSpPrXml('<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/></a:ln>');
    expect(parseShapeDrawing(drawingElement(dashXml), readCtx)!.shape.outline!.dash).toBe("dash");

    const dotXml = shapeWithSpPrXml('<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="sysDot"/></a:ln>');
    expect(parseShapeDrawing(drawingElement(dotXml), readCtx)!.shape.outline!.dash).toBe("dot");
  });

  it("a:ln with a:noFill means no outline", () => {
    const xml = shapeWithSpPrXml("<a:ln><a:noFill/></a:ln>");
    expect(parseShapeDrawing(drawingElement(xml), readCtx)!.shape.outline).toBeUndefined();
  });

  it("no a:ln at all means no outline", () => {
    const xml = shapeWithSpPrXml("");
    expect(parseShapeDrawing(drawingElement(xml), readCtx)!.shape.outline).toBeUndefined();
  });

  it("defaults outline width to 0.75pt when a:ln has no @w", () => {
    const xml = shapeWithSpPrXml('<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>');
    expect(parseShapeDrawing(drawingElement(xml), readCtx)!.shape.outline!.widthPt).toBeCloseTo(0.75, 6);
  });
});

// --- rotation / altText / group ordinal -------------------------------------------

describe("parseShapeDrawing — rotation, alt text, groupOrdinal", () => {
  it("reads a:xfrm/@rot as whole degrees", () => {
    const xml =
      `<w:drawing xmlns:wp="${WP_NS}"><wp:inline><wp:extent cx="1800000" cy="900000"/><wp:docPr id="1" name="Shape 1"/>` +
      `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="x">` +
      `<wps:wsp xmlns:wps="${WPS_NS}"><wps:spPr><a:xfrm rot="2700000"><a:off x="0" y="0"/><a:ext cx="1800000" cy="900000"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr></wps:wsp>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing>`;
    expect(parseShapeDrawing(drawingElement(xml), readCtx)!.shape.rotationDeg).toBe(45);
  });

  it("reads docPr/@descr as alt text", () => {
    const shape = roundTrip(baseShape({ altText: "A warning box" }))!.shape;
    expect(shape.altText).toBe("A warning box");
  });

  it("has no altText field when docPr carries no descr", () => {
    const shape = roundTrip(baseShape())!.shape;
    expect(shape.altText).toBeUndefined();
  });

  it("round-trips groupOrdinal through the docu-edit extension attribute", () => {
    const shape = roundTrip(baseShape({ wrap: "square", groupOrdinal: 4 }))!.shape;
    expect(shape.groupOrdinal).toBe(4);
  });

  it("omits the groupOrdinal attribute entirely when absent", () => {
    const xml = shapeDrawingXml(baseShape(), 1, writeCtx);
    expect(xml).not.toContain("de:group");
  });
});

// --- text frame (wps:txbx) ---------------------------------------------------------

describe("parseShapeDrawing — text frame (wps:txbx)", () => {
  it("no txbx means no text field", () => {
    const shape = roundTrip(baseShape())!.shape;
    expect(shape.text).toBeUndefined();
  });

  it("round-trips paragraph text and vAlign through wps:txbx/wps:bodyPr", () => {
    const source = baseShape({ text: { blocks: [{ type: "paragraph", runs: [{ text: "hello shape", formatting: {} }] }], vAlign: "middle" } });
    const shape = roundTrip(source)!.shape;
    expect(shape.text?.blocks).toHaveLength(1);
    expect((shape.text!.blocks[0] as Paragraph).runs[0].text).toBe("hello shape");
    expect(shape.text?.vAlign).toBe("middle");
  });

  it("round-trips top and bottom vAlign too", () => {
    const top = roundTrip(baseShape({ text: { blocks: [{ type: "paragraph", runs: [] }], vAlign: "top" } }))!.shape;
    expect(top.text?.vAlign).toBe("top");
    const bottom = roundTrip(baseShape({ text: { blocks: [{ type: "paragraph", runs: [] }], vAlign: "bottom" } }))!.shape;
    expect(bottom.text?.vAlign).toBe("bottom");
  });

  it("round-trips multiple paragraphs", () => {
    const source = baseShape({
      text: {
        blocks: [
          { type: "paragraph", runs: [{ text: "first", formatting: {} }] },
          { type: "paragraph", runs: [{ text: "second", formatting: {} }] },
        ],
        vAlign: "top",
      },
    });
    const shape = roundTrip(source)!.shape;
    expect(shape.text?.blocks.map((b) => (b as Paragraph).runs[0]?.text)).toEqual(["first", "second"]);
  });
});

// --- shapeDrawingXml round trip through parseShapeDrawing --------------------------

describe("shapeDrawingXml round-trips through parseShapeDrawing", () => {
  it("round-trips an inline shape's size and geometry", () => {
    const source = baseShape({ geometry: "ellipse", widthCm: 4, heightCm: 6, wrap: "inline" });
    const parsed = roundTrip(source);
    expect(parsed).not.toBeNull();
    expect(parsed!.shape.wrap).toBe("inline");
    expect(parsed!.shape.geometry).toBe("ellipse");
    expect(parsed!.shape.widthCm).toBeCloseTo(4, 6);
    expect(parsed!.shape.heightCm).toBeCloseTo(6, 6);
  });

  it("round-trips a floating (square) shape's offsets", () => {
    const source = baseShape({ wrap: "square", offsetXCm: 3, offsetYCm: 1.5, relativeToH: "column", relativeToV: "paragraph" });
    const parsed = roundTrip(source)!;
    expect(parsed.shape.wrap).toBe("square");
    expect(parsed.shape.offsetXCm).toBeCloseTo(3, 5);
    expect(parsed.shape.offsetYCm).toBeCloseTo(1.5, 5);
  });

  it("round-trips a behind-wrapped shape", () => {
    const source = baseShape({ wrap: "behind", zOrder: -2 });
    const parsed = roundTrip(source)!;
    expect(parsed.shape.wrap).toBe("behind");
    expect(parsed.shape.zOrder).toBe(-2);
  });

  it("round-trips an infront-wrapped shape", () => {
    const source = baseShape({ wrap: "infront", zOrder: 3 });
    const parsed = roundTrip(source)!;
    expect(parsed.shape.wrap).toBe("infront");
    expect(parsed.shape.zOrder).toBe(3);
  });

  it("round-trips a topAndBottom-wrapped shape", () => {
    const source = baseShape({ wrap: "topAndBottom" });
    const parsed = roundTrip(source)!;
    expect(parsed.shape.wrap).toBe("topAndBottom");
  });

  it("round-trips rotation", () => {
    const source = baseShape({ rotationDeg: 270 });
    const parsed = roundTrip(source)!;
    expect(parsed.shape.rotationDeg).toBe(270);
  });

  it("round-trips fill", () => {
    const source = baseShape({ fill: "#336699" });
    const parsed = roundTrip(source)!;
    expect(parsed.shape.fill).toBe("#336699");
  });

  it("round-trips an outline (colour/width/dash)", () => {
    const source = baseShape({ outline: { color: "#123456", widthPt: 2.25, dash: "dot" } });
    const parsed = roundTrip(source)!;
    expect(parsed.shape.outline?.color).toBe("#123456");
    expect(parsed.shape.outline?.widthPt).toBeCloseTo(2.25, 3);
    expect(parsed.shape.outline?.dash).toBe("dot");
  });

  it("a Line's spPr never carries a fillXml element, even when a fill colour is set (fill is meaningless for a zero-area geometry)", () => {
    const xml = shapeDrawingXml(baseShape({ geometry: "line", fill: "#ff0000", outline: { color: "#000000", widthPt: 1, dash: "solid" } }), 1, writeCtx);
    // The only a:solidFill in the document is the outline's own — fillXml
    // is skipped entirely for a line geometry (isLine short-circuits it),
    // so the shape's #ff0000 fill never appears anywhere in the XML.
    expect(xml.match(/<a:solidFill>/g)).toHaveLength(1);
    expect(xml).not.toContain("ff0000");
    expect(xml).not.toContain("a:noFill");
  });

  it("round-trips an unsupported geometry's XML verbatim, marking the geometry field as the rect stand-in (AC-60, AC-61)", () => {
    const unsupportedXml = `<a:prstGeom xmlns:a="${A_NS}" prst="cube"><a:avLst/></a:prstGeom>`;
    const source = baseShape({ geometry: "rect", unsupportedGeometryXml: unsupportedXml });
    const xml = shapeDrawingXml(source, 1, writeCtx);
    // The writer never computes a normal prstGeom for the shape's own
    // `geometry` field once unsupportedGeometryXml is set.
    expect(xml).not.toContain('prst="rect"');
    expect(xml).toContain('prst="cube"');

    const parsed = parseShapeDrawing(drawingElement(xml), readCtx)!;
    expect(parsed.shape.geometry).toBe("rect");
    expect(parsed.shape.unsupportedGeometryXml).toBeDefined();
    expect(canonicalizeXml(parsed.shape.unsupportedGeometryXml!)).toBe(canonicalizeXml(unsupportedXml));
  });
});
