// SPEC-051: the DrawingML shape reader/writer. Mirrors docx/drawing.spec.ts's
// own fixture conventions (a `drawingElement` wrapper supplying the ambient
// `xmlns:w` a real word/document.xml would already declare).
import { DEFAULT_THEME } from "../core/theme";
import { describe, it, expect } from "vitest";
import { parseShapeDrawing, shapeDrawingXml, type ShapeTextContext, type ShapeTextWriteContext } from "./shape";
import { parseXml } from "../core/xml";
import { canonicalizeXml } from "../core/xml";
import { type Paragraph, type Shape } from "../documentModel";
import { SHAPE_CATALOGUE } from "../core/shapes";

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
  theme: DEFAULT_THEME,
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

// --- adjustments (SPEC-069 AC-6a, AC-7b) ---------------------------------------

describe("shape adjustments round-trip through a:avLst/a:gd", () => {
  // One geometry per adjustment-bearing group (AC-6c).
  const cases: { geometry: Shape["geometry"]; adjustments: number[] }[] = [
    { geometry: "roundRect", adjustments: [0.25] }, // Basic shapes
    { geometry: "rightArrow", adjustments: [0.5, 0.3] }, // Block arrows
    { geometry: "star6", adjustments: [0.4] }, // Stars and banners
    { geometry: "wedgeRectCallout", adjustments: [0.2, 0.9] }, // Callouts
  ];

  for (const { geometry, adjustments } of cases) {
    it(`writes a:gd@fmla and reads it back for ${geometry}`, () => {
      const xml = shapeDrawingXml(baseShape({ geometry, adjustments }), 1, writeCtx);
      expect(xml).toContain('<a:gd name="');
      expect(xml).toContain('fmla="val ');
      const parsed = roundTrip(baseShape({ geometry, adjustments }));
      expect(parsed).not.toBeNull();
      expect(parsed!.shape.adjustments).toBeDefined();
      parsed!.shape.adjustments!.forEach((value, i) => expect(value).toBeCloseTo(adjustments[i], 4));
    });
  }

  it("writes an empty a:avLst for a shape with no adjustments (byte-unchanged)", () => {
    const xml = shapeDrawingXml(baseShape({ geometry: "roundRect" }), 1, writeCtx);
    expect(xml).toContain("<a:avLst/>");
    expect(roundTrip(baseShape({ geometry: "roundRect" }))!.shape.adjustments).toBeUndefined();
  });

  it("preserves an out-of-range adjustment value verbatim on a same-format save (AC-7b)", () => {
    // 2.0 is well outside roundRect's 0..0.5 range; it must survive unchanged.
    const parsed = roundTrip(baseShape({ geometry: "roundRect", adjustments: [2] }));
    expect(parsed!.shape.adjustments![0]).toBeCloseTo(2, 4);
  });
});

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
      `<a:prstGeom prst="gear9"><a:avLst/></a:prstGeom></wps:spPr></wps:wsp>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing>`;
    const result = parseShapeDrawing(drawingElement(xml), readCtx);
    expect(result).not.toBeNull();
    expect(result!.shape.geometry).toBe("rect");
    expect(result!.shape.unsupportedGeometryXml).toBeDefined();
    expect(result!.shape.unsupportedGeometryXml).toContain('prst="gear9"');
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
    expect(parseShapeDrawing(drawingElement(xml), readCtx)!.shape.fill).toEqual({ kind: "solid", color: "#ff0000" });
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
    const source = baseShape({ fill: { kind: "solid", color: "#336699" } });
    const parsed = roundTrip(source)!;
    expect(parsed.shape.fill).toEqual({ kind: "solid", color: "#336699" });
  });

  it("round-trips an outline (colour/width/dash)", () => {
    const source = baseShape({ outline: { color: "#123456", widthPt: 2.25, dash: "dot" } });
    const parsed = roundTrip(source)!;
    expect(parsed.shape.outline?.color).toBe("#123456");
    expect(parsed.shape.outline?.widthPt).toBeCloseTo(2.25, 3);
    expect(parsed.shape.outline?.dash).toBe("dot");
  });

  it("a Line's spPr never carries a fillXml element, even when a fill colour is set (fill is meaningless for a zero-area geometry)", () => {
    const xml = shapeDrawingXml(baseShape({ geometry: "line", fill: { kind: "solid", color: "#ff0000" }, outline: { color: "#000000", widthPt: 1, dash: "solid" } }), 1, writeCtx);
    // The only a:solidFill in the document is the outline's own — fillXml
    // is skipped entirely for a line geometry (isLine short-circuits it),
    // so the shape's #ff0000 fill never appears anywhere in the XML.
    expect(xml.match(/<a:solidFill>/g)).toHaveLength(1);
    expect(xml).not.toContain("ff0000");
    expect(xml).not.toContain("a:noFill");
  });

  it("round-trips an unsupported geometry's XML verbatim, marking the geometry field as the rect stand-in (AC-60, AC-61)", () => {
    const unsupportedXml = `<a:prstGeom xmlns:a="${A_NS}" prst="gear9"><a:avLst/></a:prstGeom>`;
    const source = baseShape({ geometry: "rect", unsupportedGeometryXml: unsupportedXml });
    const xml = shapeDrawingXml(source, 1, writeCtx);
    // The writer never computes a normal prstGeom for the shape's own
    // `geometry` field once unsupportedGeometryXml is set.
    expect(xml).not.toContain('prst="rect"');
    expect(xml).toContain('prst="gear9"');

    const parsed = parseShapeDrawing(drawingElement(xml), readCtx)!;
    expect(parsed.shape.geometry).toBe("rect");
    expect(parsed.shape.unsupportedGeometryXml).toBeDefined();
    expect(canonicalizeXml(parsed.shape.unsupportedGeometryXml!)).toBe(canonicalizeXml(unsupportedXml));
  });

  it("does not retroactively reinterpret a preserved geometry now in the catalogue on save (AC-10)", () => {
    // 'cube' joined the catalogue in SPEC-069, but a stored shape that still
    // carries its verbatim prstGeom in unsupportedGeometryXml (from an
    // earlier app version) is not upgraded: a same-format save re-emits the
    // preserved bytes exactly and never computes a fresh cube prstGeom from
    // the geometry field. (Re-opening the source file is what would pick up
    // the catalogue — that is a new read, not this stored model.)
    const unsupportedXml = `<a:prstGeom xmlns:a="${A_NS}" prst="cube"><a:avLst/></a:prstGeom>`;
    const source = baseShape({ geometry: "rect", unsupportedGeometryXml: unsupportedXml });
    const xml = shapeDrawingXml(source, 1, writeCtx);
    // Exactly one prstGeom, and it is the preserved one — no second, upgraded
    // <a:prstGeom> derived from the catalogue.
    expect(xml.match(/<a:prstGeom/g)).toHaveLength(1);
    expect(xml).toContain('prst="cube"');
  });
});

// --- custom geometry / paths (SPEC-071) -------------------------------------

describe("path shapes round-trip through a:custGeom (SPEC-071 AC-26)", () => {
  const path = {
    subpaths: [
      {
        steps: [
          { op: "move" as const, x: 0, y: 1 },
          { op: "line" as const, x: 1, y: 1 },
          { op: "cubic" as const, x1: 0.75, y1: 0.5, x2: 0.25, y2: 0.5, x: 0.5, y: 0 },
          { op: "close" as const },
        ],
        closed: true,
      },
    ],
  };

  it("emits a:custGeom with moveTo/lnTo/cubicBezTo/close on a 21600 grid", () => {
    const xml = shapeDrawingXml(baseShape({ geometry: "rect", path, fill: { kind: "solid", color: "#ff0000" } }), 7, writeCtx);
    expect(xml).toContain("<a:custGeom>");
    expect(xml).toContain('<a:path w="21600" h="21600">');
    expect(xml).toContain("<a:moveTo><a:pt");
    expect(xml).toContain("<a:cubicBezTo><a:pt");
    expect(xml).toContain("<a:close/>");
    // The three lists the spec forbids alongside a modelled path.
    expect(xml).not.toContain("<a:gdLst");
    expect(xml).not.toContain("<a:ahLst");
    expect(xml).not.toContain("<a:cxnLst");
    expect(xml).not.toContain("<a:prstGeom");
  });

  it("reads back into an equal ShapePath, no unsupportedGeometryXml", () => {
    const result = roundTrip(baseShape({ geometry: "rect", path, fill: { kind: "solid", color: "#ff0000" } }));
    expect(result?.shape.path).toEqual(path);
    expect(result?.shape.unsupportedGeometryXml).toBeUndefined();
    expect(result?.shape.geometry).toBe("rect");
  });

  it("converts a:quadBezTo to a cubic on read and preserves the source", () => {
    const custGeom =
      `<wps:wsp xmlns:wps="${WPS_NS}"><wps:cNvSpPr/><wps:spPr xmlns:a="${A_NS}">` +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="360000" cy="360000"/></a:xfrm>` +
      `<a:custGeom><a:pathLst><a:path w="21600" h="21600">` +
      `<a:moveTo><a:pt x="0" y="0"/></a:moveTo>` +
      `<a:quadBezTo><a:pt x="10800" y="21600"/><a:pt x="21600" y="0"/></a:quadBezTo>` +
      `<a:close/></a:path></a:pathLst></a:custGeom>` +
      `</wps:spPr></wps:wsp>`;
    const drawing = drawingElement(
      `<w:drawing xmlns:wp="${WP_NS}"><wp:inline><wp:extent cx="360000" cy="360000"/>` +
        `<wp:docPr id="1" name="s"/><a:graphic xmlns:a="${A_NS}">` +
        `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${custGeom}</a:graphicData>` +
        `</a:graphic></wp:inline></w:drawing>`,
    );
    const result = parseShapeDrawing(drawing, readCtx);
    const steps = result?.shape.path?.subpaths[0].steps ?? [];
    expect(steps[1].op).toBe("cubic");
    // Exact degree-elevated control points: p0=(0,0), c=(0.5,1), p1=(1,0).
    if (steps[1].op === "cubic") {
      expect(steps[1].x1).toBeCloseTo(1 / 3);
      expect(steps[1].y1).toBeCloseTo(2 / 3);
    }
    expect(result?.shape.pathSource?.format).toBe("docx");
    expect(result?.shape.unsupportedGeometryXml).toBeUndefined();
  });

  it("re-emits the preserved source verbatim on a same-format save", () => {
    const shape = baseShape({
      geometry: "rect",
      path,
      pathSource: { format: "docx", xml: "<a:custGeom><a:pathLst><a:path w=\"21600\" h=\"21600\"><a:arcTo wR=\"1\" hR=\"1\" stAng=\"0\" swAng=\"5400000\"/></a:path></a:pathLst></a:custGeom>" },
    });
    const xml = shapeDrawingXml(shape, 7, writeCtx);
    expect(xml).toContain("<a:arcTo");
  });

  it("preserves a guide-driven custGeom verbatim (no path modelled)", () => {
    const custGeom =
      `<wps:wsp xmlns:wps="${WPS_NS}"><wps:cNvSpPr/><wps:spPr xmlns:a="${A_NS}">` +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="360000" cy="360000"/></a:xfrm>` +
      `<a:custGeom><a:gdLst><a:gd name="x" fmla="*/ w 1 2"/></a:gdLst><a:pathLst><a:path w="21600" h="21600">` +
      `<a:moveTo><a:pt x="x" y="0"/></a:moveTo><a:close/></a:path></a:pathLst></a:custGeom>` +
      `</wps:spPr></wps:wsp>`;
    const drawing = drawingElement(
      `<w:drawing xmlns:wp="${WP_NS}"><wp:inline><wp:extent cx="360000" cy="360000"/>` +
        `<wp:docPr id="1" name="s"/><a:graphic xmlns:a="${A_NS}">` +
        `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${custGeom}</a:graphicData>` +
        `</a:graphic></wp:inline></w:drawing>`,
    );
    const result = parseShapeDrawing(drawing, readCtx);
    expect(result?.shape.path).toBeUndefined();
    expect(result?.shape.unsupportedGeometryXml).toContain("custGeom");
  });
});

// --- SPEC-073: the widened styling model ------------------------------------

describe("SPEC-073 fill round-trip (AC-32)", () => {
  it("writes a:solidFill with a:alpha for a semi-transparent solid and reads it back", () => {
    const source = baseShape({ fill: { kind: "solid", color: "#336699", alpha: 0.4 } });
    const xml = shapeDrawingXml(source, 1, writeCtx);
    expect(xml).toContain('<a:solidFill><a:srgbClr val="336699"><a:alpha val="40000"/></a:srgbClr></a:solidFill>');
    expect(roundTrip(source)!.shape.fill).toEqual({ kind: "solid", color: "#336699", alpha: 0.4 });
  });

  it("writes a linear a:gradFill (gsLst positions, a:lin angle) and reads it back", () => {
    const source = baseShape({
      fill: {
        kind: "gradient",
        gradient: {
          type: "linear",
          angleDeg: 90,
          stops: [
            { position: 0, color: "#4472c4" },
            { position: 0.5, color: "#8faadc", alpha: 0.5 },
            { position: 1, color: "#ffffff" },
          ],
        },
      },
    });
    const xml = shapeDrawingXml(source, 1, writeCtx);
    expect(xml).toContain("<a:gradFill><a:gsLst>");
    expect(xml).toContain('<a:gs pos="0"><a:srgbClr val="4472C4"/></a:gs>');
    expect(xml).toContain('<a:gs pos="50000"><a:srgbClr val="8FAADC"><a:alpha val="50000"/></a:srgbClr></a:gs>');
    expect(xml).toContain('<a:lin ang="5400000" scaled="1"/>');
    expect(roundTrip(source)!.shape.fill).toEqual(source.fill);
  });

  it("writes a radial gradient as a:path path=circle and reads it back", () => {
    const source = baseShape({
      fill: {
        kind: "gradient",
        gradient: {
          type: "radial",
          stops: [
            { position: 0, color: "#000000" },
            { position: 1, color: "#ffffff" },
          ],
        },
      },
    });
    const xml = shapeDrawingXml(source, 1, writeCtx);
    expect(xml).toContain('<a:path path="circle">');
    expect(roundTrip(source)!.shape.fill).toEqual(source.fill);
  });

  it("writes a:pattFill with prst/fgClr/bgClr and reads it back", () => {
    const source = baseShape({
      fill: { kind: "pattern", preset: "ltUpDiag", foreground: "#4472c4", background: "#ffffff" },
    });
    const xml = shapeDrawingXml(source, 1, writeCtx);
    expect(xml).toContain('<a:pattFill prst="ltUpDiag">');
    expect(xml).toContain('<a:fgClr><a:srgbClr val="4472C4"/></a:fgClr>');
    expect(xml).toContain('<a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr>');
    expect(roundTrip(source)!.shape.fill).toEqual(source.fill);
  });

  it("writes a picture fill as a:blipFill via the media plan's relId and reads it back through the resolver (AC-4)", () => {
    const source = baseShape({ fill: { kind: "picture", mediaId: "word/media/image1.png", mode: "tile", alpha: 0.8 } });
    const xml = shapeDrawingXml(source, 1, { ...writeCtx, relIdForMedia: () => "rId9" });
    expect(xml).toContain('<a:blipFill><a:blip r:embed="rId9"');
    expect(xml).toContain('<a:alphaModFix amt="80000"/>');
    expect(xml).toContain("<a:tile ");
    const parsed = parseShapeDrawing(drawingElement(xml), {
      ...readCtx,
      resolveBlip: (embedId) => (embedId === "rId9" ? "word/media/image1.png" : null),
    });
    expect(parsed!.shape.fill).toEqual(source.fill);
  });

  it("a picture fill whose blip cannot resolve is preserved verbatim, not dropped (AC-35)", () => {
    const source = baseShape({ fill: { kind: "picture", mediaId: "word/media/image1.png", mode: "stretch" } });
    const xml = shapeDrawingXml(source, 1, { ...writeCtx, relIdForMedia: () => "rId9" });
    const parsed = parseShapeDrawing(drawingElement(xml), readCtx)!;
    expect(parsed.shape.fill).toBeUndefined();
    expect(parsed.shape.unsupportedStyleXml).toContain("blipFill");
  });
});

describe("SPEC-073 outline round-trip (AC-2, AC-32)", () => {
  it("writes alpha, cap, compound, join and the extended dashes and reads them back", () => {
    const source = baseShape({
      outline: {
        color: "#2e5395",
        widthPt: 2,
        dash: "longDashDot",
        alpha: 0.6,
        join: "bevel",
        cap: "round",
        compound: "double",
      },
    });
    const xml = shapeDrawingXml(source, 1, writeCtx);
    expect(xml).toContain('cap="rnd"');
    expect(xml).toContain('cmpd="dbl"');
    expect(xml).toContain('<a:prstDash val="lgDashDot"/>');
    expect(xml).toContain("<a:bevel/>");
    expect(xml).toContain('<a:alpha val="60000"/>');
    expect(roundTrip(source)!.shape.outline).toEqual(source.outline);
  });

  it("the three legacy dash values keep their spelling on the wire (AC-2)", () => {
    for (const [dash, wire] of [
      ["solid", null],
      ["dash", 'val="dash"'],
      ["dot", 'val="sysDot"'],
    ] as const) {
      const xml = shapeDrawingXml(baseShape({ outline: { color: "#000000", widthPt: 1, dash } }), 1, writeCtx);
      if (wire) {
        expect(xml).toContain(wire);
      } else {
        expect(xml).not.toContain("<a:prstDash");
      }
    }
  });

  it("round join and miter (absent) round-trip", () => {
    const round = baseShape({ outline: { color: "#000000", widthPt: 1, dash: "solid", join: "round" } });
    expect(shapeDrawingXml(round, 1, writeCtx)).toContain("<a:round/>");
    expect(roundTrip(round)!.shape.outline?.join).toBe("round");
    const miter = baseShape({ outline: { color: "#000000", widthPt: 1, dash: "solid" } });
    expect(roundTrip(miter)!.shape.outline?.join).toBeUndefined();
  });
});

describe("SPEC-073 arrowheads on line geometries (AC-13, AC-14, AC-15, AC-32)", () => {
  it("writes sized a:headEnd/a:tailEnd for an explicit double-headed line and reads them back", () => {
    const source = baseShape({
      geometry: "line",
      heightCm: 0,
      outline: {
        color: "#2e5395",
        widthPt: 1,
        dash: "solid",
        startArrow: { kind: "open", size: "large" },
        endArrow: { kind: "stealth", size: "small" },
      },
    });
    const xml = shapeDrawingXml(source, 1, writeCtx);
    expect(xml).toContain('<a:headEnd type="arrow" w="lg" len="lg"/>');
    expect(xml).toContain('<a:tailEnd type="stealth" w="sm" len="sm"/>');
    const parsed = roundTrip(source)!;
    expect(parsed.shape.outline?.startArrow).toEqual({ kind: "open", size: "large" });
    expect(parsed.shape.outline?.endArrow).toEqual({ kind: "stealth", size: "small" });
  });

  it("the legacy arrow geometry still writes a bare tailEnd and reads back as geometry arrow with no explicit endArrow (AC-15)", () => {
    const source = baseShape({ geometry: "arrow", heightCm: 0, outline: { color: "#2e5395", widthPt: 1, dash: "solid" } });
    const xml = shapeDrawingXml(source, 1, writeCtx);
    expect(xml).toContain('<a:tailEnd type="triangle"/>');
    const parsed = roundTrip(source)!;
    expect(parsed.shape.geometry).toBe("arrow");
    expect(parsed.shape.outline?.endArrow).toBeUndefined();
  });

  it("an arrow geometry with an explicit endArrow keeps the explicit sized tail", () => {
    const source = baseShape({
      geometry: "arrow",
      heightCm: 0,
      outline: { color: "#2e5395", widthPt: 1, dash: "solid", endArrow: { kind: "diamond", size: "medium" } },
    });
    const xml = shapeDrawingXml(source, 1, writeCtx);
    expect(xml).toContain('<a:tailEnd type="diamond" w="med" len="med"/>');
    const parsed = roundTrip(source)!;
    expect(parsed.shape.outline?.endArrow).toEqual({ kind: "diamond", size: "medium" });
  });
});

describe("SPEC-073 shadow round-trip (AC-8, AC-10, AC-32)", () => {
  it("writes a:effectLst/a:outerShdw with blurRad/dist/dir and the colour's alpha and reads it back", () => {
    const source = baseShape({ shadow: { color: "#000000", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 90 } });
    const xml = shapeDrawingXml(source, 1, writeCtx);
    expect(xml).toContain('<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="5400000" rotWithShape="0">');
    expect(xml).toContain('<a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>');
    expect(roundTrip(source)!.shape.shadow).toEqual(source.shadow);
  });

  it("no shadow writes no a:effectLst at all (AC-40's model half)", () => {
    expect(shapeDrawingXml(baseShape(), 1, writeCtx)).not.toContain("effectLst");
  });
});

describe("SPEC-073 preservation of unmodelled styling (AC-35, AC-36)", () => {
  const wsp = (spPrChildren: string) =>
    `<w:drawing xmlns:wp="${WP_NS}"><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    '<wp:extent cx="1800000" cy="1080000"/><wp:docPr id="1" name="Shape 1"/>' +
    `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp xmlns:wps="${WPS_NS}"><wps:cNvSpPr/><wps:spPr>` +
    '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="1080000"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    spPrChildren +
    "</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>";

  it("a scheme-colour fill is preserved verbatim, never flattened, and re-emitted on a same-format save", () => {
    const schemeFill = '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>';
    const parsed = parseShapeDrawing(drawingElement(wsp(schemeFill)), readCtx)!;
    expect(parsed.shape.fill).toBeUndefined();
    expect(parsed.shape.unsupportedStyleXml).toContain("schemeClr");
    const rewritten = shapeDrawingXml(parsed.shape, 1, writeCtx);
    expect(rewritten).toContain('<a:schemeClr val="accent1"/>');
    // The preserved fill takes the modelled fill's slot — no second fill.
    expect(rewritten).not.toContain("<a:noFill/></wps:spPr>");
    const again = parseShapeDrawing(drawingElement(rewritten), readCtx)!;
    expect(canonicalizeXml(`<x xmlns:a="${A_NS}">${again.shape.unsupportedStyleXml}</x>`)).toBe(
      canonicalizeXml(`<x xmlns:a="${A_NS}">${parsed.shape.unsupportedStyleXml}</x>`),
    );
  });

  it("a glow beside a parsed outerShdw preserves only the glow (exactly one source per property, AC-36)", () => {
    const effects =
      "<a:effectLst>" +
      '<a:glow rad="63500"><a:srgbClr val="00FF00"/></a:glow>' +
      '<a:outerShdw blurRad="50800" dist="38100" dir="5400000"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw>' +
      "</a:effectLst>";
    const parsed = parseShapeDrawing(drawingElement(wsp(effects)), readCtx)!;
    expect(parsed.shape.shadow).toEqual({ color: "#000000", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 90 });
    expect(parsed.shape.unsupportedStyleXml).toContain("a:glow");
    expect(parsed.shape.unsupportedStyleXml).not.toContain("outerShdw");
    // The writer merges the model's shadow and the preserved glow into ONE
    // effectLst — and in CT_EffectList's sequence order: glow (position 3)
    // must precede outerShdw (position 5), or the output is out-of-schema
    // OOXML a strict consumer flags as unreadable (AC-36 "schema order").
    const rewritten = shapeDrawingXml(parsed.shape, 1, writeCtx);
    expect(rewritten.match(/<a:effectLst>/g)).toHaveLength(1);
    expect(rewritten).toContain("a:glow");
    expect(rewritten).toContain("a:outerShdw");
    expect(rewritten.indexOf("<a:glow")).toBeLessThan(rewritten.indexOf("<a:outerShdw"));
  });

  it("a preserved reflection re-emits AFTER the modelled outerShdw (CT_EffectList order, AC-36)", () => {
    const effects =
      "<a:effectLst>" +
      '<a:outerShdw blurRad="50800" dist="38100" dir="5400000"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw>' +
      '<a:reflection blurRad="6350" stA="52000" endA="300"/>' +
      "</a:effectLst>";
    const parsed = parseShapeDrawing(drawingElement(wsp(effects)), readCtx)!;
    expect(parsed.shape.shadow).toBeDefined();
    const rewritten = shapeDrawingXml(parsed.shape, 1, writeCtx);
    expect(rewritten.indexOf("<a:outerShdw")).toBeLessThan(rewritten.indexOf("<a:reflection"));
  });

  it("an edited modelled fill wins over a stale preserved scheme-colour fill (AC-36 one-source)", () => {
    // Open a doc whose fill the model cannot hold (schemeClr -> preserved),
    // then the user picks a solid red: the canvas shows red, so the save
    // must write red — never resurrect the preserved theme fill. The edit
    // path clears the preserved slot (see shapes.spec.ts); the writer's
    // model-first guard covers any stale leftover, tested here directly.
    const schemeFill = '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>';
    const parsed = parseShapeDrawing(drawingElement(wsp(schemeFill)), readCtx)!;
    const edited: typeof parsed.shape = { ...parsed.shape, fill: { kind: "solid", color: "#ff0000" } };
    const rewritten = shapeDrawingXml(edited, 1, writeCtx);
    expect(rewritten).toContain('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>');
    expect(rewritten).not.toContain("schemeClr");
  });

  it("an edited modelled outline wins over a stale preserved a:ln (AC-36 one-source)", () => {
    const gradientStroke = '<a:ln w="25400"><a:gradFill><a:gsLst/></a:gradFill></a:ln>';
    const parsed = parseShapeDrawing(drawingElement(wsp(gradientStroke)), readCtx)!;
    expect(parsed.shape.outline).toBeUndefined();
    expect(parsed.shape.unsupportedStyleXml).toContain("gradFill");
    const edited: typeof parsed.shape = { ...parsed.shape, outline: { color: "#2e5395", widthPt: 1, dash: "solid" } };
    const rewritten = shapeDrawingXml(edited, 1, writeCtx);
    expect(rewritten).toContain('<a:ln w="12700">');
    expect(rewritten).not.toContain("gradFill");
    expect(rewritten.match(/<a:ln[ >]/g)).toHaveLength(1);
  });

  it("scene3d/sp3d append after effectLst in schema order", () => {
    const extras = '<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>';
    const parsed = parseShapeDrawing(drawingElement(wsp(extras)), readCtx)!;
    expect(parsed.shape.unsupportedStyleXml).toContain("scene3d");
    const rewritten = shapeDrawingXml({ ...parsed.shape, shadow: { color: "#000000", alpha: 1, blurPt: 0, distancePt: 1, angleDeg: 0 } }, 1, writeCtx);
    expect(rewritten.indexOf("<a:scene3d")).toBeGreaterThan(rewritten.indexOf("</a:effectLst>"));
  });
});
