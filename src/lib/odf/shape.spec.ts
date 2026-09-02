// SPEC-051: the ODF shape reader/writer (draw:custom-shape / draw:line).
// Mirrors odf/drawing.spec.ts's own fixture conventions exactly — a `wrap()`
// helper supplying the ambient namespaces content.xml would already declare,
// and a Document built via DOMParser for parseShapeElement's style lookup.
import { describe, it, expect } from "vitest";
import {
  createOdfDrawingDefs,
  parseShapeElement,
  shapeElementXml,
  shapeGraphicStyleXml,
  type OdfShapeTextContext,
  type OdfShapeTextWriteContext,
} from "./shape";
import { parseXml } from "../core/xml";
import { canonicalizeXml } from "../core/xml";
import { DRAW_NAMESPACE, STYLE_NAMESPACE, SVG_NAMESPACE, TEXT_NAMESPACE, XLINK_NAMESPACE } from "../core/odfXml";
import { type Paragraph, type Shape } from "../documentModel";
import { SHAPE_CATALOGUE } from "../core/shapes";

function wrap(xml: string): Element {
  const wrapped =
    `<root xmlns:text="${TEXT_NAMESPACE}" xmlns:style="${STYLE_NAMESPACE}" ` +
    `xmlns:draw="${DRAW_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}" xmlns:xlink="${XLINK_NAMESPACE}">${xml}</root>`;
  return parseXml(wrapped).documentElement.children[0];
}

function wrapDoc(xml: string): Document {
  const wrapped =
    `<root xmlns:text="${TEXT_NAMESPACE}" xmlns:style="${STYLE_NAMESPACE}" ` +
    `xmlns:draw="${DRAW_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}" xmlns:xlink="${XLINK_NAMESPACE}">${xml}</root>`;
  return parseXml(wrapped);
}

/** Minimal paragraph parse/write stubs, exercising shape.ts's own text-box
 * plumbing rather than the full paragraph parser (out of scope here). A
 * paragraph round-trips as a single run's plain text. */
function escapeXmlText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

const readCtx: OdfShapeTextContext = {
  parseParagraph: (pEl: Element): Paragraph => {
    const text = pEl.textContent ?? "";
    return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
  },
};

const writeCtx: OdfShapeTextWriteContext = {
  paragraphXml: (paragraph: Paragraph): string => {
    const text = paragraph.runs.map((r) => r.text).join("");
    return `<text:p>${escapeXmlText(text)}</text:p>`;
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

/** Round-trips a Shape through shapeElementXml/shapeGraphicStyleXml and back
 * through parseShapeElement — the generated style AND the shared drawing
 * definitions (SPEC-073: gradients/hatches/fill images/markers/dashes,
 * normally routed into styles.xml) live in the same document as the shape
 * element, exactly as the reader resolves them across the package's parts. */
function roundTrip(
  shape: Shape,
  styleName = "gr1",
  opts: { hrefForMedia?: (mediaId: string) => string | undefined; resolveFillImage?: (href: string) => string | null } = {},
) {
  const defs = createOdfDrawingDefs();
  const styleXml = shapeGraphicStyleXml(styleName, shape, { defs, hrefForMedia: opts.hrefForMedia });
  const shapeXml = shapeElementXml(shape, styleName, writeCtx);
  const contentDoc = wrapDoc(defs.defsXml() + styleXml + shapeXml);
  const shapeEl = contentDoc.documentElement.lastElementChild!;
  return parseShapeElement(shapeEl, contentDoc, undefined, {
    ...readCtx,
    resolveFillImage: opts.resolveFillImage,
  });
}

// --- adjustments (SPEC-069 AC-6b, AC-7b) ----------------------------------------

describe("shape adjustments round-trip through draw:modifiers", () => {
  // One geometry per adjustment-bearing group (AC-6c).
  const cases: { geometry: Shape["geometry"]; adjustments: number[] }[] = [
    { geometry: "roundRect", adjustments: [0.25] }, // Basic shapes
    { geometry: "rightArrow", adjustments: [0.5, 0.3] }, // Block arrows
    { geometry: "star6", adjustments: [0.4] }, // Stars and banners
    { geometry: "wedgeRectCallout", adjustments: [0.2, 0.9] }, // Callouts
  ];

  for (const { geometry, adjustments } of cases) {
    it(`writes draw:modifiers and reads it back for ${geometry}`, () => {
      const shapeXml = shapeElementXml(baseShape({ geometry, adjustments }), "gr1", writeCtx);
      expect(shapeXml).toContain("draw:modifiers=");
      const parsed = roundTrip(baseShape({ geometry, adjustments }));
      expect(parsed).not.toBeNull();
      expect(parsed!.shape.adjustments).toBeDefined();
      parsed!.shape.adjustments!.forEach((value, i) => expect(value).toBeCloseTo(adjustments[i], 4));
    });
  }

  it("omits draw:modifiers for a shape with no adjustments", () => {
    const shapeXml = shapeElementXml(baseShape({ geometry: "roundRect" }), "gr1", writeCtx);
    expect(shapeXml).not.toContain("draw:modifiers=");
    expect(roundTrip(baseShape({ geometry: "roundRect" }))!.shape.adjustments).toBeUndefined();
  });

  it("preserves an out-of-range adjustment value verbatim on a same-format save (AC-7b)", () => {
    const parsed = roundTrip(baseShape({ geometry: "roundRect", adjustments: [2] }));
    expect(parsed!.shape.adjustments![0]).toBeCloseTo(2, 4);
  });
});

// --- parseShapeElement — null cases ---------------------------------------------

describe("parseShapeElement — null cases", () => {
  it("returns null for an element that is neither draw:custom-shape nor draw:line", () => {
    const el = wrap(`<draw:frame svg:width="5cm" svg:height="3cm"/>`);
    expect(parseShapeElement(el, wrapDoc(""), undefined, readCtx)).toBeNull();
  });

  it("returns null for a draw:custom-shape with no width/height", () => {
    const el = wrap(`<draw:custom-shape text:anchor-type="as-char"><draw:enhanced-geometry draw:type="rectangle"/></draw:custom-shape>`);
    expect(parseShapeElement(el, wrapDoc(""), undefined, readCtx)).toBeNull();
  });
});

// --- draw:custom-shape geometry mapping ------------------------------------------

describe("parseShapeElement — draw:custom-shape geometry mapping", () => {
  it("maps every catalogue area geometry's draw:type back to itself, round-tripped through the writer", () => {
    for (const entry of SHAPE_CATALOGUE) {
      if (entry.geometry === "line" || entry.geometry === "arrow") {
        continue;
      }
      const result = roundTrip(baseShape({ geometry: entry.geometry }));
      expect(result, entry.geometry).not.toBeNull();
      expect(result!.shape.geometry).toBe(entry.geometry);
      expect(result!.shape.unsupportedGeometryXml).toBeUndefined();
    }
  });

  it("an unrecognised draw:type falls back to 'rect' and captures the enhanced-geometry element verbatim (AC-60)", () => {
    const contentDoc = wrapDoc(
      `<draw:custom-shape text:anchor-type="as-char" svg:width="5cm" svg:height="3cm">` +
        `<draw:enhanced-geometry draw:type="gear"/>` +
        `</draw:custom-shape>`,
    );
    const el = contentDoc.documentElement.firstElementChild!;
    const result = parseShapeElement(el, contentDoc, undefined, readCtx);
    expect(result).not.toBeNull();
    expect(result!.shape.geometry).toBe("rect");
    expect(result!.shape.unsupportedGeometryXml).toContain('draw:type="gear"');
  });

  it("a custom-shape with no enhanced-geometry child at all is unsupported with an empty capture", () => {
    const contentDoc = wrapDoc(`<draw:custom-shape text:anchor-type="as-char" svg:width="5cm" svg:height="3cm"/>`);
    const el = contentDoc.documentElement.firstElementChild!;
    const result = parseShapeElement(el, contentDoc, undefined, readCtx);
    expect(result!.shape.geometry).toBe("rect");
    expect(result!.shape.unsupportedGeometryXml).toBe("");
  });
});

// --- draw:line geometry (line/arrow) ---------------------------------------------

describe("parseShapeElement — draw:line (Line/Arrow)", () => {
  it("computes width/height from |x2-x1| / |y2-y1|", () => {
    const el = wrap(`<draw:line text:anchor-type="paragraph" svg:x1="1cm" svg:y1="2cm" svg:x2="6cm" svg:y2="5cm"/>`);
    const result = parseShapeElement(el, wrapDoc(""), undefined, readCtx)!;
    expect(result.shape.widthCm).toBeCloseTo(5, 5);
    expect(result.shape.heightCm).toBeCloseTo(3, 5);
  });

  it("is 'line' with no marker-end, 'arrow' with one", () => {
    const line = wrap(`<draw:line text:anchor-type="paragraph" svg:x1="0cm" svg:y1="0cm" svg:x2="5cm" svg:y2="0cm"/>`);
    expect(parseShapeElement(line, wrapDoc(""), undefined, readCtx)!.shape.geometry).toBe("line");

    const arrow = wrap(
      `<draw:line text:anchor-type="paragraph" svg:x1="0cm" svg:y1="0cm" svg:x2="5cm" svg:y2="0cm" draw:marker-end="Arrow"/>`,
    );
    expect(parseShapeElement(arrow, wrapDoc(""), undefined, readCtx)!.shape.geometry).toBe("arrow");
  });
});

// --- resolveGraphicStyleProps: fill/outline/wrap from the referenced style -------

// resolveGraphicStyleProps (odf/shape.ts) reads fill/fill-color/stroke via
// `drawAttr` (DRAW_NAMESPACE, i.e. a `draw:fill` attribute) and
// stroke-color/stroke-width via `svgAttr` (SVG_NAMESPACE) — matching real
// ODF and this same module's own writer, graphicPropertiesXml. Only
// wrap/run-through are genuinely in the style: namespace.
describe("parseShapeElement — fill/outline/wrap resolved through the graphic style", () => {
  function customShapeWithStyle(styleProps: string, anchorType: string | null = "paragraph"): { el: Element; doc: Document } {
    const anchorAttr = anchorType === null ? "" : ` text:anchor-type="${anchorType}"`;
    const xml =
      `<style:style style:name="gr1" style:family="graphic"><style:graphic-properties ${styleProps}/></style:style>` +
      `<draw:custom-shape${anchorAttr} draw:style-name="gr1" svg:width="5cm" svg:height="3cm">` +
      `<draw:enhanced-geometry draw:type="rectangle"/>` +
      `</draw:custom-shape>`;
    const doc = wrapDoc(xml);
    return { el: doc.documentElement.lastElementChild!, doc };
  }

  it("reads a solid fill colour from draw:fill/draw:fill-color", () => {
    const { el, doc } = customShapeWithStyle('draw:fill="solid" draw:fill-color="#ff0000"');
    const result = parseShapeElement(el, doc, undefined, readCtx)!;
    expect(result.shape.fill).toEqual({ kind: "solid", color: "#ff0000" });
  });

  it("fill='none' means no fill", () => {
    const { el, doc } = customShapeWithStyle('draw:fill="none"');
    const result = parseShapeElement(el, doc, undefined, readCtx)!;
    expect(result.shape.fill).toBeUndefined();
  });

  it("style:fill/style:fill-color (the wrong namespace) is NOT read", () => {
    const { el, doc } = customShapeWithStyle('style:fill="solid" style:fill-color="#ff0000"');
    const result = parseShapeElement(el, doc, undefined, readCtx)!;
    expect(result.shape.fill).toBeUndefined();
  });

  it("reads a stroke into an outline (colour/width/solid) from draw:stroke/svg:stroke-color/svg:stroke-width", () => {
    const { el, doc } = customShapeWithStyle('draw:stroke="solid" svg:stroke-color="#000000" svg:stroke-width="0.035cm"');
    const result = parseShapeElement(el, doc, undefined, readCtx)!;
    expect(result.shape.outline?.color).toBe("#000000");
    expect(result.shape.outline?.dash).toBe("solid");
    expect(result.shape.outline?.widthPt).toBeGreaterThan(0);
  });

  it("stroke='dash' maps to dash", () => {
    const { el, doc } = customShapeWithStyle('draw:stroke="dash" svg:stroke-color="#000000" svg:stroke-width="0.035cm"');
    const result = parseShapeElement(el, doc, undefined, readCtx)!;
    expect(result.shape.outline?.dash).toBe("dash");
  });

  it("stroke='none' means no outline", () => {
    const { el, doc } = customShapeWithStyle('draw:stroke="none"');
    const result = parseShapeElement(el, doc, undefined, readCtx)!;
    expect(result.shape.outline).toBeUndefined();
  });

  it("maps style:wrap='parallel'/'dynamic' to 'square'", () => {
    const parallel = customShapeWithStyle('style:wrap="parallel"');
    expect(parseShapeElement(parallel.el, parallel.doc, undefined, readCtx)!.shape.wrap).toBe("square");
    const dynamic = customShapeWithStyle('style:wrap="dynamic"');
    expect(parseShapeElement(dynamic.el, dynamic.doc, undefined, readCtx)!.shape.wrap).toBe("square");
  });

  it("maps style:wrap='run-through' + run-through='background' to 'behind'", () => {
    const { el, doc } = customShapeWithStyle('style:wrap="run-through" style:run-through="background"');
    expect(parseShapeElement(el, doc, undefined, readCtx)!.shape.wrap).toBe("behind");
  });

  it("maps style:wrap='run-through' + run-through='foreground' to 'infront'", () => {
    const { el, doc } = customShapeWithStyle('style:wrap="run-through" style:run-through="foreground"');
    expect(parseShapeElement(el, doc, undefined, readCtx)!.shape.wrap).toBe("infront");
  });

  it("maps style:wrap='none' to 'topAndBottom'", () => {
    const { el, doc } = customShapeWithStyle('style:wrap="none"');
    expect(parseShapeElement(el, doc, undefined, readCtx)!.shape.wrap).toBe("topAndBottom");
  });

  it("text:anchor-type='as-char' forces wrap 'inline' regardless of the style's own wrap", () => {
    const { el, doc } = customShapeWithStyle('style:wrap="parallel"', "as-char");
    expect(parseShapeElement(el, doc, undefined, readCtx)!.shape.wrap).toBe("inline");
  });

  it("a missing anchor-type attribute also defaults to 'inline'", () => {
    const { el, doc } = customShapeWithStyle('style:wrap="parallel"', null);
    expect(parseShapeElement(el, doc, undefined, readCtx)!.shape.wrap).toBe("inline");
  });

  it("falls back to the stylesDoc when the style isn't in contentDoc", () => {
    const stylesDoc = wrapDoc(`<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#00ff00"/></style:style>`);
    const contentDoc = wrapDoc(
      `<draw:custom-shape text:anchor-type="paragraph" draw:style-name="gr1" svg:width="5cm" svg:height="3cm">` +
        `<draw:enhanced-geometry draw:type="rectangle"/></draw:custom-shape>`,
    );
    const el = contentDoc.documentElement.firstElementChild!;
    const result = parseShapeElement(el, contentDoc, stylesDoc, readCtx)!;
    expect(result.shape.fill).toEqual({ kind: "solid", color: "#00ff00" });
  });
});

// --- alt text / groupOrdinal / rotation / offsets ---------------------------------

describe("parseShapeElement — alt text, groupOrdinal, rotation, offsets", () => {
  it("reads svg:desc as alt text", () => {
    const shape = roundTrip(baseShape({ altText: "A red circle" }))!.shape;
    expect(shape.altText).toBe("A red circle");
  });

  it("has no altText when svg:desc is absent", () => {
    const shape = roundTrip(baseShape())!.shape;
    expect(shape.altText).toBeUndefined();
  });

  it("round-trips groupOrdinal and rotationDeg through the docu-edit extension attributes", () => {
    const shape = roundTrip(baseShape({ wrap: "square", groupOrdinal: 3, rotationDeg: 45 }))!.shape;
    expect(shape.groupOrdinal).toBe(3);
    expect(shape.rotationDeg).toBe(45);
  });

  it("omits both extension attributes (and the xmlns:de declaration) when neither is set", () => {
    const xml = shapeElementXml(baseShape(), "gr1", writeCtx);
    expect(xml).not.toContain("de:group");
    expect(xml).not.toContain("de:rotation");
    expect(xml).not.toContain("xmlns:de");
  });

  it("reads svg:x/svg:y into offsetXCm/offsetYCm for a non-inline shape", () => {
    const shape = roundTrip(baseShape({ wrap: "square", offsetXCm: 2, offsetYCm: 3 }))!.shape;
    expect(shape.offsetXCm).toBeCloseTo(2, 5);
    expect(shape.offsetYCm).toBeCloseTo(3, 5);
  });

  it("an inline shape has no offsets/relativeTo fields", () => {
    const shape = roundTrip(baseShape({ wrap: "inline" }))!.shape;
    expect(shape.offsetXCm).toBeUndefined();
    expect(shape.offsetYCm).toBeUndefined();
    expect(shape.relativeToH).toBeUndefined();
    expect(shape.relativeToV).toBeUndefined();
  });
});

// --- text (draw:text-box) --------------------------------------------------------

describe("parseShapeElement — text frame (draw:text-box)", () => {
  it("no paragraphs means no text field", () => {
    const shape = roundTrip(baseShape())!.shape;
    expect(shape.text).toBeUndefined();
  });

  it("round-trips paragraph text and vAlign through draw:text-box", () => {
    const source = baseShape({ text: { blocks: [{ type: "paragraph", runs: [{ text: "hello shape", formatting: {} }] }], vAlign: "middle" } });
    const shape = roundTrip(source)!.shape;
    expect(shape.text?.blocks).toHaveLength(1);
    expect((shape.text!.blocks[0] as Paragraph).runs[0].text).toBe("hello shape");
    expect(shape.text?.vAlign).toBe("middle");
  });

  it("round-trips top/bottom vAlign", () => {
    const top = roundTrip(baseShape({ text: { blocks: [{ type: "paragraph", runs: [] }], vAlign: "top" } }))!.shape;
    expect(top.text?.vAlign).toBe("top");
    const bottom = roundTrip(baseShape({ text: { blocks: [{ type: "paragraph", runs: [] }], vAlign: "bottom" } }))!.shape;
    expect(bottom.text?.vAlign).toBe("bottom");
  });

  it("round-trips a bulleted-list-worthy multi-paragraph text frame", () => {
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

  it("a Line never gets a text-box (parseShapeText is only consulted for area geometries)", () => {
    const el = wrap(`<draw:line text:anchor-type="paragraph" svg:x1="0cm" svg:y1="0cm" svg:x2="5cm" svg:y2="0cm"><text:p>ignored</text:p></draw:line>`);
    const result = parseShapeElement(el, wrapDoc(""), undefined, readCtx)!;
    expect(result.shape.text).toBeUndefined();
  });
});

// --- shapeElementXml round trip ----------------------------------------------------

describe("shapeElementXml round-trips through parseShapeElement", () => {
  it("round-trips an inline shape's size and geometry", () => {
    const source = baseShape({ geometry: "ellipse", widthCm: 4, heightCm: 6, wrap: "inline" });
    const result = roundTrip(source)!;
    expect(result.shape.wrap).toBe("inline");
    expect(result.shape.geometry).toBe("ellipse");
    expect(result.shape.widthCm).toBeCloseTo(4, 5);
    expect(result.shape.heightCm).toBeCloseTo(6, 5);
  });

  it("round-trips a floating (square) shape's offsets", () => {
    const source = baseShape({ wrap: "square", offsetXCm: 3, offsetYCm: 1.5 });
    const result = roundTrip(source)!;
    expect(result.shape.wrap).toBe("square");
    expect(result.shape.offsetXCm).toBeCloseTo(3, 5);
    expect(result.shape.offsetYCm).toBeCloseTo(1.5, 5);
  });

  it("round-trips a behind-wrapped shape", () => {
    const result = roundTrip(baseShape({ wrap: "behind" }))!;
    expect(result.shape.wrap).toBe("behind");
  });

  it("round-trips an infront-wrapped shape", () => {
    const result = roundTrip(baseShape({ wrap: "infront" }))!;
    expect(result.shape.wrap).toBe("infront");
  });

  it("round-trips a topAndBottom-wrapped shape", () => {
    const result = roundTrip(baseShape({ wrap: "topAndBottom" }))!;
    expect(result.shape.wrap).toBe("topAndBottom");
  });

  it("round-trips fill end-to-end through the real writer+reader", () => {
    const result = roundTrip(baseShape({ fill: { kind: "solid", color: "#336699" } }))!;
    expect(result.shape.fill).toEqual({ kind: "solid", color: "#336699" });
  });

  it("round-trips an outline end-to-end (SPEC-073: the dash KIND survives via draw:stroke-dash, no longer degrading to plain dash)", () => {
    const result = roundTrip(baseShape({ outline: { color: "#123456", widthPt: 2.25, dash: "dot" } }))!;
    expect(result.shape.outline?.color).toBe("#123456");
    expect(result.shape.outline?.dash).toBe("dot");
    expect(result.shape.outline?.widthPt).toBeGreaterThan(0);
  });

  it("round-trips a Line's width/height (derived from x1/y1/x2/y2) and its anchored offset (derived from x1/y1, which a draw:line carries instead of svg:x/svg:y)", () => {
    const source = baseShape({ geometry: "line", wrap: "square", offsetXCm: 1, offsetYCm: 2, widthCm: 5, heightCm: 3 });
    const xml = shapeElementXml(source, "gr1", writeCtx);
    // The writer places the anchored position in x1/y1 (there is no
    // separate svg:x/svg:y on a draw:line at all).
    expect(xml).toContain('svg:x1="1cm"');
    expect(xml).toContain('svg:y1="2cm"');
    const result = roundTrip(source)!;
    expect(result.shape.geometry).toBe("line");
    expect(result.shape.widthCm).toBeCloseTo(5, 5);
    expect(result.shape.heightCm).toBeCloseTo(3, 5);
    expect(result.shape.offsetXCm).toBeCloseTo(1, 5);
    expect(result.shape.offsetYCm).toBeCloseTo(2, 5);
  });

  it("round-trips an Arrow with its marker-end", () => {
    const source = baseShape({ geometry: "arrow", wrap: "square", widthCm: 5, heightCm: 0 });
    const xml = shapeElementXml(source, "gr1", writeCtx);
    expect(xml).toContain("draw:marker-end");
    const result = roundTrip(source)!;
    expect(result.shape.geometry).toBe("arrow");
  });

  it("round-trips an unsupported geometry's XML verbatim, marking the geometry field as the rect stand-in (AC-60, AC-61)", () => {
    const unsupportedXml = `<draw:enhanced-geometry xmlns:draw="${DRAW_NAMESPACE}" draw:type="gear"/>`;
    const source = baseShape({ geometry: "rect", unsupportedGeometryXml: unsupportedXml });
    const xml = shapeElementXml(source, "gr1", writeCtx);
    expect(xml).not.toContain('draw:type="rectangle"');
    expect(xml).toContain('draw:type="gear"');

    const contentDoc = wrapDoc(shapeGraphicStyleXml("gr1", source) + xml);
    const shapeEl = contentDoc.documentElement.lastElementChild!;
    const result = parseShapeElement(shapeEl, contentDoc, undefined, readCtx)!;
    expect(result.shape.geometry).toBe("rect");
    expect(canonicalizeXml(result.shape.unsupportedGeometryXml!)).toBe(canonicalizeXml(unsupportedXml));
  });
});

// --- shapeGraphicStyleXml ----------------------------------------------------------

describe("shapeGraphicStyleXml", () => {
  it("emits draw:fill='none' and draw:stroke='none' for a bare text box (SPEC-051's text box definition)", () => {
    const textBox = baseShape({ text: { blocks: [{ type: "paragraph", runs: [] }], vAlign: "top" } });
    const xml = shapeGraphicStyleXml("gr1", textBox);
    expect(xml).toContain('draw:fill="none"');
    expect(xml).toContain('draw:stroke="none"');
  });

  it("carries the style family and parent (Graphics)", () => {
    const xml = shapeGraphicStyleXml("gr1", baseShape());
    expect(xml).toContain('style:family="graphic"');
    expect(xml).toContain('style:parent-style-name="Graphics"');
  });
});

// --- custom geometry / paths (SPEC-071) -------------------------------------

describe("path shapes round-trip through draw:enhanced-path (SPEC-071 AC-27)", () => {
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

  it("emits a non-primitive enhanced-geometry with viewBox and enhanced-path", () => {
    const xml = shapeElementXml(baseShape({ geometry: "rect", path, fill: { kind: "solid", color: "#00ff00" } }), "gr1", writeCtx);
    expect(xml).toContain('draw:type="non-primitive"');
    expect(xml).toContain('svg:viewBox="0 0 21600 21600"');
    expect(xml).toContain("draw:enhanced-path=");
    expect(xml).toContain("N"); // trailing endpath terminator
  });

  it("reads back into an equal ShapePath, no unsupportedGeometryXml", () => {
    const parsed = roundTrip(baseShape({ geometry: "rect", path, fill: { kind: "solid", color: "#00ff00" } }));
    expect(parsed?.shape.path).toEqual(path);
    expect(parsed?.shape.unsupportedGeometryXml).toBeUndefined();
    expect(parsed?.shape.geometry).toBe("rect");
  });

  it("converts a quadratic on read and preserves the source", () => {
    const shapeXml =
      `<draw:custom-shape xmlns:draw="${DRAW_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}" xmlns:text="${TEXT_NAMESPACE}" ` +
      `text:anchor-type="as-char" draw:style-name="gr1" svg:width="4cm" svg:height="4cm">` +
      `<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="non-primitive" ` +
      `draw:enhanced-path="M 0 0 Q 10800 21600 21600 0 Z N"/></draw:custom-shape>`;
    const contentDoc = wrapDoc(shapeXml);
    const shapeEl = contentDoc.documentElement.lastElementChild!;
    const parsed = parseShapeElement(shapeEl, contentDoc, undefined, readCtx);
    const steps = parsed?.shape.path?.subpaths[0].steps ?? [];
    expect(steps[1].op).toBe("cubic");
    expect(parsed?.shape.pathSource?.format).toBe("odt");
    expect(parsed?.shape.unsupportedGeometryXml).toBeUndefined();
  });
});

// --- SPEC-073: fills (AC-33) -------------------------------------------------

describe("SPEC-073 fills — written attributes and round trip", () => {
  it("solid fill with alpha writes draw:opacity as a whole percent and parses back", () => {
    const shape = baseShape({ fill: { kind: "solid", color: "#336699", alpha: 0.4 } });
    const styleXml = shapeGraphicStyleXml("gr1", shape, { defs: createOdfDrawingDefs() });
    expect(styleXml).toContain('draw:fill="solid"');
    expect(styleXml).toContain('draw:fill-color="#336699"');
    expect(styleXml).toContain('draw:opacity="40%"');
    const back = roundTrip(shape)!.shape;
    expect(back.fill).toEqual({ kind: "solid", color: "#336699", alpha: 0.4 });
  });

  it("both no-fill model forms write draw:fill='none' and read back as NO fill (undefined, never kind:'none')", () => {
    for (const fill of [undefined, { kind: "none" } as const]) {
      const shape = baseShape({ fill });
      expect(shapeGraphicStyleXml("gr1", shape)).toContain('draw:fill="none"');
      expect(roundTrip(shape)!.shape.fill).toBeUndefined();
    }
  });

  it("a linear gradient writes draw:fill='gradient' referencing a deGradN definition with the two end colours", () => {
    const shape = baseShape({
      fill: {
        kind: "gradient",
        gradient: {
          type: "linear",
          angleDeg: 45,
          stops: [
            { position: 0, color: "#ff0000" },
            { position: 1, color: "#0000ff" },
          ],
        },
      },
    });
    const defs = createOdfDrawingDefs();
    const styleXml = shapeGraphicStyleXml("gr1", shape, { defs });
    expect(styleXml).toContain('draw:fill="gradient"');
    expect(styleXml).toContain('draw:fill-gradient-name="deGrad1"');
    const defsXml = defs.defsXml();
    expect(defsXml).toContain('draw:name="deGrad1"');
    expect(defsXml).toContain('draw:style="linear"');
    expect(defsXml).toContain('draw:start-color="#ff0000"');
    expect(defsXml).toContain('draw:end-color="#0000ff"');
    // angle10 = ((90 - 45) mod 360) * 10 — the documented bijection.
    expect(defsXml).toContain('draw:angle="450"');

    const back = roundTrip(shape)!.shape;
    expect(back.fill).toEqual(shape.fill);
  });

  it("a radial gradient writes draw:style='radial' and round-trips its type and colours", () => {
    const shape = baseShape({
      fill: {
        kind: "gradient",
        gradient: {
          type: "radial",
          stops: [
            { position: 0, color: "#ffffff" },
            { position: 1, color: "#222222" },
          ],
        },
      },
    });
    const defs = createOdfDrawingDefs();
    shapeGraphicStyleXml("gr1", shape, { defs });
    expect(defs.defsXml()).toContain('draw:style="radial"');
    const back = roundTrip(shape)!.shape;
    expect(back.fill).toEqual(shape.fill);
  });

  it("a gradient's intermediate stops do not write (only the two ends land in the definition)", () => {
    const shape = baseShape({
      fill: {
        kind: "gradient",
        gradient: {
          type: "linear",
          angleDeg: 90,
          stops: [
            { position: 0, color: "#ff0000" },
            { position: 0.5, color: "#00ff00" },
            { position: 1, color: "#0000ff" },
          ],
        },
      },
    });
    const back = roundTrip(shape)!.shape;
    expect(back.fill?.kind).toBe("gradient");
    if (back.fill?.kind === "gradient") {
      expect(back.fill.gradient.stops).toEqual([
        { position: 0, color: "#ff0000" },
        { position: 1, color: "#0000ff" },
      ]);
    }
  });

  it("identical gradients on two shapes dedupe to one definition", () => {
    const gradient = {
      type: "linear" as const,
      angleDeg: 90,
      stops: [
        { position: 0, color: "#ff0000" },
        { position: 1, color: "#0000ff" },
      ],
    };
    const defs = createOdfDrawingDefs();
    shapeGraphicStyleXml("gr1", baseShape({ fill: { kind: "gradient", gradient } }), { defs });
    shapeGraphicStyleXml("gr2", baseShape({ fill: { kind: "gradient", gradient: { ...gradient } } }), { defs });
    expect(defs.defsXml().match(/<draw:gradient/g)).toHaveLength(1);
  });

  it("pattern preset 'ltUpDiag' writes draw:fill='hatch' + a single hatch rotated 450 tenths, and round-trips", () => {
    const shape = baseShape({
      fill: { kind: "pattern", preset: "ltUpDiag", foreground: "#112233", background: "#ffffee" },
    });
    const defs = createOdfDrawingDefs();
    const styleXml = shapeGraphicStyleXml("gr1", shape, { defs });
    expect(styleXml).toContain('draw:fill="hatch"');
    expect(styleXml).toContain('draw:fill-hatch-name="deHatch1"');
    expect(styleXml).toContain('draw:fill-hatch-solid="true"');
    expect(styleXml).toContain('draw:fill-color="#ffffee"'); // background
    const defsXml = defs.defsXml();
    expect(defsXml).toContain('draw:style="single"');
    expect(defsXml).toContain('draw:rotation="450"');
    expect(defsXml).toContain('draw:color="#112233"'); // foreground

    const back = roundTrip(shape)!.shape;
    expect(back.fill).toEqual(shape.fill);
  });

  it("a preset outside the closed hatch table flattens to a solid fill of the foreground colour (drop reported centrally)", () => {
    const shape = baseShape({
      fill: { kind: "pattern", preset: "pct50", foreground: "#112233", background: "#ffffee" },
    });
    const styleXml = shapeGraphicStyleXml("gr1", shape, { defs: createOdfDrawingDefs() });
    expect(styleXml).toContain('draw:fill="solid"');
    expect(styleXml).toContain('draw:fill-color="#112233"');
    expect(styleXml).not.toContain("hatch");
  });

  it("picture fill writes draw:fill='bitmap' + a deFillImgN definition and style:repeat, and round-trips mediaId/mode/alpha", () => {
    const shape = baseShape({ fill: { kind: "picture", mediaId: "Pictures/pic.png", mode: "tile", alpha: 0.75 } });
    const hrefForMedia = (mediaId: string) => (mediaId === "Pictures/pic.png" ? "Pictures/image1.png" : undefined);
    const resolveFillImage = (href: string) => (href === "Pictures/image1.png" ? "Pictures/pic.png" : null);
    const defs = createOdfDrawingDefs();
    const styleXml = shapeGraphicStyleXml("gr1", shape, { defs, hrefForMedia });
    expect(styleXml).toContain('draw:fill="bitmap"');
    expect(styleXml).toContain('draw:fill-image-name="deFillImg1"');
    expect(styleXml).toContain('style:repeat="repeat"');
    expect(styleXml).toContain('draw:opacity="75%"');
    const defsXml = defs.defsXml();
    expect(defsXml).toContain('draw:name="deFillImg1"');
    expect(defsXml).toContain('xlink:href="Pictures/image1.png"');
    expect(defsXml).toContain('xlink:show="embed"');

    const back = roundTrip(shape, "gr1", { hrefForMedia, resolveFillImage })!.shape;
    expect(back.fill).toEqual(shape.fill);
  });

  it("a stretched picture fill writes style:repeat='stretch' and round-trips mode 'stretch'", () => {
    const shape = baseShape({ fill: { kind: "picture", mediaId: "m1", mode: "stretch" } });
    const hrefForMedia = () => "Pictures/image1.png";
    const resolveFillImage = () => "m1";
    const styleXml = shapeGraphicStyleXml("gr1", shape, { defs: createOdfDrawingDefs(), hrefForMedia });
    expect(styleXml).toContain('style:repeat="stretch"');
    const back = roundTrip(shape, "gr1", { hrefForMedia, resolveFillImage })!.shape;
    expect(back.fill).toEqual({ kind: "picture", mediaId: "m1", mode: "stretch" });
  });
});

// --- SPEC-073: outline depth (AC-33) ----------------------------------------

describe("SPEC-073 outline — opacity, join, cap, dash kinds", () => {
  it("writes svg:stroke-opacity / svg:stroke-linejoin / svg:stroke-linecap and parses them back", () => {
    const shape = baseShape({
      outline: { color: "#123456", widthPt: 2, dash: "solid", alpha: 0.5, join: "bevel", cap: "round" },
    });
    const styleXml = shapeGraphicStyleXml("gr1", shape, { defs: createOdfDrawingDefs() });
    expect(styleXml).toContain('svg:stroke-opacity="50%"');
    expect(styleXml).toContain('svg:stroke-linejoin="bevel"');
    expect(styleXml).toContain('svg:stroke-linecap="round"');
    const back = roundTrip(shape)!.shape;
    expect(back.outline?.alpha).toBe(0.5);
    expect(back.outline?.join).toBe("bevel");
    expect(back.outline?.cap).toBe("round");
  });

  it("omits join/cap/opacity attrs when unset (flat cap and default join write nothing)", () => {
    const styleXml = shapeGraphicStyleXml("gr1", baseShape({ outline: { color: "#000000", widthPt: 1, dash: "solid" } }));
    expect(styleXml).not.toContain("stroke-linejoin");
    expect(styleXml).not.toContain("stroke-linecap");
    expect(styleXml).not.toContain("stroke-opacity");
  });

  it("every non-solid dash kind writes its own de-named draw:stroke-dash (plus the shared definition) and round-trips", () => {
    const expected: Record<string, string> = {
      dash: "deDash",
      dot: "deDot",
      dashDot: "deDashDot",
      longDash: "deLongDash",
      longDashDot: "deLongDashDot",
    };
    for (const [dash, name] of Object.entries(expected)) {
      const shape = baseShape({ outline: { color: "#000000", widthPt: 1, dash: dash as never } });
      const defs = createOdfDrawingDefs();
      const styleXml = shapeGraphicStyleXml("gr1", shape, { defs });
      expect(styleXml, dash).toContain('draw:stroke="dash"');
      expect(styleXml, dash).toContain(`draw:stroke-dash="${name}"`);
      expect(defs.defsXml(), dash).toContain(`<draw:stroke-dash xmlns:draw="${DRAW_NAMESPACE}" draw:name="${name}"`);
      expect(roundTrip(shape)!.shape.outline?.dash, dash).toBe(dash);
    }
  });

  it("an unknown draw:stroke-dash name still reads as a plain dash (legacy behaviour)", () => {
    const contentDoc = wrapDoc(
      `<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:stroke="dash" draw:stroke-dash="Fine Dashed" svg:stroke-color="#000000" svg:stroke-width="0.035cm"/></style:style>` +
        `<draw:custom-shape text:anchor-type="as-char" draw:style-name="gr1" svg:width="5cm" svg:height="3cm"><draw:enhanced-geometry draw:type="rectangle"/></draw:custom-shape>`,
    );
    const el = contentDoc.documentElement.lastElementChild!;
    expect(parseShapeElement(el, contentDoc, undefined, readCtx)!.shape.outline?.dash).toBe("dash");
  });

  it("compound is not expressible: a double-compound outline writes a single stroke (no extra attrs)", () => {
    const styleXml = shapeGraphicStyleXml(
      "gr1",
      baseShape({ outline: { color: "#000000", widthPt: 1, dash: "solid", compound: "double" } }),
    );
    expect(styleXml).toContain('draw:stroke="solid"');
    expect(styleXml).not.toContain("compound");
  });
});

// --- SPEC-073: shadow (AC-33) -----------------------------------------------

describe("SPEC-073 shadow — draw:shadow-* plus the de:shadow-blur-pt extension", () => {
  it("writes visible/colour/opacity/offsets and the blur extension", () => {
    const shape = baseShape({ shadow: { color: "#000000", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 90 } });
    const styleXml = shapeGraphicStyleXml("gr1", shape, { defs: createOdfDrawingDefs() });
    expect(styleXml).toContain('draw:shadow="visible"');
    expect(styleXml).toContain('draw:shadow-color="#000000"');
    expect(styleXml).toContain('draw:shadow-opacity="40%"');
    // angle 90 = straight down: x offset 0, y offset 3pt in cm.
    expect(styleXml).toContain('draw:shadow-offset-x="0cm"');
    expect(styleXml).toMatch(/draw:shadow-offset-y="0\.105[0-9]*cm"/);
    expect(styleXml).toContain('de:shadow-blur-pt="4"');
    expect(styleXml).toContain(`xmlns:de="urn:docu-edit:shape"`);
  });

  it("round-trips the full shadow exactly", () => {
    const shadow = { color: "#332211", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 90 };
    const back = roundTrip(baseShape({ shadow }))!.shape;
    expect(back.shadow).toEqual(shadow);
  });

  it("round-trips a diagonal shadow (distance and angle recovered from the cartesian offsets)", () => {
    const shadow = { color: "#000000", alpha: 0.25, blurPt: 0, distancePt: 6, angleDeg: 315 };
    const back = roundTrip(baseShape({ shadow }))!.shape;
    expect(back.shadow?.color).toBe("#000000");
    expect(back.shadow?.alpha).toBe(0.25);
    expect(back.shadow?.blurPt).toBe(0);
    expect(back.shadow?.distancePt).toBeCloseTo(6, 1);
    expect(back.shadow?.angleDeg).toBe(315);
  });

  it("no shadow writes no draw:shadow attrs at all (AC-40's skipped pass)", () => {
    const styleXml = shapeGraphicStyleXml("gr1", baseShape());
    expect(styleXml).not.toContain("draw:shadow");
    expect(styleXml).not.toContain("shadow-blur-pt");
  });
});

// --- SPEC-073: arrowheads on plain lines (AC-33) ----------------------------

describe("SPEC-073 line arrowheads — de-named markers with sized widths", () => {
  const arrowLine = baseShape({
    geometry: "line",
    widthCm: 5,
    heightCm: 0,
    outline: {
      color: "#000000",
      widthPt: 1,
      dash: "solid",
      startArrow: { kind: "open", size: "large" },
      endArrow: { kind: "stealth", size: "small" },
    },
  });

  it("writes draw:marker-start/-end naming the closed deArrow* set with per-size widths, and the marker definitions", () => {
    const defs = createOdfDrawingDefs();
    const styleXml = shapeGraphicStyleXml("gr1", arrowLine, { defs });
    expect(styleXml).toContain('draw:marker-start="deArrowOpen"');
    expect(styleXml).toContain('draw:marker-start-width="0.45cm"');
    expect(styleXml).toContain('draw:marker-end="deArrowStealth"');
    expect(styleXml).toContain('draw:marker-end-width="0.2cm"');
    const defsXml = defs.defsXml();
    expect(defsXml).toContain('draw:name="deArrowOpen"');
    expect(defsXml).toContain('draw:name="deArrowStealth"');
    expect(defsXml).toContain('svg:viewBox="0 0 10 10"');
  });

  it("round-trips kinds and sizes on a plain line, geometry staying 'line'", () => {
    const back = roundTrip(arrowLine)!.shape;
    expect(back.geometry).toBe("line");
    expect(back.outline?.startArrow).toEqual({ kind: "open", size: "large" });
    expect(back.outline?.endArrow).toEqual({ kind: "stealth", size: "small" });
  });

  it("the legacy arrow geometry still writes exactly draw:marker-end='Arrow' 0.3cm on the element and parses back with NO outline.endArrow", () => {
    const legacy = baseShape({ geometry: "arrow", widthCm: 5, heightCm: 0 });
    const xml = shapeElementXml(legacy, "gr1", writeCtx);
    expect(xml).toContain('draw:marker-end="Arrow" draw:marker-end-width="0.3cm"');
    expect(xml).not.toContain("de:arrow-geometry");
    const back = roundTrip(legacy)!.shape;
    expect(back.geometry).toBe("arrow");
    expect(back.outline?.endArrow).toBeUndefined();
  });

  it("an arrow geometry WITH an explicit endArrow moves to de-named markers plus de:arrow-geometry, round-tripping both", () => {
    const styled = baseShape({
      geometry: "arrow",
      widthCm: 5,
      heightCm: 0,
      outline: { color: "#000000", widthPt: 1, dash: "solid", endArrow: { kind: "triangle", size: "large" } },
    });
    const elementXml = shapeElementXml(styled, "gr1", writeCtx);
    expect(elementXml).toContain('de:arrow-geometry="1"');
    expect(elementXml).not.toContain('draw:marker-end="Arrow"');
    const styleXml = shapeGraphicStyleXml("gr1", styled, { defs: createOdfDrawingDefs() });
    expect(styleXml).toContain('draw:marker-end="deArrowTriangle"');
    const back = roundTrip(styled)!.shape;
    expect(back.geometry).toBe("arrow");
    expect(back.outline?.endArrow).toEqual({ kind: "triangle", size: "large" });
  });
});

// --- SPEC-073: preservation (AC-35/AC-36) -----------------------------------

describe("SPEC-073 preservation — unmodelled graphic-properties survive verbatim", () => {
  function parseWithProps(styleProps: string, children = "") {
    const contentDoc = wrapDoc(
      `<style:style style:name="gr1" style:family="graphic"><style:graphic-properties ${styleProps}>${children}</style:graphic-properties></style:style>` +
        `<draw:custom-shape text:anchor-type="as-char" draw:style-name="gr1" svg:width="5cm" svg:height="3cm"><draw:enhanced-geometry draw:type="rectangle"/></draw:custom-shape>`,
    );
    const el = contentDoc.documentElement.lastElementChild!;
    return parseShapeElement(el, contentDoc, undefined, readCtx)!.shape;
  }

  it("an unknown attribute and child element land in unsupportedStyleXml; the modelled fields stay clean", () => {
    const shape = parseWithProps(
      'draw:fill="solid" draw:fill-color="#ff0000" draw:textarea-horizontal-align="center" fo:min-height="1cm"'.replace(
        "fo:min-height",
        "style:flow-with-text",
      ),
      `<style:background-image/>`,
    );
    expect(shape.fill).toEqual({ kind: "solid", color: "#ff0000" });
    expect(shape.unsupportedStyleXml).toBeDefined();
    expect(shape.unsupportedStyleXml).toContain('draw:textarea-horizontal-align="center"');
    expect(shape.unsupportedStyleXml).toContain("background-image");
    // Interpreted attributes never land in the preserved set (AC-36).
    expect(shape.unsupportedStyleXml).not.toContain("draw:fill=");
    expect(shape.unsupportedStyleXml).not.toContain("fill-color");
  });

  it("a same-format write re-emits the preserved attribute and child verbatim, merged into the generated properties", () => {
    const shape = parseWithProps(
      'draw:fill="solid" draw:fill-color="#ff0000" draw:textarea-horizontal-align="center"',
      `<style:background-image/>`,
    );
    const styleXml = shapeGraphicStyleXml("gr1", shape, { defs: createOdfDrawingDefs() });
    expect(styleXml).toContain('draw:fill="solid"');
    expect(styleXml).toContain('draw:fill-color="#ff0000"');
    expect(styleXml).toContain('draw:textarea-horizontal-align="center"');
    expect(styleXml).toContain("background-image");
  });

  it("write→parse→write is stable (the preserved set does not grow or mutate)", () => {
    const first = parseWithProps(
      'draw:fill="solid" draw:fill-color="#ff0000" draw:textarea-horizontal-align="center"',
      `<style:background-image/>`,
    );
    const second = roundTrip(first)!.shape;
    expect(second.fill).toEqual(first.fill);
    expect(canonicalizeXml(second.unsupportedStyleXml!)).toBe(canonicalizeXml(first.unsupportedStyleXml!));
    const third = roundTrip(second)!.shape;
    expect(canonicalizeXml(third.unsupportedStyleXml!)).toBe(canonicalizeXml(second.unsupportedStyleXml!));
  });

  it("cross-format writes do NOT merge preserved styling (includePreserved: false)", () => {
    const shape = parseWithProps('draw:fill="solid" draw:fill-color="#ff0000" draw:textarea-horizontal-align="center"');
    const styleXml = shapeGraphicStyleXml("gr1", shape, { includePreserved: false });
    expect(styleXml).not.toContain("textarea-horizontal-align");
  });

  it("an unresolvable gradient reference keeps the fill attrs preserved, the model fill undefined, and the writer defers to them", () => {
    const shape = parseWithProps('draw:fill="gradient" draw:fill-gradient-name="Missing 3"');
    expect(shape.fill).toBeUndefined();
    expect(shape.unsupportedStyleXml).toContain('draw:fill="gradient"');
    expect(shape.unsupportedStyleXml).toContain('draw:fill-gradient-name="Missing 3"');
    // The writer must not fight the preserved draw:fill with its own (AC-35).
    const styleXml = shapeGraphicStyleXml("gr1", shape, { defs: createOdfDrawingDefs() });
    expect(styleXml).toContain('draw:fill="gradient"');
    expect(styleXml).not.toContain('draw:fill="none"');
  });

  it("an unmodelled hatch (triple style) preserves the reference AND inlines the definition for the defs pre-pass", () => {
    const contentDoc = wrapDoc(
      `<draw:hatch draw:name="Odd Hatch" draw:style="triple" draw:color="#123456" draw:rotation="150"/>` +
        `<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="hatch" draw:fill-hatch-name="Odd Hatch" draw:fill-color="#ffffff"/></style:style>` +
        `<draw:custom-shape text:anchor-type="as-char" draw:style-name="gr1" svg:width="5cm" svg:height="3cm"><draw:enhanced-geometry draw:type="rectangle"/></draw:custom-shape>`,
    );
    const el = contentDoc.documentElement.lastElementChild!;
    const shape = parseShapeElement(el, contentDoc, undefined, readCtx)!.shape;
    expect(shape.fill).toBeUndefined();
    expect(shape.unsupportedStyleXml).toContain('draw:fill-hatch-name="Odd Hatch"');
    expect(shape.unsupportedStyleXml).toContain('draw:style="triple"');

    // The writer re-emits the reference and routes the definition into the
    // defs registry (bound for styles.xml office:styles).
    const defs = createOdfDrawingDefs();
    const styleXml = shapeGraphicStyleXml("gr1", shape, { defs });
    expect(styleXml).toContain('draw:fill-hatch-name="Odd Hatch"');
    expect(styleXml).not.toContain("triple"); // definition routed out, not left inline
    expect(defs.defsXml()).toContain('draw:name="Odd Hatch"');
    expect(defs.defsXml()).toContain('draw:style="triple"');
  });

  it("an edited modelled fill wins over stale preserved fill attrs — no resurrection, no orphaned companions (AC-36 one-source)", () => {
    // Same unmodelled triple-hatch document, but the user then picks a solid
    // red fill: the canvas shows red, so the save must write red — never the
    // preserved hatch reference. The edit path clears the preserved slot
    // (shapes.spec.ts); the writer's model-first guard covers any stale
    // leftover, tested here directly with the leftover still in place.
    const contentDoc = wrapDoc(
      `<draw:hatch draw:name="Odd Hatch" draw:style="triple" draw:color="#123456" draw:rotation="150"/>` +
        `<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="hatch" draw:fill-hatch-name="Odd Hatch" draw:fill-color="#ffffff" draw:auto-grow-height="true"/></style:style>` +
        `<draw:custom-shape text:anchor-type="as-char" draw:style-name="gr1" svg:width="5cm" svg:height="3cm"><draw:enhanced-geometry draw:type="rectangle"/></draw:custom-shape>`,
    );
    const el = contentDoc.documentElement.lastElementChild!;
    const shape = parseShapeElement(el, contentDoc, undefined, readCtx)!.shape;
    expect(shape.fill).toBeUndefined();
    const edited: Shape = { ...shape, fill: { kind: "solid", color: "#ff0000" } };
    const defs = createOdfDrawingDefs();
    const styleXml = shapeGraphicStyleXml("gr1", edited, { defs });
    expect(styleXml).toContain('draw:fill="solid"');
    expect(styleXml).toContain('draw:fill-color="#ff0000"');
    expect(styleXml).not.toContain("fill-hatch-name");
    // Non-fill preserved styling still merges in untouched.
    expect(styleXml).toContain('draw:auto-grow-height="true"');
    // The hatch definition rode along only for the replaced fill.
    expect(defs.defsXml()).not.toContain("Odd Hatch");
  });

  it("a shape with no foreign styling has no unsupportedStyleXml after a round trip of our own output", () => {
    const styled = baseShape({
      wrap: "square",
      fill: { kind: "solid", color: "#4472c4", alpha: 0.8 },
      outline: { color: "#2e5395", widthPt: 1.5, dash: "dashDot", alpha: 0.9, join: "round", cap: "square" },
      shadow: { color: "#000000", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 45 },
    });
    const back = roundTrip(styled)!.shape;
    expect(back.unsupportedStyleXml).toBeUndefined();
    expect(back.fill).toEqual(styled.fill);
    expect(back.outline).toEqual(styled.outline);
    expect(back.shadow?.color).toBe("#000000");
    expect(back.shadow?.distancePt).toBeCloseTo(3, 1);
    expect(back.shadow?.angleDeg).toBe(45);
  });
});

describe("path shapes round-trip through draw:enhanced-path (SPEC-071 AC-27) — formula preservation", () => {
  it("preserves a formula-driven enhanced path verbatim (no path modelled)", () => {
    const shapeXml =
      `<draw:custom-shape xmlns:draw="${DRAW_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}" xmlns:text="${TEXT_NAMESPACE}" ` +
      `text:anchor-type="as-char" draw:style-name="gr1" svg:width="4cm" svg:height="4cm">` +
      `<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="non-primitive" ` +
      `draw:enhanced-path="M 0 0 L ?f0 21600 N"><draw:equation draw:name="f0" draw:formula="0"/>` +
      `</draw:enhanced-geometry></draw:custom-shape>`;
    const contentDoc = wrapDoc(shapeXml);
    const shapeEl = contentDoc.documentElement.lastElementChild!;
    const parsed = parseShapeElement(shapeEl, contentDoc, undefined, readCtx);
    expect(parsed?.shape.path).toBeUndefined();
    expect(parsed?.shape.unsupportedGeometryXml).toContain("enhanced-geometry");
  });
});
