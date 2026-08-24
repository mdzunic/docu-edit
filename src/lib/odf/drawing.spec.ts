import { describe, it, expect } from "vitest";
import {
  collectGraphicWrapStyles,
  drawFrameXml,
  graphicStyleXml,
  parseDrawFrame,
} from "./drawing";
import { parseXml } from "../core/xml";
import { DRAW_NAMESPACE, STYLE_NAMESPACE, SVG_NAMESPACE, TEXT_NAMESPACE, XLINK_NAMESPACE } from "../core/odfXml";
import type { Image, ImageWrapMode } from "../documentModel";

/** A minimal 24-byte PNG (see imageDimensions.spec.ts's own builder), just
 * enough for decodeImageDimensions inside parseDrawFrame to read real
 * dimensions from. */
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

// drawFrameXml/graphicStyleXml emit fragments meant to be embedded inside
// content.xml, which already declares every ODF namespace at its root — so,
// like the .docx wp:drawing fixtures, a standalone parse needs a wrapper
// that supplies the namespaces the fragment itself doesn't redeclare
// (text:, style:) alongside the ones it does (draw:, xlink:, svg:).
function wrap(xml: string): Element {
  const wrapped =
    `<root xmlns:text="${TEXT_NAMESPACE}" xmlns:style="${STYLE_NAMESPACE}" ` +
    `xmlns:draw="${DRAW_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}" xmlns:xlink="${XLINK_NAMESPACE}">${xml}</root>`;
  return parseXml(wrapped).documentElement.children[0];
}

function frameXml(opts: {
  anchorType?: string | null;
  styleName?: string;
  width?: string;
  height?: string;
  href?: string;
  desc?: string;
  hasImage?: boolean;
}): string {
  const anchorAttr = opts.anchorType === null ? "" : ` text:anchor-type="${opts.anchorType ?? "as-char"}"`;
  const styleAttr = opts.styleName ? ` draw:style-name="${opts.styleName}"` : "";
  const width = opts.width ?? "5cm";
  const height = opts.height ?? "3cm";
  const imageEl =
    opts.hasImage === false
      ? "" // a shape with no draw:image child
      : `<draw:image xlink:href="${opts.href ?? "Pictures/image1.png"}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`;
  const descXml = opts.desc ? `<svg:desc>${opts.desc}</svg:desc>` : "";
  return (
    `<draw:frame${anchorAttr}${styleAttr} svg:width="${width}" svg:height="${height}">` +
    imageEl +
    descXml +
    `</draw:frame>`
  );
}

function graphicStyle(name: string, props: string): string {
  return `<style:style style:name="${name}" style:family="graphic"><style:graphic-properties ${props}/></style:style>`;
}

describe("collectGraphicWrapStyles", () => {
  function styleDoc(bodyXml: string): Document {
    return parseXml(`<root xmlns:style="${STYLE_NAMESPACE}">${bodyXml}</root>`);
  }

  it("maps wrap='parallel' to 'square'", () => {
    const doc = styleDoc(graphicStyle("fr1", 'style:wrap="parallel"'));
    expect(collectGraphicWrapStyles(doc, undefined).get("fr1")).toBe("square");
  });

  it("maps wrap='dynamic' to 'square'", () => {
    const doc = styleDoc(graphicStyle("fr1", 'style:wrap="dynamic"'));
    expect(collectGraphicWrapStyles(doc, undefined).get("fr1")).toBe("square");
  });

  it("maps wrap='run-through' + run-through='background' to 'behind'", () => {
    const doc = styleDoc(graphicStyle("fr1", 'style:wrap="run-through" style:run-through="background"'));
    expect(collectGraphicWrapStyles(doc, undefined).get("fr1")).toBe("behind");
  });

  it("maps wrap='run-through' + run-through='foreground' to 'infront'", () => {
    const doc = styleDoc(graphicStyle("fr1", 'style:wrap="run-through" style:run-through="foreground"'));
    expect(collectGraphicWrapStyles(doc, undefined).get("fr1")).toBe("infront");
  });

  it("maps wrap='none' to 'infront'", () => {
    const doc = styleDoc(graphicStyle("fr1", 'style:wrap="none"'));
    expect(collectGraphicWrapStyles(doc, undefined).get("fr1")).toBe("infront");
  });

  it("ignores styles of a different family", () => {
    const doc = styleDoc(
      `<style:style style:name="p1" style:family="paragraph"><style:paragraph-properties/></style:style>`,
    );
    expect(collectGraphicWrapStyles(doc, undefined).size).toBe(0);
  });

  it("scans both the content and styles documents", () => {
    const contentDoc = styleDoc(graphicStyle("fr1", 'style:wrap="parallel"'));
    const stylesDoc = styleDoc(graphicStyle("fr2", 'style:wrap="run-through" style:run-through="background"'));
    const styles = collectGraphicWrapStyles(contentDoc, stylesDoc);
    expect(styles.get("fr1")).toBe("square");
    expect(styles.get("fr2")).toBe("behind");
  });

  it("tolerates an absent styles document", () => {
    expect(() => collectGraphicWrapStyles(styleDoc(""), undefined)).not.toThrow();
  });
});

describe("parseDrawFrame — inline (AC-13)", () => {
  const parts = new Map([["Pictures/image1.png", PNG_BYTES]]);

  it("text:anchor-type='as-char' parses as wrap 'inline'", () => {
    const frame = wrap(frameXml({ anchorType: "as-char" }));
    const result = parseDrawFrame(frame, parts, new Map());
    expect(result).not.toBeNull();
    expect(result!.image.wrap).toBe("inline");
    expect(result!.image.mediaId).toBe("Pictures/image1.png");
    expect(result!.image.naturalWidthPx).toBe(400);
    expect(result!.image.naturalHeightPx).toBe(200);
  });

  it("a missing anchor-type attribute also defaults to 'inline'", () => {
    const frame = wrap(frameXml({ anchorType: null }));
    const result = parseDrawFrame(frame, parts, new Map());
    expect(result!.image.wrap).toBe("inline");
  });

  it("reads width/height from svg:width/svg:height", () => {
    const frame = wrap(frameXml({ width: "4cm", height: "2cm" }));
    const result = parseDrawFrame(frame, parts, new Map())!;
    expect(result.image.widthCm).toBeCloseTo(4, 6);
    expect(result.image.heightCm).toBeCloseTo(2, 6);
  });

  it("reads svg:desc as alt text", () => {
    const frame = wrap(frameXml({ desc: "A red circle" }));
    const result = parseDrawFrame(frame, parts, new Map())!;
    expect(result.image.altText).toBe("A red circle");
  });

  it("has no altText when svg:desc is absent", () => {
    const frame = wrap(frameXml({}));
    const result = parseDrawFrame(frame, parts, new Map())!;
    expect(result.image.altText).toBeUndefined();
  });
});

describe("parseDrawFrame — floating (AC-13, AC-14)", () => {
  const parts = new Map([["Pictures/image1.png", PNG_BYTES]]);

  it("resolves the wrap mode through the graphic style map", () => {
    const wrapStyles = new Map<string, ImageWrapMode>([["fr1", "square"]]);
    const frame = wrap(frameXml({ anchorType: "paragraph", styleName: "fr1" }));
    const result = parseDrawFrame(frame, parts, wrapStyles)!;
    expect(result.image.wrap).toBe("square");
  });

  it("falls back to 'infront' when the style name isn't in the wrap-style map", () => {
    const frame = wrap(frameXml({ anchorType: "paragraph", styleName: "unknown-style" }));
    const result = parseDrawFrame(frame, parts, new Map())!;
    expect(result.image.wrap).toBe("infront");
  });

  it("falls back to 'infront' when a floating frame has no style-name at all", () => {
    const frame = wrap(frameXml({ anchorType: "paragraph" }));
    const result = parseDrawFrame(frame, parts, new Map())!;
    expect(result.image.wrap).toBe("infront");
  });
});

describe("parseDrawFrame — null cases (SPEC-008 opaque fallback)", () => {
  const parts = new Map([["Pictures/image1.png", PNG_BYTES]]);

  it("returns null for a frame with no draw:image child (a shape)", () => {
    const frame = wrap(frameXml({ hasImage: false }));
    expect(parseDrawFrame(frame, parts, new Map())).toBeNull();
  });

  it("returns null when the href points at bytes the package doesn't have", () => {
    const frame = wrap(frameXml({ href: "Pictures/missing.png" }));
    expect(parseDrawFrame(frame, parts, new Map())).toBeNull();
  });

  it("returns null when width/height are missing", () => {
    const raw =
      `<draw:frame text:anchor-type="as-char">` +
      `<draw:image xlink:href="Pictures/image1.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
      `</draw:frame>`;
    expect(parseDrawFrame(wrap(raw), parts, new Map())).toBeNull();
  });

  it("returns null for a zero/negative width or height", () => {
    expect(parseDrawFrame(wrap(frameXml({ width: "0cm" })), parts, new Map())).toBeNull();
    expect(parseDrawFrame(wrap(frameXml({ height: "-1cm" })), parts, new Map())).toBeNull();
  });

  it("returns null when the extension has no known content type", () => {
    const frame = wrap(frameXml({ href: "Pictures/image1.bmp" }));
    const bmpParts = new Map([["Pictures/image1.bmp", PNG_BYTES]]);
    expect(parseDrawFrame(frame, bmpParts, new Map())).toBeNull();
  });
});

describe("drawFrameXml / graphicStyleXml round-trip through parseDrawFrame and collectGraphicWrapStyles", () => {
  const parts = new Map([["Pictures/image1.png", PNG_BYTES]]);

  it("round-trips an inline image", () => {
    const source: Image = {
      type: "image",
      mediaId: "Pictures/image1.png",
      naturalWidthPx: 400,
      naturalHeightPx: 200,
      widthCm: 5,
      heightCm: 2.5,
      wrap: "inline",
      altText: "A diagram",
    };
    const xml = drawFrameXml(source, "Pictures/image1.png", undefined);
    const parsed = parseDrawFrame(wrap(xml), parts, new Map());
    expect(parsed).not.toBeNull();
    expect(parsed!.image.wrap).toBe("inline");
    expect(parsed!.image.widthCm).toBeCloseTo(5, 6);
    expect(parsed!.image.heightCm).toBeCloseTo(2.5, 6);
    expect(parsed!.image.altText).toBe("A diagram");
    expect(parsed!.image.mediaId).toBe("Pictures/image1.png");
  });

  it("round-trips a floating image through a generated graphic style", () => {
    const source: Image = {
      type: "image",
      mediaId: "Pictures/image1.png",
      naturalWidthPx: 400,
      naturalHeightPx: 200,
      widthCm: 8,
      heightCm: 4,
      wrap: "square",
    };
    const styleXml = graphicStyleXml("fr1", "square");
    // The generated style feeds back through collectGraphicWrapStyles exactly
    // as it would from a real content.xml automatic-styles scan.
    const stylesContentDoc = parseXml(`<root xmlns:style="${STYLE_NAMESPACE}">${styleXml}</root>`);
    const wrapStyles = collectGraphicWrapStyles(stylesContentDoc, undefined);
    expect(wrapStyles.get("fr1")).toBe("square");

    const frameXmlOut = drawFrameXml(source, "Pictures/image1.png", "fr1");
    const parsed = parseDrawFrame(wrap(frameXmlOut), parts, wrapStyles);
    expect(parsed).not.toBeNull();
    expect(parsed!.image.wrap).toBe("square");
    expect(parsed!.image.widthCm).toBeCloseTo(8, 6);
    expect(parsed!.image.heightCm).toBeCloseTo(4, 6);
  });
});
