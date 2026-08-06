// Reads/writes the `draw:frame` markup SPEC-024 models: a single image
// (`draw:image`), inline (`text:anchor-type="as-char"`) or floating
// (anchor-type "paragraph", wrap mode from the frame's graphic style). A
// frame holding anything else — a shape, an OLE object, a text box — is not
// modelled; `parseDrawFrame` returns null for those, and the caller (the
// reader's inline-content walk) keeps its existing SPEC-008 opaque-inline
// fallback, exactly as it did before this spec.
import type { Image, ImageWrapMode, MediaPart } from "../documentModel";
import { contentTypeForPartName } from "../docx/drawing";
import { firstChildElement } from "../docx/xml";
import { decodeImageDimensions } from "../imageDimensions";
import { bytesToBase64 } from "../opaque";
import { cmToOdfLength, odfLengthToCm } from "./units";
import { DRAW_NAMESPACE, SVG_NAMESPACE, STYLE_NAMESPACE, XLINK_NAMESPACE, drawAttr, styleAttr, svgAttr, textAttr, xlinkAttr } from "./xml";

/** A frame's graphic-family style resolved to SPEC-024's wrap enum — the
 * inverse of graphicPropertiesXml's own encoding below. Keyed by
 * `style:name`, collected once per document (mirroring reader.ts's other
 * collectTable*Styles scans of both content.xml and styles.xml). */
export function collectGraphicWrapStyles(contentDoc: Document, stylesDoc: Document | undefined): Map<string, ImageWrapMode> {
  const styles = new Map<string, ImageWrapMode>();
  const scan = (doc: Document | undefined) => {
    if (!doc) {
      return;
    }
    for (const el of Array.from(doc.getElementsByTagNameNS(STYLE_NAMESPACE, "style"))) {
      if (el.getAttributeNS(STYLE_NAMESPACE, "family") !== "graphic") {
        continue;
      }
      const name = el.getAttributeNS(STYLE_NAMESPACE, "name");
      const props = firstChildElement(el, STYLE_NAMESPACE, "graphic-properties");
      if (!name || !props) {
        continue;
      }
      const wrap = styleAttr(props, "wrap");
      const runThrough = styleAttr(props, "run-through");
      if (wrap === "parallel" || wrap === "dynamic") {
        styles.set(name, "square");
      } else if (wrap === "run-through" && runThrough === "background") {
        styles.set(name, "behind");
      } else if (wrap === "run-through" || wrap === "none") {
        styles.set(name, "infront");
      }
    }
  };
  scan(contentDoc);
  scan(stylesDoc);
  return styles;
}

export interface ParsedFrame {
  image: Image;
  media: MediaPart;
}

/** Parses a `draw:frame` into a modelled Image + MediaPart, or null when it
 * is not a single, plain picture this spec models (a shape/OLE object/text
 * box, or a href the package doesn't actually carry bytes for) — the
 * caller's existing opaque-inline fallback then preserves it verbatim. */
export function parseDrawFrame(
  frameEl: Element,
  parts: Map<string, Uint8Array>,
  graphicWrapStyles: Map<string, ImageWrapMode>,
): ParsedFrame | null {
  const imageEl = firstChildElement(frameEl, DRAW_NAMESPACE, "image");
  if (!imageEl) {
    return null;
  }
  const href = xlinkAttr(imageEl, "href");
  if (!href) {
    return null;
  }
  const partPath = href.replace(/^\.\//, "");
  const bytes = parts.get(partPath);
  const contentType = contentTypeForPartName(partPath);
  if (!bytes || !contentType) {
    return null;
  }

  const widthCm = odfLengthToCm(svgAttr(frameEl, "width"));
  const heightCm = odfLengthToCm(svgAttr(frameEl, "height"));
  if (widthCm === undefined || heightCm === undefined || widthCm <= 0 || heightCm <= 0) {
    return null;
  }

  const anchorType = textAttr(frameEl, "anchor-type");
  let wrap: ImageWrapMode;
  if (anchorType === null || anchorType === "as-char") {
    wrap = "inline";
  } else {
    const styleName = drawAttr(frameEl, "style-name");
    wrap = (styleName && graphicWrapStyles.get(styleName)) || "infront";
  }

  const descEl = firstChildElement(frameEl, SVG_NAMESPACE, "desc");
  const altText = descEl?.textContent || undefined;
  const natural = decodeImageDimensions(bytes, contentType);

  const image: Image = {
    type: "image",
    mediaId: partPath,
    naturalWidthPx: natural?.widthPx ?? Math.round((widthCm / 2.54) * 96),
    naturalHeightPx: natural?.heightPx ?? Math.round((heightCm / 2.54) * 96),
    widthCm,
    heightCm,
    wrap,
  };
  if (altText) {
    image.altText = altText;
  }
  return { image, media: { mediaId: partPath, contentType, data: bytesToBase64(bytes) } };
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** A generated graphic-family style's properties for a floating wrap mode
 * (never called for "inline", which has no style of its own — see
 * drawFrameXml). The inverse of collectGraphicWrapStyles' own mapping. */
function graphicPropertiesXml(wrap: ImageWrapMode): string {
  if (wrap === "square") {
    return '<style:graphic-properties style:wrap="parallel" style:number-wrapped-paragraphs="no-limit" style:wrap-contour="false" style:vertical-pos="from-top" style:vertical-rel="paragraph" style:horizontal-pos="from-left" style:horizontal-rel="paragraph"/>';
  }
  if (wrap === "behind") {
    return '<style:graphic-properties style:wrap="run-through" style:run-through="background" style:vertical-pos="from-top" style:vertical-rel="paragraph" style:horizontal-pos="from-left" style:horizontal-rel="paragraph"/>';
  }
  return '<style:graphic-properties style:wrap="run-through" style:run-through="foreground" style:vertical-pos="from-top" style:vertical-rel="paragraph" style:horizontal-pos="from-left" style:horizontal-rel="paragraph"/>';
}

/** One generated graphic-family `style:style` for a floating image's wrap
 * mode, to be added to content.xml's automatic styles — "inline" needs none
 * (see drawFrameXml). */
export function graphicStyleXml(name: string, wrap: ImageWrapMode): string {
  return `<style:style style:name="${escapeXml(name)}" style:family="graphic">${graphicPropertiesXml(wrap)}</style:style>`;
}

const FRAME_XMLNS =
  ` xmlns:draw="${DRAW_NAMESPACE}" xmlns:xlink="${XLINK_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}"`;

/** Serialises an Image into a `draw:frame` referencing `href` (the media
 * part's package path). `graphicStyleName` is a floating image's own
 * generated graphic style (see graphicStyleXml); undefined for "inline",
 * which anchors as a plain character with no positioning style of its own. */
export function drawFrameXml(image: Image, href: string, graphicStyleName: string | undefined): string {
  const anchorType = image.wrap === "inline" ? "as-char" : "paragraph";
  const styleAttrXml = graphicStyleName ? ` draw:style-name="${escapeXml(graphicStyleName)}"` : "";
  const descXml = image.altText ? `<svg:desc>${escapeXml(image.altText)}</svg:desc>` : "";
  return (
    `<draw:frame${FRAME_XMLNS} text:anchor-type="${anchorType}"${styleAttrXml}` +
    ` svg:width="${cmToOdfLength(image.widthCm)}" svg:height="${cmToOdfLength(image.heightCm)}">` +
    `<draw:image xlink:href="${escapeXml(href)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
    descXml +
    "</draw:frame>"
  );
}
