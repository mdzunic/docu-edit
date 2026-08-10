// ODF namespace constants and attribute helpers. The generic XML machinery
// (parseXml, childElements, firstChildElement) is shared with the .docx
// reader — see src/lib/docx/xml.ts — since both formats are namespaced XML
// parsed with the platform's DOMParser. Attributes are read by namespace URI
// rather than prefix, since producers are free to rebind prefixes.

export const OFFICE_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
export const TEXT_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
export const TABLE_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
export const STYLE_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
export const FO_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
export const SVG_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
/** Drawing shapes/frames (SPEC-024) — `draw:frame`, `draw:image`. */
export const DRAW_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
/** XLink (SPEC-024) — `draw:image`'s `xlink:href`. */
export const XLINK_NAMESPACE = "http://www.w3.org/1999/xlink";
/** Dublin Core (SPEC-038) — `office:annotation`'s `dc:creator`/`dc:date`
 * child elements. */
export const DC_NAMESPACE = "http://purl.org/dc/elements/1.1/";
/** Metadata (SPEC-041) — `meta.xml`'s `office:meta` children. */
export const META_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";

/** Reads an `office:*` attribute regardless of the bound prefix. */
export function officeAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(OFFICE_NAMESPACE, localName);
}

/** Reads a `text:*` attribute regardless of the bound prefix. */
export function textAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(TEXT_NAMESPACE, localName);
}

/** Reads a `table:*` attribute regardless of the bound prefix. */
export function tableAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(TABLE_NAMESPACE, localName);
}

/** Reads a `style:*` attribute regardless of the bound prefix. */
export function styleAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(STYLE_NAMESPACE, localName);
}

/** Reads an `fo:*` (XSL-FO-compatible) attribute regardless of the bound
 * prefix. */
export function foAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(FO_NAMESPACE, localName);
}

/** Reads an `svg:*` (SVG-compatible) attribute regardless of the bound
 * prefix. */
export function svgAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(SVG_NAMESPACE, localName);
}

/** Reads a `draw:*` attribute regardless of the bound prefix. */
export function drawAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(DRAW_NAMESPACE, localName);
}

/** Reads an `xlink:*` attribute regardless of the bound prefix. */
export function xlinkAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(XLINK_NAMESPACE, localName);
}
