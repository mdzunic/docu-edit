// ODF namespace constants and attribute helpers, shared by every
// OpenDocument format this repo reads (`.odt` text in src/lib/odf/, `.odp`
// presentations in src/lib/presentation/odp/). The generic XML machinery
// (parseXml, childElements, firstChildElement) lives in ./xml — both
// package families are namespaced XML parsed with the platform's DOMParser.
// Attributes are read by namespace URI rather than prefix, since producers
// are free to rebind prefixes.
import { CM_PER_INCH, POINTS_PER_INCH } from "./units";

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
/** Data styles (SPEC-053) — `number:date-style`/`number:time-style` bodies a
 * date/time field's `style:data-style-name` points at. */
export const NUMBER_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
/** Presentations (`.odp`) — `office:presentation`, `presentation:class`. */
export const PRESENTATION_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
/** SMIL-compatible animation/transition attributes (SPEC-085) —
 * `smil:type`/`smil:subtype`/`smil:direction`/`smil:dur` on a page's
 * `style:drawing-page-properties`. */
export const SMIL_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";

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

/** Reads a `presentation:*` attribute regardless of the bound prefix. */
export function presentationAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(PRESENTATION_NAMESPACE, localName);
}

/** Reads a `smil:*` attribute regardless of the bound prefix. */
export function smilAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(SMIL_NAMESPACE, localName);
}

// --- ODF measured-value parsers (SPEC-079's move from odf/units.ts) ---------
// ODF writes measured values as strings with an explicit unit suffix
// ("2.54cm", "1in", "12pt"); the reading direction is product-neutral (both
// `.odt` text and `.odp` slides parse the same strings), so it lives in core
// and odf/units.ts re-exports it. The writer direction (searching for a
// byte-exact round-trip string) stays in odf/units.ts.

const CM_PER_ODF_UNIT: Record<string, number> = {
  cm: 1,
  mm: 0.1,
  in: CM_PER_INCH,
  pt: CM_PER_INCH / POINTS_PER_INCH,
  pc: (CM_PER_INCH / POINTS_PER_INCH) * 12,
};

const ODF_LENGTH_PATTERN = /^(-?\d+(?:\.\d+)?)(cm|mm|in|pt|pc)$/;

/** Parses an ODF length ("1.27cm", "-0.5in", "12pt") to centimetres, or
 * undefined for a missing/unitless/percentage value. */
export function odfLengthToCm(value: string | null | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const match = ODF_LENGTH_PATTERN.exec(value.trim());
  if (!match) {
    return undefined;
  }
  return Number(match[1]) * CM_PER_ODF_UNIT[match[2]];
}

/** Parses an ODF length to points (the model's unit for font sizes and
 * paragraph spacing), or undefined if it is not an absolute length. */
export function odfLengthToPt(value: string | null | undefined): number | undefined {
  const cm = odfLengthToCm(value);
  return cm === undefined ? undefined : (cm / CM_PER_INCH) * POINTS_PER_INCH;
}

const ODF_PERCENTAGE_PATTERN = /^(\d+(?:\.\d+)?)%$/;

/** Parses an ODF percentage ("130%") to its number (130), or undefined for
 * anything else. */
export function odfPercentage(value: string | null | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const match = ODF_PERCENTAGE_PATTERN.exec(value.trim());
  if (!match) {
    return undefined;
  }
  return Number(match[1]);
}

/** Normalises an ODF colour ("#C9211E") to a CSS colour ("#c9211e"), or
 * undefined for "transparent"/keyword/empty values (meaning "inherit"). */
export function odfColorToCss(value: string | null | undefined): string | undefined {
  if (!value || !value.startsWith("#")) {
    return undefined;
  }
  return value.toLowerCase();
}
