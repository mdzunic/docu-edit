// SPEC-056 — reads/writes the ODF markup for an equation: a `draw:frame`
// holding a `draw:object` reference to an embedded formula subpackage
// ("Formula N/") whose own `content.xml` carries
// `office:document-content` -> `office:body` -> `office:math` -> a MathML
// `<math>` element. This mirrors chart.ts's embedded-object split (SPEC-052)
// exactly, reusing the same package machinery rather than building a second
// one; only the payload differs (MathML instead of chart XML). An inline
// equation is anchored `as-char` (ODF's spelling of "in the text baseline");
// the model's `display` flag rides on the MathML `<math display>` attribute so
// display and inline equations stay one mechanism.
import type { EquationSpec } from "../documentModel";
import { escapeXmlAttribute, parseXml, XmlParseError } from "../core/xml";
import { MATHML_NAMESPACE, parseMathML, serializeMathML } from "../math/mathml";
import { DRAW_NAMESPACE, OFFICE_NAMESPACE, SVG_NAMESPACE, XLINK_NAMESPACE, xlinkAttr } from "./xml";

/** The manifest media-type of an embedded formula object. */
export const FORMULA_OBJECT_CONTENT_TYPE = "application/vnd.oasis.opendocument.formula";

function firstChildNS(parent: Element, ns: string, localName: string): Element | undefined {
  for (const child of Array.from(parent.children)) {
    if (child.namespaceURI === ns && child.localName === localName) {
      return child;
    }
  }
  return undefined;
}

// --- reading ---------------------------------------------------------------

/** The MathML `<math>` element inside a formula sub-document's content.xml,
 * ignoring any StarMath `annotation` (AC-47) — the query targets the MathML
 * namespace directly, so a `math:annotation` never matches. */
function mathElement(contentXml: string): Element | null {
  let doc: Document;
  try {
    doc = parseXml(contentXml);
  } catch (error) {
    if (error instanceof XmlParseError) return null;
    throw error;
  }
  const maths = doc.getElementsByTagNameNS(MATHML_NAMESPACE, "math");
  return maths.length > 0 ? maths[0] : null;
}

/** Resolves a `draw:frame` whose `draw:object` references a formula
 * sub-document into an EquationSpec, or null when the frame is not a formula
 * (no object, unresolved href, or the referenced part is not MathML — a chart
 * or other embedded object). The reader takes the MathML and ignores the
 * StarMath annotation (AC-47). */
export function parseFormulaFrame(frameEl: Element, parts: Map<string, Uint8Array>): EquationSpec | null {
  const objectEl = firstChildNS(frameEl, DRAW_NAMESPACE, "object");
  if (!objectEl) return null;
  const href = xlinkAttr(objectEl, "href");
  if (!href) return null;
  const objectDir = href.replace(/^\.\//, "").replace(/\/$/, "");
  const bytes = parts.get(`${objectDir}/content.xml`);
  if (!bytes) return null;
  const contentXml = new TextDecoder("utf-8").decode(bytes);
  const math = mathElement(contentXml);
  if (!math) return null;
  const eq = parseMathML(math);
  // An explicit svg:desc overrides the default (linear-form) alt text.
  const descEl = firstChildNS(frameEl, SVG_NAMESPACE, "desc");
  const alt = descEl?.textContent;
  if (alt) eq.altText = alt;
  return eq;
}

/** Whether a `draw:frame` is a formula object (used by the reader to keep a
 * formula frame from being mistaken for a chart or image). */
export function isFormulaFrame(frameEl: Element, parts: Map<string, Uint8Array>): boolean {
  return parseFormulaFrame(frameEl, parts) !== null;
}

// --- writing ---------------------------------------------------------------

const FRAME_XMLNS = ` xmlns:draw="${DRAW_NAMESPACE}" xmlns:xlink="${XLINK_NAMESPACE}" xmlns:svg="${SVG_NAMESPACE}"`;

/** The standalone `content.xml` of a formula sub-document: MathML under
 * `office:math`, no StarMath annotation (AC-48). LibreOffice regenerates its
 * own annotation from the MathML on open, so writing one would mean
 * implementing a second notation to produce something the reader throws away. */
export function odfFormulaContentXml(equation: EquationSpec): string {
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    `<office:document-content xmlns:office="${OFFICE_NAMESPACE}" office:version="1.2">` +
    "<office:body><office:math>" +
    serializeMathML(equation) +
    "</office:math></office:body>" +
    "</office:document-content>"
  );
}

/** The `draw:frame` carrier for an equation, referencing `objectHref` (the
 * embedded object's package-relative directory, e.g. "./Formula 1"). Always
 * `as-char` with a baseline vertical relation — ODF's spelling of an inline
 * equation that sits in the text baseline (AC-31). The alt text goes in
 * `svg:desc` so assistive technology reaches it (AC-21). */
export function formulaFrameXml(equation: EquationSpec, objectHref: string, altText: string): string {
  const descXml = altText ? `<svg:desc>${escapeXmlAttribute(altText)}</svg:desc>` : "";
  return (
    `<draw:frame${FRAME_XMLNS} text:anchor-type="as-char"` +
    ` svg:width="2cm" svg:height="1cm" style:vertical-rel="baseline">` +
    `<draw:object xlink:href="${escapeXmlAttribute(objectHref)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
    descXml +
    "</draw:frame>"
  );
}

/** Every embedded-object directory the preserved manifest declares as a
 * formula — a regex scan over the already-decoded manifest text, mirroring
 * chartObjectDirsInManifest. */
export function formulaObjectDirsInManifest(manifestXml: string): Set<string> {
  const dirs = new Set<string>();
  const escaped = FORMULA_OBJECT_CONTENT_TYPE.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const forward = new RegExp(
    `<manifest:file-entry\\b[^>]*manifest:full-path="([^"]+)/"[^>]*manifest:media-type="${escaped}"[^>]*/>`,
    "g",
  );
  const reverse = new RegExp(
    `<manifest:file-entry\\b[^>]*manifest:media-type="${escaped}"[^>]*manifest:full-path="([^"]+)/"[^>]*/>`,
    "g",
  );
  let match: RegExpExecArray | null;
  while ((match = forward.exec(manifestXml))) dirs.add(match[1]);
  while ((match = reverse.exec(manifestXml))) dirs.add(match[1]);
  return dirs;
}
