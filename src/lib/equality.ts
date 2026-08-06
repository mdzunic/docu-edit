// The equality relations SPEC-008's round-trip guarantees are stated in
// (see specs/008-round-trip-preservation.md "Equality"):
//
//  - `equalBytes`      — byte identity; used only for preserved opaque parts.
//  - `equalStructural` — deep equality of two document models after
//    canonicalization, reporting the path of the first difference so a
//    failing round-trip test says *where* it broke. Opaque XML is compared
//    with XML canonicalization (C14N-style), so attribute order, namespace
//    prefixes and quoting can never cause false failures.
//
// Framework-free on purpose: this must run under Vitest (jsdom) and in the
// browser with nothing but DOMParser.
import { parseXml, XmlParseError } from "./docx/xml";

/** Byte identity between two buffers — the relation preserved opaque parts
 * are compared with before and after a save. */
export function equalBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) {
    return false;
  }
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) {
      return false;
    }
  }
  return true;
}

// --- XML canonicalization --------------------------------------------------

const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";

function escapeXmlText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replace(/"/g, "&quot;");
}

function canonicalizeElement(el: Element, out: string[]): void {
  const name = `{${el.namespaceURI ?? ""}}${el.localName}`;
  out.push(`<${name}`);

  // Namespace declarations are context, not content — the expanded {ns}local
  // names above already carry what they declare — so they are dropped and
  // the remaining attributes sorted by expanded name.
  const attributes = Array.from(el.attributes)
    .filter((attr) => attr.namespaceURI !== XMLNS_NAMESPACE)
    .map((attr) => ({
      name: `{${attr.namespaceURI ?? ""}}${attr.localName}`,
      value: attr.value,
    }))
    .sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
  for (const attr of attributes) {
    out.push(` ${attr.name}="${escapeXmlAttribute(attr.value)}"`);
  }
  out.push(">");

  for (const child of Array.from(el.childNodes)) {
    if (child.nodeType === Node.ELEMENT_NODE) {
      canonicalizeElement(child as Element, out);
    } else if (child.nodeType === Node.TEXT_NODE || child.nodeType === Node.CDATA_SECTION_NODE) {
      // Text is significant inside document content — kept verbatim.
      out.push(escapeXmlText((child as CharacterData).data));
    }
    // Comments and processing instructions carry no document content.
  }

  out.push(`</${name}>`);
}

/** A deterministic canonical form of an XML fragment: expanded `{ns}local`
 * names (so the prefix a producer chose is irrelevant), attributes sorted by
 * expanded name, namespace declarations dropped, comments dropped, text
 * verbatim. Two fragments are equivalent iff their canonical forms are
 * string-equal. Throws `XmlParseError` for malformed input. */
export function canonicalizeXml(xml: string): string {
  const doc = parseXml(xml);
  const out: string[] = [];
  canonicalizeElement(doc.documentElement, out);
  return out.join("");
}

function xmlEquivalent(a: string, b: string): boolean {
  try {
    return canonicalizeXml(a) === canonicalizeXml(b);
  } catch (error) {
    if (error instanceof XmlParseError) {
      // Unparseable opaque payloads fall back to string identity.
      return a === b;
    }
    throw error;
  }
}

// --- structural equality ---------------------------------------------------

/** Model fields holding XML strings, compared canonically. */
const XML_STRING_KEYS = new Set([
  "xml",
  "docxSectPr",
  "odtFontFaceDecls",
  // Verbatim table properties preserved from a .docx source (SPEC-022).
  "docxTblPr",
  "docxTrPr",
  "docxTcPr",
]);

/** Model fields holding arrays of XML strings. */
const XML_ARRAY_KEYS = new Set(["odtAutomaticStyles"]);

/** Fields excluded from structural equality: `id`s are runtime-only and
 * regenerated freely; the rest are UI state, not content — none of them is
 * written to a `.docx`/`.odt` file, so a same-format round trip never
 * carries them either way. */
const IGNORED_KEYS = new Set([
  "id",
  "suppressConversionReport",
  "measurementUnit",
  "showRuler",
]);

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Absent, undefined, empty-object and empty-array values are all "nothing
 * to say" — a reader that writes `formatting: {}` and one that omits the
 * field describe the same document. */
function isEmpty(value: unknown): boolean {
  if (value === undefined) {
    return true;
  }
  if (Array.isArray(value)) {
    return value.length === 0;
  }
  if (isPlainObject(value)) {
    return Object.keys(value).every((key) => IGNORED_KEYS.has(key) || isEmpty(value[key]));
  }
  return false;
}

function difference(
  a: unknown,
  b: unknown,
  path: string,
  xmlValued: boolean,
): string | null {
  if (xmlValued && typeof a === "string" && typeof b === "string") {
    return xmlEquivalent(a, b) ? null : path;
  }
  if (Array.isArray(a) && Array.isArray(b)) {
    if (a.length !== b.length) {
      return `${path}.length`;
    }
    for (let i = 0; i < a.length; i++) {
      const diff = difference(a[i], b[i], `${path}[${i}]`, xmlValued);
      if (diff !== null) {
        return diff;
      }
    }
    return null;
  }
  if (isPlainObject(a) && isPlainObject(b)) {
    const keys = [...new Set([...Object.keys(a), ...Object.keys(b)])].sort();
    for (const key of keys) {
      if (IGNORED_KEYS.has(key)) {
        continue;
      }
      const left = a[key];
      const right = b[key];
      if (isEmpty(left) && isEmpty(right)) {
        continue;
      }
      const childPath = path.length > 0 ? `${path}.${key}` : key;
      if (isEmpty(left) !== isEmpty(right)) {
        return childPath;
      }
      const diff = difference(
        left,
        right,
        childPath,
        XML_STRING_KEYS.has(key) || XML_ARRAY_KEYS.has(key),
      );
      if (diff !== null) {
        return diff;
      }
    }
    return null;
  }
  if (typeof a === "number" && typeof b === "number" && Number.isNaN(a) && Number.isNaN(b)) {
    return null;
  }
  return a === b ? null : path;
}

/** The path of the first structural difference between two models (e.g.
 * "sections[0].blocks[3].runs[1].text"), or null when they are equal. */
export function firstStructuralDifference(a: unknown, b: unknown): string | null {
  return difference(a, b, "", false);
}

/** Structural equality between two document models: deep equality after
 * dropping ids and empty/undefined fields, with opaque XML compared via
 * `canonicalizeXml`. */
export function equalStructural(a: unknown, b: unknown): boolean {
  return firstStructuralDifference(a, b) === null;
}
