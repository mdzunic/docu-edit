// XML parsing, serialisation, OPC relationship resolution and
// canonicalization, built on the platform's DOMParser (available in every
// browser and in jsdom under Vitest) — no bundled XML library needed.
// Format-neutral: the WordprocessingML-specific helpers stay in
// docx/xml.ts, but everything here is shared by any OPC/ODF package
// (.docx/.odt today, .pptx/.odp for a Slides product). Elements are looked
// up by namespace URI rather than by prefix, since different producers
// (Word, LibreOffice, Google Docs) are free to bind a namespace to any
// prefix.

export const RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
export const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";

export class XmlParseError extends Error {}

/** Parses an XML string into a `Document`, throwing `XmlParseError` if the
 * markup is not well-formed (DOMParser reports this via an inline
 * `<parsererror>` node rather than throwing). */
export function parseXml(text: string): Document {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) {
    throw new XmlParseError("Malformed XML");
  }
  return doc;
}

/** A self-contained serialisation of an element (the platform serialiser
 * declares any in-scope namespaces the fragment uses), stored verbatim on
 * SPEC-008's opaque nodes and re-emitted as-is on save. */
export function serializeElement(el: Element): string {
  return new XMLSerializer().serializeToString(el);
}

export function childElements(parent: Element | Document, namespace: string, localName: string): Element[] {
  const result: Element[] = [];
  for (const child of Array.from(parent.children)) {
    if (child.namespaceURI === namespace && child.localName === localName) {
      result.push(child);
    }
  }
  return result;
}

export function firstChildElement(parent: Element | Document, namespace: string, localName: string): Element | undefined {
  return childElements(parent, namespace, localName)[0];
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
}

/** Parses a `_rels/*.rels` part into its relationships, keyed by `Id`. */
export function parseRelationships(xmlText: string): Map<string, Relationship> {
  const doc = parseXml(xmlText);
  const relationships = new Map<string, Relationship>();
  const root = doc.documentElement;
  if (!root) {
    return relationships;
  }
  for (const el of childElements(root, RELATIONSHIPS_NAMESPACE, "Relationship")) {
    const id = el.getAttribute("Id");
    const type = el.getAttribute("Type");
    const target = el.getAttribute("Target");
    if (id && type && target) {
      relationships.set(id, { id, type, target });
    }
  }
  return relationships;
}

/** Resolves a relationship `Target` against the part that owns the `.rels`
 * file referencing it — e.g. base "word/document.xml", target "styles.xml"
 * resolves to "word/styles.xml"; a leading "/" is package-root-relative. */
export function resolvePartPath(basePartPath: string, target: string): string {
  if (target.startsWith("/")) {
    return target.slice(1);
  }
  const baseDir = basePartPath.includes("/")
    ? basePartPath.slice(0, basePartPath.lastIndexOf("/"))
    : "";
  const segments = baseDir.length > 0 ? baseDir.split("/") : [];
  for (const part of target.split("/")) {
    if (part === "." || part === "") {
      continue;
    }
    if (part === "..") {
      segments.pop();
    } else {
      segments.push(part);
    }
  }
  return segments.join("/");
}

/** The `.rels` part path for a given package part, e.g.
 * "word/document.xml" -> "word/_rels/document.xml.rels". */
export function relsPathFor(partPath: string): string {
  const lastSlash = partPath.lastIndexOf("/");
  const dir = lastSlash >= 0 ? partPath.slice(0, lastSlash) : "";
  const fileName = lastSlash >= 0 ? partPath.slice(lastSlash + 1) : partPath;
  return `${dir.length > 0 ? `${dir}/` : ""}_rels/${fileName}.rels`;
}

// --- XML escaping ------------------------------------------------------------
// The three escape sets in use across the writers and the canonicalizer.
// Saved-byte fidelity pins each caller to its historical set (SPEC-008's
// round-trip rule; e2e asserts byte-identical saves), so these are three
// distinct functions, not one: swapping a caller to a different set would
// change its output wherever a quote or apostrophe occurs.

/** Escapes text content: `&`, `<`, `>` only. */
export function escapeXmlText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Escapes a double-quoted attribute value: `&`, `<`, `>`, `"`. */
export function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replace(/"/g, "&quot;");
}

/** Escapes all five predefined entities: `&`, `<`, `>`, `"`, `'`. The
 * docx/odf document writers (and docProps/meta) have always escaped both
 * text and attribute values with this full set — `&apos;` is redundant in
 * text and in double-quoted attributes, but harmless, and their saved bytes
 * pin it. */
export function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// --- XML canonicalization --------------------------------------------------
// Moved from lib/equality.ts: the C14N-style canonical form SPEC-008's
// opaque-XML comparisons are stated in, so attribute order, namespace
// prefixes and quoting can never cause false failures.

const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";

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
