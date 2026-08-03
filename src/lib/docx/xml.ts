// XML parsing built on the platform's DOMParser (available in every browser
// and in jsdom under Vitest) — no bundled XML library needed. Elements are
// looked up by namespace URI rather than by prefix, since different
// producers (Word, LibreOffice, Google Docs) are free to bind the
// WordprocessingML namespace to any prefix.

export const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
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

/** Reads an attribute bound to the WordprocessingML namespace (e.g. the
 * `w:val` on `<w:b w:val="0"/>`) regardless of which prefix the document
 * actually bound that namespace to. */
export function wordAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(WORD_NAMESPACE, localName);
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
