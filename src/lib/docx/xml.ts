// WordprocessingML-specific XML helpers. The format-neutral machinery —
// parsing, serialisation, namespace-aware child lookup and OPC
// relationships — moved to src/lib/core/xml.ts (shared with any OPC/ODF
// package, e.g. a future .pptx reader); it is re-exported here so existing
// import paths keep working.
export {
  childElements,
  firstChildElement,
  parseRelationships,
  parseXml,
  RELATIONSHIPS_NAMESPACE,
  relsPathFor,
  resolvePartPath,
  serializeElement,
  XML_NAMESPACE,
  XmlParseError,
} from "../core/xml";
export type { Relationship } from "../core/xml";

export const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** Reads an attribute bound to the WordprocessingML namespace (e.g. the
 * `w:val` on `<w:b w:val="0"/>`) regardless of which prefix the document
 * actually bound that namespace to. */
export function wordAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(WORD_NAMESPACE, localName);
}
