// WordprocessingML-specific XML helpers. The format-neutral machinery —
// parsing, serialisation, namespace-aware child lookup and OPC
// relationships — lives in src/lib/core/xml.ts (shared with any OPC/ODF
// package, e.g. a future .pptx reader).

export const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** Reads an attribute bound to the WordprocessingML namespace (e.g. the
 * `w:val` on `<w:b w:val="0"/>`) regardless of which prefix the document
 * actually bound that namespace to. */
export function wordAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(WORD_NAMESPACE, localName);
}
