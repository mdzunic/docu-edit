// A small element tree of an XML text or slice (SPEC-129), built from
// core/xmlScanner.ts's events, every node keeping its source offsets — what
// the readers of the few, small, verbatim-kept structures use (a worksheet's
// `<conditionalFormatting>` blocks and `<extLst>`, styles.xml's `<dxfs>`)
// where a DOM would lose the exact source text a writer re-emits.
//
// A slice is read inside `context`: the start tags (outermost first — the
// part's root start tag, say) whose namespace declarations are in scope where
// the slice was cut, so a prefixed `x:cfRule` or an `x14:cfRule` whose prefix
// the root or an `<ext>` declares resolves as it did in the part. Offsets are
// into the slice itself. Text directly inside an element is joined into its
// `text` (decoded, line ends normalised, as the scanner reports it).
import { scanXml, XmlScanError, type XmlScanAttribute } from "../../core/xmlScanner";

export interface XmlNode {
  /** The qualified name as written. */
  name: string;
  local: string;
  ns: string | undefined;
  attrs: readonly XmlScanAttribute[];
  /** Offset of the element's `<`. */
  start: number;
  /** Offset just past the start tag's `>`. */
  tagEnd: number;
  /** Offset just past the element (its close tag, or the start tag when self-closing). */
  end: number;
  selfClosing: boolean;
  children: XmlNode[];
  text: string;
}

/** The qualified name a start tag is written with. */
function tagName(tag: string): string {
  return /^<([^\s/>]+)/.exec(tag)?.[1] ?? "";
}

/** The top-level elements of `xml` (one, the root, for a whole document),
 * read inside `context` (see the module header); undefined when the text is
 * not well-formed. */
export function elementTree(xml: string, context: readonly string[] = []): XmlNode[] | undefined {
  const open = context.map((tag) => (tag.endsWith("/>") ? `${tag.slice(0, -2).trimEnd()}>` : tag));
  const head = open.join("");
  const tail = open
    .map((tag) => `</${tagName(tag)}>`)
    .toReversed()
    .join("");
  const offset = head.length;
  const top: XmlNode[] = [];
  const stack: XmlNode[] = [];
  let depth = 0;
  try {
    scanXml(`${head}${xml}${tail}`, (event) => {
      if (event.type === "open") {
        depth += 1;
        if (depth > open.length) {
          const node: XmlNode = {
            name: event.name,
            local: event.local,
            ns: event.ns,
            attrs: event.attrs,
            start: event.start - offset,
            tagEnd: event.end - offset,
            end: event.end - offset,
            selfClosing: event.selfClosing,
            children: [],
            text: "",
          };
          (stack.at(-1)?.children ?? top).push(node);
          if (!event.selfClosing) {
            stack.push(node);
          }
        }
        if (event.selfClosing) {
          depth -= 1;
        }
      } else if (event.type === "close") {
        if (depth > open.length) {
          stack.pop()!.end = event.end - offset;
        }
        depth -= 1;
      } else if (depth > open.length) {
        stack.at(-1)!.text += event.text;
      }
    });
  } catch (error) {
    if (error instanceof XmlScanError) {
      return undefined;
    }
    throw error;
  }
  return top;
}

/** An attribute's value by its local name (unprefixed attributes only). */
export function nodeAttr(node: XmlNode, local: string): string | undefined {
  for (const attr of node.attrs) {
    if (attr.local === local && attr.ns === undefined) {
      return attr.value;
    }
  }
  return undefined;
}

/** The children of `node` named `local` in a namespace `inNs` accepts. */
export function childNodes(node: XmlNode, local: string, inNs: (ns: string | undefined) => boolean): XmlNode[] {
  return node.children.filter((child) => child.local === local && inNs(child.ns));
}

/** The whitespace run that ends at `at` in `xml` — its start offset. */
export function whitespaceBefore(xml: string, at: number): number {
  let start = at;
  while (start > 0 && /\s/.test(xml[start - 1]!)) {
    start -= 1;
  }
  return start;
}
