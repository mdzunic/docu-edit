// Offset splicing of retained XML (SPEC-121): the writers and the sheet
// commands change a retained part by replacing exact source ranges — an
// element core/xmlScanner.ts located, an attribute value inside its start
// tag — and never by re-serialising a parsed tree, so every byte nothing
// touched survives (SPEC-116's invariant 3, "the source is the splice base").
//
// - A start tag is edited textually: findAttribute locates one attribute's
//   raw (still-escaped) value inside the tag by its qualified name as
//   written; setAttribute replaces it, inserts it before the tag's end, or
//   removes it with the whitespace before it.
// - applySplices replaces non-overlapping [start, end) ranges of a text in
//   one pass; the ranges may arrive in any order.
// - Values written into attributes are escaped with core/xml.ts's
//   escapeXmlAttribute; element text with escapeXmlText. decodeXmlText
//   reverses the five predefined entities and numeric references, for raw
//   text sliced out of a part.

/** One replacement of `text.slice(start, end)`. */
export interface Splice {
  start: number;
  end: number;
  text: string;
}

/** `text` with every splice applied (ranges must not overlap). */
export function applySplices(text: string, splices: readonly Splice[]): string {
  if (splices.length === 0) {
    return text;
  }
  const ordered = splices.toSorted((a, b) => a.start - b.start || a.end - b.end);
  let out = "";
  let at = 0;
  for (const splice of ordered) {
    if (splice.start < at) {
      throw new Error(`overlapping splices at ${splice.start}`);
    }
    out += text.slice(at, splice.start) + splice.text;
    at = splice.end;
  }
  return out + text.slice(at);
}

/** Where an attribute sits in a start tag's text: the whole `name="value"`
 * (with the whitespace before it) and its raw value. */
export interface AttributeSpan {
  /** Index of the whitespace before the name. */
  start: number;
  /** Index just past the closing quote. */
  end: number;
  valueStart: number;
  valueEnd: number;
  /** The raw value, entities still escaped. */
  raw: string;
}

const NAME_CHAR = /[\w:.\-·À-￿]/;

/** The attribute named `qname` (as written, prefix included) in the start
 * tag at the head of `tag`. */
export function findAttribute(tag: string, qname: string): AttributeSpan | undefined {
  let i = 1;
  // Skip the element name.
  while (i < tag.length && NAME_CHAR.test(tag[i]!)) {
    i += 1;
  }
  while (i < tag.length) {
    const wsStart = i;
    while (i < tag.length && /\s/.test(tag[i]!)) {
      i += 1;
    }
    if (i >= tag.length || tag[i] === ">" || tag[i] === "/") {
      return undefined;
    }
    const nameStart = i;
    while (i < tag.length && NAME_CHAR.test(tag[i]!)) {
      i += 1;
    }
    const name = tag.slice(nameStart, i);
    while (i < tag.length && /\s/.test(tag[i]!)) {
      i += 1;
    }
    if (tag[i] !== "=") {
      return undefined;
    }
    i += 1;
    while (i < tag.length && /\s/.test(tag[i]!)) {
      i += 1;
    }
    const quote = tag[i];
    if (quote !== '"' && quote !== "'") {
      return undefined;
    }
    const valueStart = i + 1;
    const valueEnd = tag.indexOf(quote, valueStart);
    if (valueEnd < 0) {
      return undefined;
    }
    i = valueEnd + 1;
    if (name === qname) {
      return { start: wsStart, end: i, valueStart, valueEnd, raw: tag.slice(valueStart, valueEnd) };
    }
  }
  return undefined;
}

/** The decoded value of an attribute of the start tag at the head of
 * `tag`, or undefined. */
export function attributeValue(tag: string, qname: string): string | undefined {
  const span = findAttribute(tag, qname);
  return span ? decodeXmlText(span.raw) : undefined;
}

/** `tag` with attribute `qname` set to `value` (escaped here) — replaced in
 * place, or appended before the tag's end — or removed when `value` is
 * null. The same string when nothing changes. */
export function setAttribute(tag: string, qname: string, value: string | null): string {
  const span = findAttribute(tag, qname);
  if (value === null) {
    return span ? tag.slice(0, span.start) + tag.slice(span.end) : tag;
  }
  const escaped = escapeAttributeValue(value);
  if (span) {
    return span.raw === escaped ? tag : tag.slice(0, span.valueStart) + escaped + tag.slice(span.valueEnd);
  }
  const close = tag.endsWith("/>") ? tag.length - 2 : tag.length - 1;
  let insertAt = close;
  while (insertAt > 0 && /\s/.test(tag[insertAt - 1]!)) {
    insertAt -= 1;
  }
  return `${tag.slice(0, insertAt)} ${qname}="${escaped}"${tag.slice(insertAt)}`;
}

/** An attribute value as written between double quotes. */
export function escapeAttributeValue(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Element text as written: `&`, `<` and `>` escaped. */
export function escapeElementText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

const ENTITIES: Readonly<Record<string, string>> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };

/** Raw XML text (an attribute value or element content) with the five
 * predefined entities and numeric character references decoded. */
export function decodeXmlText(raw: string): string {
  if (raw.indexOf("&") < 0) {
    return raw;
  }
  return raw.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (whole, body: string) => {
    if (body.startsWith("#x") || body.startsWith("#X")) {
      return String.fromCodePoint(Number.parseInt(body.slice(2), 16));
    }
    if (body.startsWith("#")) {
      return String.fromCodePoint(Number.parseInt(body.slice(1), 10));
    }
    return ENTITIES[body] ?? whole;
  });
}

/** The start tag beginning at `start` in `xml` (through its `>`). */
export function startTagAt(xml: string, start: number): string {
  let i = start + 1;
  let quote = "";
  while (i < xml.length) {
    const ch = xml[i]!;
    if (quote) {
      if (ch === quote) {
        quote = "";
      }
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === ">") {
      return xml.slice(start, i + 1);
    }
    i += 1;
  }
  return xml.slice(start);
}
