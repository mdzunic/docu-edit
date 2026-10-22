// A forward-only XML scanner (SPEC-116): reads a part's text once, start to
// finish, and reports what it meets as events — open tags (with their
// attributes, namespace-resolved), close tags and text — each with its exact
// UTF-16 source offsets. Nothing is built: no tree, no node objects kept
// between events.
//
// It exists for the parts that are too large for DOMParser — a million-cell
// worksheet as a DOM costs an order of magnitude more memory than its text —
// and for readers that must slice the source by offset (an `.ods` reader cuts
// each table's verbatim XML out of content.xml at the offsets reported here).
// Small parts keep using core/xml.ts's parseXml.
//
// What it does, and deliberately does not:
// - namespaces are resolved from the in-scope `xmlns` / `xmlns:p`
//   declarations (an element's own declarations apply to it and its
//   attributes); unprefixed attributes have no namespace; the `xml` prefix is
//   predeclared; an unbound prefix resolves to `undefined` rather than failing;
// - the five predefined entities and numeric character references are
//   decoded, CDATA sections are unwrapped into text, line ends are normalised
//   (CRLF/CR → LF) and attribute values get XML's whitespace normalisation;
// - comments and processing instructions (the XML declaration included) are
//   skipped;
// - a `<!DOCTYPE` is refused with an XmlScanError: no producer of the formats
//   read with this writes one, and refusing it closes the entity-expansion
//   door (an undeclared entity reference is malformed for the same reason);
// - a self-closing element reports ONE event, `open` with `selfClosing: true`
//   — there is no matching `close` event for it;
// - returning "stop" from the callback ends the scan at once, with no further
//   events and no end-of-input checks (a reader's cell ceiling uses this).
// Product-neutral: nothing here knows about spreadsheets.

const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";

export interface XmlScanAttribute {
  name: string;
  local: string;
  ns: string | undefined;
  value: string;
}

export type XmlScanEvent =
  | {
      type: "open";
      name: string;
      local: string;
      ns: string | undefined;
      attrs: XmlScanAttribute[];
      selfClosing: boolean;
      start: number;
      end: number;
    }
  | { type: "close"; name: string; local: string; ns: string | undefined; start: number; end: number }
  | { type: "text"; text: string; start: number; end: number };

/** Malformed input, or a construct the scanner refuses (a DTD). `offset` is
 * the UTF-16 index into the scanned text where the problem was found. */
export class XmlScanError extends Error {
  readonly offset: number;

  constructor(message: string, offset: number) {
    super(`${message} (at offset ${offset})`);
    this.name = "XmlScanError";
    this.offset = offset;
  }
}

/** prefix → namespace URI; "" is the default namespace. */
type Scope = ReadonlyMap<string, string>;

interface OpenElement {
  name: string;
  local: string;
  ns: string | undefined;
  start: number;
  /** The scope the element's parent had — restored when it closes. */
  parentScope: Scope;
}

const EMPTY_SCOPE: Scope = new Map();

function isWhitespace(code: number): boolean {
  return code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d;
}

/** Name characters, loosely: letters, digits, `_ : . -` and anything
 * non-ASCII. Real names are a subset of this; the scanner only needs to know
 * where a name ends. */
function isNameChar(code: number): boolean {
  return (
    (code >= 0x61 && code <= 0x7a) || // a-z
    (code >= 0x41 && code <= 0x5a) || // A-Z
    (code >= 0x30 && code <= 0x39) || // 0-9
    code === 0x5f || // _
    code === 0x3a || // :
    code === 0x2e || // .
    code === 0x2d || // -
    code >= 0x80
  );
}

function localOf(name: string): string {
  const colon = name.indexOf(":");
  return colon < 0 ? name : name.slice(colon + 1);
}

function prefixOf(name: string): string {
  const colon = name.indexOf(":");
  return colon < 0 ? "" : name.slice(0, colon);
}

function resolvePrefix(prefix: string, scope: Scope): string | undefined {
  if (prefix === "xml") {
    return XML_NAMESPACE;
  }
  if (prefix === "xmlns") {
    return XMLNS_NAMESPACE;
  }
  const uri = scope.get(prefix);
  return uri === undefined || uri === "" ? undefined : uri;
}

/** Whether `code` is a character XML allows in a document (production
 * [2] Char) — what a numeric character reference may name. */
function isXmlChar(code: number): boolean {
  return (
    code === 0x09 ||
    code === 0x0a ||
    code === 0x0d ||
    (code >= 0x20 && code <= 0xd7ff) ||
    (code >= 0xe000 && code <= 0xfffd) ||
    (code >= 0x10000 && code <= 0x10ffff)
  );
}

/** Decodes one entity reference starting at `amp` (the index of `&` in
 * `raw`), returning the replacement and the index just past its `;`. */
function decodeEntity(raw: string, amp: number, base: number): [string, number] {
  const semi = raw.indexOf(";", amp + 1);
  if (semi < 0 || semi - amp > 12) {
    throw new XmlScanError("Unterminated entity reference", base + amp);
  }
  const body = raw.slice(amp + 1, semi);
  switch (body) {
    case "lt":
      return ["<", semi + 1];
    case "gt":
      return [">", semi + 1];
    case "amp":
      return ["&", semi + 1];
    case "quot":
      return ['"', semi + 1];
    case "apos":
      return ["'", semi + 1];
  }
  if (body.startsWith("#")) {
    const hex = body.startsWith("#x");
    const digits = body.slice(hex ? 2 : 1);
    const valid = hex ? /^[0-9A-Fa-f]+$/.test(digits) : /^[0-9]+$/.test(digits);
    const code = valid ? Number.parseInt(digits, hex ? 16 : 10) : Number.NaN;
    if (!Number.isFinite(code) || !isXmlChar(code)) {
      throw new XmlScanError(`Invalid character reference &${body};`, base + amp);
    }
    return [String.fromCodePoint(code), semi + 1];
  }
  throw new XmlScanError(`Unknown entity &${body};`, base + amp);
}

/** Character data: CRLF/CR → LF, entities decoded. `base` is the offset of
 * `raw` in the scanned text, for error offsets. */
function decodeText(raw: string, base: number): string {
  if (raw.indexOf("&") < 0 && raw.indexOf("\r") < 0) {
    return raw;
  }
  let out = "";
  let from = 0;
  for (let i = 0; i < raw.length; i += 1) {
    const code = raw.charCodeAt(i);
    if (code === 0x26 /* & */) {
      out += raw.slice(from, i);
      const [replacement, next] = decodeEntity(raw, i, base);
      out += replacement;
      from = next;
      i = next - 1;
    } else if (code === 0x0d /* \r */) {
      out += `${raw.slice(from, i)}\n`;
      if (raw.charCodeAt(i + 1) === 0x0a) {
        i += 1;
      }
      from = i + 1;
    }
  }
  return out + raw.slice(from);
}

/** An attribute value: as text, plus XML's attribute-value normalisation —
 * each literal tab, line feed or carriage return (a CRLF pair counting once)
 * becomes a space; characters written as references are kept as they are. */
function decodeAttributeValue(raw: string, base: number): string {
  if (raw.indexOf("&") < 0 && !/[\t\n\r]/.test(raw)) {
    return raw;
  }
  let out = "";
  let from = 0;
  for (let i = 0; i < raw.length; i += 1) {
    const code = raw.charCodeAt(i);
    if (code === 0x26 /* & */) {
      out += raw.slice(from, i);
      const [replacement, next] = decodeEntity(raw, i, base);
      out += replacement;
      from = next;
      i = next - 1;
    } else if (code === 0x0d || code === 0x0a || code === 0x09) {
      out += `${raw.slice(from, i)} `;
      if (code === 0x0d && raw.charCodeAt(i + 1) === 0x0a) {
        i += 1;
      }
      from = i + 1;
    }
  }
  return out + raw.slice(from);
}

/** CDATA content is literal — only line ends are normalised. */
function normaliseLineEnds(raw: string): string {
  return raw.indexOf("\r") < 0 ? raw : raw.replace(/\r\n?/g, "\n");
}

/** Scans `text` from start to end, calling `onEvent` for every open tag,
 * close tag and text run in document order (see the module header for the
 * event rules). Throws XmlScanError on malformed input or a DTD, with the
 * offending offset. Returning "stop" from `onEvent` ends the scan. */
export function scanXml(text: string, onEvent: (event: XmlScanEvent) => void | "stop"): void {
  const length = text.length;
  let pos = text.charCodeAt(0) === 0xfeff ? 1 : 0;
  const stack: OpenElement[] = [];
  let scope: Scope = EMPTY_SCOPE;
  let rootClosed = false;
  let sawRoot = false;

  while (pos < length) {
    const lt = text.indexOf("<", pos);
    const textEnd = lt < 0 ? length : lt;
    if (textEnd > pos) {
      if (stack.length === 0) {
        for (let i = pos; i < textEnd; i += 1) {
          if (!isWhitespace(text.charCodeAt(i))) {
            throw new XmlScanError("Text outside the root element", i);
          }
        }
      } else if (onEvent({ type: "text", text: decodeText(text.slice(pos, textEnd), pos), start: pos, end: textEnd }) === "stop") {
        return;
      }
      pos = textEnd;
      if (lt < 0) {
        break;
      }
    }

    const next = text.charCodeAt(pos + 1);
    if (next === 0x3f /* ? */) {
      const close = text.indexOf("?>", pos + 2);
      if (close < 0) {
        throw new XmlScanError("Unterminated processing instruction", pos);
      }
      pos = close + 2;
      continue;
    }
    if (next === 0x21 /* ! */) {
      if (text.startsWith("<!--", pos)) {
        const close = text.indexOf("-->", pos + 4);
        if (close < 0) {
          throw new XmlScanError("Unterminated comment", pos);
        }
        pos = close + 3;
        continue;
      }
      if (text.startsWith("<![CDATA[", pos)) {
        if (stack.length === 0) {
          throw new XmlScanError("CDATA outside the root element", pos);
        }
        const close = text.indexOf("]]>", pos + 9);
        if (close < 0) {
          throw new XmlScanError("Unterminated CDATA section", pos);
        }
        const content = normaliseLineEnds(text.slice(pos + 9, close));
        if (onEvent({ type: "text", text: content, start: pos, end: close + 3 }) === "stop") {
          return;
        }
        pos = close + 3;
        continue;
      }
      if (text.startsWith("<!DOCTYPE", pos)) {
        throw new XmlScanError("A document type declaration (DTD) is not accepted", pos);
      }
      throw new XmlScanError("Malformed markup declaration", pos);
    }

    if (next === 0x2f /* / */) {
      let i = pos + 2;
      const nameStart = i;
      while (i < length && isNameChar(text.charCodeAt(i))) {
        i += 1;
      }
      const name = text.slice(nameStart, i);
      while (i < length && isWhitespace(text.charCodeAt(i))) {
        i += 1;
      }
      if (name.length === 0 || text.charCodeAt(i) !== 0x3e /* > */) {
        throw new XmlScanError("Malformed closing tag", pos);
      }
      const open = stack.pop();
      if (!open || open.name !== name) {
        throw new XmlScanError(
          open ? `Closing tag </${name}> does not match <${open.name}>` : `Closing tag </${name}> has no open element`,
          pos,
        );
      }
      scope = open.parentScope;
      if (stack.length === 0) {
        rootClosed = true;
      }
      const end = i + 1;
      if (onEvent({ type: "close", name, local: open.local, ns: open.ns, start: pos, end }) === "stop") {
        return;
      }
      pos = end;
      continue;
    }

    // An open tag.
    if (rootClosed) {
      throw new XmlScanError("Content after the root element", pos);
    }
    let i = pos + 1;
    const nameStart = i;
    while (i < length && isNameChar(text.charCodeAt(i))) {
      i += 1;
    }
    if (i === nameStart) {
      throw new XmlScanError("Expected an element name after '<'", pos);
    }
    const name = text.slice(nameStart, i);
    const rawAttrs: Array<{ name: string; value: string }> = [];
    let declaresNamespaces = false;
    let selfClosing = false;
    for (;;) {
      const wsStart = i;
      while (i < length && isWhitespace(text.charCodeAt(i))) {
        i += 1;
      }
      if (i >= length) {
        throw new XmlScanError(`Unterminated tag <${name}>`, pos);
      }
      const code = text.charCodeAt(i);
      if (code === 0x3e /* > */) {
        i += 1;
        break;
      }
      if (code === 0x2f /* / */) {
        if (text.charCodeAt(i + 1) !== 0x3e) {
          throw new XmlScanError(`Malformed tag <${name}>`, i);
        }
        i += 2;
        selfClosing = true;
        break;
      }
      if (i === wsStart) {
        throw new XmlScanError(`Expected whitespace before an attribute in <${name}>`, i);
      }
      const attrStart = i;
      while (i < length && isNameChar(text.charCodeAt(i))) {
        i += 1;
      }
      if (i === attrStart) {
        throw new XmlScanError(`Malformed attribute in <${name}>`, i);
      }
      const attrName = text.slice(attrStart, i);
      while (i < length && isWhitespace(text.charCodeAt(i))) {
        i += 1;
      }
      if (text.charCodeAt(i) !== 0x3d /* = */) {
        throw new XmlScanError(`Attribute ${attrName} has no value`, i);
      }
      i += 1;
      while (i < length && isWhitespace(text.charCodeAt(i))) {
        i += 1;
      }
      const quote = text.charCodeAt(i);
      if (quote !== 0x22 && quote !== 0x27) {
        throw new XmlScanError(`Attribute ${attrName} value is not quoted`, i);
      }
      const valueStart = i + 1;
      const valueEnd = text.indexOf(quote === 0x22 ? '"' : "'", valueStart);
      if (valueEnd < 0) {
        throw new XmlScanError(`Unterminated value for attribute ${attrName}`, i);
      }
      const raw = text.slice(valueStart, valueEnd);
      const stray = raw.indexOf("<");
      if (stray >= 0) {
        throw new XmlScanError(`'<' in the value of attribute ${attrName}`, valueStart + stray);
      }
      for (const existing of rawAttrs) {
        if (existing.name === attrName) {
          throw new XmlScanError(`Duplicate attribute ${attrName}`, attrStart);
        }
      }
      if (attrName === "xmlns" || attrName.startsWith("xmlns:")) {
        declaresNamespaces = true;
      }
      rawAttrs.push({ name: attrName, value: decodeAttributeValue(raw, valueStart) });
      i = valueEnd + 1;
    }

    let elementScope = scope;
    if (declaresNamespaces) {
      const declared = new Map(scope);
      for (const attr of rawAttrs) {
        if (attr.name === "xmlns") {
          declared.set("", attr.value);
        } else if (attr.name.startsWith("xmlns:")) {
          declared.set(attr.name.slice(6), attr.value);
        }
      }
      elementScope = declared;
    }
    const prefix = prefixOf(name);
    const local = localOf(name);
    const ns = prefix === "" ? resolvePrefix("", elementScope) : resolvePrefix(prefix, elementScope);
    const attrs: XmlScanAttribute[] = rawAttrs.map((attr) => {
      const attrPrefix = attr.name === "xmlns" ? "xmlns" : prefixOf(attr.name);
      return {
        name: attr.name,
        local: attr.name === "xmlns" ? "xmlns" : localOf(attr.name),
        ns: attrPrefix === "" ? undefined : resolvePrefix(attrPrefix, elementScope),
        value: attr.value,
      };
    });

    sawRoot = true;
    if (selfClosing) {
      if (stack.length === 0) {
        rootClosed = true;
      }
    } else {
      stack.push({ name, local, ns, start: pos, parentScope: scope });
      scope = elementScope;
    }
    if (onEvent({ type: "open", name, local, ns, attrs, selfClosing, start: pos, end: i }) === "stop") {
      return;
    }
    pos = i;
  }

  const unclosed = stack.at(-1);
  if (unclosed) {
    throw new XmlScanError(`Element <${unclosed.name}> is never closed`, unclosed.start);
  }
  if (!sawRoot) {
    throw new XmlScanError("No root element", length);
  }
}

/** The value of the attribute with local name `local` in namespace `ns`
 * (`undefined` for an unprefixed attribute), or undefined when absent. */
export function scanAttr(
  attrs: readonly XmlScanAttribute[],
  local: string,
  ns?: string,
): string | undefined {
  for (const attr of attrs) {
    if (attr.local === local && attr.ns === ns) {
      return attr.value;
    }
  }
  return undefined;
}
