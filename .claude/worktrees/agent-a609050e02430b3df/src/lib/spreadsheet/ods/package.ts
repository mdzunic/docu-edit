// The ODF package plumbing the workbook writer and the sheet commands share
// (SPEC-121): a sheet's `<table:table>` start tag inside its retained
// `sourceXml`, the source workbook's table names (walked out of the retained
// content.xml without parsing a cell), `META-INF/manifest.xml` entries and
// their splice, the embedded objects (`Object N/`, `ObjectReplacements/`)
// a table's text points at, and the id renaming a duplicate needs.
//
// content.xml is the reader's splice base: envelope prefix + every table's
// sourceXml (the whitespace before a table included) + suffix. Between the
// first and the last table the ODF schema allows only tables, so the source
// tables are found by walking that middle stretch table by table — counting
// nested `table:table` elements (a table inside a drawing) — never by
// scanning the cells.
import { scanAttr, scanXml } from "../../core/xmlScanner";
import { applySplices, attributeValue, escapeAttributeValue, findAttribute, startTagAt, type Splice } from "../xmlSplice";

/** The `<table:table …>` start tag of a sheet's sourceXml and where it
 * starts (after the whitespace that precedes the table). */
export function tableStartTag(sourceXml: string): { start: number; tag: string } {
  const start = sourceXml.indexOf("<");
  return { start, tag: startTagAt(sourceXml, start) };
}

/** A table's `table:name` as its sourceXml declares it. */
export function tableNameOf(sourceXml: string): string | undefined {
  return attributeValue(tableStartTag(sourceXml).tag, "table:name");
}

/** A table's `table:style-name`. */
export function tableStyleNameOf(sourceXml: string): string | undefined {
  return attributeValue(tableStartTag(sourceXml).tag, "table:style-name");
}

/** sourceXml with its start tag's `attribute` set (or removed when null). */
export function setTableAttribute(sourceXml: string, attribute: string, value: string | null): string {
  const { start, tag } = tableStartTag(sourceXml);
  const next = setTagAttribute(tag, attribute, value);
  return next === tag ? sourceXml : sourceXml.slice(0, start) + next + sourceXml.slice(start + tag.length);
}

function setTagAttribute(tag: string, attribute: string, value: string | null): string {
  const span = findAttribute(tag, attribute);
  if (value === null) {
    return span ? tag.slice(0, span.start) + tag.slice(span.end) : tag;
  }
  const escaped = escapeAttributeValue(value);
  if (span) {
    return tag.slice(0, span.valueStart) + escaped + tag.slice(span.valueEnd);
  }
  const close = tag.endsWith("/>") ? tag.length - 2 : tag.length - 1;
  return `${tag.slice(0, close)} ${attribute}="${escaped}"${tag.slice(close)}`;
}

/** One table of the retained content.xml. */
export interface SourceTable {
  name: string;
  /** Offsets of the table's text (its leading whitespace included). */
  start: number;
  end: number;
}

/** The top-level tables of a retained content.xml, in order, between the
 * envelope's prefix and suffix. */
export function sourceTables(contentXml: string, prefixLength: number, suffixLength: number): SourceTable[] {
  const tables: SourceTable[] = [];
  const stop = contentXml.length - suffixLength;
  let pos = prefixLength;
  while (pos < stop) {
    const lt = contentXml.indexOf("<", pos);
    if (lt < 0 || lt >= stop) {
      break;
    }
    if (contentXml.startsWith("<!--", lt)) {
      const close = contentXml.indexOf("-->", lt);
      pos = close < 0 ? stop : close + 3;
      continue;
    }
    const tag = startTagAt(contentXml, lt);
    const elementName = /^<([\w.:-]+)/.exec(tag)?.[1] ?? "";
    const name = attributeValue(tag, "table:name") ?? "";
    let end: number;
    if (tag.endsWith("/>")) {
      end = lt + tag.length;
    } else {
      end = matchingClose(contentXml, lt + tag.length, elementName);
    }
    tables.push({ name, start: pos, end });
    pos = end;
  }
  return tables;
}

/** The index just past the `</name>` closing the element whose start tag
 * ended at `from`, nested same-name elements counted. */
function matchingClose(xml: string, from: number, name: string): number {
  const open = `<${name}`;
  const close = `</${name}>`;
  let depth = 1;
  let pos = from;
  while (depth > 0) {
    const nextClose = xml.indexOf(close, pos);
    if (nextClose < 0) {
      return xml.length;
    }
    let nextOpen = xml.indexOf(open, pos);
    while (nextOpen >= 0 && nextOpen < nextClose && !/[\s/>]/.test(xml[nextOpen + open.length] ?? "")) {
      nextOpen = xml.indexOf(open, nextOpen + open.length);
    }
    if (nextOpen >= 0 && nextOpen < nextClose) {
      const tag = startTagAt(xml, nextOpen);
      if (!tag.endsWith("/>")) {
        depth += 1;
      }
      pos = nextOpen + tag.length;
    } else {
      depth -= 1;
      pos = nextClose + close.length;
    }
  }
  return pos;
}

// --- the manifest ---------------------------------------------------------------------

export interface ManifestEntry {
  fullPath: string;
  mediaType: string;
  start: number;
  end: number;
}

/** The `manifest:file-entry` elements, in order. */
export function readManifest(xml: string): ManifestEntry[] {
  const entries: ManifestEntry[] = [];
  let pending: { attrs: Parameters<typeof scanAttr>[0]; start: number } | null = null;
  const MANIFEST = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
  const push = (attrs: Parameters<typeof scanAttr>[0], start: number, end: number): void => {
    entries.push({
      fullPath: scanAttr(attrs, "full-path", MANIFEST) ?? "",
      mediaType: scanAttr(attrs, "media-type", MANIFEST) ?? "",
      start,
      end,
    });
  };
  scanXml(xml, (event) => {
    if (event.type === "open" && event.local === "file-entry") {
      if (event.selfClosing) {
        push(event.attrs, event.start, event.end);
      } else {
        pending = { attrs: event.attrs, start: event.start };
      }
    } else if (event.type === "close" && event.local === "file-entry" && pending) {
      push(pending.attrs, pending.start, event.end);
      pending = null;
    }
  });
  return entries;
}

/** The manifest with the entries for `remove` spliced out and `add`
 * (whole elements) inserted before its closing tag. */
export function spliceManifest(xml: string, remove: ReadonlySet<string>, add: readonly string[]): string {
  const splices: Splice[] = readManifest(xml)
    .filter((entry) => remove.has(entry.fullPath))
    .map((entry) => ({ start: entry.start, end: entry.end, text: "" }));
  let out = applySplices(xml, splices);
  if (add.length > 0) {
    const close = out.lastIndexOf("</");
    out = out.slice(0, close) + add.join("") + out.slice(close);
  }
  return out;
}

/** A manifest entry element for `fullPath`, shaped like `like` (its
 * prefix and attribute order kept) — the copy of an entry under a new
 * path. */
export function manifestEntryLike(xml: string, like: ManifestEntry, fullPath: string): string {
  const element = xml.slice(like.start, like.end);
  const tag = startTagAt(element, 0);
  const span = findAttribute(tag, "manifest:full-path");
  if (!span) {
    return element;
  }
  return element.slice(0, span.valueStart) + escapeAttributeValue(fullPath) + element.slice(span.valueEnd);
}

// --- embedded objects -------------------------------------------------------------------

const OBJECT_HREF = /\bxlink:href\s*=\s*"(?:\.\/)?(Object [^"/]+|ObjectReplacements\/Object [^"/]+)"/g;

/** The embedded objects (`Object N` directory names) a table's text points
 * at, in order of first reference. */
export function objectsReferenced(xml: string): string[] {
  const found: string[] = [];
  for (const match of xml.matchAll(OBJECT_HREF)) {
    const target = match[1]!;
    const object = target.startsWith("ObjectReplacements/") ? target.slice("ObjectReplacements/".length) : target;
    if (!found.includes(object)) {
      found.push(object);
    }
  }
  return found;
}

/** The pictures (`Pictures/…`) a text points at. */
export function picturesReferenced(xml: string): string[] {
  const found = new Set<string>();
  for (const match of xml.matchAll(/\bxlink:href\s*=\s*"(?:\.\/)?(Pictures\/[^"]+)"/g)) {
    found.add(match[1]!);
  }
  return [...found];
}

/** Whether a package part belongs to embedded object `object` (its
 * directory, or its replacement image). */
export function isObjectPart(partName: string, object: string): boolean {
  return partName.startsWith(`${object}/`) || partName === `ObjectReplacements/${object}`;
}

/** The next free `Object N` name over the package's parts and `taken`. */
export function nextFreeObjectName(partNames: Iterable<string>, taken: ReadonlySet<string> = new Set()): string {
  let max = 0;
  const consider = (name: string): void => {
    const match = /^(?:ObjectReplacements\/)?Object (\d+)(?:\/|$)/.exec(name);
    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  };
  for (const name of partNames) {
    consider(name);
  }
  for (const name of taken) {
    consider(name);
  }
  return `Object ${max + 1}`;
}

/** `xml` with every href to `from` (directory or replacement) pointing at
 * `to` instead. */
export function retargetObject(xml: string, from: string, to: string): string {
  return xml.replace(OBJECT_HREF, (whole, target: string) => {
    if (target === from) {
      return whole.replace(from, to);
    }
    if (target === `ObjectReplacements/${from}`) {
      return whole.replace(`ObjectReplacements/${from}`, `ObjectReplacements/${to}`);
    }
    return whole;
  });
}

// --- ids ----------------------------------------------------------------------------------

/** The attributes whose values are (or point at) element ids. */
const ID_ATTRIBUTES = ["xml:id", "draw:id", "draw:start-shape", "draw:end-shape", "draw:caption-id", "text:id"];
const ID_DECLARATIONS = /\s(?:xml:id|draw:id)\s*=\s*"([^"]*)"/g;

/** The ids a text declares (`xml:id`, `draw:id`). */
export function declaredIds(xml: string): Set<string> {
  const ids = new Set<string>();
  for (const match of xml.matchAll(ID_DECLARATIONS)) {
    ids.add(match[1]!);
  }
  return ids;
}

/** `xml` with every id it declares suffixed `suffix`, consistently: the
 * declarations and every id-valued attribute pointing at one of them. */
export function suffixIds(xml: string, suffix: string): string {
  const ids = declaredIds(xml);
  if (ids.size === 0) {
    return xml;
  }
  const pattern = new RegExp(`(\\s(?:${ID_ATTRIBUTES.map((name) => name.replace(/[.:-]/g, "\\$&")).join("|")})\\s*=\\s*")([^"]*)"`, "g");
  return xml.replace(pattern, (whole, head: string, value: string) => (ids.has(value) ? `${head}${value}${suffix}"` : whole));
}
