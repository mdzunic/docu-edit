// The OPC package plumbing the workbook writer and the sheet commands share
// (SPEC-121): relationships read WITH their element offsets (so a `.rels`
// part is edited by splicing — core/xml.ts's parseRelationships drops
// `TargetMode`, and re-serialising would lose it), the source workbook's
// sheet list as the retained `xl/workbook.xml` declares it, the package's
// relationship graph (what reaches what, from `_rels/.rels`), the "next
// free name" rule for parts a duplicate adds, and the `[Content_Types].xml`
// Override splice.
//
// Every reader here takes a PartLookup — a decoded part map, or LazyParts
// over the model's `opaque`, which decodes a part's base64 only when asked
// (a sheet command must not pay for the whole package).
import { base64ToBytes, type OpaquePart } from "../../core/packageParts";
import { relsPathFor, resolvePartPath } from "../../core/xml";
import { scanAttr, scanXml } from "../../core/xmlScanner";
import { applySplices, escapeAttributeValue, type Splice } from "../xmlSplice";

/** What the package helpers read parts through. */
export interface PartLookup {
  get(name: string): Uint8Array | undefined;
  has(name: string): boolean;
  keys(): Iterable<string>;
}

const utf8 = new TextDecoder("utf-8");
const encoder = new TextEncoder();

export function decodeUtf8(bytes: Uint8Array): string {
  return utf8.decode(bytes);
}

export function encodeUtf8(text: string): Uint8Array {
  return encoder.encode(text);
}

/** A part's text, or undefined when the package has no such part. */
export function partTextOf(parts: PartLookup, name: string): string | undefined {
  const bytes = parts.get(name);
  return bytes ? utf8.decode(bytes) : undefined;
}

/** Base64 characters decoded at each end of a large part by
 * worksheetTextWithoutCells (a multiple of 4: whole quartets). */
const WORKSHEET_END_CHARS = 65_536;

/** The tail worksheetTextWithoutCells searches at most (base64 characters)
 * before it decodes the part whole. */
const WORKSHEET_TAIL_MAX_CHARS = 4 * 1024 * 1024;

/** SPEC-137: a retained worksheet part's text with its `<sheetData>`
 * content cut out — what the readers of the parts around the cells (the
 * hyperlinks, the legacy drawing a note needs …) take, as the read phases
 * hand them the worksheet cut (xlsx/phases.ts). Decoded from the part's
 * base64 a slice at a time from each end, so a large sheet's cells — all but
 * a few KB of its part — are never decoded: a 25 MB sheet's whole decode was
 * a ~50 ms task right after its open. A part where the cut does not show
 * that way (no sheetData start tag in its head; a tail past
 * WORKSHEET_TAIL_MAX_CHARS) is decoded whole. */
export function worksheetTextWithoutCells(base64: string): string {
  if (base64.length <= 4 * WORKSHEET_END_CHARS) {
    return decodeUtf8(base64ToBytes(base64));
  }
  // A slice may end, or start, inside a character: the decoder marks it,
  // and both cuts below are at ASCII markup away from those ends.
  const head = utf8.decode(base64ToBytes(base64.slice(0, WORKSHEET_END_CHARS)));
  const open = /<([\w.-]+:)?sheetData\b[^>]*>/.exec(head);
  if (open !== null && !open[0].endsWith("/>")) {
    const close = `</${open[1] ?? ""}sheetData`;
    for (let size = WORKSHEET_END_CHARS; size <= WORKSHEET_TAIL_MAX_CHARS; size *= 4) {
      const start = Math.max(WORKSHEET_END_CHARS, base64.length - size - ((base64.length - size) % 4));
      const tail = utf8.decode(base64ToBytes(base64.slice(start)));
      const at = tail.lastIndexOf(close);
      if (at >= 0) {
        return head.slice(0, open.index + open[0].length) + tail.slice(at);
      }
      if (start === WORKSHEET_END_CHARS) {
        break;
      }
    }
  }
  return decodeUtf8(base64ToBytes(base64));
}

/** The model's opaque parts as a PartLookup that decodes each part's
 * base64 at most once, when first asked for. */
export class LazyParts implements PartLookup {
  private readonly byName = new Map<string, OpaquePart>();
  private readonly decoded = new Map<string, Uint8Array>();

  constructor(opaque: readonly OpaquePart[]) {
    for (const part of opaque) {
      this.byName.set(part.name, part);
    }
  }

  get(name: string): Uint8Array | undefined {
    const cached = this.decoded.get(name);
    if (cached) {
      return cached;
    }
    const part = this.byName.get(name);
    if (!part) {
      return undefined;
    }
    const bytes = base64ToBytes(part.data);
    this.decoded.set(name, bytes);
    return bytes;
  }

  has(name: string): boolean {
    return this.byName.has(name);
  }

  keys(): Iterable<string> {
    return this.byName.keys();
  }
}

// --- relationships ----------------------------------------------------------------

/** Relationship types by their last segment (transitional and strict URIs,
 * and Microsoft's 2017 ones, share it). */
export function relationshipKind(type: string): string {
  return type.slice(type.lastIndexOf("/") + 1);
}

export const RELATIONSHIP_TYPES = {
  worksheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
  officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  calcChain: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
} as const;

/** One `<Relationship>` element of a `.rels` part. */
export interface RelEntry {
  id: string;
  type: string;
  target: string;
  external: boolean;
  /** The element's offsets in the `.rels` text. */
  start: number;
  end: number;
}

/** The `<Relationship>` elements of a `.rels` part, in order. */
export function readRels(xml: string): RelEntry[] {
  const entries: RelEntry[] = [];
  let open: { attrs: Parameters<typeof scanAttr>[0]; start: number } | null = null;
  scanXml(xml, (event) => {
    if (event.type === "open" && event.local === "Relationship") {
      if (event.selfClosing) {
        push(event.attrs, event.start, event.end);
      } else {
        open = { attrs: event.attrs, start: event.start };
      }
    } else if (event.type === "close" && event.local === "Relationship" && open) {
      push(open.attrs, open.start, event.end);
      open = null;
    }
  });
  return entries;

  function push(attrs: Parameters<typeof scanAttr>[0], start: number, end: number): void {
    const id = scanAttr(attrs, "Id");
    const type = scanAttr(attrs, "Type");
    const target = scanAttr(attrs, "Target");
    if (id !== undefined && type !== undefined && target !== undefined) {
      entries.push({ id, type, target, external: scanAttr(attrs, "TargetMode") === "External", start, end });
    }
  }
}

/** The part a `.rels` part describes ("" for the package's `_rels/.rels`),
 * or undefined when the name is not a relationships part. */
export function relsOwner(relsName: string): string | undefined {
  const match = /^(.*?)_rels\/([^/]*)\.rels$/.exec(relsName);
  if (!match) {
    return undefined;
  }
  return `${match[1]}${match[2]}`;
}

/** The relationships of `part` ("" = the package), resolved. */
export function relationshipsOf(parts: PartLookup, part: string): RelEntry[] {
  const xml = partTextOf(parts, part === "" ? "_rels/.rels" : relsPathFor(part));
  return xml === undefined ? [] : readRels(xml);
}

/** An internal relationship's target part path, resolved against `owner`. */
export function targetPart(owner: string, rel: RelEntry): string | undefined {
  return rel.external ? undefined : resolvePartPath(owner, rel.target);
}

/** A `<Relationship>` element. */
export function relationshipXml(id: string, type: string, target: string): string {
  return `<Relationship Id="${escapeAttributeValue(id)}" Type="${escapeAttributeValue(type)}" Target="${escapeAttributeValue(target)}"/>`;
}

/** The next `rId<n>` not in `taken`. */
export function nextRelationshipId(taken: Iterable<string>): string {
  let max = 0;
  for (const id of taken) {
    const match = /^rId(\d+)$/.exec(id);
    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  }
  return `rId${max + 1}`;
}

/** `xml` (a `.rels` part) with the relationships whose ids are in `remove`
 * spliced out and `append` added before the closing tag; every other byte
 * verbatim (so `TargetMode` survives). */
export function spliceRels(xml: string, remove: ReadonlySet<string>, append: readonly string[]): string {
  const splices: Splice[] = readRels(xml)
    .filter((rel) => remove.has(rel.id))
    .map((rel) => ({ start: rel.start, end: rel.end, text: "" }));
  let out = applySplices(xml, splices);
  if (append.length > 0) {
    const close = out.lastIndexOf("</");
    out = close >= 0 ? out.slice(0, close) + append.join("") + out.slice(close) : out.replace(/\/>\s*$/, `>${append.join("")}</Relationships>`);
  }
  return out;
}

/** The relative `Target` from `fromPart` to `toPart`. */
export function relativeTarget(fromPart: string, toPart: string): string {
  const from = fromPart.split("/").slice(0, -1);
  const to = toPart.split("/");
  let common = 0;
  while (common < from.length && common < to.length - 1 && from[common] === to[common]) {
    common += 1;
  }
  return [...from.slice(common).map(() => ".."), ...to.slice(common)].join("/");
}

// --- the package graph --------------------------------------------------------------

/** Every internal relationship edge of a package: part → target parts. The
 * package itself is the part "". */
export function packageEdges(parts: PartLookup): Map<string, string[]> {
  const edges = new Map<string, string[]>();
  for (const name of parts.keys()) {
    const owner = relsOwner(name);
    if (owner === undefined) {
      continue;
    }
    const xml = partTextOf(parts, name);
    if (xml === undefined) {
      continue;
    }
    let rels: RelEntry[];
    try {
      rels = readRels(xml);
    } catch {
      continue;
    }
    edges.set(
      owner,
      rels.flatMap((rel) => {
        const target = targetPart(owner, rel);
        return target === undefined ? [] : [target];
      }),
    );
  }
  return edges;
}

/** Every part reachable from `roots` through `edges`. */
export function reachableFrom(roots: Iterable<string>, edges: ReadonlyMap<string, readonly string[]>): Set<string> {
  const seen = new Set<string>();
  const queue = [...roots];
  while (queue.length > 0) {
    const part = queue.pop()!;
    if (seen.has(part)) {
      continue;
    }
    seen.add(part);
    for (const target of edges.get(part) ?? []) {
      if (!seen.has(target)) {
        queue.push(target);
      }
    }
  }
  return seen;
}

// --- part names -----------------------------------------------------------------------

/** The next free name for a part like `like` (SPEC-121 "Duplicating a
 * sheet"): `drawing3.xml` → `drawing<max+1>.xml` over the folder's parts of
 * the same stem; a name without trailing digits gains `_2` (then `_3`, …). */
export function nextFreePartName(taken: ReadonlySet<string>, like: string): string {
  const slash = like.lastIndexOf("/");
  const folder = like.slice(0, slash + 1);
  const file = like.slice(slash + 1);
  const dot = file.lastIndexOf(".");
  const base = dot > 0 ? file.slice(0, dot) : file;
  const ext = dot > 0 ? file.slice(dot) : "";
  const digits = /^(.*?)(\d+)$/.exec(base);
  if (digits) {
    const stem = digits[1]!;
    const pattern = new RegExp(`^${escapeRegExp(folder + stem)}(\\d+)${escapeRegExp(ext)}$`);
    let max = 0;
    for (const name of taken) {
      const match = pattern.exec(name);
      if (match) {
        max = Math.max(max, Number(match[1]));
      }
    }
    return `${folder}${stem}${max + 1}${ext}`;
  }
  for (let n = 2; ; n += 1) {
    const candidate = `${folder}${base}_${n}${ext}`;
    if (!taken.has(candidate)) {
      return candidate;
    }
  }
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// --- the source workbook ----------------------------------------------------------------

/** One `<sheet>` of the retained workbook part. */
export interface SourceSheet {
  name: string;
  sheetId: number;
  rId: string;
  state?: string;
  /** The part its relationship targets ("" when the relationship is missing). */
  partName: string;
  relType: string;
  /** The `<sheet>` element's offsets in the workbook part. */
  start: number;
  end: number;
}

export interface SourceWorkbook {
  path: string;
  xml: string;
  relsPath: string;
  relsXml: string;
  rels: RelEntry[];
  sheets: SourceSheet[];
}

const RELATIONSHIP_ID_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);

/** The workbook part's path, from `_rels/.rels`. */
export function workbookPathOf(parts: PartLookup): string | undefined {
  const rel = relationshipsOf(parts, "").find((entry) => relationshipKind(entry.type) === "officeDocument");
  return rel ? targetPart("", rel) : undefined;
}

/** The retained workbook part and its sheet list, or undefined for a
 * package without one. */
export function readSourceWorkbook(parts: PartLookup): SourceWorkbook | undefined {
  const path = workbookPathOf(parts);
  const xml = path === undefined ? undefined : partTextOf(parts, path);
  if (path === undefined || xml === undefined) {
    return undefined;
  }
  const relsPath = relsPathFor(path);
  const relsXml = partTextOf(parts, relsPath) ?? "";
  const rels = relsXml === "" ? [] : readRels(relsXml);
  const byId = new Map(rels.map((rel) => [rel.id, rel] as const));
  const sheets: SourceSheet[] = [];
  let depth = 0;
  let sheetsDepth = -1;
  let pending: { attrs: Parameters<typeof scanAttr>[0]; start: number } | null = null;
  const push = (attrs: Parameters<typeof scanAttr>[0], start: number, end: number): void => {
    const rIdAttr = attrs.find((attr) => attr.local === "id" && attr.ns !== undefined && RELATIONSHIP_ID_NAMESPACES.has(attr.ns));
    const rId = rIdAttr?.value ?? "";
    const rel = byId.get(rId);
    const sheet: SourceSheet = {
      name: scanAttr(attrs, "name") ?? "",
      sheetId: Number(scanAttr(attrs, "sheetId") ?? 0),
      rId,
      partName: rel ? (targetPart(path, rel) ?? "") : "",
      relType: rel?.type ?? "",
      start,
      end,
    };
    const state = scanAttr(attrs, "state");
    if (state !== undefined) {
      sheet.state = state;
    }
    sheets.push(sheet);
  };
  scanXml(xml, (event) => {
    if (event.type === "open") {
      depth += 1;
      if (event.local === "sheets" && depth === 2) {
        sheetsDepth = depth;
      } else if (event.local === "sheet" && depth === sheetsDepth + 1 && sheetsDepth > 0) {
        if (event.selfClosing) {
          push(event.attrs, event.start, event.end);
        } else {
          pending = { attrs: event.attrs, start: event.start };
        }
      }
      if (event.selfClosing) {
        depth -= 1;
      }
    } else if (event.type === "close") {
      if (event.local === "sheet" && pending && depth === sheetsDepth + 1) {
        push(pending.attrs, pending.start, event.end);
        pending = null;
      }
      if (depth === sheetsDepth) {
        sheetsDepth = -1;
      }
      depth -= 1;
    }
  });
  return { path, xml, relsPath, relsXml, rels, sheets };
}

// --- content types ----------------------------------------------------------------------

/** One `<Override>` of `[Content_Types].xml`. */
export interface ContentTypeOverride {
  partName: string;
  contentType: string;
  start: number;
  end: number;
}

export interface ContentTypes {
  overrides: ContentTypeOverride[];
  /** lower-case extension → content type */
  defaults: Map<string, string>;
}

export function readContentTypes(xml: string): ContentTypes {
  const overrides: ContentTypeOverride[] = [];
  const defaults = new Map<string, string>();
  scanXml(xml, (event) => {
    if (event.type !== "open") {
      return;
    }
    if (event.local === "Override") {
      const partName = scanAttr(event.attrs, "PartName");
      const contentType = scanAttr(event.attrs, "ContentType");
      if (partName !== undefined && contentType !== undefined) {
        overrides.push({ partName: partName.replace(/^\//, ""), contentType, start: event.start, end: event.end });
      }
    } else if (event.local === "Default") {
      const extension = scanAttr(event.attrs, "Extension");
      const contentType = scanAttr(event.attrs, "ContentType");
      if (extension !== undefined && contentType !== undefined) {
        defaults.set(extension.toLowerCase(), contentType);
      }
    }
  });
  return { overrides, defaults };
}

/** The content type a part has (an Override, else its extension's Default). */
export function contentTypeIn(types: ContentTypes, partName: string): string | undefined {
  const override = types.overrides.find((entry) => entry.partName === partName);
  if (override) {
    return override.contentType;
  }
  const dot = partName.lastIndexOf(".");
  return dot >= 0 ? types.defaults.get(partName.slice(dot + 1).toLowerCase()) : undefined;
}

/** `[Content_Types].xml` with the Overrides for `remove` spliced out and
 * `add` appended before `</Types>`; everything else verbatim. */
export function spliceContentTypes(
  xml: string,
  remove: ReadonlySet<string>,
  add: readonly { partName: string; contentType: string }[],
): string {
  const types = readContentTypes(xml);
  const splices: Splice[] = types.overrides
    .filter((entry) => remove.has(entry.partName))
    .map((entry) => ({ start: entry.start, end: entry.end, text: "" }));
  let out = applySplices(xml, splices);
  if (add.length > 0) {
    const elements = add
      .map((entry) => `<Override PartName="/${escapeAttributeValue(entry.partName)}" ContentType="${escapeAttributeValue(entry.contentType)}"/>`)
      .join("");
    const close = out.lastIndexOf("</");
    out = out.slice(0, close) + elements + out.slice(close);
  }
  return out;
}
