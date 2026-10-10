// Duplicating an xlsx sheet's retained source (SPEC-121 "Duplicating a
// sheet"), at command time and copy-on-write: the copy gets its own part —
// the origin's bytes under the next free `xl/worksheets/sheet<N>.xml` — and
// its own relationship closure, so a later command-time rewrite of one
// sheet's retained text (SPEC-124) never reaches the other's.
//
// - Each part reachable ONLY through the origin's part is copied under its
//   folder's next free number (`drawing3.xml` → `drawing<max+1>.xml`; a name
//   with no trailing digits gains `_2`), with a copied `.rels` pointing at
//   the copies. Parts also reachable from elsewhere (images, pivot caches,
//   embedded packages, persons) are shared by reference.
// - Fix-ups, so Excel never answers with its repair prompt: a copied table
//   gets `id` = max table id + 1 and `name`/`displayName` `<name>_<n>`
//   (smallest free n ≥ 2 among tables and defined names), structured
//   references to it in the copy renamed; a copied VML part gets `o:idmap
//   data` = max over all VML parts + 1 and its `_x0000_s<id>` shape ids
//   rebased by 1024 per block; threaded-comment ids and the matching legacy
//   `xr:uid`s are re-minted deterministically (a v4-shaped UUID from FNV-1a
//   of the copy's sheet id + the original id), so writes are reproducible; a
//   copied pivot table's `name` is uniqued.
// - The Content-Types Overrides, the workbook relationship and the sheetId
//   are the writer's, added on save (xlsx/writer.ts).
// Deleting a copy removes what duplicating added (removeOwnedParts); an
// original's parts stay in `opaque` and the writer skips them on save.
import { bytesToBase64, type OpaquePart } from "../../core/packageParts";
import { relsPathFor } from "../../core/xml";
import { scanXml } from "../../core/xmlScanner";
import { joinTokens, tokenize } from "../formula/tokens";
import { applySplices, attributeValue, findAttribute, startTagAt, type Splice } from "../xmlSplice";
import {
  decodeUtf8,
  encodeUtf8,
  LazyParts,
  nextFreePartName,
  packageEdges,
  partTextOf,
  readRels,
  relationshipKind,
  relativeTarget,
  targetPart,
  workbookPathOf,
  type PartLookup,
} from "./package";

/** The worksheet-level relationship kinds a duplicate can carry. */
export const SAFE_SHEET_RELATIONSHIPS: ReadonlySet<string> = new Set([
  "drawing",
  "comments",
  "vmlDrawing",
  "threadedComment",
  "table",
  "pivotTable",
  "printerSettings",
  "hyperlink",
]);

/** What a copied drawing (charts, chart styles/colours, images) and a
 * copied chart may relate to. */
const SAFE_NESTED_RELATIONSHIPS: Readonly<Record<string, ReadonlySet<string>>> = {
  drawing: new Set(["chart", "image", "hyperlink"]),
  chart: new Set(["chartStyle", "chartColorStyle", "image", "package", "themeOverride"]),
};

/** Whether a part's relationships are all on `safe` (a hyperlink only as
 * an external link), checking a drawing's and a chart's own in turn. */
function onlySafe(parts: PartLookup, partName: string, safe: ReadonlySet<string>): boolean {
  const xml = partTextOf(parts, relsPathFor(partName));
  if (xml === undefined) {
    return true;
  }
  return readRels(xml).every((rel) => {
    const kind = relationshipKind(rel.type);
    if (!safe.has(kind) || (kind === "hyperlink" && !rel.external)) {
      return false;
    }
    const nested = SAFE_NESTED_RELATIONSHIPS[kind];
    const target = targetPart(partName, rel);
    return nested === undefined || target === undefined || onlySafe(parts, target, nested);
  });
}

/** Whether every relationship of a worksheet part is on the safe list
 * (`hyperlink` only as an external link; a drawing holding only charts,
 * their styles and colours, and images). */
export function hasOnlySafeRelationships(parts: PartLookup, partName: string): boolean {
  return onlySafe(parts, partName, SAFE_SHEET_RELATIONSHIPS);
}

/** The parts reachable from `start` that nothing outside that set refers
 * to — what a copy of `start` must copy, or what deleting a copy may drop. */
export function ownedClosure(edges: ReadonlyMap<string, readonly string[]>, start: string): string[] {
  const referrers = new Map<string, string[]>();
  for (const [from, targets] of edges) {
    for (const target of targets) {
      referrers.set(target, [...(referrers.get(target) ?? []), from]);
    }
  }
  const order: string[] = [start];
  const seen = new Set([start]);
  for (let i = 0; i < order.length; i += 1) {
    for (const target of edges.get(order[i]!) ?? []) {
      if (!seen.has(target)) {
        seen.add(target);
        order.push(target);
      }
    }
  }
  const owned = new Set([start]);
  let grew = true;
  while (grew) {
    grew = false;
    for (const part of order) {
      if (!owned.has(part) && (referrers.get(part) ?? []).every((from) => owned.has(from))) {
        owned.add(part);
        grew = true;
      }
    }
  }
  return order.filter((part) => owned.has(part));
}

// --- deterministic ids ----------------------------------------------------------------------

function fnv1a(text: string, seed: number): number {
  let hash = seed >>> 0;
  for (let i = 0; i < text.length; i += 1) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return hash >>> 0;
}

/** A v4-shaped UUID minted from FNV-1a of the copy's sheet id and an
 * original id — the same inputs always mint the same id. Braced and
 * upper-cased like the original when it was. */
export function mintedId(copySheetId: string, originalId: string): string {
  const input = `${copySheetId}\u0000${originalId}`;
  const words = [0x811c9dc5, 0x01000193, 0x2545f491, 0x9e3779b9].map((seed, i) => fnv1a(`${i}:${input}`, seed));
  const hex = words.map((word) => word.toString(16).padStart(8, "0")).join("");
  const variant = ((Number.parseInt(hex[16]!, 16) & 0x3) | 0x8).toString(16);
  const uuid = `${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-${variant}${hex.slice(17, 20)}-${hex.slice(20, 32)}`;
  const braced = originalId.startsWith("{");
  const upper = originalId.toUpperCase() === originalId;
  const text = upper ? uuid.toUpperCase() : uuid;
  return braced ? `{${text}}` : text;
}

// --- the copy ------------------------------------------------------------------------------

export interface XlsxSheetCopy {
  /** `opaque` with the copied parts appended (a new array). */
  opaque: OpaquePart[];
  /** The copy's worksheet part. */
  partName: string;
  /** Table display names renamed in the copy: OLD (upper-cased) → new. */
  tableRenames: Map<string, string>;
}

/** Every element of a part whose local name is `local`, with its start
 * tag's offsets. */
function startTags(xml: string, local: string): { start: number; tag: string }[] {
  const found: { start: number; tag: string }[] = [];
  scanXml(xml, (event) => {
    if (event.type === "open" && event.local === local) {
      found.push({ start: event.start, tag: xml.slice(event.start, event.end) });
    }
  });
  return found;
}

/** The text ranges of every `local` element (between its tags). */
function elementTexts(xml: string, local: string): { start: number; end: number }[] {
  const ranges: { start: number; end: number }[] = [];
  let open = -1;
  scanXml(xml, (event) => {
    if (event.type === "open" && event.local === local && !event.selfClosing) {
      open = event.end;
    } else if (event.type === "close" && event.local === local && open >= 0) {
      ranges.push({ start: open, end: event.start });
      open = -1;
    }
  });
  return ranges;
}

/** A tag with one attribute's raw value replaced (the attribute must exist). */
function withRawAttribute(tag: string, name: string, raw: string): string {
  const span = findAttribute(tag, name);
  return span ? tag.slice(0, span.valueStart) + raw + tag.slice(span.valueEnd) : tag;
}

/** Structured references to a renamed table, renamed in formula text (wire
 * or canonical — the tokenizer reads both). */
export function renameTablesInFormula(text: string, renames: ReadonlyMap<string, string>): string {
  if (renames.size === 0) {
    return text;
  }
  const tokens = tokenize(text);
  let changed = false;
  for (const token of tokens) {
    if (token.kind !== "structured") {
      continue;
    }
    const bracket = token.text.indexOf("[");
    const name = bracket > 0 ? token.text.slice(0, bracket) : "";
    const next = name ? renames.get(name.toUpperCase()) : undefined;
    if (next !== undefined) {
      token.text = next + token.text.slice(bracket);
      changed = true;
    }
  }
  return changed ? joinTokens(tokens) : text;
}

function renameTablesInPart(xml: string, renames: ReadonlyMap<string, string>, elements: readonly string[]): string {
  if (renames.size === 0) {
    return xml;
  }
  const splices: Splice[] = [];
  for (const local of elements) {
    for (const range of elementTexts(xml, local)) {
      const raw = xml.slice(range.start, range.end);
      const next = renameTablesInFormula(raw, renames);
      if (next !== raw) {
        splices.push({ ...range, text: next });
      }
    }
  }
  return applySplices(xml, splices);
}

/** Every name the workbook already uses for a table or a defined name
 * (upper-cased), and the largest table id. */
function tableNamespace(parts: PartLookup): { names: Set<string>; maxId: number } {
  const names = new Set<string>();
  let maxId = 0;
  for (const name of parts.keys()) {
    if (!/^xl\/tables\/[^/]+\.xml$/.test(name)) {
      continue;
    }
    const xml = partTextOf(parts, name) ?? "";
    const root = startTags(xml, "table")[0];
    if (!root) {
      continue;
    }
    maxId = Math.max(maxId, Number(attributeValue(root.tag, "id") ?? 0) || 0);
    for (const attr of ["name", "displayName"]) {
      const value = attributeValue(root.tag, attr);
      if (value !== undefined) {
        names.add(value.toUpperCase());
      }
    }
  }
  const workbookPath = workbookPathOf(parts);
  const workbook = workbookPath === undefined ? undefined : partTextOf(parts, workbookPath);
  for (const element of workbook === undefined ? [] : startTags(workbook, "definedName")) {
    const value = attributeValue(element.tag, "name");
    if (value !== undefined) {
      names.add(value.toUpperCase());
    }
  }
  return { names, maxId };
}

function uniqueName(base: string, taken: Set<string>): string {
  for (let n = 2; ; n += 1) {
    const candidate = `${base}_${n}`;
    if (!taken.has(candidate.toUpperCase())) {
      taken.add(candidate.toUpperCase());
      return candidate;
    }
  }
}

/** The largest `o:idmap data` block over the package's VML parts. */
function maxVmlBlock(parts: PartLookup): number {
  let max = 0;
  for (const name of parts.keys()) {
    if (!name.toLowerCase().endsWith(".vml")) {
      continue;
    }
    for (const match of (partTextOf(parts, name) ?? "").matchAll(/<o:idmap\b[^>]*\bdata="([^"]*)"/g)) {
      for (const value of match[1]!.split(",")) {
        max = Math.max(max, Number(value) || 0);
      }
    }
  }
  return max;
}

/** A copied VML part moved to shape-id block `block`. */
function rebaseVml(xml: string, block: number): string {
  const match = /<o:idmap\b[^>]*\bdata="([^"]*)"/.exec(xml);
  const old = Number(match?.[1]?.split(",")[0] ?? 0) || 0;
  let out = xml.replace(/(<o:idmap\b[^>]*\bdata=")([^"]*)"/, `$1${block}"`);
  out = out.replace(/_x0000_s(\d+)/g, (_whole, id: string) => `_x0000_s${Number(id) - 1024 * old + 1024 * block}`);
  return out;
}

/** A copied threaded-comments part with its ids re-minted; `ids` collects
 * the old → new mapping for the legacy comments' `xr:uid`s. */
function remintThreadedComments(xml: string, copySheetId: string, ids: Map<string, string>): string {
  const splices: Splice[] = [];
  for (const element of startTags(xml, "threadedComment")) {
    let tag = element.tag;
    for (const attr of ["id", "parentId"]) {
      const value = attributeValue(tag, attr);
      if (value !== undefined) {
        const minted = ids.get(value) ?? mintedId(copySheetId, value);
        ids.set(value, minted);
        tag = withRawAttribute(tag, attr, minted);
      }
    }
    if (tag !== element.tag) {
      splices.push({ start: element.start, end: element.start + element.tag.length, text: tag });
    }
  }
  return applySplices(xml, splices);
}

/** A copied legacy comments part with each `xr:uid` re-mapped. */
function remapCommentUids(xml: string, ids: ReadonlyMap<string, string>): string {
  const splices: Splice[] = [];
  for (const element of startTags(xml, "comment")) {
    const uid = attributeValue(element.tag, "xr:uid");
    const minted = uid === undefined ? undefined : ids.get(uid);
    if (minted !== undefined) {
      splices.push({ start: element.start, end: element.start + element.tag.length, text: withRawAttribute(element.tag, "xr:uid", minted) });
    }
  }
  return applySplices(xml, splices);
}

/** Pivot table names already used in the package (upper-cased). */
function pivotNames(parts: PartLookup): Set<string> {
  const names = new Set<string>();
  for (const name of parts.keys()) {
    if (/^xl\/pivotTables\/[^/]+\.xml$/.test(name)) {
      const root = startTags(partTextOf(parts, name) ?? "", "pivotTableDefinition")[0];
      const value = root ? attributeValue(root.tag, "name") : undefined;
      if (value !== undefined) {
        names.add(value.toUpperCase());
      }
    }
  }
  return names;
}

/** A root element's start tag rewritten by `edit`. */
function editRoot(xml: string, local: string, edit: (tag: string) => string): string {
  const root = startTags(xml, local)[0];
  if (!root) {
    return xml;
  }
  const next = edit(root.tag);
  return next === root.tag ? xml : xml.slice(0, root.start) + next + xml.slice(root.start + root.tag.length);
}

/** Copies the origin's worksheet part and its closure into `opaque` (see
 * the module header). `copySheetId` seeds the re-minted ids. */
export function copyXlsxSheetParts(opaque: readonly OpaquePart[], originPart: string, copySheetId: string): XlsxSheetCopy {
  const parts = new LazyParts(opaque);
  const edges = packageEdges(parts);
  const owned = ownedClosure(edges, originPart);
  const taken = new Set(parts.keys());
  const renamed = new Map<string, string>();
  for (const part of owned) {
    const name = part === originPart ? nextFreePartName(taken, "xl/worksheets/sheet1.xml") : nextFreePartName(taken, part);
    taken.add(name);
    renamed.set(part, name);
  }

  const tables = tableNamespace(parts);
  const tableRenames = new Map<string, string>();
  let vmlBlock = maxVmlBlock(parts);
  const pivots = pivotNames(parts);
  const threadIds = new Map<string, string>();
  const texts = new Map<string, string>();
  // Pass 1: the parts whose fix-ups others depend on (tables, threaded
  // comments), then everything else.
  const ordered = owned.toSorted((a, b) => rank(a) - rank(b));
  for (const part of ordered) {
    const bytes = parts.get(part);
    if (!bytes) {
      continue;
    }
    if (/^xl\/tables\/[^/]+\.xml$/.test(part)) {
      texts.set(
        part,
        editRoot(decodeUtf8(bytes), "table", (tag) => {
          let next = tag;
          tables.maxId += 1;
          if (findAttribute(next, "id")) {
            next = withRawAttribute(next, "id", String(tables.maxId));
          }
          const base = attributeValue(tag, "name") ?? attributeValue(tag, "displayName") ?? "Table";
          const name = uniqueName(base, tables.names);
          const display = attributeValue(tag, "displayName");
          if (display !== undefined) {
            tableRenames.set(display.toUpperCase(), name);
            next = withRawAttribute(next, "displayName", name);
          }
          if (findAttribute(next, "name")) {
            next = withRawAttribute(next, "name", name);
          }
          const uid = attributeValue(next, "xr:uid");
          if (uid !== undefined) {
            next = withRawAttribute(next, "xr:uid", mintedId(copySheetId, uid));
          }
          return next;
        }),
      );
    } else if (part.toLowerCase().endsWith(".vml")) {
      vmlBlock += 1;
      texts.set(part, rebaseVml(decodeUtf8(bytes), vmlBlock));
    } else if (part.startsWith("xl/threadedComments/")) {
      texts.set(part, remintThreadedComments(decodeUtf8(bytes), copySheetId, threadIds));
    } else if (/^xl\/pivotTables\/[^/]+\.xml$/.test(part)) {
      texts.set(
        part,
        editRoot(decodeUtf8(bytes), "pivotTableDefinition", (tag) => {
          const name = attributeValue(tag, "name");
          return name === undefined ? tag : withRawAttribute(tag, "name", uniqueName(name, pivots));
        }),
      );
    }
  }
  for (const part of owned) {
    const bytes = parts.get(part);
    if (!bytes || texts.has(part)) {
      continue;
    }
    if (/^xl\/comments[^/]*\.xml$/.test(part) && threadIds.size > 0) {
      texts.set(part, remapCommentUids(decodeUtf8(bytes), threadIds));
    } else if (part === originPart && tableRenames.size > 0) {
      texts.set(part, renameTablesInPart(decodeUtf8(bytes), tableRenames, ["f"]));
    }
  }
  for (const [part, xml] of texts) {
    if (part.startsWith("xl/tables/")) {
      texts.set(part, renameTablesInPart(xml, tableRenames, ["calculatedColumnFormula", "totalsRowFormula"]));
    }
  }

  const added: OpaquePart[] = [];
  for (const part of owned) {
    const copyName = renamed.get(part)!;
    const text = texts.get(part);
    const bytes = text !== undefined ? encodeUtf8(text) : parts.get(part);
    if (!bytes) {
      continue;
    }
    added.push({ name: copyName, data: bytesToBase64(bytes) });
    const relsXml = partTextOf(parts, relsPathFor(part));
    if (relsXml !== undefined) {
      added.push({ name: relsPathFor(copyName), data: bytesToBase64(encodeUtf8(retargetRels(relsXml, part, copyName, renamed))) });
    }
  }
  return { opaque: [...opaque, ...added], partName: renamed.get(originPart)!, tableRenames };
}

/** Fix-up order: tables and threaded comments before the parts that
 * depend on their renames. */
function rank(part: string): number {
  if (part.startsWith("xl/tables/")) {
    return 0;
  }
  if (part.startsWith("xl/threadedComments/")) {
    return 1;
  }
  return 2;
}

/** A copied `.rels` part: every internal target the copy owns points at its
 * copy; shared targets keep pointing at the shared part. */
function retargetRels(xml: string, part: string, copyName: string, renamed: ReadonlyMap<string, string>): string {
  const splices: Splice[] = [];
  for (const rel of readRels(xml)) {
    const target = targetPart(part, rel);
    if (target === undefined) {
      continue;
    }
    const copyTarget = renamed.get(target) ?? target;
    const value = relativeTarget(copyName, copyTarget);
    const tag = startTagAt(xml, rel.start);
    const span = findAttribute(tag, "Target");
    if (span && rel.target !== value && (renamed.has(target) || copyName.slice(0, copyName.lastIndexOf("/")) !== part.slice(0, part.lastIndexOf("/")))) {
      splices.push({ start: rel.start + span.valueStart, end: rel.start + span.valueEnd, text: value });
    }
  }
  return applySplices(xml, splices);
}

/** `opaque` without the parts a deleted copy owned: its worksheet part, the
 * parts reachable only through it, and their `.rels`. */
export function removeOwnedParts(opaque: readonly OpaquePart[], copyPart: string): OpaquePart[] {
  const parts = new LazyParts(opaque);
  const owned = new Set(ownedClosure(packageEdges(parts), copyPart));
  const drop = new Set<string>();
  for (const part of owned) {
    drop.add(part);
    drop.add(relsPathFor(part));
  }
  return opaque.filter((part) => !drop.has(part.name));
}
