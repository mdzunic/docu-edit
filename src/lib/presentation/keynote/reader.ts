// The Keynote importer (SPEC-087): turns an iWork Archive package (a ZIP of
// Snappy-compressed protobuf object graphs) into the presentation model's
// parts, one-way. There is no .key writer — the imported deck has no
// retained package (`opaque` stays empty) and every slide is rebuilt on
// save via the SPEC-080/081 writers under `format: "pptx"`.
//
// Fidelity is bounded and declared: text placeholders, images, geometry and
// speaker notes import; shapes, tables, charts, movies, groups, Keynote's
// own run-level text styling, transitions and builds are dropped and
// counted into the import notice (AC-27/AC-36 — per-run styling is a
// documented out-of-scope, not a counted drop: the text still imports with
// role defaults, so counting each styled run would misstate what was lost).
// The format facts this decodes against are the spec's Reference, verified
// against Apple-shipped iWork data on 2026-08-24.
import { unzip, ZipError } from "../../core/zip";
import { bytesToBase64 } from "../../core/packageParts";
import { CONTENT_TYPE_BY_EXTENSION } from "../../core/media";
import type { MediaPart } from "../../core/media";
import type { ParsedPresentation, PlaceholderRole, Slide, SlideFrame, SlideParagraph } from "../model";
import { KeynoteOpenError } from "./errors";
import { decodeIwa, parseArchives, type ArchiveMessage } from "./iwa";
import { fields, float32Of, WIRE_FIXED32, WIRE_LENGTH_DELIMITED, WIRE_VARINT, type ProtoField } from "./protobuf";

export { KeynoteOpenError };

// The three user-facing rejection strings (AC-8..AC-11) — every failure
// path in this module throws one of them as a KeynoteOpenError (AC-12).
export const KEYNOTE_LEGACY_MESSAGE =
  "This is a Keynote '09 file — open it in Keynote and save it again, or export it as .pptx, then try that file";
export const NOT_KEYNOTE_MESSAGE = "This iWork file isn't a Keynote presentation we can open";
export const KEYNOTE_UNREADABLE_MESSAGE =
  "This Keynote file can't be opened — export it as .pptx from Keynote and try that file";

// Verified archive type numbers (spec Reference, fact #6).
const TYPE_DOCUMENT = 1;
const TYPE_SHOW = 2;
const TYPE_SLIDE_NODE = 4;
const TYPE_SLIDE = 5;
const TYPE_TEXT_STORAGE = 2001;
const TYPE_PACKAGE_METADATA = 11006;

/** How deep the breadth-first sub-message searches look for a text-storage
 * or data reference on a drawable's graph (spec: depth ≤ 6). */
const MESSAGE_SEARCH_DEPTH = 6;

/** Image extensions that import. tiff/heic (and anything else) are dropped
 * and counted — browsers can't decode them portably (spec Reference). */
const IMPORTABLE_IMAGE_EXTENSIONS = new Set(["png", "jpg", "jpeg", "gif"]);

/** What `parseKeynote` counts as lost: whole drawables it could not import
 * (`objects`) and slide transitions (`effects`). */
export interface KeynoteDropCounts {
  objects: number;
  effects: number;
}

/** `parseKeynote`'s result: the shared presentation-assembly input plus the
 * drop counts the import notice reports. `opaque` is always `[]` (AC-14) and
 * no slide carries `partName`/`sourceXml` (AC-15) — the deck saves as a
 * from-scratch `.pptx`. */
export interface ParsedKeynote extends ParsedPresentation {
  dropped: KeynoteDropCounts;
}

/** Whether an already-unzipped package is an iWork Archive: it holds
 * `Index/Document.iwa`, or an `Index.zip` (the older nesting — its entries
 * are checked when parsing) (AC-5/AC-6). Content-based; the extension never
 * decides. */
export function isIwaPackage(parts: Map<string, Uint8Array>): boolean {
  return parts.has("Index/Document.iwa") || parts.has("Index.zip");
}

/** Whether a ZIP is an iWork '05–'09 file (`index.apxl`/`index.apxl.gz`),
 * which uses a different, unsupported format (AC-8). */
export function isLegacyKeynote(parts: Map<string, Uint8Array>): boolean {
  return parts.has("index.apxl") || parts.has("index.apxl.gz");
}

/** The non-blocking import notice for a finished import, or null when
 * nothing was dropped (AC-37/AC-38). Singular/zero forms handled: a zero
 * count's clause is omitted entirely. */
export function keynoteImportNotice(dropped: KeynoteDropCounts): string | null {
  if (dropped.objects === 0 && dropped.effects === 0) {
    return null;
  }
  const clause = (count: number, noun: string): string => `${count} ${noun}${count === 1 ? "" : "s"}`;
  const clauses: string[] = [];
  if (dropped.objects > 0) {
    clauses.push(clause(dropped.objects, "object"));
  }
  if (dropped.effects > 0) {
    clauses.push(clause(dropped.effects, "effect"));
  }
  return `Imported from Keynote. ${clauses.join(" and ")} couldn't be imported.`;
}

// ---------------------------------------------------------------------------
// Field helpers over the schema-less wire reader.

type Fields = ProtoField[];

function messageAt(messageFields: Fields, num: number): Fields | undefined {
  const field = messageFields.find((f) => f.num === num && f.wire === WIRE_LENGTH_DELIMITED);
  if (!field?.bytes) {
    return undefined;
  }
  try {
    return fields(field.bytes);
  } catch {
    return undefined;
  }
}

/** A TSP.Reference is a message with exactly one field: `{1: varint id}`
 * (fact #5). Returns the id, or undefined when `bytes` isn't one. */
function strictReferenceId(bytes: Uint8Array): number | undefined {
  let decoded: Fields;
  try {
    decoded = fields(bytes);
  } catch {
    return undefined;
  }
  if (decoded.length !== 1 || decoded[0].num !== 1 || decoded[0].wire !== WIRE_VARINT) {
    return undefined;
  }
  return decoded[0].value;
}

/** The id a reference-shaped sub-message at `num` points to. */
function referencedId(messageFields: Fields, num: number): number | undefined {
  const field = messageFields.find((f) => f.num === num && f.wire === WIRE_LENGTH_DELIMITED);
  return field?.bytes ? strictReferenceId(field.bytes) : undefined;
}

/** Every id the repeated reference-shaped sub-messages at `num` point to. */
function referencedIds(messageFields: Fields, num: number): number[] {
  const ids: number[] = [];
  for (const field of messageFields) {
    if (field.num !== num || field.wire !== WIRE_LENGTH_DELIMITED || !field.bytes) {
      continue;
    }
    const id = strictReferenceId(field.bytes);
    if (id !== undefined) {
      ids.push(id);
    }
  }
  return ids;
}

// ---------------------------------------------------------------------------
// The decoded package.

interface KeynoteObjects {
  /** Every archived object across every `.iwa` file, by identifier. */
  byId: Map<bigint, ArchiveMessage[]>;
  /** `Index/Document.iwa`'s archives in stream order — its first is the
   * package's root document archive. */
  documentArchives: Map<bigint, ArchiveMessage[]>;
  /** PackageMetadata field 4: data id → preferred file name (fact #6). */
  dataNames: Map<number, string>;
  /** PackageMetadata field 3: component root object ids, in order — the
   * documented fallback for slide ordering (fact #8, rule (b)). */
  componentRootIds: number[];
  /** The package's `Data/*` entries, name → bytes. */
  dataEntries: Map<string, Uint8Array>;
}

function firstMessageOf(archives: Map<bigint, ArchiveMessage[]> | undefined, id: bigint): ArchiveMessage | undefined {
  return archives?.get(id)?.[0];
}

function archiveOf(objects: KeynoteObjects, id: number | undefined): ArchiveMessage | undefined {
  return id === undefined ? undefined : firstMessageOf(objects.byId, BigInt(id));
}

async function collectIwaEntries(parts: Map<string, Uint8Array>): Promise<Map<string, Uint8Array>> {
  const entries = new Map<string, Uint8Array>();
  for (const [name, bytes] of parts) {
    if (name.startsWith("Index/") && name.endsWith(".iwa")) {
      entries.set(name.slice("Index/".length), bytes);
    }
  }
  if (entries.size === 0) {
    const nested = parts.get("Index.zip");
    if (nested) {
      let nestedParts: Map<string, Uint8Array>;
      try {
        nestedParts = await unzip(nested);
      } catch (error) {
        if (error instanceof ZipError) {
          throw new KeynoteOpenError(NOT_KEYNOTE_MESSAGE);
        }
        throw error;
      }
      for (const [name, bytes] of nestedParts) {
        if (name.endsWith(".iwa")) {
          const base = name.split("/").pop() ?? name;
          entries.set(base, bytes);
        }
      }
    }
  }
  if (!entries.has("Document.iwa")) {
    throw new KeynoteOpenError(NOT_KEYNOTE_MESSAGE);
  }
  return entries;
}

function decodePackage(entries: Map<string, Uint8Array>, parts: Map<string, Uint8Array>): KeynoteObjects {
  const byId = new Map<bigint, ArchiveMessage[]>();
  let documentArchives: Map<bigint, ArchiveMessage[]> | undefined;
  const dataNames = new Map<number, string>();
  const componentRootIds: number[] = [];

  for (const [name, bytes] of entries) {
    const archives = parseArchives(decodeIwa(bytes));
    if (name === "Document.iwa") {
      documentArchives = archives;
    }
    for (const [id, messages] of archives) {
      if (!byId.has(id)) {
        byId.set(id, messages);
      }
      for (const message of messages) {
        if (message.type !== TYPE_PACKAGE_METADATA) {
          continue;
        }
        // PackageMetadata: field 3 = repeated component {1: root object id},
        // field 4 = repeated data {1: data id, 3: preferred file name}.
        for (const field of fields(message.payload)) {
          if (field.wire !== WIRE_LENGTH_DELIMITED || !field.bytes) {
            continue;
          }
          if (field.num === 3) {
            const component = fields(field.bytes);
            const rootId = component.find((f) => f.num === 1 && f.wire === WIRE_VARINT)?.value;
            if (rootId !== undefined) {
              componentRootIds.push(rootId);
            }
          } else if (field.num === 4) {
            const data = fields(field.bytes);
            const dataId = data.find((f) => f.num === 1 && f.wire === WIRE_VARINT)?.value;
            const nameField = data.find((f) => f.num === 3 && f.wire === WIRE_LENGTH_DELIMITED);
            if (dataId !== undefined) {
              dataNames.set(dataId, nameField?.bytes ? new TextDecoder().decode(nameField.bytes) : "");
            }
          }
        }
      }
    }
  }

  if (!documentArchives || documentArchives.size === 0) {
    throw new KeynoteOpenError(NOT_KEYNOTE_MESSAGE);
  }

  const dataEntries = new Map<string, Uint8Array>();
  for (const [name, bytes] of parts) {
    if (name.startsWith("Data/")) {
      dataEntries.set(name, bytes);
    }
  }

  return { byId, documentArchives, dataNames, componentRootIds, dataEntries };
}

// ---------------------------------------------------------------------------
// Sub-message searches (iterative, bounded — AC-43/AC-44).

/** Breadth-first over a message's length-delimited sub-messages, depth ≤
 * MESSAGE_SEARCH_DEPTH, returning the first id for which `matches` is true.
 * Reference-shaped sub-messages are tested, not descended into — a
 * reference is a leaf pointing out of this object. */
function findReferencedId(
  root: Fields,
  matches: (id: number) => boolean,
): number | undefined {
  const queue: { messageFields: Fields; depth: number }[] = [{ messageFields: root, depth: 0 }];
  while (queue.length > 0) {
    const { messageFields, depth } = queue.shift() as { messageFields: Fields; depth: number };
    for (const field of messageFields) {
      if (field.wire !== WIRE_LENGTH_DELIMITED || !field.bytes) {
        continue;
      }
      const id = strictReferenceId(field.bytes);
      if (id !== undefined) {
        if (matches(id)) {
          return id;
        }
        continue;
      }
      if (depth + 1 < MESSAGE_SEARCH_DEPTH) {
        try {
          queue.push({ messageFields: fields(field.bytes), depth: depth + 1 });
        } catch {
          // Not a decodable sub-message (a string, raw bytes) — skip.
        }
      }
    }
  }
  return undefined;
}

/** The first text-storage reference on an object's graph, following
 * references across objects (a notes ref points at a container whose
 * placeholder chain reaches the storage). Visited-set bounded (AC-43). */
function findStorageIdAcrossObjects(objects: KeynoteObjects, startId: number): number | undefined {
  const visited = new Set<number>();
  const queue: number[] = [startId];
  while (queue.length > 0) {
    const id = queue.shift() as number;
    if (visited.has(id) || visited.size > 64) {
      continue;
    }
    visited.add(id);
    const archive = archiveOf(objects, id);
    if (!archive) {
      continue;
    }
    if (archive.type === TYPE_TEXT_STORAGE) {
      return id;
    }
    let decoded: Fields;
    try {
      decoded = fields(archive.payload);
    } catch {
      continue;
    }
    const storageId = findReferencedId(decoded, (refId) => archiveOf(objects, refId)?.type === TYPE_TEXT_STORAGE);
    if (storageId !== undefined) {
      return storageId;
    }
    // No storage on this object — follow its outgoing references one hop.
    findReferencedId(decoded, (refId) => {
      if (!visited.has(refId)) {
        queue.push(refId);
      }
      return false;
    });
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Text.

/** TSWP.StorageArchive field 3 strings, concatenated, cleaned and split
 * into paragraphs on `\n` and U+2029 (AC-24); U+FFFC and zero-width
 * characters dropped (AC-25); empty paragraphs skipped. */
function storageParagraphText(objects: KeynoteObjects, storageId: number): string[] {
  const archive = archiveOf(objects, storageId);
  if (!archive) {
    return [];
  }
  let decoded: Fields;
  try {
    decoded = fields(archive.payload);
  } catch {
    return [];
  }
  const decoder = new TextDecoder();
  let text = "";
  for (const field of decoded) {
    if (field.num === 3 && field.wire === WIRE_LENGTH_DELIMITED && field.bytes) {
      text += decoder.decode(field.bytes);
    }
  }
  const cleaned = text.replace(/[\uFFFC\u200B-\u200D\uFEFF]/g, "");
  return cleaned.split(/\r\n|[\r\n\u2029]/).filter((paragraph) => paragraph.length > 0);
}

/** Role-based default styling (AC-26/AC-28/AC-29): title → 36pt bold
 * #000000, body → 18pt with disc bullets, other → 20pt. Keynote's own
 * run-level formatting is out of scope (AC-27) — these defaults are the
 * whole of the imported text's look. */
function styledParagraphs(role: PlaceholderRole, lines: string[]): SlideParagraph[] {
  return lines.map((line) => {
    if (role === "title") {
      return { runs: [{ text: line, sizePt: 36, bold: true as const, color: "#000000" }] };
    }
    if (role === "body") {
      return { runs: [{ text: line, sizePt: 18 }], bullet: "disc" as const };
    }
    return { runs: [{ text: line, sizePt: 20 }] };
  });
}

// ---------------------------------------------------------------------------
// Geometry.

interface FrameBox {
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
}

/** Whether a decoded message is TSD.GeometryArchive-shaped: field 1 =
 * position `{1: float32 x, 2: float32 y}` (optional), field 2 = size
 * `{1: float32 w, 2: float32 h}` (required) — fact #6. */
function tryGeometry(messageFields: Fields): { x: number; y: number; w: number; h: number } | undefined {
  const size = messageAt(messageFields, 2);
  if (!size) {
    return undefined;
  }
  const wField = size.find((f) => f.num === 1 && f.wire === WIRE_FIXED32);
  const hField = size.find((f) => f.num === 2 && f.wire === WIRE_FIXED32);
  if (!wField || !hField) {
    return undefined;
  }
  const w = float32Of(wField);
  const h = float32Of(hField);
  if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) {
    return undefined;
  }
  let x = 0;
  let y = 0;
  const position = messageAt(messageFields, 1);
  if (position) {
    const xField = position.find((f) => f.num === 1 && f.wire === WIRE_FIXED32);
    const yField = position.find((f) => f.num === 2 && f.wire === WIRE_FIXED32);
    if (xField && yField) {
      x = float32Of(xField);
      y = float32Of(yField);
    }
  }
  return { x, y, w, h };
}

/** The first GeometryArchive-shaped message on the drawable's super chain —
 * the chain of field-1 sub-messages (`{1: {1: {1: GeometryArchive}}}`,
 * fact #6), walked iteratively. */
function findGeometry(drawableFields: Fields, pointToCm: number): FrameBox | undefined {
  let current: Fields | undefined = drawableFields;
  for (let depth = 0; depth < MESSAGE_SEARCH_DEPTH && current; depth++) {
    const geometry = tryGeometry(current);
    if (geometry) {
      return {
        xCm: geometry.x * pointToCm,
        yCm: geometry.y * pointToCm,
        widthCm: geometry.w * pointToCm,
        heightCm: geometry.h * pointToCm,
      };
    }
    current = messageAt(current, 1);
  }
  return undefined;
}

/** The role's default band when a drawable has no resolvable geometry
 * (spec Reference: title 8%/8%/84%/16%, body 8%/30%/84%/60%, other and
 * images a centred half-width 4:3 box). */
function defaultBand(role: PlaceholderRole | "image", slideWidthCm: number, slideHeightCm: number): FrameBox {
  if (role === "title") {
    return { xCm: slideWidthCm * 0.08, yCm: slideHeightCm * 0.08, widthCm: slideWidthCm * 0.84, heightCm: slideHeightCm * 0.16 };
  }
  if (role === "body") {
    return { xCm: slideWidthCm * 0.08, yCm: slideHeightCm * 0.3, widthCm: slideWidthCm * 0.84, heightCm: slideHeightCm * 0.6 };
  }
  const widthCm = slideWidthCm * 0.5;
  const heightCm = widthCm * 0.75;
  return {
    xCm: (slideWidthCm - widthCm) / 2,
    yCm: (slideHeightCm - heightCm) / 2,
    widthCm,
    heightCm,
  };
}

// ---------------------------------------------------------------------------
// Media.

interface MediaResolver {
  /** data id → minted media id, so an image used twice shares one part. */
  minted: Map<number, string>;
  parts: MediaPart[];
  nextIndex: number;
}

function dataExtension(name: string): string {
  const dot = name.lastIndexOf(".");
  return dot < 0 ? "" : name.slice(dot + 1).toLowerCase();
}

/** The `Data/` entry for a data id: its base name ends `-N` before the
 * extension (fact #7), falling back to the metadata's preferred file name. */
function dataEntryFor(objects: KeynoteObjects, dataId: number): { name: string; bytes: Uint8Array } | undefined {
  const suffix = `-${dataId}`;
  for (const [name, bytes] of objects.dataEntries) {
    const base = name.slice(name.lastIndexOf("/") + 1);
    const dot = base.lastIndexOf(".");
    const stem = dot < 0 ? base : base.slice(0, dot);
    if (stem.endsWith(suffix)) {
      return { name: base, bytes };
    }
  }
  const preferred = objects.dataNames.get(dataId);
  if (preferred) {
    for (const [name, bytes] of objects.dataEntries) {
      if (name.slice(name.lastIndexOf("/") + 1) === preferred) {
        return { name: preferred, bytes };
      }
    }
  }
  return undefined;
}

/** Mints the media part for a data id per SPEC-082's naming rule
 * (`ppt/media/image{n}.{ext}` — the deck imports as pptx), reusing the part
 * when the same data id appears again (AC-33/AC-34). */
function mintMediaPart(resolver: MediaResolver, dataId: number, extension: string, bytes: Uint8Array): string {
  const existing = resolver.minted.get(dataId);
  if (existing) {
    return existing;
  }
  const mediaId = `ppt/media/image${resolver.nextIndex}.${extension}`;
  resolver.nextIndex += 1;
  resolver.minted.set(dataId, mediaId);
  resolver.parts.push({
    mediaId,
    contentType: CONTENT_TYPE_BY_EXTENSION[extension] ?? "image/png",
    data: bytesToBase64(bytes),
  });
  return mediaId;
}

// ---------------------------------------------------------------------------
// Slides.

/** The root node's child slide-node ids (fact #8, resolution rule (a)): the
 * lowest-numbered repeated field whose entries are all reference-shaped and
 * whose targets are all type-4 slide nodes. */
function childNodeIds(objects: KeynoteObjects, nodeFields: Fields): number[] {
  const byNum = new Map<number, Uint8Array[]>();
  for (const field of nodeFields) {
    if (field.wire !== WIRE_LENGTH_DELIMITED || !field.bytes) {
      continue;
    }
    const list = byNum.get(field.num) ?? [];
    list.push(field.bytes);
    byNum.set(field.num, list);
  }
  const nums = [...byNum.keys()].sort((a, b) => a - b);
  for (const num of nums) {
    const entries = byNum.get(num) ?? [];
    const ids: number[] = [];
    let allNodes = true;
    for (const bytes of entries) {
      const id = strictReferenceId(bytes);
      if (id === undefined || archiveOf(objects, id)?.type !== TYPE_SLIDE_NODE) {
        allNodes = false;
        break;
      }
      ids.push(id);
    }
    if (allNodes && ids.length > 0) {
      return ids;
    }
  }
  return [];
}

/** Every slide-node id under the root, in slide order — an explicit stack
 * with a visited set, so cycles terminate and depth never recurses
 * (AC-43/AC-44). Falls back to PackageMetadata component order (rule (b))
 * when the tree yields nothing. */
function orderedSlideNodeIds(objects: KeynoteObjects, rootNodeIds: number[]): number[] {
  // KN.SlideTreeArchive field 2 is the REPEATED, ordered list of top-level
  // slide nodes (a real Keynote 14 file has one per slide, each referencing
  // its slide at field 2); a node's own children (grouped/nested slides)
  // hang below it. The nodes themselves are slides — walk from them, not
  // from their children.
  const visited = new Set<number>();
  const ordered: number[] = [];
  const stack: number[] = [];
  for (let i = rootNodeIds.length - 1; i >= 0; i--) {
    stack.push(rootNodeIds[i]);
  }
  while (stack.length > 0) {
    const id = stack.pop() as number;
    if (visited.has(id)) {
      continue;
    }
    visited.add(id);
    const archive = archiveOf(objects, id);
    if (!archive || archive.type !== TYPE_SLIDE_NODE) {
      continue;
    }
    ordered.push(id);
    let decoded: Fields;
    try {
      decoded = fields(archive.payload);
    } catch {
      continue;
    }
    const children = childNodeIds(objects, decoded);
    for (let i = children.length - 1; i >= 0; i--) {
      if (!visited.has(children[i])) {
        stack.push(children[i]);
      }
    }
  }
  if (ordered.length > 0) {
    return ordered;
  }
  // Rule (b): the type-4 archives in PackageMetadata component order.
  return objects.componentRootIds.filter(
    (id) => !rootNodeIds.includes(id) && archiveOf(objects, id)?.type === TYPE_SLIDE_NODE,
  );
}

interface SlideBuildContext {
  objects: KeynoteObjects;
  pointToCm: number;
  slideWidthCm: number;
  slideHeightCm: number;
  resolver: MediaResolver;
  dropped: KeynoteDropCounts;
}

/** One drawable → a TextFrame, an ImageFrame, or a counted drop (spec
 * Import mapping). A drawable whose graph reaches a text storage is text;
 * else one whose strict data reference resolves through the metadata index
 * to an importable image is an image; anything else is dropped and counted
 * (AC-32..AC-36). An empty text storage imports as nothing, uncounted — an
 * empty placeholder isn't lost content. */
function buildFrame(context: SlideBuildContext, drawableId: number, role: PlaceholderRole): SlideFrame | null {
  const archive = archiveOf(context.objects, drawableId);
  if (!archive) {
    context.dropped.objects += 1;
    return null;
  }
  let decoded: Fields;
  try {
    decoded = fields(archive.payload);
  } catch {
    context.dropped.objects += 1;
    return null;
  }

  const box = (kind: PlaceholderRole | "image"): FrameBox =>
    findGeometry(decoded, context.pointToCm) ?? defaultBand(kind, context.slideWidthCm, context.slideHeightCm);

  const storageId = findReferencedId(
    decoded,
    (id) => archiveOf(context.objects, id)?.type === TYPE_TEXT_STORAGE,
  );
  if (storageId !== undefined) {
    const lines = storageParagraphText(context.objects, storageId);
    if (lines.length === 0) {
      return null;
    }
    return {
      id: crypto.randomUUID(),
      kind: "text",
      role,
      ...box(role),
      paragraphs: styledParagraphs(role, lines),
    };
  }

  const dataId = findReferencedId(decoded, (id) => context.objects.dataNames.has(id));
  if (dataId !== undefined) {
    const entry = dataEntryFor(context.objects, dataId);
    const extension = entry ? dataExtension(entry.name) : "";
    if (!entry || !IMPORTABLE_IMAGE_EXTENSIONS.has(extension)) {
      // Unresolvable data, or a recognised-but-unportable (tiff/heic) or
      // non-image payload — dropped and counted (AC-35, spec Reference).
      context.dropped.objects += 1;
      return null;
    }
    return {
      id: crypto.randomUUID(),
      kind: "image",
      mediaId: mintMediaPart(context.resolver, dataId, extension, entry.bytes),
      ...box("image"),
    };
  }

  context.dropped.objects += 1;
  return null;
}

/** One KN.SlideArchive → a model Slide (fact #6: field 5 title ref, field 6
 * body ref, field 7 repeated drawable refs in z-order, field 27 notes ref,
 * field 4 transition info). */
function buildSlide(context: SlideBuildContext, slideFields: Fields): Slide {
  const titleId = referencedId(slideFields, 5);
  const bodyId = referencedId(slideFields, 6);
  const drawableIds = referencedIds(slideFields, 7);

  // z-order from field 7; the title/body placeholders join at the back when
  // field 7 doesn't already list them.
  const ordered: number[] = [];
  for (const id of [titleId, bodyId]) {
    if (id !== undefined && !drawableIds.includes(id)) {
      ordered.push(id);
    }
  }
  ordered.push(...drawableIds);

  const frames: SlideFrame[] = [];
  const seen = new Set<number>();
  for (const id of ordered) {
    if (seen.has(id)) {
      continue;
    }
    seen.add(id);
    const role: PlaceholderRole = id === titleId ? "title" : id === bodyId ? "body" : "other";
    const frame = buildFrame(context, id, role);
    if (frame) {
      frames.push(frame);
    }
  }

  const slide: Slide = { id: crypto.randomUUID(), frames };

  // Notes (AC-30/AC-31): slide field 27's target (or its placeholder chain)
  // reaching a storage imports as Slide.notes; absent or empty → no field.
  const notesId = referencedId(slideFields, 27);
  if (notesId !== undefined) {
    const storageId = findStorageIdAcrossObjects(context.objects, notesId);
    if (storageId !== undefined) {
      const notes = storageParagraphText(context.objects, storageId);
      if (notes.length > 0) {
        slide.notes = notes;
      }
    }
  }

  // Transition info (field 4) is out of scope — dropped and counted as an
  // effect when the slide carries any (AC-36).
  const transition = slideFields.find((f) => f.num === 4 && f.wire === WIRE_LENGTH_DELIMITED);
  if (transition?.bytes && transition.bytes.length > 0) {
    context.dropped.effects += 1;
  }

  return slide;
}

// ---------------------------------------------------------------------------
// The importer.

/** Keynote points → model centimetres (spec Import mapping): 16:9 decks
 * target 33.867cm wide, 4:3 decks 25.4cm, anything else 33.867cm. */
function targetWidthCm(showWidthPt: number, showHeightPt: number): number {
  const aspect = showWidthPt / showHeightPt;
  if (Math.abs(aspect / (4 / 3) - 1) <= 0.02) {
    return 25.4;
  }
  return 33.867;
}

/** Parses an already-unzipped iWork Archive package into the presentation
 * model's parts (AC-13..AC-36). Establishing "this is a Keynote deck" —
 * decoding the object files, finding the KN.DocumentArchive root, its show
 * and slide tree — fails with NOT_KEYNOTE_MESSAGE (AC-9/AC-10); failures
 * after that point (a deck resolving to zero slides) fail with
 * KEYNOTE_UNREADABLE_MESSAGE (AC-11). Everything thrown is a
 * KeynoteOpenError (AC-12). */
export async function parseKeynote(parts: Map<string, Uint8Array>): Promise<ParsedKeynote> {
  let objects: KeynoteObjects;
  let showFields: Fields;
  let rootNodeIds: number[];
  try {
    const entries = await collectIwaEntries(parts);
    objects = decodePackage(entries, parts);

    // The root document archive is Document.iwa's first archive; a Keynote
    // package's is KN.DocumentArchive (type 1) whose field 2 references the
    // show (type 2) whose field 3 slide tree references the root slide node.
    const rootEntry = objects.documentArchives.entries().next();
    const rootMessages: ArchiveMessage[] = rootEntry.done ? [] : rootEntry.value[1];
    const root = rootMessages[0];
    if (!root || root.type !== TYPE_DOCUMENT) {
      throw new KeynoteOpenError(NOT_KEYNOTE_MESSAGE);
    }
    const showId = referencedId(fields(root.payload), 2);
    const show = archiveOf(objects, showId);
    if (!show || show.type !== TYPE_SHOW) {
      throw new KeynoteOpenError(NOT_KEYNOTE_MESSAGE);
    }
    showFields = fields(show.payload);
    const slideTree = messageAt(showFields, 3);
    const treeNodeIds = slideTree ? referencedIds(slideTree, 2) : [];
    if (treeNodeIds.length === 0) {
      throw new KeynoteOpenError(NOT_KEYNOTE_MESSAGE);
    }
    rootNodeIds = treeNodeIds;
  } catch (error) {
    // Any decode failure before the deck is established as Keynote — bad
    // Snappy, bad protobuf, a Pages/Numbers root — is "not a Keynote
    // presentation we can open" (AC-9/AC-10), always typed (AC-12).
    throw error instanceof KeynoteOpenError && error.message === NOT_KEYNOTE_MESSAGE
      ? error
      : new KeynoteOpenError(NOT_KEYNOTE_MESSAGE);
  }

  try {
    // Show size (field 4: {1: float32 width, 2: float32 height} in Keynote
    // points; 1920×1080 for 16:9) → the model's cm space (AC-22).
    let showWidthPt = 1920;
    let showHeightPt = 1080;
    const size = messageAt(showFields, 4);
    if (size) {
      const wField = size.find((f) => f.num === 1 && f.wire === WIRE_FIXED32);
      const hField = size.find((f) => f.num === 2 && f.wire === WIRE_FIXED32);
      if (wField && hField && float32Of(wField) > 0 && float32Of(hField) > 0) {
        showWidthPt = float32Of(wField);
        showHeightPt = float32Of(hField);
      }
    }
    const slideWidthCm = targetWidthCm(showWidthPt, showHeightPt);
    const pointToCm = slideWidthCm / showWidthPt;
    const slideHeightCm = showHeightPt * pointToCm;

    const resolver: MediaResolver = { minted: new Map(), parts: [], nextIndex: 1 };
    const dropped: KeynoteDropCounts = { objects: 0, effects: 0 };
    const context: SlideBuildContext = { objects, pointToCm, slideWidthCm, slideHeightCm, resolver, dropped };

    const slides: Slide[] = [];
    for (const nodeId of orderedSlideNodeIds(objects, rootNodeIds)) {
      const node = archiveOf(objects, nodeId);
      if (!node) {
        continue;
      }
      const slideId = referencedId(fields(node.payload), 2);
      const slideArchive = archiveOf(objects, slideId);
      if (!slideArchive || slideArchive.type !== TYPE_SLIDE) {
        // An unknown or missing archive where a slide should be — ignored
        // (AC-41), keeping the deck importable.
        continue;
      }
      slides.push(buildSlide(context, fields(slideArchive.payload)));
    }

    if (slides.length === 0) {
      throw new KeynoteOpenError(KEYNOTE_UNREADABLE_MESSAGE);
    }

    return {
      slides,
      layouts: [],
      slideWidthCm,
      slideHeightCm,
      media: resolver.parts,
      opaque: [],
      properties: {},
      dropped,
    };
  } catch (error) {
    // Past the "it is Keynote" gate, a failure is a Keynote file we can't
    // read (AC-11) — normalised so only the user-facing strings ever leave
    // this module, and always as the typed error (AC-12).
    throw error instanceof KeynoteOpenError && error.message === KEYNOTE_UNREADABLE_MESSAGE
      ? error
      : new KeynoteOpenError(KEYNOTE_UNREADABLE_MESSAGE);
  }
}
