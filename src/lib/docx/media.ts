// Plans the package additions a .docx save needs for the model's images
// (SPEC-024): media parts under word/media/, `r:embed` relationships in
// word/_rels/document.xml.rels, and Default content-type extensions in
// [Content_Types].xml. Every save (same-format or not) mints a fresh set of
// media part names/relationship ids from `DocumentModel.media` rather than
// trying to reuse whatever an opened .docx originally had — the reader
// excludes media parts from `opaque` precisely so this is the one source of
// truth, and a resaved image is byte-identical even though its package path
// may differ from the source file's own (Word does not care, and nothing
// else in the package references a media path except the relationship this
// module writes).
import type { Block, DocumentModel } from "../documentModel";
import { ensureMediaEncoded, extensionForContentType, type MediaPart } from "../core/media";
import { referencedMediaIds } from "../mediaRefs";
import { base64ToBytes } from "../opaque";
import { yieldToMainThread } from "../core/idleScheduler";
import type { Relationship } from "../core/xml";
import type { ZipEntryInput } from "../core/zipWriter";

export const IMAGE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/** The next unused `rId` at or after `start`, given the relationship ids a
 * preserved rels part (or an in-progress plan) already claims. */
export function nextFreeRelId(taken: Set<string>, start = 1): string {
  let n = start;
  while (taken.has(`rId${n}`)) {
    n++;
  }
  return `rId${n}`;
}

export interface MediaWritePlan {
  /** The relationship id an Image's `w:drawing` should `r:embed` — "" for a
   * media id this plan dropped (an SVG: .docx cannot store one, AC-2's
   * insert-time conversion is the normal path there; this only fires for a
   * cross-format Save As of a .odt document with an SVG image, reported via
   * opaque.ts's collectDropRecords "svgImage" record). The caller omits the
   * image block entirely when this returns "". */
  relIdFor(mediaId: string): string;
  /** word/media/imageN.ext parts to add to the package. */
  zipEntries: ZipEntryInput[];
  /** `<Relationship>` elements to append to word/_rels/document.xml.rels. */
  relationshipsXml: string;
  /** Distinct file extensions used, for `[Content_Types].xml`'s Default
   * entries — added only if not already declared (see media.spec.ts). */
  extensions: string[];
  /** The ids this plan claimed — seeds later allocators (hyperlinks,
   * SPEC-029's header/footer relationships) so nothing collides. */
  relIds: Set<string>;
}

const EMPTY_PLAN: MediaWritePlan = {
  relIdFor: () => "",
  zipEntries: [],
  relationshipsXml: "",
  extensions: [],
  relIds: new Set(),
};

/** Builds a MediaWritePlan for an explicit list of referenced media ids —
 * the shared engine `planMediaWrite` (the whole model's body images) and
 * SPEC-029's per-header/footer-part media planning both call, so a header's
 * images mint part names/relationship ids from the exact same allocator
 * body images do (no separate "reset the counter" bug class). `counter`
 * starts part naming past whatever the caller has already minted (threaded
 * across calls so a document's images and its headers'/footers' images
 * never collide on `word/media/imageN.ext`), and returns the counter's new
 * value for the next call. */
/** Mutable accumulator threaded through `planOneMedia` — factored out so the
 * synchronous per-image step (base64 decode included) is identical whether
 * it is driven by a plain `for` loop (`planMediaWriteFor`) or one that
 * yields to the main thread between images (`planMediaWriteForChunked`). */
interface MediaWriteAccum {
  relIdOf: Map<string, string>;
  zipEntries: ZipEntryInput[];
  relParts: string[];
  extensions: Set<string>;
  taken: Set<string>;
  existingPartNames: Set<string>;
  n: number;
}

function planOneMedia(accum: MediaWriteAccum, mediaId: string, byId: Map<string, MediaPart>): void {
  const part = byId.get(mediaId);
  // .docx has no SVG representation (AC-2 converts on insert; this only
  // ever fires on a cross-format Save As carrying a .odt SVG through) —
  // dropped, not written; the caller skips the image block entirely.
  if (!part || part.contentType === "image/svg+xml") {
    return;
  }
  // SPEC-047: force-encode a still-deferred part before reading its bytes, so
  // a save that happens before the background encode pass has run still
  // writes every image (round-trip fidelity is the project's first
  // constraint) — a no-op once the part carries its `data`.
  ensureMediaEncoded(part);
  const ext = extensionForContentType(part.contentType);
  let partName = `word/media/image${accum.n}.${ext}`;
  while (accum.existingPartNames.has(partName)) {
    accum.n++;
    partName = `word/media/image${accum.n}.${ext}`;
  }
  accum.extensions.add(ext);
  const relId = nextFreeRelId(accum.taken);
  accum.taken.add(relId);
  accum.relIdOf.set(mediaId, relId);
  accum.zipEntries.push({ name: partName, data: base64ToBytes(part.data) });
  accum.relParts.push(
    `<Relationship Id="${relId}" Type="${IMAGE_RELATIONSHIP}" Target="media/image${accum.n}.${ext}"/>`,
  );
  accum.existingPartNames.add(partName);
  accum.n++;
}

function planFromAccum(accum: MediaWriteAccum): MediaWritePlan & { nextCounter: number } {
  return {
    relIdFor: (mediaId: string) => accum.relIdOf.get(mediaId) ?? "",
    zipEntries: accum.zipEntries,
    relationshipsXml: accum.relParts.join(""),
    extensions: Array.from(accum.extensions),
    relIds: new Set(accum.relIdOf.values()),
    nextCounter: accum.n,
  };
}

function planMediaWriteFor(
  referenced: string[],
  mediaParts: MediaPart[] | undefined,
  takenRelIds: Set<string>,
  existingPartNames: Set<string>,
  counter: number,
): MediaWritePlan & { nextCounter: number } {
  if (referenced.length === 0) {
    return { ...EMPTY_PLAN, nextCounter: counter };
  }
  const byId = new Map((mediaParts ?? []).map((part) => [part.mediaId, part]));
  const accum: MediaWriteAccum = {
    relIdOf: new Map(),
    zipEntries: [],
    relParts: [],
    extensions: new Set(),
    taken: new Set(takenRelIds),
    existingPartNames,
    n: counter,
  };
  for (const mediaId of referenced) {
    planOneMedia(accum, mediaId, byId);
  }
  return planFromAccum(accum);
}

// SPEC-047 AC-15's ≤100ms main-thread-block-during-save ceiling: each
// image's base64 decode (media.ts's own `base64ToBytes`, an `atob` plus a
// byte-by-byte copy) has no `await` of its own, so a fixture with hundreds
// of images reads as one long synchronous task. Time-budgeted like
// zipWriter.ts's SAVE_YIELD_BUDGET_MS, for the same reason: a large image
// costs more to decode than a small one.
const MEDIA_YIELD_BUDGET_MS = 4;

async function planMediaWriteForChunked(
  referenced: string[],
  mediaParts: MediaPart[] | undefined,
  takenRelIds: Set<string>,
  existingPartNames: Set<string>,
  counter: number,
): Promise<MediaWritePlan & { nextCounter: number }> {
  if (referenced.length === 0) {
    return { ...EMPTY_PLAN, nextCounter: counter };
  }
  const byId = new Map((mediaParts ?? []).map((part) => [part.mediaId, part]));
  const accum: MediaWriteAccum = {
    relIdOf: new Map(),
    zipEntries: [],
    relParts: [],
    extensions: new Set(),
    taken: new Set(takenRelIds),
    existingPartNames,
    n: counter,
  };
  let lastYield = performance.now();
  for (const mediaId of referenced) {
    planOneMedia(accum, mediaId, byId);
    if (performance.now() - lastYield > MEDIA_YIELD_BUDGET_MS) {
      await yieldToMainThread();
      lastYield = performance.now();
    }
  }
  return planFromAccum(accum);
}

/** Builds a MediaWritePlan for every image the model's blocks still
 * reference. `takenRelIds` seeds the id allocator so a same-format save
 * (appending to a preserved rels part) never collides with an id that part
 * already uses; `existingPartNames` does the same for the word/media/imageN
 * naming, so a package that already has an (unmodelled) word/media/image1.*
 * of its own never collides with a freshly written one. */
export function planMediaWrite(
  model: DocumentModel,
  takenRelIds: Set<string> = new Set(),
  existingPartNames: Set<string> = new Set(),
): MediaWritePlan {
  const referenced = referencedMediaIds(model);
  return planMediaWriteFor(referenced, model.media, takenRelIds, existingPartNames, 1);
}

/** The chunked twin of `planMediaWrite` (SPEC-047 AC-15): `writeDocx`'s own
 * entry point for the body's media plan, yielding to the main thread
 * between images instead of decoding every referenced image's base64 in
 * one synchronous pass. Every other caller (tests, header/footer parts,
 * which realistically hold at most a handful of images) uses the
 * synchronous `planMediaWrite` above. */
export async function planMediaWriteChunked(
  model: DocumentModel,
  takenRelIds: Set<string> = new Set(),
  existingPartNames: Set<string> = new Set(),
): Promise<MediaWritePlan> {
  const referenced = referencedMediaIds(model);
  return planMediaWriteForChunked(referenced, model.media, takenRelIds, existingPartNames, 1);
}

/** A header/footer part's own MediaWritePlan (SPEC-029 AC-6): images inside
 * header/footer content need a `r:embed` relationship in THAT part's own
 * `_rels/headerN.xml.rels`, never the main document's — headers/footers are
 * independent OPC parts. `existingPartNames`/`counter` are threaded across
 * every media plan a save builds (the main document's and every header's/
 * footer's) so `word/media/imageN.ext` naming never collides between them;
 * the caller reads back `nextCounter` and passes it into the next call. */
export function planMediaWriteForBlocks(
  blocks: Block[],
  mediaParts: MediaPart[] | undefined,
  existingPartNames: Set<string>,
  counter: number,
): MediaWritePlan & { nextCounter: number } {
  const referenced: string[] = [];
  const seen = new Set<string>();
  for (const block of blocks) {
    if (block.type === "image" && !seen.has(block.mediaId)) {
      seen.add(block.mediaId);
      referenced.push(block.mediaId);
    }
  }
  // A header/footer part's own rels file starts numbering at rId1 — it has
  // no styles/numbering relationship occupying it the way the main
  // document's does.
  return planMediaWriteFor(referenced, mediaParts, new Set(), existingPartNames, counter);
}

/** The next free rId a preserved relationships part's own ids leave open —
 * seeds planMediaWrite's allocator for a same-format save appending to an
 * existing word/_rels/document.xml.rels. */
export function takenRelIdsOf(rels: Iterable<Relationship>): Set<string> {
  return new Set(Array.from(rels, (rel) => rel.id));
}
