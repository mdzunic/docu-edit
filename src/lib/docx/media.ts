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
import type { DocumentModel } from "../documentModel";
import { base64ToBytes } from "../opaque";
import { extensionForContentType } from "./drawing";
import type { Relationship } from "./xml";
import type { ZipEntryInput } from "./zipWriter";

export const IMAGE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/** Every mediaId an Image block in the document still references — the
 * writer's own mirror of lib/images.ts's referencedMediaIds, computed here
 * too since a save must never trust stale bookkeeping over the model itself
 * (the same principle every other writer derivation in this file follows). */
function referencedMediaIds(model: DocumentModel): string[] {
  const ids: string[] = [];
  const seen = new Set<string>();
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (block.type === "image" && !seen.has(block.mediaId)) {
        seen.add(block.mediaId);
        ids.push(block.mediaId);
      }
    }
  }
  return ids;
}

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
}

const EMPTY_PLAN: MediaWritePlan = {
  relIdFor: () => "",
  zipEntries: [],
  relationshipsXml: "",
  extensions: [],
};

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
  if (referenced.length === 0) {
    return EMPTY_PLAN;
  }
  const byId = new Map((model.media ?? []).map((part) => [part.mediaId, part]));
  const relIdOf = new Map<string, string>();
  const zipEntries: ZipEntryInput[] = [];
  const relParts: string[] = [];
  const extensions = new Set<string>();
  const taken = new Set(takenRelIds);
  let counter = 1;

  for (const mediaId of referenced) {
    const part = byId.get(mediaId);
    // .docx has no SVG representation (AC-2 converts on insert; this only
    // ever fires on a cross-format Save As carrying a .odt SVG through) —
    // dropped, not written; the caller skips the image block entirely.
    if (!part || part.contentType === "image/svg+xml") {
      continue;
    }
    const ext = extensionForContentType(part.contentType);
    let partName = `word/media/image${counter}.${ext}`;
    while (existingPartNames.has(partName)) {
      counter++;
      partName = `word/media/image${counter}.${ext}`;
    }
    extensions.add(ext);
    const relId = nextFreeRelId(taken);
    taken.add(relId);
    relIdOf.set(mediaId, relId);
    zipEntries.push({ name: partName, data: base64ToBytes(part.data) });
    relParts.push(`<Relationship Id="${relId}" Type="${IMAGE_RELATIONSHIP}" Target="media/image${counter}.${ext}"/>`);
    counter++;
  }

  return {
    relIdFor: (mediaId: string) => relIdOf.get(mediaId) ?? "",
    zipEntries,
    relationshipsXml: relParts.join(""),
    extensions: Array.from(extensions),
  };
}

/** The next free rId a preserved relationships part's own ids leave open —
 * seeds planMediaWrite's allocator for a same-format save appending to an
 * existing word/_rels/document.xml.rels. */
export function takenRelIdsOf(rels: Iterable<Relationship>): Set<string> {
  return new Set(Array.from(rels, (rel) => rel.id));
}
