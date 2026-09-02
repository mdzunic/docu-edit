// The media reference walk shared by both format writers' media planners
// (docx/media.ts, odf/media.ts) and lib/images.ts's reference-counted
// pruning. A tiny model-layer module on purpose: the walk is format-neutral
// — it reads only the document model — so it must not live in either
// format's writer with the other importing it across the docx/odf boundary
// (see core/boundary.spec.ts's "format layer boundaries").
import type { DocumentModel } from "./documentModel";

/** Every mediaId an Image block — or, since SPEC-073, a Shape's picture
 * fill — still references, in document order; computed at the point of use
 * (a save, a delete's media pruning) since a save must never trust stale
 * bookkeeping over the model itself, the same principle every other writer
 * derivation follows. A picture-filled shape and an Image of the same file
 * therefore share one part and one media record in every saved package
 * (SPEC-073 AC-4). */
export function referencedMediaIds(model: DocumentModel): string[] {
  const ids: string[] = [];
  const seen = new Set<string>();
  const add = (mediaId: string) => {
    if (!seen.has(mediaId)) {
      seen.add(mediaId);
      ids.push(mediaId);
    }
  };
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (block.type === "image") {
        add(block.mediaId);
      } else if (block.type === "shape" && block.fill?.kind === "picture") {
        add(block.fill.mediaId);
      }
    }
  }
  return ids;
}
