// The media reference walk shared by both format writers' media planners
// (docx/media.ts, odf/media.ts) and lib/images.ts's reference-counted
// pruning. A tiny model-layer module on purpose: the walk is format-neutral
// — it reads only the document model — so it must not live in either
// format's writer with the other importing it across the docx/odf boundary
// (see core/boundary.spec.ts's "format layer boundaries").
import { isDrawingGroup, isImage, type DocumentModel, type DrawingMember } from "./documentModel";

/** Every mediaId an Image block in the document still references, in
 * document order — computed at the point of use (a save, a delete's media
 * pruning) since a save must never trust stale bookkeeping over the model
 * itself, the same principle every other writer derivation follows. An Image
 * nested inside a drawing group (SPEC-057) counts too — its media must be
 * written for the group member to render. */
export function referencedMediaIds(model: DocumentModel): string[] {
  const ids: string[] = [];
  const seen = new Set<string>();
  const collectMember = (member: DrawingMember): void => {
    if (isDrawingGroup(member)) {
      member.members.forEach(collectMember);
    } else if (isImage(member) && !seen.has(member.mediaId)) {
      seen.add(member.mediaId);
      ids.push(member.mediaId);
    }
  };
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (block.type === "image" && !seen.has(block.mediaId)) {
        seen.add(block.mediaId);
        ids.push(block.mediaId);
      } else if (isDrawingGroup(block)) {
        block.members.forEach(collectMember);
      }
    }
  }
  return ids;
}
