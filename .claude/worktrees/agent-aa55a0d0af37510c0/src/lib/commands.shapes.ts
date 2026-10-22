// --- shape text-frame editing seam (SPEC-051) -------------------------------
import { isShape, type Block, type DocumentModel, type Shape } from "./documentModel";

/** Addresses a Shape's text frame as a throwaway single-section document
 * (SPEC-051 AC-17, AC-18) — withNoteModel's seam reused for shape text, so
 * every body-editing primitive (character AND paragraph features, including
 * lists and tables — the same Block content body text accepts) works
 * unmodified against `Shape.text.blocks`. A shape with no text frame yet (or
 * an empty one) starts from one empty paragraph; entering edit mode is what
 * gives the shape a real text frame (see useShapes.ts's handleShapeEditText,
 * which populates it before this is ever called). */
export function withShapeModel(model: DocumentModel, location: { section: number; block: number }): DocumentModel {
  const block = model.sections[location.section]?.blocks[location.block];
  const blocks: Block[] =
    isShape(block) && block.text && block.text.blocks.length > 0
      ? block.text.blocks
      : [{ type: "paragraph", runs: [] }];
  return { ...model, sections: [{ blocks }] };
}

/** Folds a `withShapeModel` command result back into the real document —
 * commitNoteModel's sibling: the synthetic section's blocks become the
 * shape's `text.blocks` (its `vAlign` is untouched), the real sections come
 * back from `before`, and every other model-level change the inner command
 * made is kept from `result`. */
export function commitShapeModel(
  before: DocumentModel,
  location: { section: number; block: number },
  result: DocumentModel,
): DocumentModel {
  const blocks = result.sections[0].blocks;
  const sections = before.sections.map((section, sectionIndex) => {
    if (sectionIndex !== location.section) {
      return section;
    }
    return {
      ...section,
      blocks: section.blocks.map((block, blockIndex) => {
        if (blockIndex !== location.block) {
          return block;
        }
        const shape = block as Shape;
        return { ...shape, text: { blocks, vAlign: shape.text?.vAlign ?? "top" } };
      }),
    };
  });
  return { ...result, sections };
}
