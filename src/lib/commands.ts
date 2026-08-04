// The editing command API (SPEC-011): every content mutation the editor
// performs goes through one of the commands below. Each command takes the
// current model plus a selection in model coordinates and returns a NEW
// model state — the input model is never modified, so callers can rely on
// reference inequality to know a command produced a change, and the DOM is
// never consulted (it is a projection of the model, never the source of
// truth).
//
// Immutability contract: commands clone the model's containers (the model
// object, its sections array, each section object and its blocks array) and
// delegate the actual edit to the SPEC-010 primitives in `textEditing`,
// which replace blocks rather than mutating them. Untouched Paragraph
// objects are therefore carried into the new model BY REFERENCE — that
// structural sharing is what keeps a command O(edit) instead of O(document),
// and it is safe precisely because nothing in the editing stack ever mutates
// a block in place. It is also what the canvas's `v-memo` keys on to skip
// re-rendering untouched paragraphs.
import type { DocumentModel } from "./documentModel";
import {
  blockLength,
  clampCaret,
  compareCaret,
  deleteBackward as deleteBackwardInPlace,
  deleteForward as deleteForwardInPlace,
  deleteRange as deleteRangeInPlace,
  insertText as insertTextInPlace,
  splitParagraph as splitParagraphInPlace,
  type CaretPosition,
} from "./textEditing";

/** A selection in model coordinates. `anchor` is where the selection was
 * started, `focus` is the moving end (the caret); they are equal for a
 * collapsed selection (a plain caret) and may be in either document order. */
export interface ModelSelection {
  anchor: CaretPosition;
  focus: CaretPosition;
}

/** What every command returns: the new model state plus the selection after
 * the edit (always collapsed — a caret at the natural post-edit position). */
export interface CommandResult {
  model: DocumentModel;
  selection: ModelSelection;
}

/** A collapsed selection (a caret) at a position. */
export function caretSelection(caret: CaretPosition): ModelSelection {
  return { anchor: caret, focus: caret };
}

export function isCollapsed(selection: ModelSelection): boolean {
  return compareCaret(selection.anchor, selection.focus) === 0;
}

/** The selection's ends in document order, clamped into the model. */
export function orderedRange(
  model: DocumentModel,
  selection: ModelSelection,
): { start: CaretPosition; end: CaretPosition } {
  const a = clampCaret(model, selection.anchor);
  const b = clampCaret(model, selection.focus);
  return compareCaret(a, b) <= 0 ? { start: a, end: b } : { start: b, end: a };
}

// Clones the containers the primitives mutate (see the header comment);
// blocks and everything below them are shared with the input model.
function cloneContainers(model: DocumentModel): DocumentModel {
  return {
    ...model,
    sections: model.sections.map((section) => ({
      ...section,
      blocks: [...section.blocks],
    })),
  };
}

function result(model: DocumentModel, caret: CaretPosition): CommandResult {
  return { model, selection: caretSelection(caret) };
}

/** Inserts text at the selection. A non-collapsed selection is replaced by
 * the inserted text (the "typing over a selection" rule); the text inherits
 * the formatting of the run it lands in. The caret ends up after the
 * inserted text. */
export function insertText(
  model: DocumentModel,
  selection: ModelSelection,
  text: string,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection)
    ? start
    : deleteRangeInPlace(next, start, end);
  return result(next, insertTextInPlace(next, base, text));
}

/** Inserts a forced line break (Shift+Enter) at the selection: a literal
 * "\n" inside the paragraph, which pre-wrap rendering displays as a new
 * line without creating a new paragraph. Replaces a non-collapsed
 * selection. */
export function insertLineBreak(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  return insertText(model, selection, "\n");
}

/** Splits the paragraph at the selection in two (the Enter key), after
 * first deleting a non-collapsed selection. Both halves keep the source
 * paragraph's block-level formatting (alignment, indents, spacing, style
 * reference, heading level), so the new paragraph continues the old one's
 * paragraph formatting. The caret ends up at the start of the second
 * half. */
export function splitBlock(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection)
    ? start
    : deleteRangeInPlace(next, start, end);
  return result(next, splitParagraphInPlace(next, base));
}

/** Deletes the selection's range. Deleting across a paragraph boundary
 * merges the boundary paragraphs (the survivor keeps the first paragraph's
 * block formatting). A collapsed selection is a no-op that still returns a
 * fresh model state. The caret ends up at the deletion point. */
export function deleteRange(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  return result(next, deleteRangeInPlace(next, start, end));
}

/** Backspace: deletes a non-collapsed selection, else the character before
 * the caret; at a paragraph start it merges the paragraph into the previous
 * one. No-op at the very start of the document. */
export function deleteBackward(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  if (!isCollapsed(selection)) {
    return deleteRange(model, selection);
  }
  const next = cloneContainers(model);
  return result(next, deleteBackwardInPlace(next, selection.focus));
}

/** Forward delete: deletes a non-collapsed selection, else the character
 * after the caret; at a paragraph end it merges the next paragraph into
 * this one. No-op at the very end of the document. */
export function deleteForward(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  if (!isCollapsed(selection)) {
    return deleteRange(model, selection);
  }
  const next = cloneContainers(model);
  return result(next, deleteForwardInPlace(next, selection.focus));
}

// --- caret/selection queries (pure reads for keyboard navigation) ----------

/** The caret at the very start of the document. */
export function documentStart(): CaretPosition {
  return { section: 0, block: 0, offset: 0 };
}

/** The caret at the very end of the document (after the last block's last
 * character). */
export function documentEnd(model: DocumentModel): CaretPosition {
  const section = model.sections.length - 1;
  const blocks = model.sections[section].blocks;
  const block = blocks.length - 1;
  return { section, block, offset: blockLength(blocks[block]) };
}

/** The selection covering the whole document (Ctrl/Cmd+A). */
export function selectAll(model: DocumentModel): ModelSelection {
  return { anchor: documentStart(), focus: documentEnd(model) };
}

/** How many blocks the document has across all sections. */
export function blockCount(model: DocumentModel): number {
  let count = 0;
  for (const section of model.sections) {
    count += section.blocks.length;
  }
  return count;
}

/** The 1-based ordinal of a caret's block counting across all sections —
 * "paragraph 3 of 7" for screen-reader caret announcements. */
export function blockOrdinal(model: DocumentModel, caret: CaretPosition): number {
  const c = clampCaret(model, caret);
  let ordinal = c.block + 1;
  for (let s = 0; s < c.section; s += 1) {
    ordinal += model.sections[s].blocks.length;
  }
  return ordinal;
}
