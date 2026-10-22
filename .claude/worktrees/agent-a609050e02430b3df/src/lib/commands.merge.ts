// --- mail-merge chips (SPEC-046) --------------------------------------------
import {
  isCollapsed,
  orderedRange,
  result,
  type CommandResult,
  type ModelSelection,
} from "./commands.editing";
import { cloneContainers, createChipRun, type DocumentModel, type MergeFieldSpec } from "./documentModel";
import {
  deleteRange as deleteRangeInPlace,
  formattingAtOffset,
  insertRun as insertRunInPlace,
  paragraphAt,
} from "./textEditing";

/** Inserts a mail-merge chip at the caret (AC-12): a field run built the
 * same way SPEC-029's page-number fields are, so it is atomic for caret/
 * selection arithmetic. A non-collapsed selection is deleted first (the
 * typing-over rule), same as insertNote. */
export function insertMergeChip(model: DocumentModel, selection: ModelSelection, field: MergeFieldSpec): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection) ? start : deleteRangeInPlace(next, start, end);
  const paragraph = paragraphAt(next, base);
  if (!paragraph) {
    return result(next, base);
  }
  const formatting = { ...formattingAtOffset(paragraph, base.offset) };
  delete formatting.linkHref;
  delete formatting.linkAnchor;
  const caret = insertRunInPlace(next, base, createChipRun(field, formatting));
  return result(next, caret);
}
