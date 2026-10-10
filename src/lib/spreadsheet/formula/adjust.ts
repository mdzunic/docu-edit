// Structural adjustment of formula text (SPEC-124 "Formula adjustment") —
// THE rule every reference in formula text follows when whole rows or
// columns are inserted or deleted. Token-level over SPEC-119's
// rewriteReferences, so trivia, `$` markers, sheet-name quoting and every
// non-reference token survive byte-for-byte, and a formula nothing touches
// comes back as the identical string.
//
// - A reference is on the edited sheet when its sheet (the host's when it
//   names none) is the edited sheet — for model formulas, names compared
//   case-insensitively with sheetNameOf(edit.sheetId); for retained text
//   (structure.ts's structuralMapper) whatever `onEditedSheet` decides,
//   since source-time names resolve through SPEC-121's
//   resolveRetainedSheetName.
// - Its parts move per structure.ts's geometry table whatever their `$`: a
//   single cell through adjustIndex, a range's two ends through adjustSpan
//   (an insert at its first row moves it, strictly inside expands it, a
//   delete straddling it shrinks it). A whole-column reference ignores row
//   edits and a whole-row one column edits. A reference whose area is gone
//   prints `#REF!` (`Sheet2!#REF!` when it named a sheet).
// - A 3-D reference (`Sheet1:Sheet3!A5`) and an external one (`[1]Book!A1`,
//   an `external` token) are left alone; strings, names and structured
//   references are not references here. Callers never pass
//   `formulaUnparsed` text (wire text is never transformed).
import { adjustIndex, adjustSpan, editAxis, type StructuralEdit } from "../structure";
import { rewriteReferences, type RefPart, type ReferenceNode } from "./references";

/** Whether a reference's sheet — as written, undefined when it names none —
 * is the edited sheet. */
export type OnEditedSheet = (sheet: string | undefined) => boolean;

/** `text` (canonical, no "=") with every reference into the edited sheet
 * adjusted for `edit` (see the module header). */
export function adjustFormula(text: string, hostSheetId: string, edit: StructuralEdit, sheetNameOf: (sheetId: string) => string): string {
  const edited = sheetNameOf(edit.sheetId).toUpperCase();
  const host = sheetNameOf(hostSheetId);
  return adjustFormulaReferences(text, edit, (sheet) => (sheet ?? host).toUpperCase() === edited);
}

/** adjustFormula with the sheet test given (see the module header). */
export function adjustFormulaReferences(text: string, edit: StructuralEdit, onEditedSheet: OnEditedSheet): string {
  return rewriteReferences(text, (ref) => (ref.sheetEnd === undefined && onEditedSheet(ref.sheet) ? adjustReference(ref, edit) : undefined));
}

/** A part with its position on the edit's axis set. */
function withIndex(part: RefPart, axis: "row" | "col", index: number): RefPart {
  return axis === "row" ? { ...part, row: index } : { ...part, col: index };
}

/** One reference adjusted: undefined when nothing moves, null when its area
 * is gone. */
export function adjustReference(ref: ReferenceNode, edit: StructuralEdit): ReferenceNode | null | undefined {
  const axis = editAxis(edit);
  const start = ref.start[axis];
  if (start === undefined) {
    // A whole column under a row edit, or a whole row under a column edit.
    return undefined;
  }
  // A whole column or row has one part; it stands for both ends.
  const endPart = ref.end ?? (ref.start.row === undefined || ref.start.col === undefined ? ref.start : undefined);
  if (!endPart) {
    const index = adjustIndex(start, edit);
    if (index === null) {
      return null;
    }
    return index === start ? undefined : { ...ref, start: withIndex(ref.start, axis, index) };
  }
  const end = endPart[axis] ?? start;
  const reversed = start > end;
  const span = adjustSpan(Math.min(start, end), Math.max(start, end), edit);
  if (span === null) {
    return null;
  }
  const [low, high] = span;
  const nextStart = reversed ? high : low;
  const nextEnd = reversed ? low : high;
  if (nextStart === start && nextEnd === end) {
    return undefined;
  }
  return { ...ref, start: withIndex(ref.start, axis, nextStart), end: withIndex(endPart, axis, nextEnd) };
}
