// Retained references (SPEC-124 "Retained references") — the textual
// adjustment of the reference-bearing XML this product does not model yet.
// SPEC-116's splice base re-emits every unmodelled element verbatim, so a
// structural edit (and SPEC-125's moves, with their own mapper) must move the
// row and column numbers inside that text too, element by element, until the
// spec that models each element registers its own ReferenceAdjuster and
// retires its rows here. The format halves — xlsx/retainedReferences.ts and
// ods/retainedReferences.ts — hold the tables; this module is their shared
// contract:
//
// - A RetainedReferenceMapper answers in MODEL terms. `range` and `index`
//   take the current model name of the sheet the text is on — the format
//   half resolves every source-time sheet name it reads (a formula prefix,
//   an ODS address's sheet, `worksheetSource@sheet`) through SPEC-121's
//   resolveRetainedSheetName first, and a host part's sheet through the
//   sheet part's relationships (xlsx) or the table holding it (ods).
//   `range` returns the moved range, the SAME object when it does not move,
//   or null when its area is gone; `index` the moved index (structuralMapper
//   maps an index in a deleted band to the band's start). `formula` takes
//   canonical text exactly as the retained text names sheets (source time)
//   and the id of the sheet hosting it ("" for workbook-level text) and
//   resolves names itself; it returns the identical string when nothing
//   moves.
// - rewriteRetainedReferences replaces, copy-on-write, what it changes: a new
//   OpaquePart (and a new `opaque` array), a new `odsContentEnvelope`, a new
//   sheet object for a new `sourceXml` or rows — never mutating one an undo
//   snapshot may hold — and leaves every part it does not change, by object
//   identity. Every sheet whose own retained text changed is marked dirty.
import type { GridRange } from "./address";
import type { SpreadsheetModel } from "./model";
import { rewriteOdsRetainedReferences } from "./ods/retainedReferences";
import { rewriteXlsxRetainedReferences } from "./xlsx/retainedReferences";

export interface RetainedReferenceMapper {
  /** null: the area is gone. */
  range(sheetName: string, range: GridRange): GridRange | null;
  index(sheetName: string, axis: "row" | "col", index: number): number | null;
  formula(canonical: string, hostSheetId: string): string;
}

/** Adjusts the retained source's reference-bearing text through `mapper`
 * (see the module header). */
export function rewriteRetainedReferences(model: SpreadsheetModel, mapper: RetainedReferenceMapper): void {
  const changed = model.format === "xlsx" ? rewriteXlsxRetainedReferences(model, mapper) : rewriteOdsRetainedReferences(model, mapper);
  if (changed.size === 0) {
    return;
  }
  model.sheets = model.sheets.map((sheet) => (changed.has(sheet.id) && !sheet.dirty ? { ...sheet, dirty: true as const } : sheet));
}
