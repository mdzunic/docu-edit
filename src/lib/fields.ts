// The unified field evaluator (SPEC-053). A field is a run carrying a
// `FieldSpec` (see documentModel.ts) whose displayed value is *computed* from
// the document and the clock rather than stored as literal text. This module
// owns the closed date/time format catalogue, the per-format renderer, and
// the evaluator that turns a `FieldSpec` into the string a reader displays.
//
// Framework-free per CLAUDE.md: no Vue imports, unit-testable without mounting
// a component. Determinism comes from `clock.ts` — the evaluator never calls
// `Date.now()`/`new Date()` directly; the caller passes the instant read from
// `now()`.
import type { CaretPosition } from "./textEditing";
import { blockText, compareCaret, paragraphAt } from "./textEditing";
import { formatOrdinal } from "./core/lists";
import {
  DATE_TIME_CATALOGUE,
  canonicalDateTimeFormatFromDocx,
  dateTimeFormatEntry,
  formatFromOdtDateStyleName,
  isSupportedDateTimeFormat,
  odtDateStyleName,
  renderDateTime,
  type DateTimeFormatEntry,
} from "./core/dateTimeFormats";
import { documentBodyStats } from "./wordCount";
import { effectiveLanguage } from "./core/language";
import { bookmarkCaretIncludingHidden } from "./links";
import { captionRefText, isCaptionParagraph } from "./captions";
import { isEvaluatedField, isMergeChipField, isPaginationDependentField } from "./documentModel";
import { chipLabel } from "./mergeChips";
import type { DocumentModel, FieldSpec, PageNumberFormat, RefFormat } from "./documentModel";

/** The text a cross-reference shows when its target bookmark is gone
 * (SPEC-053) — Word's exact wording, so the string round-trips as content
 * both applications recognise. */
export const REF_SOURCE_NOT_FOUND = "Error! Reference source not found.";

// The date/time format catalogue, its renderer and the `.docx`/`.odt` name
// mappings moved to src/lib/core/dateTimeFormats.ts in SPEC-093's two-phase
// core move (so lib/presentation can share one date machinery). They are
// re-exported here unchanged, names included, so this module's consumers —
// FieldDialog.vue, useFields.ts, the docx/odt readers and writers, and
// fields.spec.ts — and their imports stay exactly as they were.
export {
  DATE_TIME_CATALOGUE,
  canonicalDateTimeFormatFromDocx,
  dateTimeFormatEntry,
  formatFromOdtDateStyleName,
  isSupportedDateTimeFormat,
  odtDateStyleName,
  renderDateTime,
};
export type { DateTimeFormatEntry };

// --- field evaluation --------------------------------------------------------

/** The context a field evaluation reads from. `model`, `now` and `fileName`
 * are always available; the pagination-dependent members are supplied by the
 * caller that knows layout (rendering, the update commands), and left absent
 * where the value should fall back to `cached`. */
export interface FieldContext {
  model: DocumentModel;
  /** The instant from `clock.now()` — never read the wall clock here. */
  now: Date;
  /** The document's file name (SPEC-053 `fileName`); "" when unsaved. */
  fileName: string;
  /** The field run's own position, for `ref` above/below and self page. */
  here?: CaretPosition | null;
  /** The page the field sits on (`page`, and `ref` page context). */
  pageNumber?: number;
  /** The document's total page count (`numpages`). */
  totalPages?: number;
  /** The page a bookmark target sits on (`pageRef`, `ref` show:page). */
  pageOfBookmark?: (name: string) => number | null;
}

/** A custom document property's value by name (SPEC-041), or "" when absent. */
function customPropertyValue(model: DocumentModel, name: string): string {
  return model.properties?.customProperties?.find((property) => property.name === name)?.value ?? "";
}

/** The paragraph a bookmark names, for a cross-reference's text — or null
 * when the target no longer exists. */
function bookmarkTargetText(model: DocumentModel, name: string): string | null {
  const caret = bookmarkCaretIncludingHidden(model, name);
  if (!caret) {
    return null;
  }
  const paragraph = paragraphAt(model, caret);
  return paragraph ? blockText(paragraph) : null;
}

/** The above/below position of a cross-reference target relative to the
 * reference itself (SPEC-053), using textEditing's document-order comparator. */
function relativePosition(model: DocumentModel, targetName: string, here: CaretPosition | null | undefined): string | null {
  if (!here) {
    return null;
  }
  const target = bookmarkCaretIncludingHidden(model, targetName);
  if (!target) {
    return null;
  }
  return compareCaret(target, here) <= 0 ? "above" : "below";
}

/** Evaluates a cross-reference (`ref`/`pageRef`) to its display string,
 * returning `REF_SOURCE_NOT_FOUND` when the target is gone (SPEC-053). */
function evaluateRef(field: Extract<FieldSpec, { kind: "ref" | "pageRef" }>, ctx: FieldContext): string {
  const exists = bookmarkCaretIncludingHidden(ctx.model, field.target) !== null;
  if (!exists) {
    return REF_SOURCE_NOT_FOUND;
  }
  const show: RefFormat = field.kind === "pageRef" ? "page" : field.show;
  if (show === "page") {
    const page = ctx.pageOfBookmark?.(field.target);
    return page != null ? String(page) : field.cached;
  }
  if (show === "position") {
    return relativePosition(ctx.model, field.target, ctx.here) ?? field.cached;
  }
  if (show === "labelAndNumber" || show === "numberOnly" || show === "textOnly" || show === "full") {
    const caret = bookmarkCaretIncludingHidden(ctx.model, field.target);
    const paragraph = caret ? paragraphAt(ctx.model, caret) : null;
    if (paragraph && isCaptionParagraph(paragraph)) {
      return captionRefText(paragraph, show);
    }
    // A non-caption target (a heading/bookmark) has no label/number to show —
    // fall back to its plain text, the same value `text`/`full` renders.
    return bookmarkTargetText(ctx.model, field.target) ?? REF_SOURCE_NOT_FOUND;
  }
  return bookmarkTargetText(ctx.model, field.target) ?? REF_SOURCE_NOT_FOUND;
}

/** Computes a field's live display string (SPEC-053). Fields whose value is
 * not recomputable from `ctx` (unsupported types, unsupported date pictures,
 * pagination values with no layout in `ctx`) fall back to their `cached`
 * result — the value the file was saved with, which is the truth for a field
 * this editor does not recalculate. Merge chips and `page`/`numpages` are not
 * handled here: chips resolve via `mergeChips.ts`, and page fields render
 * through `pageNumbering.ts`'s `fieldValueText`, which needs page context. */
export function evaluateField(field: FieldSpec, ctx: FieldContext): string {
  switch (field.kind) {
    case "page":
      return formatOrdinal(field.format, ctx.pageNumber ?? 1);
    case "numpages":
      return formatOrdinal(field.format, ctx.totalPages ?? 1);
    case "date":
    case "time":
      return renderDateTime(field.format, ctx.now, effectiveLanguage(undefined, ctx.model.defaultStyles.language)) ?? field.cached;
    case "author":
      return ctx.model.properties?.author ?? "";
    case "docTitle":
      return ctx.model.properties?.title ?? "";
    case "docSubject":
      return ctx.model.properties?.subject ?? "";
    case "docKeywords":
      return ctx.model.properties?.keywords ?? "";
    case "fileName":
      return field.path ? ctx.fileName : ctx.fileName.replace(/^.*[\\/]/, "");
    case "wordCount":
      return String(documentBodyStats(ctx.model).words);
    case "docProperty":
      return customPropertyValue(ctx.model, field.name);
    case "ref":
    case "pageRef":
      return evaluateRef(field, ctx);
    case "seq":
      // A caption's number is computed by captions.ts's renumberCaptions over
      // the whole document (per-label document order), not per field — this
      // path just returns the last renumbered value (SPEC-055 AC-13).
      return field.cached;
    case "unsupported":
      return field.cached;
    default:
      // page/numpages/merge chips — not evaluated here (see doc comment).
      return "";
  }
}

/** The message shown for a field this editor cannot recompute (SPEC-053):
 * an unsupported field type, or a date/time field whose picture string is
 * outside the catalogue. Both are shown read-only with their saved value. */
export const UNSUPPORTED_FIELD_MESSAGE =
  "This field type isn't supported — its value is shown as it was saved";

/** Whether a field is shown read-only with `UNSUPPORTED_FIELD_MESSAGE`
 * (SPEC-053): an `unsupported` type, or a date/time field carrying a picture
 * string outside the catalogue. */
export function isReadOnlyField(field: FieldSpec): boolean {
  if (field.kind === "unsupported") {
    return true;
  }
  if (field.kind === "date" || field.kind === "time") {
    return !isSupportedDateTimeFormat(field.format);
  }
  return false;
}

/** The `.docx` numbering-switch fragment for a page-number format, shared by
 * the writer and any picture builder. */
export function pageNumberFormatOrdinal(format: PageNumberFormat, value: number): string {
  return formatOrdinal(format, value);
}

/** A field's saved display value without any recomputation (SPEC-053): its
 * `cached` result, a merge chip's label, or "" for `page`/`numpages` (which
 * carry no cached value and need page context). Used where a renderer has the
 * field but not the evaluation context — a table cell, say — since a field is
 * never updated on open, its `cached` string is exactly what the document
 * displays there. */
export function cachedFieldText(field: FieldSpec): string {
  if (isMergeChipField(field)) {
    return chipLabel(field);
  }
  return isEvaluatedField(field) ? field.cached : "";
}

/** The value a field displays on screen (SPEC-053 AC-19, AC-20): its saved
 * `cached` result — a field is NEVER recomputed on open (Recalculation rule 3),
 * so what the file was saved with is exactly what shows. The one exception is a
 * pagination-dependent field (`page`, `numpages`, `pageRef`, `ref` show:page —
 * rule 1), which has no stable cached value in a live editor and is evaluated
 * live from the page context in `ctx`. Merge chips resolve to their label
 * through `cachedFieldText`. This is the seam every on-screen render path uses;
 * evaluating unconditionally (the AC-19/20 defect) shows today's date/word
 * count for a document saved months ago. */
export function fieldDisplayText(field: FieldSpec, ctx: FieldContext): string {
  return isPaginationDependentField(field) ? evaluateField(field, ctx) : cachedFieldText(field);
}

/** The read-only explanation for a field, or "" when the field is editable —
 * bound as an atom's `title` (SPEC-053 AC-42). */
export function isReadOnlyFieldTitle(field: FieldSpec): string {
  return isReadOnlyField(field) ? UNSUPPORTED_FIELD_MESSAGE : "";
}

// Typed pass-throughs for the SPEC-053 dialogs' runtime Object/Array props —
// the espree constraint bars `PropType<…>`/`as` in `.vue` scripts (see
// CommentsPane.vue / MergeRunDialog.vue for the same pattern), so these give
// the templates typed access to their list props.

/** One row of the Date and time dialog's format list. */
export interface DateTimePreview {
  format: string;
  label: string;
  isTime: boolean;
}

export function asDateTimePreviews(value: unknown): DateTimePreview[] {
  return value as DateTimePreview[];
}

/** One heading target row of the Cross-reference dialog. */
export interface CrossRefHeading {
  index: number;
  text: string;
  level: number;
}

export function asCrossRefHeadings(value: unknown): CrossRefHeading[] {
  return value as CrossRefHeading[];
}

/** One bookmark target row of the Cross-reference dialog. */
export interface CrossRefBookmark {
  name: string;
}

export function asCrossRefBookmarks(value: unknown): CrossRefBookmark[] {
  return value as CrossRefBookmark[];
}

/** One caption target row of the Cross-reference dialog (SPEC-055): its
 * derived anchor, label (the category it groups under) and full text. */
export interface CrossRefCaption {
  anchor: string;
  label: string;
  text: string;
}

export function asCrossRefCaptions(value: unknown): CrossRefCaption[] {
  return value as CrossRefCaption[];
}

/** Typed pass-through for a plain string list prop (the Field dialog's
 * custom-property names). */
export function asStringList(value: unknown): string[] {
  return value as string[];
}

/** Typed pass-through for a `FieldSpec` built in a `.vue` script from plain
 * option state (the Field dialog's draft field). */
export function asFieldSpec(value: unknown): FieldSpec {
  return value as FieldSpec;
}

/** Typed `null` seed for a `FieldSpec | null` ref in a `.vue` script (espree
 * bars the annotation there) — the selected-field state's initial value. */
export function noFieldSpec(): FieldSpec | null {
  return null;
}
