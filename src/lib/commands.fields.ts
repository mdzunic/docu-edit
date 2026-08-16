// --- SPEC-053 field commands ------------------------------------------------
// Inserting a field, and the update paths F9 / "Update all fields" drive: a
// field's cached value is recomputed only when asked (never on open), locked
// fields are skipped, and the SPEC-033 table of contents now shares this same
// update code path rather than owning its own.
import {
  isCollapsed,
  orderedRange,
  result,
  type CommandResult,
  type ModelSelection,
} from "./commands.editing";
import {
  cloneContainers,
  createFieldSpecRun,
  isEvaluatedField,
  isFieldLocked,
  isMergeChipField,
  isShape,
  isTable,
  isToc,
  type Block,
  type DocumentModel,
  type FieldSpec,
  type Paragraph,
  type Run,
  type Table,
} from "./documentModel";
import {
  formattingAtOffset,
  insertRun as insertRunInPlace,
  paragraphAt,
  type CaretPosition,
} from "./textEditing";
import { deleteRange as deleteRangeInPlace } from "./textEditing";
import { evaluateField, isReadOnlyField, REF_SOURCE_NOT_FOUND, type FieldContext } from "./fields";
import { updateTocEntirely } from "./toc";
import type { TocPageNumberLookup } from "./toc";

/** Inserts a field run at the caret (SPEC-053): a one-character caret atom,
 * exactly like a page-number field or a merge chip. A non-collapsed selection
 * is replaced (the typing-over rule), same as insertMergeChip/insertNote. */
export function insertFieldCommand(model: DocumentModel, selection: ModelSelection, field: FieldSpec): CommandResult {
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
  const caret = insertRunInPlace(next, base, createFieldSpecRun(field, formatting));
  return result(next, caret);
}

/** The data a field update reads (SPEC-053 Determinism/Recalculation): the
 * clock instant, the file name, and the optional pagination resolvers the
 * caller supplies from layout. */
export interface FieldUpdateContext {
  now: Date;
  fileName: string;
  pageOfBookmark?: (name: string) => number | null;
  totalPages?: number;
}

/** Whether a field is one an update path recomputes (SPEC-053 Recalculation):
 * evaluated, unlocked, and not a read-only unsupported type/picture. */
function isUpdatable(field: FieldSpec): boolean {
  return isEvaluatedField(field) && !isFieldLocked(field) && !isReadOnlyField(field);
}

/** Recomputes one field's cached value, returning the (possibly unchanged)
 * field and the target name of a cross-reference whose source is missing —
 * so "Update all fields" can list it rather than fail (AC-28). */
function recomputeField(
  field: FieldSpec,
  here: CaretPosition,
  model: DocumentModel,
  ctx: FieldUpdateContext,
): { field: FieldSpec; brokenRef: string | null } {
  if (!isUpdatable(field)) {
    return { field, brokenRef: null };
  }
  const evalCtx: FieldContext = {
    model,
    now: ctx.now,
    fileName: ctx.fileName,
    here,
    totalPages: ctx.totalPages,
    pageOfBookmark: ctx.pageOfBookmark,
  };
  const value = evaluateField(field, evalCtx);
  const brokenRef =
    (field.kind === "ref" || field.kind === "pageRef") && value === REF_SOURCE_NOT_FOUND ? field.target : null;
  // A cross-reference to a deleted target keeps the error string as its cached
  // value (what Word shows) — still "updated", just to the error. `field` is
  // an evaluated (cached-carrying) member here (isUpdatable guards it), so the
  // cast to the cached-bearing subtype is sound.
  const updated = { ...(field as Extract<FieldSpec, { cached: string }>), cached: value };
  return { field: updated, brokenRef };
}

/** Rewrites every field run of a paragraph via `map`, immutably — a new
 * paragraph only when something changed. */
function mapParagraphFields(
  paragraph: Paragraph,
  base: Omit<CaretPosition, "offset">,
  map: (field: FieldSpec, here: CaretPosition) => FieldSpec,
): Paragraph {
  let changed = false;
  let offset = 0;
  const runs = paragraph.runs.map((run) => {
    const start = offset;
    offset += run.text.length;
    if (run.field && !isMergeChipField(run.field)) {
      const next = map(run.field, { ...base, offset: start });
      if (next !== run.field) {
        changed = true;
        return { ...run, field: next };
      }
    }
    return run;
  });
  return changed ? { ...paragraph, runs } : paragraph;
}

function mapTableFields(
  table: Table,
  base: Omit<CaretPosition, "offset" | "cell">,
  map: (field: FieldSpec, here: CaretPosition) => FieldSpec,
): Table {
  return {
    ...table,
    rows: table.rows.map((row, row_) => ({
      ...row,
      cells: row.cells.map((cell, col) => ({
        ...cell,
        blocks: cell.blocks.map((cellBlock, cellBlockIndex) =>
          mapParagraphFields(cellBlock, { ...base, cell: { row: row_, col, block: cellBlockIndex } }, map),
        ),
      })),
    })),
  };
}

function mapBlockFields(
  block: Block,
  base: Omit<CaretPosition, "offset" | "cell">,
  map: (field: FieldSpec, here: CaretPosition) => FieldSpec,
): Block {
  if (isTable(block)) {
    return mapTableFields(block, base, map);
  }
  if (isShape(block) && block.text) {
    return { ...block, text: { ...block.text, blocks: block.text.blocks.map((b) => mapBlockFields(b, base, map)) } };
  }
  if (block.type === "paragraph") {
    return mapParagraphFields(block, base, map);
  }
  return block;
}

/** Rewrites every field run in the whole document (body, table cells, shape
 * text, headers/footers and notes) via `map`, immutably. Merge chips are
 * never touched (they resolve once, by a merge run). */
function mapAllFields(
  model: DocumentModel,
  map: (field: FieldSpec, here: CaretPosition) => FieldSpec,
): DocumentModel {
  const mapHeaderFooter = (
    group: Partial<Record<string, { blocks: Block[] }>> | undefined,
    section: number,
  ): Partial<Record<string, { blocks: Block[] }>> | undefined => {
    if (!group) {
      return group;
    }
    const next: Record<string, { blocks: Block[] }> = {};
    for (const [kind, content] of Object.entries(group)) {
      if (content) {
        next[kind] = { ...content, blocks: content.blocks.map((b) => mapBlockFields(b, { section, block: -1 }, map)) };
      }
    }
    return next;
  };
  return {
    ...model,
    sections: model.sections.map((section, sectionIndex) => ({
      ...section,
      blocks: section.blocks.map((block, blockIndex) => mapBlockFields(block, { section: sectionIndex, block: blockIndex }, map)),
      headers: mapHeaderFooter(section.headers, sectionIndex) as typeof section.headers,
      footers: mapHeaderFooter(section.footers, sectionIndex) as typeof section.footers,
    })),
    notes: model.notes?.map((note) => ({
      ...note,
      blocks: note.blocks.map((b) => mapBlockFields(b, { section: -1, block: -1 }, map)),
    })),
  };
}

/** F9 (SPEC-053 AC-15): recomputes the field at the caret. A locked or
 * unsupported field is left untouched; a caret not on a field is a no-op. */
export function updateFieldAtCaret(
  model: DocumentModel,
  selection: ModelSelection,
  ctx: FieldUpdateContext,
): CommandResult {
  const caret = orderedRange(model, selection).start;
  const paragraph = paragraphAt(model, caret);
  if (!paragraph) {
    return result(model, caret);
  }
  const next = mapAllFields(model, (field, here) => {
    const sameParagraph =
      here.section === caret.section &&
      here.block === caret.block &&
      here.cell?.row === caret.cell?.row &&
      here.cell?.col === caret.cell?.col &&
      here.cell?.block === caret.cell?.block;
    // The field atom occupies [here.offset, here.offset + 1]; the caret is
    // "on it" when it sits at either edge.
    const onField = sameParagraph && caret.offset >= here.offset && caret.offset <= here.offset + 1;
    return onField ? recomputeField(field, here, model, ctx).field : field;
  });
  return result(next, caret);
}

/** The result of "Update all fields" (SPEC-053 AC-16): the new model plus the
 * cross-references whose target is missing (AC-28 — listed, not fatal). */
export interface UpdateAllFieldsResult {
  model: DocumentModel;
  selection: ModelSelection;
  brokenRefs: string[];
}

/** "Tools → Update all fields" (SPEC-053 AC-16..18): recomputes every
 * unlocked field in the document and rebuilds every SPEC-033 table of
 * contents through the shared TOC update path. Deleted-target
 * cross-references are collected in `brokenRefs` rather than failing. */
export function updateAllFields(
  model: DocumentModel,
  selection: ModelSelection,
  ctx: FieldUpdateContext,
  tocPageNumberFor: TocPageNumberLookup,
): UpdateAllFieldsResult {
  const brokenRefs: string[] = [];
  let next = mapAllFields(model, (field, here) => {
    const { field: updated, brokenRef } = recomputeField(field, here, model, ctx);
    if (brokenRef) {
      brokenRefs.push(brokenRef);
    }
    return updated;
  });
  // Every TOC shares this path now (AC-17, AC-18): no more per-TOC update
  // button owning its own recompute.
  next.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (isToc(block)) {
        next = updateTocEntirely(next, selection, sectionIndex, blockIndex, tocPageNumberFor).model;
      }
    });
  });
  return { model: next, selection, brokenRefs };
}

/** Finds the field run at the caret, with its paragraph and run index. */
function fieldRunAtCaret(
  model: DocumentModel,
  caret: CaretPosition,
): { paragraph: Paragraph; runIndex: number; run: Run; start: number } | null {
  const paragraph = paragraphAt(model, caret);
  if (!paragraph) {
    return null;
  }
  let offset = 0;
  for (let runIndex = 0; runIndex < paragraph.runs.length; runIndex++) {
    const run = paragraph.runs[runIndex];
    const start = offset;
    offset += run.text.length;
    if (run.field && !isMergeChipField(run.field) && caret.offset >= start && caret.offset <= start + run.text.length) {
      return { paragraph, runIndex, run, start };
    }
  }
  return null;
}

/** "Lock field" / "Unlock field" (SPEC-053 AC-13/AC-14): toggles the caret
 * field's `locked` flag so update paths skip (or resume touching) it. */
export function setFieldLockedAtCaret(model: DocumentModel, selection: ModelSelection, locked: boolean): CommandResult {
  const caret = orderedRange(model, selection).start;
  const found = fieldRunAtCaret(model, caret);
  if (!found || !isEvaluatedField(found.run.field!)) {
    return result(model, caret);
  }
  const nextField: FieldSpec = { ...found.run.field } as FieldSpec;
  if (locked) {
    (nextField as { locked?: true }).locked = true;
  } else {
    delete (nextField as { locked?: true }).locked;
  }
  const next = cloneContainers(model);
  const paragraph = paragraphAt(next, caret)!;
  const runs = paragraph.runs.map((run, index) => (index === found.runIndex ? { ...run, field: nextField } : run));
  replaceParagraphRuns(next, caret, runs);
  return result(next, caret);
}

/** "Unlink field" (SPEC-053 AC-13/AC-14): replaces the field run with a plain
 * text run carrying its current displayed result — the field becomes ordinary
 * text and no longer updates. */
export function unlinkFieldAtCaret(
  model: DocumentModel,
  selection: ModelSelection,
  ctx: FieldUpdateContext,
): CommandResult {
  const caret = orderedRange(model, selection).start;
  const found = fieldRunAtCaret(model, caret);
  if (!found) {
    return result(model, caret);
  }
  const field = found.run.field!;
  const value = isEvaluatedField(field)
    ? evaluateField(field, {
        model,
        now: ctx.now,
        fileName: ctx.fileName,
        here: { ...caret, offset: found.start },
        pageOfBookmark: ctx.pageOfBookmark,
        totalPages: ctx.totalPages,
      })
    : found.run.text;
  const plain: Run = { text: value.length > 0 ? value : found.run.text, formatting: { ...found.run.formatting } };
  const next = cloneContainers(model);
  const paragraph = paragraphAt(next, caret)!;
  const runs = paragraph.runs.map((run, index) => (index === found.runIndex ? plain : run));
  replaceParagraphRuns(next, caret, runs);
  return result(next, { ...caret, offset: found.start + plain.text.length });
}

/** Writes a paragraph's runs back into a cloned model at a caret address. */
function replaceParagraphRuns(model: DocumentModel, caret: CaretPosition, runs: Run[]): void {
  const paragraph = paragraphAt(model, caret);
  if (!paragraph) {
    return;
  }
  // paragraphAt returns the live (cloned-shallow) paragraph object; rebuild it.
  const rebuilt: Paragraph = { ...paragraph, runs };
  writeParagraphAt(model, caret, rebuilt);
}

/** Places a rebuilt paragraph at the caret's address inside a cloned model,
 * covering the body, table cells, headers/footers, notes and shape text. */
function writeParagraphAt(model: DocumentModel, caret: CaretPosition, paragraph: Paragraph): void {
  const section = model.sections[caret.section];
  if (!section) {
    return;
  }
  const block = section.blocks[caret.block];
  if (caret.cell && isTable(block)) {
    const cell = block.rows[caret.cell.row]?.cells[caret.cell.col];
    if (cell) {
      const rows = block.rows.map((row, r) =>
        r === caret.cell!.row
          ? {
              ...row,
              cells: row.cells.map((c, col) =>
                col === caret.cell!.col
                  ? { ...c, blocks: c.blocks.map((b, i) => (i === caret.cell!.block ? paragraph : b)) }
                  : c,
              ),
            }
          : row,
      );
      section.blocks[caret.block] = { ...block, rows } as Block;
    }
    return;
  }
  if (block?.type === "paragraph") {
    section.blocks[caret.block] = paragraph;
  }
}
