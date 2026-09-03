// --- table styles, sorting and text↔table conversion (SPEC-076) -------------
//
// The DocumentModel-level commands the Table menu dispatches. The rules
// themselves live in the three framework-free modules this wraps —
// tableStyles.ts, tableSort.ts, tableConvert.ts (AC-78) — so what is here is
// only the model surgery: which block is replaced, which catalogue entry is
// materialised, and which preserved XML a command has to drop.
//
// Follows commands.ts's immutability contract like every sibling module:
// containers are cloned, blocks replaced never mutated, so each command is a
// single undo entry for the snapshot history for free (AC-33, AC-51, AC-60,
// AC-64).
import { cloneContainers, isTable, type DocumentModel, type NamedStyle, type Paragraph, type Table, type TableLook } from "./documentModel";
import { caretSelection, orderedRange, type CommandResult, type ModelSelection } from "./commands.editing";
import { builtInTableStyle } from "./tableStyles";
import { sortTableRows, type SortKey } from "./tableSort";
import { tableToText, textToTable } from "./tableConvert";
import { sectionColumnWidthCm } from "./documentModel";
import type { TableLocation } from "./tables";

/** Replaces the table at `location` with `next` in an already-cloned model. */
function putTable(model: DocumentModel, location: TableLocation, next: Table): void {
  model.sections[location.section].blocks.splice(location.block, 1, next);
}

/** A table whose `look` this spec has edited can no longer re-emit its
 * preserved `w:tblPr` verbatim: `w:tblLook` lives inside it and is now
 * modelled, so the preserved copy would overwrite the edit on save. Dropping
 * the whole block is exactly what editing alignment already does (AC-72) —
 * every other part of `w:tblPr` this spec does not model is preserved for as
 * long as nothing modelled inside it is touched. */
function clearPreservedTableProps(table: Table): void {
  delete table.docxTblPr;
}

/** Sets (or clears) the table's `look` flags (AC-35). `look` undefined
 * removes the field entirely — "no banding at all" is expressed as an absent
 * look, not as six false flags (AC-23). */
export function setTableLook(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
  look: TableLook | undefined,
): CommandResult {
  const next = cloneContainers(model);
  const table = next.sections[location.section]?.blocks[location.block];
  if (!isTable(table)) {
    return { model: next, selection };
  }
  const replacement: Table = { ...table };
  if (look === undefined) {
    delete replacement.look;
  } else {
    replacement.look = { ...look };
  }
  clearPreservedTableProps(replacement);
  putTable(next, location, replacement);
  return { model: next, selection };
}

/** Applies a gallery entry in one action (AC-33): sets `styleId` and `look`,
 * and materialises the built-in style into `DocumentModel.styles` when the
 * document does not already carry it — so an applied look round-trips as a
 * real style definition rather than a dangling reference.
 *
 * `styleId` null is the gallery's "None" entry: it clears both fields
 * (AC-34). The catalogue is left alone in that case — a style the document
 * already carries stays available in the Styles pane. */
export function applyTableStyleLook(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
  styleId: string | null,
  look: TableLook | undefined,
): CommandResult {
  const next = cloneContainers(model);
  const table = next.sections[location.section]?.blocks[location.block];
  if (!isTable(table)) {
    return { model: next, selection };
  }
  const replacement: Table = { ...table };
  if (styleId === null) {
    delete replacement.styleId;
    delete replacement.look;
  } else {
    replacement.styleId = styleId;
    if (look === undefined) {
      delete replacement.look;
    } else {
      replacement.look = { ...look };
    }
    const styles: NamedStyle[] = [...(next.styles ?? [])];
    if (!styles.some((style) => style.id === styleId)) {
      const builtIn = builtInTableStyle(styleId);
      if (builtIn) {
        styles.push(builtIn);
        next.styles = styles;
      }
    }
  }
  clearPreservedTableProps(replacement);
  putTable(next, location, replacement);
  return { model: next, selection };
}

/** Sorts the table's rows (AC-51): one undo entry, the caret collapsed to
 * the table's first sorted cell so it is never left pointing at a row that
 * moved out from under it. */
export function sortTable(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
  keys: SortKey[],
  hasHeader: boolean,
  language?: string,
): CommandResult {
  const next = cloneContainers(model);
  const table = next.sections[location.section]?.blocks[location.block];
  if (!isTable(table)) {
    return { model: next, selection };
  }
  putTable(next, location, sortTableRows(table, keys, hasHeader, language));
  return {
    model: next,
    selection: caretSelection({
      section: location.section,
      block: location.block,
      offset: 0,
      cell: { row: 0, col: 0, block: 0 },
    }),
  };
}

/** Whether the selection is a run of body paragraphs this spec can convert
 * (AC-52): at least one paragraph, no table anywhere in the range, and not
 * inside a cell (a nested table is out of scope, as it is for Insert table).
 * Returns the block range, or null when the conversion does not apply. */
export function convertibleParagraphRange(
  model: DocumentModel,
  selection: ModelSelection,
): { section: number; first: number; last: number } | null {
  const { start, end } = orderedRange(model, selection);
  if (start.cell || end.cell || start.section !== end.section) {
    return null;
  }
  const blocks = model.sections[start.section]?.blocks ?? [];
  const first = Math.min(start.block, end.block);
  const last = Math.max(start.block, end.block);
  if (first > last || last >= blocks.length) {
    return null;
  }
  for (let index = first; index <= last; index += 1) {
    if (blocks[index].type !== "paragraph") {
      return null;
    }
  }
  return { section: start.section, first, last };
}

/** Replaces the selected paragraphs with the table they describe (AC-60).
 * Column widths divide the SECTION's content width equally (AC-59), so a
 * conversion inside a multi-column section fits its own column. */
export function convertTextToTable(
  model: DocumentModel,
  selection: ModelSelection,
  separator: string,
): CommandResult {
  const range = convertibleParagraphRange(model, selection);
  if (!range) {
    return { model, selection };
  }
  const next = cloneContainers(model);
  const blocks = next.sections[range.section].blocks;
  const paragraphs = blocks.slice(range.first, range.last + 1) as Paragraph[];
  const table = textToTable(paragraphs, separator, sectionColumnWidthCm(next, range.section));
  blocks.splice(range.first, range.last - range.first + 1, table);
  return {
    model: next,
    selection: caretSelection({
      section: range.section,
      block: range.first,
      offset: 0,
      cell: { row: 0, col: 0, block: 0 },
    }),
  };
}

/** Replaces the table with the paragraphs it flattens to (AC-64). */
export function convertTableToText(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
  separator: string,
): CommandResult {
  const next = cloneContainers(model);
  const blocks = next.sections[location.section]?.blocks ?? [];
  const table = blocks[location.block];
  if (!isTable(table)) {
    return { model: next, selection };
  }
  const paragraphs = tableToText(table, separator);
  // A table of no rows would leave a hole in the flow; one empty paragraph
  // keeps the section non-empty and the caret placeable.
  const replacement: Paragraph[] =
    paragraphs.length > 0 ? paragraphs : [{ type: "paragraph", runs: [] }];
  blocks.splice(location.block, 1, ...replacement);
  return {
    model: next,
    selection: caretSelection({ section: location.section, block: location.block, offset: 0 }),
  };
}
