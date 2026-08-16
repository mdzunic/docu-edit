// --- bulleted and numbered lists (SPEC-018) ---------------------------------
import {
  orderedRange,
  result,
  type CommandResult,
  type ModelSelection,
} from "./commands.editing";
import { touchedBlocks, type BlockAddress } from "./commands.paragraph";
import {
  clampListLevel,
  markerStyleOf,
  markerType,
  type ListFormatting,
  type ListMarkerStyle,
  type ListType,
} from "./core/lists";
import { cloneContainers, isTable, type DocumentModel, type Paragraph } from "./documentModel";
import { listRunAround } from "./lists";
import { paragraphAt, paragraphAtCaret, replaceParagraphAt } from "./textEditing";

/** Runs a per-paragraph list update over every touched paragraph, with the
 * same replace-don't-mutate and selection-preserving contract as
 * updateTouchedParagraphs. The updater returns the paragraph's next list
 * formatting, null to remove list formatting, or `undefined` for "leave
 * this paragraph untouched" (object identity kept for v-memo). */
function updateTouchedLists(
  model: DocumentModel,
  selection: ModelSelection,
  update: (paragraph: Paragraph) => ListFormatting | null | undefined,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  for (const address of touchedBlocks(next, { anchor: start, focus: end })) {
    const block = paragraphAt(next, address);
    if (!block) {
      continue;
    }
    const updated = update(block);
    if (updated === undefined) {
      continue;
    }
    const replacement: Paragraph = { ...block };
    if (updated === null) {
      delete replacement.list;
    } else {
      replacement.list = updated;
    }
    replaceParagraphAt(next, address, replacement);
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** Whether every touched paragraph is a list item of the given type — the
 * AC-1/AC-2 toggle predicate, also the toolbar buttons' pressed state. */
export function allTouchedOfListType(
  model: DocumentModel,
  selection: ModelSelection,
  type: ListType,
): boolean {
  const addresses = touchedBlocks(model, selection);
  return addresses.every(
    (address) => paragraphAt(model, address)?.list?.type === type,
  );
}

/** The bulleted/numbered toolbar toggle (AC-1, AC-2): if every touched
 * paragraph is already a list item of `type`, list formatting is removed
 * from all of them; otherwise all of them become items of `type` — a
 * paragraph already in a list (of either type) keeps its level, a plain
 * paragraph joins at level 1. An explicit marker is dropped when the type
 * changes (it belonged to the other type's dropdown). */
export function toggleList(
  model: DocumentModel,
  selection: ModelSelection,
  type: ListType,
): CommandResult {
  const removeAll = allTouchedOfListType(model, selection, type);
  return updateTouchedLists(model, selection, (paragraph) => {
    if (removeAll) {
      return null;
    }
    const current = paragraph.list;
    if (!current) {
      return { type, level: 1 };
    }
    if (current.type === type) {
      return undefined;
    }
    const next: ListFormatting = { ...current, type };
    if (next.marker !== undefined && markerType(next.marker) !== type) {
      delete next.marker;
    }
    return next;
  });
}

/** Tab / Shift+Tab in a list item (AC-3, AC-4): moves every touched list
 * item's level by `delta`, clamped to 1..9 — a no-op at the bounds, and
 * non-list paragraphs in the selection are untouched. */
export function changeListLevel(
  model: DocumentModel,
  selection: ModelSelection,
  delta: number,
): CommandResult {
  return updateTouchedLists(model, selection, (paragraph) => {
    const current = paragraph.list;
    if (!current) {
      return undefined;
    }
    const level = clampListLevel(current.level + delta);
    if (level === current.level) {
      return undefined;
    }
    // The imported per-level indent belongs to the level the item is
    // leaving; the new level renders at the editor default.
    const next: ListFormatting = { ...current, level };
    delete next.indentCm;
    return next;
  });
}

/** A marker-style pick from a list button's dropdown (AC-12, AC-13):
 * applies the style — and the type it implies — to every item at the
 * current level of the current list, i.e. the contiguous run of list
 * paragraphs around the first touched block. The level is the first touched
 * list item's own. No-op when the selection is not in a list. */
export function setListMarker(
  model: DocumentModel,
  selection: ModelSelection,
  marker: ListMarkerStyle,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const result: CommandResult = { model: next, selection: { anchor: start, focus: end } };
  const startParagraph = paragraphAtCaret(next, start);
  const list = startParagraph?.list;
  // The containing block list: the cell's own blocks for a caret inside a
  // table (SPEC-022 — a cell is its own list context), else the section's.
  const table = next.sections[start.section].blocks[start.block];
  const container =
    start.cell && isTable(table)
      ? table.rows[start.cell.row].cells[start.cell.col].blocks
      : next.sections[start.section].blocks;
  const index = start.cell ? start.cell.block : start.block;
  const run = listRunAround(container, index);
  if (!list || !run) {
    return result;
  }
  const level = list.level;
  const type = markerType(marker);
  for (let blockIndex = run.first; blockIndex <= run.last; blockIndex++) {
    const block = container[blockIndex];
    if (block.type !== "paragraph" || block.list?.level !== level) {
      continue;
    }
    const address: BlockAddress = start.cell
      ? { section: start.section, block: start.block, cell: { ...start.cell, block: blockIndex } }
      : { section: start.section, block: blockIndex };
    replaceParagraphAt(next, address, { ...block, list: { ...block.list, type, marker } });
  }
  return result;
}

/** Removes list formatting from every touched paragraph without touching
 * its text — Backspace at the start of an item (AC-9) and Enter on an
 * empty level-1 item (AC-8). */
export function removeListFormat(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  return updateTouchedLists(model, selection, (paragraph) =>
    paragraph.list ? null : undefined,
  );
}

/** The list-autoformat conversion (AC-14, AC-15): the caret paragraph —
 * which the caller has verified contains exactly a typed trigger prefix —
 * becomes an empty level-1 list item of `type`; the prefix text is
 * consumed. Recorded as its own undo step by the caller, so Ctrl/Cmd+Z
 * restores the literal prefix text (AC-16). */
export function applyListAutoformat(
  model: DocumentModel,
  selection: ModelSelection,
  type: ListType,
): CommandResult {
  const next = cloneContainers(model);
  const { start } = orderedRange(next, selection);
  const block = paragraphAtCaret(next, start);
  if (!block) {
    return result(next, start);
  }
  replaceParagraphAt(next, start, { ...block, runs: [], list: { type, level: 1 } });
  return result(next, { ...start, offset: 0 });
}

/** What the toolbar reflects about the selection's list state: the two
 * buttons' pressed states (every touched paragraph of that type, matching
 * the toggle predicate) and the first touched item's effective marker
 * style, "" when the selection is not in a list. */
export interface ToolbarListState {
  bulleted: boolean;
  numbered: boolean;
  marker: string;
}

/** The initial/no-document state — a typed seed for `.vue` refs (espree
 * bars the annotation, same constraint as noSelection). */
export function emptyToolbarListState(): ToolbarListState {
  return { bulleted: false, numbered: false, marker: "" };
}

export function toolbarListState(
  model: DocumentModel,
  selection: ModelSelection,
): ToolbarListState {
  const { start } = orderedRange(model, selection);
  const list = paragraphAtCaret(model, start)?.list;
  return {
    bulleted: allTouchedOfListType(model, selection, "bullet"),
    numbered: allTouchedOfListType(model, selection, "number"),
    marker: list ? markerStyleOf(list) : "",
  };
}
