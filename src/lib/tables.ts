// Table structure commands and navigation queries (SPEC-022), framework-
// free like every lib module. The character-level behaviour inside cells
// lives in textEditing.ts (cell-aware primitives); this module owns the
// grid: inserting a table, inserting/deleting rows and columns, Tab/arrow
// movement between cells, and the rectangular cell-range selection the
// canvas highlights.
//
// Commands follow commands.ts's immutability contract: containers are
// cloned, blocks are replaced (never mutated), untouched blocks are shared
// by reference with the input model — so every command here is undoable by
// the snapshot-based history for free (SPEC-015 AC-20).
import {
  isTable,
  type DocumentModel,
  type Paragraph,
  type Table,
  type TableCell,
  type TableRow,
} from "./documentModel";
import { cmToTwips, twipsToCm } from "./docx/units";
import {
  blockLength,
  clampCaret,
  type CaretPosition,
  type CellPath,
} from "./textEditing";
import { caretSelection, type CommandResult, type ModelSelection } from "./commands";

/** The picker grid's bounds (AC-1) and the dialog's upper limits (AC-2). */
export const PICKER_MAX_ROWS = 10;
export const PICKER_MAX_COLS = 10;
export const MAX_TABLE_ROWS = 200;
export const MAX_TABLE_COLS = 63;

// commands.ts's container clone, restated for the same reason clipboard.ts
// restates it: the undo-registry guard test bars commands.ts from exporting
// helpers.
function cloneContainers(model: DocumentModel): DocumentModel {
  return {
    ...model,
    sections: model.sections.map((section) => ({
      ...section,
      blocks: [...section.blocks],
    })),
  };
}

function emptyParagraph(): Paragraph {
  return { type: "paragraph", runs: [] };
}

function emptyCell(): TableCell {
  return { blocks: [emptyParagraph()] };
}

function emptyRow(columns: number): TableRow {
  return { cells: Array.from({ length: columns }, emptyCell) };
}

// Widths snap to the twip grid both file formats reproduce exactly, the
// same rule as commands.ts's snapCm (SPEC-014 AC-26).
function snapCm(cm: number): number {
  return twipsToCm(cmToTwips(cm));
}

/** Equal column widths spanning the text area (AC-3): the page width minus
 * the horizontal margins, split evenly and snapped to the twip grid. */
export function equalColumnWidthsCm(model: DocumentModel, columns: number): number[] {
  const { pageSize, margins } = model.defaultStyles;
  const textAreaCm = Math.max(1, pageSize.widthCm - margins.leftCm - margins.rightCm);
  const width = snapCm(textAreaCm / columns);
  return Array.from({ length: columns }, () => width);
}

/** The table a caret's block addresses, or null. */
export function tableAt(model: DocumentModel, section: number, block: number): Table | null {
  const target = model.sections[section]?.blocks[block];
  return isTable(target) ? target : null;
}

/** Whether the caret sits inside a table cell — the predicate the editor's
 * Tab/arrow handling branches on (usable from espree-parsed `.vue`
 * scripts). */
export function isInTableCell(caret: CaretPosition | null): boolean {
  return !!caret && caret.cell !== undefined;
}

function cellCaret(section: number, block: number, cell: CellPath): CaretPosition {
  return { section, block, offset: 0, cell };
}

/** Inserts a rows×cols table at the caret (AC-3, AC-4): the caret's
 * paragraph splits around the insertion point — an empty half is not kept
 * as an empty leftover, except that the trailing half always survives so a
 * paragraph follows the table and the caret can never be trapped below it
 * (AC-14). The caret lands in the first cell. A non-collapsed selection
 * first collapses to its focus (Word inserts at the caret, replacing
 * nothing). No-op inside a table: nested tables are out of scope. */
export function insertTable(
  model: DocumentModel,
  selection: ModelSelection,
  rows: number,
  cols: number,
): CommandResult {
  const next = cloneContainers(model);
  const at = clampCaret(next, selection.focus);
  if (at.cell || rows < 1 || cols < 1 || rows > MAX_TABLE_ROWS || cols > MAX_TABLE_COLS) {
    return { model: next, selection: caretSelection(at) };
  }
  const blocks = next.sections[at.section].blocks;
  const target = blocks[at.block];
  const table: Table = {
    type: "table",
    rows: Array.from({ length: rows }, () => emptyRow(cols)),
    columnWidthsCm: equalColumnWidthsCm(next, cols),
  };
  if (isTable(target)) {
    return { model: next, selection: caretSelection(at) };
  }
  const length = blockLength(target);
  const replacement: (Paragraph | Table)[] = [];
  if (at.offset === 0) {
    // At the paragraph's start the whole paragraph moves below the table —
    // it doubles as the guaranteed following paragraph.
    replacement.push(table, target);
  } else if (at.offset >= length) {
    replacement.push(target, table, emptyParagraph());
  } else {
    const [before, after] = splitParagraphAt(target, at.offset);
    replacement.push(before, table, after);
  }
  blocks.splice(at.block, 1, ...replacement);
  const tableIndex = at.block + replacement.indexOf(table);
  return {
    model: next,
    selection: caretSelection(cellCaret(at.section, tableIndex, { row: 0, col: 0, block: 0 })),
  };
}

// A paragraph split into two around a character offset, both halves keeping
// the block-level formatting (the textEditing splitParagraph rule, restated
// locally because that primitive writes into the model in place).
function splitParagraphAt(paragraph: Paragraph, offset: number): [Paragraph, Paragraph] {
  let consumed = 0;
  const before: Paragraph = { ...paragraph, runs: [] };
  const after: Paragraph = {
    ...paragraph,
    runs: [],
    formatting: paragraph.formatting ? { ...paragraph.formatting } : undefined,
  };
  const beforeRuns = [];
  const afterRuns = [];
  for (const run of paragraph.runs) {
    const end = consumed + run.text.length;
    if (end <= offset) {
      beforeRuns.push(run);
    } else if (consumed >= offset) {
      afterRuns.push(run);
    } else {
      const cut = offset - consumed;
      beforeRuns.push({ text: run.text.slice(0, cut), formatting: { ...run.formatting } });
      afterRuns.push({ text: run.text.slice(cut), formatting: { ...run.formatting } });
    }
    consumed = end;
  }
  before.runs = beforeRuns;
  after.runs = afterRuns;
  return [before, after];
}

/** Where a structural command applies: a table block plus the row/column
 * the action anchors to (the caret's cell, or the handle/menu target). */
export interface TableLocation {
  section: number;
  block: number;
  row: number;
  col: number;
}

/** The location of a caret inside a table cell, or null. */
export function caretTableLocation(
  model: DocumentModel,
  caret: CaretPosition | null,
): TableLocation | null {
  if (!caret) {
    return null;
  }
  const c = clampCaret(model, caret);
  if (!c.cell || !tableAt(model, c.section, c.block)) {
    return null;
  }
  return { section: c.section, block: c.block, row: c.cell.row, col: c.cell.col };
}

function replaceTable(
  model: DocumentModel,
  location: TableLocation,
  next: Table,
): void {
  model.sections[location.section].blocks.splice(location.block, 1, next);
}

/** Inserts an empty row above or below the location's row (AC-17, AC-20).
 * The caret stays put; a caret in the shifted rows follows its cell. */
export function insertTableRow(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
  where: "above" | "below",
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, location.section, location.block);
  if (!table) {
    return { model: next, selection };
  }
  const index = where === "above" ? location.row : location.row + 1;
  const rows = [...table.rows];
  rows.splice(index, 0, emptyRow(table.columnWidthsCm.length));
  replaceTable(next, location, { ...table, rows });
  return {
    model: next,
    selection: caretSelection(
      shiftCellCaret(clampCaret(next, selection.focus), location, (cell) =>
        cell.row >= index ? { ...cell, row: cell.row + 1 } : cell,
      ),
    ),
  };
}

/** Inserts an empty column left or right of the location's column (AC-18,
 * AC-20). Column widths rescale so the table keeps its overall width. */
export function insertTableColumn(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
  where: "left" | "right",
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, location.section, location.block);
  if (!table || table.columnWidthsCm.length >= MAX_TABLE_COLS) {
    return { model: next, selection };
  }
  const index = where === "left" ? location.col : location.col + 1;
  const reference = table.columnWidthsCm[location.col];
  const widths = [...table.columnWidthsCm];
  widths.splice(index, 0, reference);
  const total = table.columnWidthsCm.reduce((sum, cm) => sum + cm, 0);
  const grownTotal = total + reference;
  const scaled = widths.map((cm) => snapCm((cm * total) / grownTotal));
  const columnStyleNames = table.odtColumnStyleNames ? [...table.odtColumnStyleNames] : undefined;
  columnStyleNames?.splice(index, 0, undefined);
  const rows = table.rows.map((row) => {
    const cells = [...row.cells];
    cells.splice(index, 0, emptyCell());
    return { ...row, cells };
  });
  const grown: Table = { ...table, rows, columnWidthsCm: scaled };
  if (columnStyleNames) {
    // Rescaled widths no longer match the preserved column styles' declared
    // widths; drop the references so the writer regenerates them.
    delete grown.odtColumnStyleNames;
  }
  replaceTable(next, location, grown);
  return {
    model: next,
    selection: caretSelection(
      shiftCellCaret(clampCaret(next, selection.focus), location, (cell) =>
        cell.col >= index ? { ...cell, col: cell.col + 1 } : cell,
      ),
    ),
  };
}

// A caret's cell path adjusted after a structural change to its own table;
// carets elsewhere pass through unchanged.
function shiftCellCaret(
  caret: CaretPosition,
  location: TableLocation,
  shift: (cell: CellPath) => CellPath,
): CaretPosition {
  if (!caret.cell || caret.section !== location.section || caret.block !== location.block) {
    return caret;
  }
  const cell = shift(caret.cell);
  return cell === caret.cell ? caret : { ...caret, cell };
}

// Removing a table block entirely: the caret lands where the table was
// (AC-23) — the block that follows it takes its index; an emptied section
// keeps one empty paragraph.
function removeTableBlock(next: DocumentModel, location: TableLocation): CommandResult {
  const blocks = next.sections[location.section].blocks;
  blocks.splice(location.block, 1);
  if (blocks.length === 0) {
    blocks.push(emptyParagraph());
  }
  const landing = Math.min(location.block, blocks.length - 1);
  return {
    model: next,
    selection: caretSelection({ section: location.section, block: landing, offset: 0 }),
  };
}

/** Deletes the location's row (AC-19, AC-21); deleting the last remaining
 * row deletes the table and lands the caret where it was (AC-22, AC-23).
 * Otherwise the caret moves to the same column of the row that takes the
 * deleted row's place. */
export function deleteTableRow(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, location.section, location.block);
  if (!table) {
    return { model: next, selection };
  }
  if (table.rows.length <= 1) {
    return removeTableBlock(next, location);
  }
  const rows = [...table.rows];
  rows.splice(location.row, 1);
  replaceTable(next, location, { ...table, rows });
  const row = Math.min(location.row, rows.length - 1);
  return {
    model: next,
    selection: caretSelection(
      cellCaret(location.section, location.block, { row, col: location.col, block: 0 }),
    ),
  };
}

/** Deletes the location's column (AC-19, AC-21); deleting the last
 * remaining column deletes the table (AC-22, AC-23). */
export function deleteTableColumn(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, location.section, location.block);
  if (!table) {
    return { model: next, selection };
  }
  if (table.columnWidthsCm.length <= 1) {
    return removeTableBlock(next, location);
  }
  const widths = [...table.columnWidthsCm];
  widths.splice(location.col, 1);
  const columnStyleNames = table.odtColumnStyleNames ? [...table.odtColumnStyleNames] : undefined;
  columnStyleNames?.splice(location.col, 1);
  const rows = table.rows.map((row) => {
    const cells = [...row.cells];
    cells.splice(location.col, 1);
    return { ...row, cells };
  });
  const shrunk: Table = { ...table, rows, columnWidthsCm: widths };
  if (columnStyleNames) {
    shrunk.odtColumnStyleNames = columnStyleNames;
  }
  replaceTable(next, location, shrunk);
  const col = Math.min(location.col, widths.length - 1);
  return {
    model: next,
    selection: caretSelection(
      cellCaret(location.section, location.block, { row: location.row, col, block: 0 }),
    ),
  };
}

/** Deletes the whole table (AC-19); the caret lands where it was (AC-23). */
export function deleteTable(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
): CommandResult {
  const next = cloneContainers(model);
  if (!tableAt(next, location.section, location.block)) {
    return { model: next, selection };
  }
  return removeTableBlock(next, location);
}

/** Appends a row at the table's end and puts the caret in its first cell —
 * Tab in the last cell (AC-12). */
export function appendTableRowForTab(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, location.section, location.block);
  if (!table) {
    return { model: next, selection };
  }
  const rows = [...table.rows, emptyRow(table.columnWidthsCm.length)];
  replaceTable(next, location, { ...table, rows });
  return {
    model: next,
    selection: caretSelection(
      cellCaret(location.section, location.block, { row: rows.length - 1, col: 0, block: 0 }),
    ),
  };
}

// --- navigation queries (pure reads) ---------------------------------------

/** The caret at the start of the next cell in reading order (AC-10), or
 * null when the caret is in the last cell (the caller appends a row,
 * AC-12) or not in a table. */
export function nextCellCaret(model: DocumentModel, caret: CaretPosition): CaretPosition | null {
  const c = clampCaret(model, caret);
  const table = tableAt(model, c.section, c.block);
  if (!table || !c.cell) {
    return null;
  }
  const cols = table.columnWidthsCm.length;
  if (c.cell.col + 1 < cols) {
    return cellCaret(c.section, c.block, { row: c.cell.row, col: c.cell.col + 1, block: 0 });
  }
  if (c.cell.row + 1 < table.rows.length) {
    return cellCaret(c.section, c.block, { row: c.cell.row + 1, col: 0, block: 0 });
  }
  return null;
}

/** The caret at the end of the previous cell in reading order (AC-11), or
 * null in the first cell (Shift+Tab is then a no-op) or outside a table. */
export function prevCellCaret(model: DocumentModel, caret: CaretPosition): CaretPosition | null {
  const c = clampCaret(model, caret);
  const table = tableAt(model, c.section, c.block);
  if (!table || !c.cell) {
    return null;
  }
  let row = c.cell.row;
  let col = c.cell.col - 1;
  if (col < 0) {
    row -= 1;
    if (row < 0) {
      return null;
    }
    col = table.columnWidthsCm.length - 1;
  }
  const blocks = table.rows[row].cells[col].blocks;
  const block = blocks.length - 1;
  return {
    section: c.section,
    block: c.block,
    offset: blockLength(blocks[block]),
    cell: { row, col, block },
  };
}

/** Where an arrow key leaves a cell when the browser's native caret
 * movement made no progress (AC-13, AC-14): the neighbouring cell in the
 * arrow's direction, or the block before/after the table at its edges.
 * Null when there is nowhere to go (e.g. ArrowUp in the first row of a
 * table that starts the document). */
export function arrowEscapeCaret(
  model: DocumentModel,
  caret: CaretPosition,
  direction: "up" | "down" | "left" | "right",
): CaretPosition | null {
  const c = clampCaret(model, caret);
  const table = tableAt(model, c.section, c.block);
  if (!table || !c.cell) {
    return null;
  }
  const { row, col } = c.cell;
  if (direction === "right") {
    return nextCellCaret(model, c) ?? blockAfterTable(model, c.section, c.block);
  }
  if (direction === "left") {
    return prevCellCaret(model, c) ?? blockBeforeTable(model, c.section, c.block);
  }
  if (direction === "down") {
    if (row + 1 < table.rows.length) {
      const blocks = table.rows[row + 1].cells[col].blocks;
      return {
        section: c.section,
        block: c.block,
        offset: Math.min(c.offset, blockLength(blocks[0])),
        cell: { row: row + 1, col, block: 0 },
      };
    }
    return blockAfterTable(model, c.section, c.block);
  }
  if (row > 0) {
    const blocks = table.rows[row - 1].cells[col].blocks;
    const block = blocks.length - 1;
    return {
      section: c.section,
      block: c.block,
      offset: Math.min(c.offset, blockLength(blocks[block])),
      cell: { row: row - 1, col, block },
    };
  }
  return blockBeforeTable(model, c.section, c.block);
}

function blockAfterTable(
  model: DocumentModel,
  section: number,
  block: number,
): CaretPosition | null {
  const blocks = model.sections[section].blocks;
  if (block + 1 < blocks.length && !isTable(blocks[block + 1])) {
    return { section, block: block + 1, offset: 0 };
  }
  return null;
}

function blockBeforeTable(
  model: DocumentModel,
  section: number,
  block: number,
): CaretPosition | null {
  const blocks = model.sections[section].blocks;
  if (block > 0 && !isTable(blocks[block - 1])) {
    return { section, block: block - 1, offset: blockLength(blocks[block - 1]) };
  }
  return null;
}

// --- cell-range selection (AC-15, AC-16) ------------------------------------

/** A rectangular cell range within one table: `anchor` is where the drag
 * started, `focus` the cell under the pointer. Held by the view while a
 * drag is active; the model's own text selection stays collapsed. */
export interface CellRange {
  section: number;
  block: number;
  anchor: { row: number; col: number };
  focus: { row: number; col: number };
}

/** Typed `null` seed for cell-range refs in `.vue` scripts (espree bars the
 * annotation, same constraint as commands.noSelection). */
export function noCellRange(): CellRange | null {
  return null;
}

/** Typed pass-through for a cell range arriving as a runtime-declared
 * Object prop (espree bars the prop generic — see asParagraphBlock). */
export function asCellRange(value: unknown): CellRange | null {
  return (value ?? null) as CellRange | null;
}

/** The range's normalised bounds (inclusive). */
export function cellRangeBounds(range: CellRange): {
  top: number;
  left: number;
  bottom: number;
  right: number;
} {
  return {
    top: Math.min(range.anchor.row, range.focus.row),
    left: Math.min(range.anchor.col, range.focus.col),
    bottom: Math.max(range.anchor.row, range.focus.row),
    right: Math.max(range.anchor.col, range.focus.col),
  };
}

/** Whether a cell lies inside the (normalised) range — the per-cell
 * highlight predicate (AC-16). */
export function isCellInRange(
  range: CellRange | null,
  section: number,
  block: number,
  row: number,
  col: number,
): boolean {
  if (!range || range.section !== section || range.block !== block) {
    return false;
  }
  const bounds = cellRangeBounds(range);
  return row >= bounds.top && row <= bounds.bottom && col >= bounds.left && col <= bounds.right;
}
