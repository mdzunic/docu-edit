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
  createChipRun,
  isBreak,
  isImage,
  isShape,
  isTable,
  isToc,
  sectionColumnWidthCm,
  type BorderSide,
  type CellBorders,
  type CellVerticalAlign,
  type DocumentModel,
  type Paragraph,
  type Table,
  type TableAlignment,
  type TableCell,
  type TableRow,
  type TableWidthMode,
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

/** Equal column widths spanning `availableCm`, split evenly and snapped to
 * the twip grid. */
function equalWidthsAcross(availableCm: number, columns: number): number[] {
  const width = snapCm(Math.max(0.1, availableCm) / columns);
  return Array.from({ length: columns }, () => width);
}

/** Equal column widths spanning the document's text area (AC-3): the page
 * width minus the horizontal margins, split evenly and snapped to the twip
 * grid. */
export function equalColumnWidthsCm(model: DocumentModel, columns: number): number[] {
  const { pageSize, margins } = model.defaultStyles;
  return equalWidthsAcross(pageSize.widthCm - margins.leftCm - margins.rightCm, columns);
}

/** Equal column widths for a table inserted into `sectionIndex` (SPEC-030
 * AC-13/AC-15): a table in a multi-column section must fit within one text
 * column, not the whole page text area, or it overflows the column and is
 * clipped. Falls back to the section's full content width for single-column
 * sections. */
export function tableColumnWidthsCm(
  model: DocumentModel,
  sectionIndex: number,
  columns: number,
): number[] {
  return equalWidthsAcross(sectionColumnWidthCm(model, sectionIndex), columns);
}

/** The table a caret's block addresses, or null. */
export function tableAt(model: DocumentModel, section: number, block: number): Table | null {
  const target = model.sections[section]?.blocks[block];
  return isTable(target) ? target : null;
}

/** Typed pass-through for a Table arriving as a runtime-declared Object prop
 * (espree bars the type annotation the canvas's own props declaration would
 * need — see documentModel.ts's asParagraphBlock for the same pattern). */
export function asTable(value: unknown): Table {
  return value as Table;
}

/** One row's real (non-covered) cells, each carrying its own grid column
 * index — a merged cell's covered neighbours (SPEC-023) render nothing of
 * their own; the spanning cell's own colspan/rowspan covers their space.
 * Lives here (rather than as a .vue-local helper) so the canvas's template
 * gets a properly typed `TableCell`/`colIndex` pair back regardless of how
 * loosely its own runtime-declared props are typed — see asTable. */
export interface VisibleCellEntry {
  cell: TableCell;
  colIndex: number;
}

export function visibleRowCells(row: TableRow): VisibleCellEntry[] {
  return row.cells
    .map((cell, colIndex) => ({ cell, colIndex }))
    .filter((entry) => !entry.cell.covered);
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
    columnWidthsCm: tableColumnWidthsCm(next, at.section, cols),
  };
  if (isTable(target) || isImage(target) || isShape(target) || isBreak(target) || isToc(target)) {
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

/** Toggles "Table -> Repeat row for each..." on a row (SPEC-046 AC-29,
 * AC-30): the `{{#each name}}` chip is the entire state — no
 * `TableRow.repeat` attribute (see "Repeating table rows" in
 * specs/046-template-gallery-and-mail-merge.md for why) — inserted at the
 * start of the row's first cell's first paragraph. Called again on an
 * already-repeating row, it removes the chip instead. No-op on a header row
 * or an out-of-range location. */
export function toggleRowRepeat(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
  name: string,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, location.section, location.block);
  const row = table?.rows[location.row];
  const firstCell = row?.cells[0];
  const firstParagraph = firstCell?.blocks[0];
  if (!table || !row || !firstCell || !firstParagraph || row.isHeader) {
    return { model: next, selection };
  }
  const existingIndex = firstParagraph.runs.findIndex((run) => run.field?.kind === "each");
  const runs =
    existingIndex !== -1
      ? firstParagraph.runs.filter((_, index) => index !== existingIndex)
      : [createChipRun({ kind: "each", name }), ...firstParagraph.runs];
  const rows = [...table.rows];
  rows[location.row] = {
    ...row,
    cells: [{ ...firstCell, blocks: [{ ...firstParagraph, runs }, ...firstCell.blocks.slice(1)] }, ...row.cells.slice(1)],
  };
  replaceTable(next, location, { ...table, rows });
  return { model: next, selection };
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

// --- cell/border formatting and merge/split (SPEC-023) ----------------------

/** The smallest width a dragged column boundary can leave a column (AC-2). */
export const MIN_COLUMN_WIDTH_CM = 0.5;

/** The smallest minimum height a dragged row boundary can set (AC-4). */
export const MIN_ROW_HEIGHT_CM = 0.4;

/** A grid slot's owning (real) cell: the cell at `originRow`/`originCol`
 * whose `rowSpan`×`colSpan` rectangle covers the slot — itself when the slot
 * is not covered by anything. Every merge/split/border/shading/padding
 * operation resolves through this so a click anywhere in a merged cell's
 * rectangle (or a range that only partly overlaps it) resolves to the same
 * whole cell. */
interface CellOwner {
  originRow: number;
  originCol: number;
  rowSpan: number;
  colSpan: number;
}

/** Every grid slot's owner, computed in one forward pass: row-major order
 * always visits a cell's own (top-left) slot before any slot it covers, so
 * each covered slot's owner is already known by the time it is reached. */
function computeOwnerGrid(table: Table): CellOwner[][] {
  const rowCount = table.rows.length;
  const colCount = table.columnWidthsCm.length;
  const grid: CellOwner[][] = Array.from({ length: rowCount }, () => Array.from<CellOwner>({ length: colCount }));
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      if (grid[r][c]) {
        continue;
      }
      const cell = table.rows[r].cells[c];
      const owner: CellOwner = {
        originRow: r,
        originCol: c,
        rowSpan: cell.rowSpan ?? 1,
        colSpan: cell.colSpan ?? 1,
      };
      for (let dr = 0; dr < owner.rowSpan && r + dr < rowCount; dr++) {
        for (let dc = 0; dc < owner.colSpan && c + dc < colCount; dc++) {
          grid[r + dr][c + dc] = owner;
        }
      }
    }
  }
  return grid;
}

interface GridBounds {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/** Grows a selection's bounds to whole cells: a range that only partly
 * overlaps a merged cell (dragged from inside it, or a stale selection after
 * an earlier merge) expands to cover that cell's whole rectangle, repeating
 * until stable (a newly included cell can itself pull the bounds further). */
function expandRangeToWholeCells(table: Table, bounds: GridBounds): GridBounds {
  const grid = computeOwnerGrid(table);
  let { top, left, bottom, right } = bounds;
  let changed = true;
  while (changed) {
    changed = false;
    for (let r = top; r <= bottom; r++) {
      for (let c = left; c <= right; c++) {
        const owner = grid[r][c];
        if (owner.originRow < top) {
          top = owner.originRow;
          changed = true;
        }
        if (owner.originCol < left) {
          left = owner.originCol;
          changed = true;
        }
        if (owner.originRow + owner.rowSpan - 1 > bottom) {
          bottom = owner.originRow + owner.rowSpan - 1;
          changed = true;
        }
        if (owner.originCol + owner.colSpan - 1 > right) {
          right = owner.originCol + owner.colSpan - 1;
          changed = true;
        }
      }
    }
  }
  return { top, left, bottom, right };
}

/** `total` split into `parts` contiguous integer-sized groups, as equal as
 * possible with any remainder going to the earliest groups — how a split
 * cell's requested rows/columns are distributed across the grid rows/columns
 * its rectangle already occupies (AC-9). */
function partitionSizes(total: number, parts: number): number[] {
  const base = Math.floor(total / parts);
  const remainder = total % parts;
  return Array.from({ length: parts }, (_, i) => base + (i < remainder ? 1 : 0));
}

/** Drops a cell's preserved source XML/style reference: any command that
 * changes a modelled property must not let a stale `docxTcPr`/`odtStyleName`
 * be re-emitted verbatim over top of it (the same rule insertTableColumn
 * already applies to `odtColumnStyleNames`). */
function clearPreservedCell(cell: TableCell): TableCell {
  const next = { ...cell };
  delete next.docxTcPr;
  delete next.odtStyleName;
  return next;
}

/** The same rule as clearPreservedCell, for the table's own preserved
 * `w:tblPr`. */
function clearPreservedTable(table: Table): Table {
  const next = { ...table };
  delete next.docxTblPr;
  return next;
}

/** Inserts `count` new grid columns at `atCol` (AC-9's grow-then-split path
 * for a plain cell split into more columns than it currently spans). Any
 * cell whose span already crosses `atCol - 1` — including the cell being
 * split, whose own span ends there — extends by `count` columns; every row
 * gets `count` new covered placeholder slots at `atCol`, later overwritten
 * by the caller for the rows/columns actually being split. Column widths
 * rescale so the table's overall width is unchanged (the same rule
 * insertTableColumn uses). Must only be called with `atCol >= 1`. */
function insertGridColumns(table: Table, atCol: number, count: number): Table {
  if (count <= 0) {
    return table;
  }
  const grid = computeOwnerGrid(table);
  const refWidth = table.columnWidthsCm[atCol - 1] ?? table.columnWidthsCm[0] ?? 1;
  const originalTotal = table.columnWidthsCm.reduce((sum, w) => sum + w, 0);
  const widths = [...table.columnWidthsCm];
  for (let i = 0; i < count; i++) {
    widths.splice(atCol, 0, refWidth);
  }
  const grownTotal = widths.reduce((sum, w) => sum + w, 0);
  const scaledWidths = widths.map((w) => snapCm((w * originalTotal) / grownTotal));
  const rows = table.rows.map((row, r) => {
    const cells = [...row.cells];
    const owner = grid[r][atCol - 1];
    if (owner.originRow === r) {
      const origin = cells[owner.originCol];
      cells[owner.originCol] = { ...origin, colSpan: (origin.colSpan ?? 1) + count };
    }
    const inserted: TableCell[] = Array.from({ length: count }, () => ({ blocks: [], covered: true }));
    cells.splice(atCol, 0, ...inserted);
    return { ...row, cells };
  });
  const result = clearPreservedTable({ ...table, rows, columnWidthsCm: scaledWidths });
  delete result.odtColumnStyleNames;
  return result;
}

/** Inserts `count` new grid rows at `atRow` — the row analogue of
 * insertGridColumns, used the same way. Must only be called with
 * `atRow >= 1`. */
function insertGridRows(table: Table, atRow: number, count: number): Table {
  if (count <= 0) {
    return table;
  }
  const grid = computeOwnerGrid(table);
  const colCount = table.columnWidthsCm.length;
  const rows = table.rows.map((row) => ({ ...row, cells: [...row.cells] }));
  const bumped = new Set<string>();
  for (let c = 0; c < colCount; c++) {
    const owner = grid[atRow - 1][c];
    const key = `${owner.originRow}:${owner.originCol}`;
    if (bumped.has(key)) {
      continue;
    }
    bumped.add(key);
    const origin = rows[owner.originRow].cells[owner.originCol];
    rows[owner.originRow].cells[owner.originCol] = { ...origin, rowSpan: (origin.rowSpan ?? 1) + count };
  }
  const insertedRows: TableRow[] = Array.from({ length: count }, () => ({
    cells: Array.from({ length: colCount }, () => ({ blocks: [], covered: true })),
  }));
  rows.splice(atRow, 0, ...insertedRows);
  return { ...table, rows };
}

/** Selecting a cell range and choosing "Merge cells" (AC-6, AC-7): the
 * range's bounds expand to whole cells, then every cell in the resulting
 * rectangle collapses into one, at the top-left corner, whose content is
 * every covered cell's paragraphs concatenated in reading order. A range
 * that resolves to a single cell is a no-op. */
export function mergeCells(
  model: DocumentModel,
  selection: ModelSelection,
  range: CellRange,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, range.section, range.block);
  if (!table) {
    return { model: next, selection };
  }
  const bounds = expandRangeToWholeCells(table, cellRangeBounds(range));
  if (bounds.top === bounds.bottom && bounds.left === bounds.right) {
    return { model: next, selection };
  }
  const grid = computeOwnerGrid(table);
  const mergedBlocks: Paragraph[] = [];
  const seen = new Set<string>();
  for (let r = bounds.top; r <= bounds.bottom; r++) {
    for (let c = bounds.left; c <= bounds.right; c++) {
      const owner = grid[r][c];
      const key = `${owner.originRow}:${owner.originCol}`;
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      const cell = table.rows[owner.originRow].cells[owner.originCol];
      if (!cell.covered) {
        mergedBlocks.push(...cell.blocks);
      }
    }
  }
  const originBefore = clearPreservedCell(table.rows[bounds.top].cells[bounds.left]);
  const rowSpan = bounds.bottom - bounds.top + 1;
  const colSpan = bounds.right - bounds.left + 1;
  const rows = table.rows.map((row) => ({ ...row, cells: [...row.cells] }));
  for (let r = bounds.top; r <= bounds.bottom; r++) {
    for (let c = bounds.left; c <= bounds.right; c++) {
      if (r === bounds.top && c === bounds.left) {
        rows[r].cells[c] = {
          ...originBefore,
          blocks: mergedBlocks.length > 0 ? mergedBlocks : [emptyParagraph()],
          colSpan: colSpan > 1 ? colSpan : undefined,
          rowSpan: rowSpan > 1 ? rowSpan : undefined,
          covered: undefined,
        };
      } else {
        rows[r].cells[c] = { blocks: [], covered: true };
      }
    }
  }
  replaceTable(
    next,
    { section: range.section, block: range.block, row: bounds.top, col: bounds.left },
    { ...table, rows },
  );
  const caret = cellCaret(range.section, range.block, { row: bounds.top, col: bounds.left, block: 0 });
  return { model: next, selection: caretSelection(caret) };
}

/** "Split cell…" (AC-8, AC-9): splits the cell at `location` into `rows`
 * rows × `cols` columns. Splitting back down within the cell's current
 * span (e.g. undoing an earlier merge) just repartitions its existing grid
 * rectangle; splitting into more rows/columns than the cell currently spans
 * grows the table's grid first (inserting just enough rows/columns, whose
 * siblings extend across them to stay whole) and then partitions the grown
 * rectangle — so both cases share one code path. The original content lands
 * in the top-left sub-cell; the rest start empty. */
export function splitCell(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
  rows: number,
  cols: number,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, location.section, location.block);
  if (!table || rows < 1 || cols < 1) {
    return { model: next, selection };
  }
  const grid = computeOwnerGrid(table);
  const owner = grid[location.row]?.[location.col];
  if (!owner) {
    return { model: next, selection };
  }
  const { originRow, originCol } = owner;
  let rowSpanOld = owner.rowSpan;
  let colSpanOld = owner.colSpan;
  const projectedCols = table.columnWidthsCm.length + Math.max(0, cols - colSpanOld);
  const projectedRows = table.rows.length + Math.max(0, rows - rowSpanOld);
  if (projectedCols > MAX_TABLE_COLS || projectedRows > MAX_TABLE_ROWS) {
    return { model: next, selection };
  }
  let working = table;
  if (cols > colSpanOld) {
    working = insertGridColumns(working, originCol + colSpanOld, cols - colSpanOld);
    colSpanOld = cols;
  }
  if (rows > rowSpanOld) {
    working = insertGridRows(working, originRow + rowSpanOld, rows - rowSpanOld);
    rowSpanOld = rows;
  }
  const rowGroups = partitionSizes(rowSpanOld, rows);
  const colGroups = partitionSizes(colSpanOld, cols);
  const originalCell = clearPreservedCell(working.rows[originRow].cells[originCol]);
  const originalBlocks = originalCell.blocks.length > 0 ? originalCell.blocks : [emptyParagraph()];
  const newRows = working.rows.map((row) => ({ ...row, cells: [...row.cells] }));
  let rowOffset = 0;
  for (let i = 0; i < rows; i++) {
    const rg = rowGroups[i];
    let colOffset = 0;
    for (let j = 0; j < cols; j++) {
      const cg = colGroups[j];
      const isFirstSubcell = i === 0 && j === 0;
      for (let dr = 0; dr < rg; dr++) {
        for (let dc = 0; dc < cg; dc++) {
          const r = originRow + rowOffset + dr;
          const c = originCol + colOffset + dc;
          newRows[r].cells[c] =
            dr === 0 && dc === 0
              ? {
                  blocks: isFirstSubcell ? originalBlocks : [emptyParagraph()],
                  colSpan: cg > 1 ? cg : undefined,
                  rowSpan: rg > 1 ? rg : undefined,
                }
              : { blocks: [], covered: true };
        }
      }
      colOffset += cg;
    }
    rowOffset += rg;
  }
  replaceTable(next, location, { ...working, rows: newRows });
  const caret = cellCaret(location.section, location.block, { row: originRow, col: originCol, block: 0 });
  return { model: next, selection: caretSelection(caret) };
}

/** Applies `fn` to every distinct real cell overlapping `range` (expanded to
 * whole cells first) and returns the table's updated rows — the shared
 * traversal for the per-cell formatting commands below. `fn` also gets the
 * cell's owner rectangle, for commands (border presets) that need to know
 * where the cell sits relative to the range's own boundary. */
function mapCellsInRange(
  table: Table,
  range: CellRange,
  fn: (cell: TableCell, owner: CellOwner, bounds: GridBounds) => TableCell,
): { rows: TableRow[]; bounds: GridBounds } {
  const bounds = expandRangeToWholeCells(table, cellRangeBounds(range));
  const grid = computeOwnerGrid(table);
  const rows = table.rows.map((row) => ({ ...row, cells: [...row.cells] }));
  const seen = new Set<string>();
  for (let r = bounds.top; r <= bounds.bottom; r++) {
    for (let c = bounds.left; c <= bounds.right; c++) {
      const owner = grid[r][c];
      const key = `${owner.originRow}:${owner.originCol}`;
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      rows[owner.originRow].cells[owner.originCol] = fn(
        table.rows[owner.originRow].cells[owner.originCol],
        owner,
        bounds,
      );
    }
  }
  return { rows, bounds };
}

/** A cell range spanning a whole table — "apply to the whole table" (AC-13)
 * is "apply to every cell", the same per-cell commands "apply to selected
 * cells" (AC-14) use with a smaller range. Null when the location has no
 * table. */
export function fullTableRange(model: DocumentModel, location: TableLocation): CellRange | null {
  const table = tableAt(model, location.section, location.block);
  if (!table) {
    return null;
  }
  return {
    section: location.section,
    block: location.block,
    anchor: { row: 0, col: 0 },
    focus: { row: table.rows.length - 1, col: table.columnWidthsCm.length - 1 },
  };
}

/** The standard border presets a table properties dialog offers (AC-18..21). */
export type BorderPreset = "all" | "outer" | "inner" | "none";

/** Typed pass-through for a border preset coming from a `.vue` template's
 * plain string (same espree constraint as documentModel.ts's
 * asCharacterFormatKey). */
export function asBorderPreset(value: string): BorderPreset {
  return value as BorderPreset;
}

const NO_BORDER: BorderSide = { style: "none", widthPt: 0, color: "#000000" };

function presetEdge(preset: BorderPreset, isOuterEdge: boolean, side: BorderSide): BorderSide {
  if (preset === "none") {
    return NO_BORDER;
  }
  if (preset === "all") {
    return side;
  }
  return isOuterEdge === (preset === "outer") ? side : NO_BORDER;
}

/** A table properties dialog's border section (AC-10..14, AC-18..21): sets
 * `side` on every edge `preset` selects, over every cell `range` covers
 * (the whole table, via fullTableRange, for AC-13; a drag-selected range for
 * AC-14). An edge is "outer" when it sits on the (whole-cell-expanded)
 * range's own boundary; "all" paints every edge, "none" clears every edge,
 * "outer"/"inner" paint one set and clear the other so a repeated
 * application is idempotent rather than additive. */
export function applyBorderPreset(
  model: DocumentModel,
  selection: ModelSelection,
  range: CellRange,
  preset: BorderPreset,
  side: BorderSide,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, range.section, range.block);
  if (!table) {
    return { model: next, selection };
  }
  const { rows, bounds } = mapCellsInRange(table, range, (cell, owner, bounds) => {
    const rectBottom = owner.originRow + owner.rowSpan - 1;
    const rectRight = owner.originCol + owner.colSpan - 1;
    const borders: CellBorders = {
      top: presetEdge(preset, owner.originRow === bounds.top, side),
      bottom: presetEdge(preset, rectBottom === bounds.bottom, side),
      left: presetEdge(preset, owner.originCol === bounds.left, side),
      right: presetEdge(preset, rectRight === bounds.right, side),
    };
    return { ...clearPreservedCell(cell), borders };
  });
  replaceTable(
    next,
    { section: range.section, block: range.block, row: bounds.top, col: bounds.left },
    { ...table, rows },
  );
  return { model: next, selection };
}

/** A table properties dialog's shading section (AC-15), over `range` (the
 * whole table or a selected range, same convention as applyBorderPreset).
 * `fill` undefined clears the cells' shading back to none. */
export function setCellShading(
  model: DocumentModel,
  selection: ModelSelection,
  range: CellRange,
  fill: string | undefined,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, range.section, range.block);
  if (!table) {
    return { model: next, selection };
  }
  const { rows, bounds } = mapCellsInRange(table, range, (cell) => {
    const updated = clearPreservedCell(cell);
    if (fill === undefined) {
      delete updated.shadingFill;
    } else {
      updated.shadingFill = fill;
    }
    return updated;
  });
  replaceTable(
    next,
    { section: range.section, block: range.block, row: bounds.top, col: bounds.left },
    { ...table, rows },
  );
  return { model: next, selection };
}

/** A table properties dialog's cell padding section (AC-16), over `range`. */
export function setCellPadding(
  model: DocumentModel,
  selection: ModelSelection,
  range: CellRange,
  paddingCm: number | undefined,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, range.section, range.block);
  if (!table) {
    return { model: next, selection };
  }
  const { rows, bounds } = mapCellsInRange(table, range, (cell) => {
    const updated = clearPreservedCell(cell);
    if (paddingCm === undefined) {
      delete updated.paddingCm;
    } else {
      updated.paddingCm = Math.max(0, paddingCm);
    }
    return updated;
  });
  replaceTable(
    next,
    { section: range.section, block: range.block, row: bounds.top, col: bounds.left },
    { ...table, rows },
  );
  return { model: next, selection };
}

/** A table properties dialog's vertical alignment section (AC-17), over
 * `range`. */
export function setCellVerticalAlign(
  model: DocumentModel,
  selection: ModelSelection,
  range: CellRange,
  align: CellVerticalAlign | undefined,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, range.section, range.block);
  if (!table) {
    return { model: next, selection };
  }
  const { rows, bounds } = mapCellsInRange(table, range, (cell) => {
    const updated = clearPreservedCell(cell);
    if (align === undefined) {
      delete updated.verticalAlign;
    } else {
      updated.verticalAlign = align;
    }
    return updated;
  });
  replaceTable(
    next,
    { section: range.section, block: range.block, row: bounds.top, col: bounds.left },
    { ...table, rows },
  );
  return { model: next, selection };
}

/** A table properties dialog's alignment section (AC-22): the table's own
 * position on the page. */
export function setTableAlignment(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
  alignment: TableAlignment,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, location.section, location.block);
  if (!table) {
    return { model: next, selection };
  }
  replaceTable(next, location, clearPreservedTable({ ...table, alignment }));
  return { model: next, selection };
}

/** A table properties dialog's width section (AC-23): auto (content-sized),
 * fixed (the `columnWidthsCm` sum) or a percentage of the text area.
 * `percent` is required for, and only meaningful with, `"percent"`. */
export function setTableWidthMode(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
  mode: TableWidthMode,
  percent?: number,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, location.section, location.block);
  if (!table) {
    return { model: next, selection };
  }
  const updated = clearPreservedTable({ ...table, widthMode: mode });
  if (mode === "percent") {
    updated.widthPercent = Math.min(100, Math.max(1, percent ?? table.widthPercent ?? 100));
  } else {
    delete updated.widthPercent;
  }
  replaceTable(next, location, updated);
  return { model: next, selection };
}

/** Dragging a column boundary (AC-1..3): the view live-resizes the DOM
 * during the drag and calls this once, on release, with the full widths
 * array — never the model. Widths are clamped to MIN_COLUMN_WIDTH_CM and
 * snapped to the twip grid; anything else (wrong column count) is a no-op. */
export function setColumnWidths(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
  widthsCm: number[],
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, location.section, location.block);
  if (!table || widthsCm.length !== table.columnWidthsCm.length) {
    return { model: next, selection };
  }
  const columnWidthsCm = widthsCm.map((w) => snapCm(Math.max(MIN_COLUMN_WIDTH_CM, w)));
  const updated = clearPreservedTable({ ...table, columnWidthsCm });
  delete updated.odtColumnStyleNames;
  replaceTable(next, location, updated);
  return { model: next, selection };
}

/** Dragging a row boundary (AC-4, AC-5): sets the row's minimum height —
 * rendering still lets the row grow taller than this for its content, since
 * `minHeightCm` is a floor, never a fixed height. */
export function setRowMinHeight(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
  rowIndex: number,
  minHeightCm: number,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, location.section, location.block);
  if (!table || !table.rows[rowIndex]) {
    return { model: next, selection };
  }
  const rows = [...table.rows];
  const row = { ...rows[rowIndex], minHeightCm: snapCm(Math.max(MIN_ROW_HEIGHT_CM, minHeightCm)) };
  delete row.docxTrPr;
  delete row.odtStyleName;
  rows[rowIndex] = row;
  replaceTable(next, location, { ...table, rows });
  return { model: next, selection };
}

/** Marks (or unmarks) a row as the table's repeating header row (SPEC-043's
 * "mark header row" mechanical fix). Clears the row's preserved source
 * markup the same way `setRowMinHeight` does — `isHeader` is a modelled
 * field now, so the writer must regenerate this row's properties instead of
 * reproducing a stale `docxTrPr`/`odtStyleName` that disagrees with it. */
export function setRowIsHeader(
  model: DocumentModel,
  selection: ModelSelection,
  location: TableLocation,
  rowIndex: number,
  isHeader: boolean,
): CommandResult {
  const next = cloneContainers(model);
  const table = tableAt(next, location.section, location.block);
  if (!table || !table.rows[rowIndex]) {
    return { model: next, selection };
  }
  const rows = [...table.rows];
  const row = { ...rows[rowIndex] };
  if (isHeader) {
    row.isHeader = true;
  } else {
    delete row.isHeader;
  }
  delete row.docxTrPr;
  delete row.odtStyleName;
  rows[rowIndex] = row;
  replaceTable(next, location, { ...table, rows });
  return { model: next, selection };
}
