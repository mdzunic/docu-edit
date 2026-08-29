// The framework-free table command layer for slide editing (SPEC-091), the
// presentation counterpart of the word processor's src/lib/tables.ts. It owns
// the grid of a `TableFrame`: creating a table, inserting/deleting rows and
// columns, resizing columns/rows, merging and splitting cells, per-cell fill /
// border / padding / vertical-alignment, the table-level header/banding look
// flags, and Tab/Shift+Tab cell navigation.
//
// Every command mutates the passed `TableFrame` in place (the presentation
// editing convention — see text.ts's applyRunFormat; the undo stack snapshots
// the whole slide BEFORE the command, so in-place mutation is undoable for
// free). The rectangular cell-range arithmetic (`CellRange`, `cellRangeBounds`,
// `isCellInRange`), the grid limits, the border types and the merged-cell owner
// grid (`computeOwnerGrid`) all come from the product-neutral src/lib/core/
// tables.ts — a slide's tables carry the same border and range semantics as the
// document's; only cell *content* (SlideParagraph[]) is product-specific.
//
// Layering rule (presentation/boundary.spec.ts): imports only presentation
// siblings and src/lib/core.
import {
  cellRangeBounds,
  computeOwnerGrid,
  MAX_TABLE_COLS,
  MAX_TABLE_ROWS,
  MIN_COLUMN_WIDTH_CM,
  MIN_ROW_HEIGHT_CM,
  type BorderPreset,
  type BorderSide,
  type CellBorders,
  type CellOwner,
  type CellRange,
  type CellVerticalAlign,
} from "../core/tables";
import {
  DEFAULT_TABLE_ROW_HEIGHT_CM,
  type SlideFill,
  type SlideParagraph,
  type SlideTableCell,
  type SlideTableRow,
  type TableFrame,
} from "./model";

// --- construction -----------------------------------------------------------

function emptyParagraph(): SlideParagraph {
  return { runs: [] };
}

function emptyCell(): SlideTableCell {
  return { paragraphs: [emptyParagraph()] };
}

function coveredCell(): SlideTableCell {
  return { paragraphs: [], covered: true };
}

function emptyRow(columns: number, heightCm?: number): SlideTableRow {
  const row: SlideTableRow = { cells: Array.from({ length: columns }, emptyCell) };
  if (heightCm !== undefined) {
    row.heightCm = heightCm;
  }
  return row;
}

/** cm rounded to 3 decimals, matching frames.ts's geometry precision so a
 * table's widths survive a build → re-parse round trip. */
function r3(value: number): number {
  return Math.round(value * 1000) / 1000;
}

/** Equal column widths spanning `availableCm`, split evenly and 3-dp snapped. */
function equalWidths(availableCm: number, columns: number): number[] {
  const width = r3(Math.max(0.1, availableCm) / columns);
  return Array.from({ length: columns }, () => width);
}

/** Builds a fresh `rows`×`cols` `TableFrame` centred at 80% of the slide width
 * (AC-5c): equal columns, `DEFAULT_TABLE_ROW_HEIGHT_CM` (0.9cm) rows, the
 * header-row look on, and a minted id. Vertically centred on the slide. */
export function createTableFrame(
  rows: number,
  cols: number,
  slideWidthCm: number,
  slideHeightCm: number,
): TableFrame {
  const r = Math.min(Math.max(1, Math.floor(rows)), MAX_TABLE_ROWS);
  const c = Math.min(Math.max(1, Math.floor(cols)), MAX_TABLE_COLS);
  const widthCm = r3(slideWidthCm * 0.8);
  const heightCm = r3(r * DEFAULT_TABLE_ROW_HEIGHT_CM);
  return {
    id: crypto.randomUUID(),
    xCm: r3((slideWidthCm - widthCm) / 2),
    yCm: r3((slideHeightCm - heightCm) / 2),
    widthCm,
    heightCm,
    kind: "table",
    columnWidthsCm: equalWidths(widthCm, c),
    rows: Array.from({ length: r }, () => emptyRow(c, DEFAULT_TABLE_ROW_HEIGHT_CM)),
    firstRow: true,
  };
}

// --- preservation clearing (AC-12) ------------------------------------------

/** Drops a cell's preserved source markup: any command that edits a modelled
 * cell property must not let a stale `pptxTcPr`/`odpCellStyleName` be re-emitted
 * verbatim over it — exactly Table.docxTblPr's documented rule. */
function clearPreservedCell(cell: SlideTableCell): SlideTableCell {
  const next = { ...cell };
  delete next.pptxTcPr;
  delete next.odpCellStyleName;
  return next;
}

/** The table-level analogue: drops `pptxTblPr` when a command edits the look
 * flags (firstRow/bandRow/styleId) it shadows. */
function clearPreservedTableLook(frame: TableFrame): void {
  delete frame.pptxTblPr;
}

/** Column-width edits invalidate the odp per-column automatic styles (which
 * encode the widths) — drop them so the writer regenerates them. */
function clearColumnStyles(frame: TableFrame): void {
  delete frame.odpColumnStyleNames;
}

// --- geometry helpers -------------------------------------------------------

function widthSum(frame: TableFrame): number {
  return r3(frame.columnWidthsCm.reduce((sum, w) => sum + w, 0));
}

// --- row / column structure (AC-7) ------------------------------------------

/** Inserts an empty row above/below `rowIndex` (AC-7a). */
export function insertTableRow(frame: TableFrame, rowIndex: number, where: "above" | "below"): void {
  const index = where === "above" ? rowIndex : rowIndex + 1;
  frame.rows.splice(index, 0, emptyRow(frame.columnWidthsCm.length, DEFAULT_TABLE_ROW_HEIGHT_CM));
  frame.heightCm = r3(frame.heightCm + DEFAULT_TABLE_ROW_HEIGHT_CM);
}

/** Appends a row at the table's end — Tab in the last cell (AC-6b). */
export function appendTableRow(frame: TableFrame): void {
  insertTableRow(frame, frame.rows.length - 1, "below");
}

/** Inserts an empty column left/right of `colIndex` (AC-7a). The new column
 * takes its neighbour's width and the frame's `widthCm` grows to the new
 * `columnWidthsCm` sum (AC-7b) — no rescaling. */
export function insertTableColumn(frame: TableFrame, colIndex: number, where: "left" | "right"): void {
  if (frame.columnWidthsCm.length >= MAX_TABLE_COLS) {
    return;
  }
  const index = where === "left" ? colIndex : colIndex + 1;
  const reference = frame.columnWidthsCm[colIndex] ?? frame.columnWidthsCm[0] ?? 1;
  frame.columnWidthsCm.splice(index, 0, reference);
  for (const row of frame.rows) {
    row.cells.splice(index, 0, emptyCell());
  }
  frame.widthCm = widthSum(frame);
  clearColumnStyles(frame);
}

/** Deletes `rowIndex`. Returns `false` when this was the last row (the caller
 * deletes the whole frame — AC-7c); otherwise mutates and returns `true`. */
export function deleteTableRow(frame: TableFrame, rowIndex: number): boolean {
  if (frame.rows.length <= 1) {
    return false;
  }
  const removed = frame.rows[rowIndex];
  frame.rows.splice(rowIndex, 1);
  frame.heightCm = r3(Math.max(MIN_ROW_HEIGHT_CM, frame.heightCm - (removed?.heightCm ?? DEFAULT_TABLE_ROW_HEIGHT_CM)));
  normaliseSpans(frame);
  return true;
}

/** Deletes `colIndex`. Returns `false` when this was the last column (AC-7c);
 * otherwise mutates and returns `true`. The frame's `widthCm` shrinks to the
 * new column-width sum. */
export function deleteTableColumn(frame: TableFrame, colIndex: number): boolean {
  if (frame.columnWidthsCm.length <= 1) {
    return false;
  }
  frame.columnWidthsCm.splice(colIndex, 1);
  frame.odpColumnStyleNames?.splice(colIndex, 1);
  for (const row of frame.rows) {
    row.cells.splice(colIndex, 1);
  }
  frame.widthCm = widthSum(frame);
  normaliseSpans(frame);
  return true;
}

/** After a row/column delete a surviving owner cell may still claim a span that
 * now overruns the grid, or a covered slot may have lost its owner. Clamp every
 * owner's span to the grid and re-mark covered slots so the grid stays a
 * regular R×C addressed through `computeOwnerGrid` (AC-9c). */
function normaliseSpans(frame: TableFrame): void {
  const rowCount = frame.rows.length;
  const colCount = frame.columnWidthsCm.length;
  // First clamp spans and reset every slot to "not yet covered".
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      const cell = frame.rows[r].cells[c];
      if (cell.covered) {
        continue;
      }
      if (cell.colSpan && c + cell.colSpan > colCount) {
        cell.colSpan = colCount - c;
      }
      if (cell.rowSpan && r + cell.rowSpan > rowCount) {
        cell.rowSpan = rowCount - r;
      }
      if (cell.colSpan !== undefined && cell.colSpan <= 1) {
        delete cell.colSpan;
      }
      if (cell.rowSpan !== undefined && cell.rowSpan <= 1) {
        delete cell.rowSpan;
      }
    }
  }
  // Recompute the covered map: any slot whose owner is itself becomes a real
  // (empty) cell if it was orphaned; any slot an owner still covers stays
  // covered.
  const owned: boolean[][] = Array.from({ length: rowCount }, () => Array.from<boolean>({ length: colCount }).fill(false));
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      const cell = frame.rows[r].cells[c];
      if (cell.covered || owned[r][c]) {
        continue;
      }
      const rowSpan = cell.rowSpan ?? 1;
      const colSpan = cell.colSpan ?? 1;
      for (let dr = 0; dr < rowSpan; dr++) {
        for (let dc = 0; dc < colSpan; dc++) {
          if (dr === 0 && dc === 0) {
            continue;
          }
          owned[r + dr][c + dc] = true;
        }
      }
    }
  }
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      const cell = frame.rows[r].cells[c];
      if (cell.covered && !owned[r][c]) {
        // Orphaned covered slot → becomes a real empty cell.
        frame.rows[r].cells[c] = emptyCell();
      } else if (!cell.covered && owned[r][c]) {
        // A real cell now sits under a surviving span → becomes covered.
        frame.rows[r].cells[c] = coveredCell();
      }
    }
  }
}

// --- resize (AC-8) ----------------------------------------------------------

/** Applies the widths from a column-boundary drag (AC-8a): clamped to
 * `MIN_COLUMN_WIDTH_CM`, 3-dp snapped, and the frame's `widthCm` set to their
 * sum. A wrong column count is a no-op. */
export function setColumnWidths(frame: TableFrame, widthsCm: number[]): void {
  if (widthsCm.length !== frame.columnWidthsCm.length) {
    return;
  }
  frame.columnWidthsCm = widthsCm.map((w) => r3(Math.max(MIN_COLUMN_WIDTH_CM, w)));
  frame.widthCm = widthSum(frame);
  clearColumnStyles(frame);
}

/** Sets a row's height from a row-boundary drag (AC-8b): a floor, never a fixed
 * height, so the row still grows taller for its content. Clears the row's
 * preserved odp style (which encoded the old height). */
export function setRowHeight(frame: TableFrame, rowIndex: number, heightCm: number): void {
  const row = frame.rows[rowIndex];
  if (!row) {
    return;
  }
  row.heightCm = r3(Math.max(MIN_ROW_HEIGHT_CM, heightCm));
  delete row.odpRowStyleName;
  frame.heightCm = r3(frame.rows.reduce((sum, rr) => sum + (rr.heightCm ?? DEFAULT_TABLE_ROW_HEIGHT_CM), 0));
}

/** Scales the grid proportionally when the frame itself is resized through
 * SPEC-082's handles (AC-8c): every column width by `scaleX`, every row height
 * by `scaleY`, so the grid keeps its relative proportions. The caller
 * (resizeFrame) has already set `widthCm`/`heightCm`. */
export function scaleTableGrid(frame: TableFrame, scaleX: number, scaleY: number): void {
  if (scaleX > 0 && Number.isFinite(scaleX)) {
    frame.columnWidthsCm = frame.columnWidthsCm.map((w) => r3(Math.max(MIN_COLUMN_WIDTH_CM, w * scaleX)));
    clearColumnStyles(frame);
  }
  if (scaleY > 0 && Number.isFinite(scaleY)) {
    for (const row of frame.rows) {
      if (row.heightCm !== undefined) {
        row.heightCm = r3(Math.max(MIN_ROW_HEIGHT_CM, row.heightCm * scaleY));
        delete row.odpRowStyleName;
      }
    }
  }
}

// --- merge / split (AC-9) ---------------------------------------------------

interface GridBounds {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/** Grows a range's bounds to whole cells: a range that only partly overlaps a
 * merged cell expands to cover that cell's whole rectangle, repeating until
 * stable. */
function expandRangeToWholeCells(frame: TableFrame, bounds: GridBounds): GridBounds {
  const grid = computeOwnerGrid(frame);
  let { top, left, bottom, right } = bounds;
  const rowCount = frame.rows.length;
  const colCount = frame.columnWidthsCm.length;
  top = Math.max(0, Math.min(top, rowCount - 1));
  bottom = Math.max(0, Math.min(bottom, rowCount - 1));
  left = Math.max(0, Math.min(left, colCount - 1));
  right = Math.max(0, Math.min(right, colCount - 1));
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

/** `total` split into `parts` contiguous integer groups, remainder to the
 * earliest — how a split cell's requested rows/columns spread across the grid
 * rows/columns its rectangle already occupies. */
function partitionSizes(total: number, parts: number): number[] {
  const base = Math.floor(total / parts);
  const remainder = total % parts;
  return Array.from({ length: parts }, (_, i) => base + (i < remainder ? 1 : 0));
}

/** Merges a rectangular range into its top-left cell (AC-9a): spans set,
 * content concatenated in reading order, every other slot becoming covered. A
 * range that resolves to a single cell is a no-op. */
export function mergeCells(frame: TableFrame, range: CellRange): void {
  const bounds = expandRangeToWholeCells(frame, cellRangeBounds(range));
  if (bounds.top === bounds.bottom && bounds.left === bounds.right) {
    return;
  }
  const grid = computeOwnerGrid(frame);
  const merged: SlideParagraph[] = [];
  const seen = new Set<string>();
  for (let r = bounds.top; r <= bounds.bottom; r++) {
    for (let c = bounds.left; c <= bounds.right; c++) {
      const owner = grid[r][c];
      const key = `${owner.originRow}:${owner.originCol}`;
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      const cell = frame.rows[owner.originRow].cells[owner.originCol];
      if (!cell.covered) {
        merged.push(...cell.paragraphs);
      }
    }
  }
  const originBefore = clearPreservedCell(frame.rows[bounds.top].cells[bounds.left]);
  const rowSpan = bounds.bottom - bounds.top + 1;
  const colSpan = bounds.right - bounds.left + 1;
  for (let r = bounds.top; r <= bounds.bottom; r++) {
    for (let c = bounds.left; c <= bounds.right; c++) {
      if (r === bounds.top && c === bounds.left) {
        const owner: SlideTableCell = { ...originBefore, paragraphs: merged.length > 0 ? merged : [emptyParagraph()] };
        delete owner.covered;
        if (colSpan > 1) {
          owner.colSpan = colSpan;
        } else {
          delete owner.colSpan;
        }
        if (rowSpan > 1) {
          owner.rowSpan = rowSpan;
        } else {
          delete owner.rowSpan;
        }
        frame.rows[r].cells[c] = owner;
      } else {
        frame.rows[r].cells[c] = coveredCell();
      }
    }
  }
}

/** Inserts `count` grid columns at `atCol`, extending any span that crosses the
 * boundary and adding covered placeholders — the grow-then-split path when a
 * cell is split into more columns than it spans. The frame's `widthCm` grows.
 * Must be called with `atCol >= 1`. */
function insertGridColumns(frame: TableFrame, atCol: number, count: number): void {
  if (count <= 0) {
    return;
  }
  const grid = computeOwnerGrid(frame);
  const refWidth = frame.columnWidthsCm[atCol - 1] ?? frame.columnWidthsCm[0] ?? 1;
  for (let i = 0; i < count; i++) {
    frame.columnWidthsCm.splice(atCol, 0, refWidth);
  }
  frame.rows.forEach((row, r) => {
    const owner = grid[r][atCol - 1];
    if (owner.originRow === r) {
      const origin = row.cells[owner.originCol];
      origin.colSpan = (origin.colSpan ?? 1) + count;
    }
    const inserted = Array.from({ length: count }, coveredCell);
    row.cells.splice(atCol, 0, ...inserted);
  });
  frame.widthCm = widthSum(frame);
  clearColumnStyles(frame);
}

/** The row analogue of insertGridColumns. Must be called with `atRow >= 1`. */
function insertGridRows(frame: TableFrame, atRow: number, count: number): void {
  if (count <= 0) {
    return;
  }
  const grid = computeOwnerGrid(frame);
  const colCount = frame.columnWidthsCm.length;
  const bumped = new Set<string>();
  for (let c = 0; c < colCount; c++) {
    const owner = grid[atRow - 1][c];
    const key = `${owner.originRow}:${owner.originCol}`;
    if (bumped.has(key)) {
      continue;
    }
    bumped.add(key);
    const origin = frame.rows[owner.originRow].cells[owner.originCol];
    origin.rowSpan = (origin.rowSpan ?? 1) + count;
  }
  const insertedRows: SlideTableRow[] = Array.from({ length: count }, () => ({
    cells: Array.from({ length: colCount }, coveredCell),
    heightCm: DEFAULT_TABLE_ROW_HEIGHT_CM,
  }));
  frame.rows.splice(atRow, 0, ...insertedRows);
  frame.heightCm = r3(frame.heightCm + count * DEFAULT_TABLE_ROW_HEIGHT_CM);
}

/** Splits the cell at (`rowIndex`, `colIndex`) into `rows`×`cols` (AC-9b):
 * repartitions the cell's current span, growing the grid first when asked for
 * more rows/columns than it currently spans. The original content lands in the
 * top-left sub-cell. */
export function splitCell(frame: TableFrame, rowIndex: number, colIndex: number, rows: number, cols: number): void {
  if (rows < 1 || cols < 1) {
    return;
  }
  const grid = computeOwnerGrid(frame);
  const owner = grid[rowIndex]?.[colIndex];
  if (!owner) {
    return;
  }
  const { originRow, originCol } = owner;
  let rowSpanOld = owner.rowSpan;
  let colSpanOld = owner.colSpan;
  const projectedCols = frame.columnWidthsCm.length + Math.max(0, cols - colSpanOld);
  const projectedRows = frame.rows.length + Math.max(0, rows - rowSpanOld);
  if (projectedCols > MAX_TABLE_COLS || projectedRows > MAX_TABLE_ROWS) {
    return;
  }
  if (cols > colSpanOld) {
    insertGridColumns(frame, originCol + colSpanOld, cols - colSpanOld);
    colSpanOld = cols;
  }
  if (rows > rowSpanOld) {
    insertGridRows(frame, originRow + rowSpanOld, rows - rowSpanOld);
    rowSpanOld = rows;
  }
  const rowGroups = partitionSizes(rowSpanOld, rows);
  const colGroups = partitionSizes(colSpanOld, cols);
  const originalCell = clearPreservedCell(frame.rows[originRow].cells[originCol]);
  const originalParagraphs = originalCell.paragraphs.length > 0 ? originalCell.paragraphs : [emptyParagraph()];
  let rowOffset = 0;
  for (let i = 0; i < rows; i++) {
    const rg = rowGroups[i];
    let colOffset = 0;
    for (let j = 0; j < cols; j++) {
      const cg = colGroups[j];
      const isFirst = i === 0 && j === 0;
      for (let dr = 0; dr < rg; dr++) {
        for (let dc = 0; dc < cg; dc++) {
          const r = originRow + rowOffset + dr;
          const c = originCol + colOffset + dc;
          if (dr === 0 && dc === 0) {
            const sub: SlideTableCell = { paragraphs: isFirst ? originalParagraphs : [emptyParagraph()] };
            if (cg > 1) {
              sub.colSpan = cg;
            }
            if (rg > 1) {
              sub.rowSpan = rg;
            }
            frame.rows[r].cells[c] = sub;
          } else {
            frame.rows[r].cells[c] = coveredCell();
          }
        }
      }
      colOffset += cg;
    }
    rowOffset += rg;
  }
}

// --- per-cell formatting over a range (AC-10) -------------------------------

/** A range spanning the whole table — "apply to the whole table" is "apply to
 * every cell", the same commands "apply to selected cells" use with a smaller
 * range. `section`/`block` are unused by the presentation table commands (one
 * table is edited at a time); they carry 0 for the shared core `CellRange`. */
export function fullTableRange(frame: TableFrame): CellRange {
  return {
    section: 0,
    block: 0,
    anchor: { row: 0, col: 0 },
    focus: { row: frame.rows.length - 1, col: frame.columnWidthsCm.length - 1 },
  };
}

function mapCellsInRange(
  frame: TableFrame,
  range: CellRange,
  fn: (cell: SlideTableCell, owner: CellOwner, bounds: GridBounds) => SlideTableCell,
): void {
  const bounds = expandRangeToWholeCells(frame, cellRangeBounds(range));
  const grid = computeOwnerGrid(frame);
  const seen = new Set<string>();
  for (let r = bounds.top; r <= bounds.bottom; r++) {
    for (let c = bounds.left; c <= bounds.right; c++) {
      const owner = grid[r][c];
      const key = `${owner.originRow}:${owner.originCol}`;
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      frame.rows[owner.originRow].cells[owner.originCol] = fn(
        frame.rows[owner.originRow].cells[owner.originCol],
        owner,
        bounds,
      );
    }
  }
}

/** Sets (or clears with `undefined`) the fill of every cell in `range`
 * (AC-10a). */
export function setCellFill(frame: TableFrame, range: CellRange, fill: SlideFill | undefined): void {
  mapCellsInRange(frame, range, (cell) => {
    const next = clearPreservedCell(cell);
    if (fill === undefined) {
      delete next.fill;
    } else {
      next.fill = fill;
    }
    return next;
  });
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

/** Applies a border preset over `range` (AC-10b), reusing BorderPreset /
 * BorderLineStyle / BorderSide. "outer"/"inner" paint one set and clear the
 * other so a repeated application is idempotent. */
export function applyCellBorderPreset(
  frame: TableFrame,
  range: CellRange,
  preset: BorderPreset,
  side: BorderSide,
): void {
  mapCellsInRange(frame, range, (cell, owner, bounds) => {
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
}

/** Sets (or clears) cell padding over `range` (AC-10c). */
export function setCellPadding(frame: TableFrame, range: CellRange, paddingCm: number | undefined): void {
  mapCellsInRange(frame, range, (cell) => {
    const next = clearPreservedCell(cell);
    if (paddingCm === undefined) {
      delete next.paddingCm;
    } else {
      next.paddingCm = Math.max(0, r3(paddingCm));
    }
    return next;
  });
}

/** Sets (or clears) cell vertical alignment over `range` (AC-10c). */
export function setCellVerticalAlign(
  frame: TableFrame,
  range: CellRange,
  align: CellVerticalAlign | undefined,
): void {
  mapCellsInRange(frame, range, (cell) => {
    const next = clearPreservedCell(cell);
    if (align === undefined) {
      delete next.verticalAlign;
    } else {
      next.verticalAlign = align;
    }
    return next;
  });
}

// --- table-level look flags (AC-10d) ----------------------------------------

/** Toggles the header-row look (`firstRow`). Clears the preserved `pptxTblPr`
 * (which encoded the old look). */
export function setTableHeaderRow(frame: TableFrame, on: boolean): void {
  if (on) {
    frame.firstRow = true;
  } else {
    delete frame.firstRow;
  }
  clearPreservedTableLook(frame);
}

/** Toggles the banded-rows look (`bandRow`). */
export function setTableBandRows(frame: TableFrame, on: boolean): void {
  if (on) {
    frame.bandRow = true;
  } else {
    delete frame.bandRow;
  }
  clearPreservedTableLook(frame);
}

// --- navigation (AC-6b) -----------------------------------------------------

/** The owner cell address of the next non-covered cell in reading order, or
 * `null` when `(row, col)` is the last cell (the caller appends a row). Covered
 * slots are never a navigation target — a click anywhere in a merged rectangle
 * addresses its owner via `computeOwnerGrid`. */
export function nextCellAddress(frame: TableFrame, row: number, col: number): { row: number; col: number } | null {
  const grid = computeOwnerGrid(frame);
  const colCount = frame.columnWidthsCm.length;
  const rowCount = frame.rows.length;
  const owner = grid[row]?.[col];
  if (!owner) {
    return null;
  }
  // Step past the whole owner rectangle so a wide/tall merge is visited once.
  let c = owner.originCol + owner.colSpan;
  let r = owner.originRow;
  while (r < rowCount) {
    while (c < colCount) {
      const o = grid[r][c];
      if (o.originRow === r && o.originCol === c) {
        return { row: r, col: c };
      }
      c += 1;
    }
    r += 1;
    c = 0;
  }
  return null;
}

/** The owner cell address of the previous non-covered cell in reading order, or
 * `null` in the first cell (Shift+Tab is then a no-op). */
export function prevCellAddress(frame: TableFrame, row: number, col: number): { row: number; col: number } | null {
  const grid = computeOwnerGrid(frame);
  const colCount = frame.columnWidthsCm.length;
  const owner = grid[row]?.[col];
  if (!owner) {
    return null;
  }
  let c = owner.originCol - 1;
  let r = owner.originRow;
  while (r >= 0) {
    while (c >= 0) {
      const o = grid[r][c];
      if (o.originRow === r && o.originCol === c) {
        return { row: r, col: c };
      }
      c -= 1;
    }
    r -= 1;
    c = colCount - 1;
  }
  return null;
}
