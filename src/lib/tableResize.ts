// DOM helpers for the live column/row resize preview (SPEC-023 AC-1..5).
// Kept out of DocumentView.vue because typed DOM element access needs real
// type annotations, which ESLint's espree parser rejects in `.vue` scripts
// (see domCaret.ts's similar framework-free DOM helpers) — the view just
// calls these with whatever loosely-typed values its own template/event
// handlers carry.

/** CSS's own fixed cm-to-px ratio (96 CSS px per inch): both resize axes
 * convert a drag's pixel delta straight to centimetres without depending on
 * the table's own measured geometry, which a live width change can
 * invalidate mid-drag (e.g. by reflowing the page). Moved to core/units.ts;
 * re-exported so existing import paths keep working. */
export { CM_PER_PX } from "./core/units";

// The preview painters and row lookup moved to core/tables.ts (model-free
// DOM helpers); re-exported so existing import paths keep working. The drag
// state types and findTableElement stay here — they encode the document
// renderer's section/block addressing scheme.
export { previewColumnWidths, previewRowHeight, tableRowElement } from "./core/tables";

export interface ColumnResizeDrag {
  section: number;
  block: number;
  colIndex: number;
  startX: number;
  startWidthsCm: number[];
  tableEl: HTMLTableElement;
  liveWidthsCm: number[] | null;
}

/** Typed `null` seed for a column-resize-drag ref/variable in `.vue`
 * scripts (espree bars the annotation — see domCaret.ts's noElement). */
export function noColumnResizeDrag(): ColumnResizeDrag | null {
  return null;
}

export interface RowResizeDrag {
  section: number;
  block: number;
  rowIndex: number;
  startY: number;
  startHeightCm: number;
  trEl: HTMLTableRowElement;
  liveHeightCm: number | null;
}

/** Typed `null` seed for a row-resize-drag ref/variable in `.vue` scripts
 * (same constraint as noColumnResizeDrag). */
export function noRowResizeDrag(): RowResizeDrag | null {
  return null;
}

/** The table element a table block's own paragraph data attributes address
 * (SPEC-022's contract): resize needs the live DOM table to preview into
 * directly, unlike every other table action, which only ever touches the
 * model. */
export function findTableElement(root: Element, section: number, block: number): HTMLTableElement | null {
  const paragraph = root.querySelector(`[data-section="${section}"][data-block="${block}"][data-row]`);
  const table = paragraph ? paragraph.closest("table.doc-table") : null;
  return table instanceof HTMLTableElement ? table : null;
}

