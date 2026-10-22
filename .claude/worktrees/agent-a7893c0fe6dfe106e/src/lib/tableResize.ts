// DOM helpers for the live column/row resize preview (SPEC-023 AC-1..5).
// Kept out of DocumentView.vue because typed DOM element access needs real
// type annotations, which ESLint's espree parser rejects in `.vue` scripts
// (see domCaret.ts's similar framework-free DOM helpers) — the view just
// calls these with whatever loosely-typed values its own template/event
// handlers carry.

// The cm/px ratio lives in core/units.ts and the preview painters in
// core/tables.ts (model-free DOM helpers). The drag state types and
// findTableElement stay here — they encode the document renderer's
// section/block addressing scheme.
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

