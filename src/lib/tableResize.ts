// DOM helpers for the live column/row resize preview (SPEC-023 AC-1..5).
// Kept out of DocumentView.vue because typed DOM element access needs real
// type annotations, which ESLint's espree parser rejects in `.vue` scripts
// (see domCaret.ts's similar framework-free DOM helpers) — the view just
// calls these with whatever loosely-typed values its own template/event
// handlers carry.

/** CSS's own fixed cm-to-px ratio (96 CSS px per inch): both resize axes
 * convert a drag's pixel delta straight to centimetres without depending on
 * the table's own measured geometry, which a live width change can
 * invalidate mid-drag (e.g. by reflowing the page). */
export const CM_PER_PX = 2.54 / 96;

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

/** One row's `<tr>` element by index, or null out of range. */
export function tableRowElement(tableEl: HTMLTableElement, rowIndex: number): HTMLTableRowElement | null {
  const row = tableEl.querySelectorAll(":scope > tbody > tr")[rowIndex];
  return row instanceof HTMLTableRowElement ? row : null;
}

/** Live-previews a column resize (AC-1) directly on the DOM: every `<col>`'s
 * width, and the table's own width when it isn't sized by its content
 * (`table-layout: auto`, SPEC-023 AC-23's "auto" width mode) — clamping to
 * the minimum width is the caller's job (lib/tables.ts's MIN_COLUMN_WIDTH_CM);
 * this just paints whatever it is given. */
export function previewColumnWidths(tableEl: HTMLTableElement, widthsCm: number[]): void {
  tableEl.querySelectorAll("colgroup col").forEach((col, index) => {
    if (widthsCm[index] !== undefined && col instanceof HTMLElement) {
      col.style.width = `${widthsCm[index]}cm`;
    }
  });
  if (tableEl.style.tableLayout !== "auto") {
    tableEl.style.width = `${widthsCm.reduce((sum, w) => sum + w, 0)}cm`;
  }
}

/** Live-previews a row's minimum height (AC-4): a floor in table layout, so
 * the row still renders taller than this for its content (AC-5). */
export function previewRowHeight(trEl: HTMLTableRowElement, heightCm: number): void {
  trEl.style.height = `${heightCm}cm`;
}
