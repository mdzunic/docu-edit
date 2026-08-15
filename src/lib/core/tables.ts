// The product-neutral half of tables (SPEC-022, cell/border formatting
// SPEC-023): the border/alignment/width-mode types, the grid limits, the
// rectangular cell-range arithmetic and the DOM resize-preview helpers.
// Everything here is independent of the document model — a Slides product's
// tables carry the same border and range semantics; only cell *content*
// (paragraph blocks) is product-specific and stays on lib/documentModel.ts's
// TableCell. The editing commands stay in lib/tables.ts.

/** A border line's rendering (SPEC-023): the standard presets plus width and
 * colour. `"none"` paints nothing, distinct from an absent `BorderSide`
 * (unset — falls back to whatever the renderer/writer defaults to). */
export type BorderLineStyle = "none" | "single" | "double" | "dashed" | "dotted";

/** One edge of a cell's border (SPEC-023 AC-10..12): style, width in points
 * and a CSS-ready colour ("#rrggbb", or "auto" for the source's automatic
 * colour). */
export interface BorderSide {
  style: BorderLineStyle;
  widthPt: number;
  color: string;
}

/** A cell's four borders, each independently set or left unset (SPEC-023).
 * Table-wide and per-selected-cell border operations (AC-13, AC-14) both
 * write into this same per-cell shape — "apply to the whole table" is just
 * "apply to every cell" (see lib/tables.ts's border-preset commands), which
 * is also why the round trip is always per-cell (`w:tcBorders` /
 * .odt cell styles — AC-25, AC-29), never a table-level default. */
export interface CellBorders {
  top?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  right?: BorderSide;
}

/** A cell's content vertical alignment within its (possibly taller, due to
 * sibling cells) row (SPEC-023 AC-17). */
export type CellVerticalAlign = "top" | "center" | "bottom";

/** Table alignment on the page (SPEC-023 AC-22). */
export type TableAlignment = "left" | "center" | "right";

/** How a table's overall width is determined (SPEC-023 AC-23): `"fixed"`
 * (absent also means fixed, for models saved before SPEC-023) uses the
 * `columnWidthsCm` sum as-is; `"percent"` scales `widthPercent`% of the
 * text area, redistributing `columnWidthsCm` proportionally; `"auto"` sizes
 * the table to its content, ignoring `columnWidthsCm` for layout (the field
 * is kept as the last explicit widths, reused if the mode changes back). */
export type TableWidthMode = "auto" | "fixed" | "percent";

/** Typed pass-through for a border line style coming from a `.vue`
 * template's plain string (espree bars the annotation in `.vue` scripts). */
export function asBorderLineStyle(value: string): BorderLineStyle {
  return value as BorderLineStyle;
}

/** Typed pass-through for a cell vertical alignment coming from a `.vue`
 * template's plain string (same espree constraint as asBorderLineStyle). */
export function asCellVerticalAlign(value: string): CellVerticalAlign {
  return value as CellVerticalAlign;
}

/** Typed pass-through for a table alignment coming from a `.vue` template's
 * plain string (same espree constraint as asBorderLineStyle). */
export function asTableAlignment(value: string): TableAlignment {
  return value as TableAlignment;
}

/** Typed pass-through for a table width mode coming from a `.vue`
 * template's plain string (same espree constraint as asBorderLineStyle). */
export function asTableWidthMode(value: string): TableWidthMode {
  return value as TableWidthMode;
}

// --- grid limits (SPEC-022) --------------------------------------------------

/** The picker grid's bounds (AC-1) and the dialog's upper limits (AC-2). */
export const PICKER_MAX_ROWS = 10;
export const PICKER_MAX_COLS = 10;
export const MAX_TABLE_ROWS = 200;
export const MAX_TABLE_COLS = 63;

/** The smallest width a dragged column boundary can leave a column (AC-2). */
export const MIN_COLUMN_WIDTH_CM = 0.5;

/** The smallest minimum height a dragged row boundary can set (AC-4). */
export const MIN_ROW_HEIGHT_CM = 0.4;

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

// --- border presets (SPEC-023 AC-18..21) -------------------------------------

/** The standard border presets a table properties dialog offers (AC-18..21). */
export type BorderPreset = "all" | "outer" | "inner" | "none";

/** Typed pass-through for a border preset coming from a `.vue` template's
 * plain string (same espree constraint as asBorderLineStyle). */
export function asBorderPreset(value: string): BorderPreset {
  return value as BorderPreset;
}

// --- live resize preview (SPEC-023 AC-1..5) ----------------------------------
// DOM helpers for the column/row resize preview. DOM-touching but
// model-free: they paint whatever widths/heights they are given into a live
// table element (clamping is the caller's job).

/** One row's `<tr>` element by index, or null out of range. */
export function tableRowElement(tableEl: HTMLTableElement, rowIndex: number): HTMLTableRowElement | null {
  const row = tableEl.querySelectorAll(":scope > tbody > tr")[rowIndex];
  return row instanceof HTMLTableRowElement ? row : null;
}

/** Live-previews a column resize (AC-1) directly on the DOM: every `<col>`'s
 * width, and the table's own width when it isn't sized by its content
 * (`table-layout: auto`, SPEC-023 AC-23's "auto" width mode) — clamping to
 * the minimum width is the caller's job (MIN_COLUMN_WIDTH_CM); this just
 * paints whatever it is given. */
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
