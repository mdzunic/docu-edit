// Cell-style resolution (SPEC-117 "Model additions"): which style a cell
// shows and the legacy indexed palette both readers and renderers share.
// Style resolution is ONE function, effectiveStyleId — the cell's own
// `styleId`, else its row's (present only when the source row carried a
// format), else the column span covering it, else the default style 0 —
// each step a binary search over sparse, sorted data, never a scan.
import { findRow, getCell } from "./address";
import type { CellStyle, ColumnSpan, Sheet, SpreadsheetModel } from "./model";

/** The default style table: one empty ("Normal") style. */
export function defaultStyles(): CellStyle[] {
  return [{}];
}

/** ECMA-376 §18.8.27's 64-entry legacy palette (lowercase), the colour an
 * `indexed="n"` attribute names unless the file overrides it with
 * `<colors><indexedColors>`. 64 is "system foreground" (automatic) and 65
 * "system background" — the readers map those, not this table. */
export const INDEXED_COLORS: readonly string[] = [
  "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff",
  "#000000", "#ffffff", "#ff0000", "#00ff00", "#0000ff", "#ffff00", "#ff00ff", "#00ffff",
  "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#c0c0c0", "#808080",
  "#9999ff", "#993366", "#ffffcc", "#ccffff", "#660066", "#ff8080", "#0066cc", "#ccccff",
  "#000080", "#ff00ff", "#ffff00", "#00ffff", "#800080", "#800000", "#008080", "#0000ff",
  "#00ccff", "#ccffff", "#ccffcc", "#ffff99", "#99ccff", "#ff99cc", "#cc99ff", "#ffcc99",
  "#3366ff", "#33cccc", "#99cc00", "#ffcc00", "#ff9900", "#ff6600", "#666699", "#969696",
  "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
];

/** The column span covering `col` — a binary search over the sorted,
 * non-overlapping spans; undefined when the column has the sheet default. */
export function columnSpanAt(columns: readonly ColumnSpan[] | undefined, col: number): ColumnSpan | undefined {
  if (!columns) {
    return undefined;
  }
  let lo = 0;
  let hi = columns.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    const span = columns[mid]!;
    if (col < span.min) {
      hi = mid - 1;
    } else if (col > span.max) {
      lo = mid + 1;
    } else {
      return span;
    }
  }
  return undefined;
}

/** The style a position shows: the cell's `styleId`, else its row's
 * `styleId`, else the covering column span's, else 0. */
export function effectiveStyleId(sheet: Sheet, row: number, col: number): number {
  const cellStyleId = getCell(sheet, row, col)?.styleId;
  if (cellStyleId !== undefined) {
    return cellStyleId;
  }
  const rowStyleId = findRow(sheet, row)?.styleId;
  if (rowStyleId !== undefined) {
    return rowStyleId;
  }
  return columnSpanAt(sheet.columns, col)?.styleId ?? 0;
}

/** `model.styles[id]`, or the default style for an id out of range. */
export function cellStyle(model: Pick<SpreadsheetModel, "styles">, id: number): CellStyle {
  const styles = model.styles;
  if (!styles || styles.length === 0) {
    return {};
  }
  return (Number.isInteger(id) && id >= 0 ? styles[id] : undefined) ?? styles[0]!;
}
