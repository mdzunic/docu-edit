// Where and how the in-cell editor shows (SPEC-122 "Editor DOM"): over the
// active cell (its merge's range), in the paper's unscaled px, at least the
// cell's size and free to grow right to the viewport's edge (and down as
// lines grow) — a wrapping cell keeps its width and only grows down — in the
// cell's resolved font and colour on its fill (the paper's white without
// one), as SheetCellLayer paints the cell. Also its accessible name: the
// address and the value as shown ("B3, Revenue"; just "B3" when empty).
import { formatA1, getCell, type CellRef } from "../../lib/spreadsheet/address";
import { cellCss, cellDisplay, PATTERN_BACKGROUND } from "../../lib/spreadsheet/cellRender";
import type { DocumentTheme } from "../../lib/core/theme";
import type { CellStyle, Sheet } from "../../lib/spreadsheet/model";
import { plainText, type FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import type { RectItem } from "./useSheetOverlays";

/** The paper's default cell font (SheetCellLayer's `.sheet-cell`). */
const DEFAULT_FONT = { "font-family": "Calibri, Carlito, sans-serif", "font-size": "11pt", "line-height": "1.2" };
const FONT_KEYS = ["font-family", "font-size", "font-weight", "font-style", "text-decoration", "color"] as const;

export interface CellEditorView {
  boxStyle: Record<string, string>;
  textStyle: Record<string, string>;
  label: string;
}

/** The editor's box, text style and name for the active cell. `rect` is the
 * cell's on-screen rectangle (null when scrolled off: the box parks at the
 * paper's origin, still mounted and focusable); `room` the paper's width. */
export function cellEditorView(input: {
  sheet: Sheet;
  active: CellRef;
  style: CellStyle;
  rect: RectItem | null;
  room: number;
  theme: DocumentTheme;
  format: FormatContext;
}): CellEditorView {
  const { rect, style } = input;
  const left = rect?.left ?? 0;
  const width = rect?.width ?? 0;
  const css = cellCss(style, input.theme);
  const wraps = style.alignment?.wrap === true;
  const boxStyle: Record<string, string> = {
    left: `${left}px`,
    top: `${rect?.top ?? 0}px`,
    width: `${width}px`,
    height: `${rect?.height ?? 0}px`,
    "max-width": `${Math.max(width, input.room - left)}px`,
  };
  const textStyle: Record<string, string> = { ...DEFAULT_FONT, background: css.box["background-color"] ?? PATTERN_BACKGROUND };
  for (const key of FONT_KEYS) {
    const value = css.text[key];
    if (value !== undefined) {
      textStyle[key] = value;
    }
  }
  if (wraps) {
    textStyle.width = `${width}px`;
  }
  const cell = getCell(input.sheet, input.active.row, input.active.col);
  const shown = cell ? plainText(cellDisplay(cell.value, style, input.sheet, input.format)) : "";
  const address = formatA1(input.active);
  return { boxStyle, textStyle, label: shown === "" ? address : `${address}, ${shown}` };
}
