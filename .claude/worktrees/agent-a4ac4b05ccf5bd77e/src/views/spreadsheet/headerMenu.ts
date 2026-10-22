// The header menu's model (SPEC-124 "Header menu"): which bands of whole
// rows or columns a selection holds, the band containing a header, and the
// entries SheetHeaderContextMenu shows — their labels, test ids and
// enabled states. Framework-free so useStructure stays a wiring of commands
// (and under its line budget).
//
// | Entry | testid | Label (rows; columns analogous) |
// | Insert above / left | sheet-insert-before | "Insert 1 row above", "Insert 3 rows above" / "Insert 1 column left" |
// | Insert below / right | sheet-insert-after | "Insert 1 row below" / "Insert 2 columns right" |
// | Delete | sheet-delete-band | "Delete row", "Delete rows 4–6" / "Delete columns B–D" |
// | Hide | sheet-hide-band | "Hide row", "Hide rows 4–6" |
// | Unhide | sheet-unhide-band | "Unhide rows" / "Unhide columns" (enabled when the bands hold a hidden one) |
// | Size | sheet-size-band | "Row height…" / "Column width…" |
// | Autofit | sheet-autofit-band | "Autofit row height" / "Autofit column width" |
// N is the number of selected rows or columns in the band containing the
// clicked header; a selection of several bands deletes or hides them all.
import { colToLetters, type GridRange } from "../../lib/spreadsheet/address";
import type { Band, BandAxis } from "../../lib/spreadsheet/bands";
import { normaliseBands } from "../../lib/spreadsheet/bands";
import { MAX_COLS, MAX_ROWS, type Sheet } from "../../lib/spreadsheet/model";
import type { SheetSelection } from "../../lib/spreadsheet/selection";

export type HeaderMenuCommand = "insert-before" | "insert-after" | "delete" | "hide" | "unhide" | "size" | "autofit";

export interface HeaderMenuItem {
  command: HeaderMenuCommand;
  label: string;
  testid: string;
  disabled: boolean;
}

/** Whether a DOM target is inside the grid and outside any dialog. */
export function inGrid(target: EventTarget | null): boolean {
  return target instanceof Element && target.closest('[data-testid="spreadsheet-grid"]') !== null && target.closest('[role="dialog"]') === null;
}

/** The row or column header a DOM target is on, from its data-index. */
export function headerAt(target: EventTarget | null): { axis: BandAxis; index: number } | null {
  const header = target instanceof Element ? target.closest<HTMLElement>('[data-testid="sheet-row-header"],[data-testid="sheet-col-header"]') : null;
  const index = Number(header?.dataset.index);
  if (!header || !Number.isInteger(index)) {
    return null;
  }
  return { axis: header.dataset.testid === "sheet-row-header" ? "row" : "col", index };
}

/** Whether a range is whole rows (every column) or whole columns. */
export function rangeAxis(range: GridRange): BandAxis | null {
  if (range.left === 0 && range.right === MAX_COLS - 1) {
    return "row";
  }
  if (range.top === 0 && range.bottom === MAX_ROWS - 1) {
    return "col";
  }
  return null;
}

/** The axis every range of a selection is whole along, or null — the
 * selection the header menu and the insert/delete shortcuts act on
 * directly. Select-all is whole rows. */
export function selectionAxis(selection: SheetSelection): BandAxis | null {
  const axes = selection.ranges.map(rangeAxis);
  if (axes.length === 0 || axes.some((axis) => axis === null)) {
    return null;
  }
  return axes.every((axis) => axis === "row") ? "row" : axes.every((axis) => axis === "col") ? "col" : null;
}

const bandOf = (range: GridRange, axis: BandAxis): Band => (axis === "row" ? { start: range.top, end: range.bottom } : { start: range.left, end: range.right });

/** The selection's bands along `axis` (every range's rows or columns),
 * sorted and merged. */
export function selectionBands(selection: SheetSelection, axis: BandAxis): Band[] {
  return normaliseBands(
    selection.ranges.map((range) => bandOf(range, axis)),
    axis,
  );
}

/** The band of the selection containing `index`, or the index alone. */
export function bandContaining(selection: SheetSelection, axis: BandAxis, index: number): Band {
  return selectionBands(selection, axis).find((band) => index >= band.start && index <= band.end) ?? { start: index, end: index };
}

/** Whether the selection's whole bands along `axis` contain `index`. */
export function bandSelected(selection: SheetSelection, axis: BandAxis, index: number): boolean {
  return selectionAxis(selection) === axis && selectionBands(selection, axis).some((band) => index >= band.start && index <= band.end);
}

/** Whether any row or column of `bands` is hidden. */
export function bandsHoldHidden(sheet: Sheet, axis: BandAxis, bands: readonly Band[]): boolean {
  if (axis === "col") {
    return sheet.columns.some((span) => span.hidden && bands.some((band) => span.min <= band.end && span.max >= band.start));
  }
  return sheet.rows.some((row) => row.hidden && bands.some((band) => row.index >= band.start && row.index <= band.end));
}

const plural = (count: number, word: string): string => `${count} ${word}${count === 1 ? "" : "s"}`;
const name = (axis: BandAxis, index: number): string => (axis === "row" ? String(index + 1) : colToLetters(index));

/** "row" / "rows 4–6" / "columns B–D" — what a delete or hide names. */
export function bandsLabel(axis: BandAxis, bands: readonly Band[]): string {
  const word = axis === "row" ? "row" : "column";
  if (bands.length !== 1) {
    return `${word}s`;
  }
  const band = bands[0]!;
  return band.start === band.end ? word : `${word}s ${name(axis, band.start)}–${name(axis, band.end)}`;
}

/** The header menu's entries (see the module header). */
export function headerMenuItems(axis: BandAxis, clicked: Band, bands: readonly Band[], anyHidden: boolean): HeaderMenuItem[] {
  const count = clicked.end - clicked.start + 1;
  const rows = axis === "row";
  const unit = rows ? "row" : "column";
  return [
    { command: "insert-before", label: `Insert ${plural(count, unit)} ${rows ? "above" : "left"}`, testid: "sheet-insert-before", disabled: false },
    { command: "insert-after", label: `Insert ${plural(count, unit)} ${rows ? "below" : "right"}`, testid: "sheet-insert-after", disabled: false },
    { command: "delete", label: `Delete ${bandsLabel(axis, bands)}`, testid: "sheet-delete-band", disabled: false },
    { command: "hide", label: `Hide ${bandsLabel(axis, bands)}`, testid: "sheet-hide-band", disabled: false },
    { command: "unhide", label: rows ? "Unhide rows" : "Unhide columns", testid: "sheet-unhide-band", disabled: !anyHidden },
    { command: "size", label: rows ? "Row height…" : "Column width…", testid: "sheet-size-band", disabled: false },
    { command: "autofit", label: rows ? "Autofit row height" : "Autofit column width", testid: "sheet-autofit-band", disabled: false },
  ];
}
