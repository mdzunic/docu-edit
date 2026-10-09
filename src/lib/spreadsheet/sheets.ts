// Sheet commands (SPEC-121 "Sheet commands") — THIS FILE is the normative
// source for creating a workbook and for the workbook-level sheet commands:
// insert, rename, duplicate, delete, move, hide/unhide and tab colour, the
// sheet-name rules, and the cross-format copy (`stripRetention`,
// `unconvertibleContent`). Framework-free: the view's useSheetCommands drives
// these through the undo stack (record → command → commit).
//
// Every command keeps SPEC-116's invariants and SPEC-117's: a formula rewrite
// builds new Cell/Row objects and a new `rows` array for the sheets it
// touches; a changed sheet is replaced by a new object, never mutated;
// `model.sheets`, `opaque` and `odsContentEnvelope` are replaced, not spliced
// in place, when they change. Every command except setTabColor on an
// unchanged colour sets `structureDirty`. Indices are into `model.sheets`
// (hidden sheets included); the UI maps visible tab positions onto them. The
// active sheet is always visible.
import { DEFAULT_THEME } from "../core/theme";
import { DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PT } from "./geometry";
import { DEFAULT_SPREADSHEET_TITLE, type Sheet, type SpreadsheetModel } from "./model";
import { SPREADSHEET_READER_VERSION } from "./open";

/** A new, empty grid sheet — born dirty, with no source. */
export function emptySheet(name: string): Sheet {
  return {
    id: crypto.randomUUID(),
    name,
    visibility: "visible",
    kind: "grid",
    rows: [],
    columns: [],
    merges: [],
    defaultColWidthPx: DEFAULT_COL_WIDTH_PX,
    defaultRowHeightPt: DEFAULT_ROW_HEIGHT_PT,
    dirty: true,
  };
}

/** A new workbook (SPEC-121 "The blank workbook"): one visible grid sheet
 * "Sheet1", `.xlsx` from birth, the 1900 date system, the single default
 * style (Calibri through the theme's minor font, 11 pt, dk1) and the Office
 * theme. No `properties`: the from-scratch core.xml carries the title and
 * `modified` only. */
export function createBlankWorkbook(): SpreadsheetModel {
  return {
    kind: "spreadsheet",
    id: crypto.randomUUID(),
    title: DEFAULT_SPREADSHEET_TITLE,
    format: "xlsx",
    readerVersion: SPREADSHEET_READER_VERSION,
    sheets: [emptySheet("Sheet1")],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{ font: { family: { theme: "minor" }, sizePt: 11, color: { theme: "dk1" } }, namedStyle: "Normal" }],
    defaultStyleId: 0,
    theme: structuredClone(DEFAULT_THEME),
  };
}
