// Shape helper for the workbook's menu bar (SPEC-117 "The menu bar") — the
// same move as presentation/menuState.ts: SpreadsheetMenuBar takes its state
// as one runtime `Object` prop (the espree constraint on .vue scripts bars a
// typed PropType there), so it coerces that prop through here. SPEC-121 adds
// the Edit menu's undo/redo state and File ▸ "Download a copy as …" (its
// label names the other format; a reason disables it). SPEC-124 adds View ▸
// Freeze's state and the Format menu's rows-and-columns commands; SPEC-125
// Edit ▸ Paste special… and Fill ▸ Down/Right; SPEC-128 the Data menu's
// sorts and filter commands and whether the sheet has an autofilter (Filter
// is checked; without one Clear filters and Reapply filter are disabled,
// saying "Turn on Filter first"). SPEC-130 adds Data ▸ "Data validation…"
// (disabled with the other grid items on a foreign sheet). SPEC-131 adds the
// Insert menu (Image…, Shapes — disabled where the workbook cannot hold
// drawings: a CSV or TSV file, a foreign sheet) and the Arrange menu (the four
// z-order moves and Alt Text…, disabled without a selected drawing; Bring to
// Front/Forward at the front, Send Backward/to Back at the back). Later
// spreadsheet specs append fields as they append menus and items.

/** What the bar needs in order to render: what is checked, what applies. */
export interface SpreadsheetMenuState {
  /** View → Gridlines, checked from the sheet's `showGridLines`. */
  gridLines: boolean;
  /** View → Show formulas (SPEC-119), checked from the sheet's
   * `showFormulas`. */
  showFormulas: boolean;
  /** The active sheet's zoom — its preset (if any) shows checked. */
  zoomPercent: number;
  /** Whether a grid sheet is on show (the View items need one). */
  hasGrid: boolean;
  /** Edit → Undo / Redo (SPEC-121's useSpreadsheetUndo). */
  canUndo: boolean;
  canRedo: boolean;
  /** The step the next undo/redo would report ("Rename sheet"); "" for none. */
  undoLabel: string;
  redoLabel: string;
  /** File → "Download a copy as …": the item's label, naming the other
   * format ("Download a copy as Excel workbook (.xlsx)…"). */
  copyLabel: string;
  /** Why the copy cannot be made yet; "" when it can. */
  copyDisabledReason: string;
  /** SPEC-123: Format ▸ Cell styles' names (the gallery's, in its order). */
  cellStyleNames: string[];
  /** SPEC-124: View ▸ Freeze — Unfreeze is enabled only when the sheet is
   * frozen; "Freeze at current cell" is disabled at A1. */
  frozen: boolean;
  freezeAtCellDisabled: boolean;
  /** SPEC-128: the active sheet has an autofilter (Data ▸ Filter checked). */
  filterOn: boolean;
  /** SPEC-131: Insert ▸ Image… and Shapes apply (an .xlsx/.ods grid sheet). */
  canInsertDrawings: boolean;
  /** SPEC-131: a drawing is selected (the Arrange menu applies), and whether
   * it is already at the front or the back of the z-order. */
  drawingSelected: boolean;
  drawingAtFront: boolean;
  drawingAtBack: boolean;
}

/** SPEC-128: why Clear filters and Reapply filter are disabled. */
export const FILTER_OFF_REASON = "Turn on Filter first";

/** The command ids the bar emits — routed by useSpreadsheetMenu. Item test
 * ids keep the `sheet-menu-` prefix. */
export const SHEET_MENU_COMMANDS = {
  download: "file.download",
  close: "file.close",
  gridLines: "view.gridlines",
  showFormulas: "view.show-formulas",
  zoomIn: "view.zoom-in",
  zoomOut: "view.zoom-out",
  /** SPEC-120: Data ▸ Recalculate (F9). */
  recalculate: "data.recalculate",
  /** SPEC-121: Edit ▸ Undo / Redo (the registry's edit.undo / edit.redo). */
  undo: "edit.undo",
  redo: "edit.redo",
  /** SPEC-121: File ▸ Download a copy as … the other format. */
  downloadCopy: "file.download-copy",
  /** SPEC-123: Format ▸ Format cells…, Number format ▸ More formats… and
   * Clear ▸ (numberFormatCommand / cellStyleCommand name the rest). */
  formatCells: "format.cells",
  moreNumberFormats: "format.number-more",
  clearAll: "format.clear-all",
  clearFormats: "format.clear-formats",
  clearContents: "format.clear-contents",
  /** SPEC-124: View ▸ Freeze ▸ and the Format menu's rows and columns. */
  freezeTopRow: "view.freeze-top-row",
  freezeFirstColumn: "view.freeze-first-column",
  freezeAtCell: "view.freeze-at-cell",
  unfreeze: "view.unfreeze",
  rowHeight: "format.row-height",
  columnWidth: "format.column-width",
  autofitRows: "format.autofit-rows",
  autofitColumns: "format.autofit-columns",
  /** SPEC-125: Edit ▸ Paste special… and Edit ▸ Fill ▸ Down/Right. */
  pasteSpecial: "edit.paste-special",
  fillDown: "edit.fill-down",
  fillRight: "edit.fill-right",
  /** SPEC-128: Data ▸ Sort A to Z / Z to A / Sort range…, Filter, Clear
   * filters and Reapply filter. */
  sortAscending: "data.sort-asc",
  sortDescending: "data.sort-desc",
  sortRange: "data.sort-range",
  filter: "data.filter",
  clearFilters: "data.clear-filters",
  reapplyFilter: "data.reapply-filter",
  /** SPEC-129: Format ▸ Conditional formatting… (the rules manager pane). */
  conditionalFormatting: "format.conditional-formatting",
  /** SPEC-130: Data ▸ Data validation… (the validation dialog). */
  dataValidation: "data.validation",
  /** SPEC-131: Insert ▸ Image… and Insert ▸ Shapes; the Arrange menu. */
  insertImage: "insert.image",
  insertShape: "insert.shape",
  /** SPEC-132: Insert ▸ Chart (the chart gallery popover). */
  insertChart: "insert.chart",
  /** SPEC-133: Insert ▸ Table… (the Create Table dialog; no shortcut —
   * Ctrl/Cmd+T is the browser's). */
  insertTable: "insert.table",
  arrangeFront: "arrange.front",
  arrangeForward: "arrange.forward",
  arrangeBackward: "arrange.backward",
  arrangeBack: "arrange.back",
  altText: "arrange.alt-text",
  /** SPEC-133: Formulas ▸ Name Manager… (Ctrl/Cmd+F3). */
  nameManager: "formulas.name-manager",
} as const;

/** `format.number.<preset>`: Format ▸ Number format ▸ a preset (SPEC-123). */
export function numberFormatCommand(preset: string): string {
  return `format.number.${preset}`;
}

/** `format.style.<name>`: Format ▸ Cell styles ▸ a style (SPEC-123). */
export function cellStyleCommand(name: string): string {
  return `format.style.${name}`;
}

/** The copy item's label until the view names the other format. */
export const DEFAULT_COPY_LABEL = "Download a copy as …";

/** "Undo" / "Redo", plus the step when there is one ("Undo Rename sheet"). */
export function undoRedoMenuLabel(verb: "Undo" | "Redo", step: string): string {
  return step ? `${verb} ${step}` : verb;
}

/** `view.zoom-<n>`: one rung of the ZOOM_PRESETS ladder. */
export function zoomPresetCommand(percent: number): string {
  return `view.zoom-${percent}`;
}

/** The preset a `view.zoom-<n>` command names, or null. */
export function zoomPresetOf(command: string): number | null {
  const match = /^view\.zoom-(\d+)$/.exec(command);
  return match ? Number(match[1]) : null;
}

/** Coerces the untyped `state` prop into the typed shape the template reads;
 * a bar mounted without state shows gridlines on, values (not formulas) and
 * 100%, with nothing to undo or redo. */
export function asSpreadsheetMenuState(value: unknown): SpreadsheetMenuState {
  const state = (value ?? {}) as Partial<SpreadsheetMenuState>;
  return {
    gridLines: state.gridLines !== false,
    showFormulas: state.showFormulas === true,
    zoomPercent: typeof state.zoomPercent === "number" ? state.zoomPercent : 100,
    hasGrid: state.hasGrid !== false,
    canUndo: state.canUndo === true,
    canRedo: state.canRedo === true,
    undoLabel: typeof state.undoLabel === "string" ? state.undoLabel : "",
    redoLabel: typeof state.redoLabel === "string" ? state.redoLabel : "",
    copyLabel: typeof state.copyLabel === "string" && state.copyLabel !== "" ? state.copyLabel : DEFAULT_COPY_LABEL,
    copyDisabledReason: typeof state.copyDisabledReason === "string" ? state.copyDisabledReason : "",
    cellStyleNames: Array.isArray(state.cellStyleNames) ? state.cellStyleNames : [],
    frozen: state.frozen === true,
    freezeAtCellDisabled: state.freezeAtCellDisabled === true,
    filterOn: state.filterOn === true,
    canInsertDrawings: state.canInsertDrawings !== false,
    drawingSelected: state.drawingSelected === true,
    drawingAtFront: state.drawingAtFront === true,
    drawingAtBack: state.drawingAtBack === true,
  };
}
