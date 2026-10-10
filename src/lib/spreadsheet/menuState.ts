// Shape helper for the workbook's menu bar (SPEC-117 "The menu bar") — the
// same move as presentation/menuState.ts: SpreadsheetMenuBar takes its state
// as one runtime `Object` prop (the espree constraint on .vue scripts bars a
// typed PropType there), so it coerces that prop through here. SPEC-121 adds
// the Edit menu's undo/redo state and File ▸ "Download a copy as …" (its
// label names the other format; a reason disables it). Later spreadsheet
// specs append fields as they append menus and items.

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
}

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
  };
}
