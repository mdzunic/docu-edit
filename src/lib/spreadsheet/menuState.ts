// Shape helper for the workbook's menu bar (SPEC-117 "The menu bar") — the
// same move as presentation/menuState.ts: SpreadsheetMenuBar takes its state
// as one runtime `Object` prop (the espree constraint on .vue scripts bars a
// typed PropType there), so it coerces that prop through here. Later
// spreadsheet specs append fields as they append menus and items.

/** What the bar needs in order to render: what is checked. */
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
} as const;

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
 * 100%. */
export function asSpreadsheetMenuState(value: unknown): SpreadsheetMenuState {
  const state = (value ?? {}) as Partial<SpreadsheetMenuState>;
  return {
    gridLines: state.gridLines !== false,
    showFormulas: state.showFormulas === true,
    zoomPercent: typeof state.zoomPercent === "number" ? state.zoomPercent : 100,
    hasGrid: state.hasGrid !== false,
  };
}
