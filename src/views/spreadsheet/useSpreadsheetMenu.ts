// Routes the workbook menu bar's commands (SPEC-117 "The menu bar") — the
// role usePresentationMenu plays for the deck: SpreadsheetMenuBar emits a
// command id, this maps it onto the composable that owns the action, and it
// builds the bar's state (what is checked). Download and Close are the app
// bar's own actions (useWorkbookLoad); Gridlines writes the sheet's
// `showGridLines` as view state (useSheetViewState); Show formulas toggles
// the sheet's `showFormulas` (useFormulaDisplay, SPEC-119); the zoom items
// step or set the sheet's zoom (useSheetZoom). Later specs append commands
// here.
import { computed } from "vue";
import { SHEET_MENU_COMMANDS, zoomPresetOf, type SpreadsheetMenuState } from "../../lib/spreadsheet/menuState";
import type { FormulaDisplay } from "./useFormulaDisplay";
import type { SheetViewState } from "./useSheetViewState";
import type { SheetZoom } from "./useSheetZoom";

export interface SpreadsheetMenuDeps {
  book: { download(): void; close(): void };
  viewState: Pick<SheetViewState, "active" | "activeSheet" | "updateActive">;
  zoom: Pick<SheetZoom, "percent" | "set" | "zoomIn" | "zoomOut">;
  /** View → Show formulas (SPEC-119); absent: the item never shows checked. */
  formulas?: Pick<FormulaDisplay, "showFormulas" | "toggle">;
}

export function useSpreadsheetMenu(deps: SpreadsheetMenuDeps) {
  const gridLines = computed(() => deps.viewState.active.value.showGridLines !== false);

  const state = computed(
    (): SpreadsheetMenuState => ({
      gridLines: gridLines.value,
      showFormulas: deps.formulas?.showFormulas.value === true,
      zoomPercent: deps.zoom.percent.value,
      hasGrid: deps.viewState.activeSheet.value?.kind === "grid",
    }),
  );

  /** View → Gridlines: shown ↔ hidden, as view state. */
  function toggleGridLines(): void {
    deps.viewState.updateActive({ showGridLines: gridLines.value ? false : undefined });
  }

  function run(command = ""): void {
    const preset = zoomPresetOf(command);
    if (preset !== null) {
      deps.zoom.set(preset);
      return;
    }
    switch (command) {
      case SHEET_MENU_COMMANDS.download:
        deps.book.download();
        break;
      case SHEET_MENU_COMMANDS.close:
        deps.book.close();
        break;
      case SHEET_MENU_COMMANDS.gridLines:
        toggleGridLines();
        break;
      case SHEET_MENU_COMMANDS.showFormulas:
        deps.formulas?.toggle();
        break;
      case SHEET_MENU_COMMANDS.zoomIn:
        deps.zoom.zoomIn();
        break;
      case SHEET_MENU_COMMANDS.zoomOut:
        deps.zoom.zoomOut();
        break;
      default:
        break;
    }
  }

  return { state, gridLines, toggleGridLines, run };
}
