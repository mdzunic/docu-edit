// Routes the workbook menu bar's commands (SPEC-117 "The menu bar") — the
// role usePresentationMenu plays for the deck: SpreadsheetMenuBar emits a
// command id, this maps it onto the composable that owns the action, and it
// builds the bar's state (what is checked). Download and Close are the app
// bar's own actions (useWorkbookLoad); Gridlines writes the sheet's
// `showGridLines` as view state (useSheetViewState); Show formulas toggles
// the sheet's `showFormulas` (useFormulaDisplay, SPEC-119); the zoom items
// step or set the sheet's zoom (useSheetZoom); Data ▸ Recalculate runs a full
// recalculation (useRecalc, SPEC-120). SPEC-121: Edit ▸ Undo/Redo step the
// workbook's undo stack, and File ▸ Download and "Download a copy as …" are
// useWorkbookDownload's (through useWorkbookEditing) — the bar's state carries
// their enabled states and labels. Later specs append commands here.
import { computed } from "vue";
import { SHEET_MENU_COMMANDS, zoomPresetOf, type SpreadsheetMenuState } from "../../lib/spreadsheet/menuState";
import type { FormulaDisplay } from "./useFormulaDisplay";
import type { Recalc } from "./useRecalc";
import type { SheetViewState } from "./useSheetViewState";
import type { SheetZoom } from "./useSheetZoom";
import type { WorkbookEditing } from "./useWorkbookEditing";

export interface SpreadsheetMenuDeps {
  book: { download(): void; close(): void };
  viewState: Pick<SheetViewState, "active" | "activeSheet" | "updateActive">;
  zoom: Pick<SheetZoom, "percent" | "set" | "zoomIn" | "zoomOut">;
  /** View → Show formulas (SPEC-119); absent: the item never shows checked. */
  formulas?: Pick<FormulaDisplay, "showFormulas" | "toggle">;
  /** Data → Recalculate (SPEC-120). */
  recalc?: Pick<Recalc, "recalculate">;
  /** SPEC-121: Edit → Undo/Redo and File → Download / "Download a copy as …". */
  editing?: Pick<WorkbookEditing, "undo" | "download" | "appBar">;
}

export function useSpreadsheetMenu(deps: SpreadsheetMenuDeps) {
  const gridLines = computed(() => deps.viewState.active.value.showGridLines !== false);

  const state = computed(
    (): SpreadsheetMenuState => ({
      gridLines: gridLines.value,
      showFormulas: deps.formulas?.showFormulas.value === true,
      zoomPercent: deps.zoom.percent.value,
      hasGrid: deps.viewState.activeSheet.value?.kind === "grid",
      ...(deps.editing?.appBar.value ?? {
        canUndo: false,
        canRedo: false,
        undoLabel: "",
        redoLabel: "",
        copyLabel: "Download a copy as …",
        copyDisabledReason: "",
      }),
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
        if (deps.editing) {
          deps.editing.download.download();
        } else {
          deps.book.download();
        }
        break;
      case SHEET_MENU_COMMANDS.downloadCopy:
        deps.editing?.download.downloadCopy();
        break;
      case SHEET_MENU_COMMANDS.undo:
        deps.editing?.undo.undo();
        break;
      case SHEET_MENU_COMMANDS.redo:
        deps.editing?.undo.redo();
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
      case SHEET_MENU_COMMANDS.recalculate:
        deps.recalc?.recalculate();
        break;
      default:
        break;
    }
  }

  return { state, gridLines, toggleGridLines, run };
}
