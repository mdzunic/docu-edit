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
// their enabled states and labels. SPEC-123: the Format menu runs
// useFormatting's commands. SPEC-124: View ▸ Freeze and the Format menu's
// row height, column width and autofits are useStructure's. SPEC-125: Edit ▸
// Paste special… is useSheetClipboard's, Edit ▸ Fill ▸ Down/Right
// useFillHandle's. SPEC-128: the Data menu's sorts and filter commands are
// useSortFilter's, and the bar's state says whether Filter is on. Later specs
// append commands here.
import { computed } from "vue";
import { SHEET_MENU_COMMANDS, zoomPresetOf, type SpreadsheetMenuState } from "../../lib/spreadsheet/menuState";
import type { NumberPresetId } from "../../lib/spreadsheet/numberCodes";
import type { Formatting } from "./useFormatting";
import type { FormulaDisplay } from "./useFormulaDisplay";
import type { FillHandle } from "./useFillHandle";
import type { Recalc } from "./useRecalc";
import type { SheetClipboard } from "./useSheetClipboard";
import type { SortFilter } from "./useSortFilter";
import type { SheetViewState } from "./useSheetViewState";
import type { SheetZoom } from "./useSheetZoom";
import type { Structure } from "./useStructure";
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
  /** SPEC-123: the Format menu. */
  formatting?: Pick<Formatting, "run" | "toolbar">;
  /** SPEC-124: View ▸ Freeze and the Format menu's rows and columns. */
  structure?: Pick<Structure, "freeze" | "freezeState" | "formatBands">;
  /** SPEC-125: Edit ▸ Paste special… and Edit ▸ Fill ▸ Down/Right. */
  clipboard?: Pick<SheetClipboard, "openPasteSpecial">;
  fillHandle?: Pick<FillHandle, "fillDown" | "fillRight">;
  /** SPEC-128: the Data menu's sorts and filter commands. */
  sortFilter?: Pick<SortFilter, "sortAscending" | "sortDescending" | "openSortDialog" | "toggleFilter" | "clearFilters" | "reapplyFilter" | "filterOn">;
}

/** SPEC-128's Data menu commands → useSortFilter's. */
const SORT_FILTER_COMMANDS: Readonly<Record<string, "sortAscending" | "sortDescending" | "openSortDialog" | "toggleFilter" | "clearFilters" | "reapplyFilter">> = {
  [SHEET_MENU_COMMANDS.sortAscending]: "sortAscending",
  [SHEET_MENU_COMMANDS.sortDescending]: "sortDescending",
  [SHEET_MENU_COMMANDS.sortRange]: "openSortDialog",
  [SHEET_MENU_COMMANDS.filter]: "toggleFilter",
  [SHEET_MENU_COMMANDS.clearFilters]: "clearFilters",
  [SHEET_MENU_COMMANDS.reapplyFilter]: "reapplyFilter",
};

/** A Format menu command as useFormatting's (null: not one). */
function formatCommand(command: string): Parameters<Formatting["run"]>[0] | null {
  if (command.startsWith("format.number.")) {
    return { kind: "number-format", preset: command.slice("format.number.".length) as NumberPresetId };
  }
  if (command.startsWith("format.style.")) {
    return { kind: "cell-style", name: command.slice("format.style.".length) };
  }
  switch (command) {
    case SHEET_MENU_COMMANDS.formatCells:
      return { kind: "format-cells" };
    case SHEET_MENU_COMMANDS.moreNumberFormats:
      return { kind: "format-cells", tab: "number" };
    case SHEET_MENU_COMMANDS.clearAll:
      return { kind: "clear", what: "all" };
    case SHEET_MENU_COMMANDS.clearFormats:
      return { kind: "clear", what: "formats" };
    case SHEET_MENU_COMMANDS.clearContents:
      return { kind: "clear", what: "contents" };
    default:
      return null;
  }
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
      cellStyleNames: deps.formatting?.toolbar.value.cellStyles.flatMap((group) => group.styles.map((style) => style.name)) ?? [],
      frozen: deps.structure?.freezeState.value.frozen === true,
      freezeAtCellDisabled: deps.structure?.freezeState.value.atA1 !== false,
      filterOn: deps.sortFilter?.filterOn.value === true,
    }),
  );

  /** View → Gridlines: shown ↔ hidden, as view state. */
  function toggleGridLines(): void {
    deps.viewState.updateActive({ showGridLines: gridLines.value ? false : undefined });
  }

  function run(command = ""): void {
    const format = formatCommand(command);
    if (format) {
      deps.formatting?.run(format);
      return;
    }
    const sortFilter = SORT_FILTER_COMMANDS[command];
    if (sortFilter) {
      deps.sortFilter?.[sortFilter]();
      return;
    }
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
      case SHEET_MENU_COMMANDS.freezeTopRow:
        deps.structure?.freeze("top-row");
        break;
      case SHEET_MENU_COMMANDS.freezeFirstColumn:
        deps.structure?.freeze("first-column");
        break;
      case SHEET_MENU_COMMANDS.freezeAtCell:
        deps.structure?.freeze("at-cell");
        break;
      case SHEET_MENU_COMMANDS.unfreeze:
        deps.structure?.freeze("unfreeze");
        break;
      case SHEET_MENU_COMMANDS.rowHeight:
        deps.structure?.formatBands("row", "size");
        break;
      case SHEET_MENU_COMMANDS.columnWidth:
        deps.structure?.formatBands("col", "size");
        break;
      case SHEET_MENU_COMMANDS.autofitRows:
        deps.structure?.formatBands("row", "autofit");
        break;
      case SHEET_MENU_COMMANDS.autofitColumns:
        deps.structure?.formatBands("col", "autofit");
        break;
      case SHEET_MENU_COMMANDS.pasteSpecial:
        void deps.clipboard?.openPasteSpecial();
        break;
      case SHEET_MENU_COMMANDS.fillDown:
        deps.fillHandle?.fillDown();
        break;
      case SHEET_MENU_COMMANDS.fillRight:
        deps.fillHandle?.fillRight();
        break;
      default:
        break;
    }
  }

  return { state, gridLines, toggleGridLines, run };
}
