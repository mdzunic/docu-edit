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
// useSortFilter's, and the bar's state says whether Filter is on. SPEC-129:
// Format ▸ Conditional formatting… opens the side panel's rules manager
// (useWorkbookEditing's sidePanel). SPEC-130: Data ▸ "Data validation…" opens
// the validation dialog (useWorkbookEditing's validationDialog). SPEC-131:
// Insert ▸ Image… and Shapes and the Arrange menu are the drawing layer's
// (useWorkbookEditing's drawings), and the bar's state says whether a drawing
// can go in and which is selected. SPEC-133: Insert ▸ Table… opens
// useSheetTables' Create Table dialog. SPEC-135: Insert ▸ Link… opens the
// link dialog (useWorkbookEditing's links); Insert ▸ Note and View ▸ Show all
// notes are the notes' (useWorkbookEditing's notes), Insert ▸ Comment and
// View ▸ Comments the comments pane's (its comments), and the bar's state
// says which apply and what is checked. SPEC-136: File ▸ Page setup…,
// Print… and Print area ▸, View ▸ Page breaks and Insert ▸ the page breaks
// are `printing`'s (usePageSetup, useSheetPrint). SPEC-138: File ▸ Version
// history toggles the side panel's history pane (useWorkbookEditing's
// sidePanel), checked while it is open, and the bar's state says when the
// workbook is read-only (useWorkbookEditing's readOnly). Later specs append
// commands here.
import { computed } from "vue";
import { SHEET_MENU_COMMANDS, zoomPresetOf, type SpreadsheetMenuState } from "../../lib/spreadsheet/menuState";
import type { NumberPresetId } from "../../lib/spreadsheet/numberCodes";
import type { Formatting } from "./useFormatting";
import type { FormulaDisplay } from "./useFormulaDisplay";
import type { FillHandle } from "./useFillHandle";
import type { Names } from "./useNames";
import type { SheetPrinting } from "./useSheetPrint";
import type { Recalc } from "./useRecalc";
import type { SheetClipboard } from "./useSheetClipboard";
import type { SheetTables } from "./useSheetTables";
import type { DrawingAction } from "./useSheetDrawings";
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
  editing?: Pick<WorkbookEditing, "undo" | "download" | "appBar"> & Partial<Pick<WorkbookEditing, "sidePanel" | "validationDialog" | "drawings" | "charts" | "links" | "notes" | "comments" | "readOnly">>;
  /** SPEC-123: the Format menu. */
  formatting?: Pick<Formatting, "run" | "toolbar">;
  /** SPEC-124: View ▸ Freeze and the Format menu's rows and columns. */
  structure?: Pick<Structure, "freeze" | "freezeState" | "formatBands">;
  /** SPEC-125: Edit ▸ Paste special… and Edit ▸ Fill ▸ Down/Right. */
  clipboard?: Pick<SheetClipboard, "openPasteSpecial">;
  fillHandle?: Pick<FillHandle, "fillDown" | "fillRight">;
  /** SPEC-128: the Data menu's sorts and filter commands. */
  sortFilter?: Pick<SortFilter, "sortAscending" | "sortDescending" | "openSortDialog" | "toggleFilter" | "clearFilters" | "reapplyFilter" | "filterOn">;
  /** SPEC-133: Insert ▸ Table…. */
  tables?: Pick<SheetTables, "openCreate">;
  /** SPEC-133: Formulas ▸ Name Manager…. */
  names?: Pick<Names, "openManager">;
  /** SPEC-136: page setup, printing, page breaks and the print area. */
  printing?: Pick<SheetPrinting, "pageSetup" | "print">;
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

/** SPEC-131's Arrange menu commands → the drawing actions. */
const ARRANGE_COMMANDS: Readonly<Record<string, DrawingAction>> = {
  [SHEET_MENU_COMMANDS.arrangeFront]: "front",
  [SHEET_MENU_COMMANDS.arrangeForward]: "forward",
  [SHEET_MENU_COMMANDS.arrangeBackward]: "backward",
  [SHEET_MENU_COMMANDS.arrangeBack]: "back",
  [SHEET_MENU_COMMANDS.altText]: "alt-text",
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
      canInsertDrawings: deps.editing?.drawings?.canInsert.value === true,
      drawingSelected: deps.editing?.drawings?.order.value.selected === true,
      drawingAtFront: deps.editing?.drawings?.order.value.atFront === true,
      drawingAtBack: deps.editing?.drawings?.order.value.atBack === true,
      canAnnotate: deps.editing?.notes?.canAnnotate.value === true,
      canComment: deps.editing?.comments?.canComment.value === true,
      commentDisabledReason: deps.editing?.comments?.newCommentDisabledReason.value ?? "",
      noteDisabledReason: deps.editing?.notes?.noteDisabledReason.value ?? "",
      canLink: deps.editing?.links?.canLink.value ?? true,
      linkDisabledReason: deps.editing?.links?.linkDisabledReason.value ?? "",
      commentsOpen: deps.editing?.comments?.isOpen.value === true,
      showAllNotes: deps.editing?.notes?.showAll.value === true,
      historyOpen: deps.editing?.sidePanel?.activePane.value === "history",
      readOnly: deps.editing?.readOnly?.value === true,
      pageBreaks: deps.printing?.pageSetup.pageBreaks.shown.value === true,
      canInsertPageBreak: deps.printing?.pageSetup.state.value.canInsert === true,
      canRemovePageBreak: deps.printing?.pageSetup.state.value.canRemove === true,
      canResetPageBreaks: deps.printing?.pageSetup.state.value.canReset === true,
      hasPrintArea: deps.printing?.pageSetup.state.value.hasPrintArea === true,
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
    const arrange = ARRANGE_COMMANDS[command];
    if (arrange) {
      deps.editing?.drawings?.runAction(arrange);
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
      case SHEET_MENU_COMMANDS.conditionalFormatting:
        deps.editing?.sidePanel?.open("conditionalFormats");
        break;
      case SHEET_MENU_COMMANDS.dataValidation:
        deps.editing?.validationDialog?.show();
        break;
      case SHEET_MENU_COMMANDS.insertImage:
        deps.editing?.drawings?.pickImage();
        break;
      case SHEET_MENU_COMMANDS.insertShape:
        deps.editing?.drawings?.openGallery();
        break;
      case SHEET_MENU_COMMANDS.insertChart:
        deps.editing?.charts?.openGallery();
        break;
      case SHEET_MENU_COMMANDS.nameManager:
        deps.names?.openManager();
        break;
      case SHEET_MENU_COMMANDS.insertTable:
        deps.tables?.openCreate();
        break;
      case SHEET_MENU_COMMANDS.insertLink:
        deps.editing?.links?.openDialog();
        break;
      case SHEET_MENU_COMMANDS.insertNote:
        deps.editing?.notes?.openEditor();
        break;
      case SHEET_MENU_COMMANDS.insertComment:
        deps.editing?.comments?.newComment();
        break;
      case SHEET_MENU_COMMANDS.comments:
        deps.editing?.comments?.toggle();
        break;
      case SHEET_MENU_COMMANDS.showAllNotes:
        deps.editing?.notes?.toggleShowAll();
        break;
      case SHEET_MENU_COMMANDS.versionHistory:
        deps.editing?.sidePanel?.toggle("history");
        break;
      case SHEET_MENU_COMMANDS.pageSetup:
        deps.printing?.pageSetup.dialog.show();
        break;
      case SHEET_MENU_COMMANDS.print:
        deps.printing?.print.open();
        break;
      case SHEET_MENU_COMMANDS.setPrintArea:
        deps.printing?.pageSetup.setArea();
        break;
      case SHEET_MENU_COMMANDS.clearPrintArea:
        deps.printing?.pageSetup.clearArea();
        break;
      case SHEET_MENU_COMMANDS.pageBreaks:
        deps.printing?.pageSetup.pageBreaks.toggle();
        break;
      case SHEET_MENU_COMMANDS.insertPageBreak:
        deps.printing?.pageSetup.insertBreak();
        break;
      case SHEET_MENU_COMMANDS.removePageBreak:
        deps.printing?.pageSetup.removeBreak();
        break;
      case SHEET_MENU_COMMANDS.resetPageBreaks:
        deps.printing?.pageSetup.resetBreaks();
        break;
      default:
        break;
    }
  }

  return { state, gridLines, toggleGridLines, run };
}
