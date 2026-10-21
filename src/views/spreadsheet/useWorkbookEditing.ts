// The workbook view's editing wiring (SPEC-121), composed in one call so
// SpreadsheetView stays a wiring hub (SpreadsheetView.size.spec.ts): the undo
// stack and edit pipeline (useSpreadsheetUndo) over the view's model, driving
// SPEC-120's recalculation with reason "edit" and writing each step through
// the store (storeSpreadsheetEdit → touchDocument); the tab bar's commands
// (useSheetCommands); the downloads (useWorkbookDownload); and SPEC-122's
// cell entry (useCellEditing) over the same undo pipeline — `cell` is what
// the grid, the formula bar, the sheet tabs and the status bar bind.
// SPEC-129: the view's one side panel (useSheetSidePanel — the Format menu
// opens it, WorkbookBody renders it) and its conditional-formatting rules
// manager (useConditionalFormatManager), whose steps run on the same undo
// stack. SPEC-130: the Data validation dialog (useValidationDialog — the Data
// menu opens it, WorkbookBody renders it), whose OK is one step on that stack;
// and data validation (useDataValidation) — the cell editor's entry gate, and
// the alert, input message and dropdown it shows (`validation` — WorkbookBody
// renders the alert; the grid context carries it for the rest). SPEC-131: the
// sheet drawing layer (useSheetDrawings — selection, keyboard and commands
// over the same undo stack; the grid context carries it to SheetGrid and
// WorkbookBody, the menu reads it for Insert and Arrange, and a pasted image
// goes to it through useSheetClipboard).
import { computed, triggerRef, type Ref, type ShallowRef } from "vue";
import { storeSpreadsheetEdit } from "../../lib/documentStore";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { useCellEditing, type CellEditing, type CellEditingDeps } from "./useCellEditing";
import { useConditionalFormatManager } from "./useConditionalFormatManager";
import { useConditionalFormats } from "./useConditionalFormats";
import { entryGateOf, useDataValidation } from "./useDataValidation";
import type { Recalc } from "./useRecalc";
import { useSheetCommands } from "./useSheetCommands";
import { useSheetDrawings } from "./useSheetDrawings";
import type { SheetMetrics } from "./useSheetMetrics";
import { useSheetSidePanel } from "./useSheetSidePanel";
import type { SheetViewState } from "./useSheetViewState";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";
import { useValidationDialog } from "./useValidationDialog";
import { useWorkbookDownload } from "./useWorkbookDownload";

export interface WorkbookEditingDeps {
  book: {
    workbook: ShallowRef<SpreadsheetModel | null>;
    activeIndex: Ref<number>;
    activeSheet: Readonly<Ref<Sheet | null>>;
    selectSheet(index: number): void;
    sheetsChanged(): void;
  };
  viewState: Pick<SheetViewState, "syncInto">;
  recalc: Pick<Recalc, "edit" | "settle" | "onResult">;
  /** Cell entry's view inputs (SPEC-122). */
  selection: CellEditingDeps["selection"];
  sheetMetrics: CellEditingDeps["sheetMetrics"] & Pick<SheetMetrics, "metrics">;
  locale: () => SpreadsheetLocale;
}

export function useWorkbookEditing(deps: WorkbookEditingDeps) {
  const undo = useSpreadsheetUndo({
    model: deps.book.workbook,
    recalc: (changed, done) => deps.recalc.edit(changed, done),
    settle: () => deps.recalc.settle(),
    onApplied: () => {
      const model = deps.book.workbook.value;
      if (!model) {
        return;
      }
      triggerRef(deps.book.workbook);
      deps.book.sheetsChanged();
      void storeSpreadsheetEdit(model);
    },
    activateSheet: (sheetId) => {
      const index = deps.book.workbook.value?.sheets.findIndex((sheet) => sheet.id === sheetId) ?? -1;
      if (index >= 0) {
        deps.book.activeIndex.value = index;
      }
    },
  });
  const commands = useSheetCommands({
    book: deps.book,
    undo,
    syncViewState: (model) => deps.viewState.syncInto(model),
  });
  const download = useWorkbookDownload({ book: deps.book });
  const conditionalFormats = useConditionalFormats({ book: deps.book, recalc: deps.recalc, undo, locale: deps.locale });
  // The gate is the editor's dep and a pick commits through the editor.
  let editor: CellEditing | null = null;
  const validation = useDataValidation({ book: deps.book, selection: deps.selection, locale: deps.locale, editor: () => editor });
  const cell = useCellEditing({
    book: deps.book,
    selection: deps.selection,
    sheetMetrics: deps.sheetMetrics,
    undo,
    locale: deps.locale,
    syncViewState: (model) => deps.viewState.syncInto(model),
    entryGate: entryGateOf(validation),
  });
  editor = cell;
  /** SpreadsheetAppBar's undo/redo and split-download props, bound as one
   * (the Edit menu reads the same). While a cell is being edited the
   * workbook's undo and redo are off — the editor's own text undo has
   * Cmd/Ctrl+Z, and the workbook's applies only after the commit (SPEC-122). */
  const appBar = computed(() => ({
    canUndo: undo.canUndo.value && !cell.editing.value,
    canRedo: undo.canRedo.value && !cell.editing.value,
    undoLabel: undo.undoLabel.value,
    redoLabel: undo.redoLabel.value,
    copyLabel: download.copyLabel.value,
    copyDisabledReason: download.copyDisabledReason.value,
    copyOdsLabel: download.copyOdsLabel.value,
  }));
  const sidePanel = useSheetSidePanel();
  const cfManager = useConditionalFormatManager({
    model: () => deps.book.workbook.value,
    sheet: () => deps.book.activeSheet.value,
    selection: () => deps.selection.selection.value,
    undo,
    locale: deps.locale,
    syncInto: (model) => deps.viewState.syncInto(model),
    busy: () => cell.editing.value,
  });
  const validationDialog = useValidationDialog({
    model: () => deps.book.workbook.value,
    sheet: () => deps.book.activeSheet.value,
    selection: () => deps.selection.selection.value,
    undo,
    formatContext: () => deps.selection.formatContext.value,
    syncInto: (model) => deps.viewState.syncInto(model),
    busy: () => cell.editing.value,
  });
  const drawings = useSheetDrawings({
    book: deps.book,
    undo,
    selection: deps.selection,
    sheetMetrics: deps.sheetMetrics,
    syncInto: (model) => deps.viewState.syncInto(model),
    busy: () => cell.editing.value,
  });
  return { undo, commands, download, appBar, cell, conditionalFormats, sidePanel, cfManager, validationDialog, validation, drawings };
}

export type WorkbookEditing = ReturnType<typeof useWorkbookEditing>;

/** Narrows a component's runtime-typed `editing` prop (the .vue parser bars
 * casts). */
export function asWorkbookEditing(value: unknown): WorkbookEditing {
  return value as WorkbookEditing;
}
