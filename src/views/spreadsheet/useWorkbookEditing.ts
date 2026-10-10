// The workbook view's editing wiring (SPEC-121), composed in one call so
// SpreadsheetView stays a wiring hub (SpreadsheetView.size.spec.ts): the undo
// stack and edit pipeline (useSpreadsheetUndo) over the view's model, driving
// SPEC-120's recalculation with reason "edit" and writing each step through
// the store (storeSpreadsheetEdit → touchDocument); the tab bar's commands
// (useSheetCommands); and the downloads (useWorkbookDownload).
import { computed, triggerRef, type Ref, type ShallowRef } from "vue";
import { storeSpreadsheetEdit } from "../../lib/documentStore";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { Recalc } from "./useRecalc";
import { useSheetCommands } from "./useSheetCommands";
import type { SheetViewState } from "./useSheetViewState";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";
import { useWorkbookDownload } from "./useWorkbookDownload";

export interface WorkbookEditingDeps {
  book: {
    workbook: ShallowRef<SpreadsheetModel | null>;
    activeIndex: Ref<number>;
    selectSheet(index: number): void;
    sheetsChanged(): void;
  };
  viewState: Pick<SheetViewState, "syncInto">;
  recalc: Pick<Recalc, "edit" | "settle">;
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
  /** SpreadsheetAppBar's undo/redo and split-download props, bound as one. */
  const appBar = computed(() => ({
    canUndo: undo.canUndo.value,
    canRedo: undo.canRedo.value,
    undoLabel: undo.undoLabel.value,
    redoLabel: undo.redoLabel.value,
    copyLabel: download.copyLabel.value,
    copyDisabledReason: download.copyDisabledReason.value,
  }));
  return { undo, commands, download, appBar };
}

export type WorkbookEditing = ReturnType<typeof useWorkbookEditing>;

/** Narrows a component's runtime-typed `editing` prop (the .vue parser bars
 * casts). */
export function asWorkbookEditing(value: unknown): WorkbookEditing {
  return value as WorkbookEditing;
}
