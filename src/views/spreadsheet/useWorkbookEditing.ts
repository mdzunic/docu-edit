// The workbook view's editing wiring (SPEC-121), composed in one call so
// SpreadsheetView stays a wiring hub (SpreadsheetView.size.spec.ts): the undo
// stack and edit pipeline (useSpreadsheetUndo) over the view's model, driving
// SPEC-120's recalculation with reason "edit" and handing each step to the
// saving session (SPEC-138: `onApplied` calls its noteEdit() — the SPEC-010
// draft cadence — where it wrote the workbook through before; editing waits
// for the session's edit lock and stops while a version is previewed, and a
// download is a commit point); the tab bar's commands
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
// goes to it through useSheetClipboard). SPEC-132: the sheet charts
// (useSheetCharts — Insert ▸ Chart's gallery, the chart toolbar's commands
// and dialogs over the same undo stack, and the refresh each recalculation
// pass triggers; WorkbookBody renders its surfaces, the menu opens its
// gallery).
// SPEC-133 moved in, from the view, SPEC-127's function assists
// (useFunctionAssist — `assist`) and SPEC-126's CSV import (useCsvImport —
// `csvImport`), so the view has room for the names and tables wiring.
// SPEC-135: notes on cells (useCellNotes — the indicators, popovers, note
// boxes and editor the grid context carries to SheetGrid; the menu and the
// cell context menu run its commands) and the comments pane
// (useSheetComments — the side panel's "comments" pane, the app bar's
// toggle), each step on the same undo stack; and links on cells
// (useCellLinks — the link dialog WorkbookBody renders, the chip and the
// data-link the grid context carries to SheetGrid, Ctrl/Cmd+K and following).
import { computed, triggerRef, type Ref, type ShallowRef } from "vue";
import { useRouter } from "vue-router";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { useAnnotationEditing } from "./useAnnotationEditing";
import { useCellEditing, type CellEditing, type CellEditingDeps } from "./useCellEditing";
import { useCellLinks } from "./useCellLinks";
import { useCellNotes } from "./useCellNotes";
import { useConditionalFormatManager } from "./useConditionalFormatManager";
import { useConditionalFormats } from "./useConditionalFormats";
import { useCsvImport } from "./useCsvImport";
import { entryGateOf, useDataValidation } from "./useDataValidation";
import { useFunctionAssist } from "./useFunctionAssist";
import type { Recalc } from "./useRecalc";
import { useSheetCharts } from "./useSheetCharts";
import { useSheetComments } from "./useSheetComments";
import { useSheetCommands } from "./useSheetCommands";
import { useSheetDrawings } from "./useSheetDrawings";
import type { SheetMetrics } from "./useSheetMetrics";
import { useSheetSidePanel } from "./useSheetSidePanel";
import type { SheetViewState } from "./useSheetViewState";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";
import { useValidationDialog } from "./useValidationDialog";
import { useWorkbookDownload } from "./useWorkbookDownload";
import type { WorkbookSaving } from "./useWorkbookSaving";

export interface WorkbookEditingDeps {
  book: {
    workbook: ShallowRef<SpreadsheetModel | null>;
    activeIndex: Ref<number>;
    activeSheet: Readonly<Ref<Sheet | null>>;
    selectSheet(index: number): void;
    sheetsChanged(): void;
    /** SPEC-126: shows a re-read model in place of the one on show. */
    replace(model: SpreadsheetModel): void;
    /** SPEC-137: the workbook's storage came back incomplete — read-only. */
    storageIncomplete?: Readonly<Ref<boolean>>;
  };
  /** `update`: SPEC-135's comments select a thread's cell on its sheet. */
  viewState: Pick<SheetViewState, "syncInto" | "update">;
  recalc: Pick<Recalc, "edit" | "settle" | "onResult">;
  /** Cell entry's view inputs (SPEC-122). */
  selection: CellEditingDeps["selection"];
  sheetMetrics: CellEditingDeps["sheetMetrics"] & Pick<SheetMetrics, "metrics">;
  locale: () => SpreadsheetLocale;
  /** SPEC-138: the saving session (useWorkbookSaving). */
  session: Pick<WorkbookSaving, "noteEdit" | "editable" | "commitPoint">;
}

export function useWorkbookEditing(deps: WorkbookEditingDeps) {
  // SPEC-137: no edit reaches a workbook whose storage came back incomplete;
  // SPEC-138: nor one this tab does not own (another holds its edit lock, or
  // the lock is still resolving), nor a version previewed in its place.
  const locked = (): boolean => !deps.session.editable.value;
  const readOnly = (): boolean => deps.book.storageIncomplete?.value === true || locked();
  const undo = useSpreadsheetUndo({
    readOnly,
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
      deps.session.noteEdit();
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
  const download = useWorkbookDownload({ book: deps.book, onDownload: () => deps.session.commitPoint() });
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
    readOnly,
  });
  editor = cell;
  /** The grid's cell editor — withdrawn while the session is locked, so the
   * grid reads aria-readonly="true" and typing does nothing (SPEC-138). */
  const gridEditor = computed(() => (locked() ? undefined : cell));
  /** SpreadsheetAppBar's undo/redo and split-download props, bound as one
   * (the Edit menu reads the same). While a cell is being edited the
   * workbook's undo and redo are off — the editor's own text undo has
   * Cmd/Ctrl+Z, and the workbook's applies only after the commit (SPEC-122). */
  const appBar = computed(() => ({
    canUndo: undo.canUndo.value && !cell.editing.value && !locked(),
    canRedo: undo.canRedo.value && !cell.editing.value && !locked(),
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
  const charts = useSheetCharts({
    book: deps.book,
    undo,
    recalc: deps.recalc,
    selection: deps.selection,
    sheetMetrics: deps.sheetMetrics,
    drawings,
    syncInto: (model) => deps.viewState.syncInto(model),
    busy: () => cell.editing.value,
  });
  // SPEC-127's function assists over the cell entry, and SPEC-126's CSV
  // import bar and dialog (moved here from the view to keep it a hub).
  const assist = useFunctionAssist({ ...cell });
  const csvImport = useCsvImport({ book: deps.book, undo, router: useRouter(), locked });
  // SPEC-135: an .xlsx sheet whose annotations the reader left verbatim
  // refuses new notes, comments and links (AC-11).
  const annotationEditing = useAnnotationEditing({ book: deps.book });
  const comments = useSheetComments({
    book: deps.book,
    undo,
    selection: deps.selection,
    viewState: deps.viewState,
    cell,
    sidePanel,
    syncInto: (model) => deps.viewState.syncInto(model),
    busy: () => cell.editing.value,
    blocked: () => annotationEditing.notesReason.value,
  });
  const notes = useCellNotes({
    book: deps.book,
    undo,
    selection: deps.selection,
    syncInto: (model) => deps.viewState.syncInto(model),
    busy: () => cell.editing.value,
    openThread: comments.openThread,
    focusGrid: () => cell.focusCell(),
    blocked: () => annotationEditing.notesReason.value,
  });
  const links = useCellLinks({
    book: deps.book,
    selection: deps.selection,
    viewState: deps.viewState,
    undo,
    cell,
    locale: deps.locale,
    blocked: () => annotationEditing.linksReason.value,
  });
  return { undo, commands, download, appBar, cell, gridEditor, conditionalFormats, sidePanel, cfManager, validationDialog, validation, drawings, charts, assist, csvImport, notes, comments, links };
}

export type WorkbookEditing = ReturnType<typeof useWorkbookEditing>;

/** Narrows a component's runtime-typed `editing` prop (the .vue parser bars
 * casts). */
export function asWorkbookEditing(value: unknown): WorkbookEditing {
  return value as WorkbookEditing;
}
