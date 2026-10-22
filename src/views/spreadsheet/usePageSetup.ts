// SPEC-136's page setup in the workbook view: the Page Setup dialog
// (SheetPageSetupDialog — its fields are pageSetupForm.ts's, OK one undo
// step "Page setup" — none for a form left as shown — Cancel nothing), the
// manual-break and print-area
// commands (Insert ▸ Page break / Remove page break / Reset all page breaks,
// File ▸ Print area ▸ Set / Clear — each one undo step through record(label,
// { sheetId }) before mutating and commit() after), and View ▸ Page breaks:
// a session-only, never-written toggle per sheet drawing the active sheet's
// breaks (paginateSheet's, recomputed on an idle callback after a model or
// setup change). Closing a print preview turns the lines on for its sheet,
// as Excel does (showFor).
import { computed, ref, shallowRef, watch, type ComputedRef, type Ref } from "vue";
import { runWhenIdle } from "../../lib/core/idleScheduler";
import type { MeasurementUnit } from "../../lib/core/units";
import type { CellRef } from "../../lib/spreadsheet/address";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import {
  canInsertPageBreak,
  effectivePaper,
  hasManualBreaks,
  insertPageBreak,
  pageBreaksAt,
  printAreaOf,
  printTitlesOf,
  removePageBreak,
  resetPageBreaks,
  setPageSetup,
  setPrintArea,
  setPrintTitles,
} from "../../lib/spreadsheet/pageSetup";
import { pageBreakLines, paginateSheet, type PageBreakLine } from "../../lib/spreadsheet/print";
import { usePreferences } from "../../composables/usePreferences";
import { formFromSetup, setupFromForm, validateForm, type FormProblem, type PageSetupForm, type SectionFields } from "./pageSetupForm";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";
import type { WorkbookEditing } from "./useWorkbookEditing";
import type { asWorkbookBook } from "./useWorkbookLoad";

export const PAGE_SETUP_LABEL = "Page setup";
export const PAGE_BREAK_LABELS = {
  insert: "Insert page break",
  remove: "Remove page break",
  reset: "Reset page breaks",
  setArea: "Set print area",
  clearArea: "Clear print area",
} as const;

export type PageSetupTab = "page" | "margins" | "headerFooter" | "sheet";

export interface PageSetupDeps {
  book: ReturnType<typeof asWorkbookBook>;
  viewState: Pick<SheetViewState, "syncInto">;
  selection: Pick<SheetSelectionController, "selection">;
  editing: Pick<WorkbookEditing, "undo" | "cell">;
}

export interface PageSetupDialog {
  open: Ref<boolean>;
  tab: Ref<PageSetupTab>;
  form: Ref<PageSetupForm | null>;
  problem: ComputedRef<FormProblem | null>;
  unit: ComputedRef<MeasurementUnit>;
  /** The header/footer section a code insert goes into. */
  focused: Ref<{ part: "header" | "footer"; section: keyof SectionFields }>;
  show(tab?: PageSetupTab): void;
  /** A tab button's choice (its id as the template has it). */
  selectTab(id: string): void;
  ok(): void;
  cancel(): void;
  insertCode(code: string): void;
}

export function usePageSetup(deps: PageSetupDeps) {
  const { preferences } = usePreferences();
  const unit = computed((): MeasurementUnit => preferences.measurementUnit ?? "cm");
  const printPreferences = computed(() => ({ defaultPageSizeId: preferences.defaultPageSizeId }));

  const gridSheet = (): Sheet | null => {
    const sheet = deps.book.activeSheet.value;
    return sheet?.kind === "grid" ? sheet : null;
  };
  const busy = (): boolean => deps.editing.cell.editing.value;

  /** One command as ONE undo step on the sheet on show. */
  function command(label: string, mutate: (model: SpreadsheetModel, sheet: Sheet) => void): void {
    const model = deps.book.workbook.value;
    const sheet = gridSheet();
    if (!model || !sheet || busy()) {
      return;
    }
    deps.viewState.syncInto(model);
    deps.editing.undo.record(label, { sheetId: sheet.id });
    const current = model.sheets.find((each) => each.id === sheet.id) ?? sheet;
    mutate(model, current);
    deps.editing.undo.commit();
  }

  // --- the dialog ---------------------------------------------------------------

  const open = ref(false);
  const tab = ref<PageSetupTab>("page");
  const form = ref<PageSetupForm | null>(null);
  const focused = ref<{ part: "header" | "footer"; section: keyof SectionFields }>({ part: "header", section: "center" });
  const problem = computed(() => (form.value === null ? null : validateForm(form.value, unit.value)));
  /** The form as shown: OK on it unchanged writes nothing (no step, no
   * paper pinned on a sheet that names none, nothing marked dirty). */
  let shownForm = "";

  function show(at: PageSetupTab = "page"): void {
    const model = deps.book.workbook.value;
    const sheet = gridSheet();
    if (!model || !sheet) {
      return;
    }
    form.value = formFromSetup(sheet.pageSetup, effectivePaper(sheet, printPreferences.value), unit.value, printAreaOf(model, sheet), printTitlesOf(model, sheet));
    shownForm = JSON.stringify(form.value);
    tab.value = at;
    focused.value = { part: "header", section: "center" };
    open.value = true;
  }

  function cancel(): void {
    open.value = false;
    form.value = null;
  }

  function ok(): void {
    const value = form.value;
    if (value === null || problem.value !== null) {
      return;
    }
    if (JSON.stringify(value) === shownForm) {
      cancel();
      return;
    }
    const sheet = gridSheet();
    const { patch, printArea, titles } = setupFromForm(value, unit.value, sheet?.pageSetup);
    command(PAGE_SETUP_LABEL, (model, current) => {
      setPageSetup(model, current, patch);
      setPrintArea(model, current.id, printArea);
      setPrintTitles(model, current.id, titles);
    });
    cancel();
  }

  function insertCode(code: string): void {
    if (form.value === null) {
      return;
    }
    const { part, section } = focused.value;
    form.value[part][section] += code;
  }

  function selectTab(id: string): void {
    tab.value = id === "margins" || id === "headerFooter" || id === "sheet" ? id : "page";
  }

  const dialog: PageSetupDialog = { open, tab, form, problem, unit, focused, show, selectTab, ok, cancel, insertCode };

  // --- breaks and the print area ----------------------------------------------------

  const active = (): CellRef => deps.selection.selection.value.active;
  /** The percent a fit-to-page sheet prints at (Insert ▸ Page break's switch). */
  const fitScale = (model: SpreadsheetModel, sheet: Sheet): number | undefined =>
    paginateSheet(model, sheet, { preferences: printPreferences.value }).pages[0]?.scale;

  const state = computed(() => {
    const sheet = gridSheet();
    const at = deps.selection.selection.value.active;
    const breaks = sheet === null ? { row: false, col: false } : pageBreaksAt(sheet, at);
    return {
      canInsert: sheet !== null && canInsertPageBreak(at),
      canRemove: breaks.row || breaks.col,
      canReset: sheet !== null && hasManualBreaks(sheet),
      hasPrintArea: sheet !== null && deps.book.workbook.value !== null && printAreaOf(deps.book.workbook.value, sheet) !== null,
      hasGrid: sheet !== null,
    };
  });

  function insertBreak(): void {
    command(PAGE_BREAK_LABELS.insert, (model, sheet) => insertPageBreak(model, sheet, active(), fitScale(model, sheet)));
  }
  function removeBreak(): void {
    command(PAGE_BREAK_LABELS.remove, (model, sheet) => removePageBreak(model, sheet, active()));
  }
  function resetBreaks(): void {
    command(PAGE_BREAK_LABELS.reset, (model, sheet) => resetPageBreaks(model, sheet));
  }
  function setArea(): void {
    command(PAGE_BREAK_LABELS.setArea, (model, sheet) => setPrintArea(model, sheet, deps.selection.selection.value.ranges));
  }
  function clearArea(): void {
    command(PAGE_BREAK_LABELS.clearArea, (model, sheet) => setPrintArea(model, sheet, null));
  }

  // --- View ▸ Page breaks -------------------------------------------------------------

  const shownSheets = ref<ReadonlySet<string>>(new Set());
  const linesShown = computed(() => {
    const sheet = gridSheet();
    return sheet !== null && shownSheets.value.has(sheet.id);
  });
  const lines = shallowRef<PageBreakLine[]>([]);
  let generation = 0;

  function recompute(): void {
    generation += 1;
    const ticket = generation;
    if (!linesShown.value) {
      lines.value = [];
      return;
    }
    runWhenIdle(() => {
      const model = deps.book.workbook.value;
      const sheet = gridSheet();
      if (ticket !== generation || !model || !sheet) {
        return;
      }
      lines.value = pageBreakLines(paginateSheet(model, sheet, { preferences: printPreferences.value }), sheet);
    }, 300);
  }

  // Commands, undo, redo and recalculation mutate the model in place, a print
  // area (a name) without touching the sheet: the workbook ref itself (a
  // shallow ref, so its triggerRef fires this) and the book's revision say
  // the model changed, whatever changed in it.
  watch([deps.book.workbook, deps.book.revision, () => deps.book.activeSheet.value, linesShown, printPreferences], recompute, { immediate: true });

  function toggleLines(): void {
    const sheet = gridSheet();
    if (sheet === null) {
      return;
    }
    const next = new Set(shownSheets.value);
    if (next.has(sheet.id)) {
      next.delete(sheet.id);
    } else {
      next.add(sheet.id);
    }
    shownSheets.value = next;
  }

  /** Turns the lines on for a sheet (a closed print preview's). */
  function showLinesFor(sheetId: string): void {
    if (!shownSheets.value.has(sheetId)) {
      shownSheets.value = new Set([...shownSheets.value, sheetId]);
    }
  }

  return {
    dialog,
    state,
    insertBreak,
    removeBreak,
    resetBreaks,
    setArea,
    clearArea,
    printPreferences,
    pageBreaks: { shown: linesShown, lines, toggle: toggleLines, showFor: showLinesFor },
  };
}

export type PageSetup = ReturnType<typeof usePageSetup>;

/** Page-break lines with their offsets in the grid's overlay layer (those
 * `position` places; it returns null for one scrolled out of sight). */
export function placedBreakLines(lines: unknown, position: unknown): (PageBreakLine & { at: number })[] {
  const place = position as (axis: string, index: number) => number | null;
  return (lines as PageBreakLine[]).flatMap((line) => {
    const at = place(line.axis, line.index);
    return typeof at === "number" ? [{ ...line, at }] : [];
  });
}

/** Typed pass-through for the dialog's runtime prop (espree). */
export function asPageSetupDialog(value: unknown): PageSetupDialog {
  return value as PageSetupDialog;
}
