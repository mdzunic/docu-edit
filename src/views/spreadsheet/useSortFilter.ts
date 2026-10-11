// Sort and filter in the workbook view (SPEC-128 "UI"): Data ▸ Sort A to Z,
// Sort Z to A and Sort range… (SortDialog), Data ▸ Filter, Clear filters and
// Reapply filter, the header row's `filter-button`s and their popover
// (useFilterPopover), the refusal alert (SheetAlertDialog, `sheet-alert`),
// the status bar's "N of M records found" and the filtered row headers —
// over lib/spreadsheet/sort.ts and filter.ts, which own every rule.
//
// - Each command is ONE undo step through SPEC-121's record/commit, labelled
//   "Sort", "Filter", "Filter column", "Clear filter", "Clear filters" or
//   "Reapply filter". A sort hands its changed cells to SPEC-120's recalc; a
//   filter change hands visibilitySensitiveCells (SUBTOTAL and AGGREGATE
//   follow which rows show). A refusal changes nothing — its empty step is
//   dropped — and shows its message in `sheet-alert` (the Sort dialog shows
//   it inline instead).
// - What a sort covers: a multi-area selection refuses; a selection (or the
//   active cell) inside the autofilter's range sorts that range below its
//   header; a single cell (or merge) expands to its currentRegion; any other
//   selection sorts exactly itself. The header guess (guessHasHeader) seeds
//   the dialog's "My data has headers"; the quick sorts use it as is. Text
//   collates in the user's number locale, case-insensitively.
// - Data ▸ Filter on covers filter.ts autoFilterRangeFor's range (an empty
//   region refuses with "Select a cell inside the data you want to filter").
// - Keys, only in the grid while no cell is being edited (the SPEC-021
//   convention — never from an input, a dialog or the cell editor):
//   Ctrl/Cmd+Shift+L toggles the filter, Ctrl/Cmd+Alt+L reapplies it (best
//   effort on macOS, where Chrome keeps Cmd+Option+L; the menu item is the
//   guaranteed path), Alt+Down on a header cell opens its popover.
import { computed, nextTick, onScopeDispose, ref, shallowRef, type Ref, type ShallowRef } from "vue";
import { isMacPlatform } from "../../lib/core/platform";
import { findShortcut } from "../../lib/shortcuts";
import { formatA1, type CellRef, type GridRange } from "../../lib/spreadsheet/address";
import { selectionSpansMultipleCells } from "../../lib/spreadsheet/aggregates";
import {
  autoFilterRangeFor,
  clearFilters,
  filterButtonState,
  filterRecordCounts,
  hasCriteria,
  reapplyFilter,
  toggleAutoFilter,
  type FilterCommandResult,
} from "../../lib/spreadsheet/filter";
import type { CellKey } from "../../lib/spreadsheet/formula/dependencies";
import { visibilitySensitiveCells } from "../../lib/spreadsheet/formula/recalc";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import type { FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import { currentRegion, selectCell } from "../../lib/spreadsheet/selection";
import { guessHasHeader, SORT_REFUSAL_MESSAGES, sortRange, type SheetSortKey, type SortOptions, type SortOutcome } from "../../lib/spreadsheet/sort";
import type { CellEditing } from "./useCellEditing";
import { inGrid } from "./headerMenu";
import { SHEET_SHORTCUTS } from "./sheetShortcuts";
import { headerLabel, sortColumnChoices, type SortColumnChoice } from "./sortFilterLabels";
import { useFilterPopover, type PopoverAnchor } from "./useFilterPopover";
import type { SheetMetrics } from "./useSheetMetrics";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export interface SortFilterDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeSheet: Readonly<Ref<Sheet | null>> };
  viewState: Pick<SheetViewState, "syncInto">;
  selection: Pick<SheetSelectionController, "selection" | "set" | "formatContext">;
  sheetMetrics: Pick<SheetMetrics, "nav" | "merges">;
  editing: { undo: Pick<SpreadsheetUndo, "record" | "commit">; cell: Pick<CellEditing, "editing"> };
  /** The user's number locale (useSpreadsheetLocale): text collates in it. */
  locale: () => SpreadsheetLocale;
}

/** The Sort dialog's state: what it sorts and the choices it offers. */
export interface SortDialogState {
  sheetId: string;
  range: GridRange;
  /** guessHasHeader's answer — the checkbox's initial state. */
  hasHeader: boolean;
  /** The first level's column (the active cell's). */
  keyCol: number;
  withHeader: SortColumnChoice[];
  withoutHeader: SortColumnChoice[];
  /** A refusal or a duplicated column, shown inline ("" for none). */
  error: string;
}

/** A header's button as the cell layer draws it. */
export interface FilterButtonView {
  label: string;
  state: "none" | "filtered" | "sorted-asc" | "sorted-desc";
  icon: string;
  expanded: boolean;
}

const BUTTON_ICONS: Readonly<Record<FilterButtonView["state"], string>> = {
  none: "chevron-down",
  filtered: "filter",
  "sorted-asc": "sort-ascending",
  "sorted-desc": "sort-descending",
};

/** Whether `inner` lies wholly inside `outer`. */
const within = (inner: GridRange, outer: GridRange): boolean =>
  inner.top >= outer.top && inner.bottom <= outer.bottom && inner.left >= outer.left && inner.right <= outer.right;

export function useSortFilter(deps: SortFilterDeps) {
  /** The refusal on show in `sheet-alert` (null: none). */
  const alert = ref<string | null>(null);
  const sortDialog = shallowRef<SortDialogState | null>(null);

  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const gridSheet = (): Sheet | null => {
    const sheet = deps.book.activeSheet.value;
    return sheet && sheet.kind === "grid" ? sheet : null;
  };
  const ready = (): boolean => model() !== null && gridSheet() !== null && !deps.editing.cell.editing.value;
  const ctx = (): FormatContext => deps.selection.formatContext.value;

  /** One command as ONE undo step on the sheet on show; `recalc` names the
   * cells to recalculate after it (none after a refusal). */
  function step<T>(label: string, mutate: (m: SpreadsheetModel, sheetId: string) => T, recalc: (m: SpreadsheetModel, result: T) => CellKey[]): T | null {
    const m = model();
    const sheet = gridSheet();
    if (!m || !sheet || deps.editing.cell.editing.value) {
      return null;
    }
    deps.viewState.syncInto(m);
    deps.editing.undo.record(label, { sheetId: sheet.id });
    const result = mutate(m, sheet.id);
    deps.editing.undo.commit(recalc(m, result));
    return result;
  }

  /** A filter command; a refusal opens the alert. */
  function filterStep(label: string, mutate: (m: SpreadsheetModel, sheetId: string, ctx: FormatContext) => FilterCommandResult): boolean {
    const format = ctx();
    const result = step(label, (m, sheetId) => mutate(m, sheetId, format), (m, done) => (done.ok ? visibilitySensitiveCells(m) : []));
    if (result && !result.ok) {
      alert.value = result.message;
    }
    return result?.ok === true;
  }

  /** A sort as one "Sort" step (its outcome, null when nothing could run). */
  function sortStep(range: GridRange | readonly GridRange[], options: Omit<SortOptions, "localeTag">): SortOutcome | null {
    const localeTag = deps.locale().tag;
    return step("Sort", (m, sheetId) => sortRange(m, sheetId, range, { ...options, localeTag }), (_m, outcome) => (outcome.ok ? outcome.changed : []));
  }

  /** What a sort covers (see the module header): the ranges as given when
   * several (sortRange refuses those), else one range and its header guess. */
  function sortTarget(m: SpreadsheetModel, sheet: Sheet): { range: GridRange | readonly GridRange[]; hasHeader: boolean } {
    const sel = deps.selection.selection.value;
    if (sel.ranges.length > 1) {
      return { range: sel.ranges, hasHeader: false };
    }
    const selected = sel.ranges[0] ?? { top: sel.active.row, left: sel.active.col, bottom: sel.active.row, right: sel.active.col };
    const filter = sheet.autoFilter;
    if (filter && within(selected, filter.range)) {
      return { range: filter.range, hasHeader: true };
    }
    const range = selectionSpansMultipleCells(sel.ranges, deps.sheetMetrics.merges.value) ? selected : currentRegion(sel.active, deps.sheetMetrics.nav());
    return { range, hasHeader: guessHasHeader(m, sheet, range) };
  }

  /** Data ▸ Sort A to Z / Z to A: by the active cell's column. */
  function quickSort(descending: boolean): void {
    const m = model();
    const sheet = gridSheet();
    if (!m || !sheet || !ready()) {
      return;
    }
    const target = sortTarget(m, sheet);
    const keys: SheetSortKey[] = [{ col: deps.selection.selection.value.active.col, descending }];
    const outcome = sortStep(target.range, { keys, hasHeader: target.hasHeader, caseSensitive: false });
    if (outcome && !outcome.ok) {
      alert.value = outcome.message;
    }
  }

  /** Data ▸ Sort range…: the dialog over the sort's target. */
  function openSortDialog(): void {
    const m = model();
    const sheet = gridSheet();
    if (!m || !sheet || !ready()) {
      return;
    }
    const target = sortTarget(m, sheet);
    if (!("top" in target.range)) {
      alert.value = SORT_REFUSAL_MESSAGES.multiArea;
      return;
    }
    const range = target.range;
    const keyCol = Math.min(Math.max(deps.selection.selection.value.active.col, range.left), range.right);
    sortDialog.value = {
      sheetId: sheet.id,
      range,
      hasHeader: target.hasHeader,
      keyCol,
      withHeader: sortColumnChoices(m, sheet, range, true, keyCol, ctx()),
      withoutHeader: sortColumnChoices(m, sheet, range, false, keyCol, ctx()),
      error: "",
    };
  }

  /** The dialog's OK: sorts and closes, or shows the refusal inline. */
  function confirmSort(choice: { keys: SheetSortKey[]; hasHeader: boolean; caseSensitive: boolean }): void {
    const dialog = sortDialog.value;
    if (!dialog || gridSheet()?.id !== dialog.sheetId) {
      sortDialog.value = null;
      return;
    }
    const outcome = sortStep(dialog.range, choice);
    if (outcome) {
      sortDialog.value = outcome.ok ? null : { ...dialog, error: outcome.message };
    }
  }

  /** The popover's sorts: the autofilter's data rows by one column. */
  function sortFilterRange(col: number, descending: boolean): void {
    const range = gridSheet()?.autoFilter?.range;
    if (!range) {
      return;
    }
    const outcome = sortStep(range, { keys: [{ col, descending }], hasHeader: true, caseSensitive: false });
    if (outcome && !outcome.ok) {
      alert.value = outcome.message;
    }
  }

  /** Gives the grid focus — making `at` the active cell first when given. */
  function focusGrid(at: CellRef | null): void {
    if (at) {
      deps.selection.set(selectCell(at, deps.sheetMetrics.nav()));
    }
    void nextTick(() => document.querySelector<HTMLElement>('[data-testid="spreadsheet-grid"]')?.focus({ preventScroll: true }));
  }

  const popover = useFilterPopover({
    model,
    sheet: gridSheet,
    formatContext: ctx,
    locale: () => deps.locale(),
    filterStep,
    sortFilterRange,
    focusGrid,
  });

  // --- the filter commands ------------------------------------------------------

  const filterOn = computed(() => gridSheet()?.autoFilter !== undefined);

  /** Data ▸ Filter (Ctrl/Cmd+Shift+L): on over the chosen range, or off. */
  function toggleFilter(): void {
    const sheet = gridSheet();
    if (!sheet || !ready()) {
      return;
    }
    const range = sheet.autoFilter ? null : autoFilterRangeFor(deps.selection.selection.value, deps.sheetMetrics.nav());
    popover.close("outside");
    filterStep("Filter", (m, sheetId, format) => toggleAutoFilter(m, sheetId, range, format));
  }

  function clearAll(): void {
    if (filterOn.value && ready()) {
      filterStep("Clear filters", (m, sheetId) => clearFilters(m, sheetId));
    }
  }

  function reapply(): void {
    if (filterOn.value && ready()) {
      filterStep("Reapply filter", (m, sheetId, format) => reapplyFilter(m, sheetId, format));
    }
  }

  // --- what the grid and the status bar draw -----------------------------------------

  /** The filtered data rows while the filter holds a criterion (null: none). */
  const filteredRows = computed(() => {
    const filter = gridSheet()?.autoFilter;
    return filter && hasCriteria(filter) ? { top: filter.range.top + 1, bottom: filter.range.bottom } : null;
  });

  /** "8 of 12 records found" while the filter holds a criterion; "" else. */
  const statusText = computed(() => {
    const sheet = gridSheet();
    const counts = sheet ? filterRecordCounts(sheet) : null;
    return counts ? `${counts.found} of ${counts.total} records found` : "";
  });

  /** The `filter-button` in the header cell at (row, col), or null. Read
   * during the cell layer's render, so it follows the filter and popover. */
  function filterButtonAt(row: number, col: number): FilterButtonView | null {
    const m = model();
    const sheet = gridSheet();
    const filter = sheet?.autoFilter;
    if (!m || !sheet || !filter || row !== filter.range.top || col < filter.range.left || col > filter.range.right) {
      return null;
    }
    const colId = col - filter.range.left;
    if (filter.columns.some((column) => column.colId === colId && column.hideButton)) {
      return null;
    }
    const state = filterButtonState(filter, colId);
    return { label: `Filter "${headerLabel(m, sheet, row, col, ctx())}"`, state, icon: BUTTON_ICONS[state], expanded: popover.openColId.value === colId };
  }

  /** A button's press: its popover, anchored under the button. */
  function pressFilterButton(col: number, anchor: PopoverAnchor): void {
    const filter = gridSheet()?.autoFilter;
    if (filter && ready()) {
      popover.open(col - filter.range.left, anchor);
    }
  }

  /** Alt+Down on a header cell with a button: its popover. */
  function openAtActive(): boolean {
    const active = deps.selection.selection.value.active;
    const filter = gridSheet()?.autoFilter;
    if (!filter || !filterButtonAt(active.row, active.col)) {
      return false;
    }
    const cell = document.querySelector(`[data-testid="sheet-cell"][data-ref="${formatA1(active)}"]`);
    const box = (cell?.querySelector('[data-testid="filter-button"]') ?? cell)?.getBoundingClientRect();
    popover.open(active.col - filter.range.left, { left: box?.left ?? 0, top: box?.bottom ?? 0 });
    return true;
  }

  function onKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented || !inGrid(event.target) || !ready()) {
      return;
    }
    if (event.key === "ArrowDown" && event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey) {
      if (openAtActive()) {
        event.preventDefault();
        event.stopPropagation();
      }
      return;
    }
    const id = findShortcut(event, isMacPlatform(), SHEET_SHORTCUTS)?.id;
    if (id === "sheet.toggleFilter") {
      event.preventDefault();
      toggleFilter();
    } else if (id === "sheet.reapplyFilter") {
      event.preventDefault();
      reapply();
    }
  }

  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown, true);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown, true);
    }
  });

  return {
    sortAscending: (): void => quickSort(false),
    sortDescending: (): void => quickSort(true),
    sortDialog,
    openSortDialog,
    confirmSort,
    cancelSort: (): null => (sortDialog.value = null),
    filterOn,
    toggleFilter,
    clearFilters: clearAll,
    reapplyFilter: reapply,
    popover,
    filterButtonAt,
    pressFilterButton,
    filteredRows,
    isRowFiltered: (row: number): boolean => filteredRows.value !== null && row >= filteredRows.value.top && row <= filteredRows.value.bottom,
    statusText,
    alert,
    closeAlert(): void {
      alert.value = null;
      focusGrid(null);
    },
    onKeydown,
  };
}

export type SortFilter = ReturnType<typeof useSortFilter>;

/** Narrows a component's runtime-typed `sortFilter` prop (the .vue parser
 * bars casts). */
export function asSortFilter(value: unknown): SortFilter {
  return value as SortFilter;
}

/** The grid context's `sortFilter`, when the view wired one (null: none). */
export function sortFilterOf(context: unknown): SortFilter | null {
  const value = context !== null && typeof context === "object" ? Reflect.get(context, "sortFilter") : undefined;
  return value ? (value as SortFilter) : null;
}
