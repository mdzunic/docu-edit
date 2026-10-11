// The filter popover and its two dialogs (SPEC-128 "UI"), split out of
// useSortFilter, which hands in its undo-step runner and its sort.
//
// - A header's `filter-button` (or Alt+Down on the header cell) opens
//   `filter-popover` for that column: the sorts by that column over the
//   autofilter's data rows, `Clear filter from "Region"`, the conditions
//   submenu (filterConditions.ts), the search box, "(Select All)" and the
//   value list (sortFilterLabels.ts' entries over filter.ts
//   filterValueList), OK and Cancel. OK writes the checked items (all
//   checked: the criteria removed; nothing toggled: nothing written).
// - A column holding a `preserved` criterion — or any column while the ods
//   filter does not decompose (`odsUnmappedFilter`) — shows only
//   "Filtered by colour or icon — this filter can't be edited here" and
//   Clear filter.
// - Each change is ONE undo step through the runner: "Filter column" for a
//   criterion, "Clear filter" for the column's Clear, "Sort" for a sort.
// - Escape closes the popover and returns focus to the header cell (it
//   becomes the active cell, the grid takes focus); OK, Cancel and the
//   actions hand focus back to the grid; an outside press only closes it.
import { computed, nextTick, shallowRef, watch } from "vue";
import { type CellRef } from "../../lib/spreadsheet/address";
import {
  columnValueKind,
  filterValueList,
  setFilterCriteria,
  type FilterCommandResult,
  type FilterCriteria,
} from "../../lib/spreadsheet/filter";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import type { FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import {
  conditionItems,
  CONDITIONS_LABELS,
  customCriteria,
  customDraftOf,
  top10Criteria,
  type ColumnKind,
  type CustomDraft,
} from "./filterConditions";
import { headerLabel, sortOrderLabels, valueEntries, valuesCriteria } from "./sortFilterLabels";

/** Where the popover opens: the viewport point under its button. */
export interface PopoverAnchor {
  left: number;
  top: number;
}

export const PRESERVED_FILTER_MESSAGE = "Filtered by colour or icon — this filter can't be edited here";

export interface FilterPopoverDeps {
  model: () => SpreadsheetModel | null;
  /** The grid sheet on show. */
  sheet: () => Sheet | null;
  formatContext: () => FormatContext;
  locale: () => SpreadsheetLocale;
  /** Runs a filter command as ONE undo step (useSortFilter's). */
  filterStep: (label: string, mutate: (m: SpreadsheetModel, sheetId: string, ctx: FormatContext) => FilterCommandResult) => boolean;
  /** Sorts the autofilter's data rows by absolute column `col`. */
  sortFilterRange: (col: number, descending: boolean) => void;
  /** Gives the grid focus — first making `at` the active cell when given. */
  focusGrid: (at: CellRef | null) => void;
}

const filterOf = (sheet: Sheet | null) => (sheet?.kind === "grid" ? sheet.autoFilter : undefined);
const criteriaOf = (sheet: Sheet | null, colId: number): FilterCriteria | undefined =>
  filterOf(sheet)?.columns.find((column) => column.colId === colId)?.criteria;

export function useFilterPopover(deps: FilterPopoverDeps) {
  /** The open popover: its sheet, column (offset) and anchor. */
  const state = shallowRef<{ sheetId: string; colId: number; anchor: PopoverAnchor } | null>(null);
  /** The custom-filter dialog (null: closed). */
  const custom = shallowRef<{ colId: number; kind: ColumnKind; title: string; draft: CustomDraft } | null>(null);
  /** The Top 10 dialog (null: closed). */
  const top10 = shallowRef<{ colId: number; title: string } | null>(null);

  /** What the popover shows (null: closed, or its column no longer filters). */
  const view = computed(() => {
    const shown = state.value;
    const m = deps.model();
    const sheet = deps.sheet();
    const filter = filterOf(sheet);
    if (!shown || !m || !sheet || !filter || sheet.id !== shown.sheetId || shown.colId > filter.range.right - filter.range.left) {
      return null;
    }
    const ctx = deps.formatContext();
    const col = filter.range.left + shown.colId;
    const criteria = criteriaOf(sheet, shown.colId);
    const preserved = criteria?.kind === "preserved" || filter.odsUnmappedFilter !== undefined;
    const kind = columnValueKind(m, sheet, col, { top: filter.range.top + 1, bottom: filter.range.bottom }, ctx);
    const list = preserved ? { items: [], blanks: null, truncated: false } : filterValueList(m, sheet, shown.colId, ctx);
    const [asc, desc] = sortOrderLabels(kind !== "text");
    return {
      colId: shown.colId,
      col,
      header: { row: filter.range.top, col },
      title: headerLabel(m, sheet, filter.range.top, col, ctx),
      kind,
      conditionsLabel: CONDITIONS_LABELS[kind],
      conditions: conditionItems(kind).map((item) => ({ id: item.id, label: item.label })),
      sortAscLabel: `Sort ${asc}`,
      sortDescLabel: `Sort ${desc}`,
      entries: valueEntries(list, criteria),
      truncated: list.truncated,
      hasCriteria: criteria !== undefined || filter.odsUnmappedFilter !== undefined,
      preserved,
      criteria,
      anchor: shown.anchor,
    };
  });

  /** Opens the popover on column offset `colId` (a second press closes it). */
  function open(colId: number, anchor: PopoverAnchor): void {
    const sheet = deps.sheet();
    if (!sheet || !filterOf(sheet)) {
      return;
    }
    state.value = state.value?.colId === colId && state.value.sheetId === sheet.id ? null : { sheetId: sheet.id, colId, anchor };
  }

  /** Closes it; Escape returns focus to the header cell, a press elsewhere
   * leaves focus where it went. */
  function close(how: "escape" | "done" | "outside" = "done"): void {
    const header = view.value?.header ?? null;
    state.value = null;
    if (how !== "outside") {
      deps.focusGrid(how === "escape" ? header : null);
    }
  }

  /** A criterion for the open column as one "Filter column" step. */
  function setCriteria(colId: number, criteria: FilterCriteria | undefined, label = "Filter column"): void {
    deps.filterStep(label, (m, sheetId, ctx) => setFilterCriteria(m, sheetId, colId, criteria, ctx));
  }

  /** OK: the checked entries' keys. */
  function apply(checked: readonly string[]): void {
    const current = view.value;
    close();
    if (!current || current.preserved) {
      return;
    }
    const next = valuesCriteria(current.entries, new Set(checked), current.criteria);
    if (next !== "unchanged") {
      setCriteria(current.colId, next);
    }
  }

  function clear(): void {
    const current = view.value;
    close();
    if (current?.hasCriteria) {
      setCriteria(current.colId, undefined, "Clear filter");
    }
  }

  function sort(descending: boolean): void {
    const current = view.value;
    close();
    if (current) {
      deps.sortFilterRange(current.col, descending);
    }
  }

  /** A conditions submenu item: a dialog pre-set, or a dynamic criterion. */
  function condition(id: string): void {
    const current = view.value;
    const item = current ? conditionItems(current.kind).find((each) => each.id === id) : undefined;
    if (!current || !item) {
      return;
    }
    state.value = null;
    switch (item.action.kind) {
      case "custom": {
        const existing = id === "custom" ? customDraftOf(current.criteria, deps.locale()) : null;
        custom.value = { colId: current.colId, kind: current.kind, title: current.title, draft: existing ?? item.action.draft };
        break;
      }
      case "top10":
        top10.value = { colId: current.colId, title: current.title };
        break;
      case "dynamic":
        setCriteria(current.colId, { kind: "dynamic", type: item.action.type });
        deps.focusGrid(null);
        break;
    }
  }

  /** Focus goes back to the grid once a dialog has unmounted. */
  const afterDialog = (): void => void nextTick(() => deps.focusGrid(null));

  function confirmCustom(draft: CustomDraft): void {
    const dialog = custom.value;
    const criteria = dialog ? customCriteria(draft, deps.locale(), deps.formatContext().dateSystem) : null;
    if (!dialog || !criteria) {
      return;
    }
    custom.value = null;
    setCriteria(dialog.colId, criteria);
    afterDialog();
  }

  function confirmTop10(choice: { top: boolean; percent: boolean; value: string }): void {
    const dialog = top10.value;
    const criteria = dialog ? top10Criteria(choice.top, choice.percent, choice.value) : null;
    if (!dialog || !criteria) {
      return;
    }
    top10.value = null;
    setCriteria(dialog.colId, criteria);
    afterDialog();
  }

  // A different sheet on show closes everything.
  watch(
    () => deps.sheet()?.id,
    () => {
      state.value = null;
      custom.value = null;
      top10.value = null;
    },
  );

  return {
    view,
    /** The column offset whose popover is open (its button's aria-expanded). */
    openColId: computed(() => (view.value ? view.value.colId : null)),
    open,
    close,
    apply,
    clear,
    sort,
    condition,
    custom,
    confirmCustom,
    cancelCustom(): void {
      custom.value = null;
      afterDialog();
    },
    top10,
    confirmTop10,
    cancelTop10(): void {
      top10.value = null;
      afterDialog();
    },
  };
}

export type FilterPopover = ReturnType<typeof useFilterPopover>;
