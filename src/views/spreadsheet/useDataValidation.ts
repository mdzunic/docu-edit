// Data validation in the workbook view (SPEC-130 "Entry gate", the in-cell
// dropdown and the input message): THIS composable is the view's entry gate
// (validationGate.ts's checks, with the alert's state) and the state of the
// surfaces the rules show — ValidationAlert (WorkbookBody),
// ValidationInputMessage and ValidationDropdown (SheetGrid's validation
// layer).
//
// - The gate: `checkEntry(targets, candidate)` resolves "commit" at once when
//   every target accepts the candidate, and otherwise opens the alert for
//   the first failure and resolves with the user's answer. useCellEditing's
//   commit path stays synchronous for a valid entry: it asks `firstFailure`
//   and only waits on `openAlert` when there is one (`checkEntry` is exactly
//   the two together).
// - The input message: the active cell's rule's `input` while it shows
//   (`input.show`, a title or a text) and the sheet's prompts are not
//   disabled — hidden while the dropdown's list is open.
// - The dropdown: the button while the active cell's rule is a `list` with
//   `showDropdown`; `openList` (the button, or Alt+Down from the grid —
//   SPEC-128's capture listener keeps Alt+Down on an autofilter header cell)
//   lists `listItems` with the cell's current item selected (none when its
//   value is not in the list); `pick` enters the item's VALUE through the
//   cell editor's commit path (protection, the gate, one undo step) and
//   returns focus to the grid. Another cell becoming active, or another
//   sheet, closes the list.
import { computed, shallowRef, watch, type Ref, type ShallowRef } from "vue";
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import type { CellValue, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { itemIndexOf, listItems, storedValueAt, validationAt, type DataValidation, type ValidationDeps, type ValidationListItem } from "../../lib/spreadsheet/validation";
import type { SheetSelectionController } from "./useSheetSelection";
import {
  alertFor,
  ALERT_VERDICTS,
  firstValidationFailure,
  validatedFillTargets,
  validationDepsFor,
  type EntryCandidate,
  type EntryGate,
  type EntryVerdict,
  type ValidationAlertChoice,
  type ValidationAlertState,
  type ValidationFailure,
} from "./validationGate";

/** The open dropdown list. */
export interface ValidationListState {
  sheetId: string;
  at: CellRef;
  items: ValidationListItem[];
  unavailable: boolean;
  /** The selected item's index (-1: none). */
  selected: number;
}

/** What a pick commits through (useCellEditing — which takes this
 * composable as its entry gate, so it is handed over lazily). */
export interface ValidationEditor {
  enterValue(value: CellValue, text: string): boolean;
  focusCell(): void;
  editing: Readonly<Ref<boolean>>;
  /** Commits an edit in progress; false when it was held (kept open). */
  commit(): boolean;
}

export interface DataValidationDeps {
  book: {
    workbook: ShallowRef<SpreadsheetModel | null>;
    activeSheet: Readonly<Ref<Sheet | null>>;
  };
  selection: Pick<SheetSelectionController, "selection">;
  locale: () => SpreadsheetLocale;
  /** The cell editor a pick enters through; absent = picks do nothing. */
  editor?: () => ValidationEditor | null;
}

export function useDataValidation(deps: DataValidationDeps) {
  const alert = shallowRef<ValidationAlertState | null>(null);
  let resolveAlert: ((verdict: EntryVerdict) => void) | null = null;
  const list = shallowRef<ValidationListState | null>(null);

  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const evalDeps = (m: SpreadsheetModel): ValidationDeps => validationDepsFor(m, deps.locale());

  function firstFailure(sheet: Sheet, targets: Iterable<CellRef>, candidate: EntryCandidate, origin?: CellRef): ValidationFailure | null {
    const m = model();
    return m ? firstValidationFailure(m, sheet, targets, candidate, evalDeps(m), origin) : null;
  }

  /** Opens the alert for `failure` (an alert still open is cancelled). */
  function openAlert(failure: ValidationFailure): Promise<EntryVerdict> {
    resolveAlert?.("cancel");
    alert.value = alertFor(failure.rule);
    return new Promise<EntryVerdict>((resolve) => {
      resolveAlert = resolve;
    });
  }

  /** The alert's button (Escape: "cancel"). */
  function answer(choice: ValidationAlertChoice): void {
    const resolve = resolveAlert;
    resolveAlert = null;
    alert.value = null;
    resolve?.(ALERT_VERDICTS[choice]);
  }

  /** SPEC-130's entry gate: "commit" when every target accepts the
   * candidate, else the user's answer to the alert for the first failure.
   * `targets` are on the sheet `options.sheetId` names (default: the sheet
   * on show). */
  function checkEntry(targets: Iterable<CellRef>, candidate: EntryCandidate, options: { sheetId?: string; origin?: CellRef } = {}): Promise<EntryVerdict> {
    const m = model();
    const sheet = options.sheetId === undefined ? deps.book.activeSheet.value : (m?.sheets.find((each) => each.id === options.sheetId) ?? null);
    const failure = sheet ? firstFailure(sheet, targets, candidate, options.origin) : null;
    return failure ? openAlert(failure) : Promise.resolve("commit");
  }

  const gridSheet = computed((): Sheet | null => {
    const sheet = deps.book.activeSheet.value;
    return sheet?.kind === "grid" ? sheet : null;
  });
  /** The active cell's rule on the sheet on show. */
  const activeRule = computed((): DataValidation | undefined => {
    const sheet = gridSheet.value;
    return sheet ? validationAt(sheet, deps.selection.selection.value.active) : undefined;
  });

  /** The input message to show for the active cell, or null. */
  const inputMessage = computed((): { title: string; text: string } | null => {
    const rule = activeRule.value;
    if (!rule || !rule.input.show || (rule.input.title === "" && rule.input.text === "") || gridSheet.value?.validationPromptsDisabled || list.value) {
      return null;
    }
    return { title: rule.input.title, text: rule.input.text };
  });

  /** The dropdown button for the active cell, or null. */
  const dropdown = computed((): { expanded: boolean } | null => {
    const rule = activeRule.value;
    return rule?.type === "list" && rule.showDropdown ? { expanded: list.value !== null } : null;
  });

  /** What ValidationDropdown's list binds: the items' texts, the selection,
   * whether the source is unavailable; null when closed. */
  const listView = computed(() => {
    const open = list.value;
    return open ? { texts: open.items.map((item) => item.text), selected: open.selected, unavailable: open.unavailable } : null;
  });

  /** Opens the active cell's list (an edit in progress commits first);
   * false when there is no dropdown here or the commit was held. */
  function openList(): boolean {
    const m = model();
    const sheet = gridSheet.value;
    const rule = activeRule.value;
    if (!m || !sheet || rule?.type !== "list" || !rule.showDropdown) {
      return false;
    }
    const editor = deps.editor?.();
    if (editor?.editing.value && !editor.commit()) {
      return false;
    }
    const at = { ...deps.selection.selection.value.active };
    const result = listItems(m, sheet, rule, at, evalDeps(m));
    const unavailable = result.unavailable === true;
    list.value = {
      sheetId: sheet.id,
      at,
      items: unavailable ? [] : result.items,
      unavailable,
      selected: unavailable ? -1 : itemIndexOf(result.items, storedValueAt(sheet, at)),
    };
    return true;
  }

  /** Closes the list; `focusGrid` returns focus to the grid's cell editor
   * (Escape, a pick), else focus is left where it went. */
  function closeList(focusGrid: boolean): void {
    if (list.value) {
      list.value = null;
    }
    if (focusGrid) {
      deps.editor?.()?.focusCell();
    }
  }

  /** The button: opens the list, or closes it when open. */
  function toggleList(): void {
    if (list.value) {
      closeList(true);
    } else {
      openList();
    }
  }

  /** Selects an item (the list's keys). */
  function selectItem(index: number): void {
    const open = list.value;
    if (open && index >= 0 && index < open.items.length && index !== open.selected) {
      list.value = { ...open, selected: index };
    }
  }

  /** Enters item `index`'s value into the list's cell through the commit
   * path and returns focus to the grid; false when nothing was entered. */
  function pick(index: number): boolean {
    const open = list.value;
    const item = open?.items[index];
    closeList(true);
    if (!item) {
      return false;
    }
    return deps.editor?.()?.enterValue(item.value, item.text) ?? false;
  }

  /** A key on the grid (the cell editor in Ready, or the grid itself):
   * Alt+Down opens the list. True when taken. */
  function onGridKey(event: KeyboardEvent): boolean {
    if (event.key !== "ArrowDown" || !event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) {
      return false;
    }
    return openList();
  }

  // Another active cell or sheet closes the list (at once: it never shows
  // over another cell).
  watch(
    () => {
      const active = deps.selection.selection.value.active;
      return `${deps.book.activeSheet.value?.id ?? ""}:${active.row}:${active.col}`;
    },
    () => closeList(false),
    { flush: "sync" },
  );

  return {
    checkEntry,
    firstFailure,
    fillTargets: (sheet: Sheet, ranges: readonly GridRange[], origin: CellRef): Iterable<CellRef> => validatedFillTargets(sheet, ranges, origin),
    alert,
    /** Opens the alert for a failure firstFailure found (checkEntry's second
     * half — useCellEditing's commit asks firstFailure synchronously). */
    openAlert,
    answer,
    inputMessage,
    dropdown,
    list,
    listView,
    openList,
    closeList,
    toggleList,
    selectItem,
    pick,
    onGridKey,
  };
}

export type DataValidationView = ReturnType<typeof useDataValidation>;

/** The view's entry gate over useDataValidation (useCellEditing's dep). */
export function entryGateOf(validation: DataValidationView): EntryGate {
  return { firstFailure: validation.firstFailure, fillTargets: validation.fillTargets, alert: validation.openAlert };
}
