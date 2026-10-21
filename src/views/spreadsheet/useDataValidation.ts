// Data validation at entry, and its grid chrome (SPEC-130 "Entry gate", the
// in-cell dropdown and the input message): THIS composable is the view's
// entry gate and the state of the three surfaces the rules show —
// ValidationAlert (WorkbookBody), ValidationInputMessage and
// ValidationDropdown (SheetGrid's validation layer).
//
// - The gate: `checkEntry(targets, candidate)` resolves "commit" at once when
//   every target accepts the candidate (validation.ts's validateValue, in the
//   user's locale, over stored values), and otherwise opens the alert for the
//   FIRST failing target — in the order given, which for Ctrl/Cmd+Enter's
//   fill is the selection's (`fillTargets`, sparse: only the cells a rule
//   covers) — and resolves with the user's answer. A rule whose `error.show`
//   is false never alerts (the entry is committed silently, as in Excel). A
//   formula candidate is checked by its result evaluated at each target
//   (shifted from `origin` there, as the fill writes it); one that cannot be
//   computed is accepted. useCellEditing's commit stays synchronous for a
//   valid entry: it asks `firstFailure` first and only waits on `alert` when
//   there is one (`checkEntry` is exactly the two together).
// - The alert: `validation-alert`, its title `error.title` or "Invalid
//   entry", its text `error.text` or the default message. Stop: Retry →
//   "retry", Cancel → "cancel"; warning: Yes → "commit", No → "retry",
//   Cancel → "cancel"; information: OK → "commit", Cancel → "cancel";
//   Escape is Cancel.
// - The input message: the active cell's rule's `input` while it shows
//   (`input.show`, a title or a text) and the sheet's prompts are not
//   disabled — hidden while the dropdown's list is open.
// - The dropdown: the button while the active cell's rule is a `list` with
//   `showDropdown`; `openList` (the button, or Alt+Down from the grid —
//   SPEC-128's capture listener keeps Alt+Down on an autofilter header cell)
//   lists `listItems` with the cell's current item selected (none when its
//   value is not in the list); `pick` enters the item's VALUE through the
//   cell editor's commit path (protection, this gate, one undo step) and
//   returns focus to the grid. Another cell becoming active, or another
//   sheet, closes the list.
//
// Nothing else validates (validation.ts's header): paste, fill, drag-move,
// clearing, undo/redo, sort and CSV import never reach this gate.
import { computed, shallowRef, watch, type Ref, type ShallowRef } from "vue";
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import { fillTargets as cellFillTargets } from "../../lib/spreadsheet/cells";
import { storedValueContext } from "../../lib/spreadsheet/formula/evaluate";
import { shiftFormula } from "../../lib/spreadsheet/formula/references";
import { tokenize } from "../../lib/spreadsheet/formula/tokens";
import { buildMergeIndex, mergeAt } from "../../lib/spreadsheet/merges";
import { MAX_COLS, type CellValue, type Sheet, type SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { workbookFormatContext } from "../../lib/spreadsheet/numberFormat/render";
import {
  formulaCandidateValue,
  itemIndexOf,
  listItems,
  sheetValidations,
  storedValueAt,
  validateValue,
  validationAt,
  type DataValidation,
  type ValidationDeps,
  type ValidationErrorStyle,
  type ValidationListItem,
} from "../../lib/spreadsheet/validation";
import type { SheetSelectionController } from "./useSheetSelection";

/** What the gate checks: the entry's value, or its formula (canonical, as
 * written at the edited cell — `origin`). */
export interface EntryCandidate {
  value: CellValue;
  formula?: string;
}

/** The user's answer to the alert (or "commit" when nothing failed). */
export type EntryVerdict = "commit" | "retry" | "cancel";

/** The first target a candidate fails at, and the rule it breaks. */
export interface ValidationFailure {
  at: CellRef;
  rule: DataValidation;
}

/** The alert's buttons. */
export type ValidationAlertChoice = "retry" | "cancel" | "yes" | "no" | "ok";

/** What `validation-alert` shows. */
export interface ValidationAlertState {
  style: ValidationErrorStyle;
  title: string;
  text: string;
}

/** The open dropdown list. */
export interface ValidationListState {
  sheetId: string;
  at: CellRef;
  items: ValidationListItem[];
  unavailable: boolean;
  /** The selected item's index (-1: none). */
  selected: number;
}

export const DEFAULT_ALERT_TITLE = "Invalid entry";
export const DEFAULT_ALERT_TEXT = "This value doesn't match the data validation rules for this cell.";
export const LIST_UNAVAILABLE_TEXT = "The list source isn't available";

/** The verdict each alert button gives (see the module header). */
export const ALERT_VERDICTS: Readonly<Record<ValidationAlertChoice, EntryVerdict>> = {
  retry: "retry",
  cancel: "cancel",
  yes: "commit",
  no: "retry",
  ok: "commit",
};

/** SPEC-130's entry gate as useCellEditing's commit path uses it. */
export interface EntryGate {
  /** The first of `targets` (on `sheet`) refusing the candidate under a rule
   * that shows its alert, or null — synchronous, so a valid entry commits at
   * once. `origin`: where a formula candidate is written (each target
   * evaluates it shifted from there). */
  firstFailure(sheet: Sheet, targets: Iterable<CellRef>, candidate: EntryCandidate, origin?: CellRef): ValidationFailure | null;
  /** The cells of a Ctrl/Cmd+Enter fill a rule covers, in fill order. */
  fillTargets(sheet: Sheet, ranges: readonly GridRange[], origin: CellRef): Iterable<CellRef>;
  /** Opens the alert for a failure; resolves with the user's answer. */
  alert(failure: ValidationFailure): Promise<EntryVerdict>;
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

/** The deps validateValue and listItems evaluate with: stored values, the
 * user's locale. */
export function validationDepsFor(model: SpreadsheetModel, locale: SpreadsheetLocale): ValidationDeps {
  return { base: storedValueContext(model, { locale }), ctx: workbookFormatContext(model, locale) };
}

/** Whether a rule's verdict is the same at every cell it covers for a plain
 * value: no `custom` overlay, and its formulas hold only constants and
 * absolute references (no relative part, name or function — ROW() and a
 * name may depend on the cell). A fill checks such a rule once. */
export function positionIndependent(rule: DataValidation): boolean {
  if (rule.type === "custom") {
    return false;
  }
  return [rule.formula1, rule.formula2].every(
    (formula) =>
      formula === undefined ||
      tokenize(formula).every((token) => {
        if (token.kind === "ref") {
          const ref = token.ref;
          const parts = ref ? [ref.start, ...(ref.end ? [ref.end] : [])] : [];
          return parts.every((part) => (part.row === undefined || part.rowAbs) && (part.col === undefined || part.colAbs));
        }
        return token.kind !== "name" && token.kind !== "function" && token.kind !== "structured" && token.kind !== "external" && token.kind !== "unknown";
      }),
  );
}

/** The cells of a Ctrl/Cmd+Enter fill a rule covers (sparse: only the
 * rules' rectangles intersected with what the fill writes — lib's
 * fillTargets — never every cell of the selection), in fill order: range by
 * range, rows top to bottom, columns left to right; a merge's covered cells
 * and a cell already listed skipped, as the fill skips them. */
export function* validatedFillTargets(sheet: Sheet, ranges: readonly GridRange[], origin: CellRef): Generator<CellRef> {
  const rules = sheetValidations(sheet);
  if (rules.length === 0) {
    return;
  }
  const merges = buildMergeIndex(sheet.merges);
  const seen = new Set<number>();
  for (const range of cellFillTargets(sheet, ranges, origin)) {
    const pieces: GridRange[] = [];
    for (const rule of rules) {
      for (const covered of rule.ranges) {
        const top = Math.max(range.top, covered.top);
        const bottom = Math.min(range.bottom, covered.bottom);
        const left = Math.max(range.left, covered.left);
        const right = Math.min(range.right, covered.right);
        if (top <= bottom && left <= right) {
          pieces.push({ top, left, bottom, right });
        }
      }
    }
    const byTop = pieces.toSorted((a, b) => a.top - b.top);
    let row = byTop[0]?.top ?? 0;
    while (byTop.length > 0) {
      const spans = byTop.filter((piece) => piece.top <= row && row <= piece.bottom).toSorted((a, b) => a.left - b.left);
      if (spans.length === 0) {
        const next = byTop.find((piece) => piece.top > row);
        if (!next) {
          break;
        }
        row = next.top;
        continue;
      }
      let from = 0;
      for (const span of spans) {
        for (let col = Math.max(from, span.left); col <= span.right; col += 1) {
          const key = row * MAX_COLS + col;
          const merge = mergeAt(merges, row, col);
          if (seen.has(key) || (merge && (merge.top !== row || merge.left !== col))) {
            continue;
          }
          seen.add(key);
          yield { row, col };
        }
        from = Math.max(from, span.right + 1);
      }
      row += 1;
    }
  }
}

/** The first of `targets` refusing `candidate` under a rule that shows its
 * alert, or null (see the module header). */
export function firstValidationFailure(
  model: SpreadsheetModel,
  sheet: Sheet,
  targets: Iterable<CellRef>,
  candidate: EntryCandidate,
  deps: ValidationDeps,
  origin?: CellRef,
): ValidationFailure | null {
  if (sheetValidations(sheet).length === 0) {
    return null;
  }
  // A plain value under a position-independent rule is checked once per rule.
  const settled = new Map<string, boolean>();
  for (const at of targets) {
    const rule = validationAt(sheet, at);
    if (!rule || rule.type === "any" || !rule.error.show) {
      continue;
    }
    const memo = candidate.formula === undefined && positionIndependent(rule);
    let ok = memo ? settled.get(rule.id) : undefined;
    if (ok === undefined) {
      let value: CellValue | undefined = candidate.value;
      if (candidate.formula !== undefined) {
        const formula = origin ? shiftFormula(candidate.formula, at.row - origin.row, at.col - origin.col) : candidate.formula;
        value = formulaCandidateValue(sheet, at, formula, deps);
      }
      ok = value === undefined || validateValue(model, sheet, at, value, deps).ok;
      if (memo) {
        settled.set(rule.id, ok);
      }
    }
    if (!ok) {
      return { at: { ...at }, rule };
    }
  }
  return null;
}

/** The alert a failing rule shows. */
export function alertFor(rule: DataValidation): ValidationAlertState {
  return {
    style: rule.error.style,
    title: rule.error.title.trim() === "" ? DEFAULT_ALERT_TITLE : rule.error.title,
    text: rule.error.text.trim() === "" ? DEFAULT_ALERT_TEXT : rule.error.text,
  };
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

/** Narrows a component's runtime-typed `validation` prop (the .vue parser
 * bars casts). */
export function asDataValidation(value: unknown): DataValidationView {
  return value as DataValidationView;
}

/** The grid context's `validation`, when the view wired one (null: none). */
export function dataValidationOf(context: unknown): DataValidationView | null {
  const value = context !== null && typeof context === "object" ? Reflect.get(context, "validation") : undefined;
  return value ? (value as DataValidationView) : null;
}
