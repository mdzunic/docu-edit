// Formulas on screen (SPEC-119 "The formula bar", "Show formulas"): the
// formula bar's text and View ▸ Show formulas for the sheet on show.
//
// - The bar shows lib's formulaBarText for the active cell — "=" + formula,
//   an array's "{=…}", a data table's "{=TABLE(…)}", an unparsed formula as
//   stored, SPEC-118's editText for every other cell — in the view's locale
//   through the cell's effective style (useSheetSelection's format context).
//   A dynamic array's member shows its anchor's formula with `spillMember`,
//   which the bar renders muted. SPEC-122 makes the bar editable (the cell
//   editor's draft replaces this text while a cell is edited) and honours
//   protection: a cell whose effective style hides its formula
//   (`protection.hidden`) on a protected sheet shows an empty bar.
// - Show formulas is the sheet's `showFormulas` view state
//   (useSheetViewState): toggled from the View menu or with Ctrl+` — matched
//   on `event.code === "Backquote"` so a layout where ` is a dead key still
//   works, with Ctrl and no other modifier (Ctrl+Shift+` is another binding),
//   on every platform: macOS keeps Cmd+` for cycling windows and never hands
//   it to the page, and Excel for Mac uses Control+` too. A key from an input,
//   textarea, select or dialog is not a shortcut (isOutsideDocumentInput,
//   SPEC-021). Toggling is view state only: it never marks the workbook
//   dirty, and an untouched workbook still downloads byte-identical.
import { computed, onScopeDispose } from "vue";
import { isOutsideDocumentInput } from "../../lib/shortcuts";
import { formulaBarContent, type FormulaBarContent } from "../../lib/spreadsheet/formula/print";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";

export interface FormulaDisplayDeps {
  sheet: () => Sheet | null;
  /** The workbook (its format decides how an unparsed formula shows). */
  model: () => SpreadsheetModel | null;
  viewState: Pick<SheetViewState, "active" | "updateActive">;
  selection: Pick<SheetSelectionController, "selection" | "formatContext" | "styleAt">;
}

/** Ctrl+` — by physical key, with Ctrl and no other modifier, on every
 * platform. */
export function isShowFormulasShortcut(event: Pick<KeyboardEvent, "code" | "ctrlKey" | "shiftKey" | "altKey" | "metaKey">): boolean {
  return event.code === "Backquote" && event.ctrlKey && !event.shiftKey && !event.altKey && !event.metaKey;
}

export function useFormulaDisplay(deps: FormulaDisplayDeps) {
  /** Whether the sheet on show shows formulas instead of values. */
  const showFormulas = computed(() => deps.viewState.active.value.showFormulas === true);

  /** View ▸ Show formulas: on ↔ off for the sheet on show, as view state. */
  function toggle(): void {
    if (deps.sheet()?.kind === "grid") {
      deps.viewState.updateActive({ showFormulas: showFormulas.value ? undefined : true });
    }
  }

  function onKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented || !isShowFormulasShortcut(event) || isOutsideDocumentInput(event.target) || deps.sheet()?.kind !== "grid") {
      return;
    }
    event.preventDefault();
    toggle();
  }

  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown);
    }
  });

  /** The formula bar's text for the active cell, and whether it is muted. */
  const bar = computed((): FormulaBarContent => {
    const sheet = deps.sheet();
    if (!sheet || sheet.kind !== "grid") {
      return { text: "", spillMember: false };
    }
    const { row, col } = deps.selection.selection.value.active;
    const ctx = { ...deps.selection.formatContext.value, format: deps.model()?.format ?? "xlsx" };
    const style = deps.selection.styleAt(sheet, row, col);
    if (sheet.protection && style?.protection?.hidden) {
      return { text: "", spillMember: false };
    }
    return formulaBarContent(sheet, row, col, style, ctx);
  });

  return {
    showFormulas,
    toggle,
    onKeydown,
    barText: computed(() => bar.value.text),
    spillMember: computed(() => bar.value.spillMember),
  };
}

export type FormulaDisplay = ReturnType<typeof useFormulaDisplay>;
