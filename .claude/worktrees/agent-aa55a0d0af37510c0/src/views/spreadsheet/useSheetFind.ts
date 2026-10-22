// Find and replace in the workbook view (SPEC-134 "View wiring"): the state
// behind SPEC-017's FindBar (reused through its props, `options` slot and
// events — `bar` / `barEvents`), the marks SheetGrid paints (`marks`), Find
// all's results (`results`) and the keys. lib/spreadsheet/find.ts decides
// what matches, in what order, and what a replace writes; sheetFindApply.ts
// enters a replace through SPEC-122's cell entry as one undo step.
//
// - Keys (window, capture; sheetFindKeys.ts decides): Cmd/Ctrl+F opens the
//   bar and Cmd/Ctrl+H with the replace row (preventDefault, so the browser's
//   find never opens over the grid); on macOS, while the bar is open,
//   Cmd+G / Cmd+Shift+G step next / previous; Escape on the bar closes it.
//   While the cell editor is open they are consumed and do nothing.
// - The walk: the query and every option debounce FIND_DEBOUNCE_MS, then a
//   createSheetSearch runs one scheduleEagerWork slot at a time, each until
//   FIND_SLICE_BUDGET_MS, the count showing the running total with
//   "Searching…" until done. A model change (a sheet's `rows` reference or
//   the sheet list) or a scope change (Within, the active sheet, a multi-cell
//   selection's ranges) restarts it after the same debounce.
// - Navigation: a query or option change makes the current hit the first at
//   or after the active cell and selects it; Enter / find-next the first
//   after the active cell, Shift+Enter / find-prev the last before it, either
//   wrapping with "Wrapped to top"; a request the walk cannot answer yet
//   waits for it. A hit becomes the active cell through sheetJump.ts (the
//   sheet switched first in workbook scope; inside a searched selection its
//   ranges kept), is revealed, and is never an undo entry; one in a hidden
//   row or column shows FIND_HIDDEN_MATCH_HINT.
// - Replace: the current cell per "Replace rules" (one step "Replace"), then
//   on to the next match; a refusal leaves the cell and the current match
//   where they are and says why. Replace all: every hit in scope (the walk
//   finished first) as ONE step "Replace all" — scope "workbook" for Within
//   = Workbook, else the active sheet — reported by replaceReport.
// - Every option persists for the view session; closing clears the marks and
//   the results and hands focus back to the grid.
import { computed, nextTick, onScopeDispose, ref, watch, type Ref, type ShallowRef } from "vue";
import { isMacPlatform } from "../../lib/core/platform";
import { scheduleEagerWork, type IdleHandle } from "../../lib/core/idleScheduler";
import { formatA1 } from "../../lib/spreadsheet/address";
import type { CellTextMeasurer } from "../../lib/spreadsheet/autofit";
import { ARRAY_PART_MESSAGE, PROTECTED_CELL_MESSAGE } from "../../lib/spreadsheet/cells";
import {
  breaksFormulaMessage,
  compileCellQuery,
  createSheetSearch,
  FIND_HIDDEN_MATCH_HINT,
  FIND_PLACEHOLDERS,
  FIND_SCOPE_LABEL,
  FIND_SCOPE_TITLE,
  FIND_SEARCHING_HINT,
  FIND_WRAPPED_HINT,
  hitIsHidden,
  planReplace,
  replaceReport,
  searchesSelection,
  type FindHit,
  type FindLookIn,
  type FindWithin,
  type SheetFindOptions,
} from "../../lib/spreadsheet/find";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { createCanvasCellMeasurer } from "./cellMeasurer";
import { applyReplacePlan } from "./sheetFindApply";
import { findKeyAction } from "./sheetFindKeys";
import { findResultRows } from "./sheetFindRows";
import { createFindWalk } from "./sheetFindWalk";
import { findMarkKey, type FindMarks } from "./sheetLayerProps";
import { createSheetJump } from "./sheetJump";
import type { CellEditing } from "./useCellEditing";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

/** A refused single Replace's reasons other than a broken formula. */
export const FORMULA_RESULT_MESSAGE = "Formula results can't be replaced — look in formulas instead";
export const FIND_UNDO_LABELS = { replace: "Replace", replaceAll: "Replace all" } as const;

export interface SheetFindDeps {
  book: {
    workbook: ShallowRef<SpreadsheetModel | null>;
    activeIndex: Ref<number>;
    activeSheet: Readonly<Ref<Sheet | null>>;
    selectSheet(index: number): void;
  };
  viewState: Pick<SheetViewState, "update" | "syncInto">;
  selection: Pick<SheetSelectionController, "selection" | "formatContext">;
  editing: { undo: Pick<SpreadsheetUndo, "record" | "commit">; cell: Pick<CellEditing, "editing" | "focusCell" | "revealTick"> };
  locale: () => SpreadsheetLocale;
  /** Test seams: the slot scheduler, the clock, the platform, the measurer. */
  schedule?: (callback: (budgetMs: number) => void, budgetMs: number) => IdleHandle;
  now?: () => number;
  isMac?: boolean;
  measurer?: CellTextMeasurer;
}

export function useSheetFind(deps: SheetFindDeps) {
  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const jump = createSheetJump({ book: deps.book, viewState: deps.viewState, cell: deps.editing.cell });
  let measurer = deps.measurer;

  const open = ref(false);
  const expanded = ref(false);
  const query = ref("");
  const replaceText = ref("");
  const matchCase = ref(false);
  const entireCell = ref(false);
  const regex = ref(false);
  const lookIn = ref<FindLookIn>("values");
  const within = ref<FindWithin>("sheet");
  const resultsOpen = ref(false);
  /** A transient line: a wrap, a replace report, a refusal. */
  const message = ref("");
  const currentIndex = ref(-1);
  /** Bumped whenever the walk's hits or state change. */
  const revision = ref(0);

  const options = computed((): SheetFindOptions => ({ matchCase: matchCase.value, entireCell: entireCell.value, regex: regex.value, lookIn: lookIn.value, within: within.value }));
  const compiled = computed(() => compileCellQuery(query.value, options.value));
  const error = computed(() => (compiled.value !== null && "error" in compiled.value ? compiled.value.error : ""));
  const activeCell = (): FindHit | null => {
    const sheet = deps.book.activeSheet.value;
    return sheet ? { sheetId: sheet.id, ...deps.selection.selection.value.active } : null;
  };

  // --- the walk (sheetFindWalk.ts) -------------------------------------------------------------
  const inSelection = (): boolean => {
    const sheet = deps.book.activeSheet.value;
    return within.value === "sheet" && sheet !== null && searchesSelection(sheet, deps.selection.selection.value);
  };
  /** Whether the running walk searches the selection (navigation keeps its ranges). */
  let walkedSelection = false;
  const walk = createFindWalk({
    create: () => {
      const m = model();
      const sheet = deps.book.activeSheet.value;
      const q = compiled.value;
      walkedSelection = inSelection();
      currentIndex.value = -1;
      return open.value && m && sheet && q !== null && !("error" in q)
        ? createSheetSearch(m, q, options.value, sheet.id, deps.selection.selection.value, deps.selection.formatContext.value)
        : null;
    },
    key: () => [
      within.value,
      deps.book.activeSheet.value?.id,
      inSelection() ? JSON.stringify(deps.selection.selection.value.ranges) : "",
      ...(model()?.sheets ?? []).flatMap((sheet) => [sheet.rows, `${sheet.id}:${sheet.kind}:${sheet.visibility}`]),
    ],
    from: activeCell,
    changed: () => void (revision.value += 1),
    answered: (answer) => {
      if (answer.kind === "select") {
        message.value = answer.wrapped ? FIND_WRAPPED_HINT : message.value;
        void select(answer.index);
      } else {
        currentIndex.value = answer.kind === "keep" ? answer.index : -1;
      }
    },
    schedule: deps.schedule ?? scheduleEagerWork,
    now: deps.now ?? ((): number => performance.now()),
  });

  /** Makes hit `index` current and the active cell (see the header). */
  async function select(index: number): Promise<void> {
    const hit = walk.search?.hits[index];
    if (!hit) {
      return;
    }
    currentIndex.value = index;
    const ranges = walkedSelection ? deps.selection.selection.value.ranges : [{ top: hit.row, left: hit.col, bottom: hit.row, right: hit.col }];
    await jump(hit.sheetId, ranges, { active: { row: hit.row, col: hit.col }, focus: false });
  }

  function navigate(mode: "after" | "before"): void {
    const from = activeCell();
    if (open.value && error.value === "" && from && compiled.value !== null) {
      message.value = "";
      walk.ask(mode, from);
    }
  }

  watch([query, options], () => {
    message.value = "";
    if (open.value) {
      walk.soon("atOrAfter");
    }
  });
  watch([() => model()?.sheets, () => deps.book.activeSheet.value, () => deps.selection.selection.value.ranges], () => {
    if (open.value) {
      walk.soon();
    }
  });

  // --- the bar -------------------------------------------------------------------------------
  function openFind(withReplace = false): void {
    if (withReplace) {
      expanded.value = true;
    }
    if (!open.value) {
      open.value = true;
      walk.begin();
    } else if (!withReplace && typeof document !== "undefined") {
      // Already open: back to the query, selected (the bar focuses it only on opening).
      void nextTick(() => {
        const field = document.querySelector<HTMLInputElement>('[data-testid="find-query"]');
        field?.focus();
        field?.select();
      });
    }
  }

  function close(): void {
    open.value = false;
    resultsOpen.value = false;
    walk.clear();
    currentIndex.value = -1;
    message.value = "";
    deps.editing.cell.focusCell();
  }

  /** The walk's hits so far — a fresh array each revision, so dependents follow. */
  const hits = computed((): readonly FindHit[] => {
    void revision.value;
    return [...(walk.search?.hits ?? [])];
  });
  const current = computed(() => hits.value[currentIndex.value] ?? null);
  const searching = computed(() => {
    void revision.value;
    return walk.search !== null && !walk.search.done;
  });
  const hint = computed(() => {
    const m = model();
    if (message.value !== "") {
      return message.value;
    }
    if (searching.value) {
      return FIND_SEARCHING_HINT;
    }
    return m && current.value && hitIsHidden(m, current.value) ? FIND_HIDDEN_MATCH_HINT : "";
  });

  const bar = computed(() => ({
    visible: open.value,
    expanded: expanded.value,
    query: query.value,
    replaceText: replaceText.value,
    matchCase: matchCase.value,
    wholeWords: false,
    useRegex: regex.value,
    matchCount: hits.value.length,
    currentIndex: currentIndex.value,
    error: error.value,
    hint: hint.value,
    placeholder: FIND_PLACEHOLDERS[within.value],
    showWholeWords: false,
    scopeLabel: FIND_SCOPE_LABEL,
    scopeTitle: FIND_SCOPE_TITLE,
  }));

  /** The marks on the sheet on show (null: the bar is closed). */
  const marks = computed((): FindMarks | null => {
    const sheet = deps.book.activeSheet.value;
    if (!open.value || !sheet) {
      return null;
    }
    const matches = new Set<number>();
    for (const each of hits.value) {
      if (each.sheetId === sheet.id) {
        matches.add(findMarkKey(each.row, each.col));
      }
    }
    const here = current.value;
    return { sheetId: sheet.id, matches, current: here && here.sheetId === sheet.id ? findMarkKey(here.row, here.col) : null };
  });

  // --- Find all ------------------------------------------------------------------------------
  const results = computed(() => {
    const m = model();
    return resultsOpen.value && m ? findResultRows(m, hits.value, deps.selection.formatContext.value) : null;
  });

  // --- replace -------------------------------------------------------------------------------
  function write(plan: ReturnType<typeof planReplace>, label: string, scope: "workbook" | { sheetId: string }): number {
    const m = model()!;
    deps.viewState.syncInto(m);
    measurer ??= createCanvasCellMeasurer();
    return applyReplacePlan({ model: m, undo: deps.editing.undo, locale: deps.locale(), measurer, format: deps.selection.formatContext.value }, plan, label, scope);
  }

  function replaceCurrent(): void {
    const m = model();
    const q = compiled.value;
    const hit = current.value;
    if (!m || !hit || q === null || "error" in q || deps.editing.cell.editing.value) {
      return;
    }
    const plan = planReplace(m, [hit], q, replaceText.value, options.value, deps.selection.formatContext.value);
    if (plan.edits.length === 0) {
      const { breaksFormula, protected: locked, array } = plan.skipped;
      message.value = breaksFormula ? breaksFormulaMessage(formatA1(hit)) : locked ? PROTECTED_CELL_MESSAGE : array ? ARRAY_PART_MESSAGE : FORMULA_RESULT_MESSAGE;
      return;
    }
    message.value = "";
    write(plan, FIND_UNDO_LABELS.replace, { sheetId: hit.sheetId });
    walk.begin("after", hit);
  }

  function replaceAll(): void {
    const m = model();
    const q = compiled.value;
    const sheet = deps.book.activeSheet.value;
    if (!m || !sheet || q === null || "error" in q || deps.editing.cell.editing.value) {
      return;
    }
    // Over the current query, options and model: a fresh walk, finished now.
    walk.begin();
    walk.search?.step(() => false);
    const plan = planReplace(m, walk.search?.hits ?? [], q, replaceText.value, options.value, deps.selection.formatContext.value);
    const written = write(plan, FIND_UNDO_LABELS.replaceAll, within.value === "workbook" ? "workbook" : { sheetId: sheet.id });
    walk.begin();
    message.value = replaceReport(plan, written);
  }

  // --- keys (sheetFindKeys.ts) --------------------------------------------------------------
  function onKeydown(event: KeyboardEvent): void {
    const action = findKeyAction(event, open.value, deps.isMac ?? isMacPlatform());
    if (action === null) {
      return;
    }
    event.preventDefault();
    if (action === "close") {
      close();
    } else if (!deps.editing.cell.editing.value) {
      if (action === "next" || action === "prev") {
        navigate(action === "next" ? "after" : "before");
      } else {
        openFind(action === "openReplace");
      }
    }
  }
  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown, true);
  }
  onScopeDispose(() => {
    walk.stop();
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown, true);
    }
  });

  return {
    open,
    expanded,
    query,
    replaceText,
    matchCase,
    entireCell,
    regex,
    lookIn,
    within,
    hits,
    currentIndex,
    bar,
    /** FindBar's events, bound with v-on. */
    barEvents: {
      "update:query": (text: string): void => void (query.value = text),
      "update:replaceText": (text: string): void => void (replaceText.value = text),
      "update:matchCase": (on: boolean): void => void (matchCase.value = on),
      "update:useRegex": (on: boolean): void => void (regex.value = on),
      "toggle-expand": (): void => void (expanded.value = !expanded.value),
      next: (): void => navigate("after"),
      prev: (): void => navigate("before"),
      replace: replaceCurrent,
      "replace-all": replaceAll,
      close,
    },
    marks,
    results,
    resultsOpen,
    findAll: (): void => void ((resultsOpen.value = true), open.value && !walk.search && walk.begin()),
    /** A Find all row activated: that hit, exactly as find-next selects it. */
    activate: (index: number): void => void select(index),
    openFind,
    close,
    onKeydown,
  };
}

export type SheetFind = ReturnType<typeof useSheetFind>;

/** Narrows a component's runtime-typed `find` prop (the .vue parser bars casts). */
export function asSheetFind(value: unknown): SheetFind {
  return value as SheetFind;
}

/** The grid context's `find`, when the view wired it (null: none). */
export function findOf(context: unknown): SheetFind | null {
  const value = context !== null && typeof context === "object" ? Reflect.get(context, "find") : undefined;
  return value ? (value as SheetFind) : null;
}
