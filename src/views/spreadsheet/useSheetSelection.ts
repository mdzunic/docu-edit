// The sheet on show's selection (SPEC-117 "Selection", "Keyboard map",
// "Pointer", "Aggregates"): the state the grid draws, the formula bar's Name
// Box and content, and the status bar's aggregates all read, and the actions
// the grid's keyboard and pointer handlers call. Every rule lives in the
// pure functions of lib/spreadsheet/selection.ts — this composable only
// holds the state, maps events onto them and stores the result as view state
// (useSheetViewState), so selecting never touches the document.
//
// - Keys: the primary modifier comes from core/platform.ts (Cmd on macOS,
//   Ctrl elsewhere); Ctrl+Space is literal Ctrl everywhere; a key whose
//   target is an input (isOutsideDocumentInput) is ignored; a key the map
//   does not handle is left alone (the caller calls preventDefault only for
//   consumed keys).
// - Pointer: press → the cell (snapped to its merge), Shift+press extends,
//   Mod+press adds a range, drag extends; header presses select whole rows or
//   columns the same way. The Name Box shows "<rows>R × <cols>C" from press
//   until release.
// - Aggregates are computed over stored cells only; a selection over
//   AGGREGATE_DEFER_THRESHOLD stored cells is computed in browser idle time
//   (core/idleScheduler.ts) rather than on the keystroke.
import { computed, shallowRef, watch } from "vue";
import { isMacPlatform } from "../../lib/core/platform";
import { runWhenIdle } from "../../lib/core/idleScheduler";
import { isOutsideDocumentInput } from "../../lib/shortcuts";
import { formatA1, getCell, type CellRef } from "../../lib/spreadsheet/address";
import {
  formatAggregates,
  selectionAggregates,
  selectionSpansMultipleCells,
  selectionStoredCellEstimate,
  type SelectionAggregates,
} from "../../lib/spreadsheet/aggregates";
import { generalText } from "../../lib/spreadsheet/general";
import type { Sheet } from "../../lib/spreadsheet/model";
import {
  addRange,
  extendTo,
  extendToLine,
  headerSelection,
  navigateByKey,
  normaliseSelection,
  selectAll,
  selectCell,
  type SheetSelection,
} from "../../lib/spreadsheet/selection";
import type { SheetMetrics } from "./useSheetMetrics";
import type { SheetViewState } from "./useSheetViewState";

/** Selections over this many stored cells aggregate in idle time. */
export const AGGREGATE_DEFER_THRESHOLD = 100_000;

/** Modifier state of a pointer press. */
export interface PressModifiers {
  shift: boolean;
  mod: boolean;
}

/** The primary modifier of an event on this platform. */
export function primaryModifier(event: { metaKey: boolean; ctrlKey: boolean }): boolean {
  return isMacPlatform() ? event.metaKey : event.ctrlKey;
}

export interface SheetSelectionDeps {
  sheet: () => Sheet | null;
  sheetMetrics: Pick<SheetMetrics, "nav" | "metrics" | "merges">;
  viewState: Pick<SheetViewState, "active" | "updateActive">;
}

export function useSheetSelection(deps: SheetSelectionDeps) {
  /** The current selection, validated against the sheet (a file's or a
   * stored selection is clamped, snapped and made consistent). */
  const selection = computed((): SheetSelection => {
    void deps.sheetMetrics.metrics.value;
    void deps.sheetMetrics.merges.value;
    return normaliseSelection(deps.viewState.active.value.selection, deps.sheetMetrics.nav());
  });

  /** A pointer drag in progress: what it extends and from where — a header
   * drag begun with Shift (`extend`) extends from the active cell. */
  const drag = shallowRef<{ axis: "cell" | "row" | "column"; from: number; base?: SheetSelection; extend?: true } | null>(null);

  function set(next: SheetSelection): void {
    deps.viewState.updateActive({ selection: next });
  }

  /** A press on a cell. */
  function press(at: CellRef, mods: PressModifiers): void {
    const nav = deps.sheetMetrics.nav();
    const current = selection.value;
    set(mods.shift ? extendTo(current, at, nav) : mods.mod ? addRange(current, at, nav) : selectCell(at, nav));
    drag.value = { axis: "cell", from: 0 };
  }

  /** A press on a row or column header. */
  function pressHeader(axis: "row" | "column", index: number, mods: PressModifiers): void {
    const nav = deps.sheetMetrics.nav();
    const current = selection.value;
    if (mods.shift) {
      set(extendToLine(current, axis, index, nav));
      drag.value = { axis, from: index, extend: true };
    } else if (mods.mod) {
      set(headerSelection(axis, index, index, nav, current));
      drag.value = { axis, from: index, base: current };
    } else {
      set(headerSelection(axis, index, index, nav));
      drag.value = { axis, from: index };
    }
  }

  /** The pointer moved to `at` during a drag. */
  function dragTo(at: CellRef): void {
    const current = drag.value;
    if (!current) {
      return;
    }
    const nav = deps.sheetMetrics.nav();
    if (current.axis === "cell") {
      set(extendTo(selection.value, at, nav));
    } else if (current.extend) {
      set(extendToLine(selection.value, current.axis, current.axis === "row" ? at.row : at.col, nav));
    } else {
      set(headerSelection(current.axis, current.from, current.axis === "row" ? at.row : at.col, nav, current.base));
    }
  }

  function release(): void {
    drag.value = null;
  }

  /** The select-all corner. */
  function selectEverything(): void {
    set(selectAll());
  }

  /** A key on the grid: applies the keyboard map; true when consumed (the
   * caller then prevents the default and scrolls the active cell into view). */
  function handleKey(event: KeyboardEvent, page: { rows: number; cols: number }): boolean {
    if (isOutsideDocumentInput(event.target)) {
      return false;
    }
    const next = navigateByKey(
      selection.value,
      { key: event.key, shift: event.shiftKey, mod: primaryModifier(event), alt: event.altKey, ctrl: event.ctrlKey },
      deps.sheetMetrics.nav(page),
    );
    if (!next) {
      return false;
    }
    set(next);
    return true;
  }

  /** "<rows>R × <cols>C" of the active range while a press is held. */
  const dragSize = computed((): string | null => {
    if (!drag.value) {
      return null;
    }
    const sel = selection.value;
    const range = sel.ranges[sel.activeRangeIndex] ?? sel.ranges[0];
    return range ? `${range.bottom - range.top + 1}R × ${range.right - range.left + 1}C` : null;
  });

  /** The Name Box: the active cell's address (a merge's top-left), or the
   * drag size while a drag is in progress. */
  const nameBoxText = computed(() => dragSize.value ?? formatA1(selection.value.active));

  /** The formula bar's content: the active cell's General text (SPEC-118
   * refines this to the edit text), line breaks preserved. */
  const activeContent = computed(() => {
    const sheet = deps.sheet();
    const active = selection.value.active;
    return sheet ? generalText(getCell(sheet, active.row, active.col)?.value ?? null) : "";
  });

  const aggregates = shallowRef<SelectionAggregates | null>(null);
  let generation = 0;
  watch(
    [() => deps.sheet(), selection],
    ([sheet, sel]) => {
      generation += 1;
      if (!sheet || !selectionSpansMultipleCells(sel.ranges, deps.sheetMetrics.merges.value)) {
        aggregates.value = null;
        return;
      }
      const nav = deps.sheetMetrics.nav();
      if (selectionStoredCellEstimate(sheet, sel.ranges, AGGREGATE_DEFER_THRESHOLD) <= AGGREGATE_DEFER_THRESHOLD) {
        aggregates.value = selectionAggregates(sheet, sel.ranges, nav);
        return;
      }
      aggregates.value = null;
      const mine = generation;
      runWhenIdle(() => {
        if (mine === generation) {
          aggregates.value = selectionAggregates(sheet, sel.ranges, nav);
        }
      }, 1000);
    },
    { immediate: true },
  );

  /** The status bar's texts ({} when nothing is shown). */
  const statusTexts = computed(() =>
    aggregates.value
      ? formatAggregates(aggregates.value, selectionSpansMultipleCells(selection.value.ranges, deps.sheetMetrics.merges.value))
      : {},
  );

  return {
    selection,
    drag,
    press,
    pressHeader,
    dragTo,
    release,
    selectEverything,
    handleKey,
    set,
    dragSize,
    nameBoxText,
    activeContent,
    aggregates,
    statusTexts,
  };
}

export type SheetSelectionController = ReturnType<typeof useSheetSelection>;
