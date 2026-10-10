// The fill handle, its options chip, double-click fill and fill down/right in
// the workbook view (SPEC-125 "Series detection", "Double-click fill",
// "Shortcuts") over lib/spreadsheet/fill.ts.
//
// - The handle (SheetFillHandle, `sheet-fill-handle`) sits on the bottom-right
//   corner of a single-range selection while no cell is being edited. A
//   drag (the grid's pointer handler calls begin/dragTo/release) extends the
//   selection along the axis the pointer moved further on — or shrinks it
//   back inside — and the target is outlined (`sheet-fill-preview`). Release
//   runs fillRange as ONE "Fill" step: Ctrl (Windows/Linux) or Alt/Option
//   (every platform) held at release toggles copy ↔ series; a shrink clears
//   the contents of the cells left behind. Escape cancels a drag.
// - After a fill the chip (FillOptionsChip, `sheet-fill-options`) offers
//   Copy cells, Fill series, Fill formatting only and Fill without
//   formatting, the applied one pressed; choosing one re-applies the fill as
//   the SAME undo step (the fill is undone and made again under the same
//   label). Escape, any key or press elsewhere, or another sheet dismisses it.
// - Double-clicking the handle fills down to doubleClickFillTarget's extent,
//   and does nothing without adjacent data.
// - Ctrl/Cmd+D and Ctrl/Cmd+R (SHEET_SHORTCUTS sheet.fillDown /
//   sheet.fillRight, claimed only while the grid has focus) and Edit ▸ Fill ▸
//   Down/Right: one "Fill down" / "Fill right" step.
// Every step refits the rows it wrote (sheetSteps.ts); a refusal toasts.
import { computed, onScopeDispose, shallowRef, watch, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { isMacPlatform } from "../../lib/core/platform";
import { findShortcut } from "../../lib/shortcuts";
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import type { CellTextMeasurer } from "../../lib/spreadsheet/autofit";
import { doubleClickFillTarget, fillDownRight, fillRange, type AppliedFill, type FillMode, type FillResult } from "../../lib/spreadsheet/fill";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { createCanvasCellMeasurer } from "./cellMeasurer";
import { inGrid } from "./headerMenu";
import { SHEET_SHORTCUTS } from "./sheetShortcuts";
import { runStep, selectionOf, type StepContext } from "./sheetSteps";
import type { CellEditing } from "./useCellEditing";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export interface FillHandleDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeSheet: Readonly<Ref<Sheet | null>> };
  viewState: Pick<SheetViewState, "syncInto">;
  selection: Pick<SheetSelectionController, "selection" | "set" | "formatContext">;
  editing: { undo: Pick<SpreadsheetUndo, "record" | "commit" | "undo">; cell: Pick<CellEditing, "editing"> };
  locale: () => SpreadsheetLocale;
  /** Shows a toast — default: useToasts' addToast. */
  toast?: (message: string) => void;
  /** Row refit's measurer — default: the canvas one. */
  measurer?: CellTextMeasurer;
}

/** The chip's options, in order. */
export const FILL_CHIP_MODES: readonly AppliedFill[] = ["copy", "series", "formats", "no-formats"];

const sameRange = (a: GridRange, b: GridRange): boolean => a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;

/** The range a fill drag from `seed` reaches with the pointer over `at`
 * (see the module header). */
export function fillDragTarget(seed: GridRange, at: CellRef): GridRange {
  const below = at.row > seed.bottom ? at.row - seed.bottom : at.row < seed.top ? seed.top - at.row : 0;
  const beside = at.col > seed.right ? at.col - seed.right : at.col < seed.left ? seed.left - at.col : 0;
  if (below === 0 && beside === 0) {
    // Back inside: shrink along the axis moved further from the corner.
    return seed.bottom - at.row >= seed.right - at.col ? { ...seed, bottom: Math.max(seed.top, at.row) } : { ...seed, right: Math.max(seed.left, at.col) };
  }
  if (below >= beside) {
    return at.row > seed.bottom ? { ...seed, bottom: at.row } : { ...seed, top: at.row };
  }
  return at.col > seed.right ? { ...seed, right: at.col } : { ...seed, left: at.col };
}

export function useFillHandle(deps: FillHandleDeps) {
  const toast = deps.toast ?? useToasts().addToast;
  let measurer = deps.measurer;
  const drag = shallowRef<{ sheetId: string; seed: GridRange; target: GridRange } | null>(null);
  const chip = shallowRef<{ sheetId: string; seed: GridRange; target: GridRange; applied: AppliedFill } | null>(null);

  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const gridSheet = (): Sheet | null => {
    const sheet = deps.book.activeSheet.value;
    return sheet && sheet.kind === "grid" ? sheet : null;
  };

  /** The range the handle sits on: a single range of a grid sheet, Ready. */
  const seed = computed((): GridRange | null => {
    const ranges = deps.selection.selection.value.ranges;
    return gridSheet() && model() && !deps.editing.cell.editing.value && ranges.length === 1 ? ranges[0]! : null;
  });

  function context(m: SpreadsheetModel): StepContext {
    measurer ??= createCanvasCellMeasurer();
    return { model: m, undo: deps.editing.undo, measurer, format: deps.selection.formatContext.value, sync: (each) => deps.viewState.syncInto(each) };
  }

  /** One fill as one step; the result (null when refused or impossible). */
  function fill(label: string, from: GridRange, to: GridRange, mode: FillMode): FillResult | null {
    const m = model();
    const sheet = gridSheet();
    if (!m || !sheet) {
      return null;
    }
    const result = runStep(context(m), label, { sheetId: sheet.id }, () => fillRange(m, sheet.id, from, to, mode, deps.locale()), (done) =>
      done.ok ? [{ sheetId: sheet.id, ranges: [done.range] }] : [],
    );
    if (!result.ok) {
      toast(result.message);
      return null;
    }
    deps.selection.set(selectionOf(result.cleared ? to : result.range));
    return result;
  }

  /** Runs a fill from the handle and shows the chip for it. */
  function fillWithChip(from: GridRange, to: GridRange, mode: FillMode): void {
    const sheetId = gridSheet()?.id;
    const result = fill("Fill", from, to, mode);
    if (result?.ok && !result.cleared && sheetId) {
      chip.value = { sheetId, seed: from, target: result.range, applied: result.applied };
    }
  }

  function begin(): boolean {
    const sheet = gridSheet();
    const from = seed.value;
    if (!sheet || !from) {
      return false;
    }
    chip.value = null;
    drag.value = { sheetId: sheet.id, seed: from, target: from };
    return true;
  }

  function dragTo(at: CellRef): void {
    const current = drag.value;
    if (current) {
      drag.value = { ...current, target: fillDragTarget(current.seed, at) };
    }
  }

  /** The drag ends: the fill, Ctrl (Windows/Linux) or Alt/Option toggling. */
  function release(mods: { ctrl: boolean; alt: boolean }): void {
    const current = drag.value;
    drag.value = null;
    if (!current || sameRange(current.seed, current.target) || gridSheet()?.id !== current.sheetId) {
      return;
    }
    const toggled = mods.alt || (mods.ctrl && !isMacPlatform());
    fillWithChip(current.seed, current.target, toggled ? "toggle" : "auto");
  }

  function doubleClick(): void {
    const sheet = gridSheet();
    const from = seed.value;
    const to = sheet && from ? doubleClickFillTarget(sheet, from) : null;
    if (from && to) {
      fillWithChip(from, to, "auto");
    }
  }

  /** A chip option: the fill made again, as the same undo step. */
  function choose(mode: AppliedFill): void {
    const current = chip.value;
    if (!current || current.applied === mode || gridSheet()?.id !== current.sheetId) {
      return;
    }
    deps.editing.undo.undo();
    const result = fill("Fill", current.seed, current.target, mode);
    chip.value = result?.ok ? { ...current, applied: mode } : null;
  }

  function fillDirection(direction: "down" | "right"): void {
    const m = model();
    const sheet = gridSheet();
    if (!m || !sheet || deps.editing.cell.editing.value) {
      return;
    }
    const ranges = deps.selection.selection.value.ranges;
    const result = runStep(context(m), direction === "down" ? "Fill down" : "Fill right", { sheetId: sheet.id }, () => fillDownRight(m, sheet.id, ranges, direction), () => [
      { sheetId: sheet.id, ranges: [...ranges] },
    ]);
    if (!result.ok) {
      toast(result.message);
    }
  }

  function dismiss(): void {
    chip.value = null;
  }

  function onKeydown(event: KeyboardEvent): void {
    if (event.key === "Escape" && (drag.value || chip.value)) {
      drag.value = null;
      chip.value = null;
      return;
    }
    if (chip.value && !["Shift", "Control", "Alt", "Meta"].includes(event.key)) {
      chip.value = null;
    }
    if (!inGrid(event.target) || deps.editing.cell.editing.value || !gridSheet()) {
      return;
    }
    const id = findShortcut(event, isMacPlatform(), SHEET_SHORTCUTS)?.id;
    if (id === "sheet.fillDown" || id === "sheet.fillRight") {
      event.preventDefault();
      fillDirection(id === "sheet.fillDown" ? "down" : "right");
    }
  }

  /** A press anywhere but the chip dismisses it ("any other command"). */
  function onPointerDown(event: PointerEvent): void {
    if (chip.value && !(event.target instanceof Element && event.target.closest('[data-testid="sheet-fill-options"]'))) {
      chip.value = null;
    }
  }

  watch(
    () => deps.book.activeSheet.value?.id,
    () => {
      drag.value = null;
      chip.value = null;
    },
  );

  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown, true);
    window.addEventListener("pointerdown", onPointerDown, true);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown, true);
      window.removeEventListener("pointerdown", onPointerDown, true);
    }
  });

  return {
    /** The range the handle sits on (null: no handle). */
    seed,
    drag,
    /** The drag's target while it differs from the seed (`sheet-fill-preview`). */
    preview: computed(() => (drag.value && !sameRange(drag.value.seed, drag.value.target) ? drag.value.target : null)),
    chip,
    begin,
    dragTo,
    release,
    cancel: (): void => {
      drag.value = null;
    },
    doubleClick,
    choose,
    dismiss,
    fillDown: (): void => fillDirection("down"),
    fillRight: (): void => fillDirection("right"),
    onKeydown,
  };
}

export type FillHandle = ReturnType<typeof useFillHandle>;

/** Narrows a component's runtime-typed `fill` prop (the .vue parser bars
 * casts). */
export function asFillHandle(value: unknown): FillHandle {
  return value as FillHandle;
}
