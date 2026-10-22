// Drag-move in the workbook view (SPEC-125 "Cut and move"): the selection
// dragged by its outline to another place on its sheet.
//
// - Within 4px of a single-range selection's outline (not the fill handle)
//   the grid shows `cursor: move` (`hover`, set by the grid's pointer
//   handler, which also calls begin/dragTo/drop).
// - Dragging outlines the block at its snapped destination
//   (`sheet-move-preview`) with a tooltip naming the destination's address
//   ("H10:H12"); the block stays inside the grid.
// - Dropping moves the cells — moveRange's cut semantics, ONE "Move cells"
//   step — or, with Alt/Option (or Ctrl on Windows/Linux) held at the drop,
//   pastes a copy of them there (copy-paste semantics, one "Paste" step).
//   Dropping onto non-empty cells outside the source first asks "There's
//   already data here. Do you want to replace it?" (SheetMoveOverwriteDialog,
//   `sheet-move-overwrite`, Replace / Cancel); a drop partly over a merge
//   (or any other refusal) is refused like a paste, with a toast.
// - Escape, or a drop outside the grid, cancels. Moving across sheets is
//   out of scope.
import { computed, onScopeDispose, ref, shallowRef, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { formatRange, type CellRef, type GridRange } from "../../lib/spreadsheet/address";
import type { CellTextMeasurer } from "../../lib/spreadsheet/autofit";
import { buildClipboardCopy } from "../../lib/spreadsheet/clipboard";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { moveDestination, moveRange, moveRefusal, snappedMove, type RangeMove } from "../../lib/spreadsheet/move";
import { blockFromPayload, pasteBlock, pasteRefusal, planPaste, type PasteResult } from "../../lib/spreadsheet/paste";
import { contains } from "../../lib/spreadsheet/sheetEdit";
import { createCanvasCellMeasurer } from "./cellMeasurer";
import { runStep, selectionOf, type StepContext } from "./sheetSteps";
import type { CellEditing } from "./useCellEditing";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export interface RangeMoveDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeSheet: Readonly<Ref<Sheet | null>> };
  viewState: Pick<SheetViewState, "syncInto">;
  selection: Pick<SheetSelectionController, "selection" | "set" | "formatContext">;
  editing: { undo: Pick<SpreadsheetUndo, "record" | "commit">; cell: Pick<CellEditing, "editing"> };
  /** Shows a toast — default: useToasts' addToast. */
  toast?: (message: string) => void;
  /** Row refit's measurer — default: the canvas one. */
  measurer?: CellTextMeasurer;
}

/** Whether a destination holds a value or a formula outside the source. */
function overwrites(sheet: Sheet, from: GridRange, dest: GridRange): boolean {
  return sheet.rows.some(
    (row) =>
      row.index >= dest.top &&
      row.index <= dest.bottom &&
      row.cells.some(
        (cell) =>
          cell.col >= dest.left &&
          cell.col <= dest.right &&
          (cell.value !== null || cell.formula !== undefined) &&
          !contains(from, { top: row.index, left: cell.col, bottom: row.index, right: cell.col }),
      ),
  );
}

export function useRangeMove(deps: RangeMoveDeps) {
  const toast = deps.toast ?? useToasts().addToast;
  let measurer = deps.measurer;
  /** The pointer is over the selection's outline: `cursor: move`. */
  const hover = ref(false);
  const drag = shallowRef<{ sheetId: string; from: GridRange; anchor: CellRef; toTop: number; toLeft: number } | null>(null);
  /** The pending overwrite question: what Replace does (null: none). */
  const overwrite = shallowRef<{ proceed: () => void } | null>(null);

  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const gridSheet = (): Sheet | null => {
    const sheet = deps.book.activeSheet.value;
    return sheet && sheet.kind === "grid" ? sheet : null;
  };

  function context(m: SpreadsheetModel): StepContext {
    measurer ??= createCanvasCellMeasurer();
    return { model: m, undo: deps.editing.undo, measurer, format: deps.selection.formatContext.value, sync: (each) => deps.viewState.syncInto(each) };
  }

  /** A press on the outline starts a drag; false when no drag can start. */
  function begin(at: CellRef): boolean {
    const sheet = gridSheet();
    const ranges = deps.selection.selection.value.ranges;
    if (!sheet || !model() || deps.editing.cell.editing.value || ranges.length !== 1) {
      return false;
    }
    const from = ranges[0]!;
    drag.value = { sheetId: sheet.id, from, anchor: at, toTop: from.top, toLeft: from.left };
    return true;
  }

  function dragTo(at: CellRef): void {
    const current = drag.value;
    if (current) {
      drag.value = { ...current, ...snappedMove(current.from, current.anchor, at) };
    }
  }

  /** The snapped destination while it differs from the source. */
  const preview = computed(() => {
    const current = drag.value;
    if (!current || (current.toTop === current.from.top && current.toLeft === current.from.left)) {
      return null;
    }
    const range = moveDestination({ fromSheetId: current.sheetId, from: current.from, toSheetId: current.sheetId, toTop: current.toTop, toLeft: current.toLeft });
    return { range, address: formatRange(range) };
  });

  /** Carries the drop out as one step (see the module header). */
  function perform(move: RangeMove, copy: boolean): void {
    const m = model();
    if (!m) {
      return;
    }
    const dest = moveDestination(move);
    const refit = () => [
      { sheetId: move.toSheetId, ranges: [dest] },
      { sheetId: move.fromSheetId, ranges: [move.from] },
    ];
    let result: { ok: true } | { ok: false; message: string } | PasteResult;
    if (copy) {
      const source = buildClipboardCopy(m, move.fromSheetId, move.from);
      const target = { ranges: [{ top: dest.top, left: dest.left, bottom: dest.top, right: dest.left }], active: { row: dest.top, col: dest.left } };
      result = runStep(context(m), "Paste", { sheetId: move.toSheetId }, () => pasteBlock(m, move.toSheetId, target, blockFromPayload(source.payload), { what: "all" }, { styleIds: source.styleIds }), refit);
    } else {
      result = runStep(context(m), "Move cells", { sheetId: move.toSheetId }, () => moveRange(m, move), refit);
    }
    if (!result.ok) {
      toast(result.message);
      return;
    }
    deps.selection.set(selectionOf(dest));
  }

  /** The drop: `copy` when Alt/Option (or Ctrl on Windows/Linux) is held;
   * `outside` when the pointer left the grid (a cancel). */
  function drop(options: { copy: boolean; outside: boolean }): void {
    const current = drag.value;
    drag.value = null;
    const m = model();
    const sheet = gridSheet();
    if (!current || !m || !sheet || options.outside || sheet.id !== current.sheetId || (current.toTop === current.from.top && current.toLeft === current.from.left)) {
      return;
    }
    const move: RangeMove = { fromSheetId: sheet.id, from: current.from, toSheetId: sheet.id, toTop: current.toTop, toLeft: current.toLeft };
    const dest = moveDestination(move);
    let refusal: string | null;
    if (options.copy) {
      const plan = planPaste(sheet, { ranges: [dest], active: { row: dest.top, col: dest.left } }, blockFromPayload(buildClipboardCopy(m, sheet.id, current.from).payload));
      refusal = typeof plan === "string" ? plan : pasteRefusal(m, m.sheets.indexOf(sheet), plan);
    } else {
      refusal = moveRefusal(m, move);
    }
    if (refusal !== null) {
      toast(refusal);
      return;
    }
    if (overwrites(sheet, current.from, dest)) {
      overwrite.value = { proceed: () => perform(move, options.copy) };
    } else {
      perform(move, options.copy);
    }
  }

  function onKeydown(event: KeyboardEvent): void {
    if (event.key === "Escape" && drag.value) {
      drag.value = null;
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
    hover,
    setHover: (on: boolean): void => {
      hover.value = on;
    },
    drag,
    preview,
    begin,
    dragTo,
    drop,
    cancel: (): void => {
      drag.value = null;
    },
    overwrite,
    confirmOverwrite(): void {
      const pending = overwrite.value;
      overwrite.value = null;
      pending?.proceed();
    },
    cancelOverwrite(): void {
      overwrite.value = null;
    },
    onKeydown,
  };
}

export type RangeMoveController = ReturnType<typeof useRangeMove>;

/** Narrows a component's runtime-typed `move` prop (the .vue parser bars
 * casts). */
export function asRangeMove(value: unknown): RangeMoveController {
  return value as RangeMoveController;
}
