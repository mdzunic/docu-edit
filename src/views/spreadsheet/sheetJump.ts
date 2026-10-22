// The jump the Name Box (SPEC-133) and Go To (SPEC-134) share: `ranges`
// selected on `sheetId` with the active cell at the first range's top-left —
// the active sheet switched first when it differs (view state, never
// `dirty`), the top-left revealed, then the ranges selected and focus handed
// back to the grid. useSheetGoTo wraps it so every jump first remembers where
// it came from (useNames' setJump). useSheetFind's navigation selects a hit
// through it too, with `active` (the hit, kept inside a searched selection's
// ranges) and `focus: false` (focus stays in the find bar). A jump made
// while an earlier one is still between its ticks wins: the earlier stops.
import { nextTick, type Ref, type ShallowRef } from "vue";
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { CellEditing } from "./useCellEditing";
import type { SheetViewState } from "./useSheetViewState";

export interface SheetJumpDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeIndex: Ref<number>; selectSheet(index: number): void };
  viewState: Pick<SheetViewState, "update">;
  cell: Pick<CellEditing, "focusCell" | "revealTick">;
}

export interface JumpOptions {
  /** The active cell — default the first range's top-left. */
  active?: CellRef;
  /** Hand focus back to the grid — default true. */
  focus?: boolean;
}

/** Selects `ranges` on `sheetId` (see the header); nothing for an unknown
 * sheet or no range. */
export type SheetJump = (sheetId: string, ranges: GridRange[], options?: JumpOptions) => Promise<void>;

const holds = (range: GridRange, at: CellRef): boolean => at.row >= range.top && at.row <= range.bottom && at.col >= range.left && at.col <= range.right;

export function createSheetJump(deps: SheetJumpDeps): SheetJump {
  /** The latest jump: an earlier one still between its ticks stops there. */
  let latest = 0;
  return async (sheetId, ranges, options = {}) => {
    const mine = ++latest;
    const index = deps.book.workbook.value?.sheets.findIndex((sheet) => sheet.id === sheetId) ?? -1;
    const first = ranges[0];
    if (index < 0 || !first) {
      return;
    }
    if (index !== deps.book.activeIndex.value) {
      deps.book.selectSheet(index);
    }
    const active = options.active ?? { row: first.top, col: first.left };
    const selection = { active, ranges: ranges.map((range) => ({ ...range })), activeRangeIndex: Math.max(0, ranges.findIndex((range) => holds(range, active))) };
    // A move of the active cell (find) selects its ranges at once — never
    // passing through a single cell, which would read as a scope change.
    deps.viewState.update(sheetId, { selection: options.active ? selection : { active, ranges: [{ top: active.row, left: active.col, bottom: active.row, right: active.col }], activeRangeIndex: 0 } });
    await nextTick();
    if (mine !== latest) {
      return;
    }
    deps.cell.revealTick.value += 1;
    await nextTick();
    if (mine !== latest) {
      return;
    }
    deps.viewState.update(sheetId, { selection });
    if (options.focus !== false) {
      deps.cell.focusCell();
    }
  };
}
