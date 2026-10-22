// Selecting ranges on a sheet and bringing them into view — a followed
// place link's jump (SPEC-135), as the Name Box's (useNames) does it: the
// sheet activated when it is not on show, the top-left cell selected and
// revealed first (the grid reveals on `revealTick`), then the ranges
// selected and the grid focused.
import { nextTick, type Ref, type ShallowRef } from "vue";
import type { GridRange } from "../../lib/spreadsheet/address";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { CellEditing } from "./useCellEditing";
import type { SheetViewState } from "./useSheetViewState";

export interface SheetJumpDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeIndex: Ref<number>; selectSheet(index: number): void };
  viewState: Pick<SheetViewState, "update">;
  cell: Pick<CellEditing, "revealTick" | "focusCell">;
}

/** Selects `ranges` on `sheetId`, revealed and focused (nothing for an
 * unknown sheet or no range). */
export async function jumpToRanges(deps: SheetJumpDeps, sheetId: string, ranges: readonly GridRange[]): Promise<void> {
  const index = deps.book.workbook.value?.sheets.findIndex((sheet) => sheet.id === sheetId) ?? -1;
  const first = ranges[0];
  if (index < 0 || !first) {
    return;
  }
  if (index !== deps.book.activeIndex.value) {
    deps.book.selectSheet(index);
  }
  const active = { row: first.top, col: first.left };
  deps.viewState.update(sheetId, { selection: { active, ranges: [{ top: active.row, left: active.col, bottom: active.row, right: active.col }], activeRangeIndex: 0 } });
  await nextTick();
  deps.cell.revealTick.value += 1;
  await nextTick();
  deps.viewState.update(sheetId, { selection: { active, ranges: ranges.map((range) => ({ ...range })), activeRangeIndex: 0 } });
  deps.cell.focusCell();
}
