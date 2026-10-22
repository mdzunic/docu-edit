// One SPEC-125 command as ONE undo step (SPEC-121's record/commit): the view
// state synced into the model first (so the step holds the selection it was
// made from), the step recorded, the command run, SPEC-122's row refit
// (refitRows, which SPEC-124 extends to pastes and fills) over the rows the
// command touched on each sheet it changed — given that sheet's rows from
// before, so an untouched row keeps the height its file gave it — and the
// step committed through the edit pipeline, which marks the touched sheets
// dirty and hands the changed cells to recalculation. A refused command
// changes nothing, and commit() then records nothing. Shared by
// useSheetClipboard, useFillHandle and useRangeMove.
import type { GridRange } from "../../lib/spreadsheet/address";
import { refitRows, type CellTextMeasurer } from "../../lib/spreadsheet/autofit";
import type { SheetSelection } from "../../lib/spreadsheet/selection";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import type { SpreadsheetUndo, UndoScope } from "./useSpreadsheetUndo";

export interface StepContext {
  model: SpreadsheetModel;
  undo: Pick<SpreadsheetUndo, "record" | "commit">;
  measurer: CellTextMeasurer;
  format: FormatContext;
  /** Writes the view state into the model (useSheetViewState's syncInto). */
  sync: (model: SpreadsheetModel) => void;
}

/** Where a command's rows refit: per sheet, the ranges it wrote. */
export type Refits = { sheetId: string; ranges: GridRange[] }[];

/** Runs `run` as one undo step (see the module header). */
export function runStep<T extends { ok: boolean }>(ctx: StepContext, label: string, scope: UndoScope, run: () => T, refits: (result: T) => Refits): T {
  ctx.sync(ctx.model);
  ctx.undo.record(label, scope);
  const before = new Map(ctx.model.sheets.map((sheet) => [sheet.id, sheet.rows] as const));
  const result = run();
  if (result.ok) {
    for (const { sheetId, ranges } of refits(result)) {
      refitRows(ctx.model, sheetId, ranges, ctx.measurer, ctx.format, before.get(sheetId));
    }
  }
  ctx.undo.commit();
  return result;
}

/** A selection of one range, its top-left active. */
export function selectionOf(range: GridRange): SheetSelection {
  return { ranges: [range], active: { row: range.top, col: range.left }, activeRangeIndex: 0 };
}
