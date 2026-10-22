// How SPEC-134's replace plans reach the model (find.ts's planReplace says
// what to write; this enters it): ONE undo step — one record(label, scope)
// before any edit and one commit() after, so the edit pipeline recalculates
// once and marks dirty only the sheets whose rows changed — with each edit
// entered through SPEC-122's cell entry. A Formulas-mode edit is entered
// exactly as typed (cellEntry.ts's entryFromText: a leading "=" makes a
// canonical formula); a Values-mode edit goes through parseInput in the
// display locale and the workbook's date system, a formula result discarded
// so the text is stored as a string — never a formula. Each sheet's edits go
// in one copy-on-write batch (cells.ts's writeEntries: the cell's style kept,
// a proposed number format applied only over General) and its touched rows
// are refitted, as a typed entry refits them.
import { refitRows, type CellTextMeasurer } from "../../lib/spreadsheet/autofit";
import { entryFromInput, writeEntries, type CellEntry } from "../../lib/spreadsheet/cells";
import { sanitizeXmlText } from "../../lib/core/xmlText";
import type { ReplacePlan } from "../../lib/spreadsheet/find";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import { formatCodeOf } from "../../lib/spreadsheet/numberFormat/builtin";
import { parseInput } from "../../lib/spreadsheet/numberFormat/input";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import type { FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import { cellStyle, effectiveStyleId } from "../../lib/spreadsheet/styles";
import { entryFromText } from "./cellEntry";
import type { SpreadsheetUndo, UndoScope } from "./useSpreadsheetUndo";

export interface ReplaceWriteContext {
  model: SpreadsheetModel;
  undo: Pick<SpreadsheetUndo, "record" | "commit">;
  locale: SpreadsheetLocale;
  measurer: CellTextMeasurer;
  format: FormatContext;
}

/** One edit's entry (see the header); null when cell entry refuses it. */
function entryOf(edit: ReplacePlan["edits"][number], ctx: ReplaceWriteContext, sheetIndex: number): CellEntry | null {
  const sheet = ctx.model.sheets[sheetIndex]!;
  const style = cellStyle(ctx.model, effectiveStyleId(sheet, edit.row, edit.col));
  if (!edit.asValue) {
    const outcome = entryFromText(edit.text, { locale: ctx.locale, dateSystem: ctx.model.dateSystem, style });
    return "entry" in outcome ? outcome.entry : null;
  }
  const text = sanitizeXmlText(edit.text);
  const parsed = parseInput(text, ctx.locale, ctx.model.dateSystem, { existingFormat: formatCodeOf(style, ctx.locale) });
  return parsed.formula === undefined ? entryFromInput(parsed) : { value: text };
}

/** Enters `plan`'s edits as one step labelled `label`; the number of cells
 * written (an edit cell entry refuses is left out). */
export function applyReplacePlan(ctx: ReplaceWriteContext, plan: ReplacePlan, label: string, scope: UndoScope): number {
  const bySheet = new Map<string, ReplacePlan["edits"]>();
  for (const edit of plan.edits) {
    bySheet.set(edit.sheetId, [...(bySheet.get(edit.sheetId) ?? []), edit]);
  }
  ctx.undo.record(label, scope);
  let written = 0;
  for (const [sheetId, edits] of bySheet) {
    const sheetIndex = ctx.model.sheets.findIndex((sheet) => sheet.id === sheetId);
    if (sheetIndex < 0) {
      continue;
    }
    const entries = edits.flatMap((edit) => {
      const entry = entryOf(edit, ctx, sheetIndex);
      return entry ? [{ row: edit.row, col: edit.col, entry }] : [];
    });
    written += entries.length;
    const before = ctx.model.sheets[sheetIndex]!.rows;
    writeEntries(ctx.model, sheetIndex, entries);
    refitRows(ctx.model, sheetId, entries.map(({ row, col }) => ({ top: row, left: col, bottom: row, right: col })), ctx.measurer, ctx.format, before);
  }
  ctx.undo.commit();
  return written;
}
