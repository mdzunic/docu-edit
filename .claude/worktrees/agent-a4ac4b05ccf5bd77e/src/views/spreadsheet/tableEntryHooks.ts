// SPEC-133's typed-entry hooks (cellCommit.ts's EntryHooks; useSheetTables
// installs them on the cell editor) — they run for a typed entry's commit
// only, never a Fill, paste, fill handle, undo or dropdown pick:
// - `prepare` qualifies a formula typed into (or growing) a table, so
//   `=[@Price]*2` is stored `Sales[@Price]*2` (qualifyEntryFormula);
// - `within` renames a header cell's column INSIDE the entry's own step, so
//   the header text and the column name never disagree on any undo state;
// - `after` records "Expand table" when the entry grows a table (never a
//   locked one), then "Fill table column" when its formula makes a
//   calculated column — each its own step after the entry's (a step that
//   changed nothing is dropped), so one undo removes only the last.
// A cell neither in nor right below or right of a table skips them all
// (and so never settles a running recalculation pass early).
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import * as commands from "../../lib/spreadsheet/tableCommands";
import type { EntryHooks } from "./cellCommit";
import { TABLE_UNDO_LABELS } from "./tableDesign";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export interface TableEntryHooksDeps {
  model: () => SpreadsheetModel | null;
  undo: Pick<SpreadsheetUndo, "record" | "commit">;
  locale: () => SpreadsheetLocale;
}

/** Whether a cell is in, or right below or right of, one of the sheet's tables. */
const nearTable = (m: SpreadsheetModel, sheetId: string, row: number, col: number): boolean =>
  m.sheets.find((sheet) => sheet.id === sheetId)?.tables?.some(({ ref: r }) => row >= r.top && row <= r.bottom + 1 && col >= r.left && col <= r.right + 1) === true;

export function tableEntryHooks(deps: TableEntryHooksDeps): EntryHooks {
  const { model, locale } = deps;
  /** A follow-up step of the entry's: kept only when `command` changed something. */
  function followUp(label: string, sheetId: string, command: (m: SpreadsheetModel) => boolean): void {
    const m = model();
    if (m) {
      deps.undo.record(label, { sheetId });
      command(m);
      deps.undo.commit();
    }
  }

  return {
    prepare: ({ model: m, sheetId, at, entry }) =>
      entry.formula === undefined || !nearTable(m, sheetId, at.row, at.col) ? entry : { ...entry, formula: commands.qualifyEntryFormula(m, sheetId, at.row, at.col, entry.formula) },
    within: ({ model: m, sheetId, at }) => {
      const header = nearTable(m, sheetId, at.row, at.col) ? commands.headerEditColumn(m, sheetId, at.row, at.col, locale()) : undefined;
      if (header) {
        commands.renameTableColumn(m, sheetId, header.tableId, header.columnIndex, header.text);
      }
    },
    after: ({ model: m, sheetId, at, entry }) => {
      if ((entry.value === null && entry.formula === undefined) || !nearTable(m, sheetId, at.row, at.col)) {
        return;
      }
      followUp(TABLE_UNDO_LABELS.expand, sheetId, (current) => commands.expandTableForEntry(current, sheetId, at.row, at.col, locale()));
      if (entry.formula !== undefined) {
        followUp(TABLE_UNDO_LABELS.fill, sheetId, (current) => commands.fillCalculatedColumn(current, sheetId, at.row, at.col));
      }
    },
  };
}
