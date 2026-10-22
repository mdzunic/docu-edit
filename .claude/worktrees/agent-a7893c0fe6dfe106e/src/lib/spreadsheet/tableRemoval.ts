// Removing tables (SPEC-133 "Tables follow cells"): a table deleted by a
// structural edit, a move's destination, a cleared range or a deleted sheet
// leaves `Sheet.tables`, and every reference to it — in cell formulas, names
// and other tables' calculated and totals formulas — becomes #REF!. Kept
// apart from nameTableAdjusters.ts so cells.ts can use it without importing
// the adjuster registries (an import cycle through structure.ts).
import type { SheetTable, SpreadsheetModel } from "./model";
import { formulaMentions, rewriteWorkbookFormulas } from "./names";
import { removeTableReferences } from "./tables";

/** References to the tables named `tables` — and to the columns named in
 * `columns` of surviving tables — become #REF! in every cell formula, name,
 * calculated and totals formula of the workbook. */
export function removeTableReferencesEverywhere(model: SpreadsheetModel, tables: readonly string[], columns: readonly { table: string; column: string }[] = []): void {
  if (tables.length === 0 && columns.length === 0) {
    return;
  }
  rewriteWorkbookFormulas(model, (text) => {
    const mentions = formulaMentions(text);
    let out = text;
    for (const table of tables) {
      const key = table.toUpperCase();
      if (mentions.tables.has(key) || mentions.names.has(key)) {
        out = removeTableReferences(out, table);
      }
    }
    for (const { table, column } of columns) {
      if (mentions.tables.has(table.toUpperCase())) {
        out = removeTableReferences(out, table, column);
      }
    }
    return out;
  });
}

/** Removes `victims` (tables of the sheet `sheetId`); every reference to
 * them becomes #REF! (removeTableReferencesEverywhere). */
export function removeTables(model: SpreadsheetModel, sheetId: string, victims: readonly SheetTable[]): void {
  if (victims.length === 0) {
    return;
  }
  const gone = new Set(victims);
  model.sheets = model.sheets.map((sheet) =>
    sheet.id === sheetId && sheet.tables?.some((table) => gone.has(table)) ? { ...sheet, tables: sheet.tables.filter((table) => !gone.has(table)), dirty: true as const } : sheet,
  );
  removeTableReferencesEverywhere(
    model,
    victims.map((table) => table.name),
  );
}
