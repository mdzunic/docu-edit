// What a change to names or tables makes recalculate (SPEC-133
// "Resolution"): exactly the formula cells whose text mentions a changed
// name or table — formulaMentions' token scan, cached per formula text —
// including cells showing #NAME? for a name just defined. A name whose own
// formula mentions a changed name or table counts as changed too (to a
// fixpoint), so a cell using `Total = Margin*2` follows `Margin`.
//
// SPEC-121's edit pipeline (useSpreadsheetUndo) calls nameTableDependents
// with the names and sheets before a step; the graph nodes of the cells it
// returns are rebuilt (refreshGraphCells), since what their names read moved
// while their text did not.
import { cellKey, type CellKey } from "./formula/dependencies";
import type { DefinedName, Sheet, SheetTable, SpreadsheetModel } from "./model";
import { formulaMentions } from "./names";

const keyOf = (name: DefinedName): string => name.name.toUpperCase();

function sameName(a: DefinedName, b: DefinedName): boolean {
  return a.formula === b.formula && a.scopeSheetId === b.scopeSheetId && a.formulaUnparsed === b.formulaUnparsed && a.name === b.name && a.sourceAttrs === b.sourceAttrs;
}

/** The upper-cased texts of the names added, removed or changed. */
export function changedNames(before: readonly DefinedName[], after: readonly DefinedName[]): Set<string> {
  const changed = new Set<string>();
  if (before === after) {
    return changed;
  }
  const remaining = [...after];
  for (const was of before) {
    const index = remaining.findIndex((now) => now === was || sameName(now, was));
    if (index < 0) {
      changed.add(keyOf(was));
    } else {
      remaining.splice(index, 1);
    }
  }
  for (const now of remaining) {
    changed.add(keyOf(now));
  }
  return changed;
}

function sameTable(a: SheetTable, b: SheetTable): boolean {
  return (
    a === b ||
    (a.name === b.name &&
      a.headerRow === b.headerRow &&
      a.totalsRow === b.totalsRow &&
      a.ref.top === b.ref.top &&
      a.ref.left === b.ref.left &&
      a.ref.bottom === b.ref.bottom &&
      a.ref.right === b.ref.right &&
      a.columns.length === b.columns.length &&
      a.columns.every((column, i) => column.name === b.columns[i]!.name))
  );
}

/** The upper-cased names of the tables whose name, `ref`, header/totals
 * rows or columns changed, appeared or went (both names of a rename). */
export function changedTables(before: readonly Sheet[], after: readonly Sheet[]): Set<string> {
  const changed = new Set<string>();
  const was = new Map<string, { table: SheetTable; sheetId: string }>();
  for (const sheet of before) {
    for (const table of sheet.tables ?? []) {
      was.set(table.id, { table, sheetId: sheet.id });
    }
  }
  for (const sheet of after) {
    for (const table of sheet.tables ?? []) {
      const old = was.get(table.id);
      was.delete(table.id);
      if (!old || old.sheetId !== sheet.id || !sameTable(old.table, table)) {
        changed.add(table.name.toUpperCase());
        if (old) {
          changed.add(old.table.name.toUpperCase());
        }
      }
    }
  }
  for (const { table } of was.values()) {
    changed.add(table.name.toUpperCase());
  }
  return changed;
}

/** Whether a formula mentions one of the names or tables. */
function mentions(text: string, names: ReadonlySet<string>, tables: ReadonlySet<string>): boolean {
  const found = formulaMentions(text);
  for (const name of found.names) {
    if (names.has(name) || tables.has(name)) {
      return true;
    }
  }
  for (const table of found.tables) {
    if (tables.has(table)) {
      return true;
    }
  }
  return false;
}

/** Every formula cell mentioning one of the names or tables (upper-cased),
 * sheet order then row-major; names mentioning them count as changed. */
export function cellsMentioning(model: Pick<SpreadsheetModel, "sheets" | "names">, names: ReadonlySet<string>, tables: ReadonlySet<string>): CellKey[] {
  if (names.size === 0 && tables.size === 0) {
    return [];
  }
  const widened = new Set(names);
  for (let grew = true; grew; ) {
    grew = false;
    for (const name of model.names) {
      if (!widened.has(keyOf(name)) && !name.formulaUnparsed && mentions(name.formula, widened, tables)) {
        widened.add(keyOf(name));
        grew = true;
      }
    }
  }
  const keys: CellKey[] = [];
  for (const sheet of model.sheets) {
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (cell.formula !== undefined && !cell.formulaUnparsed && mentions(cell.formula, widened, tables)) {
          keys.push(cellKey(sheet.id, row.index, cell.col));
        }
      }
    }
  }
  return keys;
}

/** The formula cells a step's name and table changes reach (see the
 * module header); none when neither changed. */
export function nameTableDependents(model: SpreadsheetModel, namesBefore: readonly DefinedName[], sheetsBefore: readonly Sheet[]): CellKey[] {
  const names = changedNames(namesBefore, model.names);
  const tablesMoved = sheetsBefore.length !== model.sheets.length || sheetsBefore.some((sheet, i) => sheet.tables !== model.sheets[i]!.tables);
  const tables = tablesMoved ? changedTables(sheetsBefore, model.sheets) : new Set<string>();
  return cellsMentioning(model, names, tables);
}
