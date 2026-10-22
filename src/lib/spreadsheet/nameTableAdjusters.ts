// Defined names and tables follow their cells (SPEC-133 "Tables follow
// cells", "Tables and moves") — THIS FILE is the normative source for the
// `names` and `tables` entries of SPEC-124's REFERENCE_ADJUSTERS and SPEC-125's
// MOVE_ADJUSTERS (structure.ts and move.ts register these functions), the
// partial-table move refusal, and removeTables: a deleted table's references
// become #REF! in every cell formula, name, calculated and totals formula.
//
// names — every name's formula (not formulaUnparsed):
// - structural edit: through adjustFormula, host = its scope sheet, else the
//   edited sheet. A name's RELATIVE parts are offsets from the cell using it
//   (relative to A1, wrapping modulo the grid — `Above = Sheet1!A1048576`),
//   not positions, so only a reference whose parts along the edit's axis are
//   all absolute moves; a relative one keeps its offset (else every row insert
//   would turn `Above` into #REF!).
// - move: through repointFormula, host = its scope sheet, else the move's
//   source sheet (an unqualified reference names the cells being moved); a
//   name follows its cells when they are cut-pasted or drag-moved, as in
//   Excel. Only a reference whose every part is absolute moves (as above).
// A changed name is replaced copy-on-write and `namesDirty` set.
//
// tables — structural edit, on the edited sheet (rows; columns symmetric):
// | Edit | Effect |
// | insert at ≤ ref's first row | ref shifts |
// | insert with first < at ≤ last | ref grows (new columns `Column{n}`, smallest unused n, their header cells written as strings) |
// | delete band covering the header / totals row | headerRow / totalsRow false |
// | delete band covering every data row, or every column | the table is deleted |
// | other overlapping deletions | ref shrinks (deleted columns leave `columns`) |
// A changed table is replaced, `dirty`, and its sheet `dirty`; a change of
// its column layout drops an autoFilter `sourceXml` holding filterColumns
// (their colIds no longer fit). Then the A1 parts of every table's
// calculatedFormula and totalsFormula, on every sheet (host = the table's
// sheet), go through adjustFormula as cell formulas do, and references to a
// deleted table or column become #REF! everywhere.
//
// tables — move: a table whose whole ref lies inside the moved block moves
// by the offset onto the destination sheet (same id and partName, `dirty`;
// both sheets dirty, so their `<tableParts>` follow); references to it are
// by name and need nothing. A table lying wholly inside the destination and
// not moving is overwritten: deleted, its references #REF!. Every table's
// calculated and totals formulas go through repointFormula — a moved
// table's as moved formulas hosted on the source sheet. moveRange refuses
// (partialTableMove) a move whose source cuts through a table, or whose
// destination cuts through one not wholly inside the source:
// PARTIAL_TABLE_MOVE_MESSAGE.
import type { GridRange } from "./address";
import { adjustFormula } from "./formula/adjust";
import type { ReferenceNode } from "./formula/references";
import { joinTokens, tokenize } from "./formula/tokens";
import { rangesIntersect } from "./merges";
import type { DefinedName, Sheet, SheetTable, SheetTableColumn, SpreadsheetModel } from "./model";
import { moveDestination, repointFormula, type RangeMove } from "./move";
import { contains, SheetEdit } from "./sheetEdit";
import { adjustIndex, adjustSpan, editAxis, isInsert, type StructuralEdit } from "./structure";
import { dataBody, sheetTables } from "./tables";
import { removeTableReferencesEverywhere } from "./tableRemoval";

export { removeTableReferencesEverywhere, removeTables } from "./tableRemoval";

export const PARTIAL_TABLE_MOVE_MESSAGE = "You can't move part of a table";

/** Sheet id → name, for adjustFormula and repointFormula. */
function sheetNames(model: SpreadsheetModel): (sheetId: string) => string {
  const names = new Map(model.sheets.map((sheet) => [sheet.id, sheet.name] as const));
  return (sheetId) => names.get(sheetId) ?? "";
}

const sheetOf = (model: SpreadsheetModel, id: string): Sheet | undefined => model.sheets.find((sheet) => sheet.id === id);

// --- shared rewrites --------------------------------------------------------------------------

/** Every parsed name's formula through `rewrite`, copy-on-write; sets
 * `namesDirty` when one changed. */
export function mapNameFormulas(model: SpreadsheetModel, rewrite: (name: DefinedName) => string): void {
  let changed = false;
  const names = model.names.map((name) => {
    if (name.formulaUnparsed) {
      return name;
    }
    const formula = rewrite(name);
    if (formula === name.formula) {
      return name;
    }
    changed = true;
    return { ...name, formula };
  });
  if (changed) {
    model.names = names;
    model.namesDirty = true;
  }
}

/** Every table's calculated and totals formulas, on every sheet, through
 * `rewrite(text, the table's sheet id, table)`; a changed table is replaced
 * and `dirty`, its sheet `dirty`. */
export function mapTableFormulas(model: SpreadsheetModel, rewrite: (text: string, hostSheetId: string, table: SheetTable) => string): void {
  let changed = false;
  const sheets = model.sheets.map((sheet) => {
    if (!sheet.tables || sheet.tables.length === 0) {
      return sheet;
    }
    let tablesChanged = false;
    const tables = sheet.tables.map((table) => {
      let columnsChanged = false;
      const columns = table.columns.map((column) => {
        const calculatedFormula = column.calculatedFormula === undefined ? undefined : rewrite(column.calculatedFormula, sheet.id, table);
        const totalsFormula = column.totalsFormula === undefined ? undefined : rewrite(column.totalsFormula, sheet.id, table);
        if (calculatedFormula === column.calculatedFormula && totalsFormula === column.totalsFormula) {
          return column;
        }
        columnsChanged = true;
        const next: SheetTableColumn = { ...column };
        if (calculatedFormula !== undefined) {
          next.calculatedFormula = calculatedFormula;
        }
        if (totalsFormula !== undefined) {
          next.totalsFormula = totalsFormula;
        }
        return next;
      });
      if (!columnsChanged) {
        return table;
      }
      tablesChanged = true;
      return { ...table, columns, dirty: true as const };
    });
    if (!tablesChanged) {
      return sheet;
    }
    changed = true;
    return { ...sheet, tables, dirty: true as const };
  });
  if (changed) {
    model.sheets = sheets;
  }
}



/** `text` with each reference token that `follows` passed through `rewrite`
 * on its own (its text is a formula of one reference). */
function rewriteFollowing(text: string, follows: (ref: ReferenceNode) => boolean, rewrite: (reference: string) => string): string {
  const tokens = tokenize(text);
  let changed = false;
  for (const token of tokens) {
    if (token.kind !== "ref" || !token.ref || !follows(token.ref)) {
      continue;
    }
    const next = rewrite(token.text);
    if (next !== token.text) {
      token.text = next;
      changed = true;
    }
  }
  return changed ? joinTokens(tokens) : text;
}

const partsOf = (ref: ReferenceNode): ReferenceNode["start"][] => (ref.end ? [ref.start, ref.end] : [ref.start]);

/** Whether a reference's parts along `axis` are absolute (a part without
 * that axis — a whole column under a row edit — never moves anyway). */
function absoluteAlong(ref: ReferenceNode, axis: "row" | "col"): boolean {
  return partsOf(ref).every((part) => (axis === "row" ? part.row === undefined || part.rowAbs : part.col === undefined || part.colAbs));
}

const fullyAbsolute = (ref: ReferenceNode): boolean => absoluteAlong(ref, "row") && absoluteAlong(ref, "col");

// --- names ---------------------------------------------------------------------------------------

/** The `names` REFERENCE_ADJUSTER (see the module header). */
export function adjustNamesForEdit(model: SpreadsheetModel, edit: StructuralEdit): void {
  if (model.names.length === 0) {
    return;
  }
  const nameOf = sheetNames(model);
  const axis = editAxis(edit);
  mapNameFormulas(model, (name) =>
    rewriteFollowing(
      name.formula,
      (ref) => absoluteAlong(ref, axis),
      (reference) => adjustFormula(reference, name.scopeSheetId ?? edit.sheetId, edit, nameOf),
    ),
  );
}

/** The `names` MOVE_ADJUSTER (see the module header). */
export function moveNamesForMove(model: SpreadsheetModel, move: RangeMove): void {
  if (model.names.length === 0) {
    return;
  }
  const nameOf = sheetNames(model);
  mapNameFormulas(model, (name) =>
    rewriteFollowing(name.formula, fullyAbsolute, (reference) => repointFormula(reference, name.scopeSheetId ?? move.fromSheetId, move, nameOf)),
  );
}

// --- tables: structural edits --------------------------------------------------------------------

interface Followed {
  table: SheetTable;
  /** Header cells to write: the new columns' names. */
  headers: { row: number; col: number; name: string }[];
  /** Names of the columns that went. */
  removedColumns: string[];
}

/** `count` new columns for a table holding `columns`: `Column{n}`, smallest
 * unused n; numbered `xlsxId`s (1 + max) when every column has one. */
function newColumns(columns: readonly SheetTableColumn[], count: number): SheetTableColumn[] {
  const used = new Set(columns.map((column) => column.name.toUpperCase()));
  const numbered = columns.length > 0 && columns.every((column) => column.xlsxId !== undefined);
  let xlsxId = Math.max(0, ...columns.map((column) => column.xlsxId ?? 0));
  const out: SheetTableColumn[] = [];
  for (let n = 1; out.length < count; n += 1) {
    const name = `Column${n}`;
    if (!used.has(name.toUpperCase())) {
      used.add(name.toUpperCase());
      xlsxId += 1;
      out.push(numbered ? { name, xlsxId } : { name });
    }
  }
  return out;
}

/** A table after `edit` on its sheet (see the module header): the same
 * object when it does not change, null when it is deleted. */
function followEdit(table: SheetTable, edit: StructuralEdit): SheetTable | Followed | null {
  const rows = editAxis(edit) === "row";
  const { ref } = table;
  const lo = rows ? ref.top : ref.left;
  const hi = rows ? ref.bottom : ref.right;
  if (edit.at > hi) {
    return table;
  }
  const span = adjustSpan(lo, hi, edit);
  if (span === null) {
    return null;
  }
  const [s, e] = span;
  const next: SheetTable = { ...table, ref: rows ? { ...ref, top: s, bottom: e } : { ...ref, left: s, right: e }, dirty: true };
  const followed: Followed = { table: next, headers: [], removedColumns: [] };
  if (rows) {
    const body = dataBody(table);
    if (adjustSpan(body.top, body.bottom, edit) === null) {
      return null; // every data row went
    }
    next.headerRow = table.headerRow && adjustIndex(lo, edit) !== null;
    next.totalsRow = table.totalsRow && adjustIndex(hi, edit) !== null;
    return followed;
  }
  let columns = table.columns;
  if (isInsert(edit)) {
    if (edit.at > lo) {
      const offset = edit.at - lo;
      const added = newColumns(table.columns, edit.count);
      columns = [...table.columns.slice(0, offset), ...added, ...table.columns.slice(offset)];
      if (table.headerRow) {
        followed.headers = added.map((column, k) => ({ row: ref.top, col: edit.at + k, name: column.name }));
      }
    }
  } else {
    columns = table.columns.filter((_, i) => adjustIndex(lo + i, edit) !== null);
    if (columns.length === 0) {
      return null; // every column went
    }
  }
  // Columns pushed past the grid's last column go too.
  if (columns.length > e - s + 1) {
    columns = columns.slice(0, e - s + 1);
  }
  if (columns !== table.columns) {
    const kept = new Set(columns);
    followed.removedColumns = table.columns.filter((column) => !kept.has(column)).map((column) => column.name);
    next.columns = columns;
    if (table.autoFilter?.sourceXml?.includes("filterColumn")) {
      next.autoFilter = {};
    }
  }
  return followed;
}

/** The `tables` REFERENCE_ADJUSTER (see the module header). */
export function adjustTablesForEdit(model: SpreadsheetModel, edit: StructuralEdit): void {
  const removed: string[] = [];
  const removedColumns: { table: string; column: string }[] = [];
  const sheet = sheetOf(model, edit.sheetId);
  if (sheet?.tables && sheet.tables.length > 0) {
    let changed = false;
    const tables: SheetTable[] = [];
    const batch = new SheetEdit(sheet);
    for (const table of sheet.tables) {
      const result = followEdit(table, edit);
      if (result === table) {
        tables.push(table);
        continue;
      }
      changed = true;
      if (result === null) {
        removed.push(table.name);
        continue;
      }
      const followed = result as Followed;
      tables.push(followed.table);
      for (const { row, col, name } of followed.headers) {
        batch.change(row, col, { content: { value: name } });
      }
      removedColumns.push(...followed.removedColumns.map((column) => ({ table: table.name, column })));
    }
    if (changed) {
      const rows = batch.rows();
      model.sheets = model.sheets.map((each) => (each === sheet ? { ...sheet, rows, tables, dirty: true as const } : each));
    }
  }
  const nameOf = sheetNames(model);
  mapTableFormulas(model, (text, hostSheetId) => adjustFormula(text, hostSheetId, edit, nameOf));
  removeTableReferencesEverywhere(model, removed, removedColumns);
}

// --- tables: moves -------------------------------------------------------------------------------

function shifted(range: GridRange, dRow: number, dCol: number): GridRange {
  return { top: range.top + dRow, left: range.left + dCol, bottom: range.bottom + dRow, right: range.right + dCol };
}

/** Whether `move` would move part of a table (see the module header). */
export function partialTableMove(model: SpreadsheetModel, move: RangeMove): boolean {
  const from = sheetOf(model, move.fromSheetId);
  const to = sheetOf(model, move.toSheetId);
  if (!from || !to) {
    return false;
  }
  if (sheetTables(from).some((table) => rangesIntersect(table.ref, move.from) && !contains(move.from, table.ref))) {
    return true;
  }
  const dest = moveDestination(move);
  return sheetTables(to).some((table) => rangesIntersect(table.ref, dest) && !contains(dest, table.ref) && !(from === to && contains(move.from, table.ref)));
}

/** The `tables` MOVE_ADJUSTER (see the module header). */
export function moveTablesForMove(model: SpreadsheetModel, move: RangeMove): void {
  const from = sheetOf(model, move.fromSheetId);
  const to = sheetOf(model, move.toSheetId);
  const dRow = move.toTop - move.from.top;
  const dCol = move.toLeft - move.from.left;
  if (!from || !to || (from === to && dRow === 0 && dCol === 0)) {
    return;
  }
  const dest = moveDestination(move);
  const moving = sheetTables(from).filter((table) => contains(move.from, table.ref));
  const overwritten = sheetTables(to).filter((table) => !moving.includes(table) && contains(dest, table.ref));
  const arrived = new Map(moving.map((table) => [table, { ...table, ref: shifted(table.ref, dRow, dCol), dirty: true as const }] as const));
  if (moving.length > 0 || overwritten.length > 0) {
    model.sheets = model.sheets.map((sheet) => {
      if (sheet !== from && sheet !== to) {
        return sheet;
      }
      let tables = sheetTables(sheet).filter((table) => !overwritten.includes(table));
      if (sheet === from && sheet === to) {
        tables = tables.map((table) => arrived.get(table) ?? table);
      } else if (sheet === from) {
        tables = tables.filter((table) => !arrived.has(table));
      } else {
        tables = [...tables, ...arrived.values()];
      }
      return { ...sheet, tables, dirty: true as const };
    });
  }
  const nameOf = sheetNames(model);
  const movedIds = new Set(moving.map((table) => table.id));
  mapTableFormulas(model, (text, hostSheetId, table) =>
    movedIds.has(table.id) ? repointFormula(text, move.fromSheetId, move, nameOf, move.toSheetId) : repointFormula(text, hostSheetId, move, nameOf),
  );
  removeTableReferencesEverywhere(
    model,
    overwritten.map((table) => table.name),
  );
}
