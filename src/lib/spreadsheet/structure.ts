// Structural edits — whole rows and columns inserted or deleted (SPEC-124).
// THIS FILE is the normative source for StructuralEdit, the geometry every
// reference follows (adjustIndex / adjustRange), the REFERENCE_ADJUSTERS
// registry and applyStructuralEdit.
//
// Geometry, along the edit's axis (`last` = MAX_ROWS − 1 or MAX_COLS − 1):
// | Case | Insert n at a | Delete [a, a+n) |
// | index i < a | i | i |
// | index i ≥ a | i+n; > last → null | in the band → null; i ≥ a+n → i−n |
// | range s..e | s ≥ a → s+n; e ≥ a → e+n, clamped to last (a start past last → null) | both in the band → null; s in it → a; e in it → a−1; e ≥ a+n → e−n |
// So an insert AT a range's first row moves it, one strictly inside expands
// it and one just after its last row leaves it alone (Excel); the other axis
// is never touched. null means "no longer exists".
//
// applyStructuralEdit checks the refusals before anything changes — a
// protected sheet unless its protection permits the operation (SPEC-122's
// message), an edit splitting an array formula's range (an insert strictly
// inside it, a delete band partly inside it: PARTIAL_ARRAY_MESSAGE), an
// insert pushing a cell holding a value or formula past the grid's end
// (INSERT_PUSHES_DATA_MESSAGE) — then runs every adjuster in registry order
// and marks the edited sheet dirty. Each adjuster replaces what it changes —
// a new sheet, rows, cells, merges, columns, frozen or selection object, a
// new OpaquePart — and never mutates one an undo snapshot may hold
// (SPEC-121's snapshots are by reference); a structure of its kind that does
// not move keeps its identity.
//
// | Order | id | Moves |
// | 1 | rows | row edits: Row objects re-indexed, the band removed, inserted rows' inherited properties and styled cells; column edits: Cell.col shifted per row, inherited cell styles |
// | 2 | formulas | adjustFormula over every sheet's Cell.formula (not formulaUnparsed), Cell.arrayRange and a data table's ref/r1/r2 via the geometry |
// | 3 | merges | Sheet.merges via adjustRange; single-cell results dropped |
// | 4 | frozen | the band [0, rows−1] / [0, cols−1] via adjustRange |
// | 5 | selection | Sheet.selection's ranges and active cell (a range or cell whose area is gone stays where it was, clamped into the sheet) |
// | 6 | columns | ColumnSpans shifted, split, removed; inserted columns copy the left span's width and style (not `hidden`); adjacent equal spans re-joined |
// | 7 | retained | rewriteRetainedReferences(model, structuralMapper(model, edit)) |
// | 8 | autofilter | SPEC-128: Sheet.autoFilter (range, colIds, its sort record) and Sheet.sortState (ref, key columns) — see "8. autofilter" |
// | 9 | conditionalFormats | SPEC-129: Sheet.conditionalFormats — the edited sheet's ranges via adjustRange (a fully deleted one dropped; an emptied format removed) and its cellIs/expression formulas re-anchored with adjustAnchoredFormula; every other formula a rule carries (cfFormulas.ts), on every sheet, via adjustFormula — see "9. conditionalFormats" |
// | 10 | validations | SPEC-130: Sheet.validations — the edited sheet's ranges via adjustRange (a fully deleted one dropped; an emptied rule removed) and formula1/formula2 re-anchored with adjustAnchoredFormula; every other sheet's rule formulas via adjustFormula — see "10. validations" |
// | 11 | drawings | SPEC-131: the edited sheet's Sheet.drawings per drawings.ts's anchor-follow table (moveAndSize drawings move, grow, shrink and go with their cells; move drawings shift with their `from`; fixed drawings never move); an xlsx sheet holding drawings gets `drawingsDirty` |
// | 12 | charts | SPEC-132: adjustFormula over every modelled chart's refs on every sheet (host = the chart's sheet) — `refsDirty` where one changed — then a refresh of each refreshable chart whose refs changed or read the edited sheet (`cachesDirty` when its visual changed; read-only charts only adjust); ods chart frames' notify-on-update-of-ranges synced (charts.ts adjustChartsForEdit) |
// | 13 | names | SPEC-133: every DefinedName's formula via adjustFormula, host = its scope sheet, else the edited sheet — a name's relative parts are offsets and stay (nameTableAdjusters.ts) |
// | 14 | tables | SPEC-133 "Tables follow cells": the edited sheet's Sheet.tables shift, grow (new columns `Column{n}`, headers written), lose a deleted header/totals row, shrink or go; every table's calculated/totals formulas via adjustFormula; references to a deleted table or column #REF! everywhere (nameTableAdjusters.ts) |
// Later specs append theirs, in landing order — notes and hyperlinks (135),
// pageBreaks (136), pivots (139) — each with the registry contract test in
// structure.spec.ts, retiring its rows from retainedReferences' tables in the
// same change.
//
// Inserted rows take the row above's styleId, heightPt and customHeight and
// its non-zero cell styles as styled empty cells (columns: the column to the
// left; inserting at index 0: the old first row or column); when that would
// materialise more than MAX_FORMAT_MATERIALISED cells only the row and column
// properties are copied. Formatting-only rows, cells, spans and merges pushed
// past the end are dropped. Deleting every row leaves an empty sheet.
// Inserted rows are visible: they inherit neither `hidden` nor SPEC-128's
// `filterHidden`.
import type { CellRef, GridRange } from "./address";
import { adjustHyperlinksForEdit, adjustNotesForEdit } from "./annotations";
import { adjustPageBreaksForEdit } from "./pageSetup";
import { isEmptyRow, PROTECTED_CELL_MESSAGE } from "./cells";
import { mapFormatFormulas, mapFormatsFormulas, withoutFormatSource } from "./cfFormulas";
import { formatOrigin, renumberPriorities, sheetFormats, type ConditionalFormat } from "./conditionalFormats";
import { adjustChartsForEdit } from "./charts";
import type { AutoFilter, FilterColumn } from "./filter";
import { adjustDrawingsForEdit } from "./drawings";
import { adjustNamesForEdit, adjustTablesForEdit } from "./nameTableAdjusters";
import { adjustAnchoredFormula, adjustFormula, adjustFormulaReferences } from "./formula/adjust";
import { MAX_COLS, MAX_ROWS, type Cell, type ColumnSpan, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { adjustOdsFilterCriteriaRanges } from "./ods/retainedReferences";
import { firstRowAtOrAfter, MAX_FORMAT_MATERIALISED } from "./restyle";
import { rewriteRetainedReferences, type RetainedReferenceMapper } from "./retainedReferences";
import type { SheetSelection } from "./selection";
import type { SheetSortKey, SortState } from "./sort";
import { resolveRetainedSheetName } from "./sourceNames";
import {
  mapValidationFormulas,
  mapValidationsFormulas,
  sheetValidations,
  validationOrigin,
  withoutValidationSource,
  withX14Placement,
  type DataValidation,
} from "./validation";
import { LazyParts } from "./xlsx/package";

export { rewriteRetainedReferences, type RetainedReferenceMapper } from "./retainedReferences";

export type StructuralEdit = {
  kind: "insertRows" | "deleteRows" | "insertCols" | "deleteCols";
  sheetId: string;
  at: number; // 0-based first row/column inserted or deleted
  count: number; // ≥ 1; at + count ≤ MAX_ROWS / MAX_COLS
};

export interface ReferenceAdjuster {
  id: string;
  adjust(model: SpreadsheetModel, edit: StructuralEdit): void;
}

export type StructuralResult = { ok: true } | { ok: false; message: string };

export const INSERT_PUSHES_DATA_MESSAGE = "Can't insert — that would push filled cells off the end of the sheet";
export const PARTIAL_ARRAY_MESSAGE = "You can't change part of an array";
export const NOT_A_GRID_MESSAGE = "Only a worksheet's rows and columns can be changed";

// --- geometry ------------------------------------------------------------------------------

/** The axis an edit moves along. */
export function editAxis(edit: Pick<StructuralEdit, "kind">): "row" | "col" {
  return edit.kind === "insertRows" || edit.kind === "deleteRows" ? "row" : "col";
}

export function isInsert(edit: Pick<StructuralEdit, "kind">): boolean {
  return edit.kind === "insertRows" || edit.kind === "insertCols";
}

/** The last index of an axis. */
export function lastIndex(axis: "row" | "col"): number {
  return axis === "row" ? MAX_ROWS - 1 : MAX_COLS - 1;
}

/** An index along the edit's axis per the geometry table. */
export function adjustIndex(index: number, edit: StructuralEdit): number | null {
  const { at, count } = edit;
  if (index < at) {
    return index;
  }
  if (isInsert(edit)) {
    const moved = index + count;
    return moved > lastIndex(editAxis(edit)) ? null : moved;
  }
  return index < at + count ? null : index - count;
}

/** A span [start, end] along the edit's axis per the geometry table. */
export function adjustSpan(start: number, end: number, edit: StructuralEdit): [number, number] | null {
  const { at, count } = edit;
  if (isInsert(edit)) {
    const last = lastIndex(editAxis(edit));
    const s = start >= at ? start + count : start;
    if (s > last) {
      return null;
    }
    const e = end >= at ? Math.min(end + count, last) : end;
    return [s, e];
  }
  const bandEnd = at + count - 1;
  const inBand = (i: number): boolean => i >= at && i <= bandEnd;
  if (inBand(start) && inBand(end)) {
    return null;
  }
  const s = start < at ? start : inBand(start) ? at : start - count;
  const e = end < at ? end : inBand(end) ? at - 1 : end - count;
  return [s, e];
}

/** A range per the geometry table: the same object when it does not move,
 * null when it no longer exists. */
export function adjustRange(range: GridRange, edit: StructuralEdit): GridRange | null {
  const rows = editAxis(edit) === "row";
  const span = rows ? adjustSpan(range.top, range.bottom, edit) : adjustSpan(range.left, range.right, edit);
  if (span === null) {
    return null;
  }
  const [s, e] = span;
  if (rows) {
    return s === range.top && e === range.bottom ? range : { ...range, top: s, bottom: e };
  }
  return s === range.left && e === range.right ? range : { ...range, left: s, right: e };
}

/** A cell position along the edit's axis (`null`: deleted). */
function adjustCell(at: CellRef, edit: StructuralEdit): CellRef | null {
  const axis = editAxis(edit);
  const index = adjustIndex(axis === "row" ? at.row : at.col, edit);
  if (index === null) {
    return null;
  }
  return axis === "row" ? { row: index, col: at.col } : { row: at.row, col: index };
}

// --- shared helpers ------------------------------------------------------------------------

/** Replaces the sheet `id` with `next(sheet)` when that is a new object. */
function updateSheet(model: SpreadsheetModel, id: string, next: (sheet: Sheet) => Sheet): void {
  let changed = false;
  const sheets = model.sheets.map((sheet) => {
    if (sheet.id !== id) {
      return sheet;
    }
    const updated = next(sheet);
    changed ||= updated !== sheet;
    return updated;
  });
  if (changed) {
    model.sheets = sheets;
  }
}

/** Whether a cell holds a value or a formula (content an insert must not
 * push off the grid; a merge keeps only the upper-left one). */
export function holdsContent(cell: Cell): boolean {
  return cell.value !== null || cell.formula !== undefined;
}

/** The first cell at or after `col` (an index into `cells`). */
function firstCellAtOrAfter(cells: readonly Cell[], col: number): number {
  let lo = 0;
  let hi = cells.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (cells[mid]!.col < col) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return lo;
}

function cellAt(row: Row | undefined, col: number): Cell | undefined {
  if (!row) {
    return undefined;
  }
  const cell = row.cells[firstCellAtOrAfter(row.cells, col)];
  return cell?.col === col ? cell : undefined;
}

function rowAt(rows: readonly Row[], index: number): Row | undefined {
  const row = rows[firstRowAtOrAfter(rows, index)];
  return row?.index === index ? row : undefined;
}

/** A style id an inserted row or column inherits (0 = none). */
function inheritedStyle(styleId: number | undefined): number | undefined {
  return styleId !== undefined && styleId !== 0 ? styleId : undefined;
}

// --- refusals ------------------------------------------------------------------------------

const PERMITS = {
  insertRows: "allowInsertRows",
  insertCols: "allowInsertColumns",
  deleteRows: "allowDeleteRows",
  deleteCols: "allowDeleteColumns",
} as const;

/** Whether a protected sheet's protection permits the edit. */
function protectionPermits(sheet: Sheet, edit: StructuralEdit): boolean {
  return !sheet.protection || sheet.protection[PERMITS[edit.kind]] === true;
}

/** Whether `edit` would split an array range: an insert strictly inside it,
 * or a delete band partly inside it. */
export function splitsArray(range: GridRange, edit: StructuralEdit): boolean {
  const rows = editAxis(edit) === "row";
  const lo = rows ? range.top : range.left;
  const hi = rows ? range.bottom : range.right;
  if (isInsert(edit)) {
    return edit.at > lo && edit.at <= hi;
  }
  const bandEnd = edit.at + edit.count - 1;
  const intersects = edit.at <= hi && bandEnd >= lo;
  return intersects && !(edit.at <= lo && bandEnd >= hi);
}

/** Whether an insert would push a cell holding a value or formula off the
 * grid's end. */
function pushesData(sheet: Sheet, edit: StructuralEdit): boolean {
  if (!isInsert(edit)) {
    return false;
  }
  if (editAxis(edit) === "row") {
    const limit = MAX_ROWS - edit.count;
    for (let r = sheet.rows.length - 1; r >= 0 && sheet.rows[r]!.index >= limit; r -= 1) {
      if (sheet.rows[r]!.cells.some(holdsContent)) {
        return true;
      }
    }
    return false;
  }
  const limit = MAX_COLS - edit.count;
  return sheet.rows.some((row) => {
    for (let c = row.cells.length - 1; c >= 0 && row.cells[c]!.col >= limit; c -= 1) {
      if (holdsContent(row.cells[c]!)) {
        return true;
      }
    }
    return false;
  });
}

/** Why `edit` cannot be made, or null (see the module header). */
export function structuralRefusal(model: SpreadsheetModel, edit: StructuralEdit): string | null {
  const sheet = model.sheets.find((each) => each.id === edit.sheetId);
  const axisLength = editAxis(edit) === "row" ? MAX_ROWS : MAX_COLS;
  if (!sheet || sheet.kind !== "grid" || edit.count < 1 || edit.at < 0 || edit.at + edit.count > axisLength) {
    return NOT_A_GRID_MESSAGE;
  }
  if (!protectionPermits(sheet, edit)) {
    return PROTECTED_CELL_MESSAGE;
  }
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      if (cell.arrayRange && splitsArray(cell.arrayRange, edit)) {
        return PARTIAL_ARRAY_MESSAGE;
      }
    }
  }
  return pushesData(sheet, edit) ? INSERT_PUSHES_DATA_MESSAGE : null;
}

// --- 1. rows ---------------------------------------------------------------------------------

/** The rows inserted at `at`: the source row's properties and its styled
 * cells as styled empty cells (properties only past the cap). */
function inheritedRows(source: Row | undefined, at: number, count: number): Row[] {
  if (!source) {
    return [];
  }
  const props: Partial<Row> = {};
  if (inheritedStyle(source.styleId) !== undefined) {
    props.styleId = source.styleId;
  }
  if (source.heightPt !== undefined) {
    props.heightPt = source.heightPt;
  }
  if (source.customHeight) {
    props.customHeight = true;
  }
  let styled = source.cells.filter((cell) => inheritedStyle(cell.styleId) !== undefined);
  if (styled.length * count > MAX_FORMAT_MATERIALISED) {
    styled = [];
  }
  if (styled.length === 0 && Object.keys(props).length === 0) {
    return [];
  }
  return Array.from({ length: count }, (_, k): Row => ({
    index: at + k,
    cells: styled.map((cell): Cell => ({ col: cell.col, value: null, styleId: cell.styleId })),
    ...props,
  }));
}

function insertRows(sheet: Sheet, at: number, count: number): Row[] {
  const rows = sheet.rows;
  const first = firstRowAtOrAfter(rows, at);
  const source = at > 0 ? rowAt(rows, at - 1) : rows[0]?.index === 0 ? rows[0] : undefined;
  const inserted = inheritedRows(source, at, count);
  if (first === rows.length && inserted.length === 0) {
    return rows;
  }
  // concat, not push(...): an insert may add a million rows.
  const out = rows.slice(0, first).concat(inserted);
  for (let r = first; r < rows.length; r += 1) {
    const row = rows[r]!;
    if (row.index + count > MAX_ROWS - 1) {
      break; // formatting only: an insert pushing data is refused
    }
    out.push({ ...row, index: row.index + count });
  }
  return out;
}

function deleteRows(sheet: Sheet, at: number, count: number): Row[] {
  const rows = sheet.rows;
  const first = firstRowAtOrAfter(rows, at);
  if (first === rows.length) {
    return rows;
  }
  const out = rows.slice(0, first);
  for (let r = firstRowAtOrAfter(rows, at + count); r < rows.length; r += 1) {
    out.push({ ...rows[r]!, index: rows[r]!.index - count });
  }
  return out;
}

/** A row with columns inserted at `at`; `inherit` says whether the
 * source column's style is materialised. */
function insertColsInRow(row: Row, at: number, count: number, inherit: boolean): Row {
  const cells = row.cells;
  const first = firstCellAtOrAfter(cells, at);
  const source = inherit ? (at > 0 ? cellAt(row, at - 1) : cellAt(row, 0)) : undefined;
  const styleId = inheritedStyle(source?.styleId);
  if (first === cells.length && styleId === undefined) {
    return row;
  }
  const out = cells.slice(0, first);
  if (styleId !== undefined) {
    for (let k = 0; k < count; k += 1) {
      out.push({ col: at + k, value: null, styleId });
    }
  }
  for (let c = first; c < cells.length; c += 1) {
    const cell = cells[c]!;
    if (cell.col + count > MAX_COLS - 1) {
      break; // formatting only: an insert pushing data is refused
    }
    out.push({ ...cell, col: cell.col + count });
  }
  return { ...row, cells: out };
}

function deleteColsInRow(row: Row, at: number, count: number): Row | undefined {
  const cells = row.cells;
  const first = firstCellAtOrAfter(cells, at);
  if (first === cells.length) {
    return row;
  }
  const out = cells.slice(0, first);
  for (let c = firstCellAtOrAfter(cells, at + count); c < cells.length; c += 1) {
    out.push({ ...cells[c]!, col: cells[c]!.col - count });
  }
  const next = { ...row, cells: out };
  return isEmptyRow(next) ? undefined : next;
}

function columnEditRows(sheet: Sheet, edit: StructuralEdit): Row[] {
  let inherit = false;
  if (isInsert(edit)) {
    const sourceCol = edit.at > 0 ? edit.at - 1 : 0;
    const styledRows = sheet.rows.filter((row) => inheritedStyle(cellAt(row, sourceCol)?.styleId) !== undefined).length;
    inherit = styledRows * edit.count <= MAX_FORMAT_MATERIALISED;
  }
  let changed = false;
  const rows: Row[] = [];
  for (const row of sheet.rows) {
    const next = isInsert(edit) ? insertColsInRow(row, edit.at, edit.count, inherit) : deleteColsInRow(row, edit.at, edit.count);
    changed ||= next !== row;
    if (next) {
      rows.push(next);
    }
  }
  return changed ? rows : sheet.rows;
}

const rowsAdjuster: ReferenceAdjuster = {
  id: "rows",
  adjust(model, edit) {
    updateSheet(model, edit.sheetId, (sheet) => {
      let rows: Row[];
      if (editAxis(edit) === "row") {
        rows = isInsert(edit) ? insertRows(sheet, edit.at, edit.count) : deleteRows(sheet, edit.at, edit.count);
      } else {
        rows = columnEditRows(sheet, edit);
      }
      return rows === sheet.rows ? sheet : { ...sheet, rows };
    });
  },
};

// --- 2. formulas -----------------------------------------------------------------------------

const DATA_TABLE_REFS = /(\s(?:ref|r1|r2)=")([^"]*)(")/g;

/** A data table's `<f t="dataTable" ref r1 r2/>` with its references
 * moved; references gone are left as they were. */
function adjustDataTable(xml: string, edit: StructuralEdit): string {
  return xml.replace(DATA_TABLE_REFS, (whole, head: string, text: string, tail: string) => {
    const next = adjustFormulaReferences(text, edit, () => true);
    return next === text || next.includes("#REF!") ? whole : `${head}${next}${tail}`;
  });
}

const formulasAdjuster: ReferenceAdjuster = {
  id: "formulas",
  adjust(model, edit) {
    const names = new Map(model.sheets.map((sheet) => [sheet.id, sheet.name] as const));
    const nameOf = (id: string): string => names.get(id) ?? "";
    let sheetsChanged = false;
    const sheets = model.sheets.map((sheet) => {
      const own = sheet.id === edit.sheetId;
      let rowsChanged = false;
      const rows = sheet.rows.map((row) => {
        let cellsChanged = false;
        const cells = row.cells.map((cell) => {
          let next = cell;
          // Another sheet's formula names the edited sheet with a `!`.
          if (cell.formula !== undefined && !cell.formulaUnparsed && (own || cell.formula.includes("!"))) {
            const formula = adjustFormula(cell.formula, sheet.id, edit, nameOf);
            if (formula !== cell.formula) {
              next = { ...next, formula };
            }
          }
          if (own && cell.arrayRange) {
            const arrayRange = adjustRange(cell.arrayRange, edit);
            if (arrayRange && arrayRange !== cell.arrayRange) {
              next = { ...next, arrayRange };
            }
          }
          if (own && cell.dataTableXml !== undefined) {
            const dataTableXml = adjustDataTable(cell.dataTableXml, edit);
            if (dataTableXml !== cell.dataTableXml) {
              next = { ...next, dataTableXml };
            }
          }
          cellsChanged ||= next !== cell;
          return next;
        });
        if (!cellsChanged) {
          return row;
        }
        rowsChanged = true;
        return { ...row, cells };
      });
      if (!rowsChanged) {
        return sheet;
      }
      sheetsChanged = true;
      return { ...sheet, rows, dirty: true as const };
    });
    if (sheetsChanged) {
      model.sheets = sheets;
    }
  },
};

// --- 3. merges -------------------------------------------------------------------------------

const isSingleCell = (range: GridRange): boolean => range.top === range.bottom && range.left === range.right;

const mergesAdjuster: ReferenceAdjuster = {
  id: "merges",
  adjust(model, edit) {
    updateSheet(model, edit.sheetId, (sheet) => {
      let changed = false;
      const merges: GridRange[] = [];
      for (const merge of sheet.merges) {
        const next = adjustRange(merge, edit);
        changed ||= next !== merge;
        if (next && !isSingleCell(next)) {
          merges.push(next);
        } else {
          changed = true;
        }
      }
      return changed ? { ...sheet, merges } : sheet;
    });
  },
};

// --- 4. frozen -------------------------------------------------------------------------------

const frozenAdjuster: ReferenceAdjuster = {
  id: "frozen",
  adjust(model, edit) {
    updateSheet(model, edit.sheetId, (sheet) => {
      const frozen = sheet.frozen;
      const rows = editAxis(edit) === "row";
      const count = frozen ? (rows ? frozen.rows : frozen.cols) : 0;
      if (!frozen || count === 0) {
        return sheet;
      }
      const span = adjustSpan(0, count - 1, edit);
      const next = span === null ? 0 : span[1] + 1;
      if (next === count) {
        return sheet;
      }
      const moved = rows ? { rows: next, cols: frozen.cols } : { rows: frozen.rows, cols: next };
      const updated: Sheet = { ...sheet, frozen: moved };
      if (moved.rows === 0 && moved.cols === 0) {
        delete updated.frozen;
      }
      return updated;
    });
  },
};

// --- 5. selection ----------------------------------------------------------------------------

/** A position clamped into the sheet. */
function clampCell(at: CellRef): CellRef {
  return { row: Math.min(Math.max(at.row, 0), MAX_ROWS - 1), col: Math.min(Math.max(at.col, 0), MAX_COLS - 1) };
}

/** A selection after `edit` (see the module header): the same object when
 * nothing moves. */
export function adjustSelection(selection: SheetSelection, edit: StructuralEdit): SheetSelection {
  let changed = false;
  const ranges = selection.ranges.map((range) => {
    const next = adjustRange(range, edit) ?? range;
    changed ||= next !== range;
    return next;
  });
  const moved = adjustCell(selection.active, edit);
  const active = clampCell(moved ?? selection.active);
  changed ||= active.row !== selection.active.row || active.col !== selection.active.col;
  return changed ? { ...selection, ranges, active } : selection;
}

const selectionAdjuster: ReferenceAdjuster = {
  id: "selection",
  adjust(model, edit) {
    updateSheet(model, edit.sheetId, (sheet) => {
      if (!sheet.selection) {
        return sheet;
      }
      const selection = adjustSelection(sheet.selection, edit);
      return selection === sheet.selection ? sheet : { ...sheet, selection };
    });
  },
};

// --- 6. columns ------------------------------------------------------------------------------

/** Whether two spans carry the same properties (and can be one span). */
function sameSpanProps(a: ColumnSpan, b: ColumnSpan): boolean {
  return (
    a.widthPx === b.widthPx &&
    a.hidden === b.hidden &&
    a.styleId === b.styleId &&
    a.sourceWidthChars === b.sourceWidthChars &&
    a.sourceAttrs === b.sourceAttrs &&
    a.sourceStyleName === b.sourceStyleName
  );
}

/** Adjacent spans with the same properties joined into one. */
export function joinSpans(spans: readonly ColumnSpan[]): ColumnSpan[] {
  const out: ColumnSpan[] = [];
  for (const span of spans) {
    const last = out.at(-1);
    if (last && last.max + 1 === span.min && sameSpanProps(last, span)) {
      out[out.length - 1] = { ...last, max: span.max };
    } else {
      out.push(span);
    }
  }
  return out;
}

/** A span with its run [min, max] and every property but `hidden` and the
 * unmodelled attributes of `like` — the inserted columns' span. */
function insertedSpan(like: ColumnSpan, min: number, max: number): ColumnSpan | undefined {
  const span: ColumnSpan = { min, max };
  if (like.widthPx !== undefined) {
    span.widthPx = like.widthPx;
  }
  if (like.styleId !== undefined) {
    span.styleId = like.styleId;
  }
  if (like.sourceWidthChars !== undefined) {
    span.sourceWidthChars = like.sourceWidthChars;
  }
  if (like.sourceStyleName !== undefined) {
    span.sourceStyleName = like.sourceStyleName;
  }
  return Object.keys(span).length > 2 ? span : undefined;
}

function insertColumns(columns: readonly ColumnSpan[], at: number, count: number): ColumnSpan[] {
  const last = MAX_COLS - 1;
  const sourceCol = at > 0 ? at - 1 : 0;
  const source = columns.find((span) => span.min <= sourceCol && span.max >= sourceCol);
  const out: ColumnSpan[] = [];
  const shifted = (span: ColumnSpan, min: number, max: number): void => {
    if (min + count <= last) {
      out.push({ ...span, min: min + count, max: Math.min(max + count, last) });
    }
  };
  const inserted = source ? insertedSpan(source, at, at + count - 1) : undefined;
  let placed = false;
  const place = (): void => {
    if (!placed && inserted) {
      out.push(inserted);
    }
    placed = true;
  };
  for (const span of columns) {
    if (span.max < at) {
      out.push(span);
    } else if (span.min >= at) {
      place();
      shifted(span, span.min, span.max);
    } else {
      out.push({ ...span, max: at - 1 });
      place();
      shifted(span, at, span.max);
    }
  }
  place();
  return joinSpans(out);
}

function deleteColumns(columns: readonly ColumnSpan[], at: number, count: number): ColumnSpan[] {
  const bandEnd = at + count - 1;
  const out: ColumnSpan[] = [];
  for (const span of columns) {
    if (span.max < at) {
      out.push(span);
    } else if (span.min > bandEnd) {
      out.push({ ...span, min: span.min - count, max: span.max - count });
    } else {
      const left = span.min < at ? at - span.min : 0;
      const right = span.max > bandEnd ? span.max - bandEnd : 0;
      if (left + right > 0) {
        const min = span.min < at ? span.min : at;
        out.push({ ...span, min, max: min + left + right - 1 });
      }
    }
  }
  return joinSpans(out);
}

/** Whether two span lists hold the same spans. */
function sameSpans(a: readonly ColumnSpan[], b: readonly ColumnSpan[]): boolean {
  return a.length === b.length && a.every((span, i) => span.min === b[i]!.min && span.max === b[i]!.max && sameSpanProps(span, b[i]!));
}

const columnsAdjuster: ReferenceAdjuster = {
  id: "columns",
  adjust(model, edit) {
    if (editAxis(edit) !== "col") {
      return;
    }
    updateSheet(model, edit.sheetId, (sheet) => {
      if (sheet.columns.length === 0) {
        return sheet;
      }
      const columns = isInsert(edit) ? insertColumns(sheet.columns, edit.at, edit.count) : deleteColumns(sheet.columns, edit.at, edit.count);
      return sameSpans(columns, sheet.columns) ? sheet : { ...sheet, columns };
    });
  },
};

// --- 7. retained -----------------------------------------------------------------------------

/** The mapper a structural edit moves retained text with (see
 * retainedReferences.ts). Sheet names in retained text are source-time:
 * `formula` resolves each through resolveRetainedSheetName (the host
 * selecting a duplicated sheet's override); `range` and `index` take model
 * names, compared case-insensitively. An index in a deleted band becomes the
 * band's start; one pushed past the end, the last index. */
export function structuralMapper(model: SpreadsheetModel, edit: StructuralEdit): RetainedReferenceMapper {
  const edited = model.sheets.find((sheet) => sheet.id === edit.sheetId)?.name.toUpperCase();
  const onEdited = (sheetName: string): boolean => sheetName.toUpperCase() === edited;
  const axis = editAxis(edit);
  const parts = new LazyParts(model.opaque);
  return {
    range: (sheetName, range) => (onEdited(sheetName) ? adjustRange(range, edit) : range),
    index: (sheetName, along, index) => {
      if (along !== axis || !onEdited(sheetName)) {
        return index;
      }
      return adjustIndex(index, edit) ?? (isInsert(edit) ? lastIndex(axis) : edit.at);
    },
    formula: (canonical, hostSheetId) =>
      adjustFormulaReferences(canonical, edit, (sheet) =>
        sheet === undefined ? hostSheetId === edit.sheetId : resolveRetainedSheetName(model, parts, sheet, hostSheetId || undefined)?.id === edit.sheetId,
      ),
  };
}

const retainedAdjuster: ReferenceAdjuster = {
  id: "retained",
  adjust(model, edit) {
    if (model.opaque.length === 0) {
      return;
    }
    rewriteRetainedReferences(model, structuralMapper(model, edit));
  },
};

// --- 8. autofilter ---------------------------------------------------------------------------
// SPEC-128 "Structural edits". `autoFilter.range`, `autoFilter.sort.ref`,
// `Sheet.sortState.ref` follow adjustRange and every SheetSortKey.col
// adjustIndex. A FilterColumn's colId is relative — column range.left +
// colId — so a column edit moves its absolute column by adjustIndex and
// re-bases it on the moved range's left: colIds at or right of an inserted
// column shift by `count`; a deleted column drops its FilterColumn and its
// sort keys. A sort record left with no keys once it changes (its keys'
// columns deleted, or a keyless record whose source no longer fits) is
// removed. Deleting the header row (range.top in the band), or every column
// of the range, removes the autofilter; the rows it hid stay hidden as the
// user's (`hidden`, which Unhide clears — Excel's one flag). A changed
// column layout of a range holding an ods `odsUnmappedFilter` drops that
// verbatim filter (its field numbers no longer fit); the criteria range such
// a filter names (`table:condition-source-range-address`, on any sheet —
// SPEC-124's retained row, retired) moves through structuralMapper first,
// adjustOdsFilterCriteriaRanges's rewrite of every sheet's. Every object that
// changes is replaced — never mutated — without its `sourceXml`; one that
// does not keeps its identity. Moving or removing an autofilter sets
// `structureDirty` (the workbook part carries `_xlnm._FilterDatabase`).

/** A sort record after `edit`: the same object when nothing moves, null
 * when it no longer exists. */
export function adjustSortState(sort: SortState, edit: StructuralEdit): SortState | null {
  const ref = adjustRange(sort.ref, edit);
  if (ref === null) {
    return null;
  }
  let keys = sort.keys;
  if (editAxis(edit) === "col") {
    const moved: SheetSortKey[] = [];
    for (const key of sort.keys) {
      const col = adjustIndex(key.col, edit);
      if (col !== null) {
        moved.push(col === key.col ? key : { ...key, col });
      }
    }
    if (moved.length !== keys.length || moved.some((key, i) => key !== keys[i])) {
      keys = moved;
    }
  }
  if (ref === sort.ref && keys === sort.keys) {
    return sort;
  }
  if (keys.length === 0) {
    return null;
  }
  const { sourceXml: _source, ...rest } = sort;
  return { ...rest, ref, keys };
}

/** An autofilter after `edit`: the same object when nothing moves, null when
 * it no longer exists. */
export function adjustAutoFilter(filter: AutoFilter, edit: StructuralEdit): AutoFilter | null {
  const columnEdit = editAxis(edit) === "col";
  const { top } = filter.range;
  if (!columnEdit && !isInsert(edit) && top >= edit.at && top < edit.at + edit.count) {
    return null;
  }
  const range = adjustRange(filter.range, edit);
  if (range === null) {
    return null;
  }
  let columns = filter.columns;
  if (columnEdit) {
    const moved: FilterColumn[] = [];
    for (const column of filter.columns) {
      const col = adjustIndex(filter.range.left + column.colId, edit);
      if (col !== null) {
        const colId = col - range.left;
        moved.push(colId === column.colId ? column : { ...column, colId });
      }
    }
    if (moved.length !== columns.length || moved.some((column, i) => column !== columns[i])) {
      columns = moved;
    }
  }
  const sort = filter.sort === undefined ? undefined : (adjustSortState(filter.sort, edit) ?? undefined);
  if (range === filter.range && columns === filter.columns && sort === filter.sort) {
    return filter;
  }
  const { sourceXml: _source, sort: _sort, odsUnmappedFilter, ...rest } = filter;
  const next: AutoFilter = { ...rest, range, columns };
  if (sort !== undefined) {
    next.sort = sort;
  }
  const reshaped = columnEdit && range.right - range.left !== filter.range.right - filter.range.left;
  if (odsUnmappedFilter !== undefined && !reshaped) {
    next.odsUnmappedFilter = odsUnmappedFilter;
  }
  return next;
}

/** The rows with every `filterHidden` row made the user's `hidden` — the
 * same array when there is none. */
function filterHiddenToHidden(rows: Row[]): Row[] {
  if (!rows.some((row) => row.filterHidden)) {
    return rows;
  }
  return rows.map((row) => {
    if (!row.filterHidden) {
      return row;
    }
    const { filterHidden: _filtered, ...rest } = row;
    return { ...rest, hidden: true };
  });
}

const autofilterAdjuster: ReferenceAdjuster = {
  id: "autofilter",
  adjust(model, edit) {
    adjustOdsFilterCriteriaRanges(model, structuralMapper(model, edit));
    let structureMoved = false;
    updateSheet(model, edit.sheetId, (sheet) => {
      const was = sheet.autoFilter;
      const autoFilter = was === undefined ? undefined : (adjustAutoFilter(was, edit) ?? undefined);
      const sortState = sheet.sortState === undefined ? undefined : (adjustSortState(sheet.sortState, edit) ?? undefined);
      if (autoFilter === was && sortState === sheet.sortState) {
        return sheet;
      }
      const next: Sheet = { ...sheet };
      if (autoFilter !== undefined) {
        next.autoFilter = autoFilter;
      } else if (was !== undefined) {
        delete next.autoFilter;
        next.rows = filterHiddenToHidden(sheet.rows);
      }
      if (sortState !== undefined) {
        next.sortState = sortState;
      } else {
        delete next.sortState;
      }
      structureMoved ||= was !== undefined && autoFilter?.range !== was.range;
      return next;
    });
    if (structureMoved) {
      model.structureDirty = true;
    }
  },
};

// --- 9. conditionalFormats ---------------------------------------------------------------------
// SPEC-129 "Structural edits". On the edited sheet each format's ranges are
// mapped with adjustRange in source order (a range fully gone is dropped —
// the first one still standing leads; a format left with none is removed);
// its cellIs/expression formulas are re-anchored with adjustAnchoredFormula
// (origin = the pre-edit ranges[0] top-left, firstRange = the first range
// still standing, in pre-edit coordinates) and every other formula it
// carries (cfFormulas.ts: value objects, a preserved xlsx rule's texts) goes
// through adjustFormula. On every other sheet each formula of each format
// goes through adjustFormula with that sheet as host — only references
// naming the edited sheet move; ranges never do. A touched format is
// replaced without its sourceXml, an untouched one keeps its identity;
// priorities are renumbered when a format went (renumberPriorities), and a
// sheet whose formats changed is marked dirty, as the formulas adjuster
// marks a sheet whose cell formulas changed: the writers re-serialise a
// dirty sheet's formats from the model (SPEC-124's retained rows for them
// are retired).

/** `sheet`'s formats after `edit`: undefined when none changes. */
function adjustSheetFormats(model: SpreadsheetModel, sheet: Sheet, edit: StructuralEdit, nameOf: (sheetId: string) => string): ConditionalFormat[] | undefined {
  const formats = sheetFormats(sheet);
  const host = { format: model.format, sheetName: sheet.name };
  const plain = (formula: string): string => adjustFormula(formula, sheet.id, edit, nameOf);
  if (sheet.id !== edit.sheetId) {
    const mapped = mapFormatsFormulas(formats, plain, host);
    return mapped === formats ? undefined : mapped;
  }
  let changed = false;
  const next: ConditionalFormat[] = [];
  for (const format of formats) {
    let firstStanding: GridRange | undefined;
    const ranges: GridRange[] = [];
    for (const range of format.ranges) {
      const adjusted = adjustRange(range, edit);
      if (adjusted !== null) {
        firstStanding ??= range;
        ranges.push(adjusted);
      }
    }
    if (firstStanding === undefined) {
      changed = true; // the format is removed
      continue;
    }
    const origin = formatOrigin(format);
    const first = firstStanding;
    const anchored = (formula: string): string => adjustAnchoredFormula(formula, origin, first, sheet.id, edit, nameOf);
    const mapped = mapFormatFormulas(format, { relative: anchored, absolute: plain }, host);
    const moved = ranges.length !== format.ranges.length || ranges.some((range, i) => range !== format.ranges[i]);
    if (!moved && mapped === format) {
      next.push(format);
      continue;
    }
    changed = true;
    next.push(moved ? withoutFormatSource({ ...mapped, ranges }) : mapped);
  }
  // Surviving formats keep their rules' priorities: renumbering changes
  // something only when a format went.
  return changed ? renumberPriorities(next) : undefined;
}

const conditionalFormatsAdjuster: ReferenceAdjuster = {
  id: "conditionalFormats",
  adjust(model, edit) {
    const names = new Map(model.sheets.map((sheet) => [sheet.id, sheet.name] as const));
    const nameOf = (id: string): string => names.get(id) ?? "";
    let changed = false;
    const sheets = model.sheets.map((sheet) => {
      const formats = adjustSheetFormats(model, sheet, edit, nameOf);
      if (formats === undefined) {
        return sheet;
      }
      changed = true;
      return { ...sheet, conditionalFormats: formats, dirty: true as const };
    });
    if (changed) {
      model.sheets = sheets;
    }
  },
};

// --- 10. validations ---------------------------------------------------------------------------
// SPEC-130 "Structural edits". On the edited sheet each rule's ranges are
// mapped with adjustRange in source order (a range fully gone is dropped —
// the first one still standing leads; a rule left with none is removed);
// formula1/formula2, stored relative to the rule's origin, go through
// adjustAnchoredFormula (origin = the pre-edit ranges[0] top-left,
// firstRange = the first range still standing, in pre-edit coordinates), so
// deleting the origin's row or column re-anchors the rule instead of
// producing #REF!. On every other sheet each rule's formulas go through
// adjustFormula with that sheet as host — only references naming the edited
// sheet move; ranges never do. A touched rule is replaced without its
// sourceXml, its x14 placement re-derived (withX14Placement); an untouched
// one keeps its identity, as does a sheet stored before SPEC-130 (no
// `validations` field). A sheet whose rules changed is marked dirty, as the
// formulas adjuster marks a sheet whose cell formulas changed: the writers
// re-serialise a dirty sheet's rules from the model (SPEC-124's retained
// rows for them are retired).

/** `sheet`'s rules after `edit`: undefined when none changes. */
function adjustSheetValidations(model: SpreadsheetModel, sheet: Sheet, edit: StructuralEdit, nameOf: (sheetId: string) => string): DataValidation[] | undefined {
  const rules = sheetValidations(sheet);
  if (sheet.id !== edit.sheetId) {
    const mapped = mapValidationsFormulas(rules, (formula) => adjustFormula(formula, sheet.id, edit, nameOf), model.format, sheet.name);
    return mapped === rules ? undefined : mapped;
  }
  let changed = false;
  const next: DataValidation[] = [];
  for (const dv of rules) {
    let firstStanding: GridRange | undefined;
    const ranges: GridRange[] = [];
    for (const range of dv.ranges) {
      const adjusted = adjustRange(range, edit);
      if (adjusted !== null) {
        firstStanding ??= range;
        ranges.push(adjusted);
      }
    }
    if (firstStanding === undefined) {
      changed = true; // the rule is removed
      continue;
    }
    const origin = validationOrigin(dv);
    const first = firstStanding;
    const mapped = mapValidationFormulas(dv, (formula) => adjustAnchoredFormula(formula, origin, first, sheet.id, edit, nameOf));
    const moved = ranges.length !== dv.ranges.length || ranges.some((range, i) => range !== dv.ranges[i]);
    if (!moved && mapped === dv) {
      next.push(dv);
      continue;
    }
    changed = true;
    next.push(withX14Placement(withoutValidationSource(moved ? { ...mapped, ranges } : mapped), model.format, sheet.name));
  }
  return changed ? next : undefined;
}

const validationsAdjuster: ReferenceAdjuster = {
  id: "validations",
  adjust(model, edit) {
    const names = new Map(model.sheets.map((sheet) => [sheet.id, sheet.name] as const));
    const nameOf = (id: string): string => names.get(id) ?? "";
    let changed = false;
    const sheets = model.sheets.map((sheet) => {
      const validations = adjustSheetValidations(model, sheet, edit, nameOf);
      if (validations === undefined) {
        return sheet;
      }
      changed = true;
      return { ...sheet, validations, dirty: true as const };
    });
    if (changed) {
      model.sheets = sheets;
    }
  },
};

// --- 11. drawings ------------------------------------------------------------------------------
// SPEC-131 "Anchors follow cells" (drawings.ts is the normative source). The
// edited sheet's drawings follow per the anchor-follow table — a drawing that
// does not move keeps its identity, the array its identity when none moves.
// SPEC-124's retained rows for drawing anchors (xlsx `xdr:from`/`xdr:to`,
// ods `draw:*@table:end-cell-address`) are retired: the model moves them and
// the writers re-emit a touched drawing's anchor from it.

const drawingsAdjuster: ReferenceAdjuster = {
  id: "drawings",
  adjust(model, edit) {
    adjustDrawingsForEdit(model, edit);
  },
};

// --- 12. charts --------------------------------------------------------------------------------
// SPEC-132 "Reference adjustment" (charts.ts is the normative source). Every
// modelled chart's references, on every sheet (a chart may plot another
// sheet's data), go through adjustFormula; the chart drawing itself is
// moved by `drawings` above. SPEC-124's retained rows for chart parts now
// cover only the parts this spec does not model (out of catalogue,
// chartex), so nothing is adjusted twice.

const chartsAdjuster: ReferenceAdjuster = {
  id: "charts",
  adjust: adjustChartsForEdit,
};

// --- the registry ----------------------------------------------------------------------------

export const REFERENCE_ADJUSTERS: ReferenceAdjuster[] = [
  rowsAdjuster,
  formulasAdjuster,
  mergesAdjuster,
  frozenAdjuster,
  selectionAdjuster,
  columnsAdjuster,
  retainedAdjuster,
  autofilterAdjuster,
  conditionalFormatsAdjuster,
  validationsAdjuster,
  drawingsAdjuster,
  chartsAdjuster,
  // SPEC-133 (nameTableAdjusters.ts); called through arrows, as the modules import each other.
  { id: "names", adjust: (model, edit) => adjustNamesForEdit(model, edit) },
  { id: "tables", adjust: (model, edit) => adjustTablesForEdit(model, edit) },
  // SPEC-135 (annotations.ts): notes and threads, then hyperlinks and every place location.
  { id: "notes", adjust: (model, edit) => adjustNotesForEdit(model, edit) },
  { id: "hyperlinks", adjust: (model, edit) => adjustHyperlinksForEdit(model, edit) },
  // SPEC-136 (pageSetup.ts): the edited sheet's manual page breaks. Print areas and titles have none — `names` moves them.
  { id: "pageBreaks", adjust: (model, edit) => adjustPageBreaksForEdit(model, edit) },
];

/** One structural edit (see the module header). */
export function applyStructuralEdit(model: SpreadsheetModel, edit: StructuralEdit): StructuralResult {
  const refusal = structuralRefusal(model, edit);
  if (refusal !== null) {
    return { ok: false, message: refusal };
  }
  for (const adjuster of REFERENCE_ADJUSTERS) {
    adjuster.adjust(model, edit);
  }
  updateSheet(model, edit.sheetId, (sheet) => (sheet.dirty ? sheet : { ...sheet, dirty: true }));
  return { ok: true };
}
