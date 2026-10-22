// Cut and move (SPEC-125 "Cut and move") — THIS FILE is the normative source
// for RangeMove, repointFormula, moveMapper, the MOVE_ADJUSTERS registry and
// moveRange: what pasting a cut (and a drag-move) does to the workbook.
//
// | A reference anywhere in the workbook, moved formulas included | becomes |
// | entirely inside the source (cell or range) | moved by the offset, onto the destination sheet (prefix added, changed or dropped to match the host) |
// | range partly overlapping the source | unchanged |
// | entirely inside the destination and not in the source | `#REF!` (its cells were overwritten) |
// | whole-row/column, 3-D, external, anything else | unchanged |
// Moved formulas are otherwise NOT shifted (Excel): a moved formula keeps
// pointing at the cells it pointed at — across sheets its references to its
// old sheet gain that sheet's prefix.
//
// moveRange checks the refusals before anything changes — a sheet that is
// not a worksheet, a destination past the grid's edge, more than
// MAX_COPY_CELLS positions, a merge or an array range crossing the source's
// or the destination's edge (one entirely inside the source moves with it),
// a locked cell of a protected sheet in either, part of a table (SPEC-133) —
// then runs every adjuster in registry order and marks both sheets dirty.
// Source and destination may
// overlap (moving a block down one row): the move reads the source before it
// writes. Each adjuster replaces what it changes copy-on-write and keeps the
// identity of everything else.
//
// | Order | id | Moves |
// | 1 | cells | values, formulas, styles (each position's effective style), runs, preserved content and array anchors; the source left as unstyled blanks, the destination overwritten |
// | 2 | formulas | repointFormula over every sheet's formulas (not formulaUnparsed) |
// | 3 | merges | source merges move, destination merges are replaced |
// | 4 | retained | SPEC-124's rewriteRetainedReferences with moveMapper |
// | 5 | autofilter | SPEC-128 "Moves": an autofilter or worksheet sort record entirely inside the source moves by the offset onto the destination sheet (see "5. autofilter") |
// | 6 | conditionalFormats | SPEC-129 "Moves": formats travel with their cells — a format wholly inside the source moves, the part of one partly inside becomes a new format at the destination, the destination's previous rules go, every rule formula re-pointed (see "6. conditionalFormats") |
// | 7 | validations | SPEC-130 "Moves": data-validation rules travel with their cells by the same rule — a rule wholly inside the source moves, the part of one partly inside becomes a new rule at the destination, the destination's previous rules go, every rule formula re-pointed (see "7. validations") |
// | 8 | charts | SPEC-132 "Reference adjustment": repointFormula over every modelled chart's refs on every sheet (`refsDirty` where one changed), then a refresh of each refreshable chart whose refs changed or read either sheet; the chart drawing stays put (charts.ts moveChartsForMove) |
// | 9 | names | SPEC-133: repointFormula over every DefinedName's formula, host = its scope sheet, else the source sheet — a name follows its cells; its relative parts are offsets and stay (nameTableAdjusters.ts) |
// | 10 | tables | SPEC-133 "Tables and moves": a table wholly inside the source moves onto the destination sheet, one wholly inside the destination is overwritten (deleted, its references #REF!), every table's calculated/totals formulas re-pointed (nameTableAdjusters.ts); the refusal pass refuses a partial table (PARTIAL_TABLE_MOVE_MESSAGE) |
// Later specs that register a SPEC-124 ReferenceAdjuster for something that
// travels with cells or points at them — notes and hyperlinks
// (135), pivots (139) — register a MoveAdjuster with the same id here, with
// its row in move.spec.ts's registry contract; drawings (131) and page
// breaks (136) deliberately do not.
import { subtractRanges, type CellRef, type GridRange } from "./address";
import { mapFormatFormulas, withoutFormatSource } from "./cfFormulas";
import { moveChartsForMove } from "./charts";
import { forEachPosition, PASTE_MERGE_MESSAGE, PASTE_PAST_EDGE_MESSAGE, TOO_MANY_TO_CHANGE_MESSAGE } from "./clipboard";
import { firstLockedCell, isEmptyRow, PROTECTED_CELL_MESSAGE } from "./cells";
import { MAX_COPY_CELLS } from "./clipboardPayload";
import { formatOrigin, mintCfId, reanchorRule, renumberPriorities, sheetFormats, type CfRule, type ConditionalFormat } from "./conditionalFormats";
import { applyAutoFilter, type AutoFilter } from "./filter";
import { moveNamesForMove, moveTablesForMove, PARTIAL_TABLE_MOVE_MESSAGE, partialTableMove } from "./nameTableAdjusters";
import { rewriteReferences, type RefPart, type ReferenceNode } from "./formula/references";
import { rangesIntersect } from "./merges";
import { MAX_COLS, MAX_ROWS, type Cell, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { workbookFormatContext } from "./numberFormat/render";
import { adjustOdsFilterCriteriaRanges } from "./ods/retainedReferences";
import { rewriteRetainedReferences, type RetainedReferenceMapper } from "./retainedReferences";
import { contains, inheritedStyleId, SheetEdit } from "./sheetEdit";
import type { SortState } from "./sort";
import { resolveRetainedSheetName } from "./sourceNames";
import { NOT_A_GRID_MESSAGE, PARTIAL_ARRAY_MESSAGE } from "./structure";
import {
  mapValidationFormulas,
  mintValidationId,
  reanchorValidation,
  sheetValidations,
  validationOrigin,
  withoutValidationSource,
  withX14Placement,
  type DataValidation,
} from "./validation";
import { LazyParts } from "./xlsx/package";

export interface RangeMove {
  fromSheetId: string;
  from: GridRange;
  toSheetId: string;
  toTop: number;
  toLeft: number;
}

export interface MoveAdjuster {
  id: string;
  move(model: SpreadsheetModel, move: RangeMove): void;
}

export type MoveResult = { ok: true } | { ok: false; message: string };

/** The rectangle a move writes. */
export function moveDestination(move: RangeMove): GridRange {
  return {
    top: move.toTop,
    left: move.toLeft,
    bottom: move.toTop + move.from.bottom - move.from.top,
    right: move.toLeft + move.from.right - move.from.left,
  };
}

const offsetOf = (move: RangeMove): { dRow: number; dCol: number } => ({ dRow: move.toTop - move.from.top, dCol: move.toLeft - move.from.left });

const sheetOf = (model: SpreadsheetModel, id: string): Sheet | undefined => model.sheets.find((sheet) => sheet.id === id);

function shifted(range: GridRange, dRow: number, dCol: number): GridRange {
  return { top: range.top + dRow, left: range.left + dCol, bottom: range.bottom + dRow, right: range.right + dCol };
}

/** Whether any array formula's range on `sheet` crosses `range`'s edge,
 * ignoring those entirely inside `moving` (which travel with the block). */
function arrayCrosses(sheet: Sheet, range: GridRange, moving?: GridRange): boolean {
  return sheet.rows.some((row) =>
    row.cells.some((cell) => cell.arrayRange !== undefined && rangesIntersect(cell.arrayRange, range) && !contains(range, cell.arrayRange) && !(moving && contains(moving, cell.arrayRange))),
  );
}

/** Why `move` cannot be made, or null (see the module header). */
export function moveRefusal(model: SpreadsheetModel, move: RangeMove): string | null {
  const from = sheetOf(model, move.fromSheetId);
  const to = sheetOf(model, move.toSheetId);
  if (!from || !to || from.kind !== "grid" || to.kind !== "grid") {
    return NOT_A_GRID_MESSAGE;
  }
  const dest = moveDestination(move);
  if (dest.top < 0 || dest.left < 0 || dest.bottom >= MAX_ROWS || dest.right >= MAX_COLS) {
    return PASTE_PAST_EDGE_MESSAGE;
  }
  if ((dest.bottom - dest.top + 1) * (dest.right - dest.left + 1) > MAX_COPY_CELLS) {
    return TOO_MANY_TO_CHANGE_MESSAGE;
  }
  const same = from === to;
  if (
    from.merges.some((merge) => rangesIntersect(merge, move.from) && !contains(move.from, merge)) ||
    to.merges.some((merge) => rangesIntersect(merge, dest) && !contains(dest, merge) && !(same && contains(move.from, merge)))
  ) {
    return PASTE_MERGE_MESSAGE;
  }
  if (arrayCrosses(from, move.from) || arrayCrosses(to, dest, same ? move.from : undefined)) {
    return PARTIAL_ARRAY_MESSAGE;
  }
  if (partialTableMove(model, move)) {
    return PARTIAL_TABLE_MOVE_MESSAGE;
  }
  const fromIndex = model.sheets.indexOf(from);
  const toIndex = model.sheets.indexOf(to);
  return firstLockedCell(model, fromIndex, [move.from]) || firstLockedCell(model, toIndex, [dest]) ? PROTECTED_CELL_MESSAGE : null;
}

// --- 1. cells ----------------------------------------------------------------------------------

const cellsAdjuster: MoveAdjuster = {
  id: "cells",
  move(model, move) {
    const from = sheetOf(model, move.fromSheetId);
    const to = sheetOf(model, move.toSheetId);
    const { dRow, dCol } = offsetOf(move);
    if (!from || !to || (from === to && dRow === 0 && dCol === 0)) {
      return;
    }
    // Read the whole source first: source and destination may overlap.
    const moved = new Map<number, Cell>();
    forEachPosition(from, move.from, (row, col, cell, effective) => {
      if (!cell && effective === 0) {
        return;
      }
      const next: Cell = cell ? { ...cell, col: col + dCol } : { col: col + dCol, value: null };
      if (next.styleId === undefined && effective !== 0) {
        next.styleId = effective;
      }
      if (next.arrayRange) {
        next.arrayRange = shifted(next.arrayRange, dRow, dCol);
      }
      moved.set((row + dRow) * MAX_COLS + col + dCol, next);
    });
    const fromEdit = new SheetEdit(from);
    const toEdit = from === to ? fromEdit : new SheetEdit(to);
    // The source becomes unstyled blanks.
    forEachPosition(from, move.from, (row, col, cell) => {
      if (inheritedStyleId(from, row, col) !== 0) {
        fromEdit.set(row, col, { col, value: null, styleId: 0 });
      } else if (cell) {
        fromEdit.set(row, col, undefined);
      }
    });
    // The destination is overwritten.
    const dest = moveDestination(move);
    for (let row = dest.top; row <= dest.bottom; row += 1) {
      for (let col = dest.left; col <= dest.right; col += 1) {
        const cell = moved.get(row * MAX_COLS + col);
        const inherited = inheritedStyleId(to, row, col);
        if (cell) {
          toEdit.set(row, col, cell.styleId === undefined && inherited !== 0 ? { ...cell, styleId: 0 } : cell);
        } else {
          toEdit.set(row, col, inherited !== 0 ? { col, value: null, styleId: 0 } : undefined);
        }
      }
    }
    const fromRows = fromEdit.rows();
    const toRows = toEdit === fromEdit ? fromRows : toEdit.rows();
    if (fromRows === from.rows && toRows === to.rows) {
      return;
    }
    model.sheets = model.sheets.map((sheet) => {
      if (sheet === from && fromRows !== from.rows) {
        return { ...sheet, rows: fromRows };
      }
      if (sheet === to && to !== from && toRows !== to.rows) {
        return { ...sheet, rows: toRows };
      }
      return sheet;
    });
  },
};

// --- 2. formulas -------------------------------------------------------------------------------

/** A reference's area (whole rows and columns are null: never moved). */
function areaOf(ref: ReferenceNode): GridRange | null {
  const end = ref.end ?? ref.start;
  if (ref.start.row === undefined || ref.start.col === undefined || end.row === undefined || end.col === undefined) {
    return null;
  }
  return {
    top: Math.min(ref.start.row, end.row),
    left: Math.min(ref.start.col, end.col),
    bottom: Math.max(ref.start.row, end.row),
    right: Math.max(ref.start.col, end.col),
  };
}

function movedPart(part: RefPart, dRow: number, dCol: number): RefPart {
  return { ...part, row: part.row! + dRow, col: part.col! + dCol };
}

/** `text` (canonical, no "=") hosted on `hostSheetId` with every reference
 * re-pointed for `move` (see the module header). `movedTo`: the formula
 * itself moved there with the cells (its host was `hostSheetId`). */
export function repointFormula(text: string, hostSheetId: string, move: RangeMove, sheetNameOf: (id: string) => string, movedTo?: string): string {
  const fromName = sheetNameOf(move.fromSheetId).toUpperCase();
  const toName = sheetNameOf(move.toSheetId).toUpperCase();
  const hostName = sheetNameOf(hostSheetId);
  const newHost = movedTo ?? hostSheetId;
  const crossSheet = move.fromSheetId !== move.toSheetId;
  const { dRow, dCol } = offsetOf(move);
  const dest = moveDestination(move);
  return rewriteReferences(text, (ref) => {
    // Whole rows and columns and 3-D references have no area: never moved.
    const area = ref.sheetEnd === undefined ? areaOf(ref) : null;
    if (area) {
      const on = (ref.sheet ?? hostName).toUpperCase();
      if (on === fromName && contains(move.from, area)) {
        const moved: ReferenceNode = { start: movedPart(ref.start, dRow, dCol) };
        if (ref.end) {
          moved.end = movedPart(ref.end, dRow, dCol);
        }
        if (!crossSheet) {
          if (ref.sheet !== undefined) {
            moved.sheet = ref.sheet;
          }
        } else if (move.toSheetId !== newHost) {
          moved.sheet = sheetNameOf(move.toSheetId);
        }
        return moved;
      }
      if (on === toName && contains(dest, area)) {
        return null;
      }
    }
    if (movedTo !== undefined && movedTo !== hostSheetId && ref.sheet === undefined) {
      // A moved formula keeps pointing at its old sheet's cells — whole rows
      // and columns included (`B:B` must not become the new sheet's B:B).
      return { ...ref, sheet: hostName };
    }
    return undefined;
  });
}

const formulasAdjuster: MoveAdjuster = {
  id: "formulas",
  move(model, move) {
    const names = new Map(model.sheets.map((sheet) => [sheet.id, sheet.name] as const));
    const nameOf = (id: string): string => names.get(id) ?? "";
    const dest = moveDestination(move);
    let sheetsChanged = false;
    const sheets = model.sheets.map((sheet) => {
      const involved = sheet.id === move.fromSheetId || sheet.id === move.toSheetId;
      let rowsChanged = false;
      const rows = sheet.rows.map((row) => {
        let cellsChanged = false;
        const cells = row.cells.map((cell) => {
          // Another sheet's formula reaches the moved cells only with a `!`.
          if (cell.formula === undefined || cell.formulaUnparsed || (!involved && !cell.formula.includes("!"))) {
            return cell;
          }
          const moved = sheet.id === move.toSheetId && row.index >= dest.top && row.index <= dest.bottom && cell.col >= dest.left && cell.col <= dest.right;
          const formula = moved ? repointFormula(cell.formula, move.fromSheetId, move, nameOf, move.toSheetId) : repointFormula(cell.formula, sheet.id, move, nameOf);
          if (formula === cell.formula) {
            return cell;
          }
          cellsChanged = true;
          return { ...cell, formula };
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

// --- 3. merges ---------------------------------------------------------------------------------

const mergesAdjuster: MoveAdjuster = {
  id: "merges",
  move(model, move) {
    const from = sheetOf(model, move.fromSheetId);
    const to = sheetOf(model, move.toSheetId);
    const { dRow, dCol } = offsetOf(move);
    if (!from || !to || (from === to && dRow === 0 && dCol === 0)) {
      return;
    }
    const dest = moveDestination(move);
    const moving = from.merges.filter((merge) => contains(move.from, merge));
    const replaced = to.merges.filter((merge) => contains(dest, merge) && !(from === to && contains(move.from, merge)));
    if (moving.length === 0 && replaced.length === 0) {
      return;
    }
    const arrived = moving.map((merge) => shifted(merge, dRow, dCol));
    model.sheets = model.sheets.map((sheet) => {
      if (sheet === from && from === to) {
        return { ...sheet, merges: [...sheet.merges.filter((merge) => !moving.includes(merge) && !replaced.includes(merge)), ...arrived] };
      }
      if (sheet === from) {
        return moving.length > 0 ? { ...sheet, merges: sheet.merges.filter((merge) => !moving.includes(merge)) } : sheet;
      }
      if (sheet === to) {
        return { ...sheet, merges: [...sheet.merges.filter((merge) => !replaced.includes(merge)), ...arrived] };
      }
      return sheet;
    });
  },
};

// --- 4. retained -------------------------------------------------------------------------------

/** The mapper a move re-points retained text with (see
 * retainedReferences.ts): on the source sheet, a range or reference entirely
 * inside the source moves with the cells — on the same sheet only, since no
 * sheet name in retained text is ever rewritten; on the destination sheet,
 * one entirely inside the destination is dropped (a range) or becomes
 * `#REF!` (a reference); partial overlaps, whole rows and columns and
 * indices (drawings, page breaks) stay. Retained formulas name sheets as the
 * source file did: each resolves through SPEC-121's
 * resolveRetainedSheetName. */
export function moveMapper(model: SpreadsheetModel, move: RangeMove): RetainedReferenceMapper {
  const fromName = sheetOf(model, move.fromSheetId)?.name.toUpperCase();
  const toName = sheetOf(model, move.toSheetId)?.name.toUpperCase();
  const same = move.fromSheetId === move.toSheetId;
  const { dRow, dCol } = offsetOf(move);
  const dest = moveDestination(move);
  const parts = new LazyParts(model.opaque);
  return {
    range: (sheetName, range) => {
      const name = sheetName.toUpperCase();
      if (name === fromName && contains(move.from, range)) {
        return same ? shifted(range, dRow, dCol) : range;
      }
      return name === toName && contains(dest, range) ? null : range;
    },
    index: (_sheetName, _axis, index) => index,
    formula: (canonical, hostSheetId) =>
      rewriteReferences(canonical, (ref) => {
        const area = ref.sheetEnd === undefined ? areaOf(ref) : null;
        const sheetId = ref.sheet === undefined ? hostSheetId || undefined : resolveRetainedSheetName(model, parts, ref.sheet, hostSheetId || undefined)?.id;
        if (!area || sheetId === undefined) {
          return undefined;
        }
        if (sheetId === move.fromSheetId && contains(move.from, area)) {
          return same ? { ...ref, start: movedPart(ref.start, dRow, dCol), ...(ref.end ? { end: movedPart(ref.end, dRow, dCol) } : {}) } : undefined;
        }
        return sheetId === move.toSheetId && contains(dest, area) ? null : undefined;
      }),
  };
}

const retainedAdjuster: MoveAdjuster = {
  id: "retained",
  move(model, move) {
    if (model.opaque.length === 0) {
      return;
    }
    rewriteRetainedReferences(model, moveMapper(model, move));
  },
};

// --- 5. autofilter -----------------------------------------------------------------------------
// SPEC-128 "Moves". An autofilter whose range lies entirely inside the source
// moves by the offset onto the destination sheet — its colIds are relative
// and stay; its sort record's ref and key columns shift with it — replacing
// that sheet's autofilter. Row properties never move with cells: the old data
// rows' `filterHidden` (and a replaced autofilter's) are cleared and the
// filter is reapplied at its new place (applyAutoFilter, the workbook's
// format context in the default locale). A worksheet sort record entirely
// inside the source moves by the same rule. Anything overlapping the source
// or the destination only partly is left alone. What moves is replaced —
// never mutated — without its `sourceXml`; an ods database range keeps its
// name on its own sheet and takes the replaced one's (or none) on another.
// Moving an autofilter sets `structureDirty` (`_xlnm._FilterDatabase`). The
// criteria range an ods verbatim filter names, on any sheet, moves first
// through moveMapper (adjustOdsFilterCriteriaRanges — SPEC-124's retained
// row for it is retired).

/** A sort record moved by the offset, without its source. */
function movedSortState(sort: SortState, dRow: number, dCol: number): SortState {
  const { sourceXml: _source, ...rest } = sort;
  return { ...rest, ref: shifted(sort.ref, dRow, dCol), keys: sort.keys.map((key) => ({ ...key, col: key.col + dCol })) };
}

/** The rows with `filterHidden` cleared on an autofilter's data rows —
 * emptied rows removed, the same array when none was set. */
function withoutFilterHidden(rows: Row[], range: GridRange): Row[] {
  let changed = false;
  const out: Row[] = [];
  for (const row of rows) {
    if (!row.filterHidden || row.index <= range.top || row.index > range.bottom) {
      out.push(row);
      continue;
    }
    changed = true;
    const { filterHidden: _filtered, ...rest } = row;
    if (!isEmptyRow(rest)) {
      out.push(rest);
    }
  }
  return changed ? out : rows;
}

const autofilterAdjuster: MoveAdjuster = {
  id: "autofilter",
  move(model, move) {
    const { dRow, dCol } = offsetOf(move);
    if (move.fromSheetId === move.toSheetId && dRow === 0 && dCol === 0) {
      return;
    }
    adjustOdsFilterCriteriaRanges(model, moveMapper(model, move));
    const from = sheetOf(model, move.fromSheetId);
    const to = sheetOf(model, move.toSheetId);
    if (!from || !to) {
      return;
    }
    const filter = from.autoFilter !== undefined && contains(move.from, from.autoFilter.range) ? from.autoFilter : undefined;
    const sortState = from.sortState !== undefined && contains(move.from, from.sortState.ref) ? from.sortState : undefined;
    if (!filter && !sortState) {
      return;
    }
    let arrived: AutoFilter | undefined;
    if (filter) {
      const { sourceXml: _source, sort, odsName: _name, ...rest } = filter;
      arrived = { ...rest, range: shifted(filter.range, dRow, dCol) };
      if (sort) {
        arrived.sort = movedSortState(sort, dRow, dCol);
      }
      const odsName = from === to ? filter.odsName : to.autoFilter?.odsName;
      if (odsName !== undefined) {
        arrived.odsName = odsName;
      }
    }
    model.sheets = model.sheets.map((sheet) => {
      if (sheet !== from && sheet !== to) {
        return sheet;
      }
      const next: Sheet = { ...sheet };
      if (sheet === from) {
        if (filter) {
          delete next.autoFilter;
          next.rows = withoutFilterHidden(next.rows, filter.range);
        }
        if (sortState) {
          delete next.sortState;
        }
      }
      if (sheet === to) {
        if (arrived) {
          if (sheet !== from && sheet.autoFilter) {
            next.rows = withoutFilterHidden(next.rows, sheet.autoFilter.range);
          }
          next.autoFilter = arrived;
        }
        if (sortState) {
          next.sortState = movedSortState(sortState, dRow, dCol);
        }
      }
      return next;
    });
    if (filter) {
      model.structureDirty = true;
      applyAutoFilter(model, move.toSheetId, workbookFormatContext(model, SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE]));
    }
  },
};

// --- 6. conditionalFormats ----------------------------------------------------------------------
// SPEC-129 "Moves": formats travel with their cells, as in Excel. The source
// is read before anything is written (source and destination may overlap):
// - Each format on the source sheet loses the part of its ranges inside
//   `move.from`. A format whose every range lies inside travels whole —
//   same id, its ranges offset onto the destination sheet. Otherwise the
//   cut-out pieces become a new format on the destination (fresh format and
//   rule ids, no x14Id: the writer pairs it afresh) carrying the source's
//   rules rebased (reanchorRule) onto the pieces' origin, so each piece
//   means what it meant in place.
// - Every format on the destination sheet that did not just arrive loses the
//   destination rectangle (subtractRanges) — the overwritten cells' rules go
//   — and a format left with no range is removed. A staying format whose
//   origin changed (cut or overwritten) is re-anchored onto its new origin,
//   once, from the origin it had.
// - Priorities: on the same sheet each new rule ranks right after the rule
//   it copies; on another sheet the arriving rules (a whole format's too)
//   rank lowest, in their source order — a whole format crossing sheets also
//   drops its x14Ids, as a GUID pairs only within one worksheet. Both
//   sheets are then renumbered densely.
// - Every formula of every format on every sheet goes through
//   repointFormula (cfFormulas.ts's table), an arriving format's as a moved
//   formula hosted on the source sheet — references to moved cells follow
//   them, its unqualified references gain the source sheet's name across
//   sheets — every other format's as its own sheet's.
// Touched formats are replaced without their sourceXml; a sheet whose
// formats changed is marked dirty (moveRange marks both ends anyway).

/** The cells two ranges share, or null. */
function overlap(a: GridRange, b: GridRange): GridRange | null {
  const top = Math.max(a.top, b.top);
  const bottom = Math.min(a.bottom, b.bottom);
  const left = Math.max(a.left, b.left);
  const right = Math.min(a.right, b.right);
  return top <= bottom && left <= right ? { top, left, bottom, right } : null;
}

/** `format` on `ranges` (some of its own, or what is left of them): the
 * same format when they are its ranges, null when none is left, else
 * re-anchored from its origin onto theirs, without its sourceXml. */
function onRanges(format: ConditionalFormat, ranges: GridRange[]): ConditionalFormat | null {
  if (ranges.length === 0) {
    return null;
  }
  if (ranges.length === format.ranges.length && ranges.every((range, i) => range === format.ranges[i])) {
    return format;
  }
  const from = formatOrigin(format);
  const to = formatOrigin({ ranges });
  return withoutFormatSource({ ...format, ranges, rules: format.rules.map((rule) => reanchorRule(rule, from, to)) });
}

/** A rule without its x14 GUID. */
function withoutX14Id(rule: CfRule): CfRule {
  if (rule.x14Id === undefined) {
    return rule;
  }
  const { x14Id: _id, ...rest } = rule;
  return rest as CfRule;
}

const conditionalFormatsAdjuster: MoveAdjuster = {
  id: "conditionalFormats",
  move(model, move) {
    const from = sheetOf(model, move.fromSheetId);
    const to = sheetOf(model, move.toSheetId);
    const { dRow, dCol } = offsetOf(move);
    if (!from || !to || (from === to && dRow === 0 && dCol === 0)) {
      return;
    }
    const same = from === to;
    const dest = moveDestination(move);
    // The source's formats with the moved part cut out, and what travels.
    const staying: { format: ConditionalFormat; ranges: GridRange[]; arrived?: true }[] = [];
    const arriving: ConditionalFormat[] = [];
    for (const format of sheetFormats(from)) {
      const inside = format.ranges.flatMap((range) => overlap(range, move.from) ?? []);
      if (inside.length === 0) {
        staying.push({ format, ranges: format.ranges });
        continue;
      }
      if (format.ranges.every((range) => contains(move.from, range))) {
        const whole = withoutFormatSource({
          ...format,
          ranges: format.ranges.map((range) => shifted(range, dRow, dCol)),
          rules: same ? format.rules : format.rules.map(withoutX14Id),
        });
        if (same) {
          staying.push({ format: whole, ranges: whole.ranges, arrived: true });
        } else {
          arriving.push(whole);
        }
        continue;
      }
      staying.push({ format, ranges: subtractRanges(format.ranges, move.from) });
      const origin = formatOrigin(format);
      const pieceOrigin = formatOrigin({ ranges: inside });
      const { id: _id, ranges: _ranges, sourceXml: _source, rules, ...fields } = format;
      arriving.push({
        ...fields,
        id: mintCfId(),
        ranges: inside.map((range) => shifted(range, dRow, dCol)),
        rules: rules.map((rule) => ({
          ...withoutX14Id(reanchorRule(rule, origin, pieceOrigin)),
          id: mintCfId(),
          // Right after the rule it copies (another sheet: placed below).
          priority: same ? rule.priority + 0.5 : rule.priority,
        })),
      });
    }
    // The destination loses the overwritten rectangle — what arrived excepted.
    const overwritten = (format: ConditionalFormat, ranges: GridRange[], arrived?: true): ConditionalFormat | null =>
      arrived ? format : onRanges(format, subtractRanges(ranges, dest));
    const fromNext = same
      ? staying.flatMap(({ format, ranges, arrived }) => overwritten(format, ranges, arrived) ?? [])
      : staying.flatMap(({ format, ranges }) => onRanges(format, ranges) ?? []);
    const toNext = same ? fromNext : sheetFormats(to).flatMap((format) => overwritten(format, format.ranges) ?? []);
    // On another sheet the arrivals rank below every rule already there.
    const lowest = same ? 0 : Math.max(0, ...toNext.flatMap((format) => format.rules.map((rule) => rule.priority)));
    toNext.push(...arriving.map((format) => (lowest === 0 ? format : { ...format, rules: format.rules.map((rule) => ({ ...rule, priority: lowest + rule.priority })) })));
    const moved = new Set<ConditionalFormat>([...toNext.slice(toNext.length - arriving.length), ...staying.flatMap(({ format, arrived }) => (arrived ? [format] : []))]);
    // Every formula re-pointed; then each touched sheet renumbered.
    const names = new Map(model.sheets.map((sheet) => [sheet.id, sheet.name] as const));
    const nameOf = (id: string): string => names.get(id) ?? "";
    let changed = false;
    const sheets = model.sheets.map((sheet) => {
      const formats = sheet === to ? toNext : sheet === from ? fromNext : sheetFormats(sheet);
      const host = { format: model.format, sheetName: sheet.name };
      const own = (formula: string): string => repointFormula(formula, sheet.id, move, nameOf);
      const travelled = (formula: string): string => repointFormula(formula, move.fromSheetId, move, nameOf, move.toSheetId);
      let next = formats.map((format) => {
        const map = moved.has(format) ? travelled : own;
        return mapFormatFormulas(format, { relative: map, absolute: map }, host);
      });
      if (sheet === from || sheet === to) {
        next = renumberPriorities(next);
      }
      const before = sheetFormats(sheet);
      if (next.length === before.length && next.every((format, i) => format === before[i])) {
        return sheet;
      }
      changed = true;
      return { ...sheet, conditionalFormats: next, dirty: true as const };
    });
    if (changed) {
      model.sheets = sheets;
    }
  },
};

// --- 7. validations -----------------------------------------------------------------------------
// SPEC-130 "Moves": data-validation rules travel with their cells, as in
// Excel (a cut-paste or a drag-move; a copy-paste never touches them). The
// source is read before anything is written (source and destination may
// overlap), as for conditional formats above:
// - Each rule on the source sheet loses the part of its ranges inside
//   `move.from`. A rule whose every range lies inside travels whole — same
//   id, its ranges offset onto the destination sheet (across sheets without
//   its `odsName`: a content-validation's base cell names its sheet).
//   Otherwise the cut-out pieces become a new rule on the destination (a
//   fresh id, no `sourceXml` or `odsName`) whose formulas are the source's
//   rebased (reanchorValidation, shiftFormula) onto the pieces' origin, so
//   each piece means what it meant in place.
// - Every rule on the destination sheet that did not just arrive loses the
//   destination rectangle (subtractRanges) — the overwritten cells' rules go
//   — and a rule left with no range is removed. A staying rule whose origin
//   changed (cut or overwritten) is re-anchored onto its new origin, once,
//   from the origin it had. An arriving rule is never merged into an
//   identical one at the destination.
// - Every formula of every rule on every sheet then goes through
//   repointFormula exactly as cell formulas do — an arriving rule's as a
//   moved formula hosted on the source sheet (references to moved cells
//   follow them; across sheets its unqualified references gain the source
//   sheet's name, so they keep meaning the same cells), every other rule's
//   as its own sheet's.
// Touched rules are replaced without their sourceXml, their x14 placement
// re-derived for the sheet they end on (withX14Placement — a moved rule that
// now names its old sheet is an x14 rule in an .xlsx); a sheet whose rules
// changed is marked dirty (moveRange marks both ends anyway).

/** `dv` on `ranges` (some of its own, or what is left of them): the same
 * rule when they are its ranges, null when none is left, else re-anchored
 * from its origin onto theirs, without its sourceXml. */
function validationOnRanges(dv: DataValidation, ranges: GridRange[]): DataValidation | null {
  if (ranges.length === 0) {
    return null;
  }
  if (ranges.length === dv.ranges.length && ranges.every((range, i) => range === dv.ranges[i])) {
    return dv;
  }
  return withoutValidationSource(reanchorValidation({ ...dv, ranges }, validationOrigin(dv), validationOrigin({ ranges })));
}

const validationsAdjuster: MoveAdjuster = {
  id: "validations",
  move(model, move) {
    const from = sheetOf(model, move.fromSheetId);
    const to = sheetOf(model, move.toSheetId);
    const { dRow, dCol } = offsetOf(move);
    if (!from || !to || (from === to && dRow === 0 && dCol === 0)) {
      return;
    }
    const same = from === to;
    const dest = moveDestination(move);
    // The source's rules with the moved part cut out, and what travels.
    const staying: { dv: DataValidation; ranges: GridRange[]; arrived?: true }[] = [];
    const arriving: DataValidation[] = [];
    for (const dv of sheetValidations(from)) {
      const inside = dv.ranges.flatMap((range) => overlap(range, move.from) ?? []);
      if (inside.length === 0) {
        staying.push({ dv, ranges: dv.ranges });
        continue;
      }
      if (dv.ranges.every((range) => contains(move.from, range))) {
        const whole: DataValidation = { ...withoutValidationSource(dv), ranges: dv.ranges.map((range) => shifted(range, dRow, dCol)) };
        if (same) {
          staying.push({ dv: whole, ranges: whole.ranges, arrived: true });
        } else {
          const { odsName: _name, ...crossing } = whole;
          arriving.push(crossing);
        }
        continue;
      }
      staying.push({ dv, ranges: subtractRanges(dv.ranges, move.from) });
      const { id: _id, ranges: _ranges, sourceXml: _source, odsName: _name, ...settings } = dv;
      arriving.push({
        id: mintValidationId(),
        ranges: inside.map((range) => shifted(range, dRow, dCol)),
        ...reanchorValidation(settings, validationOrigin(dv), validationOrigin({ ranges: inside })),
      });
    }
    // The destination loses the overwritten rectangle — what arrived excepted.
    const overwritten = (dv: DataValidation, ranges: GridRange[], arrived?: true): DataValidation | null =>
      arrived ? dv : validationOnRanges(dv, subtractRanges(ranges, dest));
    const fromNext = same
      ? staying.flatMap(({ dv, ranges, arrived }) => overwritten(dv, ranges, arrived) ?? [])
      : staying.flatMap(({ dv, ranges }) => validationOnRanges(dv, ranges) ?? []);
    const toNext = same ? fromNext : sheetValidations(to).flatMap((dv) => overwritten(dv, dv.ranges) ?? []);
    toNext.push(...arriving);
    const moved = new Set<DataValidation>([...arriving, ...staying.flatMap(({ dv, arrived }) => (arrived ? [dv] : []))]);
    // Every formula re-pointed; a touched rule placed for its sheet.
    const names = new Map(model.sheets.map((sheet) => [sheet.id, sheet.name] as const));
    const nameOf = (id: string): string => names.get(id) ?? "";
    let changed = false;
    const sheets = model.sheets.map((sheet) => {
      const before = sheetValidations(sheet);
      const rules = sheet === to ? toNext : sheet === from ? fromNext : before;
      const own = (formula: string): string => repointFormula(formula, sheet.id, move, nameOf);
      const travelled = (formula: string): string => repointFormula(formula, move.fromSheetId, move, nameOf, move.toSheetId);
      const untouched = new Set(before);
      const next = rules.map((dv) => {
        const mapped = mapValidationFormulas(dv, moved.has(dv) ? travelled : own);
        return mapped === dv && untouched.has(dv) ? dv : withX14Placement(mapped, model.format, sheet.name);
      });
      if (next.length === before.length && next.every((dv, i) => dv === before[i])) {
        return sheet;
      }
      changed = true;
      return { ...sheet, validations: next, dirty: true as const };
    });
    if (changed) {
      model.sheets = sheets;
    }
  },
};

// --- the registry ------------------------------------------------------------------------------

export const MOVE_ADJUSTERS: MoveAdjuster[] = [
  cellsAdjuster,
  formulasAdjuster,
  mergesAdjuster,
  retainedAdjuster,
  autofilterAdjuster,
  conditionalFormatsAdjuster,
  validationsAdjuster,
  { id: "charts", move: moveChartsForMove },
  // SPEC-133 (nameTableAdjusters.ts); called through arrows, as the modules import each other.
  { id: "names", move: (model, move) => moveNamesForMove(model, move) },
  { id: "tables", move: (model, move) => moveTablesForMove(model, move) },
];

/** One move (see the module header). */
export function moveRange(model: SpreadsheetModel, move: RangeMove): MoveResult {
  const refusal = moveRefusal(model, move);
  if (refusal !== null) {
    return { ok: false, message: refusal };
  }
  const { dRow, dCol } = offsetOf(move);
  if (move.fromSheetId === move.toSheetId && dRow === 0 && dCol === 0) {
    return { ok: true };
  }
  for (const adjuster of MOVE_ADJUSTERS) {
    adjuster.move(model, move);
  }
  model.sheets = model.sheets.map((sheet) => ((sheet.id === move.fromSheetId || sheet.id === move.toSheetId) && !sheet.dirty ? { ...sheet, dirty: true as const } : sheet));
  return { ok: true };
}

/** The top-left a block lands on when moved from `anchor` to `at`, kept
 * inside the grid (a drag-move's snapped destination). */
export function snappedMove(from: GridRange, anchor: CellRef, at: CellRef): { toTop: number; toLeft: number } {
  const rows = from.bottom - from.top;
  const cols = from.right - from.left;
  return {
    toTop: Math.min(Math.max(0, from.top + at.row - anchor.row), MAX_ROWS - 1 - rows),
    toLeft: Math.min(Math.max(0, from.left + at.col - anchor.col), MAX_COLS - 1 - cols),
  };
}
