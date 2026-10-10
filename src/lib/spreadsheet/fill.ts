// The fill handle, double-click fill and fill down/right (SPEC-125 "Series
// detection", "Double-click fill") — THIS FILE is the normative source for
// fillRange, the double-click extent and Ctrl/Cmd+D / Ctrl/Cmd+R. The series
// themselves are series.ts's detectSeries.
//
// - fillRange(model, sheetId, seed, target, mode): `target` is the range the
//   drag ends on — the seed extended along one axis (down, up, right or
//   left), or shrunk back inside it. Extended: every seed line (each column
//   of a vertical fill, each row of a horizontal one) continues into the new
//   positions, filling up or left extending it backwards; each filled cell
//   takes the style of its seed position (k mod n), and the seed's merges
//   are repeated wherever a whole copy fits. Shrunk: the contents (not the
//   styles) of the cells left behind are cleared.
// - Modes: "auto" — each line's series where a plain drag fills one
//   (detectSeries' byDefault), else a copy; "toggle" — the other one (Ctrl on
//   Windows/Linux, Alt/Option everywhere, held at release); the chip's "copy"
//   (the pattern repeats), "series" (the detected series; a non-series seed
//   copies), "formats" (styles only) and "no-formats" (the auto values,
//   destination styles kept). Formulas are shifted by their distance in
//   every mode.
// - Refusals, checked before anything changes: more than MAX_COPY_CELLS
//   positions to write, a merge or an array range crossing the written
//   band's edge, a locked cell of a protected sheet.
// - Double-click fill: with the selection's rows top..bottom and columns
//   left..right, the guide column is left−1 if the cell at (bottom+1,
//   left−1) is non-empty, else right+1 if (bottom+1, right+1) is; none →
//   nothing. The fill ends at the last row of the guide column's contiguous
//   non-empty run from bottom+1, or the row before the first non-empty cell
//   below the selection in any of its own columns, whichever comes first.
// - Fill down / right copy each column's top cell (each row's left cell)
//   into the rest of the range — formulas shifted, styles copied, never a
//   series; a one-row (one-column) range copies from the row above (the
//   column to the left). Whole rows and columns are clipped to the used
//   range first.
import { getCell, usedRange, type GridRange } from "./address";
import { arrayRangeTouched, clearContents, firstLockedCell, PROTECTED_CELL_MESSAGE } from "./cells";
import { forEachPosition, PASTE_MERGE_MESSAGE, TOO_MANY_TO_CHANGE_MESSAGE } from "./clipboard";
import { MAX_COPY_CELLS } from "./clipboardPayload";
import { shiftFormula } from "./formula/references";
import { buildMergeIndex, mergeAt, rangesIntersect } from "./merges";
import { MAX_COLS, MAX_ROWS, type Cell, type Sheet, type SpreadsheetModel, type TextRun } from "./model";
import { formatCodeOf } from "./numberFormat/builtin";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES, type SpreadsheetLocale } from "./numberFormat/locale";
import { firstRowAtOrAfter } from "./restyle";
import { detectSeries, type SeriesSeed, type SeriesValue } from "./series";
import { contains, SheetEdit, type CellChange } from "./sheetEdit";
import { NOT_A_GRID_MESSAGE, PARTIAL_ARRAY_MESSAGE } from "./structure";
import { cellStyle } from "./styles";

export { detectSeries, type Series, type SeriesSeed, type SeriesKind } from "./series";

export type FillMode = "auto" | "toggle" | "copy" | "series" | "formats" | "no-formats";
/** What a fill did — the chip shows it pressed. */
export type AppliedFill = "copy" | "series" | "formats" | "no-formats";

export type FillResult = { ok: true; applied: AppliedFill; range: GridRange; cleared?: true } | { ok: false; message: string };

const mod = (a: number, n: number): number => ((a % n) + n) % n;

/** The fill's direction from `seed` to `target`, or null when target is no
 * one-axis extension or shrink of it. */
export function fillDirection(seed: GridRange, target: GridRange): "down" | "up" | "right" | "left" | "shrink-rows" | "shrink-cols" | null {
  if (target.left === seed.left && target.right === seed.right) {
    if (target.top === seed.top && target.bottom > seed.bottom) {
      return "down";
    }
    if (target.bottom === seed.bottom && target.top < seed.top) {
      return "up";
    }
    if (target.top === seed.top && target.bottom < seed.bottom && target.bottom >= seed.top) {
      return "shrink-rows";
    }
  }
  if (target.top === seed.top && target.bottom === seed.bottom) {
    if (target.left === seed.left && target.right > seed.right) {
      return "right";
    }
    if (target.right === seed.right && target.left < seed.left) {
      return "left";
    }
    if (target.left === seed.left && target.right < seed.right && target.right >= seed.left) {
      return "shrink-cols";
    }
  }
  return null;
}

/** Why writing `band` is refused, or null (see the module header). */
export function fillRefusal(model: SpreadsheetModel, sheetIndex: number, band: GridRange): string | null {
  const sheet = model.sheets[sheetIndex];
  if (!sheet || sheet.kind !== "grid") {
    return NOT_A_GRID_MESSAGE;
  }
  if ((band.bottom - band.top + 1) * (band.right - band.left + 1) > MAX_COPY_CELLS) {
    return TOO_MANY_TO_CHANGE_MESSAGE;
  }
  if (sheet.merges.some((merge) => rangesIntersect(merge, band) && !contains(band, merge))) {
    return PASTE_MERGE_MESSAGE;
  }
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      if (cell.arrayRange && rangesIntersect(cell.arrayRange, band) && !contains(band, cell.arrayRange)) {
        return PARTIAL_ARRAY_MESSAGE;
      }
    }
  }
  return firstLockedCell(model, sheetIndex, [band]) ? PROTECTED_CELL_MESSAGE : null;
}

/** One seed position as the fill reads it. */
interface SeedCell extends SeriesSeed {
  styleId: number;
  runs?: TextRun[];
}

/** The seed positions of `range` (row-major), merged-away positions empty. */
function seedCells(model: SpreadsheetModel, sheet: Sheet, range: GridRange, locale: SpreadsheetLocale): SeedCell[][] {
  const merges = buildMergeIndex(sheet.merges);
  const grid: SeedCell[][] = [];
  forEachPosition(sheet, range, (row, col, cell, styleId) => {
    if (row - range.top >= grid.length) {
      grid.push([]);
    }
    const merge = mergeAt(merges, row, col);
    const covered = merge !== undefined && (merge.top !== row || merge.left !== col);
    const seed: SeedCell = { value: covered ? null : (cell?.value ?? null), styleId, code: formatCodeOf(cellStyle(model, styleId), locale) };
    if (!covered && cell?.formula !== undefined && !cell.formulaUnparsed) {
      seed.formula = cell.formula;
    }
    if (!covered && cell?.runs && typeof cell.value === "string") {
      seed.runs = cell.runs;
    }
    grid[grid.length - 1]!.push(seed);
  });
  return grid;
}

/** The value a copy writes at index i of a line (the pattern repeated,
 * formulas shifted by their distance). */
function copyAt(line: readonly SeedCell[], i: number, axis: "row" | "col"): SeriesValue & { runs?: TextRun[] } {
  const j = mod(i, line.length);
  const seed = line[j]!;
  const out: SeriesValue & { runs?: TextRun[] } = { value: seed.value };
  if (seed.formula !== undefined) {
    const distance = i - j;
    out.formula = shiftFormula(seed.formula, axis === "row" ? distance : 0, axis === "col" ? distance : 0);
  } else if (seed.runs) {
    out.runs = structuredClone(seed.runs);
  }
  return out;
}

/** Fills `seed` out to `target` on the sheet `sheetId` (see the module
 * header). */
export function fillRange(
  model: SpreadsheetModel,
  sheetId: string,
  seed: GridRange,
  target: GridRange,
  mode: FillMode,
  locale: SpreadsheetLocale = SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE],
): FillResult {
  const sheetIndex = model.sheets.findIndex((each) => each.id === sheetId);
  const sheet = model.sheets[sheetIndex];
  if (!sheet || sheet.kind !== "grid") {
    return { ok: false, message: NOT_A_GRID_MESSAGE };
  }
  const direction = fillDirection(seed, target);
  if (direction === null) {
    return { ok: true, applied: "copy", range: seed };
  }
  if (direction === "shrink-rows" || direction === "shrink-cols") {
    const band = direction === "shrink-rows" ? { ...seed, top: target.bottom + 1 } : { ...seed, left: target.right + 1 };
    if (firstLockedCell(model, sheetIndex, [band])) {
      return { ok: false, message: PROTECTED_CELL_MESSAGE };
    }
    if (arrayRangeTouched(model, sheetIndex, [band])) {
      return { ok: false, message: PARTIAL_ARRAY_MESSAGE };
    }
    clearContents(model, sheetIndex, [band]);
    return { ok: true, applied: "copy", range: target, cleared: true };
  }
  const axis = direction === "down" || direction === "up" ? "row" : "col";
  const band: GridRange =
    direction === "down"
      ? { ...seed, top: seed.bottom + 1, bottom: target.bottom }
      : direction === "up"
        ? { ...seed, top: target.top, bottom: seed.top - 1 }
        : direction === "right"
          ? { ...seed, left: seed.right + 1, right: target.right }
          : { ...seed, left: target.left, right: seed.left - 1 };
  const refusal = fillRefusal(model, sheetIndex, band);
  if (refusal !== null) {
    return { ok: false, message: refusal };
  }
  const grid = seedCells(model, sheet, seed, locale);
  const lines: SeedCell[][] = axis === "row" ? grid[0]!.map((_, c) => grid.map((row) => row[c]!)) : grid;
  const n = axis === "row" ? seed.bottom - seed.top + 1 : seed.right - seed.left + 1;
  const edit = new SheetEdit(sheet);
  let usedSeries = false;
  lines.forEach((line, k) => {
    const series = detectSeries(line, locale, { axis, dateSystem: model.dateSystem });
    const extends_ = series.kind !== "repeat" && (mode === "series" || (mode === "toggle" ? !series.byDefault : mode === "auto" || mode === "no-formats" ? series.byDefault : false));
    usedSeries ||= extends_;
    const from = axis === "row" ? band.top : band.left;
    const to = axis === "row" ? band.bottom : band.right;
    for (let at = from; at <= to; at += 1) {
      const i = at - (axis === "row" ? seed.top : seed.left);
      const row = axis === "row" ? at : seed.top + k;
      const col = axis === "row" ? seed.left + k : at;
      const source = line[mod(i, n)]!;
      const change: CellChange = {};
      if (mode !== "formats") {
        const value = extends_ ? series.at(i) : copyAt(line, i, axis);
        change.content = { value: value.value };
        if (value.formula !== undefined) {
          change.content.formula = value.formula;
        }
        const runs = (value as { runs?: TextRun[] }).runs;
        if (runs && mode !== "no-formats") {
          change.content.runs = runs;
        }
      }
      if (mode !== "no-formats") {
        change.style = source.styleId;
      }
      edit.change(row, col, change);
    }
  });
  const merges = fillMerges(sheet, seed, band, axis, n);
  const rows = edit.rows();
  if (rows !== sheet.rows || merges !== sheet.merges) {
    model.sheets = model.sheets.map((each, index) => (index === sheetIndex ? { ...each, rows, merges, dirty: true as const } : each));
  }
  const applied: AppliedFill = mode === "formats" || mode === "no-formats" ? mode : usedSeries ? "series" : "copy";
  const range = { top: Math.min(seed.top, target.top), left: Math.min(seed.left, target.left), bottom: Math.max(seed.bottom, target.bottom), right: Math.max(seed.right, target.right) };
  return { ok: true, applied, range };
}

/** The sheet's merges after a fill: those inside the band replaced by the
 * seed's merges repeated at every whole copy that fits in it. */
function fillMerges(sheet: Sheet, seed: GridRange, band: GridRange, axis: "row" | "col", n: number): GridRange[] {
  const own = sheet.merges.filter((merge) => contains(seed, merge));
  const inside = sheet.merges.filter((merge) => contains(band, merge));
  if (own.length === 0 && inside.length === 0) {
    return sheet.merges;
  }
  const copies: GridRange[] = [];
  const first = axis === "row" ? Math.floor((band.top - seed.top) / n) : Math.floor((band.left - seed.left) / n);
  const last = axis === "row" ? Math.floor((band.bottom - seed.top) / n) : Math.floor((band.right - seed.left) / n);
  for (let k = first; k <= last; k += 1) {
    for (const merge of own) {
      const copy = axis === "row" ? { ...merge, top: merge.top + k * n, bottom: merge.bottom + k * n } : { ...merge, left: merge.left + k * n, right: merge.right + k * n };
      if (contains(band, copy)) {
        copies.push(copy);
      }
    }
  }
  return [...sheet.merges.filter((merge) => !inside.includes(merge)), ...copies];
}

// --- fill down / right ----------------------------------------------------------------------

const isWholeColumns = (range: GridRange): boolean => range.top === 0 && range.bottom === MAX_ROWS - 1;
const isWholeRows = (range: GridRange): boolean => range.left === 0 && range.right === MAX_COLS - 1;

/** Ctrl/Cmd+D ("down") or Ctrl/Cmd+R ("right") over the selection's ranges
 * (see the module header). */
export function fillDownRight(model: SpreadsheetModel, sheetId: string, ranges: readonly GridRange[], direction: "down" | "right"): FillResult {
  const sheetIndex = model.sheets.findIndex((each) => each.id === sheetId);
  const sheet = model.sheets[sheetIndex];
  if (!sheet || sheet.kind !== "grid") {
    return { ok: false, message: NOT_A_GRID_MESSAGE };
  }
  const used = usedRange(sheet);
  const plans: { source: GridRange; band: GridRange }[] = [];
  for (const selected of ranges) {
    const range = { ...selected };
    if (used && isWholeColumns(range)) {
      range.bottom = Math.max(range.top, used.bottom);
    }
    if (used && isWholeRows(range)) {
      range.right = Math.max(range.left, used.right);
    }
    if (direction === "down") {
      const single = range.top === range.bottom;
      if (single && range.top === 0) {
        continue;
      }
      const sourceRow = single ? range.top - 1 : range.top;
      plans.push({ source: { ...range, top: sourceRow, bottom: sourceRow }, band: single ? range : { ...range, top: range.top + 1 } });
    } else {
      const single = range.left === range.right;
      if (single && range.left === 0) {
        continue;
      }
      const sourceCol = single ? range.left - 1 : range.left;
      plans.push({ source: { ...range, left: sourceCol, right: sourceCol }, band: single ? range : { ...range, left: range.left + 1 } });
    }
  }
  if (plans.reduce((sum, plan) => sum + (plan.band.bottom - plan.band.top + 1) * (plan.band.right - plan.band.left + 1), 0) > MAX_COPY_CELLS) {
    return { ok: false, message: TOO_MANY_TO_CHANGE_MESSAGE };
  }
  for (const plan of plans) {
    const refusal = fillRefusal(model, sheetIndex, plan.band);
    if (refusal !== null) {
      return { ok: false, message: refusal };
    }
  }
  const locale = SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE];
  const edit = new SheetEdit(sheet);
  for (const { source, band } of plans) {
    const cells = seedCells(model, sheet, source, locale).flat();
    forEachPosition(sheet, band, (row, col) => {
      const k = direction === "down" ? col - source.left : row - source.top;
      const seed = cells[k]!;
      const change: CellChange = { content: { value: seed.value }, style: seed.styleId };
      if (seed.formula !== undefined) {
        change.content!.formula = direction === "down" ? shiftFormula(seed.formula, row - source.top, 0) : shiftFormula(seed.formula, 0, col - source.left);
      } else if (seed.runs) {
        change.content!.runs = structuredClone(seed.runs);
      }
      edit.change(row, col, change);
    });
  }
  const rows = edit.rows();
  if (rows !== sheet.rows) {
    model.sheets = model.sheets.map((each, index) => (index === sheetIndex ? { ...each, rows, dirty: true as const } : each));
  }
  const first = plans[0]?.band ?? ranges[0] ?? { top: 0, left: 0, bottom: 0, right: 0 };
  return { ok: true, applied: "copy", range: first };
}

// --- double-click fill ---------------------------------------------------------------------

const holdsContent = (cell: Cell | undefined): boolean => cell !== undefined && (cell.value !== null || cell.formula !== undefined);

/** The range a double-click on the fill handle fills (see the module
 * header): the selection extended down, or null when there is nothing to
 * follow. */
export function doubleClickFillTarget(sheet: Sheet, range: GridRange): GridRange | null {
  const { top, left, bottom, right } = range;
  if (bottom >= MAX_ROWS - 1) {
    return null;
  }
  let guide: number | null = null;
  if (left > 0 && holdsContent(getCell(sheet, bottom + 1, left - 1))) {
    guide = left - 1;
  } else if (right < MAX_COLS - 1 && holdsContent(getCell(sheet, bottom + 1, right + 1))) {
    guide = right + 1;
  }
  if (guide === null) {
    return null;
  }
  // The guide column's contiguous run from bottom + 1.
  let end = bottom;
  for (let r = firstRowAtOrAfter(sheet.rows, bottom + 1); r < sheet.rows.length; r += 1) {
    const row = sheet.rows[r]!;
    if (row.index !== end + 1 || !holdsContent(getCell(sheet, row.index, guide))) {
      break;
    }
    end = row.index;
  }
  // Stop before the first non-empty cell below the selection in its own columns.
  for (let r = firstRowAtOrAfter(sheet.rows, bottom + 1); r < sheet.rows.length && sheet.rows[r]!.index <= end; r += 1) {
    const row = sheet.rows[r]!;
    if (row.cells.some((cell) => cell.col >= left && cell.col <= right && holdsContent(cell))) {
      end = row.index - 1;
      break;
    }
  }
  return end > bottom ? { top, left, bottom: end, right } : null;
}
