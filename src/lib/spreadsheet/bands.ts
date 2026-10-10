// Row and column properties over whole bands (SPEC-124 "Sizes", "Header
// menu", the freeze commands) — the commands the header menu, the size
// dialogs, a header-boundary drag, the Format menu and View ▸ Freeze run.
// Each works on the sheet it names copy-on-write (a new rows array or span
// list, new Row and ColumnSpan objects for what it changes, every untouched
// one kept by identity), marks that sheet dirty, and returns a FormatResult;
// the view wraps each in ONE undo step (SPEC-121's record/commit).
//
// - Resize writes a row's `heightPt` + `customHeight` (a changed height
//   drops the row's ods `sourceStyleName`, SPEC-122's drop rule) or a
//   column's `widthPx` — dropping `sourceWidthChars` unless the column-width
//   dialog states the width typed (`sourceWidthChars`, the xlsx `<col
//   width>` it stands for). A size under 4 px hides instead, keeping the
//   size for unhide; any other size unhides.
// - Hide and unhide set `Row.hidden` (a property-only Row where none is
//   stored) and `ColumnSpan.hidden`, never touching a size. Every Row a
//   command creates for rows that store nothing counts against
//   MAX_FORMAT_MATERIALISED: past it the command is refused with
//   TOO_MANY_ROWS_MESSAGE.
// - Autofit writes the measured height (autofit.ts's autofitRowHeightPt)
//   and clears `customHeight`, or the measured width
//   (autofitColumnWidthPx); a size equal to the sheet default is stored as
//   no size at all.
// - On a protected sheet resizing, hiding and autofitting rows need its
//   formatRows permit and columns its formatColumns permit (SPEC-123), else
//   SPEC-122's protected message.
// - freezePanes sets `Sheet.frozen` (counted from row and column 0; 0 × 0
//   unfreezes).
import { autofitColumnWidthsPx, autofitRowHeightPt, type CellTextMeasurer } from "./autofit";
import { isEmptyRow, PROTECTED_CELL_MESSAGE } from "./cells";
import type { FormatResult } from "./formatting";
import { MAX_COLS, MAX_ROWS, type ColumnSpan, type Row, type Sheet, type SpreadsheetModel } from "./model";
import type { FormatContext } from "./numberFormat/render";
import { firstRowAtOrAfter, MAX_FORMAT_MATERIALISED } from "./restyle";
import { joinSpans, NOT_A_GRID_MESSAGE } from "./structure";

export type BandAxis = "row" | "col";

/** A run of whole rows or columns, inclusive. */
export interface Band {
  start: number;
  end: number;
}

export const TOO_MANY_ROWS_MESSAGE = "That's too many rows to change at once";

/** Rows and columns narrower than this hide. */
export const HIDE_BELOW_PX = 4;
/** The row height dialog's limit, in points. */
export const MAX_ROW_HEIGHT_PT = 409;
/** The column width dialog's limit, in characters. */
export const MAX_COLUMN_WIDTH_CHARS = 255;

/** Characters (the column-width dialog's unit) → the xlsx `<col width>`
 * they stand for: trunc((chars · mdw + 5) / mdw · 256) / 256, Excel's inverse
 * of SPEC-117's MDW formula (8.43 → 9.140625 → 64 px at MDW 7). */
export function charsToXlsxWidth(chars: number, mdw: number): number {
  return Math.trunc(((chars * mdw + 5) / mdw) * 256) / 256;
}

/** A column's px as the characters Excel shows: trunc((px − 5) / mdw · 100 +
 * 0.5) / 100, never below 0 (64 px → 8.43 at MDW 7). */
export function pxToChars(px: number, mdw: number): number {
  return Math.max(0, Math.trunc(((px - 5) / mdw) * 100 + 0.5) / 100);
}

/** An xlsx `<col width>` as the characters Excel shows. */
export function xlsxWidthToChars(width: number, mdw: number): number {
  return Math.max(0, Math.trunc(((width * mdw - 5) / mdw) * 100 + 0.5) / 100);
}

/** A height in points rounded to the nearest 0.75 pt (one whole px). */
export function roundRowHeightPt(pt: number): number {
  return Math.round(pt / 0.75) * 0.75;
}

/** The sheet a command works on, or why it cannot. */
function target(model: SpreadsheetModel, sheetId: string, axis: BandAxis): { sheet: Sheet; index: number } | { message: string } {
  const index = model.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = model.sheets[index];
  if (!sheet || sheet.kind !== "grid") {
    return { message: NOT_A_GRID_MESSAGE };
  }
  const protection = sheet.protection;
  if (protection && !(axis === "row" ? protection.allowFormatRows : protection.allowFormatColumns)) {
    return { message: PROTECTED_CELL_MESSAGE };
  }
  return { sheet, index };
}

/** Commits a changed sheet copy-on-write, marked dirty. */
function commit(model: SpreadsheetModel, index: number, before: Sheet, next: Sheet): void {
  if (next !== before) {
    model.sheets = model.sheets.map((sheet, i) => (i === index ? { ...next, dirty: true as const } : sheet));
  }
}

/** Bands sorted, clamped to the axis and merged where they touch. */
export function normaliseBands(bands: readonly Band[], axis: BandAxis): Band[] {
  const last = (axis === "row" ? MAX_ROWS : MAX_COLS) - 1;
  const sorted = bands
    .map((band) => ({ start: Math.max(0, Math.min(band.start, band.end)), end: Math.min(last, Math.max(band.start, band.end)) }))
    .toSorted((a, b) => a.start - b.start);
  const out: Band[] = [];
  for (const band of sorted) {
    const prev = out.at(-1);
    if (prev && band.start <= prev.end + 1) {
      prev.end = Math.max(prev.end, band.end);
    } else {
      out.push({ ...band });
    }
  }
  return out;
}

// --- rows ------------------------------------------------------------------------------------

/** The sheet's rows with every row of `bands` passed through `patch` (a
 * property-only row created where none is stored when `create`); null when
 * that would create more than MAX_FORMAT_MATERIALISED rows. */
function patchRows(sheet: Sheet, bands: readonly Band[], patch: (row: Row) => Row, create: boolean): Row[] | null {
  if (create) {
    let missing = 0;
    for (const band of bands) {
      let stored = 0;
      for (let r = firstRowAtOrAfter(sheet.rows, band.start); r < sheet.rows.length && sheet.rows[r]!.index <= band.end; r += 1) {
        stored += 1;
      }
      missing += band.end - band.start + 1 - stored;
      if (missing > MAX_FORMAT_MATERIALISED) {
        return null;
      }
    }
  }
  const rows = sheet.rows;
  const out: Row[] = [];
  let changed = false;
  let r = 0;
  const push = (before: Row | undefined, next: Row): void => {
    if (next !== before) {
      changed = true;
    }
    if (!isEmptyRow(next)) {
      out.push(next);
    } else if (before) {
      changed = true;
    }
  };
  for (const band of bands) {
    while (r < rows.length && rows[r]!.index < band.start) {
      out.push(rows[r]!);
      r += 1;
    }
    for (let index = band.start; index <= band.end; index += 1) {
      const stored = rows[r]?.index === index ? rows[r] : undefined;
      if (stored) {
        push(stored, patch(stored));
        r += 1;
      } else if (create) {
        push(undefined, patch({ index, cells: [] }));
      } else {
        // Jump to the next stored row inside the band.
        const nextStored = rows[r]?.index;
        if (nextStored === undefined || nextStored > band.end) {
          break;
        }
        index = nextStored - 1;
      }
    }
  }
  while (r < rows.length) {
    out.push(rows[r]!);
    r += 1;
  }
  return changed ? out : rows;
}

function withHeight(row: Row, heightPt: number | undefined, customHeight: boolean): Row {
  const next: Row = { ...row };
  if (heightPt === undefined) {
    delete next.heightPt;
  } else {
    next.heightPt = heightPt;
  }
  if (customHeight) {
    next.customHeight = true;
  } else {
    delete next.customHeight;
  }
  if (next.heightPt !== row.heightPt) {
    delete next.sourceStyleName;
  }
  return next.heightPt === row.heightPt && next.customHeight === row.customHeight ? row : next;
}

function withRowHidden(row: Row, hidden: boolean): Row {
  if ((row.hidden === true) === hidden) {
    return row;
  }
  const next: Row = { ...row };
  if (hidden) {
    next.hidden = true;
  } else {
    delete next.hidden;
  }
  return next;
}

function rowCommand(model: SpreadsheetModel, sheetId: string, bands: readonly Band[], patch: (row: Row, sheet: Sheet) => Row, create: boolean): FormatResult {
  const found = target(model, sheetId, "row");
  if ("message" in found) {
    return { ok: false, message: found.message };
  }
  const rows = patchRows(found.sheet, normaliseBands(bands, "row"), (row) => patch(row, found.sheet), create);
  if (rows === null) {
    return { ok: false, message: TOO_MANY_ROWS_MESSAGE };
  }
  commit(model, found.index, found.sheet, rows === found.sheet.rows ? found.sheet : { ...found.sheet, rows });
  return { ok: true };
}

/** Sets every row of `bands` `heightPt` tall with `customHeight` (rounded to
 * 0.75 pt, at most 409); 0 hides them instead, keeping their size. */
export function resizeRows(model: SpreadsheetModel, sheetId: string, bands: readonly Band[], heightPt: number): FormatResult {
  const pt = Math.min(MAX_ROW_HEIGHT_PT, Math.max(0, roundRowHeightPt(heightPt)));
  if ((pt * 4) / 3 < HIDE_BELOW_PX) {
    return setRowsHidden(model, sheetId, bands, true);
  }
  return rowCommand(model, sheetId, bands, (row) => withRowHidden(withHeight(row, pt, true), false), true);
}

/** Hides or unhides every row of `bands` (see the module header). */
export function setRowsHidden(model: SpreadsheetModel, sheetId: string, bands: readonly Band[], hidden: boolean): FormatResult {
  return rowCommand(model, sheetId, bands, (row) => withRowHidden(row, hidden), hidden);
}

/** Autofits every stored row of `bands`: the measured height, customHeight
 * cleared (see the module header). */
export function autofitRows(
  model: SpreadsheetModel,
  sheetId: string,
  bands: readonly Band[],
  measurer: CellTextMeasurer,
  ctx: FormatContext,
): FormatResult {
  return rowCommand(
    model,
    sheetId,
    bands,
    (row, sheet) => {
      const pt = autofitRowHeightPt(model, sheet, row.index, measurer, ctx);
      return withHeight(row, pt === sheet.defaultRowHeightPt ? undefined : pt, false);
    },
    false,
  );
}

// --- columns ---------------------------------------------------------------------------------

/** Whether a span still stores something. */
function spanHoldsProperties(span: ColumnSpan): boolean {
  return Object.keys(span).some((key) => key !== "min" && key !== "max" && span[key as keyof ColumnSpan] !== undefined);
}

/** `columns` with every column of `bands` passed through `patch` — spans
 * split at the bands' edges, default columns materialised as spans for the
 * patch, spans left with no property dropped, equal neighbours joined. */
export function patchColumns(columns: readonly ColumnSpan[], bands: readonly Band[], patch: (span: ColumnSpan) => ColumnSpan): ColumnSpan[] {
  const pieces: ColumnSpan[] = [];
  // Bands are sorted and disjoint (they may touch): one cursor walks them.
  let b = 0;
  const inBand = (col: number): boolean => {
    while (b < bands.length && bands[b]!.end < col) {
      b += 1;
    }
    return b < bands.length && bands[b]!.start <= col;
  };
  // Cut points: every span edge and band edge.
  const cuts = new Set<number>();
  for (const span of columns) {
    cuts.add(span.min);
    cuts.add(span.max + 1);
  }
  for (const band of bands) {
    cuts.add(band.start);
    cuts.add(band.end + 1);
  }
  const points = [...cuts].filter((p) => p >= 0 && p <= MAX_COLS).toSorted((x, y) => x - y);
  let s = 0;
  for (let k = 0; k + 1 < points.length; k += 1) {
    const min = points[k]!;
    const max = points[k + 1]! - 1;
    while (s < columns.length && columns[s]!.max < min) {
      s += 1;
    }
    const covering = columns[s] && columns[s]!.min <= min ? columns[s] : undefined;
    const base: ColumnSpan = covering ? { ...covering, min, max } : { min, max };
    const piece = inBand(min) ? { ...patch(base), min, max } : base;
    if (spanHoldsProperties(piece)) {
      pieces.push(piece);
    }
  }
  return joinSpans(pieces);
}

/** Whether two span lists hold the same spans. */
function sameColumns(a: readonly ColumnSpan[], b: readonly ColumnSpan[]): boolean {
  return a.length === b.length && a.every((span, i) => JSON.stringify(span) === JSON.stringify(b[i]));
}

function columnCommand(model: SpreadsheetModel, sheetId: string, bands: readonly Band[], patch: (span: ColumnSpan, sheet: Sheet) => ColumnSpan): FormatResult {
  const found = target(model, sheetId, "col");
  if ("message" in found) {
    return { ok: false, message: found.message };
  }
  const columns = patchColumns(found.sheet.columns, normaliseBands(bands, "col"), (span) => patch(span, found.sheet));
  commit(model, found.index, found.sheet, sameColumns(columns, found.sheet.columns) ? found.sheet : { ...found.sheet, columns });
  return { ok: true };
}

function withWidth(span: ColumnSpan, widthPx: number | undefined, sourceWidthChars: number | undefined): ColumnSpan {
  const next: ColumnSpan = { ...span };
  if (widthPx === undefined) {
    delete next.widthPx;
  } else {
    next.widthPx = widthPx;
  }
  if (sourceWidthChars === undefined) {
    delete next.sourceWidthChars;
  } else {
    next.sourceWidthChars = sourceWidthChars;
  }
  return next;
}

function withColumnHidden(span: ColumnSpan, hidden: boolean): ColumnSpan {
  const next: ColumnSpan = { ...span };
  if (hidden) {
    next.hidden = true;
  } else {
    delete next.hidden;
  }
  return next;
}

/** Sets every column of `bands` `widthPx` wide (`sourceWidthChars`: the
 * xlsx width the column-width dialog's typed value stands for); under 4 px
 * hides them instead, keeping their size. */
export function resizeColumns(model: SpreadsheetModel, sheetId: string, bands: readonly Band[], widthPx: number, sourceWidthChars?: number): FormatResult {
  if (widthPx < HIDE_BELOW_PX) {
    return setColumnsHidden(model, sheetId, bands, true);
  }
  const px = Math.round(widthPx);
  return columnCommand(model, sheetId, bands, (span, sheet) =>
    withColumnHidden(withWidth(span, px === sheet.defaultColWidthPx && sourceWidthChars === undefined ? undefined : px, sourceWidthChars), false),
  );
}

/** Hides or unhides every column of `bands`. */
export function setColumnsHidden(model: SpreadsheetModel, sheetId: string, bands: readonly Band[], hidden: boolean): FormatResult {
  return columnCommand(model, sheetId, bands, (span) => withColumnHidden(span, hidden));
}

/** Autofits every column of `bands` (see the module header). */
export function autofitColumns(
  model: SpreadsheetModel,
  sheetId: string,
  bands: readonly Band[],
  measurer: CellTextMeasurer,
  ctx: FormatContext,
): FormatResult {
  const found = target(model, sheetId, "col");
  if ("message" in found) {
    return { ok: false, message: found.message };
  }
  const sheet = found.sheet;
  // Runs of columns measuring the same width — each patched as one piece
  // (the runs touch, so they are cut, never merged, by patchColumns).
  const runs: (Band & { px: number })[] = [];
  for (const band of normaliseBands(bands, "col")) {
    for (const [col, px] of autofitColumnWidthsPx(model, sheet, band.start, band.end, measurer, ctx)) {
      const last = runs.at(-1);
      if (last && last.px === px && last.end === col - 1) {
        last.end = col;
      } else {
        runs.push({ start: col, end: col, px });
      }
    }
  }
  const widthAt = (col: number): number => runs.find((run) => col >= run.start && col <= run.end)?.px ?? sheet.defaultColWidthPx;
  const columns = patchColumns(sheet.columns, runs, (span) => {
    const px = widthAt(span.min);
    return withWidth(span, px === sheet.defaultColWidthPx ? undefined : px, undefined);
  });
  commit(model, found.index, sheet, sameColumns(columns, sheet.columns) ? sheet : { ...sheet, columns });
  return { ok: true };
}

// --- freeze ----------------------------------------------------------------------------------

/** Freezes the first `rows` rows and `cols` columns (0 × 0 unfreezes). */
export function freezePanes(model: SpreadsheetModel, sheetId: string, rows: number, cols: number): FormatResult {
  const index = model.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = model.sheets[index];
  if (!sheet || sheet.kind !== "grid") {
    return { ok: false, message: NOT_A_GRID_MESSAGE };
  }
  const current = sheet.frozen ?? { rows: 0, cols: 0 };
  if (current.rows === rows && current.cols === cols) {
    return { ok: true };
  }
  const next: Sheet = { ...sheet };
  if (rows === 0 && cols === 0) {
    delete next.frozen;
  } else {
    next.frozen = { rows: Math.min(rows, MAX_ROWS - 1), cols: Math.min(cols, MAX_COLS - 1) };
  }
  commit(model, index, sheet, next);
  return { ok: true };
}
