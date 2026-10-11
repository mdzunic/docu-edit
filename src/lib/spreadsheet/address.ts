// A1 addressing and sparse-sheet lookups (SPEC-116) — the normative helpers
// every later spreadsheet spec addresses cells through. Coordinates are
// 0-based everywhere in the model; A1 text is 1-based rows and lettered
// columns. Letters are case-insensitive on input and upper-case on output.
// (core/tables.ts's unrelated `CellRange` is the text product's; spreadsheet
// code never uses that name.)
import { MAX_COLS, MAX_ROWS, type Cell, type Row, type Sheet } from "./model";

export interface CellRef {
  row: number;
  col: number;
} // 0-based
export interface GridRange {
  top: number;
  left: number;
  bottom: number;
  right: number;
} // inclusive, normalised

/** 0 → "A", 25 → "Z", 26 → "AA", 16383 → "XFD" (bijective base 26). */
export function colToLetters(col: number): string {
  if (!Number.isInteger(col) || col < 0) {
    throw new RangeError(`Not a column index: ${col}`);
  }
  let letters = "";
  let n = col + 1;
  while (n > 0) {
    const digit = (n - 1) % 26;
    letters = String.fromCharCode(65 + digit) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/** "A" → 0 … "XFD" → 16383, case-insensitively; null for anything that is
 * not one to three letters naming a column inside the grid. */
export function lettersToCol(letters: string): number | null {
  if (!/^[A-Za-z]{1,3}$/.test(letters)) {
    return null;
  }
  let n = 0;
  for (let i = 0; i < letters.length; i += 1) {
    n = n * 26 + ((letters.charCodeAt(i) | 0x20) - 0x60);
  }
  const col = n - 1;
  return col < MAX_COLS ? col : null;
}

/** { row: 2, col: 1 } → "B3". */
export function formatA1(ref: CellRef): string {
  return `${colToLetters(ref.col)}${ref.row + 1}`;
}

const A1_PATTERN = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})$/;

/** "B3", "$B$3", "b$3" → the 0-based cell plus which parts were absolute;
 * null (never a throw) for anything outside A1..XFD1048576 or not an A1
 * reference at all. */
export function parseA1(text: string): (CellRef & { rowAbs: boolean; colAbs: boolean }) | null {
  const match = A1_PATTERN.exec(text);
  if (!match) {
    return null;
  }
  const col = lettersToCol(match[2]!);
  const row = Number(match[4]) - 1;
  if (col === null || row >= MAX_ROWS) {
    return null;
  }
  return { row, col, colAbs: match[1] === "$", rowAbs: match[3] === "$" };
}

/** A single cell → "B3"; anything larger → "A1:C3". */
export function formatRange(range: GridRange): string {
  const start = formatA1({ row: range.top, col: range.left });
  if (range.top === range.bottom && range.left === range.right) {
    return start;
  }
  return `${start}:${formatA1({ row: range.bottom, col: range.right })}`;
}

/** "A1:C3", "C3:A1" (normalised to top-left:bottom-right) or "B3"; null for
 * anything else. */
export function parseRange(text: string): GridRange | null {
  const parts = text.split(":");
  if (parts.length > 2) {
    return null;
  }
  const first = parseA1(parts[0]!);
  const second = parts.length === 2 ? parseA1(parts[1]!) : first;
  if (!first || !second) {
    return null;
  }
  return {
    top: Math.min(first.row, second.row),
    left: Math.min(first.col, second.col),
    bottom: Math.max(first.row, second.row),
    right: Math.max(first.col, second.col),
  };
}

/** A space-separated range list (`sqref="A1:B2 D4"`); invalid entries are
 * dropped. */
export function parseSqref(text: string): GridRange[] {
  const ranges: GridRange[] = [];
  for (const entry of text.split(/\s+/)) {
    if (entry.length === 0) {
      continue;
    }
    const range = parseRange(entry);
    if (range) {
      ranges.push(range);
    }
  }
  return ranges;
}

/** The stored row at 0-based `row` — a binary search over the sparse,
 * ascending rows; undefined when the row stores nothing. */
export function findRow(sheet: Sheet, row: number): Row | undefined {
  const rows = sheet.rows;
  let lo = 0;
  let hi = rows.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    const index = rows[mid]!.index;
    if (index === row) {
      return rows[mid];
    }
    if (index < row) {
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return undefined;
}

/** The stored cell at (row, col) — binary searches over the sparse rows and
 * then the row's sparse cells; undefined for an empty cell. */
export function getCell(sheet: Sheet, row: number, col: number): Cell | undefined {
  const cells = findRow(sheet, row)?.cells;
  if (!cells) {
    return undefined;
  }
  let lo = 0;
  let hi = cells.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    const at = cells[mid]!.col;
    if (at === col) {
      return cells[mid];
    }
    if (at < col) {
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return undefined;
}

/** The bounding box of the sheet's stored cells; null for a sheet that
 * stores none. */
export function usedRange(sheet: Sheet): GridRange | null {
  let range: GridRange | null = null;
  for (const row of sheet.rows) {
    const first = row.cells[0];
    const last = row.cells.at(-1);
    if (!first || !last) {
      continue;
    }
    if (!range) {
      range = { top: row.index, left: first.col, bottom: row.index, right: last.col };
      continue;
    }
    range.bottom = row.index;
    range.left = Math.min(range.left, first.col);
    range.right = Math.max(range.right, last.col);
  }
  return range;
}

/** The cells of `ranges` outside `cut` (SPEC-129, the contract SPEC-130
 * states identically), as rectangles: each range in source order, the same
 * object when `cut` misses it, else its uncovered parts — the band above
 * the cut, then the parts left and right of it, then the band below — each
 * remainder cut again by the next range of `cut`. A range `cut` covers
 * entirely contributes nothing. Pure. */
export function subtractRanges(ranges: readonly GridRange[], cut: GridRange | readonly GridRange[]): GridRange[] {
  const cuts: readonly GridRange[] = "top" in cut ? [cut as GridRange] : (cut as readonly GridRange[]);
  let out: GridRange[] = [...ranges];
  for (const piece of cuts) {
    out = out.flatMap((range) => subtractOne(range, piece));
  }
  return out;
}

function subtractOne(range: GridRange, cut: GridRange): GridRange[] {
  if (cut.top > range.bottom || cut.bottom < range.top || cut.left > range.right || cut.right < range.left) {
    return [range];
  }
  const pieces: GridRange[] = [];
  const top = Math.max(range.top, cut.top);
  const bottom = Math.min(range.bottom, cut.bottom);
  if (range.top < cut.top) {
    pieces.push({ top: range.top, left: range.left, bottom: cut.top - 1, right: range.right });
  }
  if (range.left < cut.left) {
    pieces.push({ top, left: range.left, bottom, right: cut.left - 1 });
  }
  if (range.right > cut.right) {
    pieces.push({ top, left: cut.right + 1, bottom, right: range.right });
  }
  if (range.bottom > cut.bottom) {
    pieces.push({ top: cut.bottom + 1, left: range.left, bottom: range.bottom, right: range.right });
  }
  return pieces;
}
