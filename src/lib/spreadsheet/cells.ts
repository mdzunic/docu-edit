// Cell commands (SPEC-122 "Cell commands") — THE normative source for how a
// committed entry, a Fill (Ctrl/Cmd+Enter) and Clear contents change the
// model, and for the protection and array checks that run before them so
// nothing partial is ever written.
//
// - Copy-on-write (model.ts invariant 1): an edit replaces `model.sheets`
//   with a new array holding a NEW object for the touched sheet, gives it a
//   new `rows` array, and builds new Row objects only for touched rows and
//   new Cell objects only for changed cells; untouched rows and cells keep
//   their identity. A command that changes nothing (an entry equal to the
//   cell's content) leaves `model.sheets` and `rows` untouched, so the undo
//   pipeline's identity diff records nothing. `model.styles` grows only
//   through internStyle. Marking the sheet dirty is the undo pipeline's job
//   (SPEC-121's commit), never this module's.
// - Sparse (invariant 2): a cell left with no value, formula, style or
//   preserved content is removed — any style, formula or preserved field
//   keeps it (a cell holding only preserved content is content) — and a row
//   left with no cells and no properties is removed. A row's properties are
//   what the readers store an empty row for — a height, `hidden` (or
//   SPEC-128's `filterHidden`), a `styleId` — and any unmodelled source
//   attribute other than Excel's
//   layout caches (`spans`, `x14ac:dyDescent`, which the writer drops); an
//   ods row's `sourceStyleName` alone is no property (its height is
//   `heightPt`), so an emptied row reads back the way it is stored.
// - Styles: the new style is the position's effective style (effectiveStyleId
//   → cellStyle) with the entry's `numberFormat`/`builtinFormatId` applied
//   only when that style's format is General (isGeneralFormat: the code it
//   renders with is General — `parseInput` only suggests; this is where the
//   General-only rule is applied; a `numberFormat` without a
//   `builtinFormatId` clears it, SPEC-118's rule), `quotePrefix` set when the
//   entry has it and removed otherwise, and `alignment.wrap` set for `wrap` —
//   interned through internStyle. An unchanged effective style keeps the
//   cell's `styleId` exactly as it was (absent stays absent); a changed one
//   is the interned index, stored on the cell — except the default style 0
//   where the position inherits 0 anyway (no row or column style), which is
//   left absent (the same thing to every reader and writer).
// - Drop rules ("Model additions"): a changed value or runs — or a changed
//   or removed formula, whose result the cached value was — removes
//   `sstIndex`, `preservedText` and ` vm="…"` (and `runs`: a changed text on
//   a rich-text cell stores plain text, an unchanged one keeps its runs); a
//   changed or removed formula removes `cm`, `dynamicArray`, `dataTableXml`
//   and `formulaUnparsed` (the new one is canonical). A literal typed over a
//   data-table anchor replaces its formula. An unchanged formula keeps the
//   cell's cached value (the entry of a formula carries no value;
//   recalculation computes it).
// - Arrays: entering a formula into an array anchor keeps its `arrayRange`
//   (an edited dynamic array becomes a legacy array: `cm` and `dynamicArray`
//   go). An anchor that loses its formula — a literal typed over it, or
//   Clear contents — loses `arrayRange` too, and the members' cached values
//   (the array's results) are cleared with it; a Fill that writes the anchor
//   and other cells of its range replaces the whole array with plain cells.
//   Editing a member alone is refused before any command runs
//   (arrayRangeTouched).
// - fillWithEntry writes each target the entry with its formula shifted
//   relatively by shiftFormula(formula, row − origin.row, col − origin.col);
//   a merged range receives it in its anchor only; a whole-row or
//   whole-column target is clipped to the used range — the used range the
//   active cell's own entry leaves, so the origin is always filled.
// - clearContents removes values and formulas of the STORED cells in the
//   ranges — never iterating empty positions, so a whole column costs its
//   stored cells — keeping styles, `preservedXml` (notes, annotations) and
//   other preserved attributes.
// - firstLockedCell / arrayRangeTouched answer over stored rows, cells and
//   column spans only, never per position of a whole-row/column selection.
import { findRow, getCell, usedRange, type CellRef, type GridRange } from "./address";
import { shiftFormula } from "./formula/references";
import { buildMergeIndex, mergeAt, rangesIntersect } from "./merges";
import { MAX_COLS, MAX_ROWS, dropValueBound, isCellError, rowSourceAttrsHoldProperties, type Cell, type CellStyle, type CellValue, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { BUILTIN_NUMBER_FORMATS } from "./numberFormat/builtin";
import type { ParsedInput } from "./numberFormat/input";
import { cellStyle, columnSpanAt, effectiveStyleId, internStyle, styleKey } from "./styles";

/** The refusals checked before a cell or formatting command runs (SPEC-122
 * "Protection"; SPEC-123's formatting commands refuse with the same text). */
export const PROTECTED_CELL_MESSAGE = "This cell is on a protected sheet";
export const ARRAY_PART_MESSAGE = "You can't change part of an array";

/** What a commit writes into a cell (parseInput's result, plus `wrap` for a
 * text holding a line break). */
export interface CellEntry {
  value: CellValue;
  formula?: string;
  numberFormat?: string;
  builtinFormatId?: number;
  quotePrefix?: true;
  wrap?: true;
}

/** A CellEntry from parseInput's result. `wrap` defaults to whether the
 * entry is a text holding a line break (Alt/Option+Enter): such a commit
 * turns wrapping on, as in Excel. */
export function entryFromInput(parsed: ParsedInput, wrap?: boolean): CellEntry {
  const entry: CellEntry = { value: parsed.value };
  if (parsed.formula !== undefined) {
    entry.formula = parsed.formula;
  }
  if (parsed.numberFormat !== undefined) {
    entry.numberFormat = parsed.numberFormat;
  }
  if (parsed.builtinFormatId !== undefined) {
    entry.builtinFormatId = parsed.builtinFormatId;
  }
  if (parsed.quotePrefix) {
    entry.quotePrefix = true;
  }
  const wraps = wrap ?? (parsed.formula === undefined && typeof parsed.value === "string" && parsed.value.includes("\n"));
  if (wraps) {
    entry.wrap = true;
  }
  return entry;
}

/** Whether two cell values are the same (errors by code). */
export function sameCellValue(a: CellValue, b: CellValue): boolean {
  if (isCellError(a) || isCellError(b)) {
    return isCellError(a) && isCellError(b) && a.error === b.error;
  }
  return a === b;
}

/** Whether two entries are equal field by field. */
export function cellEntryEquals(a: CellEntry, b: CellEntry): boolean {
  return (
    sameCellValue(a.value, b.value) &&
    a.formula === b.formula &&
    a.numberFormat === b.numberFormat &&
    a.builtinFormatId === b.builtinFormatId &&
    a.quotePrefix === b.quotePrefix &&
    a.wrap === b.wrap
  );
}

/** Whether a style's number format is General: the code it renders with
 * (formatCodeOf's choice — a built-in id with an entry, else `numberFormat`)
 * is absent or "General". Locale-dependent built-ins are never General, so
 * the answer is the same in every locale. */
export function isGeneralFormat(style: Pick<CellStyle, "numberFormat" | "builtinFormatId">): boolean {
  const code = (style.builtinFormatId === undefined ? undefined : BUILTIN_NUMBER_FORMATS[style.builtinFormatId]) ?? style.numberFormat;
  return code === undefined || code.toLowerCase() === "general";
}

// --- structural helpers -------------------------------------------------------------

function deepEqual(a: unknown, b: unknown): boolean {
  if (a === b) {
    return true;
  }
  if (typeof a !== "object" || typeof b !== "object" || a === null || b === null) {
    return false;
  }
  if (Array.isArray(a) !== Array.isArray(b)) {
    return false;
  }
  const ra = a as Record<string, unknown>;
  const rb = b as Record<string, unknown>;
  const ka = Object.keys(ra).filter((key) => ra[key] !== undefined);
  const kb = Object.keys(rb).filter((key) => rb[key] !== undefined);
  return ka.length === kb.length && ka.every((key) => deepEqual(ra[key], rb[key]));
}

/** A cell storing nothing: a null value and no other field. */
function isEmptyCell(cell: Cell): boolean {
  if (cell.value !== null) {
    return false;
  }
  const record = cell as unknown as Record<string, unknown>;
  return Object.keys(record).every((key) => key === "col" || key === "value" || record[key] === undefined);
}

/** A row storing nothing (see the module header): no cells, no height,
 * `hidden` or `styleId` (nor any later field), no unmodelled source attribute
 * but layout caches; an ods row style name alone does not count. */
export function isEmptyRow(row: Row): boolean {
  if (row.cells.length > 0 || rowSourceAttrsHoldProperties(row.sourceAttrs)) {
    return false;
  }
  const record = row as unknown as Record<string, unknown>;
  return Object.keys(record).every(
    (key) => key === "index" || key === "cells" || key === "sourceAttrs" || key === "sourceStyleName" || record[key] === undefined,
  );
}

/** Removes the fields a changed or removed formula invalidates. */
function dropFormulaBound(cell: Cell): void {
  delete cell.cm;
  delete cell.dynamicArray;
  delete cell.dataTableXml;
  delete cell.formulaUnparsed;
}

/** Positions of a range in row-major order. */
function* positions(range: GridRange): Generator<CellRef> {
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= range.right; col += 1) {
      yield { row, col };
    }
  }
}

/** The first stored row at or after `row` (an index into `rows`). */
function firstRowAtOrAfter(rows: readonly Row[], row: number): number {
  let lo = 0;
  let hi = rows.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (rows[mid]!.index < row) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return lo;
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

/** Calls `visit` for every stored cell inside `range`, row-major. */
function forEachStoredCell(sheet: Sheet, range: GridRange, visit: (row: Row, cell: Cell) => void): void {
  const rows = sheet.rows;
  for (let r = firstRowAtOrAfter(rows, range.top); r < rows.length && rows[r]!.index <= range.bottom; r += 1) {
    const row = rows[r]!;
    for (let c = firstCellAtOrAfter(row.cells, range.left); c < row.cells.length && row.cells[c]!.col <= range.right; c += 1) {
      visit(row, row.cells[c]!);
    }
  }
}

// --- the edit batch -------------------------------------------------------------------

/** Pending changes to one sheet: row → col → the new cell (undefined:
 * removed). */
type Edits = Map<number, Map<number, Cell | undefined>>;

function setEdit(edits: Edits, row: number, col: number, cell: Cell | undefined): void {
  let cols = edits.get(row);
  if (!cols) {
    cols = new Map();
    edits.set(row, cols);
  }
  cols.set(col, cell);
}

/** A row with `changes` applied: the same object when nothing changed,
 * undefined when it ends up storing nothing. */
function editRow(old: Row | undefined, index: number, changes: Map<number, Cell | undefined>): Row | undefined {
  const oldCells = old?.cells ?? [];
  const effective: [number, Cell | undefined][] = [];
  for (const [col, cell] of changes) {
    const at = firstCellAtOrAfter(oldCells, col);
    const existing = oldCells[at]?.col === col ? oldCells[at] : undefined;
    if (existing !== cell) {
      effective.push([col, cell]);
    }
  }
  if (effective.length === 0) {
    return old;
  }
  const cells: Cell[] = [];
  let i = 0;
  for (const [col, cell] of effective.toSorted((a, b) => a[0] - b[0])) {
    while (i < oldCells.length && oldCells[i]!.col < col) {
      cells.push(oldCells[i]!);
      i += 1;
    }
    if (i < oldCells.length && oldCells[i]!.col === col) {
      i += 1; // replaced or removed
    }
    if (cell) {
      cells.push(cell);
    }
  }
  while (i < oldCells.length) {
    cells.push(oldCells[i]!);
    i += 1;
  }
  const row: Row = old ? { ...old, cells } : { index, cells };
  return isEmptyRow(row) ? undefined : row;
}

/** Applies a batch copy-on-write (see the module header); a batch that
 * changes nothing leaves the model untouched. */
function applyEdits(model: SpreadsheetModel, sheetIndex: number, edits: Edits): void {
  const sheet = model.sheets[sheetIndex];
  if (!sheet || edits.size === 0) {
    return;
  }
  const replaced = new Map<number, Row | undefined>();
  for (const [index, changes] of edits) {
    const old = findRow(sheet, index);
    const next = editRow(old, index, changes);
    if (next !== old) {
      replaced.set(index, next);
    }
  }
  if (replaced.size === 0) {
    return;
  }
  replaceRows(model, sheetIndex, replaced);
}

/** `sheet.rows` with rows replaced (undefined: removed) or added, by index —
 * a new rows array on a new sheet object in a new sheets array. */
export function replaceRows(model: SpreadsheetModel, sheetIndex: number, replaced: ReadonlyMap<number, Row | undefined>): void {
  const sheet = model.sheets[sheetIndex]!;
  const keys = [...replaced.keys()].toSorted((a, b) => a - b);
  const rows: Row[] = [];
  let k = 0;
  for (const row of sheet.rows) {
    while (k < keys.length && keys[k]! < row.index) {
      const added = replaced.get(keys[k]!);
      if (added) {
        rows.push(added);
      }
      k += 1;
    }
    if (k < keys.length && keys[k] === row.index) {
      const next = replaced.get(row.index);
      if (next) {
        rows.push(next);
      }
      k += 1;
    } else {
      rows.push(row);
    }
  }
  for (; k < keys.length; k += 1) {
    const added = replaced.get(keys[k]!);
    if (added) {
      rows.push(added);
    }
  }
  const next: Sheet = { ...sheet, rows };
  model.sheets = model.sheets.map((s, i) => (i === sheetIndex ? next : s));
}

// --- entries ------------------------------------------------------------------------

/** The style id an entry leaves at a position: `keep` when its effective
 * style does not change, else the interned changed style. */
type StyleOutcome = { keep: true } | { keep: false; styleId: number };

function entryStyle(model: SpreadsheetModel, effectiveId: number, entry: CellEntry): StyleOutcome {
  const style = cellStyle(model, effectiveId);
  const next: CellStyle = { ...style };
  if (entry.numberFormat !== undefined && isGeneralFormat(style)) {
    next.numberFormat = entry.numberFormat;
    if (entry.builtinFormatId !== undefined) {
      next.builtinFormatId = entry.builtinFormatId;
    } else {
      delete next.builtinFormatId;
    }
  }
  if (entry.quotePrefix) {
    next.quotePrefix = true;
  } else {
    delete next.quotePrefix;
  }
  if (entry.wrap && !style.alignment?.wrap) {
    next.alignment = { ...style.alignment, wrap: true };
  }
  if (styleKey(next) === styleKey(style)) {
    return { keep: true };
  }
  return { keep: false, styleId: internStyle(model, next) };
}

/** Writes per-position entries into one batch, caching each effective
 * style's outcome per entry style (a Fill's targets share one entry; a
 * plain-text paste's differ). */
class EntryWriter {
  private readonly styleOutcomes = new Map<string, StyleOutcome>();
  /** Array members to clear once every target is written. */
  readonly memberClears: GridRange[] = [];

  constructor(
    private readonly model: SpreadsheetModel,
    private readonly sheet: Sheet,
    readonly edits: Edits,
  ) {}

  /** Writes `entry` at a position; `dissolveArray`: the same command writes
   * other cells of an array anchored here, so the array becomes plain
   * cells. */
  write(row: number, col: number, entry: CellEntry, dissolveArray = false): void {
    const cell = getCell(this.sheet, row, col);
    const effectiveId = effectiveStyleId(this.sheet, row, col);
    const key = `${effectiveId}|${entry.numberFormat ?? ""}|${entry.builtinFormatId ?? ""}|${entry.quotePrefix ? 1 : 0}|${entry.wrap ? 1 : 0}`;
    let outcome = this.styleOutcomes.get(key);
    if (!outcome) {
      outcome = entryStyle(this.model, effectiveId, entry);
      this.styleOutcomes.set(key, outcome);
    }
    // What the position shows without a style of its own (row → column → 0).
    const inheritedId = findRow(this.sheet, row)?.styleId ?? columnSpanAt(this.sheet.columns, col)?.styleId ?? 0;
    const next = entryCell(cell, col, entry, outcome, inheritedId, dissolveArray);
    if (cell?.arrayRange && next?.arrayRange === undefined) {
      this.memberClears.push(cell.arrayRange);
    }
    setEdit(this.edits, row, col, next);
  }
}

/** The cell an entry leaves (see the module header): `cell` itself when
 * nothing changes, undefined when it stores nothing. A changed style is the
 * interned index, left absent only when it is the default 0 the position
 * inherits anyway (`inheritedId`). */
function entryCell(
  cell: Cell | undefined,
  col: number,
  entry: CellEntry,
  style: StyleOutcome,
  inheritedId: number,
  dissolveArray: boolean,
): Cell | undefined {
  const formula = entry.formula;
  const formulaChanged = formula !== cell?.formula || cell?.dataTableXml !== undefined;
  // An unchanged formula keeps its cached result.
  const value = !formulaChanged && formula !== undefined ? (cell?.value ?? null) : entry.value;
  const valueChanged = !sameCellValue(value, cell?.value ?? null);
  const next: Cell = cell ? { ...cell } : { col, value: null };
  next.value = value;
  if (formula !== undefined) {
    next.formula = formula;
  } else {
    delete next.formula;
  }
  if (formulaChanged) {
    dropFormulaBound(next);
    if (formula === undefined) {
      delete next.arrayRange;
    }
  }
  if (dissolveArray && next.arrayRange !== undefined) {
    delete next.arrayRange;
    dropFormulaBound(next);
  }
  // A changed formula's cached value was its old result: what is bound to
  // the value goes with it, as with a changed literal.
  if (valueChanged || formulaChanged) {
    dropValueBound(next);
  }
  if (!style.keep) {
    if (style.styleId === 0 && inheritedId === 0) {
      delete next.styleId;
    } else {
      next.styleId = style.styleId;
    }
  }
  if (isEmptyCell(next)) {
    return undefined;
  }
  return cell && deepEqual(next, cell) ? cell : next;
}

/** The cell Clear contents leaves: values and formulas (and everything bound
 * to them) removed, styles and other preserved content kept. */
function clearedCell(cell: Cell): Cell | undefined {
  const next: Cell = { ...cell, value: null };
  delete next.formula;
  delete next.arrayRange;
  dropFormulaBound(next);
  dropValueBound(next);
  if (isEmptyCell(next)) {
    return undefined;
  }
  return deepEqual(next, cell) ? cell : next;
}

/** Clears the cached values of array members (every stored cell of `range`
 * but its anchor) that no target of the same command wrote. */
function clearMembers(sheet: Sheet, ranges: readonly GridRange[], edits: Edits, written: (row: number, col: number) => boolean): void {
  for (const range of ranges) {
    forEachStoredCell(sheet, range, (row, cell) => {
      if ((row.index === range.top && cell.col === range.left) || written(row.index, cell.col)) {
        return;
      }
      setEdit(edits, row.index, cell.col, clearedCell(cell));
    });
  }
}

/** Commits `entry` into one cell (see the module header). */
export function setCellEntry(model: SpreadsheetModel, sheetIndex: number, at: CellRef, entry: CellEntry): void {
  const sheet = model.sheets[sheetIndex];
  if (!sheet) {
    return;
  }
  const writer = new EntryWriter(model, sheet, new Map());
  writer.write(at.row, at.col, entry);
  clearMembers(sheet, writer.memberClears, writer.edits, (row, col) => row === at.row && col === at.col);
  applyEdits(model, sheetIndex, writer.edits);
}

const isWholeColumns = (range: GridRange): boolean => range.top === 0 && range.bottom === MAX_ROWS - 1;
const isWholeRows = (range: GridRange): boolean => range.left === 0 && range.right === MAX_COLS - 1;

/** A target range with its whole axes clipped to `bounds`; null when
 * nothing is left. */
function clipWhole(range: GridRange, bounds: GridRange | null): GridRange | null {
  const clipped = { ...range };
  if (isWholeColumns(range)) {
    if (!bounds) {
      return null;
    }
    clipped.top = Math.max(range.top, bounds.top);
    clipped.bottom = Math.min(range.bottom, bounds.bottom);
  }
  if (isWholeRows(range)) {
    if (!bounds) {
      return null;
    }
    clipped.left = Math.max(range.left, bounds.left);
    clipped.right = Math.min(range.right, bounds.right);
  }
  return clipped.top <= clipped.bottom && clipped.left <= clipped.right ? clipped : null;
}

/** The ranges a Fill writes: whole rows and columns clipped to the used
 * range the origin's own entry leaves — the used range grown to include
 * `origin`, so the active cell is always filled. */
export function fillTargets(sheet: Sheet, ranges: readonly GridRange[], origin: CellRef): GridRange[] {
  const used = usedRange(sheet);
  const bounds: GridRange = used
    ? {
        top: Math.min(used.top, origin.row),
        left: Math.min(used.left, origin.col),
        bottom: Math.max(used.bottom, origin.row),
        right: Math.max(used.right, origin.col),
      }
    : { top: origin.row, left: origin.col, bottom: origin.row, right: origin.col };
  return ranges.map((range) => clipWhole(range, bounds)).filter((range): range is GridRange => range !== null);
}

/** A position as one number (row-major; exact below 2^53). */
function positionKey(row: number, col: number): number {
  return row * MAX_COLS + col;
}

/** Whether `ranges` hold a cell of `array` other than its anchor. */
function touchesBeyondAnchor(ranges: readonly GridRange[], array: GridRange): boolean {
  return ranges.some((range) => {
    if (!rangesIntersect(range, array)) {
      return false;
    }
    const top = Math.max(range.top, array.top);
    const left = Math.max(range.left, array.left);
    const bottom = Math.min(range.bottom, array.bottom);
    const right = Math.min(range.right, array.right);
    return !(top === array.top && left === array.left && bottom === array.top && right === array.left);
  });
}

/** Ctrl/Cmd+Enter: writes `entry` into every target, its formula shifted
 * relatively from `origin` (see the module header). */
export function fillWithEntry(model: SpreadsheetModel, sheetIndex: number, ranges: GridRange[], origin: CellRef, entry: CellEntry): void {
  const sheet = model.sheets[sheetIndex];
  if (!sheet) {
    return;
  }
  const merges = buildMergeIndex(sheet.merges);
  const writer = new EntryWriter(model, sheet, new Map());
  const written = new Set<number>();
  const targets = fillTargets(sheet, ranges, origin);
  for (const range of targets) {
    for (const { row, col } of positions(range)) {
      const merge = mergeAt(merges, row, col);
      if ((merge && (merge.top !== row || merge.left !== col)) || written.has(positionKey(row, col))) {
        continue;
      }
      written.add(positionKey(row, col));
      const formula = entry.formula === undefined ? undefined : shiftFormula(entry.formula, row - origin.row, col - origin.col);
      const array = getCell(sheet, row, col)?.arrayRange;
      writer.write(row, col, formula === undefined ? entry : { ...entry, formula }, array !== undefined && touchesBeyondAnchor(targets, array));
    }
  }
  clearMembers(sheet, writer.memberClears, writer.edits, (row, col) => written.has(positionKey(row, col)));
  applyEdits(model, sheetIndex, writer.edits);
}

/** SPEC-125's plain-text paste: writes each position's entry in one batch,
 * exactly as setCellEntry would (the destination's style kept, the entry's
 * number format applied only over General); a null entry clears the
 * position's value and formula as Clear contents does. Positions the batch
 * writes dissolve any array they cover (the caller refused partial ones). */
export function writeEntries(model: SpreadsheetModel, sheetIndex: number, entries: Iterable<{ row: number; col: number; entry: CellEntry | null }>): void {
  const sheet = model.sheets[sheetIndex];
  if (!sheet) {
    return;
  }
  const writer = new EntryWriter(model, sheet, new Map());
  const written = new Set<number>();
  for (const { row, col, entry } of entries) {
    written.add(positionKey(row, col));
    const cell = getCell(sheet, row, col);
    if (entry) {
      writer.write(row, col, entry, cell?.arrayRange !== undefined);
    } else if (cell) {
      if (cell.arrayRange) {
        writer.memberClears.push(cell.arrayRange);
      }
      setEdit(writer.edits, row, col, clearedCell(cell));
    }
  }
  clearMembers(sheet, writer.memberClears, writer.edits, (row, col) => written.has(positionKey(row, col)));
  applyEdits(model, sheetIndex, writer.edits);
}

/** Delete / Backspace: clears the values and formulas of every stored cell
 * in `ranges` (see the module header). */
export function clearContents(model: SpreadsheetModel, sheetIndex: number, ranges: GridRange[]): void {
  const sheet = model.sheets[sheetIndex];
  if (!sheet) {
    return;
  }
  const edits: Edits = new Map();
  const arrays: GridRange[] = [];
  for (const range of ranges) {
    forEachStoredCell(sheet, range, (row, cell) => {
      if (cell.arrayRange) {
        arrays.push(cell.arrayRange);
      }
      setEdit(edits, row.index, cell.col, clearedCell(cell));
    });
  }
  clearMembers(sheet, arrays, edits, () => false);
  applyEdits(model, sheetIndex, edits);
}

// --- protection and arrays -------------------------------------------------------------

/** Whether a style is locked (`protection.locked` defaults to true). */
function isLocked(model: SpreadsheetModel, styleId: number): boolean {
  return cellStyle(model, styleId).protection?.locked !== false;
}

/** The first column of [left, right] whose column default (span style or
 * 0) is locked and that `skip` does not exclude, or null — walking the
 * column spans, never each column. */
function firstLockedColumn(
  model: SpreadsheetModel,
  sheet: Sheet,
  left: number,
  right: number,
  rowStyleId: number | undefined,
  ownStyled: ReadonlySet<number>,
): number | null {
  if (rowStyleId !== undefined) {
    if (!isLocked(model, rowStyleId)) {
      return null;
    }
    for (let col = left; col <= right; col += 1) {
      if (!ownStyled.has(col)) {
        return col;
      }
    }
    return null;
  }
  let col = left;
  while (col <= right) {
    const span = columnSpanAt(sheet.columns, col);
    // The run of columns sharing this default: the span, or the gap up to
    // the next span.
    const runEnd = span ? Math.min(span.max, right) : Math.min(nextSpanStart(sheet, col) - 1, right);
    if (isLocked(model, span?.styleId ?? 0)) {
      for (let c = col; c <= runEnd; c += 1) {
        if (!ownStyled.has(c)) {
          return c;
        }
      }
    }
    col = runEnd + 1;
  }
  return null;
}

/** The first column span starting after `col` (its min), or MAX_COLS. */
function nextSpanStart(sheet: Sheet, col: number): number {
  for (const span of sheet.columns) {
    if (span.min > col) {
      return span.min;
    }
  }
  return MAX_COLS;
}

/** The first locked position of `range` in a stored row, or null. */
function firstLockedInRow(model: SpreadsheetModel, sheet: Sheet, row: Row, range: GridRange): number | null {
  const ownStyled = new Set<number>();
  let best: number | null = null;
  for (let c = firstCellAtOrAfter(row.cells, range.left); c < row.cells.length && row.cells[c]!.col <= range.right; c += 1) {
    const cell = row.cells[c]!;
    if (cell.styleId === undefined) {
      continue;
    }
    ownStyled.add(cell.col);
    if (best === null && isLocked(model, cell.styleId)) {
      best = cell.col;
    }
  }
  const inherited = firstLockedColumn(model, sheet, range.left, best ?? range.right, row.styleId, ownStyled);
  if (inherited === null) {
    return best;
  }
  return best === null ? inherited : Math.min(best, inherited);
}

/** The first locked position (row-major, range by range) of `ranges` on a
 * protected sheet; null when the sheet is unprotected or every position is
 * unlocked. Positions are judged by their effective style, found from the
 * stored cells, the stored rows' styles and the column spans — a whole
 * column or the whole sheet costs what is stored, never its positions. */
export function firstLockedCell(model: SpreadsheetModel, sheetIndex: number, ranges: GridRange[]): CellRef | null {
  const sheet = model.sheets[sheetIndex];
  if (!sheet?.protection) {
    return null;
  }
  for (const range of ranges) {
    // Every unstored row behaves alike: the column defaults decide.
    const unstored = firstLockedColumn(model, sheet, range.left, range.right, undefined, new Set());
    let expected = range.top;
    const rows = sheet.rows;
    for (let r = firstRowAtOrAfter(rows, range.top); r < rows.length && rows[r]!.index <= range.bottom; r += 1) {
      const row = rows[r]!;
      if (row.index > expected && unstored !== null) {
        return { row: expected, col: unstored };
      }
      const col = firstLockedInRow(model, sheet, row, range);
      if (col !== null) {
        return { row: row.index, col };
      }
      expected = row.index + 1;
    }
    if (expected <= range.bottom && unstored !== null) {
      return { row: expected, col: unstored };
    }
  }
  return null;
}

/** Each `rows` array's array-formula ranges (anchors with `arrayRange`),
 * built once — arrays are replaced, never mutated. */
const arrayRanges = new WeakMap<readonly Row[], GridRange[]>();

function arraysOf(sheet: Sheet): GridRange[] {
  let found = arrayRanges.get(sheet.rows);
  if (!found) {
    found = [];
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (cell.arrayRange) {
          found.push(cell.arrayRange);
        }
      }
    }
    arrayRanges.set(sheet.rows, found);
  }
  return found;
}

/** Whether the union of `ranges` covers every cell of `target`. */
function covers(ranges: readonly GridRange[], target: GridRange): boolean {
  if (ranges.some((r) => r.top <= target.top && r.bottom >= target.bottom && r.left <= target.left && r.right >= target.right)) {
    return true;
  }
  for (let row = target.top; row <= target.bottom; row += 1) {
    const spans = ranges
      .filter((r) => r.top <= row && r.bottom >= row && r.right >= target.left && r.left <= target.right)
      .map((r) => [r.left, r.right] as const)
      .toSorted((a, b) => a[0] - b[0]);
    let next = target.left;
    for (const [left, right] of spans) {
      if (left > next) {
        return false;
      }
      next = Math.max(next, right + 1);
    }
    if (next <= target.right) {
      return false;
    }
  }
  return true;
}

/** The range of an array formula that `ranges` touch only in part — a cell
 * other than its anchor without the whole array — or null. Editing the
 * anchor alone, or clearing a whole array, is allowed. */
export function arrayRangeTouched(model: SpreadsheetModel, sheetIndex: number, ranges: GridRange[]): GridRange | null {
  const sheet = model.sheets[sheetIndex];
  if (!sheet) {
    return null;
  }
  for (const array of arraysOf(sheet)) {
    if (touchesBeyondAnchor(ranges, array) && !covers(ranges, array)) {
      return array;
    }
  }
  return null;
}
