// A batch of position edits to one sheet, applied copy-on-write (model.ts
// invariant 1) — the writer SPEC-125's paste, move and fill share, the
// position-level sibling of cells.ts's entry batch. Only touched rows are
// rebuilt and every untouched row and cell keeps its identity; a cell left
// storing nothing and a row left storing nothing (cells.ts isEmptyRow) are
// removed (invariant 2); a batch that changes nothing returns the sheet's
// own `rows` array. Row properties set through `setRowProps` replace the
// four modelled ones (height, customHeight, hidden, style) and drop the ods
// row style name, as a changed height does (SPEC-122's drop rule).
//
// cellWith() is the per-position rule every SPEC-125 write follows: a new
// value or formula applies SPEC-122's drop rules (a changed value drops
// `runs`, `sstIndex`, `preservedText` and ` vm="…"`; a changed formula drops
// `cm`, `dynamicArray`, `dataTableXml` and `formulaUnparsed`), preserved
// content the model does not model (notes, annotations) stays, and a style
// equal to the default 0 a position inherits anyway is left absent.
import { findRow, getCell, type GridRange } from "./address";
import { isEmptyRow, sameCellValue } from "./cells";
import { dropValueBound, type Cell, type Row, type Sheet } from "./model";
import { columnSpanAt } from "./styles";

export interface RowProps {
  heightPt?: number;
  customHeight?: true;
  hidden?: true;
  styleId?: number;
}

/** What a write puts at a position: a value change (`formula` absent: no
 * formula), a style (`"inherit"`: no style of its own), or both. */
export interface CellChange {
  content?: { value: Cell["value"]; formula?: string; arrayRange?: Cell["arrayRange"]; runs?: Cell["runs"] };
  style?: number | "inherit";
}

function deepEqual(a: unknown, b: unknown): boolean {
  if (a === b) {
    return true;
  }
  if (typeof a !== "object" || typeof b !== "object" || a === null || b === null || Array.isArray(a) !== Array.isArray(b)) {
    return false;
  }
  const ra = a as Record<string, unknown>;
  const rb = b as Record<string, unknown>;
  const ka = Object.keys(ra).filter((key) => ra[key] !== undefined);
  const kb = Object.keys(rb).filter((key) => rb[key] !== undefined);
  return ka.length === kb.length && ka.every((key) => deepEqual(ra[key], rb[key]));
}

/** A cell storing nothing: a null value and no other field. */
export function isEmptyCell(cell: Cell): boolean {
  if (cell.value !== null) {
    return false;
  }
  const record = cell as unknown as Record<string, unknown>;
  return Object.keys(record).every((key) => key === "col" || key === "value" || record[key] === undefined);
}

/** Removes what a changed or removed formula invalidates. */
export function dropFormulaBound(cell: Cell): void {
  delete cell.cm;
  delete cell.dynamicArray;
  delete cell.dataTableXml;
  delete cell.formulaUnparsed;
}

/** `existing` (or a new cell at `col`) with `change` applied (see the module
 * header); `inheritedId` is what the position shows without a style of its
 * own. The same object when nothing changes; undefined when it stores
 * nothing. */
export function cellWith(existing: Cell | undefined, col: number, change: CellChange, inheritedId: number): Cell | undefined {
  const next: Cell = existing ? { ...existing, col } : { col, value: null };
  const content = change.content;
  if (content) {
    const formulaChanged = content.formula !== existing?.formula || existing?.dataTableXml !== undefined || existing?.formulaUnparsed === true;
    const valueChanged = !sameCellValue(content.value, existing?.value ?? null);
    const runsChanged = !deepEqual(content.runs, existing?.runs);
    next.value = content.value;
    if (content.formula !== undefined) {
      next.formula = content.formula;
    } else {
      delete next.formula;
    }
    if (formulaChanged) {
      dropFormulaBound(next);
    }
    if (content.arrayRange !== undefined) {
      next.arrayRange = content.arrayRange;
    } else {
      delete next.arrayRange;
    }
    if (valueChanged || formulaChanged || runsChanged) {
      dropValueBound(next);
      if (content.runs !== undefined) {
        next.runs = content.runs;
      }
    }
  }
  if (change.style !== undefined) {
    if (change.style === "inherit" || (change.style === 0 && inheritedId === 0)) {
      delete next.styleId;
    } else {
      next.styleId = change.style;
    }
  }
  if (isEmptyCell(next)) {
    return undefined;
  }
  return existing && deepEqual(next, existing) ? existing : next;
}

/** A position's style without one of its own: its row's, else its column
 * span's, else 0. */
export function inheritedStyleId(sheet: Sheet, row: number, col: number): number {
  return findRow(sheet, row)?.styleId ?? columnSpanAt(sheet.columns, col)?.styleId ?? 0;
}

/** Rows sorted by index with `replaced` applied (undefined: removed) — a
 * new array, or `rows` itself when `replaced` is empty. */
export function withRows(rows: readonly Row[], replaced: ReadonlyMap<number, Row | undefined>): Row[] {
  if (replaced.size === 0) {
    return rows as Row[];
  }
  const keys = [...replaced.keys()].toSorted((a, b) => a - b);
  const out: Row[] = [];
  let k = 0;
  for (const row of rows) {
    while (k < keys.length && keys[k]! < row.index) {
      const added = replaced.get(keys[k]!);
      if (added) {
        out.push(added);
      }
      k += 1;
    }
    if (k < keys.length && keys[k] === row.index) {
      const next = replaced.get(row.index);
      if (next) {
        out.push(next);
      }
      k += 1;
    } else {
      out.push(row);
    }
  }
  for (; k < keys.length; k += 1) {
    const added = replaced.get(keys[k]!);
    if (added) {
      out.push(added);
    }
  }
  return out;
}

export class SheetEdit {
  private readonly pending = new Map<number, Map<number, Cell | undefined>>();
  private readonly rowProps = new Map<number, RowProps>();

  constructor(readonly sheet: Sheet) {}

  /** The cell at a position as the batch leaves it so far. */
  cell(row: number, col: number): Cell | undefined {
    const cols = this.pending.get(row);
    return cols?.has(col) ? cols.get(col) : getCell(this.sheet, row, col);
  }

  set(row: number, col: number, cell: Cell | undefined): void {
    let cols = this.pending.get(row);
    if (!cols) {
      cols = new Map();
      this.pending.set(row, cols);
    }
    cols.set(col, cell === undefined ? undefined : cell.col === col ? cell : { ...cell, col });
  }

  /** Applies `change` at a position (cellWith over the batch's cell). */
  change(row: number, col: number, change: CellChange): void {
    const existing = this.cell(row, col);
    const next = cellWith(existing, col, change, inheritedStyleId(this.sheet, row, col));
    if (next !== existing) {
      this.set(row, col, next);
    }
  }

  setRowProps(row: number, props: RowProps): void {
    this.rowProps.set(row, props);
  }

  /** The sheet's rows with the batch applied (see the module header). */
  rows(): Row[] {
    const replaced = new Map<number, Row | undefined>();
    for (const index of new Set([...this.pending.keys(), ...this.rowProps.keys()])) {
      const old = findRow(this.sheet, index);
      let next: Row = old ?? { index, cells: [] };
      const changes = this.pending.get(index);
      if (changes) {
        const cells = mergedCells(old?.cells ?? [], changes);
        if (cells !== (old?.cells ?? null)) {
          next = { ...next, cells };
        }
      }
      const props = this.rowProps.get(index);
      if (props && !sameProps(next, props)) {
        const patched: Row = { ...next };
        delete patched.heightPt;
        delete patched.customHeight;
        delete patched.hidden;
        delete patched.styleId;
        delete patched.sourceStyleName;
        next = { ...patched, ...props };
      }
      if (next === old || (!old && next.cells.length === 0 && !props)) {
        continue;
      }
      replaced.set(index, isEmptyRow(next) ? undefined : next);
    }
    return withRows(this.sheet.rows, replaced);
  }
}

function sameProps(row: Row, props: RowProps): boolean {
  return row.heightPt === props.heightPt && row.customHeight === props.customHeight && row.hidden === props.hidden && row.styleId === props.styleId;
}

/** `old` with `changes` applied: the same array when no cell changes. */
function mergedCells(old: readonly Cell[], changes: ReadonlyMap<number, Cell | undefined>): Cell[] {
  const effective: [number, Cell | undefined][] = [];
  const byCol = new Map(old.map((cell) => [cell.col, cell] as const));
  for (const [col, cell] of changes) {
    if (byCol.get(col) !== cell) {
      effective.push([col, cell]);
    }
  }
  if (effective.length === 0) {
    return old as Cell[];
  }
  effective.sort((a, b) => a[0] - b[0]);
  const out: Cell[] = [];
  let i = 0;
  for (const [col, cell] of effective) {
    while (i < old.length && old[i]!.col < col) {
      out.push(old[i]!);
      i += 1;
    }
    if (i < old.length && old[i]!.col === col) {
      i += 1;
    }
    if (cell) {
      out.push(cell);
    }
  }
  while (i < old.length) {
    out.push(old[i]!);
    i += 1;
  }
  return out;
}

/** Whether a range lies wholly inside another. */
export function contains(outer: GridRange, inner: GridRange): boolean {
  return outer.top <= inner.top && outer.left <= inner.left && outer.bottom >= inner.bottom && outer.right >= inner.right;
}
