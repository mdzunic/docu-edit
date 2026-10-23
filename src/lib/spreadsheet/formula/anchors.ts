// Which formula covers a position that holds none of its own (SPEC-119
// "References", "The formula bar"): the anchor of a legacy or dynamic array
// (`Cell.arrayRange`) and of a What-If data table (the range its retained
// `<f t="dataTable" ref="…">` names). Built once per `sheet.rows` array —
// the copy-on-write invariant makes that array's identity the sheet's
// version — and kept in a WeakMap: lib state, never model data. Lookups go
// through merges.ts's index (array and table ranges never overlap in a valid
// file, as merges never do), so a position costs a binary search.
// SPEC-137: a sheet's new rows array carries its previous array's index
// forward when none of the rows that differ between the two (by identity,
// on either side) holds an anchor — every anchor then sits in a row both
// arrays share, the same cell at the same place — so an edit or a
// recalculation slice re-reads the rows it touched, not the whole sheet. The
// previous array is held weakly: a released workbook's rows stay
// collectable.
import { windowOf } from "../../core/arrayWindow";
import { parseRange, type GridRange } from "../address";
import { buildMergeIndex, mergeAt, type MergeIndex } from "../merges";
import type { Cell, Row, Sheet } from "../model";

/** A range filled by one formula, and the anchor cell holding it. */
export interface FormulaAnchor {
  kind: "array" | "dataTable";
  range: GridRange;
  cell: Cell;
}

interface AnchorIndex {
  index: MergeIndex;
  anchors: Map<GridRange, FormulaAnchor>;
}

const indexes = new WeakMap<readonly Row[], AnchorIndex>();
/** Each sheet's last indexed rows array, by sheet id (see the header). */
const latest = new Map<string, WeakRef<Row[]>>();

const ATTRIBUTE = /\s([\w:.-]+)\s*=\s*(["'])(.*?)\2/g;

/** An attribute of the start tag at the head of `xml`, entity-free values
 * only (a data table's attributes are addresses and flags). */
export function startTagAttribute(xml: string, name: string): string | undefined {
  const tagEnd = xml.indexOf(">");
  const tag = tagEnd < 0 ? xml : xml.slice(0, tagEnd);
  for (const match of tag.matchAll(ATTRIBUTE)) {
    if (match[1] === name) {
      return match[3];
    }
  }
  return undefined;
}

function buildIndex(sheet: Sheet): AnchorIndex {
  const anchors = new Map<GridRange, FormulaAnchor>();
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      if (cell.arrayRange) {
        anchors.set(cell.arrayRange, { kind: "array", range: cell.arrayRange, cell });
      } else if (cell.dataTableXml !== undefined) {
        const range = parseRange(startTagAttribute(cell.dataTableXml, "ref")?.trim() ?? "") ?? {
          top: row.index,
          left: cell.col,
          bottom: row.index,
          right: cell.col,
        };
        anchors.set(range, { kind: "dataTable", range, cell });
      }
    }
  }
  return { index: buildMergeIndex([...anchors.keys()]), anchors };
}

/** Whether a row holds an array's or a data table's anchor. */
function holdsAnchor(row: Row): boolean {
  return row.cells.some((cell) => cell.arrayRange !== undefined || cell.dataTableXml !== undefined);
}

/** The index of `sheet.rows`: cached, carried forward from the sheet's
 * previous rows array, or built (see the header). */
function anchorIndex(sheet: Sheet): AnchorIndex {
  let built = indexes.get(sheet.rows);
  if (built) {
    return built;
  }
  const previous = latest.get(sheet.id)?.deref();
  const carried = previous ? indexes.get(previous) : undefined;
  if (previous && carried) {
    const { beforeMid, afterMid } = windowOf(previous, sheet.rows);
    built = beforeMid.some(holdsAnchor) || afterMid.some(holdsAnchor) ? undefined : carried;
  }
  built ??= buildIndex(sheet);
  indexes.set(sheet.rows, built);
  latest.set(sheet.id, new WeakRef(sheet.rows));
  return built;
}

/** The array or data table whose range holds (row, col), if any. */
export function formulaAnchorAt(sheet: Sheet, row: number, col: number): FormulaAnchor | undefined {
  const built = anchorIndex(sheet);
  if (built.anchors.size === 0) {
    return undefined;
  }
  const range = mergeAt(built.index, row, col);
  return range ? built.anchors.get(range) : undefined;
}

/** The cell whose `arrayRange` contains (row, col) — the anchor itself for
 * the anchor's own position. */
export function findArrayAnchor(sheet: Sheet, row: number, col: number): Cell | undefined {
  const anchor = formulaAnchorAt(sheet, row, col);
  return anchor?.kind === "array" ? anchor.cell : undefined;
}
