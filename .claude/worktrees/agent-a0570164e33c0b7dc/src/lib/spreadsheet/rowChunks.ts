// Row chunks (SPEC-137 "The storage layout", "The parse protocol"): a sheet's
// rows are cut into bands of SHEET_ROWS_CHUNK row INDICES — chunk n holds the
// rows whose index lies in [n·512, n·512 + 512) — so the off-thread reader
// hands them to the main thread one band per message and the store keeps one
// record per band. A band is a range of row indices, not of array positions:
// inserting one sparse row changes the one band it lands in, never every band
// after it. Kept free of every reader and of the store so both sides of the
// worker boundary, and the store, can import it without the codecs.
import type { Row } from "./model";

/** Rows per chunk: at a 20-column sheet a chunk is ~10,000 cells — a few
 * milliseconds of structured clone — and 49 records per 25,000 rows. */
export const SHEET_ROWS_CHUNK = 512;

/** The chunk a row index falls in. */
export function chunkOf(rowIndex: number): number {
  return Math.floor(rowIndex / SHEET_ROWS_CHUNK);
}

/** A chunk index as a record key's last segment: zero-padded to four digits
 * (the grid's 1,048,576 rows need at most 2,048 chunks), so the keys of one
 * sheet sort in chunk order. */
export function chunkKeySegment(chunk: number): string {
  return String(chunk).padStart(4, "0");
}

/** The position of the first row whose index is at least `index` in rows
 * sorted ascending by index (rows.length when there is none). */
export function firstRowAtOrAfter(rows: readonly Row[], index: number): number {
  let low = 0;
  let high = rows.length;
  while (low < high) {
    const mid = (low + high) >>> 1;
    if (rows[mid]!.index < index) {
      low = mid + 1;
    } else {
      high = mid;
    }
  }
  return low;
}

/** The rows of chunk `chunk` — the same Row objects, in order (an empty array
 * when the band holds none). */
export function rowsInChunk(rows: readonly Row[], chunk: number): Row[] {
  const from = firstRowAtOrAfter(rows, chunk * SHEET_ROWS_CHUNK);
  const to = firstRowAtOrAfter(rows, (chunk + 1) * SHEET_ROWS_CHUNK);
  return rows.slice(from, to);
}

/** Every non-empty chunk of `rows` (sorted ascending by index), in chunk
 * order: [chunk index, its rows]. */
export function chunkRows(rows: readonly Row[]): [number, Row[]][] {
  const out: [number, Row[]][] = [];
  let at = 0;
  while (at < rows.length) {
    const chunk = chunkOf(rows[at]!.index);
    const end = firstRowAtOrAfter(rows, (chunk + 1) * SHEET_ROWS_CHUNK);
    out.push([chunk, rows.slice(at, end)]);
    at = end;
  }
  return out;
}
