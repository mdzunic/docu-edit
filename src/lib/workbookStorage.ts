// SPEC-137 "The storage layout": how a workbook lives in the `documents`
// store. One record per SHEET_ROWS_CHUNK-row band of each sheet's row
// INDICES (`sheetrows:<docId>:<sheetId>:<chunk4>` — SPEC-116's per-sheet key
// gains one segment, so a 116-era key has three segments and a chunk key
// four), one source record (`sheetsource:<docId>`: the retained package, the
// ods envelope and every sheet's `sourceXml`), and a slim shell that holds no
// cell, row or source byte and carries the `sheetStore` manifest. Pure
// functions over the model — the keys, the shell, one persist's write plan
// (which chunks changed by identity, which became empty, which sheets went
// away, whether the sources changed) and the storedBytes estimate;
// documentStore owns the state (residency, the latches) and runs the I/O.
import { windowOf } from "./core/arrayWindow";
import type { WriteBatchOp } from "./core/idb";
import type { OpaquePart } from "./core/packageParts";
import type { Cell, Row, Sheet, SpreadsheetModel } from "./spreadsheet/model";
import { chunkKeySegment, chunkOf, chunkRows, rowsInChunk, SHEET_ROWS_CHUNK } from "./spreadsheet/rowChunks";

export { SHEET_ROWS_CHUNK };
export const SHEET_ROWS_KEY_PREFIX = "sheetrows:";
export const SHEET_SOURCE_KEY_PREFIX = "sheetsource:";

/** Above every key that starts with a prefix: no key holds U+FFFF (document
 * ids are UUIDs, sheet ids `[A-Za-z0-9_-]`). */
const KEY_END = "￿";

/** One band of one sheet's rows: the rows whose index lies in
 * [chunk·512, chunk·512 + 512), in order. */
export interface SheetRowsChunkRecord {
  /** `sheetrows:${docId}:${sheetId}:${chunkKeySegment(chunk)}` */
  id: string;
  docId: string;
  sheetId: string;
  chunk: number;
  rows: Row[];
}

/** A workbook's retained sources, apart from the shell so the start-up read
 * never clones the package: `opaque`, the ods envelope and every sheet's
 * `sourceXml` (by sheet id). */
export interface WorkbookSourceRecord {
  /** `sheetsource:${docId}` */
  id: string;
  docId: string;
  opaque: OpaquePart[];
  odsContentEnvelope?: { prefix: string; suffix: string };
  sheetSources: Record<string, string>;
}

/** The shell record's manifest (on the stored record, not the model): the
 * chunk indices storage holds per sheet id, and the estimated bytes of the
 * whole workbook as last persisted. */
export interface SheetStoreManifest {
  chunks: Record<string, number[]>;
  storedBytes: number;
}

/** SPEC-116's per-sheet rows record (`sheetrows:${docId}:${sheetId}` — three
 * key segments). Read only to migrate a 116-era workbook. */
export interface SheetRowsRecord {
  id: string;
  docId: string;
  sheetId: string;
  rows: Row[];
}

export function sheetRowsChunkKey(docId: string, sheetId: string, chunk: number): string {
  return `${SHEET_ROWS_KEY_PREFIX}${docId}:${sheetId}:${chunkKeySegment(chunk)}`;
}

export function legacySheetRowsKey(docId: string, sheetId: string): string {
  return `${SHEET_ROWS_KEY_PREFIX}${docId}:${sheetId}`;
}

export function workbookSourceKey(docId: string): string {
  return SHEET_SOURCE_KEY_PREFIX + docId;
}

/** Every rows record of a workbook — its chunks and any SPEC-116 per-sheet
 * record alike — as a closed key range. */
export function workbookRowsRange(docId: string): [string, string] {
  return [`${SHEET_ROWS_KEY_PREFIX}${docId}:`, `${SHEET_ROWS_KEY_PREFIX}${docId}:${KEY_END}`];
}

/** Every chunk of one sheet as a closed key range (the trailing ":" keeps
 * sheet "a" from matching sheet "ab"). */
export function sheetRowsRange(docId: string, sheetId: string): [string, string] {
  return [`${SHEET_ROWS_KEY_PREFIX}${docId}:${sheetId}:`, `${SHEET_ROWS_KEY_PREFIX}${docId}:${sheetId}:${KEY_END}`];
}

/** The retained sources a persist compares by reference (strings by value):
 * the source record is rewritten only when one of them changed. */
export interface SourceLatch {
  opaque: OpaquePart[];
  envelope: { prefix: string; suffix: string } | undefined;
  sheetSources: Map<string, string>;
}

/** What storage holds for a resident workbook: per sheet id, the rows array
 * whose chunks it holds (so the last persisted slice of chunk k is
 * rowsInChunk(rows, k) — the per-chunk latch), and the sources. */
export interface WorkbookLatch {
  rows: Map<string, Row[]>;
  source: SourceLatch | null;
}

export function emptyLatch(): WorkbookLatch {
  return { rows: new Map(), source: null };
}

export function sourceLatchOf(model: SpreadsheetModel): SourceLatch {
  const sheetSources = new Map<string, string>();
  for (const sheet of model.sheets) {
    if (sheet.sourceXml !== undefined) {
      sheetSources.set(sheet.id, sheet.sourceXml);
    }
  }
  return { opaque: model.opaque, envelope: model.odsContentEnvelope, sheetSources };
}

export function sameSource(a: SourceLatch, b: SourceLatch): boolean {
  if (a.opaque !== b.opaque || a.sheetSources.size !== b.sheetSources.size) {
    return false;
  }
  if (a.envelope !== b.envelope && (a.envelope?.prefix !== b.envelope?.prefix || a.envelope?.suffix !== b.envelope?.suffix)) {
    return false;
  }
  for (const [sheetId, xml] of a.sheetSources) {
    if (b.sheetSources.get(sheetId) !== xml) {
      return false;
    }
  }
  return true;
}

/** Whether the latch still describes `model`: the same sheets, each with the
 * rows array storage holds, and the same sources. */
export function latchMatches(latch: WorkbookLatch, model: SpreadsheetModel): boolean {
  return (
    latch.source !== null &&
    latch.rows.size === model.sheets.length &&
    model.sheets.every((sheet) => latch.rows.get(sheet.id) === sheet.rows) &&
    sameSource(latch.source, sourceLatchOf(model))
  );
}

/** A sheet without its rows and its source text. */
function slimSheet(sheet: Sheet): Sheet {
  const slim: Sheet = { ...sheet, rows: [] };
  delete slim.sourceXml;
  return slim;
}

/** The slim shell (SPEC-137): every sheet's `rows: []`, `opaque: []`, no
 * `sourceXml`, no ods envelope — what the shell record stores and what a
 * workbook that is not resident keeps in memory. A new object; `model` is
 * untouched. */
export function workbookShell(model: SpreadsheetModel): SpreadsheetModel {
  const shell: SpreadsheetModel = { ...model, opaque: [], sheets: model.sheets.map(slimSheet) };
  delete shell.odsContentEnvelope;
  return shell;
}

/** A SPEC-116-era shell: rows emptied, the sources kept inline (that layout
 * has no source record to hold them). */
export function legacyWorkbookShell(model: SpreadsheetModel): SpreadsheetModel {
  return { ...model, sheets: model.sheets.map((sheet) => ({ ...sheet, rows: [] })) };
}

/** The whole workbook from its shell, each sheet's rows and the source
 * record (absent: no sources — an incomplete load). */
export function joinWorkbook(
  shell: SpreadsheetModel,
  rows: ReadonlyMap<string, Row[]>,
  source: WorkbookSourceRecord | undefined,
): SpreadsheetModel {
  const model: SpreadsheetModel = {
    ...shell,
    opaque: source?.opaque ?? [],
    sheets: shell.sheets.map((sheet) => {
      const joined: Sheet = { ...sheet, rows: rows.get(sheet.id) ?? [] };
      const xml = source?.sheetSources[sheet.id];
      if (xml !== undefined) {
        joined.sourceXml = xml;
      }
      return joined;
    }),
  };
  if (source?.odsContentEnvelope) {
    model.odsContentEnvelope = source.odsContentEnvelope;
  } else {
    delete model.odsContentEnvelope;
  }
  return model;
}

function sourceRecordOf(docId: string, model: SpreadsheetModel): WorkbookSourceRecord {
  const record: WorkbookSourceRecord = {
    id: workbookSourceKey(docId),
    docId,
    opaque: model.opaque,
    sheetSources: Object.fromEntries(sourceLatchOf(model).sheetSources),
  };
  if (model.odsContentEnvelope) {
    record.odsContentEnvelope = model.odsContentEnvelope;
  }
  return record;
}

function sameRows(a: readonly Row[], b: readonly Row[]): boolean {
  return a.length === b.length && a.every((row, index) => row === b[index]);
}

/** The chunks of one sheet a persist must look at. With no latch for the
 * sheet (never persisted, or its layout unknown): every chunk it has and
 * every chunk storage holds. Otherwise windowOf trims the reference-equal
 * head and tail (copy-on-write: an untouched row keeps its identity), the
 * window's row-INDEX span maps to a run of chunks, and of those only the ones
 * whose slice differs by identity from the last persisted slice remain. */
function chunksToCheck(before: Row[] | undefined, after: Row[], stored: readonly number[]): number[] {
  if (before === undefined) {
    return [...new Set([...chunkRows(after).map(([chunk]) => chunk), ...stored])].toSorted((a, b) => a - b);
  }
  const window = windowOf(before, after);
  const first = Math.min(window.beforeMid[0]?.index ?? Infinity, window.afterMid[0]?.index ?? Infinity);
  if (first === Infinity) {
    return [];
  }
  const last = Math.max(window.beforeMid.at(-1)?.index ?? -1, window.afterMid.at(-1)?.index ?? -1);
  const changed: number[] = [];
  for (let chunk = chunkOf(first); chunk <= chunkOf(last); chunk += 1) {
    if (!sameRows(rowsInChunk(before, chunk), rowsInChunk(after, chunk))) {
      changed.push(chunk);
    }
  }
  return changed;
}

/** One persist of a resident workbook: the batch operations before the
 * shell (changed chunks, deletions, the source record), and what storage
 * will hold once they commit. */
export interface WorkbookWritePlan {
  ops: WriteBatchOp[];
  latch: WorkbookLatch;
  chunks: Record<string, number[]>;
  /** A chunk was written or deleted — the card preview is rebuilt. */
  rowsChanged: boolean;
}

/** Plans the persist of `model` against what storage holds: `manifest` and
 * `latch` (ignored for a `legacy` record — SPEC-116's layout, whose
 * per-sheet records the plan deletes first and replaces wholesale). */
export function planWorkbookWrite(
  docId: string,
  model: SpreadsheetModel,
  stored: { manifest: SheetStoreManifest | undefined; latch: WorkbookLatch; legacy: boolean },
): WorkbookWritePlan {
  const ops: WriteBatchOp[] = [];
  const manifest = stored.legacy ? undefined : stored.manifest;
  const latch = manifest ? stored.latch : emptyLatch();
  if (stored.legacy) {
    ops.push({ deleteRange: workbookRowsRange(docId) });
  }
  const chunks: Record<string, number[]> = {};
  const rows = new Map<string, Row[]>();
  let rowsChanged = false;
  for (const sheet of model.sheets) {
    rows.set(sheet.id, sheet.rows);
    const storedChunks = manifest?.chunks[sheet.id] ?? [];
    const before = latch.rows.get(sheet.id);
    if (before === sheet.rows) {
      chunks[sheet.id] = storedChunks;
      continue;
    }
    const kept = new Set(storedChunks);
    for (const chunk of chunksToCheck(before, sheet.rows, storedChunks)) {
      const slice = rowsInChunk(sheet.rows, chunk);
      if (slice.length > 0) {
        const record: SheetRowsChunkRecord = { id: sheetRowsChunkKey(docId, sheet.id, chunk), docId, sheetId: sheet.id, chunk, rows: slice };
        ops.push({ put: record });
        kept.add(chunk);
        rowsChanged = true;
      } else if (kept.delete(chunk)) {
        ops.push({ delete: sheetRowsChunkKey(docId, sheet.id, chunk) });
        rowsChanged = true;
      }
    }
    chunks[sheet.id] = [...kept].toSorted((a, b) => a - b);
  }
  for (const sheetId of Object.keys(manifest?.chunks ?? {})) {
    if (!rows.has(sheetId)) {
      ops.push({ deleteRange: sheetRowsRange(docId, sheetId) });
      rowsChanged = true;
    }
  }
  const source = sourceLatchOf(model);
  if (!manifest || latch.source === null || !sameSource(latch.source, source)) {
    ops.push({ put: sourceRecordOf(docId, model) });
  }
  return { ops, latch: { rows, source }, chunks, rowsChanged };
}

// --- storedBytes (SPEC-137: estimated at persist time, never serialised) ---

/** 16 bytes per cell plus its strings' lengths. */
function cellBytes(cell: Cell): number {
  let bytes = 16;
  if (typeof cell.value === "string") {
    bytes += cell.value.length;
  }
  bytes += (cell.formula?.length ?? 0) + (cell.preservedAttrs?.length ?? 0) + (cell.preservedXml?.length ?? 0);
  bytes += (cell.preservedText?.length ?? 0) + (cell.dataTableXml?.length ?? 0);
  if (cell.runs) {
    for (const run of cell.runs) {
      bytes += run.text.length;
    }
  }
  return bytes;
}

// Rows are copy-on-write, so a row's (and a rows array's) estimate never
// changes: after an edit only the touched rows are walked again.
const rowBytesCache = new WeakMap<Row, number>();
const rowsBytesCache = new WeakMap<readonly Row[], number>();

function rowBytes(row: Row): number {
  let bytes = rowBytesCache.get(row);
  if (bytes === undefined) {
    bytes = (row.sourceAttrs?.length ?? 0) + (row.sourceStyleName?.length ?? 0);
    for (const cell of row.cells) {
      bytes += cellBytes(cell);
    }
    rowBytesCache.set(row, bytes);
  }
  return bytes;
}

/** A sheet's chunks' estimated size. */
export function rowsStoredBytes(rows: readonly Row[]): number {
  let bytes = rowsBytesCache.get(rows);
  if (bytes === undefined) {
    bytes = 0;
    for (const row of rows) {
      bytes += rowBytes(row);
    }
    rowsBytesCache.set(rows, bytes);
  }
  return bytes;
}

/** The source record's estimated size: its base64 and source strings. */
export function sourceStoredBytes(model: SpreadsheetModel): number {
  let bytes = (model.odsContentEnvelope?.prefix.length ?? 0) + (model.odsContentEnvelope?.suffix.length ?? 0);
  for (const part of model.opaque) {
    bytes += part.name.length + part.data.length;
  }
  for (const sheet of model.sheets) {
    bytes += sheet.sourceXml?.length ?? 0;
  }
  return bytes;
}

/** The shell's estimated size, by the same rule: 16 bytes per sheet, style
 * and preview cell plus their strings. */
function shellStoredBytes(model: SpreadsheetModel): number {
  let bytes = 16 * (model.sheets.length + model.styles.length) + model.title.length;
  for (const sheet of model.sheets) {
    bytes += sheet.name.length;
  }
  for (const line of model.preview?.cells ?? []) {
    for (const text of line) {
      bytes += 16 + text.length;
    }
  }
  return bytes;
}

/** The whole workbook's estimated stored size — what the manifest's
 * `storedBytes` records (for a shell, only what it holds in memory). */
export function workbookStoredBytes(model: SpreadsheetModel): number {
  let bytes = sourceStoredBytes(model) + shellStoredBytes(model);
  for (const sheet of model.sheets) {
    bytes += rowsStoredBytes(sheet.rows);
  }
  return bytes;
}
