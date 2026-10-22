// The workbook read's phases (SPEC-137 "The read phases"): the plain-data
// types every phase hands to the next, on both sides of the parse worker's
// boundary, and the one piece of main-thread state a shell phase leaves for
// assembly. open.ts runs the phases (inline, or with phases 1 and 3 in
// workers/spreadsheetParse.worker.ts through parseSession.ts); each format's
// module (xlsx/phases.ts, ods/phases.ts) implements them.
//
// 1. prepareWorkbookSources — DOMParser-free: the O(cells) source texts
//    (xlsx: the shared strings; ods: content.xml cut into its envelope and
//    each table's offsets).
// 2. readWorkbookShell — main thread, may use parseXml: every small part,
//    the workbook's shell (sheets with no rows yet, ids minted here, styles,
//    the retained `opaque`), one SheetScanJob per sheet to scan, and the
//    ScanContext those scans need (plain JSON: it is posted to the worker).
// 3. scanSheet — DOMParser-free: one worksheet part or one ods table, ALL of
//    it, emitting its rows as RowChunks of SHEET_ROWS_CHUNK row indices and
//    returning everything else in `fields`. What only the main thread can
//    finish (a part the sheet's relationships name and that must be parsed
//    with DOMParser — a drawing, a chart …) travels back as `deferred` plain
//    data; tables the scan grows (implied-format styles, x14 dxfs, ods
//    styles in first-reference order) travel back as `appended`.
// 4. assembleWorkbook — main thread: the shell, every result and every chunk
//    into the finished model, then the format's completion (setShellCompletion).
import type { Dxf } from "./conditionalFormats";
import type { CellStyle, Row, Sheet, SpreadsheetFormat, SpreadsheetModel } from "./model";
import type { OdsScanContext, OdsScanDeferred } from "./ods/phases";
import type { SharedStringTable } from "./xlsx/sharedStrings";
import type { XlsxScanContext, XlsxScanDeferred } from "./xlsx/phases";

/** Phase 1's output — plain data, DOMParser-free. The parse worker keeps
 * `sharedStrings` for its session and posts the rest back. */
export interface WorkbookSources {
  format: SpreadsheetFormat;
  /** xlsx: the shared-string table, scanned once. */
  sharedStrings?: SharedStringTable;
  /** ods: content.xml before the first table and after the last. */
  odsEnvelope?: { prefix: string; suffix: string };
  /** ods: the UTF-16 offsets of each table's `sourceXml` in content.xml. */
  odsTables?: { start: number; end: number }[];
}

/** One sheet to scan: the model sheet it fills, the source part it reads
 * and, for an ods table, its offsets in content.xml. */
export interface SheetScanJob {
  sheetIndex: number;
  partName: string;
  range?: { start: number; end: number };
}

/** What phase 3 needs from phase 2: plain JSON (the date system, the style
 * tables, the sheet names for messages …), discriminated by `format`. A
 * later spec that must resolve something per cell against a workbook-level
 * table adds it here; it never parses a part with DOMParser in phase 3. */
export type ScanContext = XlsxScanContext | OdsScanContext;

/** Rows with index in [chunk·512, chunk·512 + 512) of one sheet. */
export interface RowChunk {
  sheetIndex: number;
  chunk: number;
  rows: Row[];
}

/** One sheet's scan: everything but its rows. */
export interface SheetScanResult {
  sheetIndex: number;
  fields: Partial<Sheet>;
  /** The cells this sheet counted against MAX_WORKBOOK_CELLS. */
  cells: number;
  /** Entries the scan appended to the context's growing tables, in order —
   * assembleWorkbook appends them to the model's, result by result. */
  appended?: { styles?: CellStyle[]; dxfs?: Dxf[] };
  /** What only the main thread can finish (see the header), plain data. */
  deferred?: XlsxScanDeferred | OdsScanDeferred;
}

/** The cells every job of one read may still store (MAX_WORKBOOK_CELLS at
 * the start, counted across all sheets); a scan that would pass it refuses
 * the workbook with WORKBOOK_TOO_LARGE_MESSAGE. */
export interface ScanBudget {
  cellsLeft: number;
}

/** Phase 3's stepwise form: each step() scans at most `rows` more rows and
 * answers whether the sheet is done — so an inline open can yield to the
 * browser between steps (SCAN_YIELD_ROWS). scanSheet is the one-step use. */
export interface SheetScanSteps {
  step(rows: number): boolean;
  result(): SheetScanResult;
}

/** A shell phase's main-thread completion: given the assembled model (rows,
 * fields and appended tables in place) and the scan results, it finishes
 * what needs the package or DOMParser — drawings, tables, notes, the active
 * sheet, the preview … — in place. */
export type ShellCompletion = (model: SpreadsheetModel, results: readonly SheetScanResult[]) => void;

const completions = new WeakMap<SpreadsheetModel, ShellCompletion>();

/** Called by a format's readShell: what assembleWorkbook runs for `shell`
 * (the parts, the style book … stay in the closure, on the main thread). */
export function setShellCompletion(shell: SpreadsheetModel, complete: ShellCompletion): void {
  completions.set(shell, complete);
}

/** The completion registered for `shell`, removed — a shell assembles once. */
export function takeShellCompletion(shell: SpreadsheetModel): ShellCompletion | undefined {
  const complete = completions.get(shell);
  completions.delete(shell);
  return complete;
}
