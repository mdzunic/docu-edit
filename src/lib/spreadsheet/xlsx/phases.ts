// The .xlsx reader's phases (SPEC-137 "The read phases"; the types and the
// rules are readPhases.ts's). SKELETON — the implementation replaces every
// body below; the signatures are the contract open.ts, parseSession.ts and
// the parse client build on.
import type { SpreadsheetModel } from "../model";
import type { RowChunk, ScanBudget, SheetScanJob, SheetScanSteps, WorkbookSources } from "../readPhases";

/** What a worksheet scan needs from the shell phase — plain JSON. */
export interface XlsxScanContext {
  format: "xlsx";
}

/** What a worksheet scan leaves for the main thread to finish — plain JSON. */
export interface XlsxScanDeferred {
  format: "xlsx";
}

/** The parts phase 1 and phase 3 read (main thread, before phase 1 — the
 * parse client posts copies of exactly these to the worker). */
export function xlsxSourcePartNames(_parts: Map<string, Uint8Array>): string[] {
  throw new Error("SPEC-137: xlsx phases not implemented");
}

/** Phase 1 (DOMParser-free). */
export function prepareXlsxSources(_texts: Map<string, string>): WorkbookSources {
  throw new Error("SPEC-137: xlsx phases not implemented");
}

/** Phase 2 (main thread). */
export async function readXlsxShell(
  _parts: Map<string, Uint8Array>,
  _sources: WorkbookSources,
  _title: string,
): Promise<{ shell: SpreadsheetModel; jobs: SheetScanJob[]; context: XlsxScanContext }> {
  throw new Error("SPEC-137: xlsx phases not implemented");
}

/** Phase 3 (DOMParser-free), stepwise. */
export function startXlsxSheetScan(
  _job: SheetScanJob,
  _text: string,
  _sources: WorkbookSources,
  _context: XlsxScanContext,
  _emit: (chunk: RowChunk) => void,
  _budget: ScanBudget,
): SheetScanSteps {
  throw new Error("SPEC-137: xlsx phases not implemented");
}
