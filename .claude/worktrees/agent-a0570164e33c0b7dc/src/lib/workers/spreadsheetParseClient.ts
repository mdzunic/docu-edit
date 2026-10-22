// Main-thread client for the workbook parse worker (SPEC-137 "The parse
// protocol"). SKELETON: the off-thread protocol lands with the read phases;
// until then this reads inline through readSpreadsheetBytes so every caller
// can already be wired to it.
import type { SpreadsheetFormat, SpreadsheetModel } from "../spreadsheet/model";
import { readSpreadsheetBytes } from "../spreadsheet/open";

export interface OffThreadReadOptions {
  /** The open's progress, 0..1, as the rows arrive (chunks pulled / chunk count). */
  onProgress?: (fraction: number) => void;
}

/** Reads an inflated workbook package into a new SpreadsheetModel, parsing
 * its O(cells) parts off the main thread. */
export async function readSpreadsheetOffThread(
  parts: Map<string, Uint8Array>,
  format: SpreadsheetFormat,
  title: string,
  options: OffThreadReadOptions = {},
): Promise<SpreadsheetModel> {
  const model = await readSpreadsheetBytes(new Uint8Array(0), format, title, parts);
  options.onProgress?.(1);
  return model;
}
