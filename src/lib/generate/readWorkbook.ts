// Byte-based workbook reading (SPEC-140), mirroring read.ts and readDeck.ts:
// generated (or any) `.xlsx`/`.ods` bytes go straight back into a
// SpreadsheetModel with no `File`, no disk and no browser. The format is the
// package's own declaration — unzip, core's detectPackage and SPEC-116's
// spreadsheetFormatOf, never a file name — and SPEC-137's
// readSpreadsheetBytes reads it, the same assembly the app's open path uses —
// both behind nodeDom.ts's DOMParser polyfill (package detection and the
// readers need one outside a browser; writing never does).
import { detectPackage, spreadsheetFormatOf } from "../core/packageDetect";
import { unzip } from "../core/zip";
import { SpreadsheetOpenError } from "../spreadsheet/errors";
import { DEFAULT_SPREADSHEET_TITLE, type SpreadsheetModel } from "../spreadsheet/model";
import { readSpreadsheetBytes } from "../spreadsheet/open";
import { DocumentGenerationError } from "./errors";
import { ensureNodeXmlGlobals } from "./nodeDom";

export const NOT_A_WORKBOOK_MESSAGE = "bytes: not an .xlsx or .ods workbook";

/** Reads `.xlsx` or `.ods` bytes (either — the package says which) into a
 * SpreadsheetModel, in the browser or in Node. `title` defaults the way a
 * new workbook's does. Anything that is not such a workbook — not a ZIP, a
 * document or deck package, an `.xlsb` — is a DocumentGenerationError. */
export async function readWorkbook(bytes: Uint8Array | ArrayBuffer, title: string = DEFAULT_SPREADSHEET_TITLE): Promise<SpreadsheetModel> {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  let parts: Map<string, Uint8Array>;
  try {
    parts = await unzip(data);
  } catch {
    throw new DocumentGenerationError(NOT_A_WORKBOOK_MESSAGE);
  }
  // detectPackage reads an OPC package's relationships with DOMParser too.
  await ensureNodeXmlGlobals();
  const declared = spreadsheetFormatOf(detectPackage(parts));
  if (declared === null || declared === "xlsb") {
    throw new DocumentGenerationError(NOT_A_WORKBOOK_MESSAGE);
  }
  try {
    return await readSpreadsheetBytes(data, declared.format, title, parts);
  } catch (error) {
    if (error instanceof SpreadsheetOpenError) {
      throw new DocumentGenerationError(`bytes: ${error.message}`);
    }
    throw error;
  }
}
