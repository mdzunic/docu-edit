// Download a workbook (SPEC-116) — the spreadsheet twin of
// lib/downloadDocument.ts, presentation/download.ts and media/download.ts,
// kept framework-free. Nothing edits a workbook yet, so there is no writer:
// the download re-emits every source part byte-identical, in source order
// (the ZIP container itself is rebuilt — round trips compare parts, as the
// presentation suites do). An `.ods` keeps `mimetype` first and STOREd, as
// ODF requires; a macro-enabled workbook keeps `xl/vbaProject.bin` and its
// content types because every part goes back out. The file name's extension
// follows format + package variant (fileTypes.ts).
import { decodedPartBytes } from "../core/packageParts";
import { triggerBlobDownload } from "../core/triggerDownload";
import { zip, type ZipEntryInput } from "../core/zipWriter";
import { spreadsheetFileExtension, spreadsheetMimeType } from "./fileTypes";
import { DEFAULT_SPREADSHEET_TITLE, type SpreadsheetModel } from "./model";

/** The file name a workbook downloads under: its title plus the extension
 * of its format and variant (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.ods`,
 * `.ots`). */
export function spreadsheetFileName(model: SpreadsheetModel): string {
  const title = model.title.trim();
  return `${title.length > 0 ? title : DEFAULT_SPREADSHEET_TITLE}.${spreadsheetFileExtension(model.format, model.packageVariant)}`;
}

/** The bytes, name and MIME type of a workbook's download. */
export async function spreadsheetDownload(
  model: SpreadsheetModel,
): Promise<{ bytes: Uint8Array; fileName: string; mimeType: string }> {
  const parts = decodedPartBytes(model.opaque);
  const entries: ZipEntryInput[] = [];
  const mimetype = model.format === "ods" ? parts.get("mimetype") : undefined;
  if (mimetype) {
    entries.push({ name: "mimetype", data: mimetype, store: true });
  }
  for (const [name, data] of parts) {
    if (mimetype && name === "mimetype") {
      continue;
    }
    entries.push({ name, data });
  }
  return {
    bytes: await zip(entries),
    fileName: spreadsheetFileName(model),
    mimeType: spreadsheetMimeType(model.format, model.packageVariant),
  };
}

/** Triggers a browser download of the workbook. */
export async function downloadSpreadsheet(model: SpreadsheetModel): Promise<void> {
  const { bytes, fileName, mimeType } = await spreadsheetDownload(model);
  triggerBlobDownload(bytes, fileName, mimeType);
}
