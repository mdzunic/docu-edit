// Download a workbook (SPEC-116, SPEC-121) — the spreadsheet twin of
// lib/downloadDocument.ts, presentation/download.ts and media/download.ts,
// kept framework-free. In its own format the workbook goes through its
// format's writer: an untouched workbook re-emits every source part
// byte-identical, in source order (the ZIP container itself is rebuilt —
// round trips compare parts); an edited one takes the substitution path; a
// workbook with no package (a new one) is built from scratch. An `.ods`
// keeps `mimetype` first and STOREd, as ODF requires.
//
// "Download a copy as…" the other format (SPEC-121 "Cross-format copy")
// writes `stripRetention({ ...model, format: target })` through the target's
// from-scratch writer — the open model is never mutated — and reports, as
// whole one-line `notices` the view shows as toasts, the structures the copy
// leaves out ("Not included in the .ods copy: charts, comments"). The file
// name follows the target format with no package variant.
import { triggerBlobDownload } from "../core/triggerDownload";
import { spreadsheetFileExtension, spreadsheetMimeType } from "./fileTypes";
import { DEFAULT_SPREADSHEET_TITLE, type SpreadsheetFormat, type SpreadsheetModel } from "./model";
import type { SpreadsheetLocale } from "./numberFormat/locale";
import { stripRetention, unconvertibleContent } from "./sheets";

/** The file name a workbook downloads under: its title plus the extension
 * of its format and variant (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.ods`,
 * `.ots`). */
export function spreadsheetFileName(model: SpreadsheetModel): string {
  const title = model.title.trim();
  return `${title.length > 0 ? title : DEFAULT_SPREADSHEET_TITLE}.${spreadsheetFileExtension(model.format, model.packageVariant)}`;
}

/** The toast a copy shows when it leaves something out. */
export function notIncludedNotice(target: SpreadsheetFormat, labels: readonly string[]): string {
  return `Not included in the .${target} copy: ${labels.join(", ")}`;
}

export interface SpreadsheetDownload {
  bytes: Uint8Array;
  fileName: string;
  mimeType: string;
  /** Whole one-line messages the view shows as toasts. */
  notices: string[];
}

async function write(model: SpreadsheetModel, locale: SpreadsheetLocale | undefined): Promise<Uint8Array> {
  if (model.format === "ods") {
    const { writeOds } = await import("./ods/writer");
    return writeOds(model, locale ? { locale } : {});
  }
  const { writeXlsx } = await import("./xlsx/writer");
  return writeXlsx(model);
}

/** The bytes, name and MIME type of a workbook's download — a copy in
 * `targetFormat` when that is not the workbook's own. */
export async function spreadsheetDownload(
  model: SpreadsheetModel,
  targetFormat: SpreadsheetFormat = model.format,
  locale?: SpreadsheetLocale,
): Promise<SpreadsheetDownload> {
  if (targetFormat === model.format) {
    return {
      bytes: await write(model, locale),
      fileName: spreadsheetFileName(model),
      mimeType: spreadsheetMimeType(model.format, model.packageVariant),
      notices: [],
    };
  }
  const labels = unconvertibleContent(model, targetFormat);
  const copy = stripRetention({ ...model, format: targetFormat });
  return {
    bytes: await write(copy, locale),
    fileName: spreadsheetFileName(copy),
    mimeType: spreadsheetMimeType(targetFormat),
    notices: labels.length > 0 ? [notIncludedNotice(targetFormat, labels)] : [],
  };
}

/** Triggers a browser download of the workbook (or its copy); resolves to
 * the notices to show. `locale` is the user's number locale (SPEC-122: an
 * .ods cell's informative text is written in it). */
export async function downloadSpreadsheet(model: SpreadsheetModel, targetFormat?: SpreadsheetFormat, locale?: SpreadsheetLocale): Promise<string[]> {
  const { bytes, fileName, mimeType, notices } = await spreadsheetDownload(model, targetFormat, locale);
  triggerBlobDownload(bytes, fileName, mimeType);
  return notices;
}
