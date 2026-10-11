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
//
// SPEC-126: a csv/tsv workbook downloads as the exact source bytes
// (CSV_SOURCE_PART) while `sheets[0]` is the sheet read from the file
// (`csv.sheetId`) and is not dirty — whatever else the workbook gained —
// and otherwise as writeCsv's file in the source's dialect. Either way the
// download happens; CSV_LOSS_NOTICE is added when csvLosses(model, 0) finds
// something CSV cannot carry, and writeCsv's encoding fallback notice when
// Windows-1252 could not hold the text. "Download a copy as CSV/TSV (current
// sheet)" from any workbook is spreadsheetSheetCopy: the active sheet's
// displayed values in Excel's "CSV UTF-8" shape.
import { triggerBlobDownload } from "../core/triggerDownload";
import { csvSourceBytes } from "./csv/parse";
import { CSV_LOSS_NOTICE } from "./csv/messages";
import { csvLosses, sheetToCsv, writeCsv } from "./csv/write";
import { spreadsheetFileExtension, spreadsheetMimeType } from "./fileTypes";
import { DEFAULT_SPREADSHEET_TITLE, isCsvFormat, type CsvDialect, type SpreadsheetFormat, type SpreadsheetModel } from "./model";
import { SPREADSHEET_LOCALES, DEFAULT_SPREADSHEET_LOCALE, type SpreadsheetLocale } from "./numberFormat/locale";
import { stripRetention, unconvertibleContent } from "./sheets";

/** The title a download is named after: the workbook's, or the default
 * title for an untitled one. */
function downloadTitle(model: Pick<SpreadsheetModel, "title">): string {
  const title = model.title.trim();
  return title.length > 0 ? title : DEFAULT_SPREADSHEET_TITLE;
}

/** The file name a workbook downloads under: its title plus the extension
 * of its format and variant (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.ods`,
 * `.ots`, `.csv`, `.tsv`). */
export function spreadsheetFileName(model: SpreadsheetModel): string {
  return `${downloadTitle(model)}.${spreadsheetFileExtension(model.format, model.packageVariant)}`;
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

/** A workbook in its own format: the bytes and the notices the write
 * raised. */
async function write(model: SpreadsheetModel, locale: SpreadsheetLocale | undefined): Promise<{ bytes: Uint8Array; notices: string[] }> {
  switch (model.format) {
    case "xlsx": {
      const { writeXlsx } = await import("./xlsx/writer");
      return { bytes: await writeXlsx(model), notices: [] };
    }
    case "ods": {
      const { writeOds } = await import("./ods/writer");
      return { bytes: await writeOds(model, locale ? { locale } : {}), notices: [] };
    }
    case "csv":
    case "tsv":
      return writeCsvWorkbook(model);
  }
}

/** Whether a csv/tsv workbook still downloads as its source bytes: its first
 * sheet is the one read from the file, unedited. */
export function csvSourceIntact(model: SpreadsheetModel): boolean {
  const first = model.sheets[0];
  return model.csv !== undefined && first !== undefined && first.id === model.csv.sheetId && !first.dirty;
}

/** A csv/tsv workbook's own file (see the module header). */
function writeCsvWorkbook(model: SpreadsheetModel): { bytes: Uint8Array; notices: string[] } {
  const losses = csvLosses(model, 0).length > 0 ? [CSV_LOSS_NOTICE] : [];
  if (csvSourceIntact(model)) {
    return { bytes: csvSourceBytes(model), notices: losses };
  }
  const written = writeCsv(model);
  return { bytes: written.bytes, notices: [...losses, ...written.notices] };
}

/** "Download a copy as CSV/TSV (current sheet)" (SPEC-126 "Copies"): the
 * active sheet's displayed values as UTF-8 with a BOM, CRLF after every
 * record, the user's locale's list separator (or a tab) and that locale's
 * numbers and dates — Excel's "CSV UTF-8" shape, since a BOM-less UTF-8
 * file is the one CSV Excel misreads. Named `<title>.csv` for a one-sheet
 * workbook, else `<title> - <sheet name>.csv`; CSV_LOSS_NOTICE when the
 * sheet holds what CSV cannot carry (other sheets aside). The model is never
 * mutated. */
export function spreadsheetSheetCopy(
  model: SpreadsheetModel,
  target: "csv" | "tsv",
  locale: SpreadsheetLocale = SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE],
): SpreadsheetDownload {
  const index = model.sheets[model.activeSheetIndex] ? model.activeSheetIndex : 0;
  const sheet = model.sheets[index];
  const dialect: Omit<CsvDialect, "sheetId"> = {
    delimiter: target === "tsv" ? "\t" : locale.listSeparator === ";" ? ";" : ",",
    encoding: "utf-8",
    bom: true,
    lineEnding: "\r\n",
    finalLineEnding: true,
    quoteAll: false,
    localeTag: locale.tag,
  };
  const { bytes } = sheetToCsv(model, index, dialect);
  const name = model.sheets.length > 1 && sheet ? `${downloadTitle(model)} - ${sheet.name}` : downloadTitle(model);
  return {
    bytes,
    fileName: `${name}.${spreadsheetFileExtension(target)}`,
    mimeType: spreadsheetMimeType(target),
    notices: csvLosses(model, index).some((loss) => loss !== "sheets") ? [CSV_LOSS_NOTICE] : [],
  };
}

/** The bytes, name and MIME type of a workbook's download — a copy in
 * `targetFormat` when that is not the workbook's own (to csv/tsv, the
 * current sheet's copy). */
export async function spreadsheetDownload(
  model: SpreadsheetModel,
  targetFormat: SpreadsheetFormat = model.format,
  locale?: SpreadsheetLocale,
): Promise<SpreadsheetDownload> {
  if (targetFormat === model.format) {
    const written = await write(model, locale);
    return {
      bytes: written.bytes,
      fileName: spreadsheetFileName(model),
      mimeType: spreadsheetMimeType(model.format, model.packageVariant),
      notices: written.notices,
    };
  }
  if (isCsvFormat(targetFormat)) {
    return spreadsheetSheetCopy(model, targetFormat, locale);
  }
  const labels = unconvertibleContent(model, targetFormat);
  const copy = stripRetention({ ...model, format: targetFormat });
  return {
    bytes: (await write(copy, locale)).bytes,
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

/** Triggers a browser download of the active sheet's CSV/TSV copy;
 * resolves to the notices to show. */
export async function downloadSpreadsheetSheetCopy(model: SpreadsheetModel, target: "csv" | "tsv", locale?: SpreadsheetLocale): Promise<string[]> {
  const { bytes, fileName, mimeType, notices } = spreadsheetSheetCopy(model, target, locale);
  triggerBlobDownload(bytes, fileName, mimeType);
  return notices;
}
