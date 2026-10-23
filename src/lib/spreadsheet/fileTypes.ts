// A workbook's file type by format + package variant (SPEC-116's "File
// names and labels" table): the extension a download is named with (and the
// library badge is spelled from) and the MIME type it is served under. A
// template opens as an ordinary workbook and, being byte-preserved,
// downloads as the template it is. SPEC-126's delimited-text formats have no
// variants.
import type { PackageVariant, SpreadsheetFormat } from "./model";

interface FileType {
  extension: string;
  mimeType: string;
}

const PLAIN: Record<SpreadsheetFormat, FileType> = {
  xlsx: { extension: "xlsx", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  ods: { extension: "ods", mimeType: "application/vnd.oasis.opendocument.spreadsheet" },
  csv: { extension: "csv", mimeType: "text/csv" },
  tsv: { extension: "tsv", mimeType: "text/tab-separated-values" },
};

const VARIANTS: Record<SpreadsheetFormat, Partial<Record<PackageVariant, FileType>>> = {
  xlsx: {
    macroEnabled: { extension: "xlsm", mimeType: "application/vnd.ms-excel.sheet.macroEnabled.12" },
    template: { extension: "xltx", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.template" },
    macroEnabledTemplate: { extension: "xltm", mimeType: "application/vnd.ms-excel.template.macroEnabled.12" },
  },
  ods: {
    template: { extension: "ots", mimeType: "application/vnd.oasis.opendocument.spreadsheet-template" },
  },
  csv: {},
  tsv: {},
};

function fileTypeOf(format: SpreadsheetFormat, variant: PackageVariant | undefined): FileType {
  return (variant !== undefined ? VARIANTS[format][variant] : undefined) ?? PLAIN[format];
}

/** "xlsx", "xlsm", "xltx", "xltm", "ods", "ots", "csv" or "tsv". */
export function spreadsheetFileExtension(format: SpreadsheetFormat, variant?: PackageVariant): string {
  return fileTypeOf(format, variant).extension;
}

/** The MIME type a workbook of this format + variant is served under. */
export function spreadsheetMimeType(format: SpreadsheetFormat, variant?: PackageVariant): string {
  return fileTypeOf(format, variant).mimeType;
}

/** Every content type a workbook can be served under, so the open-from-URL
 * check admits the spreadsheet kind without knowing any of them itself. */
export const SPREADSHEET_CONTENT_TYPES: readonly string[] = (["xlsx", "ods"] as const).flatMap((format) => [
  PLAIN[format].mimeType,
  ...Object.values(VARIANTS[format]).map((type) => type.mimeType),
]);

/** SPEC-137: the extensions a workbook PACKAGE is named with (the .xlsx
 * family and .ods/.ots) — what the library's open progress shows from the
 * moment such a file is chosen, before the open funnel (which routes by the
 * package's content, never by its name) has inflated it. */
export const SPREADSHEET_PACKAGE_EXTENSIONS: readonly string[] = (["xlsx", "ods"] as const).flatMap((format) => [
  PLAIN[format].extension,
  ...Object.values(VARIANTS[format]).map((type) => type.extension),
]);
