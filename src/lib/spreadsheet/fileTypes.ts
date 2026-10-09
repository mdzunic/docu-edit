// A workbook's file type by format + package variant (SPEC-116's "File
// names and labels" table): the extension a download is named with (and the
// library badge is spelled from) and the MIME type it is served under. A
// template opens as an ordinary workbook and, being byte-preserved,
// downloads as the template it is.
import type { PackageVariant, SpreadsheetFormat } from "./model";

interface FileType {
  extension: string;
  mimeType: string;
}

const PLAIN: Record<SpreadsheetFormat, FileType> = {
  xlsx: { extension: "xlsx", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  ods: { extension: "ods", mimeType: "application/vnd.oasis.opendocument.spreadsheet" },
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
};

function fileTypeOf(format: SpreadsheetFormat, variant: PackageVariant | undefined): FileType {
  return (variant !== undefined ? VARIANTS[format][variant] : undefined) ?? PLAIN[format];
}

/** "xlsx", "xlsm", "xltx", "xltm", "ods" or "ots". */
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
