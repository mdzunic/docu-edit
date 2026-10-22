// The workbook refusals (SPEC-116): every workbook this product cannot read
// is refused with a message that names the way out — never misparsed. The
// open funnel (openDocument.ts) throws the first three before any reader
// runs; the readers throw the ceiling and their own parse errors. Kept apart
// from the readers so the funnel can name them without loading a codec.

/** A workbook the product recognised but will not (or cannot) open. */
export class SpreadsheetOpenError extends Error {}

export const XLSB_UNSUPPORTED_MESSAGE = "Excel binary workbooks (.xlsb) aren't supported — save as .xlsx first";
export const LEGACY_OR_ENCRYPTED_WORKBOOK_MESSAGE =
  "Old .xls files and password-protected workbooks aren't supported — save as an unprotected .xlsx first";
export const NUMBERS_UNSUPPORTED_MESSAGE =
  "Apple Numbers files can't be opened yet — in Numbers, choose File ▸ Export To ▸ Excel, then open the .xlsx";
export const WORKBOOK_TOO_LARGE_MESSAGE = "This workbook has more than 2,000,000 filled cells — too large to open here";
