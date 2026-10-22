// SPEC-126's CSV messages ("Messages"): the refusals the reader throws and the
// notices a download or a sheet command shows as toasts. Kept apart from the
// parser and the writer so the view and the download path can name them
// without loading either.

export const CSV_NOT_TEXT_MESSAGE = "This file doesn't contain text we can read as CSV";
export const CSV_TOO_MANY_ROWS_MESSAGE = "This file has more than 1,048,576 rows — too many to open here";
/** "Line N has more than 16,384 fields — too many to open here"; the
 * reader names N through csvTooManyColumnsMessage. */
export const CSV_TOO_MANY_COLUMNS_MESSAGE = "Line N has more than 16,384 fields — too many to open here";
export const CSV_LOSS_NOTICE =
  "Formatting, formulas and other sheets aren't saved in CSV — download a copy as .xlsx to keep them";
export const CSV_SECOND_SHEET_NOTICE = "CSV files keep only the first sheet — download a copy as .xlsx to keep the others";
export const CSV_ENCODING_FALLBACK_NOTICE = "Some characters can't be saved as Windows-1252, so this file was saved as UTF-8";

/** CSV_TOO_MANY_COLUMNS_MESSAGE naming the record's 1-based line. */
export function csvTooManyColumnsMessage(line: number): string {
  return CSV_TOO_MANY_COLUMNS_MESSAGE.replace("N", String(line));
}
