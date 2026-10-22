// The workbook readers' version (SPEC-116 "Reader versions"; the history of
// what each version reads is open.ts's header). Kept in a module of its own
// so the store, the workbook view and the sheet commands can compare a
// record's `readerVersion` without importing open.ts — which carries both
// readers' read phases (SPEC-137) and so loads on demand, off the library
// page's module graph.
export const SPREADSHEET_READER_VERSION = 15;
