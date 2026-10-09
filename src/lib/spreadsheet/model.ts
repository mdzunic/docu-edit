// The in-memory workbook model — the fourth product's counterpart of
// documentModel.ts, presentation/model.ts and media/model.ts (SPEC-116).
// THIS FILE is the normative source for `SpreadsheetModel`, `Sheet`, `Row`,
// `Cell` and `CellValue`; every later spreadsheet spec (SPEC-117–140) edits
// this model and keeps the invariants below.
//
// Invariants — normative for every later spec:
// 1. Copy-on-write rows. No command mutates a `Row` or `Cell` in place, nor a
//    sheet's `rows` array: an edit builds new objects for what it touches and
//    a new (shallow-copied) `rows` array, and every untouched row keeps its
//    object identity. The undo stack (SPEC-121), recalculation's change
//    detection (SPEC-120) and the store's per-sheet persistence split
//    (documentStore.ts) all find what changed by reference identity.
// 2. Sparse. A cell with `value: null` and nothing else is not stored; a row
//    with no cells and no properties is not stored. Nothing is ever allocated
//    per column or per row of the 1,048,576 × 16,384 grid.
// 3. The source is the splice base. A sheet's source — the xlsx part bytes
//    in `opaque`, the ods `sourceXml` — is kept for the life of the sheet,
//    even after it is edited. Writers (SPEC-121/122) re-serialise only what
//    the model models and splice it into the source; every unmodelled element
//    is re-emitted verbatim. This is what lets a later spec model a new
//    element and upgrade already-stored workbooks by reading it from the
//    retained source (see open.ts, "Reader versions").
// 4. Dates are serials. There is no date type: a date is a number whose
//    meaning comes from its number format (SPEC-118). Times are day
//    fractions (serial.ts converts).
//
// The model never holds a `Map`, `Set` or class instance, and never reaches
// the store as a Vue reactive proxy (a proxy fails IndexedDB's structured
// clone): `JSON.parse(JSON.stringify(model))` `toEqual`s it.
//
// Layering rule (enforced by boundary.spec.ts): this directory imports only
// its own siblings and src/lib/core — never documentModel.ts, the text
// commands, docx/, odf/, md/, presentation/, media/, the store or Vue.
import type { DocumentProperties } from "../core/documentProperties";
import type { MediaPart } from "../core/media";
import type { OpaquePart } from "../core/packageParts";

export type SpreadsheetFormat = "xlsx" | "ods"; // SPEC-126 widens with "csv" | "tsv"
export type PackageVariant = "macroEnabled" | "template" | "macroEnabledTemplate";

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384; // A..XFD
export const MAX_WORKBOOK_CELLS = 2_000_000; // stored cells, all sheets (SPEC-137 revisits)

export type ErrorCode =
  | "#NULL!"
  | "#DIV/0!"
  | "#VALUE!"
  | "#REF!"
  | "#NAME?"
  | "#NUM!"
  | "#N/A"
  | "#GETTING_DATA"
  | "#SPILL!"
  | "#CALC!"
  | "#FIELD!"
  | "#BLOCKED!"
  | "#CONNECT!"
  | "#BUSY!"
  | "#UNKNOWN!";
export interface CellError {
  error: ErrorCode;
}

/** Dates and times are numbers — serials in the workbook's date system. */
export type CellValue = number | string | boolean | CellError | null;

/** One run of a rich-text string. SPEC-117 adds `font?: CellFont`; until
 *  then a run carries its text only (the formatting stays in the source). */
export interface TextRun {
  text: string;
}

export interface Cell {
  col: number; // 0-based
  value: CellValue; // the literal, or a formula cell's cached result; null = no value
  runs?: TextRun[]; // rich text; `value` is then the runs' concatenated text
  // Later specs add optional fields only: formula (119), styleId (117), …
}

export interface Row {
  index: number; // 0-based
  cells: Cell[]; // ascending by col, unique
  // SPEC-117 adds heightPt?, customHeight?, hidden?, styleId?
}

export interface Sheet {
  id: string; // minted at read; persisted in the store; never written to a file
  name: string;
  visibility: "visible" | "hidden" | "veryHidden";
  kind: "grid" | "foreign"; // foreign: chartsheet/dialogsheet/macrosheet — placeholder + verbatim
  rows: Row[]; // sparse, ascending by index
  partName?: string; // xlsx: the sheet part ("xl/worksheets/sheet1.xml"); its bytes are in `opaque`
  sourceXml?: string; // ods: this sheet's <table:table> text, with the whitespace that preceded it
}

export interface SheetPreview {
  sheetName: string;
  /** Up to PREVIEW_ROWS (12) × PREVIEW_COLS (6) display strings from the
   *  active sheet's top-left corner; "" for empty cells. */
  cells: string[][];
}

export interface SpreadsheetModel {
  kind: "spreadsheet";
  id: string;
  title: string;
  format: SpreadsheetFormat;
  packageVariant?: PackageVariant;
  readerVersion: number;
  sheets: Sheet[];
  activeSheetIndex: number;
  dateSystem: "1900" | "1904";
  odsNullDate?: string; // ISO date, only when neither 1899-12-30 nor 1904-01-01
  odsContentEnvelope?: { prefix: string; suffix: string };
  opaque: OpaquePart[]; // the whole source package, verbatim, in order
  preview?: SheetPreview;
  properties?: DocumentProperties;
  media?: MediaPart[]; // unused until SPEC-131; present so the store's media split applies unchanged
}

/** The error codes a cell can hold, for readers validating what a file
 * declares (an unknown code reads as `#VALUE!`). */
export const ERROR_CODES: readonly ErrorCode[] = [
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#GETTING_DATA",
  "#SPILL!",
  "#CALC!",
  "#FIELD!",
  "#BLOCKED!",
  "#CONNECT!",
  "#BUSY!",
  "#UNKNOWN!",
];

const ERROR_CODE_SET: ReadonlySet<string> = new Set(ERROR_CODES);

/** The error a declared code stands for: the code itself when it is one of
 * ErrorCode, else `#VALUE!` (both readers' rule for unknown codes). */
export function toCellError(code: string): CellError {
  const trimmed = code.trim();
  return { error: ERROR_CODE_SET.has(trimmed) ? (trimmed as ErrorCode) : "#VALUE!" };
}

/** Whether a cell value is an error (the one object-shaped CellValue). */
export function isCellError(value: CellValue): value is CellError {
  return typeof value === "object" && value !== null;
}
