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
import type { DocumentTheme, ModelColor, ModelFont, ThemeColorRef } from "../core/theme";
import type { CellRef, GridRange } from "./address";
import type { SheetSelection } from "./selection";

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

// --- cell styles (SPEC-117 "Model additions") --------------------------------
// SPEC-117 is the normative source for CellStyle and the style tables both
// readers fill it from (xlsx/styles.ts, ods/styles.ts). Every colour is a
// core ModelColor resolved through resolveColor(color, model.theme ??
// DEFAULT_THEME) — nothing else interprets one.

/** A theme reference read from SpreadsheetML; `xlsxTint` is the source `tint`
 *  attribute text, re-emitted verbatim by SPEC-122 when the colour is unchanged
 *  (resolveColor ignores it — the same move as core/themeDrawingMl.ts's `spelling`). */
export interface SheetColorRef extends ThemeColorRef {
  xlsxTint?: string;
}

export type SheetBorderStyle =
  | "thin"
  | "medium"
  | "thick"
  | "dashed"
  | "dotted"
  | "double"
  | "hair"
  | "mediumDashed"
  | "dashDot"
  | "mediumDashDot"
  | "dashDotDot"
  | "mediumDashDotDot"
  | "slantDashDot";

export interface CellBorderEdge {
  style: SheetBorderStyle;
  color?: ModelColor; // absent = automatic (black)
}

export interface CellFont {
  family?: ModelFont;
  sizePt?: number;
  bold?: true;
  italic?: true;
  underline?: "single" | "double" | "singleAccounting" | "doubleAccounting";
  strike?: true;
  color?: ModelColor; // absent = automatic (black)
  vertAlign?: "superscript" | "subscript";
}

export type CellFill =
  | { kind: "solid"; color: ModelColor }
  | { kind: "pattern"; pattern: string; fg?: ModelColor; bg?: ModelColor } // xlsx patternType other than none/solid
  | {
      kind: "gradient";
      type: "linear" | "path";
      degree?: number;
      left?: number; // path focus, 0..1
      right?: number;
      top?: number;
      bottom?: number;
      stops: { position: number; color: ModelColor }[];
    };

export interface CellAlignment {
  horizontal?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
  vertical?: "top" | "center" | "bottom" | "justify" | "distributed"; // absent = bottom
  wrap?: true;
  indent?: number;
  textRotation?: number; // 0..180, 255 = vertical stacked
  shrinkToFit?: true;
}

export interface CellBorders {
  top?: CellBorderEdge;
  right?: CellBorderEdge;
  bottom?: CellBorderEdge;
  left?: CellBorderEdge;
  diagonal?: CellBorderEdge & { up?: true; down?: true };
}

export interface CellStyle {
  font?: CellFont;
  fill?: CellFill;
  borders?: CellBorders;
  alignment?: CellAlignment;
  numberFormat?: string; // declared here, populated by SPEC-118; absent = "General"
  builtinFormatId?: number; // SPEC-118: a locale-dependent Excel built-in (14, 22, 5–8, 41–44)
  protection?: { locked?: false; hidden?: true }; // locked defaults to true
  quotePrefix?: true; // xlsx xf@quotePrefix: the value was entered as text with a leading '
  namedStyle?: string; // "Good", "Heading 1"… — declared here, read and written by SPEC-123
  sourceXf?: number; // xlsx: the cellXfs index this was read from
  sourceStyleName?: string; // ods: the automatic/common style name this was read from
}

/** A run of columns sharing one size, visibility and style (SPEC-117) —
 * stored only when it has a width other than the sheet default, `hidden` or
 * a `styleId`. */
export interface ColumnSpan {
  min: number; // 0-based inclusive; spans sorted, non-overlapping
  max: number;
  widthPx?: number; // px at 100% zoom; absent = the sheet default
  hidden?: true;
  styleId?: number;
  sourceWidthChars?: number; // xlsx <col width>, re-emitted verbatim while widthPx is unchanged
  sourceAttrs?: string; // xlsx: unmodelled <col> attributes, verbatim, in source order
  sourceStyleName?: string; // ods: the table-column style name
}

/** One run of a rich-text string. `font` (SPEC-117) is a full run font in
 * xlsx (`<rPr>`) and the span's own properties in ods (`text:span`). */
export interface TextRun {
  text: string;
  font?: CellFont;
}

export interface Cell {
  col: number; // 0-based
  value: CellValue; // the literal, or a formula cell's cached result; null = no value
  runs?: TextRun[]; // rich text; `value` is then the runs' concatenated text
  /** Index into `SpreadsheetModel.styles` (SPEC-117). A style-only cell —
   * `value: null` with a `styleId` — is stored. */
  styleId?: number;
  // --- SPEC-119: formulas (formula/*.ts is the normative source) ---
  /** Canonical formula text without the leading "=" (Excel's en-US syntax).
   *  When `formulaUnparsed`, the WIRE text instead (xlsx <f> text; ods the
   *  whole table:formula value, prefix included). `value` stays the
   *  producer's cached result. A formula cell is stored even with
   *  `value: null` (no cached value). */
  formula?: string;
  formulaUnparsed?: true;
  /** Legacy array (CSE) formula: the range it fills. On the anchor only —
   *  the top-left cell, which holds `formula`; member cells hold values. */
  arrayRange?: GridRange;
  /** xlsx: the anchor's array formula is a dynamic-array (spill) formula —
   *  `t="array"` plus a `cm` attribute. arrayRange is its last spill extent. */
  dynamicArray?: true;
  /** xlsx `c@cm` verbatim (index into xl/metadata.xml cellMetadata), so the
   *  sheet writer (SPEC-122) re-emits it. */
  cm?: number;
  /** xlsx What-If data table: the anchor's `<f t="dataTable" …/>` element,
   *  verbatim. No `formula`; members hold cached values. Never evaluated. */
  dataTableXml?: string;
  // Later specs add optional fields only.
}

export interface Row {
  index: number; // 0-based
  cells: Cell[]; // ascending by col, unique
  // SPEC-117 — a row with no cells is stored only if it has a height,
  // `hidden` or a `styleId`:
  heightPt?: number;
  customHeight?: true;
  hidden?: true;
  /** Present only when the source row had customFormat="1" (xlsx) or a
   * default cell style (ods). */
  styleId?: number;
  sourceAttrs?: string; // xlsx: unmodelled <row> attributes, verbatim, in source order
  sourceStyleName?: string; // ods: the row style name
}

export interface Sheet {
  id: string; // minted at read; persisted in the store; never written to a file
  name: string;
  visibility: "visible" | "hidden" | "veryHidden";
  kind: "grid" | "foreign"; // foreign: chartsheet/dialogsheet/macrosheet — placeholder + verbatim
  rows: Row[]; // sparse, ascending by index
  partName?: string; // xlsx: the sheet part ("xl/worksheets/sheet1.xml"); its bytes are in `opaque`
  sourceXml?: string; // ods: this sheet's <table:table> text, with the whitespace that preceded it
  // --- SPEC-117: layout, merges and view settings ---
  columns: ColumnSpan[]; // sparse, sorted, non-overlapping
  defaultColWidthPx: number; // px at 100% zoom
  defaultRowHeightPt: number;
  /** Ranges of two or more cells, each overlapping no earlier merge. */
  merges: GridRange[];
  frozen?: { rows: number; cols: number };
  showGridLines?: false;
  showZeros?: false;
  zoomPercent?: number; // 10..400, absent = 100
  tabColor?: ModelColor;
  selection?: SheetSelection; // view state (SPEC-117 "View state")
  topLeftCell?: CellRef;
  /** View ▸ Show formulas for this sheet (xlsx sheetView@showFormulas,
   * SPEC-119). View state: never marks the sheet dirty. */
  showFormulas?: true;
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
  // --- SPEC-117 ---
  /** Every cell style, index 0 the default ("Normal") style — never empty. */
  styles: CellStyle[];
  defaultStyleId: 0;
  theme?: DocumentTheme; // xlsx: the workbook's theme part; ods: none
  indexedColors?: string[]; // the file's 64-entry override palette, lowercase #rrggbb
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
