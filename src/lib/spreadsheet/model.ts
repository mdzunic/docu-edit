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
import type { ChartVisual } from "../core/charts";
import type { DocumentProperties } from "../core/documentProperties";
import type { MediaPart } from "../core/media";
import type { OpaquePart } from "../core/packageParts";
import type { ShapeVisual } from "../core/shapes";
import type { DocumentTheme, ModelColor, ModelFont, ThemeColorRef } from "../core/theme";
import type { CellRef, GridRange } from "./address";
import type { ConditionalFormat, Dxf } from "./conditionalFormats";
import type { AutoFilter } from "./filter";
import type { SpreadsheetLocaleTag } from "./numberFormat/locale";
import type { SheetSelection } from "./selection";
import type { SortState } from "./sort";
import type { NamedCellStyle } from "./styles";
import type { DataValidation } from "./validation";

export type SpreadsheetFormat = "xlsx" | "ods" | "csv" | "tsv";
export type PackageVariant = "macroEnabled" | "template" | "macroEnabledTemplate";

// --- SPEC-126: the CSV dialect --------------------------------------------------
// SPEC-126 is the normative source for CsvDialect: how a `.csv`/`.tsv` file
// was written, recorded at read so a save writes it back the same way (and
// csv/write.ts never guesses). Present exactly when `format` is "csv" or
// "tsv"; the source bytes themselves are the single `opaque` part
// csv/parse.ts's CSV_SOURCE_PART.
export type CsvDelimiter = "," | ";" | "\t" | "|";
export type CsvEncoding = "utf-8" | "utf-16le" | "utf-16be" | "windows-1252";
export interface CsvDialect {
  delimiter: CsvDelimiter;
  encoding: CsvEncoding;
  bom: boolean; // the source began with its encoding's BOM
  lineEnding: "\r\n" | "\n" | "\r"; // the first record terminator outside quotes ("\r\n" if none)
  finalLineEnding: boolean; // the source's last record was terminated
  quoteAll: boolean; // every non-empty field in the sniff sample was quoted
  width?: number; // fields per record, when every non-blank record agreed
  sepLine?: true; // the source began with Excel's "sep=X" line
  localeTag: SpreadsheetLocaleTag; // the locale the fields were read in (and are written in)
  sheetId: string; // the sheet read from the file
}

/** Whether a workbook format is one of SPEC-126's delimited-text formats. */
export function isCsvFormat(format: SpreadsheetFormat): format is "csv" | "tsv" {
  return format === "csv" || format === "tsv";
}

/** The title a workbook gets when nothing names it — a new workbook's, and
 * the download name of an untitled one (SPEC-121). */
export const DEFAULT_SPREADSHEET_TITLE = "Untitled spreadsheet";

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
 * a `styleId`, or (SPEC-122) an unmodelled xlsx attribute other than
 * `bestFit` (an `outlineLevel`, `collapsed` … the writer re-emits). */
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
  // --- SPEC-122: what the writers re-emit for content the model does not
  // model. Drop rules, applied by every command: setting a cell's value or
  // runs removes `sstIndex`, `preservedText` and ` vm="…"`; setting or
  // removing its formula removes `cm`, `dynamicArray` and `dataTableXml`. A
  // cell holding only preserved content is stored (sparseness counts it as
  // content). A later spec that models an attribute or child stops capturing
  // it here, and its reader step moves it into the model. ---
  /** xlsx: the shared-string index the value was read from; reused while
   *  value and runs are unchanged. */
  sstIndex?: number;
  /** Unmodelled start-tag attributes, verbatim with leading spaces: xlsx
   *  ` vm="2"`, ` ph="1"` (`cm` is SPEC-119's field); ods
   *  ` table:content-validation-name="v1"`, … */
  preservedAttrs?: string;
  /** Unmodelled children, verbatim: xlsx `<extLst>` (a data table's `<f>` is
   *  SPEC-119's `dataTableXml`); ods everything except the text paragraphs
   *  (annotations, frames, unknown elements), in source order. */
  preservedXml?: string;
  /** ods: the cell's `text:p` sequence verbatim when it holds inline content
   *  the model does not (links, fields, notes). */
  preservedText?: string;
  // Later specs add optional fields only.
}

export interface Row {
  index: number; // 0-based
  cells: Cell[]; // ascending by col, unique
  // SPEC-117 — a row with no cells is stored only if it has a height,
  // `hidden` or a `styleId`, or (SPEC-122) an unmodelled xlsx attribute the
  // writer re-emits (rowSourceAttrsHoldProperties):
  heightPt?: number;
  customHeight?: true;
  hidden?: true;
  /** SPEC-128: hidden by the sheet's autofilter — distinct from the user's
   *  `hidden` (Unhide clears only `hidden`; turning the filter off or
   *  clearing its criteria clears only this). The row is displayed hidden
   *  when either is set (rowIsHidden). A filter-hidden blank row is stored
   *  as a Row with no cells. */
  filterHidden?: true;
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
  /** grid: a worksheet; chart (SPEC-132): an xlsx chartsheet whose drawing
   *  resolves to one chart — that drawing is the sheet's one `drawings`
   *  entry; foreign: dialogsheet/macrosheet or an unresolvable chartsheet —
   *  placeholder + verbatim. */
  kind: "grid" | "chart" | "foreign";
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
  // --- SPEC-121 ---
  /** The sheet's own content changed since open: SPEC-122's writer rebuilds
   *  it. Born true on a sheet with no source (new). Never set by a view
   *  change (selection, scroll, tab switch). */
  dirty?: true;
  /** Set by "Duplicate sheet": the origin's name in the SOURCE file (its
   *  `copiedFrom` when the origin is itself a copy). The copy's own retained
   *  text (its `partName` part in `opaque`, or its `sourceXml`) is a copy of
   *  the origin's source-time text, so the writer maps that name to the
   *  copy's when remapping it. A copy is never an "original". */
  copiedFrom?: string;
  // --- SPEC-122 ---
  /** The sheet is protected (xlsx `sheetProtection@sheet="1"`, ods
   *  `table:protected="true"`): a locked cell refuses entry. Read only — the
   *  protection XML stays in the retained source, never changed. SPEC-123
   *  adds what the protection still permits: xlsx `formatCells="0"`,
   *  `formatColumns="0"` and `formatRows="0"` (ods has no equivalent), which
   *  let formatting reach locked cells of the matching range class; SPEC-124
   *  the structural edits it permits: xlsx `insertRows="0"`,
   *  `insertColumns="0"`, `deleteRows="0"` and `deleteColumns="0"`. */
  protection?: {
    sheet: true;
    allowFormatCells?: true;
    allowFormatColumns?: true;
    allowFormatRows?: true;
    allowInsertRows?: true;
    allowInsertColumns?: true;
    allowDeleteRows?: true;
    allowDeleteColumns?: true;
    /** SPEC-128: xlsx `sort="0"` / `autoFilter="0"` — sorting and using
     *  the autofilter are permitted on the protected sheet. */
    allowSort?: true;
    allowAutoFilter?: true;
  };
  // --- SPEC-128 (filter.ts and sort.ts are the normative sources) ---
  /** The sheet's autofilter: its range (header row included), per-column
   *  criteria and the record of its last sort. Replaced, never mutated. */
  autoFilter?: AutoFilter;
  /** The record of the last sort outside an autofilter (xlsx worksheet
   *  `<sortState>`): never re-applied. Replaced, never mutated. */
  sortState?: SortState;
  // --- SPEC-129 (conditionalFormats.ts is the normative source) ---
  /** The sheet's conditional formats, in source order (new ones appended);
   *  precedence is each rule's sheet-global `priority`. Copy-on-write: every
   *  command replaces the array and the formats it touches. */
  conditionalFormats: ConditionalFormat[];
  // --- SPEC-130 (validation.ts is the normative source) ---
  /** The sheet's data-validation rules, in source order (new ones
   *  appended); the first rule covering a cell is its rule. Copy-on-write:
   *  every command replaces the array and the rules it touches. */
  validations: DataValidation[];
  /** xlsx `dataValidations@disablePrompts="1"`: no input message shows. */
  validationPromptsDisabled?: true;
  // --- SPEC-131 (drawings.ts is the normative source for the commands) ---
  /** The sheet's drawing layer in z-order, index 0 at the back. Absent or
   *  empty: none. Copy-on-write like `rows`: every command replaces the
   *  array and the drawings it touches; an untouched drawing keeps its
   *  identity. */
  drawings?: SheetDrawing[];
  /** xlsx: the sheet's drawing part ("xl/drawings/drawing1.xml"), whose
   *  bytes are in `opaque` — the splice base of a rebuilt layer. */
  drawingPart?: string;
  /** xlsx: the drawing layer changed since open — the writer rebuilds the
   *  drawing part (and its `.rels`) only; the worksheet part stays as its
   *  own `dirty` says. ods drawings live in the table XML, so an ods
   *  drawing command sets `dirty` instead. */
  drawingsDirty?: true;
  // --- SPEC-133 (tables.ts is the normative source for the commands) ---
  /** The sheet's Excel tables, in source order (new ones appended). Absent
   *  or empty: none. Copy-on-write: every command replaces the array and the
   *  tables it touches. A table's change marks the sheet `dirty` (its
   *  `<tableParts>` follows `tables`). */
  tables?: SheetTable[];
}

// --- SPEC-133: defined names and tables -----------------------------------------------
// SPEC-133 is the normative source for defined names (names.ts: validity,
// resolution, the reserved `_xlnm.*` names) and Excel tables (tables.ts:
// structured references, the commands; tableStyles.ts: how a style renders).

/** A defined name. Relative parts of `formula` are relative to A1. */
export interface DefinedName {
  name: string; // as written; compared case-insensitively
  scopeSheetId?: string; // absent = workbook scope
  formula: string; // canonical; relative parts are relative to A1
  hidden?: true;
  comment?: string;
  formulaUnparsed?: true; // kept in its source dialect, never evaluated
  sourceAttrs?: string; // xlsx: unmodelled attributes verbatim; ods: table:range-usable-as
}

export type TotalsFunction = "sum" | "min" | "max" | "average" | "count" | "countNums" | "stdDev" | "var" | "custom";

export interface SheetTableColumn {
  name: string; // always equals the header cell's string when there is a header row
  xlsxId?: number; // tableColumn@id (new columns: 1 + max)
  totalsFunction?: TotalsFunction;
  totalsFormula?: string; // canonical; "custom" only
  totalsLabel?: string;
  calculatedFormula?: string; // canonical
}

export interface SheetTableStyle {
  name?: string; // absent = no style ("None")
  showRowStripes: boolean;
  showColumnStripes: boolean;
  showFirstColumn: boolean;
  showLastColumn: boolean;
}

export interface SheetTable {
  id: string; // minted at read, never written
  name: string; // displayName — what formulas use
  ref: GridRange; // header + data + totals
  headerRow: boolean;
  totalsRow: boolean;
  columns: SheetTableColumn[];
  style: SheetTableStyle;
  autoFilter?: { sourceXml?: string }; // present = filter buttons; criteria are SPEC-128's
  locked?: true; // tableType queryTable/xml: style options only
  partName?: string; // "xl/tables/table1.xml" — bytes stay in opaque
  xlsxId?: number; // table@id, unique in the workbook
  dirty?: true; // regenerate the part
}

// --- SPEC-131: the drawing layer --------------------------------------------------
// SPEC-131 is the normative source for the cell-anchored drawing model. A
// drawing is pinned to cells: its anchor says where, and how it follows its
// cells under row and column edits (drawings.ts anchorBehaviour). Every
// length is EMU (914,400 per inch; drawings.ts EMU_PER_PX = 9525 at 96 dpi).

/** A cell corner plus an offset into the cell, EMU. Offsets are kept as
 *  read, even when larger than their cell, and clamped only to render. */
export interface AnchorPoint {
  col: number;
  colOffEmu: number;
  row: number;
  rowOffEmu: number;
}

export interface DrawingAnchor {
  kind: "twoCell" | "oneCell" | "absolute";
  /** twoCell only — Excel's editAs. Absent = "twoCell" (the schema default). */
  editAs?: "twoCell" | "oneCell" | "absolute";
  from?: AnchorPoint; // twoCell, oneCell
  to?: AnchorPoint; // twoCell, as read or last written
  posEmu?: { x: number; y: number }; // absolute; twoCell+editAs "absolute" (derived at read)
  extEmu?: { cx: number; cy: number }; // oneCell, absolute; twoCell+editAs "oneCell"|"absolute" (derived at read)
}

/** The render-only subset of a shape's text (xdr:txBody / text:p). */
export interface DrawingText {
  paragraphs: { text: string; align?: "left" | "center" | "right" }[];
  anchor: "top" | "middle" | "bottom";
  /** From the first run. */
  sizePt?: number;
  color?: ModelColor;
  bold?: true;
}

/** What a drawing paints. `crop` and a group member's box are fractions 0..1
 *  (a member's of the group's box). */
export type SheetDrawingContent =
  | { kind: "image"; mediaId: string; crop?: { l: number; t: number; r: number; b: number } }
  | { kind: "shape"; visual: ShapeVisual; text?: DrawingText }
  | {
      kind: "group";
      members: { x: number; y: number; w: number; h: number; content: Exclude<SheetDrawingContent, { kind: "group" | "chart" }> }[];
    }
  | SheetChartContent
  | { kind: "foreign"; label: string; face?: Exclude<SheetDrawingContent, { kind: "group" | "chart" | "foreign" }> };

// --- SPEC-132: the sheet chart ---------------------------------------------------------
// SPEC-132 is the normative source for the chart content of a SheetDrawing:
// a core ChartVisual plus its live source references (charts.ts is the
// normative source for resolving, refreshing and the commands). A sheet
// chart's data IS the cells; the visual is what they said at the last
// refresh (or the file's caches, which open trusts).

/** One series' source references — canonical range formulas (SPEC-119
 *  text, sheet-qualified as read). */
export interface SheetChartSeriesRefs {
  name?: string; // canonical single-cell reference; absent = the literal visual name
  categories?: string; // canonical range (c:cat / c:xVal; ODS chart:categories or chart:domain)
  values?: string; // canonical range (c:val / c:yVal; ODS chart:values-cell-range-address)
}

export interface SheetChartRefs {
  series: SheetChartSeriesRefs[]; // index-aligned with the part's series (and visual.series)
  plotArea?: string; // ODS chart:plot-area@table:cell-range-address
  plotVisibleOnly: boolean; // xlsx c:plotVisOnly (absent = true); ODS !chart:include-hidden-cells
}

export interface SheetChartContent {
  kind: "chart";
  visual: ChartVisual;
  refs?: SheetChartRefs; // absent: a literal chart (data sheet editing, SPEC-052 style)
  sourcePart?: string; // "xl/charts/chart1.xml" | "Object 1/content.xml" — a NAME into opaque
  edited?: true; // data sheet / options / colour / source change of shape → regenerate
  cachesDirty?: true; // visual refreshed from cells → patch caches
  refsDirty?: true; // references changed → patch c:f / addresses
  readOnlyReason?: string; // rendered from caches; never refreshed or edited
}

export interface SheetDrawing {
  /** Minted at read ("dr1"…), unique per sheet, never written. */
  id: string;
  anchor: DrawingAnchor;
  /** xdr:cNvPr@name / draw:name. */
  name: string;
  /** xdr:cNvPr@descr / svg:desc. */
  altText?: string;
  /** xdr:cNvPr@hidden="1" — kept, never painted or selectable. */
  hidden?: true;
  /** a:xfrm@rot / draw:transform rotate, whole degrees clockwise — rendered only. */
  rotationDeg?: number;
  /** The verbatim source element: xlsx the wsDr child (the anchor, or the
   *  mc:AlternateContent around it, with the whitespace before it); ods the
   *  draw:* element. Absent for a drawing created here. Never dropped by a
   *  move, resize or z-order change. */
  sourceXml?: string;
  content: SheetDrawingContent;
}

/** A workbook's calculation settings (SPEC-120 "Calculation settings"), read
 * from the file and never written back by this product: the file's own
 * setting is preserved verbatim in the retained source. Each field is
 * present only when the file sets it. */
export interface CalcSettings {
  fullCalcOnLoad?: true; // xlsx calcPr@fullCalcOnLoad="1"
  manual?: true; // xlsx calcPr@calcMode="manual"; ods settings.xml AutoCalculate=false
  iterate?: true; // xlsx calcPr@iterate="1"; ods table:iteration@table:status="enable"
  caseSensitive?: true; // ods table:case-sensitive (ODF default "true"); never for xlsx
  /** ods: regex when use-regular-expressions is in effect, plain when
   * neither it nor use-wildcards is. */
  criteria?: "regex" | "plain";
  partialCellMatch?: true; // ods table:search-criteria-must-apply-to-whole-cell="false"
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
  /** SPEC-131: the images sheet drawings show, keyed by package path
   *  (`xl/media/image1.png`, `Pictures/…`) — one part per image, shared by
   *  every drawing that embeds it. The writers emit a part from here, ahead
   *  of an `opaque` twin; the store's media split applies unchanged. */
  media?: MediaPart[];
  // --- SPEC-117 ---
  /** Every cell style, index 0 the default ("Normal") style — never empty. */
  styles: CellStyle[];
  defaultStyleId: 0;
  theme?: DocumentTheme; // xlsx: the workbook's theme part; ods: none
  indexedColors?: string[]; // the file's 64-entry override palette, lowercase #rrggbb
  // --- SPEC-120 ---
  calc?: CalcSettings;
  /** Set by a recalculation pass that left a stale (uncomputable) cell, and
   * never cleared: SPEC-122's xlsx writer emits calcPr@fullCalcOnLoad="1". */
  forceFullCalcOnLoad?: true;
  // --- SPEC-121 ---
  /** The sheet list or a sheet's workbook-level properties (name,
   *  visibility, tab colour, order) changed since open. Absent on a freshly
   *  opened workbook; a from-scratch workbook is always fully built. */
  structureDirty?: true;
  // --- SPEC-122 ---
  /** The workbook's structure is protected (xlsx
   *  `workbookProtection@lockStructure="1"`, ods
   *  `office:spreadsheet@table:structure-protected="true"`): the sheet
   *  commands are disabled. Read only, never written. */
  structureProtected?: true;
  // --- SPEC-123 ---
  /** The workbook's named cell styles, in file order (styles.ts
   *  "Named cell styles"); a built-in applied for the first time is appended
   *  and stays (undo never removes one — an unused definition is harmless). */
  namedStyles?: NamedCellStyle[];
  // --- SPEC-126 ---
  /** How a `.csv`/`.tsv` source was written — present exactly when `format`
   *  is "csv" or "tsv". */
  csv?: CsvDialect;
  // --- SPEC-129 ---
  /** The differential-style table conditional formats reference (index =
   *  the xlsx `dxfId`): append-only — styles.ts's internDxf replaces the
   *  array when it appends, never removes or reorders an entry, so every
   *  dxfId stays valid. */
  dxfs: Dxf[];
  // --- SPEC-130 ---
  /** ods: the `table:content-validation` elements no cell references,
   *  verbatim (with their preceding whitespace), re-emitted by the writer
   *  whenever it regenerates `table:content-validations`. */
  odsOrphanValidations?: string[];
  // --- SPEC-133 (names.ts is the normative source) ---
  /** Every defined name, in source order (new ones appended) — workbook and
   *  sheet scope, reserved `_xlnm.*` names included. Copy-on-write: every
   *  command replaces the array and the names it touches. */
  names: DefinedName[];
  /** A name changed since open: the writers rebuild xlsx `<definedNames>`
   *  and the ODS named expressions from `names`. */
  namesDirty?: true;
}

/** Whether a sheet "has a source" (SPEC-121): an xlsx part in `opaque` or
 * an ods `<table:table>` text. A sheet with no source is always `dirty`. */
export function hasSource(sheet: Pick<Sheet, "partName" | "sourceXml">): boolean {
  return sheet.partName !== undefined || sheet.sourceXml !== undefined;
}

/** An original (SPEC-121): a sheet with a source and no `copiedFrom`. */
export function isOriginal(sheet: Pick<Sheet, "partName" | "sourceXml" | "copiedFrom">): boolean {
  return hasSource(sheet) && sheet.copiedFrom === undefined;
}

const VM_ATTRIBUTE = /\s+vm\s*=\s*(?:"[^"]*"|'[^']*')/g;

/** Removes, from a cell copy about to take a new value or runs, what the
 * value invalidates (the drop rules, SPEC-122): `runs`, `sstIndex`,
 * `preservedText` and a ` vm="…"` in `preservedAttrs` — whoever sets the
 * value, a command or a recalculation. */
export function dropValueBound(cell: Cell): void {
  delete cell.runs;
  delete cell.sstIndex;
  delete cell.preservedText;
  if (cell.preservedAttrs !== undefined) {
    const rest = cell.preservedAttrs.replace(VM_ATTRIBUTE, "");
    if (rest.length === 0) {
      delete cell.preservedAttrs;
    } else {
      cell.preservedAttrs = rest;
    }
  }
}

const ROW_LAYOUT_CACHE_ATTRIBUTE = /^(?:spans|(?:[\w.-]+:)?dyDescent)$/;
const SOURCE_ATTRIBUTE = /([^\s=]+)\s*=\s*(?:"[^"]*"|'[^']*')/g;

/** Whether an xlsx `<row>` attribute is one of Excel's layout caches
 * (`spans`, `x14ac:dyDescent`) rather than a property: the xlsx writer drops
 * them (SPEC-122). */
export function isRowLayoutCacheAttribute(name: string): boolean {
  return ROW_LAYOUT_CACHE_ATTRIBUTE.test(name);
}

/** Whether a row's `sourceAttrs` hold an attribute other than the layout
 * caches — an `outlineLevel`, `collapsed`, `thickBot` …: a property the xlsx
 * writer re-emits, so a row without cells carrying one is stored
 * (SPEC-122). */
export function rowSourceAttrsHoldProperties(sourceAttrs: string | undefined): boolean {
  if (sourceAttrs === undefined) {
    return false;
  }
  for (const match of sourceAttrs.matchAll(SOURCE_ATTRIBUTE)) {
    if (!isRowLayoutCacheAttribute(match[1]!)) {
      return true;
    }
  }
  return false;
}

/** Whether a row is displayed hidden (SPEC-128): hidden by the user or by
 * the sheet's autofilter. */
export function rowIsHidden(row: Pick<Row, "hidden" | "filterHidden"> | undefined): boolean {
  return row !== undefined && (row.hidden === true || row.filterHidden === true);
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
