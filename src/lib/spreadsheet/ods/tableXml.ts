// One .ods table's cells and layout (SPEC-116 "Reading .ods", SPEC-117
// "Reading ODS sheet layout"): the per-table half of the reader's single
// forward-only scan of content.xml. reader.ts dispatches the scanner's
// events (and turns each cell's paragraphs into a value and runs);
// OdsTableReader turns a table's `table:table-column`, `table:table-row`
// and cell elements into the sheet's rows, column spans, default sizes and
// merges.
//
// - Columns. Each `table:table-column` run (`number-columns-repeated`) has a
//   width (`style:column-width` of its style, any ODF length → px at 96 dpi,
//   unrounded), `table:visibility` collapse/filter → hidden, and
//   `table:default-cell-style-name` → styleId. `defaultColWidthPx` is the
//   width of the style on the table's LAST column run (LibreOffice's
//   trailing repeated column), else the table-column default-style's, else
//   85.33 px; a run equal to the default with no other property stores no
//   span ("Default" as a default cell style is style 0, no property).
// - Rows. `style:row-height` → heightPt, `style:use-optimal-row-height=
//   "false"` → customHeight, `table:visibility="collapse"` → hidden and
//   SPEC-128's `"filter"` → filterHidden (as the wire says: no
//   classification), a default cell style → styleId. `defaultRowHeightPt`
//   is the height of the style on the table's LAST row run, else 12.8 pt,
//   and is only known when the table ends — so a row holding no cell is
//   kept as a pending run (never expanded) and materialised then only when
//   it differs from the default (a height of its own, either visibility, or
//   a style) — a filter-hidden blank row is a Row with no cells. LibreOffice's trailing 1048553-row run IS
//   the default and costs nothing. Every row object materialised for its
//   properties alone counts against MAX_WORKBOOK_CELLS, as does every merge,
//   so a few bytes of repeat counts can never allocate millions of objects.
// - Cells. A value (from reader.ts) is stored, every repeated copy counted
//   against MAX_WORKBOOK_CELLS before it is built; past the grid it is a
//   parse error naming its address. A cell's `table:style-name` becomes its
//   styleId. A styled EMPTY run stores nothing where its style equals the
//   position's effective default (the row's default style, else the
//   column's); a run reaching the last column (XFD) becomes the row's
//   styleId when the row has no default style of its own — the positions
//   left of it that inherited a column default keep it explicitly, unless
//   that would cost more cells than storing the run itself; every other
//   styled empty position is stored (`{ col, value: null, styleId }`) and
//   counted like a value. LibreOffice's 16380-column trailing runs, styled
//   or not, store no cell.
// - SPEC-118: a date, time or percentage value whose style names no data
//   style points at a derived style with its implied format (the style the
//   position shows — its own, else the row's or column's default — is the
//   base; OdsStyleBook.impliedStyleId decides).
// - Covered cells (`table:covered-table-cell`, under a span) now follow the
//   same rules: one carrying a value or a style is stored (rendering hides
//   it under the merge; SPEC-122 re-serialises it), one with neither stores
//   nothing.
// - Spans (`table:number-columns-spanned` / `-rows-spanned`) are merge
//   candidates in document order, clamped to the grid; `acceptMerges`
//   decides (two or more cells, overlapping no earlier merge), run per
//   band of rows no merge crosses so a long sheet is never quadratic.
// - SPEC-119: a formula cell (reader.ts translates `table:formula`) is
//   stored even with no value; every repetition — repeated columns or a
//   repeated row — gets the formula text VERBATIM (ODF formula text names
//   absolute addresses, so it is never shifted), and a matrix formula's
//   `table:number-matrix-columns-spanned`/`-rows-spanned` become each
//   repetition's own `arrayRange`, clamped to the grid.
import { DEFAULT_THEME, resolveColor } from "../../core/theme";
import { scanAttr, scanXml, type XmlScanAttribute } from "../../core/xmlScanner";
import { sanitizeXmlText } from "../../core/xmlText";
import { TABLE_NAMESPACE } from "../../core/odfXml";
import type { GridRange } from "../address";
import { buildGridMetrics, ODS_DEFAULT_COL_WIDTH_PX, ODS_DEFAULT_ROW_HEIGHT_PT } from "../geometry";
import { acceptMerges } from "../merges";
import { toOdsFormula, type FormulaFields } from "../formula/dialect";
import { expandStructuredReferences } from "../tables";
import { styleOverLook, tableCellLook } from "../tableStyles";
import {
  isCellError,
  MAX_COLS,
  MAX_ROWS,
  type Cell,
  type CellHyperlink,
  type CellNote,
  type CellStyle,
  type CellValue,
  type ColumnSpan,
  type Row,
  type Sheet,
  type SheetDrawing,
  type SheetTable,
  type SpreadsheetModel,
  type TextRun,
} from "../model";
import { formatCodeOf } from "../numberFormat/builtin";
import { SPREADSHEET_LOCALES, type SpreadsheetLocale } from "../numberFormat/locale";
import { LCID_SYSTEM_LONG_DATE, LCID_SYSTEM_LONG_TIME, parseFormatCode, type FormatSection } from "../numberFormat/parse";
import { formatValue, plainText, workbookFormatContext, type FormatContext } from "../numberFormat/render";
import { addressText, type CellBudget, outOfGridError } from "../parsed";
import { serialToOdsDateValue, serialToOdsTimeValue } from "../serial";
import { cellStyle, columnSpanAt } from "../styles";
import { sheetValidations, type DataValidation } from "../validation";
import { createSliceClock, type SliceClock } from "../serialiseSlices";
import { withoutValidationName } from "./contentValidations";
import { leadingWithShapes, odsSheetDrawingsXml, type OdsDrawingWriteContext, type OdsSheetDrawingsXml } from "./drawingXml";
import { linkedParagraphsXml, odsNoteXml, paragraphsMatchLink } from "./annotations";
import { tableStartTag } from "./package";
import type { OdsRowHeight, OdsStyleEnv, OdsStyleRegistry } from "./styles";

/** What the table reader needs from the workbook's styles (OdsStyleBook). */
export interface OdsLayoutStyles {
  /** The model style index of a table-cell style name ("Default" → 0). */
  cellStyleId(name: string): number;
  /** A column style's width in px, undefined when it (and the default
   * style) set none. */
  columnWidthPx(styleName: string | undefined): number | undefined;
  rowHeight(styleName: string | undefined): OdsRowHeight;
  /** The style a value with an implied format `code` shows over the base
   * style `baseId` (SPEC-118 "Implied formats"). */
  impliedStyleId(baseId: number, code: string): number;
}

/** A cell's formula as reader.ts read it (SPEC-119): the model's fields and
 * the matrix extent its spans declare. */
export interface OdsCellFormula extends FormulaFields {
  matrix?: { cols: number; rows: number };
}

/** A cell's captures (SPEC-122 "Model additions"), as reader.ts reads
 * them: the cell's `preservedAttrs`, `preservedXml` and `preservedText`. */
export interface OdsCellPreserved {
  attrs?: string;
  xml?: string;
  text?: string;
}

/** A table's layout, as the sheet stores it. */
export interface OdsTableLayout {
  rows: Row[];
  columns: ColumnSpan[];
  defaultColWidthPx: number;
  defaultRowHeightPt: number;
  merges: GridRange[];
}

/** A repeat count attribute (`table:number-rows-repeated`, …): 1 when absent
 * or unreadable. */
export function repeatCount(attrs: readonly XmlScanAttribute[], local: string): number {
  const value = Number(scanAttr(attrs, local, TABLE_NAMESPACE));
  return Number.isInteger(value) && value > 1 ? value : 1;
}

function isHidden(attrs: readonly XmlScanAttribute[]): boolean {
  const visibility = scanAttr(attrs, "visibility", TABLE_NAMESPACE);
  return visibility === "collapse" || visibility === "filter";
}

/** A row's `table:visibility` as the model holds it (SPEC-128): `collapse`
 * is the user's hidden, `filter` the autofilter's; undefined when visible. */
type RowVisibility = "collapse" | "filter" | undefined;

function rowVisibility(attrs: readonly XmlScanAttribute[]): RowVisibility {
  const visibility = scanAttr(attrs, "visibility", TABLE_NAMESPACE);
  return visibility === "collapse" || visibility === "filter" ? visibility : undefined;
}

interface ColumnRun {
  min: number;
  max: number;
  hidden: boolean;
  styleName?: string;
  styleId?: number;
}

/** A merge relative to its row: the anchor column and the extent. */
interface RowMerge {
  left: number;
  right: number;
  rows: number;
}

interface PendingRow {
  index: number;
  repeat: number;
  visibility: RowVisibility;
  styleName?: string;
  /** The row's default style (`table:default-cell-style-name`), or the
   * style of a styled run that reached the last column. */
  styleId?: number;
  /** Whether styleId is the row's own default style. */
  ownDefault: boolean;
  cells: Cell[];
  /** The 0-based column the next cell lands on. */
  col: number;
  merges: RowMerge[];
}

/** One `table:table-row` element as read: its stored rows (when it held a
 * cell), or a pending run decided when the default height is known. */
interface RowEntry {
  index: number;
  repeat: number;
  visibility: RowVisibility;
  styleName?: string;
  styleId?: number;
  rows?: Row[];
}

function copyRun(run: TextRun): TextRun {
  return run.font ? { text: run.text, font: { ...run.font } } : { text: run.text };
}

/** A repeated row's copy of a cell, `rowOffset` rows below it: a matrix
 * formula's range moves with it. */
function copyCell(cell: Cell, rowOffset: number): Cell {
  const copy: Cell = cell.runs ? { ...cell, runs: cell.runs.map(copyRun) } : { ...cell };
  if (cell.arrayRange) {
    copy.arrayRange = {
      ...cell.arrayRange,
      top: cell.arrayRange.top + rowOffset,
      bottom: Math.min(cell.arrayRange.bottom + rowOffset, MAX_ROWS - 1),
    };
  }
  return copy;
}

/** acceptMerges over candidates sorted by top row, run per band of rows: a
 * band ends where the next candidate starts below every earlier one's
 * bottom, so no candidate can meet a merge of an earlier band — the result
 * is acceptMerges' own, without comparing every pair of a long sheet. */
export function acceptMergesByBand(candidates: readonly GridRange[]): GridRange[] {
  const kept: GridRange[] = [];
  let band: GridRange[] = [];
  let bandBottom = -1;
  for (const range of candidates) {
    if (band.length > 0 && range.top > bandBottom) {
      kept.push(...acceptMerges(band));
      band = [];
    }
    band.push(range);
    bandBottom = band.length === 1 ? range.bottom : Math.max(bandBottom, range.bottom);
  }
  if (band.length > 0) {
    kept.push(...acceptMerges(band));
  }
  return kept;
}

/** Gives a cell its captures (the strings are shared: they are immutable). */
function withPreserved(cell: Cell, preserved: OdsCellPreserved): Cell {
  if (preserved.attrs !== undefined) {
    cell.preservedAttrs = preserved.attrs;
  }
  if (preserved.xml !== undefined) {
    cell.preservedXml = preserved.xml;
  }
  if (preserved.text !== undefined) {
    cell.preservedText = preserved.text;
  }
  return cell;
}

/** Reads one table, element by element, as reader.ts's scan meets them. */
export class OdsTableReader {
  private readonly name: string;
  private readonly styles: OdsLayoutStyles;
  private readonly budget: CellBudget;
  private readonly columnRuns: ColumnRun[] = [];
  private nextColumn = 0;
  private sawColumn = false;
  private lastColumnStyle: string | undefined;
  /** The finished spans — set at the first row (columns precede rows). */
  private columns: ColumnSpan[] | null = null;
  /** The spans that carry a default cell style, for the effective-default
   * lookups. */
  private styledColumns: ColumnSpan[] = [];
  private defaultColWidthPx = ODS_DEFAULT_COL_WIDTH_PX;
  private readonly entries: RowEntry[] = [];
  private readonly mergeCandidates: GridRange[] = [];
  private row: PendingRow | null = null;
  private nextRow = 0;
  private sawRow = false;
  private lastRowStyle: string | undefined;

  constructor(name: string, styles: OdsLayoutStyles, budget: CellBudget) {
    this.name = name;
    this.styles = styles;
    this.budget = budget;
  }

  /** The 0-based row the current (or next) row starts at. */
  get rowIndex(): number {
    return this.row?.index ?? this.nextRow;
  }

  /** The 0-based column the current row's next cell lands on. */
  get colIndex(): number {
    return this.row?.col ?? 0;
  }

  /** A `table:table-column` element. */
  column(attrs: readonly XmlScanAttribute[]): void {
    if (this.columns !== null) {
      return; // a column after the rows is malformed — its runs are settled
    }
    const repeat = repeatCount(attrs, "number-columns-repeated");
    const min = this.nextColumn;
    this.nextColumn += repeat;
    const styleName = scanAttr(attrs, "style-name", TABLE_NAMESPACE);
    this.sawColumn = true;
    this.lastColumnStyle = styleName;
    if (min >= MAX_COLS) {
      return;
    }
    const run: ColumnRun = { min, max: Math.min(min + repeat, MAX_COLS) - 1, hidden: isHidden(attrs) };
    if (styleName !== undefined) {
      run.styleName = styleName;
    }
    const defaultCellStyle = scanAttr(attrs, "default-cell-style-name", TABLE_NAMESPACE);
    if (defaultCellStyle !== undefined) {
      const id = this.styles.cellStyleId(defaultCellStyle);
      if (id !== 0) {
        run.styleId = id;
      }
    }
    this.columnRuns.push(run);
  }

  private finishColumns(): ColumnSpan[] {
    if (this.columns !== null) {
      return this.columns;
    }
    this.defaultColWidthPx =
      this.styles.columnWidthPx(this.sawColumn ? this.lastColumnStyle : undefined) ?? ODS_DEFAULT_COL_WIDTH_PX;
    const spans: ColumnSpan[] = [];
    for (const run of this.columnRuns) {
      const span: ColumnSpan = { min: run.min, max: run.max };
      const width = this.styles.columnWidthPx(run.styleName);
      if (width !== undefined && width !== this.defaultColWidthPx) {
        span.widthPx = width;
      }
      if (run.hidden) {
        span.hidden = true;
      }
      if (run.styleId !== undefined) {
        span.styleId = run.styleId;
      }
      if (span.widthPx === undefined && span.hidden === undefined && span.styleId === undefined) {
        continue;
      }
      if (run.styleName !== undefined) {
        span.sourceStyleName = run.styleName;
      }
      spans.push(span);
    }
    this.columns = spans;
    this.styledColumns = spans.filter((span) => span.styleId !== undefined);
    return spans;
  }

  /** Calls `visit(a, b, styleId)` for each stretch of [from, to] sharing one
   * column default cell style (0 where no column sets one). */
  private eachColumnDefault(from: number, to: number, visit: (a: number, b: number, styleId: number) => void): void {
    const spans = this.styledColumns;
    let lo = 0;
    let hi = spans.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if (spans[mid]!.max < from) {
        lo = mid + 1;
      } else {
        hi = mid;
      }
    }
    let at = from;
    for (let i = lo; i < spans.length && at <= to; i += 1) {
      const span = spans[i]!;
      if (span.min > at) {
        visit(at, Math.min(span.min - 1, to), 0);
        at = span.min;
      }
      if (at <= to) {
        const end = Math.min(span.max, to);
        visit(at, end, span.styleId!);
        at = end + 1;
      }
    }
    if (at <= to) {
      visit(at, to, 0);
    }
  }

  /** The default cell style of column `col`. */
  private columnDefault(col: number): number {
    let found = 0;
    this.eachColumnDefault(col, col, (_a, _b, styleId) => {
      found = styleId;
    });
    return found;
  }

  /** A `table:table-row` element opening. */
  startRow(attrs: readonly XmlScanAttribute[]): void {
    this.finishColumns();
    const styleName = scanAttr(attrs, "style-name", TABLE_NAMESPACE);
    const row: PendingRow = {
      index: this.nextRow,
      repeat: repeatCount(attrs, "number-rows-repeated"),
      visibility: rowVisibility(attrs),
      ownDefault: false,
      cells: [],
      col: 0,
      merges: [],
    };
    if (styleName !== undefined) {
      row.styleName = styleName;
    }
    const defaultCellStyle = scanAttr(attrs, "default-cell-style-name", TABLE_NAMESPACE);
    if (defaultCellStyle !== undefined) {
      const id = this.styles.cellStyleId(defaultCellStyle);
      // "Default" says something only where a column default would apply.
      if (id !== 0 || this.styledColumns.length > 0) {
        row.styleId = id;
        row.ownDefault = true;
      }
    }
    this.row = row;
    this.sawRow = true;
    this.lastRowStyle = styleName;
  }

  /** One cell element of the current row, repeated as declared: `value` is
   * its value (null for none), `runs` its rich text, `impliedCode` the
   * format its value type implies when its style gives none (SPEC-118),
   * `formula` its formula (SPEC-119) and `preserved` its captures (SPEC-122:
   * a cell holding only those is stored too). Returns false when the cell
   * ceiling stops the read. */
  cell(
    attrs: readonly XmlScanAttribute[],
    covered: boolean,
    value: CellValue,
    runs?: TextRun[],
    impliedCode?: string,
    formula?: OdsCellFormula,
    preserved?: OdsCellPreserved,
  ): boolean {
    const row = this.row!;
    const repeat = repeatCount(attrs, "number-columns-repeated");
    const first = row.col;
    row.col += repeat;
    const styleName = scanAttr(attrs, "style-name", TABLE_NAMESPACE);
    const styleId = styleName === undefined ? undefined : this.styles.cellStyleId(styleName);
    if (!covered) {
      this.collectSpan(row, attrs, first, repeat);
    }
    if (value !== null || formula !== undefined) {
      return this.storeValues(row, first, repeat, value, runs, styleId, impliedCode, formula, preserved);
    }
    if (preserved) {
      return this.storePreserved(row, first, repeat, styleId, preserved);
    }
    if (styleId !== undefined) {
      return this.storeStyled(row, first, repeat, styleId);
    }
    return true;
  }

  private collectSpan(row: PendingRow, attrs: readonly XmlScanAttribute[], first: number, repeat: number): void {
    const cols = repeatCount(attrs, "number-columns-spanned");
    const rows = repeatCount(attrs, "number-rows-spanned");
    if (cols === 1 && rows === 1) {
      return;
    }
    for (let k = 0; k < repeat && first + k < MAX_COLS; k += 1) {
      row.merges.push({ left: first + k, right: Math.min(first + k + cols, MAX_COLS) - 1, rows });
    }
  }

  private storeValues(
    row: PendingRow,
    first: number,
    repeat: number,
    value: CellValue,
    runs: TextRun[] | undefined,
    styleId: number | undefined,
    impliedCode: string | undefined,
    formula?: OdsCellFormula,
    preserved?: OdsCellPreserved,
  ): boolean {
    if (row.index >= MAX_ROWS) {
      throw outOfGridError(this.name, addressText(row.index, first));
    }
    if (first + repeat > MAX_COLS) {
      throw outOfGridError(this.name, addressText(row.index, Math.max(first, MAX_COLS)));
    }
    if (!this.budget.take(repeat)) {
      return false;
    }
    for (let k = 0; k < repeat; k += 1) {
      const cell: Cell = { col: first + k, value };
      if (runs) {
        cell.runs = k === 0 ? runs : runs.map(copyRun);
      }
      if (styleId !== undefined) {
        cell.styleId = styleId;
      }
      if (impliedCode !== undefined) {
        const base = styleId ?? (row.ownDefault ? row.styleId! : this.columnDefault(first + k));
        const shown = this.styles.impliedStyleId(base, impliedCode);
        if (shown !== base) {
          cell.styleId = shown;
        }
      }
      if (formula) {
        cell.formula = formula.formula;
        if (formula.formulaUnparsed) {
          cell.formulaUnparsed = true;
        }
        if (formula.matrix) {
          const left = first + k;
          cell.arrayRange = {
            top: row.index,
            left,
            bottom: Math.min(row.index + formula.matrix.rows, MAX_ROWS) - 1,
            right: Math.min(left + formula.matrix.cols, MAX_COLS) - 1,
          };
        }
      }
      if (preserved) {
        withPreserved(cell, preserved);
      }
      row.cells.push(cell);
    }
    return true;
  }

  /** A cell (run) holding only preserved content (SPEC-122): stored, each
   * repetition counted; past the grid it has nothing to show. */
  private storePreserved(row: PendingRow, first: number, repeat: number, styleId: number | undefined, preserved: OdsCellPreserved): boolean {
    if (row.index >= MAX_ROWS || first >= MAX_COLS) {
      return true;
    }
    const last = Math.min(first + repeat, MAX_COLS) - 1;
    if (!this.budget.take(last - first + 1)) {
      return false;
    }
    for (let col = first; col <= last; col += 1) {
      const cell: Cell = { col, value: null };
      if (styleId !== undefined) {
        cell.styleId = styleId;
      }
      row.cells.push(withPreserved(cell, preserved));
    }
    return true;
  }

  /** A styled empty run (see the module header). */
  private storeStyled(row: PendingRow, first: number, repeat: number, styleId: number): boolean {
    if (row.index >= MAX_ROWS || first >= MAX_COLS) {
      return true; // a style past the grid has nothing to show
    }
    const last = Math.min(first + repeat, MAX_COLS) - 1;
    if (row.ownDefault) {
      return styleId === row.styleId ? true : this.storeStyleOnly(row, first, last, styleId, () => true, last - first + 1);
    }
    let differing = 0;
    this.eachColumnDefault(first, last, (a, b, columnStyle) => {
      if (columnStyle !== styleId) {
        differing += b - a + 1;
      }
    });
    if (differing === 0) {
      return true;
    }
    if (last === MAX_COLS - 1 && row.styleId === undefined) {
      const kept = this.inheritedLeftOf(row, first, styleId);
      if (kept.length < differing) {
        return this.promote(row, first, styleId, kept);
      }
    }
    return this.storeStyleOnly(row, first, last, styleId, (columnStyle) => columnStyle !== styleId, differing);
  }

  /** Stores `{ col, value: null, styleId }` for each position of
   * [first, last] whose column default passes `where` — `count` of them. */
  private storeStyleOnly(
    row: PendingRow,
    first: number,
    last: number,
    styleId: number,
    where: (columnStyle: number) => boolean,
    count: number,
  ): boolean {
    if (!this.budget.take(count)) {
      return false;
    }
    this.eachColumnDefault(first, last, (a, b, columnStyle) => {
      if (where(columnStyle)) {
        for (let col = a; col <= b; col += 1) {
          row.cells.push({ col, value: null, styleId });
        }
      }
    });
    return true;
  }

  /** The style-only cells a row style `styleId` set from column `from` on
   * would need left of `from`, where nothing is stored and the position's
   * column default (which it showed) differs from `styleId`. */
  private inheritedLeftOf(row: PendingRow, from: number, styleId: number): Cell[] {
    const needed: Cell[] = [];
    let next = 0; // the first stored cell not yet passed (cells ascend by col)
    this.eachColumnDefault(0, from - 1, (a, b, columnStyle) => {
      if (columnStyle === styleId) {
        return;
      }
      for (let col = a; col <= b; col += 1) {
        while (next < row.cells.length && row.cells[next]!.col < col) {
          next += 1;
        }
        if (next >= row.cells.length || row.cells[next]!.col !== col) {
          needed.push({ col, value: null, styleId: columnStyle });
        }
      }
    });
    return needed;
  }

  /** Makes `styleId` the row's style: the cells left of `from` that showed
   * their column default keep it explicitly (`kept` are the style-only cells
   * the unstored positions need). */
  private promote(row: PendingRow, from: number, styleId: number, kept: Cell[]): boolean {
    if (kept.length > 0 && !this.budget.take(kept.length)) {
      return false;
    }
    for (const cell of row.cells) {
      if (cell.styleId === undefined && cell.col < from) {
        const columnStyle = this.columnDefault(cell.col);
        if (columnStyle !== styleId) {
          cell.styleId = columnStyle;
        }
      }
    }
    if (kept.length > 0) {
      row.cells = [...row.cells, ...kept].toSorted((a, b) => a.col - b.col);
    }
    row.styleId = styleId;
    return true;
  }

  /** The current row closing. Returns false when the cell ceiling stops the
   * read. */
  endRow(): boolean {
    const row = this.row!;
    this.row = null;
    const start = row.index;
    this.nextRow += row.repeat;
    const firstValue = row.cells.find((cell) => cell.value !== null);
    if (firstValue && start + row.repeat > MAX_ROWS) {
      throw outOfGridError(this.name, addressText(MAX_ROWS, firstValue.col));
    }
    // Styles and spans past the grid have nothing to show.
    const repeat = Math.min(row.repeat, MAX_ROWS - start);
    if (repeat <= 0) {
      return true;
    }
    if (row.merges.length > 0) {
      if (!this.budget.take(row.merges.length * repeat)) {
        return false;
      }
      for (let k = 0; k < repeat; k += 1) {
        const top = start + k;
        for (const merge of row.merges) {
          this.mergeCandidates.push({ top, left: merge.left, bottom: Math.min(top + merge.rows, MAX_ROWS) - 1, right: merge.right });
        }
      }
    }
    const entry: RowEntry = { index: start, repeat, visibility: row.visibility };
    if (row.styleName !== undefined) {
      entry.styleName = row.styleName;
    }
    if (row.styleId !== undefined) {
      entry.styleId = row.styleId;
    }
    if (row.cells.length > 0) {
      if (repeat > 1 && !this.budget.take((repeat - 1) * row.cells.length)) {
        return false;
      }
      const rows: Row[] = [];
      for (let k = 0; k < repeat; k += 1) {
        rows.push(this.rowObject(entry, start + k, k === 0 ? row.cells : row.cells.map((cell) => copyCell(cell, k))));
      }
      entry.rows = rows;
    }
    this.entries.push(entry);
    return true;
  }

  private rowObject(entry: RowEntry, index: number, cells: Cell[]): Row {
    const row: Row = { index, cells };
    if (entry.visibility === "collapse") {
      row.hidden = true;
    } else if (entry.visibility === "filter") {
      row.filterHidden = true;
    }
    if (entry.styleId !== undefined) {
      row.styleId = entry.styleId;
    }
    if (entry.styleName !== undefined) {
      row.sourceStyleName = entry.styleName;
    }
    return row;
  }

  /** The table closing: the layout, or null when the cell ceiling stopped
   * the read. */
  finish(): OdsTableLayout | null {
    const columns = this.finishColumns();
    const defaultRowHeightPt =
      this.styles.rowHeight(this.sawRow ? this.lastRowStyle : undefined).heightPt ?? ODS_DEFAULT_ROW_HEIGHT_PT;
    const rows: Row[] = [];
    for (const entry of this.entries) {
      const height = this.styles.rowHeight(entry.styleName);
      const ownHeight = height.heightPt !== undefined && height.heightPt !== defaultRowHeightPt ? height : undefined;
      let stored = entry.rows;
      if (!stored) {
        if (!ownHeight && entry.visibility === undefined && entry.styleId === undefined) {
          continue; // equal to the default: nothing to store
        }
        if (!this.budget.take(entry.repeat)) {
          return null;
        }
        stored = [];
        for (let k = 0; k < entry.repeat; k += 1) {
          stored.push(this.rowObject(entry, entry.index + k, []));
        }
      }
      for (const row of stored) {
        if (ownHeight) {
          row.heightPt = ownHeight.heightPt!;
          if (ownHeight.customHeight) {
            row.customHeight = true;
          }
        }
        rows.push(row);
      }
    }
    return {
      rows,
      columns,
      defaultColWidthPx: this.defaultColWidthPx,
      defaultRowHeightPt,
      merges: acceptMergesByBand(this.mergeCandidates),
    };
  }
}

// --- the table writer (SPEC-122 "The table writer") -------------------------------------

/** What serializeTable needs beyond the sheet. */
export interface OdsTableContext {
  /** The sheet's retained table text (sourceXml, its leading whitespace
   * included) with the writer's sheet-reference pass and SPEC-121's
   * `table:name` / `table:style-name` splices already applied; undefined for
   * a new sheet, whose start tag is generated. */
  source?: string;
  /** SPEC-133: with `sheets` (and `names`), structured references in
   *  `table:formula` are written as their A1 ranges. */
  model: Pick<SpreadsheetModel, "styles" | "theme" | "dateSystem" | "odsNullDate" | "format" | "indexedColors"> & Partial<Pick<SpreadsheetModel, "sheets" | "names">>;
  registry: OdsStyleRegistry;
  /** The user's number locale: the informative `text:p` and a
   * locale-dependent built-in's generated number style are written in it. */
  locale: SpreadsheetLocale;
  /** Whether the target package holds the source's styles (the substitution
   * path): a style, row or column read from the file then keeps its name.
   * From scratch every style is generated. */
  sourceStyleNames?: boolean;
  /** Whether styles.xml defines a common table-cell style (a generated
   * style's parent, for a named style the workbook does not define). */
  commonCellStyle?: (name: string) => boolean;
  /** SPEC-123: the common style a named style the workbook defines is
   * written under (a generated style's parent; OdsNamedStyles.parent). */
  namedStyleParent?: (name: string) => string | undefined;
  /** The source's column and row styles as the reader measured them — how
   * the source's default (last-run) column and row styles are recognised and
   * reused for default-sized columns and rows. */
  layoutStyles?: Pick<OdsLayoutStyles, "columnWidthPx" | "rowHeight">;
  /** The retained-text sheet-reference pass, for a cell's preserved content
   * (source-time text, like sourceXml). */
  passRetained?: (text: string) => string;
  /** SPEC-129: the save's conditional-format writer
   * (ods/conditionalFormats.ts), which writes the table's
   * `calcext:conditional-formats` from the model into its leading or
   * trailing children; absent, they are re-emitted as they are. */
  conditionalFormats?: {
    splice(sheet: Sheet, leading: string, trailing: string, passRetained?: (text: string) => string): { leading: string; trailing: string };
  };
  /** SPEC-129: whether the source's automatic cell style of this name
   * carries a `style:map` — such a style is never named by a rebuilt
   * table: the cells using it get a generated `de-ce` style without the
   * map, so a rule deleted from the model cannot come back. */
  mappedCellStyle?: (name: string) => boolean;
  /** SPEC-130: the name a sheet's validation rule is written under
   * (OdsValidationWriter.nameOf — the save's `table:content-validations`
   * defines it); absent, or undefined for a rule, no cell names it. */
  validationName?: (sheet: Sheet, rule: DataValidation) => string | undefined;
  /** SPEC-131: the package path a drawing's media id is written under
   * (ods/drawingXml.ts OdsMediaPaths); absent, the id itself. */
  mediaPath?: (mediaId: string) => string;
  /** SPEC-132: the embedded object the save minted for a chart drawing
   * written as a new object (ods/chartParts.ts) — its canonical frame names
   * it; absent, a created chart is not written. */
  chartObjectDir?: (drawing: SheetDrawing) => string | undefined;
  /** SPEC-133: the table's trailing children with the sheet's own names
   *  (ods/namedExpressions.ts odsTrailingWithNames); absent, they are
   *  re-emitted as they are. */
  localNames?: (sheet: Sheet, trailing: string) => string;
}

/** SPEC-130: a validated rectangle as the writer meets it — the rule's
 * place in model order (the first covering rule wins) and its name. */
interface ValidatedRange {
  range: GridRange;
  order: number;
  name: string;
}

/** A row's validated column segments, ascending and disjoint. */
interface ValidatedSegment {
  left: number;
  right: number;
  name: string;
}

type Axis = "row" | "col";

/** A grouping element of the source (header rows, row groups,
 * `table:table-rows`, and the column equivalents) and the indices it
 * covered at read. */
interface SkeletonGroup {
  axis: Axis;
  qname: string;
  /** The start tag verbatim (the whole element when self-closing). */
  open: string;
  selfClosing: boolean;
  from: number;
  to: number;
  children: SkeletonGroup[];
}

/** A source table as the writer rebuilds it (see serializeTable). */
interface TableSkeleton {
  /** The whitespace before the table and its start tag. */
  head: string;
  leading: string;
  trailing: string;
  close: string;
  rows: number;
  cols: number;
  rowGroups: SkeletonGroup[];
  colGroups: SkeletonGroup[];
  lastRowStyle?: string;
  lastColumnStyle?: string;
  sawRow: boolean;
  sawColumn: boolean;
}

const COLUMN_GROUPS: ReadonlySet<string> = new Set(["table-columns", "table-header-columns", "table-column-group"]);
const ROW_GROUPS: ReadonlySet<string> = new Set(["table-rows", "table-header-rows", "table-row-group"]);

function rawRepeat(attrs: readonly XmlScanAttribute[], qname: string): number {
  const value = Number(attrs.find((attr) => attr.name === qname)?.value);
  return Number.isInteger(value) && value > 1 ? value : 1;
}

/** Reads a source table's skeleton: its start tag, leading and trailing
 * children, grouping elements, extents and last column and row styles. The
 * text is a fragment (its prefixes unbound), so elements are matched by the
 * qualified names the table's own prefix gives. */
function readSkeleton(source: string): TableSkeleton {
  const { start, tag } = tableStartTag(source);
  const qname = /^<([^\s/>]+)/.exec(tag)?.[1] ?? "table:table";
  const prefix = qname.includes(":") ? qname.slice(0, qname.indexOf(":") + 1) : "";
  const skeleton: TableSkeleton = {
    head: source.slice(0, start) + tag.replace(/\s*\/>$/, ">"),
    leading: "",
    trailing: "",
    close: `</${qname}>`,
    rows: 0,
    cols: 0,
    rowGroups: [],
    colGroups: [],
    sawRow: false,
    sawColumn: false,
  };
  if (tag.endsWith("/>")) {
    return skeleton;
  }
  const local = (name: string): string | undefined => (name.startsWith(prefix) ? name.slice(prefix.length) : undefined);
  type Frame = { kind: "table" | "row" | "other" | "group"; top: boolean; group?: SkeletonGroup };
  const stack: Frame[] = [];
  let rowIndex = 0;
  let colIndex = 0;
  let rowWidth = 0;
  let maxRowWidth = 0;
  let firstStart = -1;
  let lastEnd = -1;
  let closeStart = source.length;
  const tagEnd = start + tag.length;
  scanXml(source, (event) => {
    if (event.type === "text") {
      return undefined;
    }
    if (event.type === "close") {
      const frame = stack.pop();
      if (frame?.kind === "table") {
        closeStart = event.start;
        skeleton.close = source.slice(event.start, event.end);
        return "stop";
      }
      if (frame?.kind === "row") {
        maxRowWidth = Math.max(maxRowWidth, rowWidth);
      } else if (frame?.group) {
        frame.group.to = frame.group.axis === "row" ? rowIndex : colIndex;
      }
      if (frame?.top) {
        lastEnd = event.end;
      }
      return undefined;
    }
    const parent = stack.at(-1);
    if (!parent) {
      stack.push({ kind: "table", top: false });
      return undefined;
    }
    let frame: Frame = { kind: "other", top: false };
    if (parent.kind === "row") {
      const name = local(event.name);
      if (name === "table-cell" || name === "covered-table-cell") {
        rowWidth += rawRepeat(event.attrs, `${prefix}number-columns-repeated`);
      }
    } else if (parent.kind !== "other") {
      const name = local(event.name);
      const attr = (attrLocal: string): string | undefined => event.attrs.find((each) => each.name === `${prefix}${attrLocal}`)?.value;
      let structural = true;
      if (name === "table-row") {
        rowIndex += rawRepeat(event.attrs, `${prefix}number-rows-repeated`);
        skeleton.sawRow = true;
        skeleton.lastRowStyle = attr("style-name");
        rowWidth = 0;
        frame = { kind: "row", top: false };
      } else if (name === "table-column") {
        colIndex += rawRepeat(event.attrs, `${prefix}number-columns-repeated`);
        skeleton.sawColumn = true;
        skeleton.lastColumnStyle = attr("style-name");
      } else if (name !== undefined && (ROW_GROUPS.has(name) || COLUMN_GROUPS.has(name))) {
        const axis: Axis = ROW_GROUPS.has(name) ? "row" : "col";
        const at = axis === "row" ? rowIndex : colIndex;
        const group: SkeletonGroup = {
          axis,
          qname: event.name,
          open: source.slice(event.start, event.end),
          selfClosing: event.selfClosing,
          from: at,
          to: at,
          children: [],
        };
        const siblings = parent.group ? parent.group.children : axis === "row" ? skeleton.rowGroups : skeleton.colGroups;
        siblings.push(group);
        frame = { kind: "group", top: false, group };
      } else if (!/(^|:)soft-page-break$/.test(event.name)) {
        structural = false;
      }
      if (structural && parent.kind === "table") {
        if (firstStart < 0) {
          firstStart = event.start;
        }
        frame.top = true;
        if (event.selfClosing) {
          lastEnd = event.end;
        }
      }
    }
    if (!event.selfClosing) {
      stack.push(frame);
    }
    return undefined;
  });
  if (firstStart < 0) {
    skeleton.leading = source.slice(tagEnd, closeStart);
  } else {
    skeleton.leading = source.slice(tagEnd, firstStart);
    skeleton.trailing = source.slice(lastEnd, closeStart);
  }
  skeleton.rows = rowIndex;
  skeleton.cols = Math.max(colIndex, maxRowWidth);
  return skeleton;
}

/** An element with a repeat count spliced after its name. */
interface Repeatable {
  /** `<table:table-cell`, `<table:table-row table:style-name="…"`… */
  head: string;
  /** The rest of the element. */
  tail: string;
}

/** Accumulates runs of identical elements and writes each run once, its
 * count in `attribute`. SPEC-137: a long sheet has no repeats worth
 * compressing (25,000 distinct rows, say), so `flush` runs once per row —
 * `out` is an array of pieces joined ONCE by the `out` getter, never a
 * string grown by repeated `+=` across a sheet's whole row count. */
class RunWriter {
  private current: Repeatable | null = null;
  private count = 0;
  private pieces: string[] = [];

  constructor(private readonly attribute: string) {}

  /** Adds `count` copies of `element`; a `unique` one (SPEC-131: a cell or
   * row holding a drawing) is never compressed with its neighbours. */
  push(element: Repeatable, count: number, unique = false): void {
    if (count <= 0) {
      return;
    }
    if (unique) {
      this.flush();
      this.pieces.push(element.head + element.tail);
      return;
    }
    if (this.current && this.current.head === element.head && this.current.tail === element.tail) {
      this.count += count;
      return;
    }
    this.flush();
    this.current = element;
    this.count = count;
  }

  flush(): void {
    if (this.current) {
      this.pieces.push(this.count > 1 ? `${this.current.head} ${this.attribute}="${this.count}"${this.current.tail}` : this.current.head + this.current.tail);
    }
    this.current = null;
    this.count = 0;
  }

  /** Every piece written so far, joined once. */
  get out(): string {
    return this.pieces.join("");
  }
}

/** An attribute value: XML-escaped, with the whitespace a parser would
 * normalise kept as character references. */
function attributeText(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/\t/g, "&#9;").replace(/\n/g, "&#10;").replace(/\r/g, "&#13;");
}

/** Character data: XML-escaped, a carriage return as `&#13;` — a literal one
 * would reach the reader as a line feed (XML's end-of-line normalisation),
 * so "\r\n" would read back as two line breaks. */
function elementText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r/g, "&#13;");
}

type NumberValueType = "float" | "percentage" | "currency" | "date" | "time";

/** A number's ODS value type, from its format code's first numeric section
 * (the inverse of the data-style family SPEC-118's odsNumberStyleXml writes
 * for it): date and time sections (a section naming a day, month or year is
 * a date), percent, an explicit `[$sym-lcid]` currency, else float. */
function numberValueType(code: string | undefined): { type: NumberValueType; symbol?: string } {
  if (code === undefined) {
    return { type: "float" };
  }
  let sections: readonly FormatSection[];
  try {
    const parsed = parseFormatCode(code);
    sections = parsed.unsupported ? [] : parsed.sections;
  } catch {
    return { type: "float" };
  }
  const section = sections.find((each) => each.kind !== "text" && each.kind !== "literal");
  if (!section) {
    return { type: "float" };
  }
  if (section.kind === "date" || section.lcid === LCID_SYSTEM_LONG_DATE || section.lcid === LCID_SYSTEM_LONG_TIME) {
    const dated =
      section.lcid === LCID_SYSTEM_LONG_DATE ||
      section.tokens.some((token) => token.type === "date" && token.part !== "hour" && token.part !== "minute" && token.part !== "second");
    return { type: dated ? "date" : "time" };
  }
  if (section.tokens.some((token) => token.type === "percent")) {
    return { type: "percentage" };
  }
  const currency = section.tokens.find((token) => token.type === "currency");
  if (currency) {
    return { type: "currency", symbol: currency.symbol };
  }
  return { type: "float" };
}

/** The ISO 4217 code `office:currency` names for a format's currency
 * symbol: the locale whose symbol it is (the user's own first), a symbol
 * that is itself a three-letter code, else the user's locale's code. */
function currencyCodeOf(symbol: string | undefined, locale: SpreadsheetLocale): string {
  if (symbol === undefined || symbol === locale.currencySymbol) {
    return locale.currencyCode;
  }
  const match = Object.values(SPREADSHEET_LOCALES).find((each) => each.currencySymbol === symbol);
  if (match) {
    return match.currencyCode;
  }
  return /^[A-Z]{3}$/.test(symbol) ? symbol : locale.currencyCode;
}

/** One piece of a cell's text and the text style it is written in. */
interface TextPiece {
  text: string;
  style?: string;
}

/** A line of text as paragraph content: XML-escaped, a run of spaces as
 * `<text:s text:c="n"/>` (`<text:s/>` for one) — a single space between
 * other characters stays literal — and tabs as `<text:tab/>`, each piece in
 * its `text:span`. Runs of spaces are counted across pieces. */
function lineXml(pieces: readonly TextPiece[]): string {
  const full = pieces.map((piece) => piece.text).join("");
  let out = "";
  let at = 0; // offset of the current piece in `full`
  for (const piece of pieces) {
    let body = "";
    let i = 0;
    while (i < piece.text.length) {
      const ch = piece.text[i]!;
      if (ch === " ") {
        let end = i;
        while (end < piece.text.length && piece.text[end] === " ") {
          end += 1;
        }
        let runStart = at + i;
        while (runStart > 0 && full[runStart - 1] === " ") {
          runStart -= 1;
        }
        let runEnd = at + end;
        while (runEnd < full.length && full[runEnd] === " ") {
          runEnd += 1;
        }
        const count = end - i;
        if (runEnd - runStart === 1 && runStart > 0 && runEnd < full.length) {
          body += " ";
        } else {
          body += count === 1 ? "<text:s/>" : `<text:s text:c="${count}"/>`;
        }
        i = end;
      } else if (ch === "\t") {
        body += "<text:tab/>";
        i += 1;
      } else {
        let end = i;
        while (end < piece.text.length && piece.text[end] !== " " && piece.text[end] !== "\t") {
          end += 1;
        }
        body += elementText(piece.text.slice(i, end));
        i = end;
      }
    }
    at += piece.text.length;
    if (body !== "") {
      out += piece.style === undefined ? body : `<text:span text:style-name="${piece.style}">${body}</text:span>`;
    }
  }
  return out;
}

/** Text as `text:p` paragraphs, one per "\n"-separated line ("" → none),
 * which the reader joins back with "\n". XML-illegal characters are
 * dropped (sanitizeXmlText: U+000B/U+000C become line breaks); a carriage
 * return is kept, as a character reference (elementText). */
function paragraphsXml(pieces: readonly TextPiece[]): string {
  const clean = pieces.map((piece) => ({ ...piece, text: sanitizeXmlText(piece.text) }));
  if (clean.every((piece) => piece.text === "")) {
    return "";
  }
  const lines: TextPiece[][] = [[]];
  for (const piece of clean) {
    piece.text.split("\n").forEach((part, index) => {
      if (index > 0) {
        lines.push([]);
      }
      if (part !== "") {
        lines.at(-1)!.push({ ...piece, text: part });
      }
    });
  }
  return lines
    .map((line) => {
      const body = lineXml(line);
      return body === "" ? "<text:p/>" : `<text:p>${body}</text:p>`;
    })
    .join("");
}

/** A merge a written row meets, and whether the row is its top. */
interface MergeInRow {
  range: GridRange;
  anchor: boolean;
}

/** Writes one sheet's table (see serializeTable). */
class TableWriter {
  private readonly sheet: Sheet;
  private readonly ctx: OdsTableContext;
  private readonly skeleton: TableSkeleton | null;
  private readonly styleNames = new Map<number, string>();
  private readonly valueTypes = new Map<string, { type: NumberValueType; symbol?: string }>();
  private readonly format: FormatContext;
  private readonly env: OdsStyleEnv;
  private defaultColumn: string | undefined;
  private defaultRow: string | undefined;
  rowExtent = 1;
  colExtent = 1;
  // Row iteration state (rows are written in order, segment by segment).
  private nextStored = 0;
  private nextMerge = 0;
  private active: GridRange[] = [];
  private readonly merges: GridRange[];
  // SPEC-130: the validated rectangles by top, those the current row meets,
  // and the rows where that set changes (a run of empty rows splits there).
  private readonly validated: ValidatedRange[];
  private nextValidated = 0;
  private activeValidated: ValidatedRange[] = [];
  private readonly validatedBreaks: number[];
  private nextBreak = 0;
  // SPEC-131: the drawings, written from the model (ods/drawingXml.ts).
  private readonly drawings: OdsSheetDrawingsXml;
  private readonly drawingRows: number[];
  private nextDrawingRow = 0;
  // SPEC-135: the notes by row and column (each its cell's first child), the
  // rows holding one, and the hyperlinks by top with those the current row
  // meets (ods/annotations.ts).
  private readonly notes = new Map<number, Map<number, CellNote>>();
  private readonly noteRows: number[];
  private nextNoteRow = 0;
  private readonly links: CellHyperlink[];
  private nextLink = 0;
  private activeLinks: CellHyperlink[] = [];
  // SPEC-133: the sheet's tables, and each look's generated style by own style id.
  private readonly tables: readonly SheetTable[];
  private readonly lookStyles = new Map<CellStyle, Map<number, string | null>>();
  private readonly workbookHasTables: boolean;

  constructor(sheet: Sheet, ctx: OdsTableContext) {
    this.sheet = sheet;
    this.ctx = ctx;
    this.skeleton = ctx.source === undefined ? null : readSkeleton(ctx.source);
    this.format = workbookFormatContext(ctx.model, ctx.locale);
    this.env = { theme: ctx.model.theme ?? DEFAULT_THEME, locale: ctx.locale };
    if (ctx.commonCellStyle) {
      this.env.commonCellStyle = ctx.commonCellStyle;
    }
    if (ctx.namedStyleParent) {
      this.env.namedStyleParent = ctx.namedStyleParent;
    }
    this.merges = sheet.merges.toSorted((a, b) => a.top - b.top || a.left - b.left);
    this.validated = validatedRanges(sheet, ctx.validationName);
    this.validatedBreaks = [...new Set(this.validated.flatMap(({ range }) => [range.top, range.bottom + 1]))].toSorted((a, b) => a - b);
    const drawingCtx: OdsDrawingWriteContext = { sheetName: sheet.name, metrics: buildGridMetrics(sheet), registry: ctx.registry, theme: this.env.theme };
    if (ctx.passRetained) {
      drawingCtx.passRetained = ctx.passRetained;
    }
    if (ctx.mediaPath) {
      drawingCtx.mediaPath = ctx.mediaPath;
    }
    if (ctx.chartObjectDir) {
      drawingCtx.chartObjectDir = ctx.chartObjectDir;
    }
    this.drawings = odsSheetDrawingsXml(sheet.drawings ?? [], drawingCtx);
    this.drawingRows = [...this.drawings.cells.keys()].toSorted((a, b) => a - b);
    for (const note of sheet.notes ?? []) {
      if (note.ref.row < MAX_ROWS && note.ref.col < MAX_COLS) {
        let byCol = this.notes.get(note.ref.row);
        if (!byCol) {
          byCol = new Map();
          this.notes.set(note.ref.row, byCol);
        }
        byCol.set(note.ref.col, note);
      }
    }
    this.noteRows = [...this.notes.keys()].toSorted((a, b) => a - b);
    this.links = (sheet.hyperlinks ?? []).toSorted((a, b) => a.range.top - b.range.top);
    this.tables = sheet.tables ?? [];
    this.workbookHasTables = ctx.model.sheets?.some((each) => (each.tables?.length ?? 0) > 0) ?? false;
    this.measure();
  }

  /** The extents (see serializeTable). */
  private measure(): void {
    const sheet = this.sheet;
    let rows = this.skeleton?.rows ?? 0;
    let cols = this.skeleton?.cols ?? 0;
    for (const row of sheet.rows) {
      rows = Math.max(rows, row.index + 1);
      const last = row.cells.at(-1);
      if (last) {
        cols = Math.max(cols, last.col + 1);
      }
    }
    for (const merge of sheet.merges) {
      rows = Math.max(rows, merge.bottom + 1);
      cols = Math.max(cols, merge.right + 1);
    }
    for (const span of sheet.columns) {
      cols = Math.max(cols, span.max + 1);
    }
    // SPEC-130: a validated cell is written even when nothing is stored there.
    for (const { range } of this.validated) {
      rows = Math.max(rows, range.bottom + 1);
      cols = Math.max(cols, range.right + 1);
    }
    // SPEC-133: every cell of a table is written (its look baked in).
    for (const table of this.tables) {
      rows = Math.max(rows, table.ref.bottom + 1);
      cols = Math.max(cols, table.ref.right + 1);
    }
    // SPEC-131: a drawing's anchor cell is written even when nothing is stored there.
    for (const [row, byCol] of this.drawings.cells) {
      rows = Math.max(rows, row + 1);
      for (const col of byCol.keys()) {
        cols = Math.max(cols, col + 1);
      }
    }
    // SPEC-135: so is a note's.
    for (const [row, byCol] of this.notes) {
      rows = Math.max(rows, row + 1);
      for (const col of byCol.keys()) {
        cols = Math.max(cols, col + 1);
      }
    }
    rows = Math.min(Math.max(rows, 1), MAX_ROWS);
    cols = Math.min(Math.max(cols, 1), MAX_COLS);
    // The reader takes the last column and row runs' sizes as the defaults:
    // a sized last one gets a default-sized one after it.
    if (cols < MAX_COLS && this.columnWidth(cols - 1) !== sheet.defaultColWidthPx) {
      cols += 1;
    }
    const lastRow = sheet.rows.at(-1);
    if (rows < MAX_ROWS && lastRow?.index === rows - 1 && lastRow.heightPt !== undefined && lastRow.heightPt !== sheet.defaultRowHeightPt) {
      rows += 1;
    }
    this.rowExtent = rows;
    this.colExtent = cols;
  }

  private columnWidth(col: number): number {
    return columnSpanAt(this.sheet.columns, col)?.widthPx ?? this.sheet.defaultColWidthPx;
  }

  /** A model style's name: "Default" for 0, the source's name for a style
   * read from the file (when its styles are retained), else a generated one. */
  styleName(id: number): string {
    let name = this.styleNames.get(id);
    if (name === undefined) {
      const style = cellStyle(this.ctx.model, id);
      if (id === 0 || style === this.ctx.model.styles[0]) {
        name = "Default";
      } else if (this.ctx.sourceStyleNames && style.sourceStyleName !== undefined && !this.ctx.mappedCellStyle?.(style.sourceStyleName)) {
        name = style.sourceStyleName;
      } else {
        name = this.ctx.registry.cellStyle(style, this.env);
      }
      this.styleNames.set(id, name);
    }
    return name;
  }

  /** SPEC-133: the generated style a position inside a table is written
   * with — its own style (`ownId`) over the table's look (tableStyles.ts
   * styleOverLook, Normal's values counting as inherited) — or undefined
   * outside every table and where the look changes nothing. `model.styles`
   * is never touched: the style is generated per write. */
  private tableStyleName(table: SheetTable | undefined, r: number, col: number, ownId: number): string | undefined {
    const look = table ? tableCellLook(this.sheet, table, r, col, this.env.theme) : undefined;
    if (!look) {
      return undefined;
    }
    let byOwn = this.lookStyles.get(look);
    if (!byOwn) {
      byOwn = new Map();
      this.lookStyles.set(look, byOwn);
    }
    let name = byOwn.get(ownId);
    if (name === undefined) {
      const own = cellStyle(this.ctx.model, ownId);
      const merged = styleOverLook(look, own, this.ctx.model.styles[0]);
      name = merged === own ? null : this.ctx.registry.cellStyle(merged, this.env);
      byOwn.set(ownId, name);
    }
    return name ?? undefined;
  }

  /** The style default-width columns are written with: the source's last
   * column style when it still measures the default, else a generated one. */
  private defaultColumnStyle(): string {
    if (this.defaultColumn === undefined) {
      const skeleton = this.skeleton;
      const name = skeleton?.sawColumn ? skeleton.lastColumnStyle : undefined;
      const layout = this.ctx.layoutStyles;
      this.defaultColumn =
        this.ctx.sourceStyleNames && name !== undefined && layout && (layout.columnWidthPx(name) ?? ODS_DEFAULT_COL_WIDTH_PX) === this.sheet.defaultColWidthPx
          ? name
          : this.ctx.registry.columnStyle(this.sheet.defaultColWidthPx);
    }
    return this.defaultColumn;
  }

  /** The style default-height rows are written with (as columns). */
  private defaultRowStyle(): string {
    if (this.defaultRow === undefined) {
      const skeleton = this.skeleton;
      const name = skeleton?.sawRow ? skeleton.lastRowStyle : undefined;
      const layout = this.ctx.layoutStyles;
      this.defaultRow =
        this.ctx.sourceStyleNames && name !== undefined && layout && (layout.rowHeight(name).heightPt ?? ODS_DEFAULT_ROW_HEIGHT_PT) === this.sheet.defaultRowHeightPt
          ? name
          : this.ctx.registry.rowStyle(this.sheet.defaultRowHeightPt, true);
    }
    return this.defaultRow;
  }

  // --- columns ---

  /** Columns [from, to): one element per run of identical columns. */
  private columns(from: number, to: number, out: RunWriter): void {
    const spans = this.sheet.columns;
    let index = spans.findIndex((span) => span.max >= from);
    let col = from;
    while (col < to) {
      const span = index >= 0 && index < spans.length ? spans[index] : undefined;
      if (span && span.min <= col) {
        const end = Math.min(span.max + 1, to);
        out.push(this.columnElement(span), end - col);
        col = end;
        index += 1;
      } else {
        const end = Math.min(span ? span.min : to, to);
        out.push(this.columnElement(undefined), end - col);
        col = end;
      }
    }
  }

  private columnElement(span: ColumnSpan | undefined): Repeatable {
    const width = span?.widthPx ?? this.sheet.defaultColWidthPx;
    const style =
      this.ctx.sourceStyleNames && span?.sourceStyleName !== undefined
        ? span.sourceStyleName
        : width === this.sheet.defaultColWidthPx
          ? this.defaultColumnStyle()
          : this.ctx.registry.columnStyle(width);
    return {
      head: `<table:table-column table:style-name="${attributeText(style)}"`,
      tail: `${span?.hidden ? ` table:visibility="collapse"` : ""} table:default-cell-style-name="${attributeText(this.styleName(span?.styleId ?? 0))}"/>`,
    };
  }

  // --- rows ---

  /** Rows [from, to), in order. */
  private rows(from: number, to: number, out: RunWriter): void {
    const stored = this.sheet.rows;
    let r = from;
    while (r < to) {
      while (this.nextMerge < this.merges.length && this.merges[this.nextMerge]!.top <= r) {
        this.active.push(this.merges[this.nextMerge]!);
        this.nextMerge += 1;
      }
      this.active = this.active.filter((merge) => merge.bottom >= r);
      this.meetValidated(r);
      this.meetLinks(r);
      while (this.nextStored < stored.length && stored[this.nextStored]!.index < r) {
        this.nextStored += 1;
      }
      const row = stored[this.nextStored]?.index === r ? stored[this.nextStored] : undefined;
      while (this.nextDrawingRow < this.drawingRows.length && this.drawingRows[this.nextDrawingRow]! < r) {
        this.nextDrawingRow += 1;
      }
      const drawn = this.drawingRows[this.nextDrawingRow] === r;
      while (this.nextNoteRow < this.noteRows.length && this.noteRows[this.nextNoteRow]! < r) {
        this.nextNoteRow += 1;
      }
      const noted = this.noteRows[this.nextNoteRow] === r;
      const inTable = this.tables.some((table) => table.ref.top <= r && table.ref.bottom >= r);
      if (row || drawn || noted || inTable || this.active.some((merge) => merge.top === r)) {
        out.push(this.rowElement(row, r), 1, drawn || noted);
        r += 1;
        continue;
      }
      let end = Math.min(to, stored[this.nextStored]?.index ?? to, this.drawingRows[this.nextDrawingRow] ?? to, this.noteRows[this.nextNoteRow] ?? to);
      if (this.nextMerge < this.merges.length) {
        end = Math.min(end, this.merges[this.nextMerge]!.top);
      }
      for (const merge of this.active) {
        end = Math.min(end, merge.bottom + 1);
      }
      end = Math.min(end, this.nextValidatedBreak(r));
      for (const table of this.tables) {
        if (table.ref.top > r) {
          end = Math.min(end, table.ref.top);
        }
      }
      out.push(this.rowElement(undefined, r), Math.max(end - r, 1));
      r = Math.max(end, r + 1);
    }
  }

  /** The async twin of `rows` (SPEC-137 "Saving stays responsive"):
   * identical row selection and run-compression logic — only every `r`'s
   * loop body is the same code, checking `clock` so a long sheet's save
   * yields to the main thread instead of blocking it for the whole table. */
  private async rowsAsync(from: number, to: number, out: RunWriter, clock: SliceClock): Promise<void> {
    const stored = this.sheet.rows;
    let r = from;
    while (r < to) {
      while (this.nextMerge < this.merges.length && this.merges[this.nextMerge]!.top <= r) {
        this.active.push(this.merges[this.nextMerge]!);
        this.nextMerge += 1;
      }
      this.active = this.active.filter((merge) => merge.bottom >= r);
      this.meetValidated(r);
      this.meetLinks(r);
      while (this.nextStored < stored.length && stored[this.nextStored]!.index < r) {
        this.nextStored += 1;
      }
      const row = stored[this.nextStored]?.index === r ? stored[this.nextStored] : undefined;
      while (this.nextDrawingRow < this.drawingRows.length && this.drawingRows[this.nextDrawingRow]! < r) {
        this.nextDrawingRow += 1;
      }
      const drawn = this.drawingRows[this.nextDrawingRow] === r;
      while (this.nextNoteRow < this.noteRows.length && this.noteRows[this.nextNoteRow]! < r) {
        this.nextNoteRow += 1;
      }
      const noted = this.noteRows[this.nextNoteRow] === r;
      const inTable = this.tables.some((table) => table.ref.top <= r && table.ref.bottom >= r);
      if (row || drawn || noted || inTable || this.active.some((merge) => merge.top === r)) {
        out.push(this.rowElement(row, r), 1, drawn || noted);
        r += 1;
      } else {
        let end = Math.min(to, stored[this.nextStored]?.index ?? to, this.drawingRows[this.nextDrawingRow] ?? to, this.noteRows[this.nextNoteRow] ?? to);
        if (this.nextMerge < this.merges.length) {
          end = Math.min(end, this.merges[this.nextMerge]!.top);
        }
        for (const merge of this.active) {
          end = Math.min(end, merge.bottom + 1);
        }
        end = Math.min(end, this.nextValidatedBreak(r));
        for (const table of this.tables) {
          if (table.ref.top > r) {
            end = Math.min(end, table.ref.top);
          }
        }
        out.push(this.rowElement(undefined, r), Math.max(end - r, 1));
        r = Math.max(end, r + 1);
      }
      if (clock.due()) {
        await clock.yield();
      }
    }
  }

  /** SPEC-135: the hyperlinks row `r` meets (rows ascend). */
  private meetLinks(r: number): void {
    if (this.links.length === 0) {
      return;
    }
    while (this.nextLink < this.links.length && this.links[this.nextLink]!.range.top <= r) {
      this.activeLinks.push(this.links[this.nextLink]!);
      this.nextLink += 1;
    }
    this.activeLinks = this.activeLinks.filter((link) => link.range.bottom >= r);
  }

  /** The current row's hyperlink at `col` — the first covering it in
   * model order (links are sorted by their start, and never overlap when
   * set through annotations.ts). */
  private linkAt(col: number): CellHyperlink | undefined {
    return this.activeLinks.find((link) => link.range.left <= col && link.range.right >= col);
  }

  /** The validated rectangles row `r` meets (rows ascend). */
  private meetValidated(r: number): void {
    if (this.validated.length === 0) {
      return;
    }
    while (this.nextValidated < this.validated.length && this.validated[this.nextValidated]!.range.top <= r) {
      this.activeValidated.push(this.validated[this.nextValidated]!);
      this.nextValidated += 1;
    }
    this.activeValidated = this.activeValidated.filter(({ range }) => range.bottom >= r);
  }

  /** The first row after `r` where the validated set changes. */
  private nextValidatedBreak(r: number): number {
    const breaks = this.validatedBreaks;
    while (this.nextBreak < breaks.length && breaks[this.nextBreak]! <= r) {
      this.nextBreak += 1;
    }
    return breaks[this.nextBreak] ?? Number.POSITIVE_INFINITY;
  }

  /** The current row's validated column segments: each column takes the
   * first covering rule in model order. */
  private validatedSegments(): ValidatedSegment[] {
    const active = this.activeValidated;
    if (active.length === 0) {
      return [];
    }
    if (active.length === 1) {
      const { range, name } = active[0]!;
      return [{ left: range.left, right: range.right, name }];
    }
    const points = [...new Set(active.flatMap(({ range }) => [range.left, range.right + 1]))].toSorted((a, b) => a - b);
    const segments: ValidatedSegment[] = [];
    for (let i = 0; i + 1 < points.length; i += 1) {
      const left = points[i]!;
      const right = points[i + 1]! - 1;
      let best: ValidatedRange | undefined;
      for (const each of active) {
        if (each.range.left <= left && each.range.right >= left && (!best || each.order < best.order)) {
          best = each;
        }
      }
      if (!best) {
        continue;
      }
      const last = segments.at(-1);
      if (last && last.name === best.name && last.right + 1 === left) {
        last.right = right;
      } else {
        segments.push({ left, right, name: best.name });
      }
    }
    return segments;
  }

  private rowElement(row: Row | undefined, r: number): Repeatable {
    let style: string;
    if (this.ctx.sourceStyleNames && row?.sourceStyleName !== undefined) {
      style = row.sourceStyleName;
    } else {
      const height = row?.heightPt ?? this.sheet.defaultRowHeightPt;
      style = height === this.sheet.defaultRowHeightPt && !row?.customHeight ? this.defaultRowStyle() : this.ctx.registry.rowStyle(height, !row?.customHeight);
    }
    // SPEC-128: the user's hidden wins — "collapse" whatever filterHidden is.
    // ODS rows read back as the wire says (never reclassified), so a row
    // both hidden and filter-hidden comes back hidden only.
    let tail = row?.hidden ? ` table:visibility="collapse"` : row?.filterHidden ? ` table:visibility="filter"` : "";
    if (row?.styleId !== undefined) {
      tail += ` table:default-cell-style-name="${attributeText(this.styleName(row.styleId))}"`;
    }
    return { head: `<table:table-row table:style-name="${attributeText(style)}"`, tail: `${tail}>${this.cells(row, r)}</table:table-row>` };
  }

  /** A row's cells, columns 0..colExtent−1. */
  private cells(row: Row | undefined, r: number): string {
    const out = new RunWriter("table:number-columns-repeated");
    const cells = row?.cells ?? [];
    // SPEC-130: the validation name each column is written with.
    const segments = this.validatedSegments();
    let s = 0;
    const segmentAt = (col: number): ValidatedSegment | undefined => {
      while (s < segments.length && segments[s]!.right < col) {
        s += 1;
      }
      return segments[s];
    };
    const merges: MergeInRow[] = this.active
      .filter((merge) => merge.top <= r && merge.bottom >= r)
      .map((range) => ({ range, anchor: range.top === r }))
      .toSorted((a, b) => a.range.left - b.range.left);
    let next = 0; // the next stored cell
    let m = 0; // the next merge
    let col = 0;
    // SPEC-131: the drawings anchored on this row, by column.
    const drawn = this.drawings.cells.get(r);
    const drawnCols = drawn ? [...drawn.keys()].toSorted((a, b) => a - b) : [];
    let d = 0;
    // SPEC-135: the notes on this row, by column.
    const noted = this.notes.get(r);
    const notedCols = noted ? [...noted.keys()].toSorted((a, b) => a - b) : [];
    let n = 0;
    // SPEC-133: the tables this row crosses — each of their cells is written.
    const rowTables = this.tables.filter((table) => table.ref.top <= r && table.ref.bottom >= r);
    while (col < this.colExtent) {
      while (m < merges.length && merges[m]!.range.right < col) {
        m += 1;
      }
      const merge = merges[m] && merges[m]!.range.left <= col ? merges[m] : undefined;
      while (next < cells.length && cells[next]!.col < col) {
        next += 1;
      }
      const cell = cells[next]?.col === col ? cells[next] : undefined;
      const anchor = merge?.anchor && merge.range.left === col ? merge.range : undefined;
      const covered = merge !== undefined && !anchor;
      const segment = segmentAt(col);
      const validation = segment && segment.left <= col ? segment.name : undefined;
      while (d < drawnCols.length && drawnCols[d]! < col) {
        d += 1;
      }
      const drawings = drawnCols[d] === col ? drawn!.get(col) : undefined;
      while (n < notedCols.length && notedCols[n]! < col) {
        n += 1;
      }
      const note = notedCols[n] === col ? noted!.get(col) : undefined;
      const table = rowTables.find((each) => each.ref.left <= col && each.ref.right >= col);
      if (cell || anchor || drawings !== undefined || note || table) {
        out.push(this.cellElement(cell, covered, anchor, row, r, col, validation, drawings, table, note), 1, drawings !== undefined || note !== undefined);
        col += 1;
        continue;
      }
      let end = Math.min(this.colExtent, cells[next]?.col ?? this.colExtent, drawnCols[d] ?? this.colExtent, notedCols[n] ?? this.colExtent);
      if (merge) {
        end = Math.min(end, merge.range.right + 1);
      } else if (merges[m]) {
        end = Math.min(end, merges[m]!.range.left);
      }
      if (segment) {
        end = Math.min(end, validation === undefined ? segment.left : segment.right + 1);
      }
      for (const each of rowTables) {
        if (each.ref.left > col) {
          end = Math.min(end, each.ref.left);
        }
      }
      const head = (covered ? "<table:covered-table-cell" : "<table:table-cell") + validationAttribute(validation);
      out.push({ head, tail: "/>" }, Math.max(end - col, 1));
      col = Math.max(end, col + 1);
    }
    out.flush();
    return out.out;
  }

  private valueTypeOf(code: string | undefined): { type: NumberValueType; symbol?: string } {
    const key = code ?? "";
    let found = this.valueTypes.get(key);
    if (!found) {
      found = numberValueType(code);
      this.valueTypes.set(key, found);
    }
    return found;
  }

  /** One cell element (see serializeTable, "Cells"). The style name is part
   * of the head, so a repeat count follows it, as LibreOffice writes. */
  private cellElement(
    cell: Cell | undefined,
    covered: boolean,
    anchor: GridRange | undefined,
    row: Row | undefined,
    r: number,
    col: number,
    validation?: string,
    drawings?: string,
    table?: SheetTable,
    note?: CellNote,
  ): Repeatable {
    let head = covered ? "<table:covered-table-cell" : "<table:table-cell";
    const close = covered ? "</table:covered-table-cell>" : "</table:table-cell>";
    let attrs = "";
    const positionDefault = row?.styleId ?? columnSpanAt(this.sheet.columns, col)?.styleId ?? 0;
    const tableStyle = this.tableStyleName(table, r, col, cell?.styleId ?? positionDefault);
    if (tableStyle !== undefined) {
      head += ` table:style-name="${attributeText(tableStyle)}"`;
    } else if (cell?.styleId !== undefined && cell.styleId !== positionDefault) {
      head += ` table:style-name="${attributeText(this.styleName(cell.styleId))}"`;
    }
    head += validationAttribute(validation);
    if (anchor) {
      attrs += ` table:number-columns-spanned="${anchor.right - anchor.left + 1}" table:number-rows-spanned="${anchor.bottom - anchor.top + 1}"`;
    }
    // SPEC-135: a note is the cell's first child (ods/annotations.ts).
    const noteXml = note ? odsNoteXml(note, (text) => paragraphsXml([{ text }])) : "";
    if (!cell) {
      const content = noteXml + (drawings ?? "");
      return { head, tail: content === "" ? `${attrs}/>` : `${attrs}>${content}${close}` };
    }
    if (cell.arrayRange && cell.formula !== undefined) {
      const range = cell.arrayRange;
      attrs += ` table:number-matrix-columns-spanned="${range.right - range.left + 1}" table:number-matrix-rows-spanned="${range.bottom - range.top + 1}"`;
    }
    if (cell.formula !== undefined) {
      attrs += ` table:formula="${attributeText(this.formulaText(cell, r, col))}"`;
    }
    const style = cellStyle(this.ctx.model, cell.styleId ?? positionDefault);
    const value = cell.value;
    let text: TextPiece[] = [];
    if (value !== null) {
      const code = formatCodeOf(style, this.ctx.locale);
      let type: string;
      if (typeof value === "number") {
        const kind = this.valueTypeOf(code);
        type = kind.type;
        attrs += ` office:value-type="${type}"`;
        if (type === "date") {
          attrs += ` office:date-value="${serialToOdsDateValue(value, this.ctx.model.dateSystem, this.ctx.model.odsNullDate)}"`;
        } else if (type === "time") {
          attrs += ` office:time-value="${serialToOdsTimeValue(value)}"`;
        } else {
          if (type === "currency") {
            attrs += ` office:currency="${attributeText(currencyCodeOf(kind.symbol, this.ctx.locale))}"`;
          }
          attrs += ` office:value="${Object.is(value, -0) ? "0" : String(value)}"`;
        }
        text = [{ text: plainText(formatValue(value, code, this.format)) }];
      } else if (typeof value === "boolean") {
        type = "boolean";
        attrs += ` office:value-type="boolean" office:boolean-value="${value}"`;
        text = [{ text: value ? "TRUE" : "FALSE" }];
      } else if (typeof value === "string") {
        type = "string";
        attrs += ` office:value-type="string"`;
        text = this.stringPieces(cell, value);
      } else {
        type = "error";
        attrs += cell.formula !== undefined ? ` office:value-type="float" office:value="0"` : ` office:value-type="string"`;
        text = [{ text: value.error }];
      }
      attrs += ` calcext:value-type="${isCellError(value) ? "error" : type}"`;
    }
    const pass = this.ctx.passRetained ?? ((retained: string): string => retained);
    if (cell.preservedAttrs !== undefined) {
      // A name captured before SPEC-130 modelled it: the model's wins.
      attrs += withoutValidationName(pass(cell.preservedAttrs));
    }
    let content = noteXml + (cell.preservedXml === undefined ? "" : pass(cell.preservedXml));
    // SPEC-131: after the note and any preserved element, before the text.
    content += drawings ?? "";
    // SPEC-135: the text wrapped in the cell's link — verbatim paragraphs
    // only while their first text:a says what the model's link says.
    const link = this.linkAt(col);
    const preservedText = cell.preservedText === undefined ? undefined : pass(cell.preservedText);
    if (preservedText !== undefined && paragraphsMatchLink(preservedText, link)) {
      content += preservedText;
    } else {
      const paragraphs = paragraphsXml(text);
      content += link ? linkedParagraphsXml(paragraphs, link) : paragraphs;
    }
    return { head, tail: content === "" ? `${attrs}/>` : `${attrs}>${content}${close}` };
  }

  /** A string's pieces: its rich runs (each in its `de-T` text style) when
   * they spell it, else the string as typed. */
  private stringPieces(cell: Cell, value: string): TextPiece[] {
    const runs = cell.runs;
    if (!runs || runs.map((run) => run.text).join("") !== value) {
      return [{ text: value }];
    }
    return runs.map((run) => {
      const style = this.ctx.registry.textStyle(run.font, this.env.theme);
      return style === undefined ? { text: run.text } : { text: run.text, style };
    });
  }

  /** `table:formula`: the stored wire value verbatim for an unparsed formula,
   * else the canonical formula through toOdsFormula — SPEC-133: its
   * structured references first replaced by their A1 ranges
   * (expandStructuredReferences, hosted at the cell). */
  private formulaText(cell: Cell, r: number, col: number): string {
    let formula = cell.formula!;
    if (cell.formulaUnparsed) {
      return formula;
    }
    const { sheets, names } = this.ctx.model;
    // A bare table name (`SUM(Sales)`) is a structured reference too.
    if (sheets !== undefined && (formula.includes("[") || this.workbookHasTables)) {
      formula = expandStructuredReferences({ sheets, names: names ?? [] }, formula, { sheetId: this.sheet.id, row: r, col });
    }
    try {
      return toOdsFormula(formula);
    } catch {
      return `of:=${formula}`;
    }
  }

  /** Writes [from, to) of an axis, re-emitting the source's grouping
   * elements around the indices they covered. */
  private axis(axis: Axis, groups: readonly SkeletonGroup[], from: number, to: number, emit: (a: number, b: number, out: RunWriter) => void): string {
    let xml = "";
    let at = from;
    const run = (a: number, b: number): void => {
      if (b > a) {
        const out = new RunWriter(axis === "row" ? "table:number-rows-repeated" : "table:number-columns-repeated");
        emit(a, b, out);
        out.flush();
        xml += out.out;
      }
    };
    for (const group of groups) {
      const start = Math.max(group.from, at);
      run(at, Math.min(start, to));
      if (group.selfClosing) {
        xml += group.open;
      } else {
        const end = Math.max(Math.min(group.to, to), start);
        xml += group.open + this.axis(axis, group.children, start, end, emit) + `</${group.qname}>`;
        at = end;
        continue;
      }
      at = Math.max(at, start);
    }
    run(at, to);
    return xml;
  }

  /** The async twin of `axis` (SPEC-137): identical grouping logic, awaiting
   * the (also async) row emitter instead of calling it synchronously. Only
   * used for the "row" axis — columns are few enough on any sheet that they
   * are never worth yielding over. */
  private async axisAsync(groups: readonly SkeletonGroup[], from: number, to: number, clock: SliceClock): Promise<string> {
    let xml = "";
    let at = from;
    const run = async (a: number, b: number): Promise<void> => {
      if (b > a) {
        const out = new RunWriter("table:number-rows-repeated");
        await this.rowsAsync(a, b, out, clock);
        out.flush();
        xml += out.out;
      }
    };
    for (const group of groups) {
      const start = Math.max(group.from, at);
      await run(at, Math.min(start, to));
      if (group.selfClosing) {
        xml += group.open;
      } else {
        const end = Math.max(Math.min(group.to, to), start);
        xml += group.open + (await this.axisAsync(group.children, start, end, clock)) + `</${group.qname}>`;
        at = end;
        continue;
      }
      at = Math.max(at, start);
    }
    await run(at, to);
    return xml;
  }

  write(): string {
    const sheet = this.sheet;
    const ctx = this.ctx;
    let head: string;
    let leading = this.skeleton?.leading ?? "";
    let trailing = this.skeleton?.trailing ?? "";
    if (ctx.conditionalFormats) {
      ({ leading, trailing } = ctx.conditionalFormats.splice(sheet, leading, trailing, ctx.passRetained));
    }
    // SPEC-133: the sheet's own names, after the rows (ods/namedExpressions.ts).
    if (ctx.localNames) {
      trailing = ctx.localNames(sheet, trailing);
    }
    // SPEC-131: the fixed drawings, from the model, where the source's table:shapes was.
    leading = leadingWithShapes(leading, this.drawings.shapes, this.skeleton ? tablePrefix(this.skeleton.close) : "table:");
    if (this.skeleton) {
      head = this.skeleton.head + leading;
    } else {
      const tabColor = sheet.tabColor === undefined ? undefined : resolveColor(sheet.tabColor, ctx.model.theme ?? DEFAULT_THEME);
      const style = ctx.registry.tableStyle({ masterPage: "Default", display: sheet.visibility === "visible", ...(tabColor === undefined ? {} : { tabColor }) });
      head = `<table:table table:name="${attributeText(sheet.name)}" table:style-name="${style}">` + leading;
    }
    const columns = this.axis("col", this.skeleton?.colGroups ?? [], 0, this.colExtent, (a, b, out) => this.columns(a, b, out));
    const rows = this.axis("row", this.skeleton?.rowGroups ?? [], 0, this.rowExtent, (a, b, out) => this.rows(a, b, out));
    return head + columns + rows + trailing + (this.skeleton ? this.skeleton.close : "</table:table>");
  }

  /** The async twin of `write` (SPEC-137 "Saving stays responsive"):
   * identical output, built the same way — only the row axis (the one
   * that can run to tens of thousands of iterations) is built by the
   * yielding `axisAsync`/`rowsAsync`; columns are always few, so they stay
   * on the plain synchronous `axis`. */
  async writeAsync(clock: SliceClock): Promise<string> {
    const sheet = this.sheet;
    const ctx = this.ctx;
    let head: string;
    let leading = this.skeleton?.leading ?? "";
    let trailing = this.skeleton?.trailing ?? "";
    if (ctx.conditionalFormats) {
      ({ leading, trailing } = ctx.conditionalFormats.splice(sheet, leading, trailing, ctx.passRetained));
    }
    if (ctx.localNames) {
      trailing = ctx.localNames(sheet, trailing);
    }
    leading = leadingWithShapes(leading, this.drawings.shapes, this.skeleton ? tablePrefix(this.skeleton.close) : "table:");
    if (this.skeleton) {
      head = this.skeleton.head + leading;
    } else {
      const tabColor = sheet.tabColor === undefined ? undefined : resolveColor(sheet.tabColor, ctx.model.theme ?? DEFAULT_THEME);
      const style = ctx.registry.tableStyle({ masterPage: "Default", display: sheet.visibility === "visible", ...(tabColor === undefined ? {} : { tabColor }) });
      head = `<table:table table:name="${attributeText(sheet.name)}" table:style-name="${style}">` + leading;
    }
    const columns = this.axis("col", this.skeleton?.colGroups ?? [], 0, this.colExtent, (a, b, out) => this.columns(a, b, out));
    const rows = await this.axisAsync(this.skeleton?.rowGroups ?? [], 0, this.rowExtent, clock);
    return head + columns + rows + trailing + (this.skeleton ? this.skeleton.close : "</table:table>");
  }
}

/** The prefix a table's close tag names it with ("table:"). */
function tablePrefix(close: string): string {
  const name = /^<\/([^\s>]+)/.exec(close)?.[1] ?? "table:table";
  return name.includes(":") ? name.slice(0, name.indexOf(":") + 1) : "";
}

/** ` table:content-validation-name="…"` for a validated position ("" for none). */
function validationAttribute(name: string | undefined): string {
  return name === undefined ? "" : ` table:content-validation-name="${attributeText(name)}"`;
}

/** A sheet's named validation rectangles, clamped to the grid, by top. */
function validatedRanges(sheet: Sheet, nameOf: OdsTableContext["validationName"]): ValidatedRange[] {
  if (!nameOf) {
    return [];
  }
  const out: ValidatedRange[] = [];
  sheetValidations(sheet).forEach((rule, order) => {
    const name = nameOf(sheet, rule);
    if (name === undefined) {
      return;
    }
    for (const range of rule.ranges) {
      if (range.top < MAX_ROWS && range.left < MAX_COLS && range.bottom >= range.top && range.right >= range.left) {
        out.push({ range: { ...range, bottom: Math.min(range.bottom, MAX_ROWS - 1), right: Math.min(range.right, MAX_COLS - 1) }, order, name });
      }
    }
  });
  return out.toSorted((a, b) => a.range.top - b.range.top || a.order - b.order);
}

/** One sheet as a `<table:table>` (SPEC-122 "The table writer"), in order:
 * 1. the start tag — the source's (with SPEC-121's splices, applied by the
 *    caller), or generated for a new sheet (its `de-ta` style);
 * 2. the leading children verbatim (`table:title`, `table:desc`,
 *    `table:table-source`, `office:dde-source`, `table:scenario`,
 *    `office:forms`, `table:shapes`);
 * 3. columns: one element per run of identical columns — the style its
 *    `sourceStyleName` (the source's styles retained), else a `de-co` by
 *    width (a default-width column the source's last-run style when that
 *    still measures the default), `table:visibility="collapse"` when hidden,
 *    `table:default-cell-style-name` from `styleId` or "Default" —
 *    compressed with `table:number-columns-repeated`;
 * 4. rows 0..rowExtent−1: each stored row serialised, gaps as empty rows,
 *    consecutive rows with identical serialisation (empty ones included)
 *    compressed with `table:number-rows-repeated` (so visibility is part of
 *    the compression key: a filter-hidden row never joins a collapsed or a
 *    visible one); the style `sourceStyleName` or a `de-ro` by height
 *    (default rows as columns), `table:visibility="collapse"` when hidden
 *    (whatever filterHidden is), `"filter"` when only filterHidden
 *    (SPEC-128), `table:default-cell-style-name` from `Row.styleId`;
 * 5. within a row, columns 0..colExtent−1: stored cells serialised, gaps as
 *    `<table:table-cell/>`, identical neighbours compressed; a merge anchor
 *    gets `table:number-columns-spanned`/`-rows-spanned` and every position
 *    it covers is a `table:covered-table-cell` (carrying content when a
 *    cell is stored there);
 * 6. the source's grouping skeleton — `table:table-header-rows`,
 *    `table:table-row-group`, `table:table-rows` and the column equivalents,
 *    start tags verbatim — re-emitted around the indices each covered at
 *    read; `text:soft-page-break` (a layout cache) is dropped;
 * 7. the trailing children verbatim (`table:named-expressions`,
 *    `calcext:conditional-formats`, anything after the rows) — but for
 *    SPEC-129: with a conditional-format writer in the context, the
 *    `calcext:conditional-formats` among the leading or trailing children
 *    is the model's (OdsConditionalFormatWriter.splice: in place, removed,
 *    or a new one last), and a cell style whose source automatic style
 *    carries a `style:map` (`mappedCellStyle`) is written as a generated
 *    `de-ce` style instead of its source name.
 * Extent per axis = max(the source's total repeated extent, the stored rows,
 * cells, merges and column spans + 1), so LibreOffice's 1048553-row and
 * 16380-column tails come back at the same length; a new sheet's is what it
 * stores (at least 1 × 1). The reader takes the LAST column and row runs'
 * sizes as the sheet's defaults, so a last column or row of another size
 * gets one default-sized one after it.
 * Cells: `table:style-name` when the cell's style differs from the
 * position's default (the row's, else the column's); `office:value-type` by
 * value and number format — date (`office:date-value`), time
 * (`office:time-value`), percentage, currency (+ `office:currency`), float,
 * boolean, string; an error is float `office:value="0"` in a formula cell
 * and a string literal otherwise — with `calcext:value-type` (the same word,
 * or "error"); `table:formula` through toOdsFormula (an unparsed formula's
 * stored wire value verbatim); an array anchor's
 * `table:number-matrix-columns-spanned`/`-rows-spanned`; then
 * `preservedAttrs`. Content: `preservedXml`, then `preservedText` verbatim
 * when present, else one `<text:p>` per line of the SPEC-118 formatted text
 * in the user's locale (strings as typed; rich runs as `text:span`s in
 * `de-T` styles; runs of spaces as `text:s`, tabs as `text:tab`). A formula
 * cell with no value has no value type; a stale cell's last known value is
 * written. Preserved content passes through `passRetained`.
 * SPEC-130: with a `validationName` in the context, every position a rule
 * of `Sheet.validations` covers (the first covering rule in model order)
 * carries `table:content-validation-name` after its style name — part of
 * the repeat-compression key, so validated and unvalidated neighbours split
 * runs, and an empty validated run is one attribute-only element
 * (`<table:table-cell table:content-validation-name="x"
 * table:number-columns-repeated="n"/>`); runs of empty rows split where the
 * validated rectangles start and end, and are never materialised. The
 * extent covers every validated rectangle. A stale
 * `table:content-validation-name` in `preservedAttrs` is never written.
 * SPEC-131: the drawings are the model's (ods/drawingXml.ts, the normative
 * source): the `fixed` ones in a `table:shapes` that replaces the source's
 * among the leading children (after `office:forms`; none → no element),
 * the `moveAndSize` and `move` ones inside their `from` cell after its
 * preservedXml (any `office:annotation`) and before its text — the cell
 * materialised when nothing else is stored there, its row too, and neither
 * ever repeat-compressed with a neighbour. The extent covers every anchor
 * cell.
 * SPEC-135: each note of `Sheet.notes` is its cell's FIRST child
 * (ods/annotations.ts odsNoteXml — the source element verbatim while
 * unchanged, without its stale caption position once moved, else
 * generated), its cell and row materialised and never repeat-compressed as
 * a drawing's are, the extent covering it; a cell a `Sheet.hyperlinks`
 * range covers has its paragraphs' content wrapped in the link's `text:a`
 * (linkedParagraphsXml), its `preservedText` kept only while that
 * text's first `text:a` still matches the model's link (paragraphsMatchLink). */
export function serializeTable(sheet: Sheet, ctx: OdsTableContext): string {
  return new TableWriter(sheet, ctx).write();
}

/** The async, yielding twin of `serializeTable` (SPEC-137 "Saving stays
 * responsive"): byte-identical output, but yields to the main thread
 * (`clock`, default `createSliceClock()`) every `SERIALISE_SLICE_MS` of
 * work. `serializeTable` itself stays synchronous and unchanged — the
 * substitution save path (`ods/writer.ts`'s `buildSubstitution`, which the
 * reader-upgrade path's `substitutedOdsParts` also calls and must stay
 * synchronous for) keeps using it; `writeOds`'s from-scratch path uses this
 * twin, which large-sheet saves actually need. */
export async function serializeTableAsync(sheet: Sheet, ctx: OdsTableContext, clock: SliceClock = createSliceClock()): Promise<string> {
  return new TableWriter(sheet, ctx).writeAsync(clock);
}
