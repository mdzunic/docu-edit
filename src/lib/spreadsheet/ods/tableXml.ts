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
//   "false"` → customHeight, visibility → hidden, a default cell style →
//   styleId. `defaultRowHeightPt` is the height of the style on the table's
//   LAST row run, else 12.8 pt, and is only known when the table ends — so
//   a row holding no cell is kept as a pending run (never expanded) and
//   materialised then only when it differs from the default (a height of
//   its own, hidden, or a style). LibreOffice's trailing 1048553-row run IS
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
import { scanAttr, type XmlScanAttribute } from "../../core/xmlScanner";
import { TABLE_NAMESPACE } from "../../core/odfXml";
import type { GridRange } from "../address";
import { ODS_DEFAULT_COL_WIDTH_PX, ODS_DEFAULT_ROW_HEIGHT_PT } from "../geometry";
import { acceptMerges } from "../merges";
import type { FormulaFields } from "../formula/dialect";
import { MAX_COLS, MAX_ROWS, type Cell, type CellValue, type ColumnSpan, type Row, type TextRun } from "../model";
import { addressText, type CellBudget, outOfGridError } from "../parsed";
import type { OdsRowHeight } from "./styles";

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
  hidden: boolean;
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
  hidden: boolean;
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
      hidden: isHidden(attrs),
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
   * format its value type implies when its style gives none (SPEC-118) and
   * `formula` its formula (SPEC-119). Returns false when the cell ceiling
   * stops the read. */
  cell(
    attrs: readonly XmlScanAttribute[],
    covered: boolean,
    value: CellValue,
    runs?: TextRun[],
    impliedCode?: string,
    formula?: OdsCellFormula,
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
      return this.storeValues(row, first, repeat, value, runs, styleId, impliedCode, formula);
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
      row.cells.push(cell);
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
    const entry: RowEntry = { index: start, repeat, hidden: row.hidden };
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
    if (entry.hidden) {
      row.hidden = true;
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
        if (!ownHeight && !entry.hidden && entry.styleId === undefined) {
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
