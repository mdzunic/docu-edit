// Sheet pagination (SPEC-136 "The pagination algorithm", normative) — how a
// grid is cut into pages: framework-free, imports nothing PDF or DOM, and
// never mutates the model it reads (print.spec.ts deep-freezes it).
//
// All lengths are pt: a column is its GridMetrics px × 0.75, a row its own
// `heightPt` (else the sheet default); `s` = scale ÷ 100.
// 1. Area: the print area's ranges if set (each starts on a new page, in
//    order; a whole-row or whole-column reference stops at the printable
//    used range), else the selection's ranges for scope "selection", else
//    the printable used range — the bounding box of cells with a value, a
//    fill or a border, merged ranges touching it and drawing/chart anchors.
//    Empty → no pages.
// 2. Visible tracks: hidden and filter-hidden rows and columns are dropped.
// 3. Printable box: the oriented paper minus the margins; for the bands, also
//    minus the print titles (× s) and, with printHeadings, a heading column
//    of max(24, 6 × digits + 6) pt and a 15 pt heading row (× s).
// 4. Scale: percent mode `s` as set; fit mode scans p from 100 down to 10 and
//    takes the first at which steps 5–6 give ≤ fitTo.width column bands and
//    ≤ fitTo.height row bands (0 = no limit), 10 if none; odsScaleToPages the
//    first at which bands-wide × bands-tall ≤ N. Manual breaks are ignored in
//    both fit modes. Scanning, not bisection: merges make counts non-monotone.
// 5. Column bands: accumulate scaled widths; a new band starts before a
//    manually broken column (percent mode) or one that would overflow the
//    box. A boundary strictly inside a merge moves back to the merge's first
//    column when the band keeps a column; a merge wider than the box splits
//    at the box edge; a single column wider than the box is its own band.
// 6. Row bands: the same with heights and row breaks.
// 7. Order: downThenOver — every row band of column band 1, then 2…;
//    overThenDown — every column band of row band 1, then 2….
// 8. Titles: a page draws the title rows above its band and the title
//    columns left of it, unless its band already reaches them.
// 9. Placement: the block (headings + titles + band) at the printable box's
//    top-left, centred on each axis set to centre.
// 10. Numbering: from firstPageNumber (else 1, or the job's running number),
//    continuing across sheets in a workbook job; a sheet with its own
//    firstPageNumber restarts; `&N` is the job's total.
import type { PageSize } from "../core/pageSetup";
import type { GridRange } from "./address";
import { drawingRectPx } from "./drawings";
import { buildGridMetrics, type GridMetrics } from "./geometry";
import type { HeaderFooterPair, HeaderFooterSections, Sheet, SheetPageSetup, SpreadsheetModel } from "./model";
import { MAX_COLS, MAX_ROWS, rowIsHidden } from "./model";
import { effectiveMargins, orientedPaper, printAreaOf, printTitlesOf, type PrintPreferences } from "./pageSetup";

export const PT_PER_PX = 0.75;
export const PT_PER_CM = 72 / 2.54;
/** The most pages one browser print job may hold. */
export const PRINT_PAGE_LIMIT = 250;
/** The most pages one PDF export may hold. */
export const MAX_SHEET_PDF_PAGES = 2000;
/** Pages an export keeps mounted off-screen at once. */
export const EXPORT_WINDOW_PAGES = 4;
export const HEADING_ROW_PT = 15;
export const PRINT_LIMIT_MESSAGE = "This print job is over 250 pages — download a PDF instead";
export const PDF_LIMIT_MESSAGE = "This export is over 2,000 pages — set a print area first";
export const NOTHING_TO_PRINT_MESSAGE = "This sheet has nothing to print.";

export type PrintScope = "sheet" | "selection" | "workbook";

export interface SheetPrintOptions {
  /** "selection" prints `selection` where no print area is set. */
  scope?: PrintScope;
  selection?: readonly GridRange[];
  /** What a sheet whose setup names no paper prints on. */
  preferences?: PrintPreferences;
  /** The first page's number when the sheet's setup sets none. */
  startNumber?: number;
}

export interface SheetPrintPage {
  sheetId: string;
  number: number;
  /** The oriented paper. */
  paperPt: { width: number; height: number };
  marginsPt: { top: number; right: number; bottom: number; left: number; header: number; footer: number };
  /** Percent (10–400) the page's grid is laid out at. */
  scale: number;
  rows: number[];
  cols: number[];
  titleRows: number[];
  titleCols: number[];
  /** Where the block (headings + titles + band) starts, from the paper's top-left. */
  originPt: { x: number; y: number };
  header?: HeaderFooterSections;
  footer?: HeaderFooterSections;
  /** Which of the sheet's print ranges the page belongs to (0-based). */
  area: number;
}

export interface SheetPrintJob {
  pages: SheetPrintPage[];
}

// --- geometry ----------------------------------------------------------------------------------

/** A column's width in pt at 100%. */
export function colWidthPt(metrics: GridMetrics, col: number): number {
  return metrics.colWidth(col) * PT_PER_PX;
}

/** A row's height in pt at 100% (0 when hidden). */
export function rowHeightPt(sheet: Sheet, row: number, cache?: Map<number, Sheet["rows"][number]>): number {
  const stored = cache ? cache.get(row) : sheet.rows.find((each) => each.index === row);
  if (rowIsHidden(stored)) {
    return 0;
  }
  return stored?.heightPt ?? sheet.defaultRowHeightPt;
}

/** The width of the row-number heading column for rows up to `lastRow`. */
export function headingColumnPt(lastRow: number): number {
  return Math.max(24, 6 * String(lastRow + 1).length + 6);
}

function rowIndex(sheet: Sheet): Map<number, Sheet["rows"][number]> {
  return new Map(sheet.rows.map((row) => [row.index, row]));
}

interface Track {
  indices: number[];
  sizes: number[];
}

function visibleColumns(metrics: GridMetrics, left: number, right: number): Track {
  const track: Track = { indices: [], sizes: [] };
  for (let col = left; col <= right; col += 1) {
    const width = colWidthPt(metrics, col);
    if (width > 0) {
      track.indices.push(col);
      track.sizes.push(width);
    }
  }
  return track;
}

function visibleRows(sheet: Sheet, rows: Map<number, Sheet["rows"][number]>, top: number, bottom: number): Track {
  const track: Track = { indices: [], sizes: [] };
  for (let row = top; row <= bottom; row += 1) {
    const height = rowHeightPt(sheet, row, rows);
    if (height > 0) {
      track.indices.push(row);
      track.sizes.push(height);
    }
  }
  return track;
}

// --- step 1: the area ---------------------------------------------------------------------------

function inks(model: Pick<SpreadsheetModel, "styles">, styleId: number | undefined): boolean {
  if (styleId === undefined) {
    return false;
  }
  const style = model.styles[styleId];
  if (style === undefined) {
    return false;
  }
  const borders = style.borders;
  return style.fill !== undefined || (borders !== undefined && (borders.top || borders.right || borders.bottom || borders.left || borders.diagonal) !== undefined);
}

/** The printable used range (step 1): cells with a value, a fill or a
 * border, merges touching them and drawing anchors; null when none. */
export function printableUsedRange(model: Pick<SpreadsheetModel, "styles">, sheet: Sheet, metrics: GridMetrics = buildGridMetrics(sheet)): GridRange | null {
  let range: GridRange | null = null;
  const include = (top: number, left: number, bottom: number, right: number): void => {
    if (range === null) {
      range = { top, left, bottom, right };
      return;
    }
    range.top = Math.min(range.top, top);
    range.left = Math.min(range.left, left);
    range.bottom = Math.max(range.bottom, bottom);
    range.right = Math.max(range.right, right);
  };
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      if ((cell.value !== null && cell.value !== "") || cell.formula !== undefined || inks(model, cell.styleId)) {
        include(row.index, cell.col, row.index, cell.col);
      }
    }
  }
  for (const drawing of sheet.drawings ?? []) {
    if (drawing.hidden) {
      continue;
    }
    const rect = drawingRectPx(drawing.anchor, metrics);
    include(
      metrics.rowAt(rect.top),
      metrics.colAt(rect.left),
      metrics.rowAt(rect.top + Math.max(0, rect.height - 1)),
      metrics.colAt(rect.left + Math.max(0, rect.width - 1)),
    );
  }
  if (range === null) {
    return null;
  }
  const found = range as GridRange;
  let grown = true;
  while (grown) {
    grown = false;
    for (const merge of sheet.merges) {
      const touches = merge.top <= found.bottom && merge.bottom >= found.top && merge.left <= found.right && merge.right >= found.left;
      if (touches && (merge.top < found.top || merge.left < found.left || merge.bottom > found.bottom || merge.right > found.right)) {
        include(merge.top, merge.left, merge.bottom, merge.right);
        grown = true;
      }
    }
  }
  return found;
}

/** The ranges a sheet prints (step 1), clipped to the used range where a
 * reference reaches the grid's edge. */
export function printRanges(model: Pick<SpreadsheetModel, "styles" | "names">, sheet: Sheet, options: SheetPrintOptions = {}, metrics: GridMetrics = buildGridMetrics(sheet)): GridRange[] {
  const area = printAreaOf(model, sheet);
  const selection = options.scope === "selection" && options.selection !== undefined && options.selection.length > 0 ? options.selection : null;
  const chosen = area ?? selection;
  if (chosen === null) {
    const used = printableUsedRange(model, sheet, metrics);
    return used === null ? [] : [used];
  }
  let used: GridRange | null | undefined;
  const usedRange = (): GridRange | null => (used === undefined ? (used = printableUsedRange(model, sheet, metrics)) : used);
  const out: GridRange[] = [];
  for (const range of chosen) {
    let { top, left, bottom, right } = range;
    if (bottom >= MAX_ROWS - 1 || right >= MAX_COLS - 1) {
      const bounds = usedRange();
      if (bounds === null) {
        continue;
      }
      if (bottom >= MAX_ROWS - 1) {
        bottom = Math.min(bottom, bounds.bottom);
      }
      if (right >= MAX_COLS - 1) {
        right = Math.min(right, bounds.right);
      }
    }
    if (top <= bottom && left <= right) {
      out.push({ top, left, bottom, right });
    }
  }
  return out;
}

// --- steps 5 and 6: bands ------------------------------------------------------------------------

/** A band: positions [start, end] into a track. */
export interface Band {
  start: number;
  end: number;
}

/** A merge's extent along one axis. */
interface Span {
  first: number;
  last: number;
}

const EPSILON = 1e-6;

function firstPositionAtOrAfter(indices: readonly number[], index: number): number {
  let lo = 0;
  let hi = indices.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (indices[mid]! < index) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return lo;
}

/** Steps 5/6 over one track: `limit` is the box length in pt, `s` the scale,
 * `breaks` the manual breaks (null in fit mode), `merges` the spans a
 * boundary must not fall strictly inside. */
export function cutBands(track: Track, s: number, limit: number, breaks: ReadonlySet<number> | null, merges: readonly Span[]): Band[] {
  const { indices, sizes } = track;
  const bands: Band[] = [];
  if (indices.length === 0) {
    return bands;
  }
  let start = 0;
  let used = 0;
  let k = 0;
  while (k < indices.length) {
    const size = sizes[k]! * s;
    if (k > start && ((breaks !== null && breaks.has(indices[k]!)) || used + size > limit + EPSILON)) {
      let cut = k;
      let moved = true;
      while (moved) {
        moved = false;
        for (const merge of merges) {
          if (merge.first <= indices[cut - 1]! && merge.last >= indices[cut]!) {
            const back = firstPositionAtOrAfter(indices, merge.first);
            if (back > start && back < cut) {
              cut = back;
              moved = true;
            }
          }
        }
      }
      bands.push({ start, end: cut - 1 });
      start = cut;
      used = 0;
      for (let j = cut; j < k; j += 1) {
        used += sizes[j]! * s;
      }
      continue;
    }
    used += size;
    k += 1;
  }
  bands.push({ start, end: indices.length - 1 });
  return bands;
}

// --- the sheet -------------------------------------------------------------------------------

interface AreaPlan {
  cols: Track;
  rows: Track;
  colMerges: Span[];
  rowMerges: Span[];
}

interface SheetGeometry {
  boxW: number;
  boxH: number;
  titleRows: Track;
  titleCols: Track;
  headingW: number;
  headingH: number;
}

function sum(values: readonly number[]): number {
  return values.reduce((total, value) => total + value, 0);
}

function bandsAt(plan: AreaPlan, geometry: SheetGeometry, s: number, fit: boolean, setup: SheetPageSetup | undefined): { cols: Band[]; rows: Band[] } {
  const limitW = geometry.boxW - s * (sum(geometry.titleCols.sizes) + geometry.headingW);
  const limitH = geometry.boxH - s * (sum(geometry.titleRows.sizes) + geometry.headingH);
  const colBreaks = fit ? null : new Set(setup?.colBreaks ?? []);
  const rowBreaks = fit ? null : new Set(setup?.rowBreaks ?? []);
  return {
    cols: cutBands(plan.cols, s, limitW, colBreaks, plan.colMerges),
    rows: cutBands(plan.rows, s, limitH, rowBreaks, plan.rowMerges),
  };
}

/** Step 4: the percent the sheet prints at. */
function chooseScale(plans: readonly AreaPlan[], geometry: SheetGeometry, setup: SheetPageSetup | undefined): number {
  const fitTo = setup?.fitTo;
  const pages = setup?.odsScaleToPages;
  if (fitTo === undefined && pages === undefined) {
    return setup?.scale ?? 100;
  }
  for (let p = 100; p >= 10; p -= 1) {
    const s = p / 100;
    const fits = plans.every((plan) => {
      const bands = bandsAt(plan, geometry, s, true, setup);
      if (fitTo !== undefined) {
        return (fitTo.width === 0 || bands.cols.length <= fitTo.width) && (fitTo.height === 0 || bands.rows.length <= fitTo.height);
      }
      return true;
    });
    const total =
      pages === undefined || fitTo !== undefined
        ? 0
        : plans.reduce((count, plan) => {
            const bands = bandsAt(plan, geometry, s, true, setup);
            return count + bands.cols.length * bands.rows.length;
          }, 0);
    if (fits && (pages === undefined || fitTo !== undefined || pages <= 0 || total <= pages)) {
      return p;
    }
  }
  return 10;
}

function titleTrack(sheet: Sheet, rows: Map<number, Sheet["rows"][number]>, metrics: GridMetrics, span: [number, number] | undefined, axis: "row" | "col"): Track {
  if (span === undefined) {
    return { indices: [], sizes: [] };
  }
  return axis === "row" ? visibleRows(sheet, rows, span[0], Math.min(span[1], MAX_ROWS - 1)) : visibleColumns(metrics, span[0], Math.min(span[1], MAX_COLS - 1));
}

function headerFooterFor(setup: SheetPageSetup | undefined, indexInSheet: number, number: number): HeaderFooterPair {
  const hf = setup?.headerFooter;
  if (hf === undefined) {
    return {};
  }
  if (hf.differentFirst && indexInSheet === 0) {
    return hf.first ?? {};
  }
  if (hf.differentOddEven && number % 2 === 0) {
    return hf.even ?? {};
  }
  return hf.odd;
}

function overlaps(band: readonly number[], titles: readonly number[]): boolean {
  return titles.length > 0 && band.length > 0 && band[0]! <= titles.at(-1)! && band.at(-1)! >= titles[0]!;
}

/** The pages of one sheet (see the module header). */
export function paginateSheet(model: Pick<SpreadsheetModel, "styles" | "names">, sheet: Sheet, options: SheetPrintOptions = {}): SheetPrintJob {
  if (sheet.kind !== "grid") {
    return { pages: [] };
  }
  const metrics = buildGridMetrics(sheet);
  const ranges = printRanges(model, sheet, options, metrics);
  if (ranges.length === 0) {
    return { pages: [] };
  }
  const setup = sheet.pageSetup;
  const rows = rowIndex(sheet);
  const plans: AreaPlan[] = [];
  for (const range of ranges) {
    const touching = sheet.merges.filter((merge) => merge.top <= range.bottom && merge.bottom >= range.top && merge.left <= range.right && merge.right >= range.left);
    const plan: AreaPlan = {
      cols: visibleColumns(metrics, range.left, range.right),
      rows: visibleRows(sheet, rows, range.top, range.bottom),
      colMerges: touching.filter((merge) => merge.left < merge.right).map((merge) => ({ first: merge.left, last: merge.right })),
      rowMerges: touching.filter((merge) => merge.top < merge.bottom).map((merge) => ({ first: merge.top, last: merge.bottom })),
    };
    if (plan.cols.indices.length > 0 && plan.rows.indices.length > 0) {
      plans.push(plan);
    }
  }
  if (plans.length === 0) {
    return { pages: [] };
  }
  const paper = orientedPaper(sheet, options.preferences);
  const margins = effectiveMargins(setup);
  const paperPt = { width: paper.widthCm * PT_PER_CM, height: paper.heightCm * PT_PER_CM };
  const marginsPt = {
    top: margins.topCm * PT_PER_CM,
    right: margins.rightCm * PT_PER_CM,
    bottom: margins.bottomCm * PT_PER_CM,
    left: margins.leftCm * PT_PER_CM,
    header: margins.headerCm * PT_PER_CM,
    footer: margins.footerCm * PT_PER_CM,
  };
  const titles = printTitlesOf(model, sheet);
  const lastRow = Math.max(...plans.map((plan) => plan.rows.indices.at(-1)!));
  const geometry: SheetGeometry = {
    boxW: paperPt.width - marginsPt.left - marginsPt.right,
    boxH: paperPt.height - marginsPt.top - marginsPt.bottom,
    titleRows: titleTrack(sheet, rows, metrics, titles?.rows, "row"),
    titleCols: titleTrack(sheet, rows, metrics, titles?.cols, "col"),
    headingW: setup?.printHeadings ? headingColumnPt(lastRow) : 0,
    headingH: setup?.printHeadings ? HEADING_ROW_PT : 0,
  };
  const fit = setup?.fitTo !== undefined || setup?.odsScaleToPages !== undefined;
  const percent = chooseScale(plans, geometry, setup);
  const s = percent / 100;
  const pages: SheetPrintPage[] = [];
  let number = setup?.firstPageNumber ?? options.startNumber ?? 1;
  plans.forEach((plan, area) => {
    const bands = bandsAt(plan, geometry, s, fit, setup);
    const order: [Band, Band][] = [];
    if (setup?.pageOrder === "overThenDown") {
      for (const rowBand of bands.rows) {
        for (const colBand of bands.cols) {
          order.push([colBand, rowBand]);
        }
      }
    } else {
      for (const colBand of bands.cols) {
        for (const rowBand of bands.rows) {
          order.push([colBand, rowBand]);
        }
      }
    }
    for (const [colBand, rowBand] of order) {
      const cols = plan.cols.indices.slice(colBand.start, colBand.end + 1);
      const rowsOnPage = plan.rows.indices.slice(rowBand.start, rowBand.end + 1);
      const titleRows = overlaps(rowsOnPage, geometry.titleRows.indices) ? [] : geometry.titleRows.indices;
      const titleCols = overlaps(cols, geometry.titleCols.indices) ? [] : geometry.titleCols.indices;
      const bandW = sum(plan.cols.sizes.slice(colBand.start, colBand.end + 1));
      const bandH = sum(plan.rows.sizes.slice(rowBand.start, rowBand.end + 1));
      const titlesW = titleCols.length === 0 ? 0 : sum(geometry.titleCols.sizes);
      const titlesH = titleRows.length === 0 ? 0 : sum(geometry.titleRows.sizes);
      const blockW = s * (geometry.headingW + titlesW + bandW);
      const blockH = s * (geometry.headingH + titlesH + bandH);
      const originPt = {
        x: marginsPt.left + (setup?.centerHorizontally ? Math.max(0, (geometry.boxW - blockW) / 2) : 0),
        y: marginsPt.top + (setup?.centerVertically ? Math.max(0, (geometry.boxH - blockH) / 2) : 0),
      };
      const pair = headerFooterFor(setup, pages.length, number);
      const page: SheetPrintPage = {
        sheetId: sheet.id,
        number,
        paperPt: { ...paperPt },
        marginsPt: { ...marginsPt },
        scale: percent,
        rows: rowsOnPage,
        cols,
        titleRows: [...titleRows],
        titleCols: [...titleCols],
        originPt,
        area,
      };
      if (pair.header !== undefined) {
        page.header = { ...pair.header };
      }
      if (pair.footer !== undefined) {
        page.footer = { ...pair.footer };
      }
      pages.push(page);
      number += 1;
    }
  });
  return { pages };
}

/** The pages of a print job: the active sheet ("sheet", "selection") or
 * every visible grid sheet in order ("workbook"), numbered on across sheets. */
export function paginateWorkbook(model: SpreadsheetModel, scope: PrintScope, options: SheetPrintOptions = {}): SheetPrintJob {
  if (scope !== "workbook") {
    const sheet = model.sheets[model.activeSheetIndex];
    return sheet === undefined ? { pages: [] } : paginateSheet(model, sheet, { ...options, scope });
  }
  const pages: SheetPrintPage[] = [];
  let next = options.startNumber ?? 1;
  for (const sheet of model.sheets) {
    if (sheet.kind !== "grid" || sheet.visibility !== "visible") {
      continue;
    }
    const job = paginateSheet(model, sheet, { ...options, scope: "sheet", startNumber: next });
    pages.push(...job.pages);
    if (job.pages.length > 0) {
      next = job.pages.at(-1)!.number + 1;
    }
  }
  return { pages };
}

/** The first rows and columns of the sheet's pages after each area's first
 * band — every page break pagination put in, manual ones included. */
export function autoBreaks(job: SheetPrintJob, sheetId: string): { rows: number[]; cols: number[] } {
  const rows = new Set<number>();
  const cols = new Set<number>();
  const firstRow = new Map<number, number>();
  const firstCol = new Map<number, number>();
  const own = job.pages.filter((page) => page.sheetId === sheetId);
  for (const page of own) {
    firstRow.set(page.area, Math.min(firstRow.get(page.area) ?? Infinity, page.rows[0]!));
    firstCol.set(page.area, Math.min(firstCol.get(page.area) ?? Infinity, page.cols[0]!));
  }
  for (const page of own) {
    if (page.rows[0] !== firstRow.get(page.area)) {
      rows.add(page.rows[0]!);
    }
    if (page.cols[0] !== firstCol.get(page.area)) {
      cols.add(page.cols[0]!);
    }
  }
  return { rows: [...rows].sort((a, b) => a - b), cols: [...cols].sort((a, b) => a - b) };
}

export interface PageBreakLine {
  axis: "row" | "col";
  /** The first row/column of the next page. */
  index: number;
  kind: "auto" | "manual";
}

/** The grid's page-break lines for a sheet: every break of its pages, a
 * manual one where the setup broke the sheet there. */
export function pageBreakLines(job: SheetPrintJob, sheet: Pick<Sheet, "id" | "pageSetup">): PageBreakLine[] {
  const breaks = autoBreaks(job, sheet.id);
  const manualRows = new Set(sheet.pageSetup?.rowBreaks ?? []);
  const manualCols = new Set(sheet.pageSetup?.colBreaks ?? []);
  return [
    ...breaks.rows.map((index): PageBreakLine => ({ axis: "row", index, kind: manualRows.has(index) ? "manual" : "auto" })),
    ...breaks.cols.map((index): PageBreakLine => ({ axis: "col", index, kind: manualCols.has(index) ? "manual" : "auto" })),
  ];
}

// --- printing -------------------------------------------------------------------------------

function pageNameOf(page: Pick<SheetPrintPage, "paperPt">): string {
  return `sheet-page-${Math.round(page.paperPt.width * 100)}x${Math.round(page.paperPt.height * 100)}`;
}

function sizeOf(page: Pick<SheetPrintPage, "paperPt">): string {
  return `${round2(page.paperPt.width)}pt ${round2(page.paperPt.height)}pt`;
}

function round2(value: number): number {
  return Math.round(value * 100) / 100;
}

/** The `@page` rules a print job prints under: one unnamed rule from the
 * first page's paper and one named rule per distinct paper and orientation,
 * all with no margin (the pages carry their own); `pageNames[i]` is the
 * `page:` name of page i. */
export function sheetPrintPageRuleCss(job: SheetPrintJob): { css: string; pageNames: string[] } {
  const first = job.pages[0];
  if (first === undefined) {
    return { css: "", pageNames: [] };
  }
  const rules = [`@page { size: ${sizeOf(first)}; margin: 0; }`];
  const seen = new Set<string>();
  for (const page of job.pages) {
    const name = pageNameOf(page);
    if (!seen.has(name)) {
      seen.add(name);
      rules.push(`@page ${name} { size: ${sizeOf(page)}; margin: 0; }`);
    }
  }
  return { css: rules.join("\n"), pageNames: job.pages.map(pageNameOf) };
}

/** The paper of a page as a PageSize in cm. */
export function pagePaperCm(page: Pick<SheetPrintPage, "paperPt">): PageSize {
  return { widthCm: page.paperPt.width / PT_PER_CM, heightCm: page.paperPt.height / PT_PER_CM };
}
