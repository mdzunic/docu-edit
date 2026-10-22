// A printed sheet page as positioned boxes (SPEC-136 "Printed pages are
// content, not screenshots of the grid") — framework-free, so
// SheetPrintPage.vue only renders what this lays out. Every length is CSS px
// on the page at 1:1 (pt × 4/3), the grid's own lengths multiplied by the
// page's scale: the scale is applied by layout — widths, heights, font sizes,
// borders and padding — never by a transform, because the PDF extractor
// reads computed font sizes.
//
// - A cell carries its resolved style (the cell's, else its row's, else its
//   column's; a table's look under it) as CSS through SPEC-117's cellCss, its
//   value through SPEC-118's formatting at the column (or merge) width, so a
//   number too wide shows "########". A string that does not fit spills into
//   empty neighbours by SPEC-117's overflowExtent, bounded by the page's band
//   — drawn by a transparent spill box, so the PDF clips it there too.
// - Gridlines (printGridLines) are 1 px #000000 edges where a cell has no
//   border of its own; headings (printHeadings) are `sheet-print-heading`
//   boxes, the letters and numbers centred in the default font.
// - An http/https/mailto link (a stored one or a top-level HYPERLINK()) is
//   the cell's `href`: the page renders it as `a[data-doc-link]`, which the
//   extractor makes a /Link annotation.
// - Header and footer sections are rendered for the page's number and the
//   job's total, in the workbook's default font at 11 pt (scaled with the
//   sheet unless scaleWithDoc is off), aligned to the margins unless
//   alignWithMargins is off.
import { DEFAULT_THEME, type DocumentTheme } from "../../lib/core/theme";
import { buildGridMetrics } from "../../lib/spreadsheet/geometry";
import { buildMergeIndex } from "../../lib/spreadsheet/merges";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { workbookFormatContext } from "../../lib/spreadsheet/numberFormat/render";
import { colToLetters, formatA1, getCell } from "../../lib/spreadsheet/address";
import { externalHref, openHyperlinkTarget } from "../../lib/spreadsheet/annotations";
import { borderShorthand, cellCss, cellDisplay, cellText, edgeColor, estimateTextWidth, overflowAlignment, overflowExtent, valueKind } from "../../lib/spreadsheet/cellRender";
import { drawingRectPx } from "../../lib/spreadsheet/drawings";
import { CELL_PADDING_PX, type GridMetrics } from "../../lib/spreadsheet/geometry";
import { renderHeaderFooterSection, type HeaderFooterContext, type HfRun } from "../../lib/spreadsheet/headerFooter";
import { mergeAt, type MergeIndex } from "../../lib/spreadsheet/merges";
import type { CellBorderEdge, CellStyle, HeaderFooterSections, Sheet, SheetDrawing, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import { colWidthPt, rowHeightPt, type SheetPrintPage } from "../../lib/spreadsheet/print";
import { cellStyle, effectiveStyleId, fontMaxDigitWidth } from "../../lib/spreadsheet/styles";
import { styleOverLook, tableCellLook } from "../../lib/spreadsheet/tableStyles";
import { linkTargetAt } from "./cellLinks";

export const PX_PER_PT = 4 / 3;
const GRIDLINE = "1px solid #000000";
const HEADER_FONT_PT = 11;

export type Css = Record<string, string>;

/** Everything a sheet's pages share. */
export interface PrintSheetContext {
  model: SpreadsheetModel;
  sheet: Sheet;
  metrics: GridMetrics;
  merges: MergeIndex;
  theme: DocumentTheme;
  format: FormatContext;
  rows: Map<number, Sheet["rows"][number]>;
  /** The default font's family, as CSS (`fontStack` applied by the page). */
  fontFamily: string | undefined;
  /** What headers and footers fill in, but the page numbers. */
  fields: Omit<HeaderFooterContext, "page" | "pages">;
}

export interface PrintCellItem {
  key: string;
  ref: string;
  left: number;
  top: number;
  width: number;
  height: number;
  box: Css;
  text: Css;
  content: string;
  href?: string;
  heading?: true;
  spill?: true;
  /** A merge's span on the page. */
  rowSpan?: number;
  colSpan?: number;
}

export interface PrintRun {
  text: string;
  css: Css;
}

export interface PrintBand {
  left: PrintRun[];
  center: PrintRun[];
  right: PrintRun[];
  /** px from the page's top (header) or bottom (footer) edge. */
  offset: number;
  insetLeft: number;
  insetRight: number;
}

export interface PrintDrawingItem {
  drawing: SheetDrawing;
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface PrintPageLayout {
  widthPx: number;
  heightPx: number;
  cells: PrintCellItem[];
  header: PrintBand | null;
  footer: PrintBand | null;
  /** The band's box (drawings are clipped to it). */
  bandBox: { left: number; top: number; width: number; height: number };
  drawings: PrintDrawingItem[];
  scale: number;
  baseFontPt: number;
}

/** Multiplies every px/pt length in a CSS value by `s`. */
export function scaleLengths(value: string, s: number): string {
  return value.replace(/(-?\d*\.?\d+)(px|pt)/g, (_, n: string, unit: string) => `${Math.round(Number(n) * s * 1000) / 1000}${unit}`);
}

function scaled(css: Css, keys: readonly string[], s: number): Css {
  const out = { ...css };
  for (const key of keys) {
    if (out[key] !== undefined) {
      out[key] = scaleLengths(out[key]!, s);
    }
  }
  return out;
}

function runCss(run: HfRun, s: number): Css {
  const css: Css = {};
  if (run.bold) {
    css["font-weight"] = "bold";
  }
  if (run.italic) {
    css["font-style"] = "italic";
  }
  const lines = [run.underline !== undefined ? "underline" : "", run.strike ? "line-through" : ""].filter((each) => each !== "");
  if (lines.length > 0) {
    css["text-decoration"] = lines.join(" ");
  }
  if (run.underline === "double") {
    css["text-decoration-style"] = "double";
  }
  if (run.fontFamily !== undefined) {
    css["font-family"] = run.fontFamily;
  }
  if (run.fontSizePt !== undefined) {
    css["font-size"] = `${Math.round(run.fontSizePt * s * 100) / 100}pt`;
  }
  if (run.color !== undefined) {
    css.color = run.color;
  }
  return css;
}

/** The style a cell prints in: its own (cell → row → column) over its
 * table's look. */
function printedStyle(ctx: PrintSheetContext, row: number, col: number): CellStyle {
  const own = cellStyle(ctx.model, effectiveStyleId(ctx.sheet, row, col));
  const table = ctx.sheet.tables?.find((each) => each.ref.top <= row && each.ref.bottom >= row && each.ref.left <= col && each.ref.right >= col);
  const look = table === undefined ? undefined : tableCellLook(ctx.sheet, table, row, col, ctx.theme);
  return look === undefined ? own : styleOverLook(look, own, cellStyle(ctx.model, 0));
}

function edgeCss(edge: CellBorderEdge | undefined, theme: DocumentTheme, s: number): string | undefined {
  if (edge === undefined) {
    return undefined;
  }
  return scaleLengths(borderShorthand(edge.style, edgeColor(edge, theme)), s);
}

function band(sections: HeaderFooterSections | undefined, ctx: PrintSheetContext, page: SheetPrintPage, pages: number, offsetPt: number): PrintBand | null {
  if (sections === undefined) {
    return null;
  }
  const hf = ctx.sheet.pageSetup?.headerFooter;
  const s = hf?.scaleWithDoc === false ? 1 : page.scale / 100;
  const fields: HeaderFooterContext = { ...ctx.fields, page: page.number, pages };
  const runs = (code: string | undefined): PrintRun[] =>
    code === undefined ? [] : renderHeaderFooterSection(code, fields).map((run) => ({ text: run.text, css: runCss(run, s) }));
  const inset = (marginPt: number): number => (hf?.alignWithMargins === false ? 0.5 * 72 : marginPt) * PX_PER_PT;
  return {
    left: runs(sections.left),
    center: runs(sections.center),
    right: runs(sections.right),
    offset: offsetPt * PX_PER_PT,
    insetLeft: inset(page.marginsPt.left),
    insetRight: inset(page.marginsPt.right),
  };
}

/** The [start, end) of the title or band segment an index lies in. */
function segmentOf(index: number, start: number): [number, number] {
  return index < start ? [0, start] : [start, Number.MAX_SAFE_INTEGER];
}

/** A page laid out (see the module header). `pages` is the job's total. */
export function layoutPrintPage(ctx: PrintSheetContext, page: SheetPrintPage, pages: number): PrintPageLayout {
  const { model, sheet, metrics, theme } = ctx;
  const s = page.scale / 100;
  const setup = sheet.pageSetup;
  const k = s * PX_PER_PT;
  const headW = (page.headingsPt?.width ?? 0) * k;
  const headH = (page.headingsPt?.height ?? 0) * k;
  const x0 = page.originPt.x * PX_PER_PT;
  const y0 = page.originPt.y * PX_PER_PT;
  const colList = [...page.titleCols, ...page.cols];
  const rowList = [...page.titleRows, ...page.rows];
  const colX: number[] = [];
  const colW: number[] = [];
  let x = x0 + headW;
  for (const col of colList) {
    colX.push(x);
    const width = colWidthPt(metrics, col) * k;
    colW.push(width);
    x += width;
  }
  const rowY: number[] = [];
  const rowH: number[] = [];
  let y = y0 + headH;
  for (const row of rowList) {
    rowY.push(y);
    const height = rowHeightPt(sheet, row, ctx.rows) * k;
    rowH.push(height);
    y += height;
  }
  const cells: PrintCellItem[] = [];
  const bandColStart = page.titleCols.length;
  const bandRowStart = page.titleRows.length;
  // Headings: the corner, the column letters and the row numbers.
  if (page.headingsPt !== undefined) {
    const headingText: Css = { "font-size": `${Math.round(HEADER_FONT_PT * s * 100) / 100}pt`, "text-align": "center" };
    const headingBox: Css = { border: GRIDLINE, "justify-content": "center", "text-align": "center" };
    cells.push({ key: "h-corner", ref: "", left: x0, top: y0, width: headW, height: headH, box: headingBox, text: headingText, content: "", heading: true });
    colList.forEach((col, j) => {
      cells.push({ key: `h-c${col}-${j}`, ref: "", left: colX[j]!, top: y0, width: colW[j]!, height: headH, box: headingBox, text: headingText, content: colToLetters(col), heading: true });
    });
    rowList.forEach((row, i) => {
      cells.push({ key: `h-r${row}-${i}`, ref: "", left: x0, top: rowY[i]!, width: headW, height: rowH[i]!, box: headingBox, text: headingText, content: String(row + 1), heading: true });
    });
  }
  rowList.forEach((row, i) => {
    colList.forEach((col, j) => {
      const merge = mergeAt(ctx.merges, row, col);
      let rowSpan = 1;
      let colSpan = 1;
      if (merge !== undefined) {
        const [rowSegStart, rowSegEnd] = segmentOf(i, bandRowStart);
        const [colSegStart, colSegEnd] = segmentOf(j, bandColStart);
        const firstRow = rowList.findIndex((each, at) => at >= rowSegStart && at < rowSegEnd && each >= merge.top && each <= merge.bottom);
        const firstCol = colList.findIndex((each, at) => at >= colSegStart && at < colSegEnd && each >= merge.left && each <= merge.right);
        if (firstRow !== i || firstCol !== j) {
          return;
        }
        while (i + rowSpan < Math.min(rowList.length, rowSegEnd) && rowList[i + rowSpan]! <= merge.bottom) {
          rowSpan += 1;
        }
        while (j + colSpan < Math.min(colList.length, colSegEnd) && colList[j + colSpan]! <= merge.right) {
          colSpan += 1;
        }
      }
      const anchor = merge === undefined || (merge.top === row && merge.left === col);
      const style = printedStyle(ctx, row, col);
      const value = anchor ? (getCell(sheet, row, col)?.value ?? null) : null;
      const kind = valueKind(value);
      const mdw = fontMaxDigitWidth(style.font, theme);
      let width100 = 0;
      for (let at = j; at < j + colSpan; at += 1) {
        width100 += colW[at]! / k / 0.75;
      }
      const formatted = cellDisplay(value, style, sheet, { ...ctx.format, ...(style.alignment?.shrinkToFit ? {} : { widthPx: width100 }), digitWidthPx: mdw });
      const css = cellCss(style, theme, kind, formatted.alignHint);
      const laid = cellText(formatted, style, kind, width100, mdw);
      const content = laid.fill ? laid.fill.before + laid.fill.run + laid.fill.after : laid.text;
      const box: Css = { ...css.box };
      const text: Css = scaled({ ...css.text, ...(formatted.color ? { color: formatted.color } : {}), ...(laid.shrunkFontSize ? { "font-size": laid.shrunkFontSize } : {}) }, ["font-size", "padding"], s);
      if (text["font-size"] === undefined) {
        text["font-size"] = `${Math.round(HEADER_FONT_PT * s * 100) / 100}pt`;
      }
      const borders = style.borders;
      const own = { top: edgeCss(borders?.top, theme, s), right: edgeCss(borders?.right, theme, s), bottom: edgeCss(borders?.bottom, theme, s), left: edgeCss(borders?.left, theme, s) };
      const grid = setup?.printGridLines === true;
      const top = own.top ?? (grid && i === 0 ? GRIDLINE : undefined);
      const left = own.left ?? (grid && j === 0 ? GRIDLINE : undefined);
      const right = own.right ?? (grid ? GRIDLINE : undefined);
      const bottom = own.bottom ?? (grid ? GRIDLINE : undefined);
      for (const [edge, line] of [
        ["border-top", top],
        ["border-right", right],
        ["border-bottom", bottom],
        ["border-left", left],
      ] as const) {
        if (line !== undefined) {
          box[edge] = line;
        }
      }
      let height = 0;
      for (let at = i; at < i + rowSpan; at += 1) {
        height += rowH[at]!;
      }
      let width = 0;
      for (let at = j; at < j + colSpan; at += 1) {
        width += colW[at]!;
      }
      const ref = formatA1({ row, col });
      const item: PrintCellItem = { key: `c-${i}-${j}`, ref, left: colX[j]!, top: rowY[i]!, width, height, box, text, content };
      if (rowSpan > 1) {
        item.rowSpan = rowSpan;
      }
      if (colSpan > 1) {
        item.colSpan = colSpan;
      }
      if (anchor && content !== "") {
        const target = linkTargetAt(model, sheet, { row, col });
        if (target?.kind === "url" && openHyperlinkTarget(target) === "external") {
          item.href = externalHref(target);
        }
      }
      // Overflow into empty neighbours, bounded by the band (SPEC-117's rule).
      const align = merge !== undefined || formatted.overflow === "hash" || content === "" ? null : overflowAlignment(style, value);
      if (align !== null && estimateTextWidth(content, mdw) > width100 - 2 * CELL_PADDING_PX) {
        const room = overflowExtent(sheet, row, col, align, metrics, ctx.merges);
        const [colSegStart, colSegEnd] = segmentOf(j, bandColStart);
        const segLeft = colX[colSegStart]!;
        const lastCol = Math.min(colList.length, colSegEnd) - 1;
        const segRight = colX[lastCol]! + colW[lastCol]!;
        const spillLeft = Math.max(segLeft, item.left - room.left * k * 0.75);
        const spillRight = Math.min(segRight, item.left + item.width + room.right * k * 0.75);
        if (spillLeft < item.left || spillRight > item.left + item.width) {
          cells.push({ ...item, content: "" });
          const spill: PrintCellItem = {
            key: `s-${i}-${j}`,
            ref,
            left: spillLeft,
            top: item.top,
            width: spillRight - spillLeft,
            height: item.height,
            box: { "text-align": box["text-align"] ?? "left", "justify-content": box["justify-content"] ?? "flex-end" },
            text: { ...text, "white-space": "pre" },
            content,
            spill: true,
          };
          if (item.href !== undefined) {
            spill.href = item.href;
          }
          cells.push(spill);
          return;
        }
      }
      cells.push(item);
    });
  });
  const pagesCount = pages;
  const bandLeft = colX[bandColStart] ?? x0 + headW;
  const bandTop = rowY[bandRowStart] ?? y0 + headH;
  const lastBandCol = colList.length - 1;
  const lastBandRow = rowList.length - 1;
  const bandBox = {
    left: bandLeft,
    top: bandTop,
    width: lastBandCol >= bandColStart ? colX[lastBandCol]! + colW[lastBandCol]! - bandLeft : 0,
    height: lastBandRow >= bandRowStart ? rowY[lastBandRow]! + rowH[lastBandRow]! - bandTop : 0,
  };
  const drawings: PrintDrawingItem[] = [];
  const firstCol = page.cols[0];
  const firstRow = page.rows[0];
  if (firstCol !== undefined && firstRow !== undefined) {
    const originX = metrics.colLeft(firstCol);
    const originY = metrics.rowTop(firstRow);
    const lastCol = page.cols.at(-1)!;
    const lastRow = page.rows.at(-1)!;
    const endX = metrics.colLeft(lastCol) + metrics.colWidth(lastCol);
    const endY = metrics.rowTop(lastRow) + metrics.rowHeight(lastRow);
    for (const drawing of sheet.drawings ?? []) {
      if (drawing.hidden) {
        continue;
      }
      const rect = drawingRectPx(drawing.anchor, metrics);
      if (rect.left >= endX || rect.top >= endY || rect.left + rect.width <= originX || rect.top + rect.height <= originY) {
        continue;
      }
      drawings.push({ drawing, left: (rect.left - originX) * s, top: (rect.top - originY) * s, width: rect.width, height: rect.height });
    }
  }
  return {
    widthPx: page.paperPt.width * PX_PER_PT,
    heightPx: page.paperPt.height * PX_PER_PT,
    cells,
    header: band(page.header, ctx, page, pagesCount, page.marginsPt.header),
    footer: band(page.footer, ctx, page, pagesCount, page.marginsPt.footer),
    bandBox,
    drawings,
    scale: s,
    baseFontPt: HEADER_FONT_PT * s,
  };
}

/** The workbook default style's font family (resolved through its theme by
 * cellCss), what headings and headers print in; undefined for none. */
export function defaultPrintFont(model: SpreadsheetModel): string | undefined {
  return cellCss(cellStyle(model, 0), model.theme ?? DEFAULT_THEME).text["font-family"];
}

/** A sheet's print context: its metrics, merges and formats in `locale`,
 * the header/footer fields of a job printed at `date`. */
export function printSheetContext(model: SpreadsheetModel, sheet: Sheet, locale: SpreadsheetLocale, date: Date): PrintSheetContext {
  const theme = model.theme ?? DEFAULT_THEME;
  return {
    model,
    sheet,
    metrics: buildGridMetrics(sheet),
    merges: buildMergeIndex(sheet.merges),
    theme,
    format: workbookFormatContext(model, locale),
    rows: new Map(sheet.rows.map((row) => [row.index, row])),
    fontFamily: defaultPrintFont(model),
    fields: { date, sheetName: sheet.name, fileName: model.title, locale, dateSystem: model.dateSystem },
  };
}

/** Typed pass-through for a layout coming from a `.vue` template's runtime
 * `Object` prop (espree bars the annotation in `.vue` scripts). */
export function asPrintPageLayout(value: unknown): PrintPageLayout {
  return value as PrintPageLayout;
}

/** A layout's header and footer bands, for the page's template. */
export function bandsOf(layout: PrintPageLayout): { band: PrintBand; edge: "top" | "bottom"; testid: string }[] {
  const out: { band: PrintBand; edge: "top" | "bottom"; testid: string }[] = [];
  if (layout.header !== null) {
    out.push({ band: layout.header, edge: "top", testid: "sheet-print-header" });
  }
  if (layout.footer !== null) {
    out.push({ band: layout.footer, edge: "bottom", testid: "sheet-print-footer" });
  }
  return out;
}

export const PRINT_SIDES: readonly ("left" | "center" | "right")[] = ["left", "center", "right"];

/** CSS px text for a length. */
export function cssPx(value: number): string {
  return `${Math.round(value * 1000) / 1000}px`;
}
