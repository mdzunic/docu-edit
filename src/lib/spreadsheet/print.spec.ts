// SPEC-136 "The pagination algorithm", step by step against hand-computed
// fixtures. Every model is deep-frozen: pagination never writes it.
//
// The reference page: A4 portrait (595.28 × 841.89 pt) with the default
// margins (left/right 0.7 in = 50.4 pt, top/bottom 0.75 in = 54 pt), so the
// printable box is 494.48 × 733.89 pt — ten 48 pt (64 px) columns and
// forty-eight 15 pt rows a page at 100%.
import { describe, expect, it } from "vitest";
import type { GridRange } from "./address";
import type { Cell, DefinedName, Row, Sheet, SheetPageSetup, SpreadsheetModel } from "./model";
import { autoBreaks, cutBands, pageBreakLines, paginateSheet, paginateWorkbook, printableUsedRange, sheetPrintPageRuleCss } from "./print";

function deepFreeze<T>(value: T): T {
  if (value !== null && typeof value === "object" && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const key of Object.keys(value)) {
      deepFreeze((value as Record<string, unknown>)[key]);
    }
  }
  return value;
}

const cell = (col: number, value: Cell["value"] = 1): Cell => ({ col, value });
const row = (index: number, cells: Cell[], extra: Partial<Row> = {}): Row => ({ index, cells, ...extra });

function sheet(id: string, rows: Row[], extra: Partial<Sheet> = {}): Sheet {
  return { id, name: id, visibility: "visible", kind: "grid", rows, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], conditionalFormats: [], validations: [], ...extra };
}

function book(sheets: Sheet[], names: DefinedName[] = []): SpreadsheetModel {
  return deepFreeze({
    kind: "spreadsheet",
    id: "b",
    title: "Book",
    format: "xlsx",
    readerVersion: 16,
    sheets,
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{}, { fill: { kind: "solid", color: { kind: "rgb", value: "#ffff00" } } }],
    defaultStyleId: 0,
    dxfs: [],
    names,
  } as SpreadsheetModel);
}

/** A sheet with values at A1 and at (lastRow, lastCol). */
function corners(lastRow: number, lastCol: number, extra: Partial<Sheet> = {}): Sheet {
  return sheet("s1", [row(0, [cell(0)]), row(lastRow, [cell(lastCol)])], extra);
}

const A4 = { preferences: { defaultPageSizeId: "A4" } };
const bands = (pages: { rows: number[]; cols: number[] }[]): string[] => pages.map((p) => `${p.rows[0]}-${p.rows.at(-1)}/${p.cols[0]}-${p.cols.at(-1)}`);

describe("paginateSheet", () => {
  it("1. the area: the printable used range — values, fills and borders, merges touching it, drawings; empty → no pages", () => {
    expect(paginateSheet(book([sheet("s1", [])]), sheet("s1", [])).pages).toEqual([]);
    const filled = sheet("s1", [row(2, [{ col: 4, value: null, styleId: 1 }]), row(5, [cell(1)])], { merges: [{ top: 5, left: 1, bottom: 7, right: 2 }] });
    expect(printableUsedRange(book([filled]), filled)).toEqual({ top: 2, left: 1, bottom: 7, right: 4 });
    const drawn = sheet("s1", [row(0, [cell(0)])], {
      drawings: [{ id: "dr1", name: "P", anchor: { kind: "twoCell", from: { col: 3, colOffEmu: 0, row: 3, rowOffEmu: 0 }, to: { col: 5, colOffEmu: 0, row: 8, rowOffEmu: 0 } }, content: { kind: "image", mediaId: "m" } }],
    });
    expect(printableUsedRange(book([drawn]), drawn)).toEqual({ top: 0, left: 0, bottom: 7, right: 4 });
  });

  it("1. the area: a print area's ranges, each from a new page, whole columns clipped to the used range; the selection for scope selection", () => {
    const s = corners(9, 9);
    const area: DefinedName = { name: "_xlnm.Print_Area", scopeSheetId: "s1", formula: "s1!$A$1:$B$2,s1!$D:$D" };
    const job = paginateSheet(book([s], [area]), s, A4);
    expect(bands(job.pages)).toEqual(["0-1/0-1", "0-9/3-3"]);
    expect(job.pages.map((p) => p.area)).toEqual([0, 1]);
    const selection: GridRange[] = [{ top: 2, left: 2, bottom: 3, right: 3 }];
    expect(bands(paginateSheet(book([s]), s, { ...A4, scope: "selection", selection }).pages)).toEqual(["2-3/2-3"]);
  });

  it("2. hidden and filter-hidden rows and hidden columns are skipped", () => {
    const s = sheet("s1", [row(0, [cell(0)]), row(1, [cell(0)], { hidden: true }), row(2, [cell(0)], { filterHidden: true }), row(3, [cell(4)])], { columns: [{ min: 2, max: 2, hidden: true }] });
    const [page] = paginateSheet(book([s]), s, A4).pages;
    expect(page!.rows).toEqual([0, 3]);
    expect(page!.cols).toEqual([0, 1, 3, 4]);
  });

  it("3, 5–7. ten columns and forty-eight rows a page; down then over by default, over then down when set", () => {
    const s = corners(99, 25);
    const down = paginateSheet(book([s]), s, A4).pages;
    expect(bands(down)).toEqual(["0-47/0-9", "48-95/0-9", "96-99/0-9", "0-47/10-19", "48-95/10-19", "96-99/10-19", "0-47/20-25", "48-95/20-25", "96-99/20-25"]);
    const over = corners(99, 25, { pageSetup: { pageOrder: "overThenDown" } });
    expect(bands(paginateSheet(book([over]), over, A4).pages).slice(0, 4)).toEqual(["0-47/0-9", "0-47/10-19", "0-47/20-25", "48-95/0-9"]);
  });

  it("3. landscape, the paper, margins and headings shrink or grow the box", () => {
    const landscape = corners(99, 25, { pageSetup: { orientation: "landscape", paper: { widthCm: 21, heightCm: 29.7, xlsxPaperSize: 9 } } });
    const [first] = paginateSheet(book([landscape]), landscape).pages;
    expect(first!.paperPt.width).toBeCloseTo(841.89, 2);
    // (841.89 − 100.8) / 48 = 15.4 → 15 columns; (595.28 − 108) / 15 = 32.5 → 32 rows.
    expect(bands([first!])).toEqual(["0-31/0-14"]);
    const headed = corners(99, 25, { pageSetup: { printHeadings: true } });
    // Headings: a 24 pt column (max(24, 6 × 3 + 6)) and a 15 pt row → 9 columns, 47 rows.
    const [page] = paginateSheet(book([headed]), headed, A4).pages;
    expect(bands([page!])).toEqual(["0-46/0-8"]);
    expect(page!.headingsPt).toEqual({ width: 24, height: 15 });
  });

  it("5–6. a manual break starts a band in percent mode", () => {
    const s = corners(99, 25, { pageSetup: { rowBreaks: [10], colBreaks: [4] } });
    expect(bands(paginateSheet(book([s]), s, A4).pages).slice(0, 4)).toEqual(["0-9/0-3", "10-57/0-3", "58-99/0-3", "0-9/4-13"]);
  });

  it("4. fit to 1 wide × automatic: the largest percent at which the columns fit; manual breaks ignored", () => {
    const s = corners(99, 25, { pageSetup: { fitTo: { width: 1, height: 0 }, rowBreaks: [10] } });
    const pages = paginateSheet(book([s]), s, A4).pages;
    // 26 × 48 × p/100 ≤ 494.48 → p = 39; 733.89 / (15 × 0.39) = 125 rows a page.
    expect(pages[0]!.scale).toBe(39);
    expect(bands(pages)).toEqual(["0-99/0-25"]);
    const tall = corners(299, 0, { pageSetup: { fitTo: { width: 0, height: 2 } } });
    // 300 rows into 2 bands: 150 × 15 × p/100 ≤ 733.89 → p = 32.
    expect(paginateSheet(book([tall]), tall, A4).pages.map((p) => p.scale)).toEqual([32, 32]);
    const pagesTotal = corners(99, 25, { pageSetup: { odsScaleToPages: 2 } });
    const two = paginateSheet(book([pagesTotal]), pagesTotal, A4).pages;
    // 2 wide × 1 tall: 100 × 15 × p/100 ≤ 733.89 → p = 48 (26 × 48 × 0.48 = 599 pt, two column bands).
    expect(two).toHaveLength(2);
    expect(two[0]!.scale).toBe(48);
  });

  it("4. nothing fits even at 10%: 10", () => {
    const s = corners(9999, 0, { pageSetup: { fitTo: { width: 1, height: 1 } } });
    expect(paginateSheet(book([s]), s, A4).pages[0]!.scale).toBe(10);
  });

  it("5–6. a boundary inside a merge moves back to its first row; a merge taller than a page splits at the box edge", () => {
    const moved = corners(99, 0, { merges: [{ top: 46, left: 0, bottom: 49, right: 1 }] });
    expect(bands(paginateSheet(book([moved]), moved, A4).pages).map((b) => b.split("/")[0])).toEqual(["0-45", "46-93", "94-99"]);
    const tall = corners(99, 0, { merges: [{ top: 0, left: 0, bottom: 60, right: 0 }] });
    expect(bands(paginateSheet(book([tall]), tall, A4).pages).map((b) => b.split("/")[0])).toEqual(["0-47", "48-95", "96-99"]);
  });

  it("cutBands: a single column wider than the box is its own band", () => {
    expect(cutBands({ indices: [0, 1, 2], sizes: [10, 500, 10] }, 1, 100, null, [])).toEqual([
      { start: 0, end: 0 },
      { start: 1, end: 1 },
      { start: 2, end: 2 },
    ]);
  });

  it("8. print titles repeat on pages whose band does not reach them, and take their height off the box", () => {
    const s = corners(99, 0);
    const titles: DefinedName = { name: "_xlnm.Print_Titles", scopeSheetId: "s1", formula: "s1!$1:$1" };
    const pages = paginateSheet(book([s], [titles]), s, A4).pages;
    // (733.89 − 15) / 15 = 47.9 → 47 rows a band.
    expect(bands(pages).map((b) => b.split("/")[0])).toEqual(["0-46", "47-93", "94-99"]);
    expect(pages.map((p) => p.titleRows)).toEqual([[], [0], [0]]);
  });

  it("9. placement: the block at the box's top-left, centred on each axis set to centre", () => {
    const s = sheet("s1", [row(0, [cell(0)])], { pageSetup: { centerHorizontally: true, centerVertically: true } });
    const [page] = paginateSheet(book([s]), s, A4).pages;
    expect(page!.originPt.x).toBeCloseTo(50.4 + (494.48 - 48) / 2, 1);
    expect(page!.originPt.y).toBeCloseTo(54 + (733.89 - 15) / 2, 1);
    const plain = sheet("s1", [row(0, [cell(0)])]);
    expect(paginateSheet(book([plain]), plain, A4).pages[0]!.originPt).toEqual({ x: expect.closeTo(50.4, 1), y: expect.closeTo(54, 1) });
  });

  it("10. numbering: from firstPageNumber, on across sheets in a workbook job, restarting where a sheet sets its own", () => {
    const one = sheet("s1", [row(0, [cell(0)]), row(60, [cell(0)])]);
    const two = sheet("s2", [row(0, [cell(0)])]);
    const three: Sheet = { ...sheet("s3", [row(0, [cell(0)])]), pageSetup: { firstPageNumber: 10 } };
    const hidden: Sheet = { ...sheet("s4", [row(0, [cell(0)])]), visibility: "hidden" };
    const job = paginateWorkbook(book([one, two, three, hidden]), "workbook", A4);
    expect(job.pages.map((p) => `${p.sheetId}:${p.number}`)).toEqual(["s1:1", "s1:2", "s2:3", "s3:10"]);
    expect(paginateWorkbook(book([one, two]), "sheet", A4).pages).toHaveLength(2);
  });

  it("header and footer per page: the first page's with differentFirst, the even pages' with differentOddEven", () => {
    const setup: SheetPageSetup = {
      headerFooter: { odd: { header: { center: "odd" } }, even: { header: { center: "even" } }, first: { header: { center: "first" } }, differentFirst: true, differentOddEven: true },
    };
    const s = corners(150, 0, { pageSetup: setup });
    expect(paginateSheet(book([s]), s, A4).pages.map((p) => p.header?.center)).toEqual(["first", "even", "odd", "even"]);
  });

  it("the paper: the preference without a setup, Letter for an unknown xlsx code", () => {
    const s = corners(0, 0);
    expect(paginateSheet(book([s]), s, { preferences: { defaultPageSizeId: "Letter" } }).pages[0]!.paperPt.width).toBeCloseTo(612, 0);
    const odd = corners(0, 0, { pageSetup: { paper: { widthCm: 21.59, heightCm: 27.94, xlsxPaperSize: 42 } } });
    expect(paginateSheet(book([odd]), odd, A4).pages[0]!.paperPt.height).toBeCloseTo(792, 0);
  });

  it("autoBreaks and pageBreakLines: every band start after the first, manual where the setup broke it", () => {
    const s = corners(99, 12, { pageSetup: { rowBreaks: [10] } });
    const job = paginateSheet(book([s]), s, A4);
    expect(autoBreaks(job, "s1")).toEqual({ rows: [10, 58], cols: [10] });
    expect(pageBreakLines(job, s)).toEqual([
      { axis: "row", index: 10, kind: "manual" },
      { axis: "row", index: 58, kind: "auto" },
      { axis: "col", index: 10, kind: "auto" },
    ]);
  });

  it("sheetPrintPageRuleCss: an unnamed rule from the first page and a named one per paper, margin 0", () => {
    const portrait = corners(0, 0);
    const landscape: Sheet = { ...corners(0, 0), id: "s2", pageSetup: { orientation: "landscape" } };
    const job = paginateWorkbook(book([portrait, landscape]), "workbook", A4);
    const { css, pageNames } = sheetPrintPageRuleCss(job);
    expect(css.split("\n")).toEqual([
      "@page { size: 595.28pt 841.89pt; margin: 0; }",
      "@page sheet-page-59528x84189 { size: 595.28pt 841.89pt; margin: 0; }",
      "@page sheet-page-84189x59528 { size: 841.89pt 595.28pt; margin: 0; }",
    ]);
    expect(pageNames).toEqual(["sheet-page-59528x84189", "sheet-page-84189x59528"]);
  });
});
