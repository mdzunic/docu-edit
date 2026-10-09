// SPEC-117 "Reading xlsx sheet layout": every rule of the worksheet reader —
// sheetFormatPr defaults (a baseColWidth-only sheet is 64 px for Calibri 11),
// cols (0-based spans, widths through the MDW formula with the source width
// kept, hidden, style, verbatim sourceAttrs, overlap clipping, the
// default-width test), rows (ht with or without customHeight, hidden, s only
// with customFormat, verbatim sourceAttrs, which empty rows are stored),
// style-only cells counted against the ceiling, merges (single-cell and
// overlapping ones rejected), the tab colour, the first sheetView (gridlines,
// zeros, zoom, top-left cell; frozen vs frozenSplit vs a split pane; the
// active pane's selection, normalised) and rPr fonts on shared and inline
// runs — over synthetic parts and spec117-styles.xlsx.
import { describe, expect, it } from "vitest";
import { getCell, parseA1, parseRange } from "../address";
import { buildSpec117StylesXlsxBytes, SPEC117_BORDER_STYLES, SPEC117_STYLES_XLSX, SPEC117_XF } from "../fixtures/build";
import { baseColWidthToPx, maxDigitWidthPx, xlsxWidthToPx } from "../geometry";
import { MAX_WORKBOOK_CELLS, type Sheet } from "../model";
import { CellBudget } from "../parsed";
import { INDEXED_COLORS } from "../styles";
import { parseXlsx } from "./reader";
import { readSharedStrings, readWorksheet, type SharedString, type WorksheetContent } from "./sheetXml";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
const CALIBRI_11 = maxDigitWidthPx("Calibri", 11);

/** A worksheet part around `body` (the root's children). */
function worksheet(body: string): string {
  return `<worksheet xmlns="${MAIN}" xmlns:x14ac="${X14AC}">${body}</worksheet>`;
}

/** Reads a worksheet whose root holds `body`, Calibri 11 metrics by default. */
function read(
  body: string,
  options: { sharedStrings?: SharedString[]; mdw?: number; budget?: CellBudget; indexedColors?: string[] } = {},
): WorksheetContent {
  const content = readWorksheet(worksheet(body), "S", {
    sharedStrings: options.sharedStrings ?? [],
    dateSystem: "1900",
    budget: options.budget ?? new CellBudget(),
    mdw: options.mdw ?? CALIBRI_11,
    indexedColors: options.indexedColors,
  });
  if (content === null) {
    throw new Error("the cell ceiling stopped the read");
  }
  return content;
}

/** A sheet with `data` as its sheetData. */
function rowsOf(data: string, sharedStrings?: SharedString[]): WorksheetContent["rows"] {
  return read(`<sheetData>${data}</sheetData>`, { sharedStrings }).rows;
}

describe("sheetFormatPr", () => {
  it("defaults to 15 pt rows and an 8-character base column when the element is absent", () => {
    const content = read(`<sheetData/>`);
    expect(content.defaultRowHeightPt).toBe(15);
    expect(content.defaultColWidthPx).toBe(64);
    expect(content).toEqual({ rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] });
  });

  it("reads defaultRowHeight in pt", () => {
    expect(read(`<sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.35"/><sheetData/>`).defaultRowHeightPt).toBe(14.5);
    expect(read(`<sheetFormatPr baseColWidth="8"/><sheetData/>`).defaultRowHeightPt).toBe(15);
  });

  it("converts defaultColWidth through the MDW formula", () => {
    const content = read(`<sheetFormatPr defaultColWidth="10" baseColWidth="20" defaultRowHeight="15"/><sheetData/>`);
    expect(content.defaultColWidthPx).toBe(xlsxWidthToPx(10, 7));
    expect(content.defaultColWidthPx).toBe(70);
  });

  it("falls back to baseColWidth (default 8): a baseColWidth-only sheet is 64 px for Calibri 11", () => {
    expect(read(`<sheetFormatPr baseColWidth="8" defaultRowHeight="15"/><sheetData/>`).defaultColWidthPx).toBe(64);
    expect(read(`<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/><sheetData/>`).defaultColWidthPx).toBe(80);
    expect(read(`<sheetFormatPr defaultRowHeight="15"/><sheetData/>`, { mdw: maxDigitWidthPx("Arial", 11) }).defaultColWidthPx).toBe(
      baseColWidthToPx(8, 8),
    );
  });
});

describe("cols", () => {
  it("reads 1-based min/max as 0-based spans, width → widthPx with the source width kept, hidden and style", () => {
    const { columns } = read(
      `<cols><col min="1" max="1" width="18.7109375" customWidth="1"/><col min="3" max="5" width="9.140625" hidden="1"/>` +
        `<col min="7" max="7" width="9.140625" style="4"/></cols><sheetData/>`,
    );
    expect(columns).toEqual([
      { min: 0, max: 0, widthPx: 131, sourceWidthChars: 18.7109375 },
      { min: 2, max: 4, widthPx: 64, hidden: true, sourceWidthChars: 9.140625 },
      { min: 6, max: 6, widthPx: 64, styleId: 4, sourceWidthChars: 9.140625 },
    ]);
  });

  it("keeps every other attribute verbatim, in source order, in sourceAttrs (customWidth is implied by a width)", () => {
    const { columns } = read(
      `<cols><col min="2" max="2" width="20" bestFit="1" customWidth="1" collapsed='1' outlineLevel="1" phonetic="1"/>` +
        `<col min="4" max="4" width="20" x14ac:foo="a&amp;b"/></cols><sheetData/>`,
    );
    expect(columns[0]!.sourceAttrs).toBe(`bestFit="1" collapsed='1' outlineLevel="1" phonetic="1"`);
    expect(columns[1]!.sourceAttrs).toBe(`x14ac:foo="a&amp;b"`);
  });

  it("stores no span for a column at the default width with nothing else", () => {
    const { columns } = read(`<cols><col min="1" max="3" width="9.140625" customWidth="1" bestFit="1"/><col min="4" max="4" style="2"/></cols><sheetData/>`);
    expect(columns).toEqual([{ min: 3, max: 3, styleId: 2 }]);
    // …measured against the sheet's own default.
    const wide = read(`<sheetFormatPr defaultColWidth="20" defaultRowHeight="15"/><cols><col min="1" max="1" width="9.140625"/></cols><sheetData/>`);
    expect(wide.columns).toEqual([{ min: 0, max: 0, widthPx: 64, sourceWidthChars: 9.140625 }]);
  });

  it("converts widths with the workbook's MDW", () => {
    const arial11 = maxDigitWidthPx("Arial", 11);
    const { columns } = read(`<cols><col min="1" max="1" width="12" customWidth="1"/></cols><sheetData/>`, { mdw: arial11 });
    expect(columns[0]!.widthPx).toBe(xlsxWidthToPx(12, arial11));
  });

  it("clips a later overlapping col to the gap the earlier ones leave", () => {
    const { columns } = read(
      `<cols><col min="2" max="3" width="20"/><col min="6" max="6" width="30"/>` +
        `<col min="1" max="8" width="40" hidden="1"/><col min="2" max="2" width="50"/></cols><sheetData/>`,
    );
    const w20 = xlsxWidthToPx(20, 7);
    const w30 = xlsxWidthToPx(30, 7);
    const w40 = xlsxWidthToPx(40, 7);
    expect(columns).toEqual([
      { min: 0, max: 0, widthPx: w40, hidden: true, sourceWidthChars: 40 },
      { min: 1, max: 2, widthPx: w20, sourceWidthChars: 20 },
      { min: 3, max: 4, widthPx: w40, hidden: true, sourceWidthChars: 40 },
      { min: 5, max: 5, widthPx: w30, sourceWidthChars: 30 },
      { min: 6, max: 7, widthPx: w40, hidden: true, sourceWidthChars: 40 },
    ]);
  });

  it("clamps a span reaching past XFD and skips unreadable cols", () => {
    const { columns } = read(
      `<cols><col min="16380" max="20000" hidden="1"/><col min="0" max="2" hidden="1"/><col min="5" max="3" hidden="1"/><col max="2" hidden="1"/></cols><sheetData/>`,
    );
    expect(columns).toEqual([{ min: 16379, max: 16383, hidden: true }]);
  });
});

describe("rows", () => {
  it("reads ht (with or without customHeight), customHeight and hidden", () => {
    expect(
      rowsOf(`<row r="1" ht="20.25" customHeight="1"><c r="A1"><v>1</v></c></row><row r="2" ht="18"><c r="A2"><v>2</v></c></row><row r="3" hidden="1"/>`),
    ).toEqual([
      { index: 0, cells: [{ col: 0, value: 1 }], heightPt: 20.25, customHeight: true },
      { index: 1, cells: [{ col: 0, value: 2 }], heightPt: 18 },
      { index: 2, cells: [], hidden: true },
    ]);
  });

  it("reads s as the row's style only with customFormat=\"1\"", () => {
    expect(rowsOf(`<row r="1" s="3" customFormat="1"/><row r="2" s="4"><c r="A2"><v>1</v></c></row><row r="3" s="5" customFormat="0"/>`)).toEqual([
      { index: 0, cells: [], styleId: 3 },
      { index: 1, cells: [{ col: 0, value: 1 }] },
    ]);
  });

  it("keeps unmodelled attributes verbatim, in source order, in sourceAttrs", () => {
    const [row] = rowsOf(
      `<row r="2" spans="1:5" s="1" customFormat="1" ht="30" hidden="1" customHeight="1" outlineLevel="1" collapsed="1" thickTop="1" thickBot="1" x14ac:dyDescent="0.25"/>`,
    );
    expect(row).toEqual({
      index: 1,
      cells: [],
      heightPt: 30,
      customHeight: true,
      hidden: true,
      styleId: 1,
      sourceAttrs: `spans="1:5" outlineLevel="1" collapsed="1" thickTop="1" thickBot="1" x14ac:dyDescent="0.25"`,
    });
    expect(rowsOf(`<row r="1" ht="15"><c r="A1"><v>1</v></c></row>`)[0]!.sourceAttrs).toBeUndefined();
    // The value is verbatim (its quotes and escapes); whitespace around = is not.
    expect(rowsOf(`<row r="1" ht = "15" spans = '1:2'\n x14ac:dyDescent="0&#46;25"/>`)[0]!.sourceAttrs).toBe(
      `spans='1:2' x14ac:dyDescent="0&#46;25"`,
    );
  });

  it("stores a row with no cells only when it has a height, hidden or a style", () => {
    expect(rowsOf(`<row r="1" spans="1:3" x14ac:dyDescent="0.25"/><row r="2" customHeight="1"/><row r="3" ht="12"/>`)).toEqual([
      { index: 2, cells: [], heightPt: 12 },
    ]);
  });
});

describe("cells", () => {
  it("reads s as the cell's styleId, as written — 0 included", () => {
    expect(rowsOf(`<row r="1"><c r="A1" s="0"><v>1</v></c><c r="B1" s="7" t="s"><v>0</v></c><c r="C1"><v>3</v></c></row>`, [{ text: "x" }])).toEqual([
      {
        index: 0,
        cells: [
          { col: 0, value: 1, styleId: 0 },
          { col: 1, value: "x", styleId: 7 },
          { col: 2, value: 3 },
        ],
      },
    ]);
  });

  it("stores a <c> carrying only a style as a style-only cell; one with neither stays unstored", () => {
    // SPEC-119: D4's formula is read too — with no cached value, `value: null`.
    expect(rowsOf(`<row r="4"><c r="A4" s="2"/><c r="B4"/><c r="C4" s="0"></c><c r="D4" s="3"><f>1+1</f></c></row>`)).toEqual([
      {
        index: 3,
        cells: [
          { col: 0, value: null, styleId: 2 },
          { col: 2, value: null, styleId: 0 },
          { col: 3, value: null, styleId: 3, formula: "1+1" },
        ],
      },
    ]);
  });

  it("counts style-only cells against MAX_WORKBOOK_CELLS", () => {
    const budget = new CellBudget();
    budget.take(MAX_WORKBOOK_CELLS - 1);
    const xml = worksheet(`<sheetData><row r="1"><c r="A1" s="1"/><c r="B1" s="1"/></row></sheetData>`);
    expect(readWorksheet(xml, "S", { sharedStrings: [], dateSystem: "1900", budget, mdw: 7 })).toBeNull();
    const roomy = new CellBudget();
    expect(readWorksheet(xml, "S", { sharedStrings: [], dateSystem: "1900", budget: roomy, mdw: 7 })!.rows[0]!.cells).toHaveLength(2);
    expect(roomy.used).toBe(2);
  });
});

describe("merges", () => {
  it("reads mergeCell@ref, keeping two-or-more-cell ranges that overlap no earlier one", () => {
    const { merges } = read(
      `<sheetData/><mergeCells count="6"><mergeCell ref="C5:E5"/><mergeCell ref="A1"/><mergeCell ref="D5:D9"/>` +
        `<mergeCell ref="B10:A8"/><mergeCell ref="bogus"/><mergeCell ref="G1:G1"/></mergeCells>`,
    );
    expect(merges).toEqual([parseRange("C5:E5"), parseRange("A8:B10")]);
  });
});

describe("sheetPr/tabColor", () => {
  it("reads the tab colour through the colour codec", () => {
    expect(read(`<sheetPr><tabColor theme="5" tint="-0.25"/></sheetPr><sheetData/>`).tabColor).toEqual({
      theme: "accent2",
      lumMod: 0.75,
      xlsxTint: "-0.25",
    });
    expect(read(`<sheetPr codeName="Sheet1"><tabColor rgb="FF00B050"/></sheetPr><sheetData/>`).tabColor).toBe("#00b050");
    const palette = [...INDEXED_COLORS];
    palette[12] = "#abcdef";
    expect(read(`<sheetPr><tabColor indexed="12"/></sheetPr><sheetData/>`, { indexedColors: palette }).tabColor).toBe("#abcdef");
    expect(read(`<sheetPr><tabColor auto="1"/></sheetPr><sheetData/>`)).not.toHaveProperty("tabColor");
  });
});

describe("sheetViews/sheetView[0]", () => {
  const view = (attrs: string, children = "") => read(`<sheetViews><sheetView ${attrs} workbookViewId="0">${children}</sheetView></sheetViews><sheetData/>`);

  it("reads showGridLines and showZeros \"0\" as false and leaves them absent otherwise", () => {
    expect(view(`showGridLines="0" showZeros="0"`)).toMatchObject({ showGridLines: false, showZeros: false });
    expect(view(`showGridLines="false" showZeros="false"`)).toMatchObject({ showGridLines: false, showZeros: false });
    const shown = view(`showGridLines="1" showZeros="1"`);
    expect(shown).not.toHaveProperty("showGridLines");
    expect(shown).not.toHaveProperty("showZeros");
  });

  it("reads zoomScale clamped to 10–400, 100 as absent", () => {
    expect(view(`zoomScale="85" zoomScaleNormal="85"`).zoomPercent).toBe(85);
    expect(view(`zoomScale="5"`).zoomPercent).toBe(10);
    expect(view(`zoomScale="900"`).zoomPercent).toBe(400);
    expect(view(`zoomScale="100"`)).not.toHaveProperty("zoomPercent");
    expect(view(``)).not.toHaveProperty("zoomPercent");
  });

  it("reads topLeftCell", () => {
    expect(view(`topLeftCell="C40"`).topLeftCell).toEqual({ row: 39, col: 2 });
  });

  it("reads only the first sheetView", () => {
    const content = read(
      `<sheetViews><sheetView workbookViewId="0"/><sheetView showGridLines="0" zoomScale="50" workbookViewId="1">` +
        `<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="C3" sqref="C3"/>` +
        `</sheetView></sheetViews><sheetData/>`,
    );
    expect(content).toEqual({ rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] });
  });

  it("reads a frozen pane as frozen ySplit rows and xSplit columns, its topLeftCell winning", () => {
    const frozen = view(`topLeftCell="A1"`, `<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>`);
    expect(frozen.frozen).toEqual({ rows: 2, cols: 1 });
    expect(frozen.topLeftCell).toEqual({ row: 2, col: 1 });
    expect(view(``, `<pane ySplit="3" topLeftCell="A40" activePane="bottomLeft" state="frozenSplit"/>`)).toMatchObject({
      frozen: { rows: 3, cols: 0 },
      topLeftCell: { row: 39, col: 0 },
    });
  });

  it("does not read a split pane (or a pane with no state) as a freeze", () => {
    for (const state of [` state="split"`, ""]) {
      const split = view(`topLeftCell="A2"`, `<pane xSplit="2010" ySplit="1500" topLeftCell="C5" activePane="bottomRight"${state}/>`);
      expect(split).not.toHaveProperty("frozen");
      expect(split.topLeftCell).toEqual({ row: 1, col: 0 }); // the sheetView's, not the pane's
    }
  });

  it("reads the active pane's selection: activeCell, a two-range sqref and activeCellId", () => {
    const content = view(
      ``,
      `<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>` +
        `<selection pane="topRight" activeCell="B1" sqref="B1"/><selection pane="bottomLeft" activeCell="A3" sqref="A3"/>` +
        `<selection pane="bottomRight" activeCell="D7" activeCellId="1" sqref="B3:B5 D7:E8"/>`,
    );
    expect(content.selection).toEqual({
      active: { row: 6, col: 3 },
      ranges: [parseRange("B3:B5"), parseRange("D7:E8")],
      activeRangeIndex: 1,
    });
  });

  it("defaults the active pane to topLeft, matching a selection with no pane", () => {
    expect(view(``, `<selection activeCell="C2" sqref="C2:D4"/>`).selection).toEqual({
      active: { row: 1, col: 2 },
      ranges: [parseRange("C2:D4")],
      activeRangeIndex: 0,
    });
    expect(view(``, `<selection pane="bottomRight" activeCell="C2" sqref="C2"/>`)).not.toHaveProperty("selection");
  });

  it("normalises a selection so the active cell lies in its active range", () => {
    // activeCellId names the wrong range: the range holding the active cell.
    expect(view(``, `<selection activeCell="E8" activeCellId="0" sqref="B3:B5 D7:E8"/>`).selection).toMatchObject({ activeRangeIndex: 1 });
    // …or out of range.
    expect(view(``, `<selection activeCell="B4" activeCellId="9" sqref="B3:B5 D7:E8"/>`).selection).toMatchObject({ activeRangeIndex: 0 });
    // The active cell in no range: the active cell alone.
    expect(view(``, `<selection activeCell="Z9" sqref="B3:B5"/>`).selection).toEqual({
      active: { row: 8, col: 25 },
      ranges: [parseRange("Z9")],
      activeRangeIndex: 0,
    });
    // No sqref: the active cell alone; no activeCell: the range's corner.
    expect(view(``, `<selection activeCell="C3"/>`).selection).toEqual({ active: { row: 2, col: 2 }, ranges: [parseRange("C3")], activeRangeIndex: 0 });
    expect(view(``, `<selection sqref="D4:E5"/>`).selection).toEqual({ active: { row: 3, col: 3 }, ranges: [parseRange("D4:E5")], activeRangeIndex: 0 });
  });

  it("reads nothing from elements deeper than the root's own children", () => {
    const content = read(
      `<sheetData/><customSheetViews><customSheetView guid="{0}" scale="50" showGridLines="0"><pane xSplit="1" ySplit="1" state="frozen"/>` +
        `<selection activeCell="C3" sqref="C3"/></customSheetView></customSheetViews>`,
    );
    expect(content).toEqual({ rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] });
  });
});

describe("rich-text runs", () => {
  it("gives a shared string's runs their rPr fonts; a plain item stays { text }", () => {
    const palette = [...INDEXED_COLORS];
    palette[44] = "#ffb6c1";
    const strings = readSharedStrings(
      `<sst xmlns="${MAIN}"><si><t>plain</t></si>` +
        `<si><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve">Styled </t></r>` +
        `<r><rPr><i/><strike/><u val="double"/><vertAlign val="superscript"/><sz val="14"/><color indexed="44"/><rFont val="Cambria"/><family val="1"/></rPr><t>workbook</t></r>` +
        `<r><t>!</t></r><r><rPr/><t>?</t></r></si></sst>`,
      palette,
    );
    expect(strings).toEqual([
      { text: "plain" },
      {
        text: "Styled workbook!?",
        runs: [
          { text: "Styled ", font: { family: { theme: "minor" }, sizePt: 11, bold: true, color: { theme: "dk1" } } },
          {
            text: "workbook",
            font: { family: "Cambria", sizePt: 14, italic: true, underline: "double", strike: true, color: "#ffb6c1", vertAlign: "superscript" },
          },
          { text: "!" },
          { text: "?" },
        ],
      },
    ]);
  });

  it("gives an inline string's runs their rPr fonts", () => {
    const [row] = rowsOf(
      `<row r="1"><c r="A1" t="inlineStr" s="2"><is><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>Red </t></r><r><t>plain</t></r></is></c></row>`,
    );
    expect(row!.cells).toEqual([
      { col: 0, value: "Red plain", runs: [{ text: "Red ", font: { bold: true, color: "#ff0000" } }, { text: "plain" }], styleId: 2 },
    ]);
  });
});

async function sheets(): Promise<Sheet[]> {
  return (await parseXlsx(await buildSpec117StylesXlsxBytes())).sheets;
}

/** An A1 address as a plain CellRef. */
function refOf(a1: string): { row: number; col: number } {
  const { row, col } = parseA1(a1)!;
  return { row, col };
}

function cellAt(sheet: Sheet, a1: string) {
  const ref = refOf(a1);
  return getCell(sheet, ref.row, ref.col);
}

describe("spec117-styles.xlsx (Excel-shaped)", () => {
  const F = SPEC117_STYLES_XLSX;

  it("reads the sheets in order, Styled active", async () => {
    const parsed = await parseXlsx(await buildSpec117StylesXlsxBytes());
    expect(parsed.sheets.map((sheet) => sheet.name)).toEqual(F.sheetNames);
    expect(parsed.activeSheetIndex).toBe(F.activeSheetIndex);
  });

  it("reads Styled's view: frozen 2 rows × 1 column at B3, 85%, the bottomRight selection and the tab colour", async () => {
    const [styled] = await sheets();
    expect(styled!.frozen).toEqual(F.frozen);
    expect(styled!.zoomPercent).toBe(F.zoomPercent);
    expect(styled!.topLeftCell).toEqual(refOf(F.topLeftCell));
    expect(styled!.selection).toEqual({ active: refOf(F.selection.active), ranges: F.selection.ranges, activeRangeIndex: F.selection.activeRangeIndex });
    expect(styled!.tabColor).toMatchObject({ theme: "accent2", lumOff: 0.39997558519241921, xlsxTint: F.tabColorTint });
    expect(styled).not.toHaveProperty("showGridLines");
  });

  it("reads Styled's sizes: Calibri 11 defaults, the custom, best-fit, hidden and styled columns, row heights", async () => {
    const [styled] = await sheets();
    expect(styled!.defaultColWidthPx).toBe(64);
    expect(styled!.defaultRowHeightPt).toBe(15);
    expect(styled!.columns).toEqual([
      { min: 0, max: 0, widthPx: F.columnWidthsPx.A, sourceWidthChars: 18.7109375 },
      { min: 1, max: 1, widthPx: F.columnWidthsPx.B, sourceWidthChars: 12.7109375 },
      { min: 5, max: 5, widthPx: F.columnWidthsPx.F, sourceWidthChars: 13.8515625, sourceAttrs: `bestFit="1"` },
      { min: F.hiddenColumn.index, max: F.hiddenColumn.index, widthPx: 64, hidden: true, sourceWidthChars: 9.140625 },
      { min: F.styledColumns.min, max: F.styledColumns.max, widthPx: 64, styleId: F.styledColumns.styleId, sourceWidthChars: 9.140625 },
    ]);
    const row = (index: number) => styled!.rows.find((candidate) => candidate.index === index)!;
    expect(row(1)).toMatchObject({ heightPt: F.rowHeightsPt[2], customHeight: true, sourceAttrs: `spans="1:26" x14ac:dyDescent="0.25"` });
    expect(row(4)).toMatchObject({ heightPt: F.rowHeightsPt[5], customHeight: true });
    expect(row(F.hiddenRow.index)).toMatchObject({ hidden: true });
    expect(row(F.styledRow.index)).toMatchObject({ styleId: F.styledRow.styleId, sourceAttrs: `spans="1:1" x14ac:dyDescent="0.25"` });
    expect(row(2)).not.toHaveProperty("heightPt");
  });

  it("reads Styled's cells: B2's style, B3:B12, the merge and its style-only members, the hidden column's value", async () => {
    const [styled] = await sheets();
    expect(cellAt(styled!, F.header.ref)).toEqual({ col: 1, value: F.header.value, styleId: F.header.styleId });
    F.columnB.values.forEach((value, i) => expect(cellAt(styled!, `B${i + 3}`)?.value).toBe(value));
    expect(cellAt(styled!, "B13")).toBeUndefined();
    for (let r = 13; r <= 18; r += 1) {
      expect(cellAt(styled!, `B${r}`), `B${r}`).toBeUndefined();
    }
    expect(styled!.merges).toEqual([F.merge.range]);
    expect(cellAt(styled!, "C5")).toEqual({ col: 2, value: F.merge.value, styleId: F.merge.styleId });
    expect(cellAt(styled!, "D5")).toEqual({ col: 3, value: null, styleId: F.merge.styleId });
    expect(cellAt(styled!, "E5")).toEqual({ col: 4, value: null, styleId: F.merge.styleId });
    expect(cellAt(styled!, F.styleOnly.ref)).toEqual({ col: 3, value: null, styleId: F.styleOnly.styleId });
    expect(cellAt(styled!, F.hiddenColumn.valueRef)?.value).toBe(F.hiddenColumn.value);
    expect(cellAt(styled!, F.hiddenRow.ref)?.value).toBe(F.hiddenRow.value);
    expect(cellAt(styled!, F.quotePrefix.ref)).toEqual({ col: 0, value: F.quotePrefix.value, styleId: F.quotePrefix.styleId });
    SPEC117_BORDER_STYLES.forEach((style, k) => {
      expect(cellAt(styled!, `${F.borderRow.columns[k]}16`), style).toMatchObject({ value: style, styleId: SPEC117_XF.borders + k });
    });
  });

  it("reads A1's rich shared string with its two run fonts", async () => {
    const [styled] = await sheets();
    expect(cellAt(styled!, F.rich.ref)).toEqual({
      col: 0,
      value: F.rich.value,
      runs: [
        { text: F.rich.runs[0], font: { family: { theme: "minor" }, sizePt: 11, bold: true, color: { theme: "dk1" } } },
        { text: F.rich.runs[1], font: { family: "Cambria", sizePt: 14, italic: true, color: "#c00000" } },
      ],
    });
  });

  it("reads Plain's hidden gridlines and Split's split pane — no freeze, the active pane's selection", async () => {
    const [, plain, split] = await sheets();
    expect(plain!.showGridLines).toBe(false);
    expect(plain!.selection).toEqual({ active: refOf(F.plain.selection), ranges: [parseRange(F.plain.selection)], activeRangeIndex: 0 });
    expect(plain!).not.toHaveProperty("frozen");
    expect(split!).not.toHaveProperty("frozen");
    expect(split!).not.toHaveProperty("topLeftCell");
    expect(split!.selection).toEqual({ active: refOf(F.split.selection), ranges: [parseRange(F.split.selection)], activeRangeIndex: 0 });
    expect(split!.merges).toEqual([]);
    expect(split!.columns).toEqual([]);
  });

  it("is plain JSON", async () => {
    const parsed = await parseXlsx(await buildSpec117StylesXlsxBytes());
    const { opaque: _opaque, ...model } = parsed;
    expect(JSON.parse(JSON.stringify(model))).toEqual(model);
  });
});
