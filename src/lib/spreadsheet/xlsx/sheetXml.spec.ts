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
//
// SPEC-122 "The worksheet writer" and "Reader upgrade" (the xlsx captures):
// the reader's sstIndex, preservedAttrs, preservedXml and sheet protection;
// serializeWorksheet's goldens per type-table row (never t="shared", arrays
// and their members, a dynamic array's cm kept and dropped, a data table's
// <f> and an unparsed formula's wire text verbatim, preserved attributes and
// children), rows (no spans or dyDescent, the other sourceAttrs kept), cols
// (the width rules), the outline of default-width columns and of rows
// without cells (stored by the reader, re-emitted through an edit), the
// splice order with every CT_Worksheet sibling present
// and with the rebuilt ones absent, the root start tag and every other element
// byte-verbatim, the first sheetView (spliced attributes, the frozen pane, the
// selections, the rest verbatim), sheetFormatPr, mergeCells, SPEC-121's
// empty-sheet template byte for byte, encodeXlsxText's round trip through
// SPEC-116's decoding (property-tested), and an Excel-shaped part rebuilt
// from an unchanged model.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../../core/equality";
import { unzip } from "../../core/zip";
import { formatA1, getCell, parseA1, parseRange } from "../address";
import { buildSpec117StylesXlsxBytes, buildSpec119FormulasXlsxBytes, SPEC117_BORDER_STYLES, SPEC117_STYLES_XLSX, SPEC117_XF, SPEC119_FORMULAS_XLSX } from "../fixtures/build";
import { baseColWidthToPx, maxDigitWidthPx, xlsxWidthToPx } from "../geometry";
import { MAX_WORKBOOK_CELLS, type Cell, type Row, type Sheet } from "../model";
import { readSpreadsheetBytes } from "../open";
import { CellBudget } from "../parsed";
import { INDEXED_COLORS } from "../styles";
import { decodeXmlText, escapeElementText } from "../xmlSplice";
import { parseXlsx } from "./reader";
import { decodeXmlEscapes, encodeXlsxText, xlsxTextElement } from "./sharedStrings";
import {
  emptyWorksheetXml,
  readSharedStrings,
  readWorksheet,
  serializeWorksheet,
  WORKSHEET_ORDER,
  xlsxWidthForPx,
  type SharedString,
  type WorksheetContent,
  type XlsxSheetContext,
} from "./sheetXml";
import { writeXlsx } from "./writer";

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

  it("keeps a span at the default width that carries an unmodelled attribute other than bestFit — an outline group the writer re-emits (SPEC-122)", () => {
    // Excel groups columns B:D without resizing them as width 9.140625 — 64 px, the default.
    const { columns } = read(
      `<cols><col min="2" max="4" width="9.140625" outlineLevel="1"/><col min="5" max="5" width="9.140625" bestFit="1" collapsed="1"/>` +
        `<col min="6" max="6" outlineLevel="2"/><col min="7" max="7" width="9.140625" bestFit="1"/></cols><sheetData/>`,
    );
    expect(columns).toEqual([
      { min: 1, max: 3, widthPx: 64, sourceWidthChars: 9.140625, sourceAttrs: `outlineLevel="1"` },
      { min: 4, max: 4, widthPx: 64, sourceWidthChars: 9.140625, sourceAttrs: `bestFit="1" collapsed="1"` },
      { min: 5, max: 5, sourceAttrs: `outlineLevel="2"` },
    ]);
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

  it("stores a row with no cells for an unmodelled attribute other than the layout caches — an outline level the writer re-emits (SPEC-122)", () => {
    expect(
      rowsOf(
        `<row r="1" spans="1:3" outlineLevel="1" x14ac:dyDescent="0.25"/><row r="2" collapsed="1"/><row r="3" spans="1:3" x14ac:dyDescent="0.25"/>` +
          // A row whose only cell stores nothing, and one whose only cell is a follower of no master.
          `<row r="4" outlineLevel="1"><c r="A4"/></row><row r="5" outlineLevel="1"><c r="A5"><f t="shared" si="9"/></c></row>`,
      ),
    ).toEqual([
      { index: 0, cells: [], sourceAttrs: `spans="1:3" outlineLevel="1" x14ac:dyDescent="0.25"` },
      { index: 1, cells: [], sourceAttrs: `collapsed="1"` },
      { index: 3, cells: [], sourceAttrs: `outlineLevel="1"` },
      { index: 4, cells: [], sourceAttrs: `outlineLevel="1"` },
    ]);
  });
});

describe("cells", () => {
  it("reads s as the cell's styleId, as written — 0 included", () => {
    // SPEC-122: the t="s" cell also keeps the item it was read from (sstIndex).
    expect(rowsOf(`<row r="1"><c r="A1" s="0"><v>1</v></c><c r="B1" s="7" t="s"><v>0</v></c><c r="C1"><v>3</v></c></row>`, [{ text: "x" }])).toEqual([
      {
        index: 0,
        cells: [
          { col: 0, value: 1, styleId: 0 },
          { col: 1, value: "x", styleId: 7, sstIndex: 0 },
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

/** spec117-styles.xlsx's `<si>` items, exactly as written, in order. */
async function spec117Items(): Promise<string[]> {
  const xml = new TextDecoder().decode((await unzip(await buildSpec117StylesXlsxBytes())).get("xl/sharedStrings.xml"));
  return xml.match(/<si>[\s\S]*?<\/si>|<si\/>/g) ?? [];
}

/** The position of an `<si>` item, as written, in spec117-styles.xlsx. */
async function spec117ItemIndex(item: string): Promise<number> {
  const index = (await spec117Items()).indexOf(item);
  expect(index, item).toBeGreaterThanOrEqual(0);
  return index;
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
    // SPEC-122: each shared-string cell also keeps the item it was read from.
    expect(cellAt(styled!, F.header.ref)).toEqual({ col: 1, value: F.header.value, styleId: F.header.styleId, sstIndex: await spec117ItemIndex(`<si><t>${F.header.value}</t></si>`) });
    F.columnB.values.forEach((value, i) => expect(cellAt(styled!, `B${i + 3}`)?.value).toBe(value));
    expect(cellAt(styled!, "B13")).toBeUndefined();
    for (let r = 13; r <= 18; r += 1) {
      expect(cellAt(styled!, `B${r}`), `B${r}`).toBeUndefined();
    }
    expect(styled!.merges).toEqual([F.merge.range]);
    expect(cellAt(styled!, "C5")).toEqual({ col: 2, value: F.merge.value, styleId: F.merge.styleId, sstIndex: await spec117ItemIndex(`<si><t>${F.merge.value}</t></si>`) });
    expect(cellAt(styled!, "D5")).toEqual({ col: 3, value: null, styleId: F.merge.styleId });
    expect(cellAt(styled!, "E5")).toEqual({ col: 4, value: null, styleId: F.merge.styleId });
    expect(cellAt(styled!, F.styleOnly.ref)).toEqual({ col: 3, value: null, styleId: F.styleOnly.styleId });
    expect(cellAt(styled!, F.hiddenColumn.valueRef)?.value).toBe(F.hiddenColumn.value);
    expect(cellAt(styled!, F.hiddenRow.ref)?.value).toBe(F.hiddenRow.value);
    expect(cellAt(styled!, F.quotePrefix.ref)).toEqual({
      col: 0,
      value: F.quotePrefix.value,
      styleId: F.quotePrefix.styleId,
      sstIndex: await spec117ItemIndex(`<si><t>${F.quotePrefix.value}</t></si>`),
    });
    SPEC117_BORDER_STYLES.forEach((style, k) => {
      expect(cellAt(styled!, `${F.borderRow.columns[k]}16`), style).toMatchObject({ value: style, styleId: SPEC117_XF.borders + k });
    });
  });

  it("reads A1's rich shared string with its two run fonts", async () => {
    const [styled] = await sheets();
    // SPEC-122: and the rich item's index, which keeps its runs on save.
    expect(cellAt(styled!, F.rich.ref)).toEqual({
      col: 0,
      value: F.rich.value,
      runs: [
        { text: F.rich.runs[0], font: { family: { theme: "minor" }, sizePt: 11, bold: true, color: { theme: "dk1" } } },
        { text: F.rich.runs[1], font: { family: "Cambria", sizePt: 14, italic: true, color: "#c00000" } },
      ],
      sstIndex: (await spec117Items()).findIndex((item) => item.includes(`<r>`)),
    });
    expect(cellAt(styled!, F.rich.ref)?.sstIndex).toBe(0);
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

// --- SPEC-122: what the reader captures for the writer ------------------------------

describe("SPEC-122 reader captures", () => {
  it("keeps a t=\"s\" cell's item index as sstIndex — rich items included — and none on other types", () => {
    const strings: SharedString[] = [{ text: "a" }, { text: "rich", runs: [{ text: "ri" }, { text: "ch", font: { bold: true } }] }];
    expect(rowsOf(`<row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v> 0 </v></c><c r="C1" t="str"><f>"a"</f><v>a</v></c><c r="D1" t="inlineStr"><is><t>a</t></is></c></row>`, strings)[0]!.cells).toEqual([
      { col: 0, value: "rich", runs: strings[1]!.runs, sstIndex: 1 },
      { col: 1, value: "a", sstIndex: 0 },
      { col: 2, value: "a", formula: '"a"' },
      { col: 3, value: "a" },
    ]);
  });

  it("keeps every <c> attribute but r, s, t and cm verbatim, with leading spaces, in preservedAttrs", () => {
    const [cell] = rowsOf(`<row r="1"><c r="A1" s="2" t="s" cm="1" vm="2" ph='1' x14ac:foo="a&amp;b"><v>0</v></c></row>`, [{ text: "x" }])[0]!.cells;
    expect(cell).toEqual({ col: 0, value: "x", styleId: 2, cm: 1, sstIndex: 0, preservedAttrs: ` vm="2" ph='1' x14ac:foo="a&amp;b"` });
    expect(rowsOf(`<row r="1"><c r="A1" s="2"><v>1</v></c></row>`)[0]!.cells[0]).not.toHaveProperty("preservedAttrs");
  });

  it("keeps every child but f, v and is verbatim, in source order, in preservedXml", () => {
    const children = `<extLst><ext uri="{X}"><x14ac:y a="1&amp;2"/></ext></extLst><x14ac:extra/><foo xmlns="urn:other"><v>not a value</v></foo>`;
    const [cell] = rowsOf(`<row r="1"><c r="A1"><f>1+1</f>${children}<v>2</v></c></row>`)[0]!.cells;
    expect(cell).toEqual({ col: 0, value: 2, formula: "1+1", preservedXml: children });
  });

  it("stores a cell holding only preserved content, counted against the ceiling like any other", () => {
    expect(rowsOf(`<row r="3"><c r="B3" vm="1"/><c r="C3"><extLst/></c><c r="D3"/></row>`)).toEqual([
      { index: 2, cells: [{ col: 1, value: null, preservedAttrs: ` vm="1"` }, { col: 2, value: null, preservedXml: `<extLst/>` }] },
    ]);
    const budget = new CellBudget();
    budget.take(MAX_WORKBOOK_CELLS - 1);
    expect(readWorksheet(worksheet(`<sheetData><row r="1"><c r="A1" vm="1"/><c r="B1" vm="1"/></row></sheetData>`), "S", { sharedStrings: [], dateSystem: "1900", budget, mdw: 7 })).toBeNull();
  });

  it('reads sheetProtection@sheet "1"/"true" as Sheet.protection, and nothing else as protection', () => {
    expect(read(`<sheetData/><sheetProtection sheet="1" objects="1" scenarios="1"/>`).protection).toEqual({ sheet: true });
    expect(read(`<sheetData/><sheetProtection algorithmName="SHA-512" sheet="true"/>`).protection).toEqual({ sheet: true });
    expect(read(`<sheetData/><sheetProtection sheet="0" objects="1"/>`)).not.toHaveProperty("protection");
    expect(read(`<sheetData/><sheetProtection objects="1"/>`)).not.toHaveProperty("protection");
    expect(read(`<sheetData/>`)).not.toHaveProperty("protection");
  });

  it('reads SPEC-123\'s format permits: formatCells/formatColumns/formatRows "0"/"false" allow formatting; absent (Excel\'s default) or "1" does not', () => {
    expect(read(`<sheetData/><sheetProtection sheet="1" formatCells="0" formatColumns="false" formatRows="0"/>`).protection).toEqual({
      sheet: true,
      allowFormatCells: true,
      allowFormatColumns: true,
      allowFormatRows: true,
    });
    expect(read(`<sheetData/><sheetProtection sheet="1" objects="1" formatCells="false"/>`).protection).toEqual({ sheet: true, allowFormatCells: true });
    expect(read(`<sheetData/><sheetProtection sheet="1" formatColumns="0" formatRows="1"/>`).protection).toEqual({ sheet: true, allowFormatColumns: true });
    expect(read(`<sheetData/><sheetProtection sheet="true" formatRows="0" formatCells="1" formatColumns="true"/>`).protection).toEqual({ sheet: true, allowFormatRows: true });
    expect(read(`<sheetData/><sheetProtection sheet="1" formatCells="" formatColumns="no"/>`).protection).toEqual({ sheet: true });
    // A permit means nothing on a sheet that is not protected.
    expect(read(`<sheetData/><sheetProtection sheet="0" formatCells="0" formatColumns="0" formatRows="0"/>`)).not.toHaveProperty("protection");
  });
});

// --- SPEC-122: the worksheet writer ----------------------------------------------------

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;

/** A grid sheet with SPEC-117's defaults and `fields`. */
function sheetOf(fields: Partial<Sheet> = {}): Sheet {
  return { id: "s", name: "S", visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...fields };
}

/** A save context with stub tables: a string cell is written with its
 * `sstIndex`, else 100 + the order its text was first asked for; a style id is
 * its own cellXfs index, so `s` shows the model id. */
function stubContext(fields: Partial<XlsxSheetContext> = {}): XlsxSheetContext {
  const asked: string[] = [];
  return {
    active: false,
    theme: undefined,
    mdw: 7,
    strings: {
      indexFor: (cell) => {
        if (cell.sstIndex !== undefined) {
          return cell.sstIndex;
        }
        const text = String(cell.value);
        return 100 + (asked.includes(text) ? asked.indexOf(text) : asked.push(text) - 1);
      },
    },
    styles: { xfFor: (id) => id },
    ...fields,
  };
}

/** What serializeWorksheet writes inside `<sheetData>` for `rows` in a part holding nothing else. */
function sheetDataOf(rows: Row[], fields: Partial<Sheet> = {}): string {
  const out = serializeWorksheet(sheetOf({ rows, ...fields }), worksheet(`<sheetData/>`), stubContext());
  return /<sheetData>(.*)<\/sheetData>/s.exec(out)?.[1] ?? "";
}

/** The `<c>` serializeWorksheet writes for one cell in row 1. */
function cellXmlOf(cell: Cell, row: Partial<Row> = {}, fields: Partial<Sheet> = {}): string {
  return /^<row\b[^>]*>(.*)<\/row>$/s.exec(sheetDataOf([{ index: 0, cells: [cell], ...row }], fields))![1]!;
}

describe("the worksheet writer: cells per the type table", () => {
  const GOLDEN: [string, Cell, string][] = [
    ["a number", { col: 0, value: 42 }, `<c r="A1"><v>42</v></c>`],
    ["a number as its shortest round-trip text", { col: 0, value: 0.1 + 0.2 }, `<c r="A1"><v>0.30000000000000004</v></c>`],
    ["-0 as 0", { col: 0, value: -0 }, `<c r="A1"><v>0</v></c>`],
    ["a tiny number", { col: 0, value: 1.234e-10 }, `<c r="A1"><v>1.234e-10</v></c>`],
    ["a string: t=\"s\" and its item index", { col: 0, value: "abc", sstIndex: 7 }, `<c r="A1" t="s"><v>7</v></c>`],
    ["a string with no item yet: the index the table gives", { col: 0, value: "new" }, `<c r="A1" t="s"><v>100</v></c>`],
    ["true", { col: 0, value: true }, `<c r="A1" t="b"><v>1</v></c>`],
    ["false", { col: 0, value: false }, `<c r="A1" t="b"><v>0</v></c>`],
    ["an error", { col: 0, value: { error: "#DIV/0!" } }, `<c r="A1" t="e"><v>#DIV/0!</v></c>`],
    ["formula + number", { col: 0, value: 3, formula: "SUM(B1:B2)" }, `<c r="A1"><f>SUM(B1:B2)</f><v>3</v></c>`],
    ["formula + boolean", { col: 0, value: true, formula: "B1>1" }, `<c r="A1" t="b"><f>B1&gt;1</f><v>1</v></c>`],
    ["formula + error", { col: 0, value: { error: "#DIV/0!" }, formula: "1/0" }, `<c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c>`],
    ["formula + string: t=\"str\", the text encoded", { col: 0, value: "a\r\n& <b>", formula: '"a"&B1' }, `<c r="A1" t="str"><f>"a"&amp;B1</f><v>a_x000D_\n&amp; &lt;b&gt;</v></c>`],
    ["formula with no value yet: <f> only", { col: 0, value: null, formula: "NOW()" }, `<c r="A1"><f>NOW()</f></c>`],
    ["the formula through toXlsxFormula", { col: 0, value: "big", formula: 'IFS(B1>1,"big",TRUE,"small")' }, `<c r="A1" t="str"><f>_xlfn.IFS(B1&gt;1,"big",TRUE,"small")</f><v>big</v></c>`],
    ["an unparsed formula's wire text, verbatim", { col: 0, value: 3, formula: "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)", formulaUnparsed: true }, `<c r="A1"><f>_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)</f><v>3</v></c>`],
    ["a data table's <f> element, verbatim", { col: 0, value: 10, dataTableXml: `<f t="dataTable" ref="A1:A3" dt2D="0" dtr="0" r1="F2"/>` }, `<c r="A1"><f t="dataTable" ref="A1:A3" dt2D="0" dtr="0" r1="F2"/><v>10</v></c>`],
    ["a styled cell", { col: 0, value: 1, styleId: 3 }, `<c r="A1" s="3"><v>1</v></c>`],
    ["s omitted for xf 0", { col: 0, value: 1, styleId: 0 }, `<c r="A1"><v>1</v></c>`],
    ["a style-only cell, self-closing", { col: 0, value: null, styleId: 2 }, `<c r="A1" s="2"/>`],
    ["r, s, t, cm, then the preserved attributes", { col: 0, value: "x", sstIndex: 4, styleId: 2, cm: 1, preservedAttrs: ` vm="2" ph="1"` }, `<c r="A1" s="2" t="s" cm="1" vm="2" ph="1"><v>4</v></c>`],
    ["f, v, then the preserved children", { col: 0, value: 2, formula: "1+1", preservedXml: `<extLst><ext uri="{X}"/></extLst>` }, `<c r="A1"><f>1+1</f><v>2</v><extLst><ext uri="{X}"/></extLst></c>`],
    ["only preserved attributes", { col: 0, value: null, preservedAttrs: ` vm="1"` }, `<c r="A1" vm="1"/>`],
    ["only a preserved child", { col: 0, value: null, preservedXml: `<extLst/>` }, `<c r="A1"><extLst/></c>`],
  ];
  for (const [label, cell, golden] of GOLDEN) {
    it(label, () => {
      expect(cellXmlOf(cell)).toBe(golden);
    });
  }

  it("writes s=\"0\" when the row or column would otherwise lend the cell its own style, and the inherited xf for a cell without one (the reader resolves cell → row → column; Excel reads a <c> without s as xf 0)", () => {
    expect(cellXmlOf({ col: 0, value: 1, styleId: 0 }, { styleId: 5 })).toBe(`<c r="A1" s="0"><v>1</v></c>`);
    expect(cellXmlOf({ col: 0, value: 1, styleId: 0 }, {}, { columns: [{ min: 0, max: 2, styleId: 4 }] })).toBe(`<c r="A1" s="0"><v>1</v></c>`);
    expect(cellXmlOf({ col: 0, value: 1 }, { styleId: 5 })).toBe(`<c r="A1" s="5"><v>1</v></c>`);
    expect(cellXmlOf({ col: 0, value: 1 }, {}, { columns: [{ min: 0, max: 2, styleId: 4 }] })).toBe(`<c r="A1" s="4"><v>1</v></c>`);
    expect(cellXmlOf({ col: 0, value: 1 })).toBe(`<c r="A1"><v>1</v></c>`);
  });

  it("never writes t=\"shared\": a shared block read from a part is written one formula per cell", () => {
    const part = worksheet(
      `<sheetData><row r="1"><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row><row r="2"><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>` +
        `<row r="3"><c r="B3"><f t="shared" si="0"/><v>6</v></c></row></sheetData>`,
    );
    const sheet = sheetOf(read(part.slice(part.indexOf(">") + 1, -"</worksheet>".length)));
    expect(/<sheetData>.*<\/sheetData>/.exec(serializeWorksheet(sheet, part, stubContext()))![0]).toBe(
      `<sheetData><row r="1"><c r="B1"><f>A1*2</f><v>2</v></c></row><row r="2"><c r="B2"><f>A2*2</f><v>4</v></c></row>` +
        `<row r="3"><c r="B3"><f>A3*2</f><v>6</v></c></row></sheetData>`,
    );
  });

  it("writes an array anchor as <f t=\"array\" ref> and its members' cached values only (a string member as t=\"str\")", () => {
    const rows: Row[] = [
      { index: 1, cells: [{ col: 4, value: 6, formula: "A2:A3*2", arrayRange: parseRange("E2:E3")! }, { col: 7, value: "a", formula: "SORT(A2:A3)", arrayRange: parseRange("H2:H3")!, dynamicArray: true, cm: 1 }] },
      { index: 2, cells: [{ col: 4, value: 2 }, { col: 7, value: "b" }] },
    ];
    expect(sheetDataOf(rows)).toBe(
      `<row r="2"><c r="E2"><f t="array" ref="E2:E3">A2:A3*2</f><v>6</v></c><c r="H2" t="str" cm="1"><f t="array" ref="H2:H3">_xlfn._xlws.SORT(A2:A3)</f><v>a</v></c></row>` +
        `<row r="3"><c r="E3"><v>2</v></c><c r="H3" t="str"><v>b</v></c></row>`,
    );
    // A one-cell array names its single cell.
    expect(cellXmlOf({ col: 2, value: "4x", formula: 'CONCAT(A1,"x")', arrayRange: parseRange("C1")! })).toBe(`<c r="C1" t="str"><f t="array" ref="C1">_xlfn.CONCAT(A1,"x")</f><v>4x</v></c>`);
  });

  it("keeps a dynamic array's cm while its formula is unchanged, and writes a legacy array once the edit's drop rules removed it", () => {
    const anchor: Cell = { col: 7, value: 1, formula: "SORT(A2:A6)", arrayRange: parseRange("H2:H6")!, dynamicArray: true, cm: 1 };
    expect(cellXmlOf(anchor)).toBe(`<c r="H1" cm="1"><f t="array" ref="H2:H6">_xlfn._xlws.SORT(A2:A6)</f><v>1</v></c>`);
    // Setting the formula removes cm and dynamicArray (SPEC-122's drop rules); arrayRange stays.
    const { cm: _cm, dynamicArray: _dynamic, ...edited } = { ...anchor, formula: "SORT(A2:A5)" };
    expect(cellXmlOf(edited)).toBe(`<c r="H1"><f t="array" ref="H2:H6">_xlfn._xlws.SORT(A2:A5)</f><v>1</v></c>`);
  });
});

describe("the worksheet writer: rows, columns, merges, dimension", () => {
  it("writes <row r [s customFormat] [ht [customHeight]] [hidden] {sourceAttrs less spans and x14ac:dyDescent}>", () => {
    expect(
      sheetDataOf([
        {
          index: 4,
          cells: [{ col: 0, value: 1 }],
          styleId: 3,
          heightPt: 20.25,
          customHeight: true,
          hidden: true,
          sourceAttrs: `spans="1:5" outlineLevel="1" collapsed="1" thickBot="1" x14ac:dyDescent="0.25"`,
        },
        { index: 5, cells: [{ col: 1, value: 2 }], heightPt: 18, sourceAttrs: `spans="2:2" x14ac:dyDescent="0.25"` },
        { index: 6, cells: [], hidden: true, sourceAttrs: `outlineLevel='2'` },
      ]),
    ).toBe(
      `<row r="5" s="3" customFormat="1" ht="20.25" customHeight="1" hidden="1" outlineLevel="1" collapsed="1" thickBot="1"><c r="A5" s="3"><v>1</v></c></row>` +
        `<row r="6" ht="18"><c r="B6"><v>2</v></c></row><row r="7" hidden="1" outlineLevel='2'/>`,
    );
  });

  it("writes one col per span: 1-based min/max, width (the source's, else the smallest 1/256 that maps to widthPx), style, hidden, customWidth, sourceAttrs", () => {
    const columns = [
      { min: 0, max: 0, widthPx: 131, sourceWidthChars: 18.7109375 },
      { min: 1, max: 2, widthPx: 100 },
      { min: 4, max: 4, hidden: true as const, styleId: 3 },
      { min: 5, max: 5, widthPx: 64, hidden: true as const, sourceWidthChars: 9.140625, sourceAttrs: `bestFit="1" outlineLevel="1" collapsed="1"` },
    ];
    const out = serializeWorksheet(sheetOf({ columns }), worksheet(`<sheetData/>`), stubContext());
    expect(out).toBe(
      worksheet(
        `<dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>` +
          `<cols><col min="1" max="1" width="18.7109375" customWidth="1"/><col min="2" max="3" width="14.21875" customWidth="1"/>` +
          `<col min="5" max="5" style="3" hidden="1"/><col min="6" max="6" width="9.140625" hidden="1" customWidth="1" bestFit="1" outlineLevel="1" collapsed="1"/></cols><sheetData/>`,
      ),
    );
    expect(xlsxWidthToPx(14.21875, 7)).toBe(100);
    expect(xlsxWidthToPx(14.21875 - 1 / 256, 7)).toBe(99);
    // Read back: the same spans, a computed width kept as the source's.
    expect(read(out.slice(out.indexOf(">") + 1, -"</worksheet>".length)).columns).toEqual([
      columns[0],
      { min: 1, max: 2, widthPx: 100, sourceWidthChars: 14.21875 },
      columns[2],
      columns[3],
    ]);
  });

  it("keeps the outline of grouped default-width columns and grouped rows without cells through an edit, and reads it back", () => {
    const body =
      `<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1" outlineLevelCol="1"/>` +
      `<cols><col min="2" max="4" width="9.140625" outlineLevel="1"/><col min="5" max="5" width="9.140625" collapsed="1"/><col min="6" max="6" outlineLevel="1"/></cols>` +
      `<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2" spans="1:1" outlineLevel="1" x14ac:dyDescent="0.25"/>` +
      `<row r="3" outlineLevel="1"/><row r="4" collapsed="1"/></sheetData>`;
    const read1 = sheetOf(read(body));
    const edited: Sheet = { ...read1, rows: [{ index: 0, cells: [{ col: 0, value: 2 }] }, ...read1.rows.slice(1)] };
    const out = serializeWorksheet(edited, worksheet(body), stubContext());
    expect(out).toContain(
      `<cols><col min="2" max="4" width="9.140625" customWidth="1" outlineLevel="1"/><col min="5" max="5" width="9.140625" customWidth="1" collapsed="1"/>` +
        `<col min="6" max="6" outlineLevel="1"/></cols>`,
    );
    expect(out).toContain(
      `<sheetData><row r="1"><c r="A1"><v>2</v></c></row><row r="2" outlineLevel="1"/><row r="3" outlineLevel="1"/><row r="4" collapsed="1"/></sheetData>`,
    );
    expect(out).toContain(`<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1" outlineLevelCol="1"/>`);
    const read2 = read(out.slice(out.indexOf(">") + 1, -"</worksheet>".length));
    expect(read2.columns).toEqual(read1.columns);
    expect(read2.rows).toEqual([edited.rows[0], { index: 1, cells: [], sourceAttrs: `outlineLevel="1"` }, ...edited.rows.slice(2)]);
  });

  it("writes the spans of a part's several <cols> (schema-legal) into the first, removing the rest", () => {
    const body = `<cols><col min="1" max="1" width="20" customWidth="1"/></cols><cols><col min="3" max="3" width="30" customWidth="1"/></cols><sheetData/>`;
    const sheet = sheetOf(read(body));
    expect(sheet.columns.map((span) => [span.min, span.max])).toEqual([
      [0, 0],
      [2, 2],
    ]);
    const out = serializeWorksheet(sheet, worksheet(body), stubContext());
    expect(out.match(/<cols>/g)).toHaveLength(1);
    expect(out).toContain(`<cols><col min="1" max="1" width="20" customWidth="1"/><col min="3" max="3" width="30" customWidth="1"/></cols><sheetData/>`);
    expect(read(out.slice(out.indexOf(">") + 1, -"</worksheet>".length)).columns).toEqual(sheet.columns);
  });

  it("xlsxWidthForPx is the smallest multiple of 1/256 SPEC-117's formula maps to the px, for every MDW", () => {
    for (const mdw of [5, 6, 7, 8, 9, 11, 13]) {
      for (let px = 1; px <= 400; px += 1) {
        const width = xlsxWidthForPx(px, mdw);
        expect(Number.isInteger(width * 256), `${px}@${mdw}`).toBe(true);
        expect(xlsxWidthToPx(width, mdw), `${px}@${mdw}`).toBe(px);
        if (width > 0) {
          expect(xlsxWidthToPx(width - 1 / 256, mdw), `${px}@${mdw}`).toBeLessThan(px);
        }
      }
    }
  });

  it("removes cols and mergeCells when the model has none; writes count + one mergeCell per merge", () => {
    const base = worksheet(`<dimension ref="A1"/><cols><col min="1" max="1" width="20" customWidth="1"/></cols><sheetData/><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><pageMargins left="1" right="1" top="1" bottom="1" header="0" footer="0"/>`);
    expect(serializeWorksheet(sheetOf(), base, stubContext())).toBe(
      worksheet(`<dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="1" right="1" top="1" bottom="1" header="0" footer="0"/>`),
    );
    const merged = serializeWorksheet(sheetOf({ merges: [parseRange("C5:E5")!, parseRange("A8:B10")!] }), base, stubContext());
    expect(merged).toContain(`<sheetData/><mergeCells count="2"><mergeCell ref="C5:E5"/><mergeCell ref="A8:B10"/></mergeCells><pageMargins `);
  });

  it("writes dimension as the used range — style-only cells included — and \"A1\" when the sheet stores no cell", () => {
    const base = worksheet(`<dimension ref="A1:Z99"/><sheetData/>`);
    const dimension = (sheet: Sheet): string => /<dimension ref="([^"]*)"\/>/.exec(serializeWorksheet(sheet, base, stubContext()))![1]!;
    expect(dimension(sheetOf({ rows: [{ index: 1, cells: [{ col: 3, value: 1 }] }, { index: 6, cells: [{ col: 1, value: null, styleId: 2 }] }] }))).toBe("B2:D7");
    expect(dimension(sheetOf({ rows: [{ index: 4, cells: [{ col: 2, value: 1 }] }] }))).toBe("C5");
    expect(dimension(sheetOf({ rows: [{ index: 4, cells: [], hidden: true }] }))).toBe("A1");
  });
});

/** Every CT_Worksheet child, in schema order, as a producer might write it. */
const SIBLINGS: [string, string][] = [
  ["sheetPr", `<sheetPr codeName="Sheet1"><outlinePr summaryBelow="0"/></sheetPr>`],
  ["dimension", `<dimension ref="A1:Z99"/>`],
  ["sheetViews", `<sheetViews><sheetView workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>`],
  ["sheetFormatPr", `<sheetFormatPr baseColWidth="10" defaultRowHeight="15" outlineLevelRow="1" x14ac:dyDescent="0.25"/>`],
  ["cols", `<cols><col min="9" max="9" width="30" customWidth="1"/></cols>`],
  ["sheetData", `<sheetData><row r="9" spans="9:9" x14ac:dyDescent="0.25"><c r="I9"><v>9</v></c></row></sheetData>`],
  ["sheetCalcPr", `<sheetCalcPr fullCalcOnLoad="1"/>`],
  ["sheetProtection", `<sheetProtection sheet="1" objects="1" scenarios="1"/>`],
  ["protectedRanges", `<protectedRanges><protectedRange sqref="A1" name="R"/></protectedRanges>`],
  ["scenarios", `<scenarios><scenario name="S" count="1"><inputCells r="A1" val="1"/></scenario></scenarios>`],
  ["autoFilter", `<autoFilter ref="A1:B2"/>`],
  ["sortState", `<sortState ref="A2:B2"><sortCondition ref="A2"/></sortState>`],
  ["dataConsolidate", `<dataConsolidate function="sum"/>`],
  ["customSheetViews", `<customSheetViews><customSheetView guid="{0}"><pane xSplit="1" state="frozen"/><selection sqref="A1"/></customSheetView></customSheetViews>`],
  ["mergeCells", `<mergeCells count="1"><mergeCell ref="X1:Y1"/></mergeCells>`],
  ["phoneticPr", `<phoneticPr fontId="1" type="noConversion"/>`],
  ["conditionalFormatting", `<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1&gt;1</formula></cfRule></conditionalFormatting>`],
  ["dataValidations", `<dataValidations count="1"><dataValidation type="whole" sqref="A1"><formula1>1</formula1></dataValidation></dataValidations>`],
  ["hyperlinks", `<hyperlinks><hyperlink ref="A1" location="S!B2" display="go"/></hyperlinks>`],
  ["printOptions", `<printOptions gridLines="1"/>`],
  ["pageMargins", `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`],
  ["pageSetup", `<pageSetup orientation="landscape" r:id="rId8"/>`],
  ["headerFooter", `<headerFooter><oddHeader>&amp;CTitle</oddHeader></headerFooter>`],
  ["rowBreaks", `<rowBreaks count="1" manualBreakCount="1"><brk id="5" max="16383" man="1"/></rowBreaks>`],
  ["colBreaks", `<colBreaks count="1" manualBreakCount="1"><brk id="3" max="1048575" man="1"/></colBreaks>`],
  ["customProperties", `<customProperties><customPr name="p" r:id="rId9"/></customProperties>`],
  ["cellWatches", `<cellWatches><cellWatch r="A1"/></cellWatches>`],
  ["ignoredErrors", `<ignoredErrors><ignoredError sqref="A1" numberStoredAsText="1"/></ignoredErrors>`],
  ["smartTags", `<smartTags><cellSmartTags r="A1"><cellSmartTag type="0"/></cellSmartTags></smartTags>`],
  ["drawing", `<drawing r:id="rId1"/>`],
  ["legacyDrawing", `<legacyDrawing r:id="rId2"/>`],
  ["legacyDrawingHF", `<legacyDrawingHF r:id="rId3"/>`],
  ["picture", `<picture r:id="rId4"/>`],
  ["oleObjects", `<oleObjects><oleObject progId="Package" shapeId="1025" r:id="rId5"/></oleObjects>`],
  ["controls", `<mc:AlternateContent><mc:Choice Requires="x14"><controls><control shapeId="1026" r:id="rId6" name="B"/></controls></mc:Choice><mc:Fallback><controls><control shapeId="1026" r:id="rId6" name="B"/></controls></mc:Fallback></mc:AlternateContent>`],
  ["webPublishItems", `<webPublishItems count="1"><webPublishItem id="1" divId="d" sourceType="sheet" destinationFile="f.htm"/></webPublishItems>`],
  ["tableParts", `<tableParts count="1"><tablePart r:id="rId7"/></tableParts>`],
  ["extLst", `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings/></ext></extLst>`],
];

/** Excel's root start tag: namespaces, mc:Ignorable and xr:uid. */
const EXCEL_ROOT =
  `<worksheet xmlns="${MAIN}" xmlns:r="${R_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="x14ac xr" xmlns:x14ac="${X14AC}" ` +
  `xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xr:uid="{00000000-0001-0000-0000-000000000000}">`;

/** A part holding the listed siblings, in order. */
const partWith = (siblings: readonly [string, string][]): string => `${DECLARATION}\r\n${EXCEL_ROOT}${siblings.map(([, xml]) => xml).join("")}</worksheet>`;

describe("the worksheet writer: the splice", () => {
  /** The model the splice tests write: two cells, a column, a merge, and
   * (SPEC-128) an autofilter and a worksheet sort record the source's do
   * not say. */
  const model = (): Sheet =>
    sheetOf({
      rows: [{ index: 1, cells: [{ col: 1, value: 2 }, { col: 3, value: "d", sstIndex: 0 }], sourceAttrs: `spans="2:4" x14ac:dyDescent="0.25"` }],
      columns: [{ min: 1, max: 1, widthPx: 100 }],
      merges: [parseRange("B2:C2")!],
      selection: { active: { row: 1, col: 1 }, ranges: [parseRange("B2")!], activeRangeIndex: 0 },
      autoFilter: { range: parseRange("A1:D2")!, columns: [{ colId: 1, criteria: { kind: "values", values: ["2"], blanks: false, dateGroups: [] } }] },
      sortState: { ref: parseRange("B2:D2")!, keys: [{ col: 3, descending: true }] },
    });
  const REBUILT: Record<string, string> = {
    dimension: `<dimension ref="B2:D2"/>`,
    sheetViews: `<sheetViews><sheetView workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>`,
    sheetFormatPr: `<sheetFormatPr baseColWidth="10" defaultRowHeight="15" outlineLevelRow="1" x14ac:dyDescent="0.25"/>`,
    cols: `<cols><col min="2" max="2" width="14.21875" customWidth="1"/></cols>`,
    sheetData: `<sheetData><row r="2"><c r="B2"><v>2</v></c><c r="D2" t="s"><v>0</v></c></row></sheetData>`,
    autoFilter: `<autoFilter ref="A1:D2"><filterColumn colId="1"><filters><filter val="2"/></filters></filterColumn></autoFilter>`,
    sortState: `<sortState ref="B2:D2"><sortCondition descending="1" ref="D2"/></sortState>`,
    mergeCells: `<mergeCells count="1"><mergeCell ref="B2:C2"/></mergeCells>`,
  };

  it("lists CT_Worksheet's children in schema order", () => {
    expect(WORKSHEET_ORDER).toEqual(SIBLINGS.map(([local]) => local));
  });

  it("replaces the eight rebuilt elements (SPEC-128's autoFilter and sortState among them) in place, with every CT_Worksheet sibling present; the declaration, root start tag and every other element stay byte-verbatim", () => {
    const out = serializeWorksheet(model(), partWith(SIBLINGS), stubContext());
    expect(out).toBe(partWith(SIBLINGS.map(([local, xml]) => [local, REBUILT[local] ?? xml])));
  });

  it("inserts an absent rebuilt element before the first source child that comes later in CT_Worksheet order", () => {
    const absent = new Set(["dimension", "sheetViews", "cols", "mergeCells"]);
    const out = serializeWorksheet(model(), partWith(SIBLINGS.filter(([local]) => !absent.has(local))), stubContext());
    expect(out).toBe(partWith(SIBLINGS.map(([local, xml]) => [local, REBUILT[local] ?? xml])));
    // With most siblings absent: an mc:AlternateContent ranks as the element it
    // wraps, a foreign element ranks nowhere, and nothing later puts it last.
    const sparse: [string, string][] = [SIBLINGS[5]!, ["foreign", `<x14ac:unknown/>`], SIBLINGS.find(([local]) => local === "controls")!];
    expect(serializeWorksheet(model(), partWith(sparse), stubContext())).toBe(
      partWith([
        ["dimension", REBUILT.dimension!],
        ["sheetViews", REBUILT.sheetViews!],
        ["sheetFormatPr", `<sheetFormatPr defaultRowHeight="15"/>`],
        ["cols", REBUILT.cols!],
        ["sheetData", REBUILT.sheetData!],
        ["foreign", `<x14ac:unknown/>`],
        ["autoFilter", REBUILT.autoFilter!],
        ["sortState", REBUILT.sortState!],
        ["mergeCells", REBUILT.mergeCells!],
        sparse[2]!,
      ]),
    );
    expect(serializeWorksheet(model(), partWith([SIBLINGS[5]!]), stubContext())).toBe(
      partWith([
        ["dimension", REBUILT.dimension!],
        ["sheetViews", REBUILT.sheetViews!],
        ["sheetFormatPr", `<sheetFormatPr defaultRowHeight="15"/>`],
        ["cols", REBUILT.cols!],
        ["sheetData", REBUILT.sheetData!],
        ["autoFilter", REBUILT.autoFilter!],
        ["sortState", REBUILT.sortState!],
        ["mergeCells", REBUILT.mergeCells!],
      ]),
    );
  });

  it("writes a prefixed part's rebuilt elements with its prefix", () => {
    const part = `<x:worksheet xmlns:x="${MAIN}"><x:dimension ref="A1"/><x:sheetData/></x:worksheet>`;
    expect(serializeWorksheet(sheetOf({ rows: [{ index: 0, cells: [{ col: 0, value: true }] }] }), part, stubContext())).toBe(
      `<x:worksheet xmlns:x="${MAIN}"><x:dimension ref="A1"/><x:sheetViews><x:sheetView workbookViewId="0"/></x:sheetViews>` +
        `<x:sheetFormatPr defaultRowHeight="15"/><x:sheetData><x:row r="1"><x:c r="A1" t="b"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>`,
    );
  });

  it("reads back what it wrote, and writes it again byte for byte", () => {
    const first = serializeWorksheet(model(), partWith(SIBLINGS), stubContext());
    const body = first.slice(first.indexOf(EXCEL_ROOT) + EXCEL_ROOT.length, -"</worksheet>".length);
    const reread = sheetOf(read(body, { sharedStrings: [{ text: "d" }] }));
    expect(reread.rows).toEqual([{ index: 1, cells: [{ col: 1, value: 2 }, { col: 3, value: "d", sstIndex: 0 }] }]);
    expect(serializeWorksheet(reread, first, stubContext())).toBe(first);
  });
});

describe("the worksheet writer: the first sheetView and sheetFormatPr", () => {
  const VIEWS =
    `<sheetViews><sheetView tabSelected="1" showGridLines="0" zoomScale="85" zoomScaleNormal="85" view="pageLayout" workbookViewId="0">` +
    `<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="B1" sqref="B1"/>` +
    `<selection pane="bottomLeft" activeCell="A3" sqref="A3"/><selection pane="bottomRight" activeCell="C4" sqref="C4"/>` +
    `<pivotSelection pane="bottomRight" showHeader="1" activeRow="2"><pivotArea type="data"/></pivotSelection><extLst><ext uri="{X}"/></extLst></sheetView>` +
    `<sheetView workbookViewId="1" zoomScale="50"/></sheetViews>`;
  const base = worksheet(`${VIEWS}<sheetData/>`);
  const views = (sheet: Sheet, active: boolean, part = base): string => /<sheetViews>.*<\/sheetViews>/.exec(serializeWorksheet(sheet, part, stubContext({ active })))![0];

  it("leaves an unchanged view byte-identical — the attributes that agree as written, the pane and selections regenerated to the same bytes", () => {
    const sheet = sheetOf(read(`${VIEWS}<sheetData/>`));
    expect(sheet).toMatchObject({ frozen: { rows: 2, cols: 1 }, showGridLines: false, zoomPercent: 85, topLeftCell: { row: 2, col: 1 } });
    expect(views(sheet, true)).toBe(VIEWS);
  });

  it("splices tabSelected, showGridLines, showFormulas, zoomScale and topLeftCell from the model; every other attribute, pivotSelection, extLst and further sheetViews verbatim", () => {
    const sheet = sheetOf({ ...read(`${VIEWS}<sheetData/>`), showFormulas: true, zoomPercent: undefined, showGridLines: undefined, frozen: undefined, topLeftCell: { row: 39, col: 2 }, selection: undefined });
    expect(views(sheet, false)).toBe(
      `<sheetViews><sheetView zoomScaleNormal="85" view="pageLayout" workbookViewId="0" showFormulas="1" topLeftCell="C40">` +
        `<pivotSelection pane="bottomRight" showHeader="1" activeRow="2"><pivotArea type="data"/></pivotSelection><extLst><ext uri="{X}"/></extLst></sheetView>` +
        `<sheetView workbookViewId="1" zoomScale="50"/></sheetViews>`,
    );
  });

  it("regenerates the frozen pane (xSplit/ySplit/topLeftCell/activePane/state=\"frozen\") and the active pane's selection", () => {
    const sheet = sheetOf({
      ...read(`${VIEWS}<sheetData/>`),
      frozen: { rows: 1, cols: 0 },
      topLeftCell: { row: 1, col: 0 },
      selection: { active: { row: 6, col: 3 }, ranges: [parseRange("B3:B5")!, parseRange("D7:E8")!], activeRangeIndex: 1 },
    });
    expect(views(sheet, true)).toBe(
      `<sheetViews><sheetView tabSelected="1" showGridLines="0" zoomScale="85" zoomScaleNormal="85" view="pageLayout" workbookViewId="0">` +
        `<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="D7" activeCellId="1" sqref="B3:B5 D7:E8"/>` +
        `<pivotSelection pane="bottomRight" showHeader="1" activeRow="2"><pivotArea type="data"/></pivotSelection><extLst><ext uri="{X}"/></extLst></sheetView>` +
        `<sheetView workbookViewId="1" zoomScale="50"/></sheetViews>`,
    );
    // A new freeze on a view with none: the pane first, then the selections.
    const plain = worksheet(`<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/>`);
    expect(views(sheetOf({ frozen: { rows: 2, cols: 1 }, selection: { active: { row: 4, col: 4 }, ranges: [parseRange("E5")!], activeRangeIndex: 0 } }), false, plain)).toBe(
      `<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>` +
        `<selection pane="bottomRight" activeCell="E5" sqref="E5"/></sheetView></sheetViews>`,
    );
  });

  it("keeps a split pane (not modelled) and its other panes' selections, regenerating the active one", () => {
    const split = worksheet(
      `<sheetViews><sheetView workbookViewId="0" topLeftCell="A2"><pane xSplit="2010" ySplit="1500" topLeftCell="C5" activePane="bottomRight"/>` +
        `<selection pane="topRight" activeCell="D1" sqref="D1"/><selection pane="bottomRight" activeCell="C5" sqref="C5"/></sheetView></sheetViews><sheetData/>`,
    );
    const sheet = sheetOf({ ...read(split.slice(split.indexOf(">") + 1, -"</worksheet>".length)), selection: { active: { row: 7, col: 3 }, ranges: [parseRange("D8")!], activeRangeIndex: 0 } });
    expect(views(sheet, false, split)).toBe(
      `<sheetViews><sheetView workbookViewId="0" topLeftCell="A2"><pane xSplit="2010" ySplit="1500" topLeftCell="C5" activePane="bottomRight"/>` +
        `<selection pane="topRight" activeCell="D1" sqref="D1"/><selection pane="bottomRight" activeCell="D8" sqref="D8"/></sheetView></sheetViews>`,
    );
  });

  it("puts a generated view first in a sheetViews that holds none, keeping the rest", () => {
    const odd = worksheet(`<sheetViews><extLst><ext uri="{Y}"/></extLst></sheetViews><sheetData/>`);
    expect(views(sheetOf({ showFormulas: true }), true, odd)).toBe(`<sheetViews><sheetView workbookViewId="0" tabSelected="1" showFormulas="1"/><extLst><ext uri="{Y}"/></extLst></sheetViews>`);
  });

  it("splices sheetFormatPr's defaultRowHeight and customHeight only when the default changed; baseColWidth, defaultColWidth, dyDescent and outline levels verbatim", () => {
    const formatPr = `<sheetFormatPr baseColWidth="10" defaultColWidth="12" defaultRowHeight="14.5" outlineLevelRow="2" x14ac:dyDescent="0.35"/>`;
    const part = worksheet(`${formatPr}<sheetData/>`);
    const written = (sheet: Sheet): string => /<sheetFormatPr[^>]*\/>/.exec(serializeWorksheet(sheet, part, stubContext()))![0];
    expect(written(sheetOf({ defaultRowHeightPt: 14.5 }))).toBe(formatPr);
    expect(written(sheetOf({ defaultRowHeightPt: 20 }))).toBe(`<sheetFormatPr baseColWidth="10" defaultColWidth="12" defaultRowHeight="20" outlineLevelRow="2" x14ac:dyDescent="0.35" customHeight="1"/>`);
  });
});

describe("the worksheet writer: a new sheet", () => {
  it("reproduces SPEC-121's empty-sheet template byte for byte for an empty sheet", () => {
    const golden =
      `${DECLARATION}<worksheet xmlns="${MAIN}" xmlns:r="${R_NS}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
      `<sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
    expect(serializeWorksheet(sheetOf(), undefined, stubContext())).toBe(golden);
    const cases: [Sheet, boolean][] = [
      [sheetOf(), true],
      [sheetOf({ tabColor: { theme: "accent2", lumMod: 0.6, lumOff: 0.4 } }), false],
      [sheetOf({ defaultColWidthPx: 100, defaultRowHeightPt: 20.25 }), true],
    ];
    for (const [sheet, active] of cases) {
      expect(serializeWorksheet(sheet, undefined, stubContext({ active }))).toBe(emptyWorksheetXml(sheet, active, { theme: undefined }));
    }
  });

  it("splices a new sheet's content into that template", () => {
    const sheet = sheetOf({ rows: [{ index: 0, cells: [{ col: 0, value: "x" }, { col: 1, value: 1, styleId: 2 }] }], merges: [parseRange("A1:B1")!] });
    expect(serializeWorksheet(sheet, undefined, stubContext({ active: true }))).toBe(
      `${DECLARATION}<worksheet xmlns="${MAIN}" xmlns:r="${R_NS}"><dimension ref="A1:B1"/><sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>` +
        `<sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="s"><v>100</v></c><c r="B1" s="2"><v>1</v></c></row></sheetData>` +
        `<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
    );
  });
});

/** A seeded generator (mulberry32): the property tests are reproducible. */
function seeded(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** Whether every character of a text is one XML 1.0 can carry (no CR either). */
function xmlCarries(text: string): boolean {
  for (const char of text) {
    const code = char.codePointAt(0)!;
    const legal = code === 0x09 || code === 0x0a || (code >= 0x20 && code <= 0xd7ff) || (code >= 0xe000 && code <= 0xfffd) || code >= 0x10000;
    if (!legal) {
      return false;
    }
  }
  return true;
}

describe("encodeXlsxText", () => {
  it("maps XML-illegal characters and CR to _xHHHH_, escapes literal _xHHHH_ sequences, then XML-escapes", () => {
    const cases: [string, string][] = [
      ["plain text", "plain text"],
      ["tab\tand\nline feed", "tab\tand\nline feed"],
      ["a\rb", "a_x000D_b"],
      ["\u0000\u0001\u001f", "_x0000__x0001__x001F_"],
      ["_x0041_", "_x005F_x0041_"],
      ["_x0041", "_x0041"],
      ["_x0041\r", "_x005F_x0041_x000D_"],
      ["_x00e9_ and _x12_ and _xZZZZ_", "_x005F_x00e9_ and _x12_ and _xZZZZ_"],
      ["\uD800", "_xD800_"],
      ["\uDC00x", "_xDC00_x"],
      ["😀 astral", "😀 astral"],
      ["￾￿", "_xFFFE__xFFFF_"],
      [`a & <b> "c" 'd'`, `a &amp; &lt;b&gt; "c" 'd'`],
    ];
    for (const [text, encoded] of cases) {
      expect(encodeXlsxText(text), JSON.stringify(text)).toBe(encoded);
    }
  });

  it("is the inverse of SPEC-116's decoding: generated strings come back exactly, through the scanner, carried by XML 1.0", () => {
    const alphabet = [
      "a", "Z", "0", "9", "F", "f", " ", "\t", "\n", "\r", "\u0000", "\u0001", "\u0008", "\u000b", "\u001f", "_", "x", "X", "_x", "_x0", "_x004", "_x0041", "_x0041_",
      "_x005F_", "_x000D_", "&", "<", ">", "\"", "'", "é", "東", "\uD800", "\uDBFF", "\uDC00", "\uDFFF", "😀", "￾", "￿", "�",
    ];
    const random = seeded(122);
    for (let n = 0; n < 2000; n += 1) {
      let text = "";
      const length = Math.floor(random() * 12);
      for (let i = 0; i < length; i += 1) {
        text += alphabet[Math.floor(random() * alphabet.length)];
      }
      const encoded = encodeXlsxText(text);
      expect(xmlCarries(decodeXmlText(encoded)), JSON.stringify(text)).toBe(true);
      expect(decodeXmlEscapes(decodeXmlText(encoded)), JSON.stringify(text)).toBe(text);
      // …and through the reader, as a shared-string item.
      expect(readSharedStrings(`<sst xmlns="${MAIN}"><si>${xlsxTextElement(text)}</si></sst>`)[0]!.text, JSON.stringify(text)).toBe(text);
    }
  });

  it("gives <t> xml:space=\"preserve\" for leading or trailing whitespace or a line break", () => {
    expect(xlsxTextElement("a b")).toBe(`<t>a b</t>`);
    expect(xlsxTextElement("a\tb")).toBe(`<t>a\tb</t>`);
    expect(xlsxTextElement(" a")).toBe(`<t xml:space="preserve"> a</t>`);
    expect(xlsxTextElement("a\t")).toBe(`<t xml:space="preserve">a\t</t>`);
    expect(xlsxTextElement("a\nb")).toBe(`<t xml:space="preserve">a\nb</t>`);
    expect(xlsxTextElement("a\rb")).toBe(`<t xml:space="preserve">a_x000D_b</t>`);
    expect(xlsxTextElement("x", "x:")).toBe(`<x:t>x</x:t>`);
  });
});

/** spec119-formulas.xlsx written with Calc (edited by `edit`) marked dirty:
 * Calc's source text, the written parts and the source parts. */
async function rebuiltCalc(edit?: (sheet: Sheet) => Sheet): Promise<{ source: string; out: Map<string, Uint8Array>; original: Map<string, Uint8Array> }> {
  const bytes = await buildSpec119FormulasXlsxBytes();
  const model = await readSpreadsheetBytes(bytes, "xlsx", "formulas");
  const calc = model.sheets[0]!;
  model.sheets[0] = { ...(edit ? edit(calc) : calc), dirty: true };
  const original = await unzip(bytes);
  return { source: new TextDecoder().decode(original.get("xl/worksheets/sheet1.xml")), out: await unzip(await writeXlsx(model)), original };
}

describe("an Excel-shaped part rebuilt (spec119-formulas.xlsx, Calc marked dirty)", () => {
  it("differs from its source only in the rows' layout caches and the shared formulas, each cell now carrying its own text", async () => {
    const { source, out, original } = await rebuiltCalc();
    const formulas: Record<string, string> = SPEC119_FORMULAS_XLSX.formulas;
    const expected = source
      .replace(/<row\b[^>]*>/g, (tag) => tag.replace(/ spans="[^"]*"/, "").replace(/ x14ac:dyDescent="[^"]*"/, ""))
      .replace(/<c r="([A-Z]+[0-9]+)"([^>]*)><f t="shared"[^>]*?(?:\/>|>[^<]*<\/f>)/g, (_, ref: string, attrs: string) => `<c r="${ref}"${attrs}><f>${escapeElementText(formulas[ref]!)}</f>`);
    const written = new TextDecoder().decode(out.get("xl/worksheets/sheet1.xml"));
    expect(written).toBe(expected);
    expect(written).not.toContain(`t="shared"`);
    // The array, the dynamic array with its cm, the data table and the unparsed LAMBDA, as Excel wrote them.
    expect(written).toContain(`<c r="E2"><f t="array" ref="E2:E4">A2:A4*2</f><v>6</v></c>`);
    expect(written).toContain(`<c r="H2" cm="1"><f t="array" ref="H2:H6">_xlfn._xlws.SORT(A2:A6)</f><v>1</v></c>`);
    expect(written).toContain(`<c r="L3">${SPEC119_FORMULAS_XLSX.dataTable.L3}<v>10</v></c>`);
    expect(written).toContain(`<c r="S2"><f>${SPEC119_FORMULAS_XLSX.unparsed.S2}</f><v>3</v></c>`);
    for (const name of ["xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml", "xl/metadata.xml", "xl/sharedStrings.xml", "xl/styles.xml"]) {
      expect(equalBytes(out.get(name)!, original.get(name)!), name).toBe(true);
    }
  });

  it("writes the dynamic array a legacy array once its formula is edited (the drop rules removed cm); xl/metadata.xml is kept", async () => {
    const { out, original } = await rebuiltCalc((sheet) => ({
      ...sheet,
      rows: sheet.rows.map((row) =>
        row.index !== 1
          ? row
          : {
              ...row,
              cells: row.cells.map((cell) => {
                if (formatA1({ row: 1, col: cell.col }) !== "H2") {
                  return cell;
                }
                const { cm: _cm, dynamicArray: _dynamic, ...rest } = cell;
                return { ...rest, formula: "SORT(A2:A5)" };
              }),
            },
      ),
    }));
    const written = new TextDecoder().decode(out.get("xl/worksheets/sheet1.xml"));
    expect(written).toContain(`<c r="H2"><f t="array" ref="H2:H6">_xlfn._xlws.SORT(A2:A5)</f><v>1</v></c>`);
    expect(equalBytes(out.get("xl/metadata.xml")!, original.get("xl/metadata.xml")!)).toBe(true);
  });
});
