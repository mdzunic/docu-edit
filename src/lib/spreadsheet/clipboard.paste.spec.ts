// SPEC-125 "Paste geometry", "Paste semantics", "TSV in": where a block
// lands (one cell, tiling, once at the top-left, whole rows and columns) and
// every refusal; what each paste-special option replaces, Skip blanks and
// Transpose; same- and cross-workbook styles and formulas; the 1900/1904
// date-system conversion; plain text through parseInput; copy-on-write.
import { describe, expect, it } from "vitest";
import type { GridRange } from "./address";
import { getCell } from "./address";
import { PROTECTED_CELL_MESSAGE } from "./cells";
import {
  buildClipboardCopy,
  formulasAsValuesMessage,
  MULTIPLE_SELECTIONS_MESSAGE,
  PASTE_MERGE_MESSAGE,
  PASTE_PAST_EDGE_MESSAGE,
  TOO_MANY_TO_CHANGE_MESSAGE,
} from "./clipboard";
import { importClipboardHtml } from "./clipboardImport";
import { MAX_COLS, MAX_ROWS, type Cell, type CellStyle, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { blockFromPayload, blockFromText, pasteBlock, pasteCapabilities, planPaste, transposeFormula, type PasteBlock, type PasteOptions, type PasteTarget } from "./paste";
import { PARTIAL_ARRAY_MESSAGE } from "./structure";
import { cellStyle } from "./styles";

const EN_US = SPREADSHEET_LOCALES["en-US"];

function sheet(id: string, name: string, fields: Partial<Sheet> = {}): Sheet {
  return { id, name, visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...fields };
}

function model(sheets: Sheet[], fields: Partial<SpreadsheetModel> = {}): SpreadsheetModel {
  return { kind: "spreadsheet", id: "book", title: "T", format: "xlsx", readerVersion: 8, sheets, activeSheetIndex: 0, dateSystem: "1900", opaque: [], styles: [{}], defaultStyleId: 0, ...fields };
}

const value = (col: number, v: Cell["value"], extra: Partial<Cell> = {}): Cell => ({ col, value: v, ...extra });
const row = (index: number, cells: Cell[], extra: Partial<Row> = {}): Row => ({ index, cells, ...extra });
const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });
const cellRange = (r: number, c: number): GridRange => range(r, c, r, c);
const at = (r: number, c: number): PasteTarget => ({ ranges: [cellRange(r, c)], active: { row: r, col: c } });
const over = (selection: GridRange): PasteTarget => ({ ranges: [selection], active: { row: selection.top, col: selection.left } });
const ALL: PasteOptions = { what: "all" };

function cellOf(m: SpreadsheetModel, r: number, c: number, sheetId = "s1"): Cell | undefined {
  return getCell(m.sheets.find((each) => each.id === sheetId)!, r, c);
}

/** Sheet1 B2:C3 = 1, 2 / =B2*10 (10), =$B$2+C2 (3); B2 bold; a source
 * D1:E1 merge; and an empty Sheet2. */
function source(): SpreadsheetModel {
  return model(
    [
      sheet("s1", "Sheet1", {
        rows: [row(1, [value(1, 1, { styleId: 1 }), value(2, 2)]), row(2, [value(1, 10, { formula: "B2*10" }), value(2, 3, { formula: "$B$2+C2" })])],
      }),
      sheet("s2", "Sheet2"),
    ],
    { styles: [{}, { font: { bold: true } }] },
  );
}

/** A copy of `range` on Sheet1 as the paste sees it. */
function copied(m: SpreadsheetModel, from: GridRange, sheetId = "s1"): { block: PasteBlock; styleIds: number[] } {
  const copy = buildClipboardCopy(m, sheetId, from);
  return { block: blockFromPayload(copy.payload), styleIds: copy.styleIds };
}

describe("paste geometry", () => {
  const m = source();
  const { block } = copied(m, range(1, 1, 2, 2));
  const s = m.sheets[0]!;

  it("pastes a block at the active cell of a one-cell selection", () => {
    expect(planPaste(s, at(9, 4), block)).toMatchObject({ dest: range(9, 4, 10, 5), down: 1, across: 1 });
  });

  it("tiles one cell over a multi-cell selection", () => {
    const one = copied(m, cellRange(1, 1)).block;
    expect(planPaste(s, over(range(4, 4, 6, 7)), one)).toMatchObject({ dest: range(4, 4, 6, 7), down: 3, across: 4 });
  });

  it("tiles an N×M block over a selection whose height and width are multiples", () => {
    expect(planPaste(s, over(range(4, 4, 7, 9)), block)).toMatchObject({ dest: range(4, 4, 7, 9), down: 2, across: 3 });
  });

  it("pastes once at the top-left of any other selection", () => {
    expect(planPaste(s, over(range(4, 4, 6, 9)), block)).toMatchObject({ dest: range(4, 4, 5, 5), down: 1, across: 1 });
  });

  it("refuses a multi-range destination, the grid's edge and more than MAX_COPY_CELLS", () => {
    expect(planPaste(s, { ranges: [cellRange(0, 0), cellRange(5, 5)], active: { row: 0, col: 0 } }, block)).toBe(MULTIPLE_SELECTIONS_MESSAGE);
    expect(planPaste(s, at(MAX_ROWS - 1, 0), block)).toBe(PASTE_PAST_EDGE_MESSAGE);
    expect(planPaste(s, at(0, MAX_COLS - 1), block)).toBe(PASTE_PAST_EDGE_MESSAGE);
    const one = copied(m, cellRange(1, 1)).block;
    expect(planPaste(s, over(range(0, 0, 999, 999)), one)).not.toBeTypeOf("string");
    expect(planPaste(s, over(range(0, 0, 1000, 999)), one)).toBe(TOO_MANY_TO_CHANGE_MESSAGE);
  });

  it("refuses a destination partly over a merge, part of an array, or a locked cell — and changes nothing", () => {
    const cases: [Partial<Sheet>, string][] = [
      [{ merges: [range(5, 5, 5, 7)] }, PASTE_MERGE_MESSAGE],
      [{ rows: [row(4, [value(5, 1, { formula: "1", arrayRange: range(4, 5, 6, 5) })])] }, PARTIAL_ARRAY_MESSAGE],
      [{ protection: { sheet: true } }, PROTECTED_CELL_MESSAGE],
    ];
    for (const [fields, message] of cases) {
      const target = model([sheet("t", "T", fields)]);
      const sheets = target.sheets;
      expect(pasteBlock(target, "t", at(4, 4), block, ALL)).toEqual({ ok: false, message });
      expect(target.sheets).toBe(sheets);
    }
  });

  it("replaces a merge entirely inside the destination, and pastes one cell into a merge's anchor keeping it", () => {
    const target = model([sheet("t", "T", { merges: [range(4, 4, 4, 5)] })]);
    expect(pasteBlock(target, "t", at(4, 4), block, ALL)).toMatchObject({ ok: true });
    expect(target.sheets[0]!.merges).toEqual([]);
    const intoMerge = model([sheet("t", "T", { merges: [range(4, 4, 5, 5)] })]);
    const one = copied(m, cellRange(1, 2)).block;
    expect(pasteBlock(intoMerge, "t", over(range(4, 4, 5, 5)), one, ALL)).toMatchObject({ ok: true, range: cellRange(4, 4) });
    expect(intoMerge.sheets[0]!.merges).toEqual([range(4, 4, 5, 5)]);
    expect(cellOf(intoMerge, 4, 4, "t")?.value).toBe(2);
    expect(cellOf(intoMerge, 5, 5, "t")).toBeUndefined();
  });
});

describe("Paste (All)", () => {
  it("replaces values, formulas (relative references shifted, absolute kept), styles and merges", () => {
    const m = source();
    m.sheets[0]!.merges = [range(1, 1, 1, 2)];
    const { block, styleIds } = copied(m, range(1, 1, 2, 2));
    const result = pasteBlock(m, "s1", at(4, 4), block, ALL, { styleIds });
    expect(result).toEqual({ ok: true, range: range(4, 4, 5, 5), formulasAsValues: 0 });
    expect(cellOf(m, 4, 4)).toEqual(value(4, 1, { styleId: 1 }));
    expect(cellOf(m, 5, 4)).toEqual(value(4, 10, { formula: "E5*10" }));
    expect(cellOf(m, 5, 5)).toEqual(value(5, 3, { formula: "$B$2+F5" }));
    expect(m.sheets[0]!.merges).toEqual([range(1, 1, 1, 2), range(4, 4, 4, 5)]);
    expect(m.sheets[0]!.dirty).toBe(true);
  });

  it("turns a relative reference shifted off the grid into #REF!", () => {
    const m = model([sheet("s1", "S", { rows: [row(5, [value(5, 0, { formula: "A1+F6" })])] })]);
    const { block } = copied(m, cellRange(5, 5));
    pasteBlock(m, "s1", at(0, 0), block, ALL);
    expect(cellOf(m, 0, 0)?.formula).toBe("#REF!+A1");
  });

  it("clears a destination under a blank source position — value and style, preserved notes kept", () => {
    const m = model([sheet("s1", "S", { rows: [row(0, [value(0, 1)]), row(4, [value(4, "x", { styleId: 1 }), value(5, "y", { preservedXml: "<note/>" })])] })], {
      styles: [{}, { font: { italic: true } }],
    });
    const { block } = copied(m, range(0, 0, 0, 2));
    pasteBlock(m, "s1", at(4, 3), block, ALL);
    expect(cellOf(m, 4, 3)?.value).toBe(1);
    expect(cellOf(m, 4, 4)).toBeUndefined();
    expect(cellOf(m, 4, 5)).toEqual(value(5, null, { preservedXml: "<note/>" }));
  });

  it("keeps style ids by identity in the same workbook, interns other workbooks' styles", () => {
    const m = source();
    const duplicate: CellStyle = { font: { bold: true } };
    m.styles = [{}, { font: { bold: true }, sourceXf: 1 }, { ...duplicate, sourceXf: 9 }];
    m.sheets[0]!.rows[0]!.cells[0]!.styleId = 2;
    const { block, styleIds } = copied(m, cellRange(1, 1));
    pasteBlock(m, "s1", at(6, 6), block, ALL, { styleIds });
    expect(cellOf(m, 6, 6)?.styleId).toBe(2); // the copy's own id, not the first equal one
    const other = model([sheet("t", "T")], { id: "other", styles: [{}, { fill: { kind: "solid", color: "#00ff00" } }] });
    pasteBlock(other, "t", at(0, 0), block, ALL, { styleIds });
    expect(other.styles).toHaveLength(3);
    expect(cellStyle(other, cellOf(other, 0, 0, "t")!.styleId!)).toEqual({ font: { bold: true } });
  });

  it("carries rich text and a whole array formula", () => {
    const m = model([
      sheet("s1", "S", {
        rows: [row(0, [value(0, 3, { formula: "B1:B2*1", arrayRange: range(0, 0, 1, 0) }), value(1, "ab", { runs: [{ text: "a", font: { bold: true } }, { text: "b" }] })]), row(1, [value(0, 4)])],
      }),
    ]);
    const { block } = copied(m, range(0, 0, 1, 1));
    pasteBlock(m, "s1", at(10, 3), block, ALL);
    expect(cellOf(m, 10, 3)).toEqual(value(3, 3, { formula: "E11:E12*1", arrayRange: range(10, 3, 11, 3) }));
    expect(cellOf(m, 10, 4)?.runs).toEqual([{ text: "a", font: { bold: true } }, { text: "b" }]);
    expect(cellOf(m, 11, 3)).toEqual(value(3, 4));
  });

  it("replaces whole rows including heights and row styles, clearing what the source rows lacked", () => {
    const m = model(
      [
        sheet("s1", "S", {
          rows: [row(0, [value(0, "a")], { heightPt: 30, customHeight: true, styleId: 1 }), row(5, [value(0, "old"), value(9, "far")], { heightPt: 20 })],
        }),
      ],
      { styles: [{}, { fill: { kind: "solid", color: "#ff0000" } }] },
    );
    const { block } = copied(m, range(0, 0, 0, MAX_COLS - 1));
    expect(pasteCapabilities(block).transpose).toBe(false);
    pasteBlock(m, "s1", at(5, 3), block, ALL);
    const pasted = m.sheets[0]!.rows.find((each) => each.index === 5)!;
    expect(pasted).toEqual(row(5, [value(0, "a")], { heightPt: 30, customHeight: true, styleId: 1 }));
  });

  it("replaces whole columns including widths and column styles", () => {
    const m = model([sheet("s1", "S", { rows: [row(0, [value(0, 1)])], columns: [{ min: 0, max: 0, widthPx: 120, styleId: 1 }, { min: 3, max: 5, widthPx: 30 }] })], {
      styles: [{}, { font: { italic: true } }],
    });
    const { block } = copied(m, range(0, 0, MAX_ROWS - 1, 0));
    pasteBlock(m, "s1", at(7, 4), block, ALL);
    expect(m.sheets[0]!.columns).toEqual([
      { min: 0, max: 0, widthPx: 120, styleId: 1 },
      { min: 3, max: 3, widthPx: 30 },
      { min: 4, max: 4, widthPx: 120, styleId: 1 },
      { min: 5, max: 5, widthPx: 30 },
    ]);
    expect(cellOf(m, 0, 4)?.value).toBe(1);
  });
});

describe("Paste special", () => {
  /** A destination D2 styled italic with a value. */
  function destination(): SpreadsheetModel {
    const m = source();
    m.styles = [...m.styles, { font: { italic: true } }, { numberFormat: "0.00%" }];
    m.sheets[0]!.rows.push(row(7, [value(3, "keep", { styleId: 2 }), value(4, "keep", { styleId: 2 })]));
    m.sheets[0]!.rows[0]!.cells[1]!.styleId = 3; // C2 0.00%
    return m;
  }

  it("Formulas: formulas shifted, constants as values, destination styles kept", () => {
    const m = destination();
    const { block } = copied(m, range(1, 1, 2, 2));
    pasteBlock(m, "s1", at(7, 3), block, { what: "formulas" });
    expect(cellOf(m, 7, 3)).toEqual(value(3, 1, { styleId: 2 }));
    expect(cellOf(m, 8, 3)).toEqual(value(3, 10, { formula: "D8*10" }));
  });

  it("Values: cached values only, formulas and runs dropped, styles kept", () => {
    const m = destination();
    const { block } = copied(m, range(1, 1, 2, 2));
    pasteBlock(m, "s1", at(7, 3), block, { what: "values" });
    expect(cellOf(m, 7, 3)).toEqual(value(3, 1, { styleId: 2 }));
    expect(cellOf(m, 8, 3)).toEqual(value(3, 10));
  });

  it("Formats: values kept, styles replaced", () => {
    const m = destination();
    const { block, styleIds } = copied(m, range(1, 1, 1, 2));
    pasteBlock(m, "s1", at(7, 3), block, { what: "formats" }, { styleIds });
    expect(cellOf(m, 7, 3)).toEqual(value(3, "keep", { styleId: 1 }));
    expect(cellOf(m, 7, 4)).toEqual(value(4, "keep", { styleId: 3 }));
  });

  it("Values and number formats: values, and only the number format merged into the destination's style", () => {
    const m = destination();
    const { block } = copied(m, range(1, 2, 1, 2));
    pasteBlock(m, "s1", at(7, 3), block, { what: "values-number-formats" });
    const pasted = cellOf(m, 7, 3)!;
    expect(pasted.value).toBe(2);
    expect(cellStyle(m, pasted.styleId!)).toEqual({ font: { italic: true }, numberFormat: "0.00%" });
  });

  it("Skip blanks: a source position with no value and no formula leaves its destination untouched, under every option", () => {
    for (const what of ["all", "formulas", "values", "formats", "values-number-formats"] as const) {
      const m = destination();
      const { block } = copied(m, range(1, 1, 1, 3)); // D2 is blank
      pasteBlock(m, "s1", at(7, 2), block, { what, skipBlanks: true });
      expect(cellOf(m, 7, 4), what).toEqual(value(4, "keep", { styleId: 2 }));
    }
  });

  it("Transpose: rows ↔ columns, self-referring formulas still self-refer, absolute and mixed shift by the host's displacement, merges transposed", () => {
    const m = model([
      sheet("s1", "S", {
        rows: [row(0, [value(0, 1), value(1, 2), value(2, 3, { formula: "A1+B1+$A$1+A$1" })]), row(1, [value(0, "under")])],
        merges: [range(0, 0, 1, 0)],
      }),
    ]);
    const { block } = copied(m, range(0, 0, 1, 2));
    expect(pasteBlock(m, "s1", at(10, 10), block, { what: "all", transpose: true })).toMatchObject({ ok: true, range: range(10, 10, 12, 11) });
    expect(cellOf(m, 10, 10)?.value).toBe(1);
    expect(cellOf(m, 11, 10)?.value).toBe(2);
    expect(cellOf(m, 10, 11)?.value).toBe("under");
    // C1 → K13: A1/B1 (offsets (0,-2)/(0,-1)) become K11/K12; $A$1 stays;
    // A$1 shifts its column by the host's displacement (+8 columns).
    expect(cellOf(m, 12, 10)?.formula).toBe("K11+K12+$A$1+I$1");
    expect(m.sheets[0]!.merges).toEqual([range(0, 0, 1, 0), range(10, 10, 10, 11)]);
    expect(transposeFormula("SUM(A1:B1)", { row: 0, col: 2 }, { row: 2, col: 0 })).toBe("SUM(A1:A2)");
  });

  it("offers only what a source can supply", () => {
    const m = source();
    expect(pasteCapabilities(copied(m, range(1, 1, 2, 2)).block)).toEqual({ formulas: true, formats: true, numberFormats: true, transpose: true });
    expect(pasteCapabilities(blockFromText("a\tb", { locale: EN_US, dateSystem: "1900" }) as PasteBlock)).toEqual({ formulas: false, formats: false, numberFormats: true, transpose: true });
    const excel = importClipboardHtml('<html xmlns:x="urn:schemas-microsoft-com:office:excel"><table><tr><td>1</td></tr></table></html>', { locale: EN_US, dateSystem: "1900" }) as PasteBlock;
    expect(pasteCapabilities(excel)).toEqual({ formulas: false, formats: true, numberFormats: true, transpose: true });
  });
});

describe("date systems and other workbooks", () => {
  /** A 1900 workbook: A1 a date (14), A2 a time-only value, A3 a number,
   * A4 =Data!A1, A5 =Sheet9!B2, A6 =[1]Ext!A1. */
  function dated(): SpreadsheetModel {
    return model(
      [
        sheet("s1", "Sheet1", {
          rows: [
            row(0, [value(0, 46096, { styleId: 1 })]),
            row(1, [value(0, 0.375, { styleId: 2 })]),
            row(2, [value(0, 46096)]),
            row(3, [value(0, 5, { formula: "Data!A1" })]),
            row(4, [value(0, 6, { formula: "Sheet9!B2*2" })]),
            row(5, [value(0, 7, { formula: "SUM(Sheet9:Sheet10!A1)" })]),
          ],
        }),
      ],
      { styles: [{}, { numberFormat: "m/d/yyyy", builtinFormatId: 14 }, { numberFormat: "h:mm" }] },
    );
  }

  it("converts date-formatted values between 1900 and 1904 (±1462 days); times and plain numbers stay", () => {
    const src = dated();
    const { block } = copied(src, range(0, 0, 2, 0));
    const into1904 = model([sheet("t", "T")], { id: "other", dateSystem: "1904" });
    pasteBlock(into1904, "t", at(0, 0), block, ALL);
    expect([0, 1, 2].map((r) => cellOf(into1904, r, 0, "t")?.value)).toEqual([46096 - 1462, 0.375, 46096]);
    const back = blockFromPayload(buildClipboardCopy(into1904, "t", range(0, 0, 2, 0)).payload);
    const into1900 = model([sheet("u", "U")], { id: "third" });
    pasteBlock(into1900, "u", at(0, 0), back, ALL);
    expect(cellOf(into1900, 0, 0, "u")?.value).toBe(46096);
  });

  it("pastes another workbook's formulas naming a sheet this one lacks as their cached values, counted for the toast", () => {
    const src = dated();
    const { block } = copied(src, range(3, 0, 5, 0));
    const target = model([sheet("t", "Data")], { id: "other" });
    const result = pasteBlock(target, "t", at(4, 1), block, ALL);
    expect(result).toEqual({ ok: true, range: range(4, 1, 6, 1), formulasAsValues: 2 });
    expect(cellOf(target, 4, 1, "t")).toEqual(value(1, 5, { formula: "Data!B2" }));
    expect(cellOf(target, 5, 1, "t")).toEqual(value(1, 6));
    expect(cellOf(target, 6, 1, "t")).toEqual(value(1, 7));
    expect(formulasAsValuesMessage(2)).toBe("2 formulas referring to other sheets were pasted as values");
    // The same workbook keeps them as formulas.
    const same = dated();
    expect(pasteBlock(same, "s1", at(10, 0), copied(same, range(4, 0, 4, 0)).block, ALL)).toMatchObject({ formulasAsValues: 0 });
    expect(cellOf(same, 10, 0)?.formula).toBe("Sheet9!B8*2");
  });

  it("converts Google Sheets R1C1 formulas at each destination cell, an unconvertible one pasting its value", () => {
    const html =
      '<google-sheets-html-origin><table data-sheets-root="1"><tr><td data-sheets-value=\'{"1":3,"3":2}\'>2</td>' +
      '<td data-sheets-value=\'{"1":3,"3":4}\' data-sheets-formula="=R[0]C[-1]*2">4</td>' +
      '<td data-sheets-value=\'{"1":3,"3":9}\' data-sheets-formula="=R[-5]C[0]">9</td></tr></table></google-sheets-html-origin>';
    const block = importClipboardHtml(html, { locale: EN_US, dateSystem: "1900" }) as PasteBlock;
    const m = model([sheet("t", "T")]);
    pasteBlock(m, "t", at(2, 3), block, ALL);
    expect(cellOf(m, 2, 4, "t")).toEqual(value(4, 4, { formula: "D3*2" }));
    expect(cellOf(m, 2, 5, "t")).toEqual(value(5, 9));
  });

  it("transposes Google Sheets formulas with the block, and pastes one naming a sheet this workbook lacks as its value, counted", () => {
    const html =
      '<google-sheets-html-origin><table data-sheets-root="1"><tr><td data-sheets-value=\'{"1":3,"3":2}\'>2</td>' +
      '<td data-sheets-value=\'{"1":3,"3":4}\' data-sheets-formula="=RC[-1]*2+R1C1">4</td>' +
      '<td data-sheets-value=\'{"1":3,"3":7}\' data-sheets-formula="=Missing!R1C1">7</td>' +
      '<td data-sheets-value=\'{"1":3,"3":8}\' data-sheets-formula="=T!R1C1">8</td></tr></table></google-sheets-html-origin>';
    const block = importClipboardHtml(html, { locale: EN_US, dateSystem: "1900" }) as PasteBlock;
    const m = model([sheet("t", "T")]);
    expect(pasteBlock(m, "t", at(2, 3), block, { what: "all", transpose: true })).toEqual({ ok: true, range: range(2, 3, 5, 3), formulasAsValues: 1 });
    // B1's left neighbour is, transposed at D4, the cell above it; R1C1 is absolute.
    expect(cellOf(m, 3, 3, "t")).toEqual(value(3, 4, { formula: "D3*2+$A$1" }));
    expect(cellOf(m, 4, 3, "t")).toEqual(value(3, 7));
    expect(cellOf(m, 5, 3, "t")).toEqual(value(3, 8, { formula: "T!$A$1" }));
    const plain = model([sheet("t", "T")]);
    expect(pasteBlock(plain, "t", at(2, 3), block, ALL)).toMatchObject({ formulasAsValues: 1 });
    expect(cellOf(plain, 2, 4, "t")?.formula).toBe("D3*2+$A$1");
  });
});

describe("plain text", () => {
  const ctx = { locale: EN_US, dateSystem: "1900" as const };

  it("reads each field through parseInput: 1,234.50 is 1234.5 with its implied format, =A1+1 a formula", () => {
    const block = blockFromText('1,234.50\t=A1+1\r\n"two\nlines"\t\r\n', ctx) as PasteBlock;
    expect([block.rows, block.cols]).toEqual([2, 2]);
    const m = model([sheet("t", "T")]);
    pasteBlock(m, "t", at(0, 0), block, ALL);
    expect(cellOf(m, 0, 0, "t")?.value).toBe(1234.5);
    expect(cellStyle(m, cellOf(m, 0, 0, "t")!.styleId!).numberFormat).toBe("#,##0.00");
    expect(cellOf(m, 0, 1, "t")).toEqual(value(1, null, { formula: "A1+1" }));
    expect(cellOf(m, 1, 0, "t")?.value).toBe("two\nlines");
  });

  it("keeps destination styles except the number format parseInput implies over General; an empty field clears the value", () => {
    const m = model([sheet("t", "T", { rows: [row(0, [value(0, "old", { styleId: 1 }), value(1, "old", { styleId: 2 })])] })], {
      styles: [{}, { font: { bold: true } }, { font: { bold: true }, numberFormat: "0.0" }],
    });
    pasteBlock(m, "t", at(0, 0), blockFromText("50%\t50%\r\n", ctx) as PasteBlock, ALL);
    expect(cellStyle(m, cellOf(m, 0, 0, "t")!.styleId!)).toEqual({ font: { bold: true }, numberFormat: "0%" });
    expect(cellOf(m, 0, 1, "t")).toEqual(value(1, 0.5, { styleId: 2 }));
    pasteBlock(m, "t", at(0, 0), blockFromText("\tx", ctx) as PasteBlock, ALL);
    expect(cellOf(m, 0, 0, "t")).toEqual(value(0, null, { styleId: 3 }));
  });

  it("pastes a single value into the active cell, or tiles it over the selection", () => {
    const m = model([sheet("t", "T")]);
    pasteBlock(m, "t", over(range(1, 1, 2, 2)), blockFromText("7", ctx) as PasteBlock, ALL);
    expect([cellOf(m, 1, 1, "t")?.value, cellOf(m, 2, 2, "t")?.value, cellOf(m, 3, 3, "t")]).toEqual([7, 7, undefined]);
  });

  it("Values pastes a formula field as its text (as typed after an apostrophe) and drops implied formats; Values and number formats keeps them", () => {
    const m = model([sheet("t", "T")]);
    pasteBlock(m, "t", at(1, 0), blockFromText("=1+1\t5%", ctx) as PasteBlock, { what: "values" });
    const text = cellOf(m, 1, 0, "t")!;
    expect([text.value, text.formula]).toEqual(["=1+1", undefined]);
    expect(cellStyle(m, text.styleId!)).toEqual({ quotePrefix: true });
    expect(cellOf(m, 1, 1, "t")).toEqual(value(1, 0.05));
    pasteBlock(m, "t", at(3, 0), blockFromText("=1+1\t5%", ctx) as PasteBlock, { what: "values-number-formats" });
    expect([cellOf(m, 3, 0, "t")?.value, cellOf(m, 3, 0, "t")?.formula]).toEqual(["=1+1", undefined]);
    expect(cellStyle(m, cellOf(m, 3, 1, "t")!.styleId!).numberFormat).toBe("0%");
  });

  it("text with no field but empty ones — a lone line break — is nothing to paste", () => {
    for (const text of ["\n", "\r\n", "\t\t\r\n", "\n\n"]) {
      expect(blockFromText(text, ctx), JSON.stringify(text)).toBeNull();
    }
    expect((blockFromText("\nx", ctx) as PasteBlock).cells).toEqual([expect.objectContaining({ r: 1, c: 0, v: "x" })]);
  });

  it("Paste (All) replaces a merge inside the destination, so every field shows; Values keeps it; one value into a merge's anchor keeps it", () => {
    const merged = (): SpreadsheetModel => model([sheet("t", "T", { merges: [range(0, 0, 1, 1)] })]);
    const fields = (m: SpreadsheetModel) => [cellOf(m, 0, 0, "t")?.value, cellOf(m, 0, 1, "t")?.value, cellOf(m, 1, 0, "t")?.value, cellOf(m, 1, 1, "t")?.value];
    const all = merged();
    expect(pasteBlock(all, "t", at(0, 0), blockFromText("a\tb\r\nc\td", ctx) as PasteBlock, ALL)).toMatchObject({ ok: true, range: range(0, 0, 1, 1) });
    expect(all.sheets[0]!.merges).toEqual([]);
    expect(fields(all)).toEqual(["a", "b", "c", "d"]);
    expect(all.sheets[0]!.dirty).toBe(true);
    const values = merged();
    pasteBlock(values, "t", at(0, 0), blockFromText("a\tb\r\nc\td", ctx) as PasteBlock, { what: "values" });
    expect(values.sheets[0]!.merges).toEqual([range(0, 0, 1, 1)]);
    expect(fields(values)).toEqual(["a", undefined, undefined, undefined]);
    const anchor = merged();
    pasteBlock(anchor, "t", over(range(0, 0, 1, 1)), blockFromText("7", ctx) as PasteBlock, ALL);
    expect(anchor.sheets[0]!.merges).toEqual([range(0, 0, 1, 1)]);
    expect(fields(anchor)).toEqual([7, undefined, undefined, undefined]);
  });
});

describe("copy-on-write", () => {
  it("keeps every untouched row and cell by identity, and a refused paste touches nothing", () => {
    const m = source();
    m.sheets[0]!.rows.push(row(40, [value(0, "far")]));
    const untouched = m.sheets[0]!.rows[2]!;
    const sheet2 = m.sheets[1];
    const { block } = copied(m, range(1, 1, 1, 2));
    pasteBlock(m, "s1", at(10, 0), block, ALL);
    expect(m.sheets[0]!.rows.find((each) => each.index === 40)).toBe(untouched);
    expect(m.sheets[1]).toBe(sheet2);
    expect(JSON.parse(JSON.stringify(m))).toEqual(m);
  });
});
