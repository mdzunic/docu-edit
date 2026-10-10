// SPEC-122 "Cell commands": setCellEntry / fillWithEntry / clearContents
// copy-on-write (a new sheets array and sheet object, a new rows array, new
// objects only for touched rows and changed cells, nothing replaced when
// nothing changes), sparseness (empty cells and property-less rows removed,
// preserved-only cells kept), the drop rules, the style rules (formats
// applied to General only, quotePrefix, wrap through internStyle, an
// unchanged style keeping styleId as it was), arrays (anchor edits keep
// arrayRange, a literal over an anchor clears its members), Fill's relative
// shift, merges and whole-row/column clipping, Clear contents keeping
// styles and notes, firstLockedCell over stored data only and
// arrayRangeTouched.
import { describe, expect, it } from "vitest";
import type { GridRange } from "./address";
import { arrayRangeTouched, cellEntryEquals, clearContents, entryFromInput, fillWithEntry, firstLockedCell, isGeneralFormat, setCellEntry, type CellEntry } from "./cells";
import { MAX_COLS, MAX_ROWS, type Cell, type CellStyle, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { parseInput } from "./numberFormat/input";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { createBlankWorkbook } from "./sheets";

function book(rows: Row[], options: { styles?: CellStyle[]; sheet?: Partial<Sheet> } = {}): SpreadsheetModel {
  const model = createBlankWorkbook();
  model.sheets = [{ ...model.sheets[0]!, rows, ...options.sheet }, { ...model.sheets[0]!, id: "other", name: "Other", rows: [] }];
  if (options.styles) {
    model.styles = options.styles;
  }
  return model;
}

const sheetOf = (model: SpreadsheetModel): Sheet => model.sheets[0]!;
const rowAt = (model: SpreadsheetModel, index: number): Row | undefined => sheetOf(model).rows.find((row) => row.index === index);
const cellAt = (model: SpreadsheetModel, row: number, col: number): Cell | undefined => rowAt(model, row)?.cells.find((cell) => cell.col === col);
const range = (top: number, left: number, bottom = top, right = left): GridRange => ({ top, left, bottom, right });
const value = (v: CellEntry["value"]): CellEntry => ({ value: v });

describe("setCellEntry — copy-on-write and sparseness", () => {
  it("replaces the sheets array, the sheet and its rows, keeping untouched rows and cells", () => {
    const row0: Row = { index: 0, cells: [{ col: 0, value: "a" }, { col: 1, value: "b" }] };
    const row5: Row = { index: 5, cells: [{ col: 0, value: 5 }] };
    const model = book([row0, row5]);
    const sheets = model.sheets;
    const before = sheetOf(model);
    setCellEntry(model, 0, { row: 0, col: 1 }, value("B"));
    expect(model.sheets).not.toBe(sheets);
    expect(sheetOf(model)).not.toBe(before);
    expect(sheetOf(model).rows).not.toBe(before.rows);
    expect(model.sheets[1]).toBe(sheets[1]);
    expect(rowAt(model, 5)).toBe(row5);
    expect(rowAt(model, 0)).not.toBe(row0);
    expect(rowAt(model, 0)!.cells[0]).toBe(row0.cells[0]);
    expect(rowAt(model, 0)!.cells[1]).toEqual({ col: 1, value: "B" });
    // The source objects are never mutated.
    expect(row0.cells[1]).toEqual({ col: 1, value: "b" });
    expect(before.rows).toEqual([row0, row5]);
  });

  it("inserts new rows and cells in order", () => {
    const model = book([{ index: 2, cells: [{ col: 3, value: 1 }] }]);
    setCellEntry(model, 0, { row: 4, col: 0 }, value(4));
    setCellEntry(model, 0, { row: 0, col: 0 }, value(0));
    setCellEntry(model, 0, { row: 2, col: 1 }, value(2));
    expect(sheetOf(model).rows).toEqual([
      { index: 0, cells: [{ col: 0, value: 0 }] },
      { index: 2, cells: [{ col: 1, value: 2 }, { col: 3, value: 1 }] },
      { index: 4, cells: [{ col: 0, value: 4 }] },
    ]);
  });

  it("changes nothing for an entry equal to the cell's content", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: 12 }, { col: 1, value: { error: "#N/A" } }] }]);
    const sheets = model.sheets;
    setCellEntry(model, 0, { row: 0, col: 0 }, value(12));
    setCellEntry(model, 0, { row: 0, col: 1 }, value({ error: "#N/A" }));
    setCellEntry(model, 0, { row: 3, col: 3 }, value(null));
    expect(model.sheets).toBe(sheets);
  });

  it("removes a cell left empty, and its row when that has no properties", () => {
    const model = book([
      { index: 0, cells: [{ col: 0, value: 1 }] },
      { index: 1, cells: [{ col: 0, value: 1 }], heightPt: 30, customHeight: true },
    ]);
    setCellEntry(model, 0, { row: 0, col: 0 }, value(null));
    setCellEntry(model, 0, { row: 1, col: 0 }, value(null));
    expect(sheetOf(model).rows).toEqual([{ index: 1, cells: [], heightPt: 30, customHeight: true }]);
  });

  it("keeps a cell holding a style or preserved content when its value goes", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }, { col: 1, value: 2, preservedXml: "<extLst/>" }] }], {
      styles: [{}, { fill: { kind: "solid", color: "#ff0000" } }],
    });
    setCellEntry(model, 0, { row: 0, col: 0 }, value(null));
    setCellEntry(model, 0, { row: 0, col: 1 }, value(null));
    expect(rowAt(model, 0)!.cells).toEqual([
      { col: 0, value: null, styleId: 1 },
      { col: 1, value: null, preservedXml: "<extLst/>" },
    ]);
  });

  it("stores a formula with no value yet and keeps a cached value while the formula is unchanged", () => {
    const model = book([]);
    setCellEntry(model, 0, { row: 1, col: 1 }, { value: null, formula: "A2*2" });
    expect(cellAt(model, 1, 1)).toEqual({ col: 1, value: null, formula: "A2*2" });
    // Recalculation fills the value in; recommitting the same formula is no change.
    const computed: Row = { index: 1, cells: [{ col: 1, value: 2400, formula: "A2*2" }] };
    sheetOf(model).rows = [computed];
    const sheets = model.sheets;
    setCellEntry(model, 0, { row: 1, col: 1 }, { value: null, formula: "A2*2" });
    expect(model.sheets).toBe(sheets);
    setCellEntry(model, 0, { row: 1, col: 1 }, value(7));
    expect(cellAt(model, 1, 1)).toEqual({ col: 1, value: 7 });
  });

  it("ignores a sheet index out of range", () => {
    const model = book([]);
    const sheets = model.sheets;
    setCellEntry(model, 5, { row: 0, col: 0 }, value(1));
    expect(model.sheets).toBe(sheets);
  });
});

describe("setCellEntry — drop rules", () => {
  it("drops sstIndex, preservedText, vm and runs with a changed value, keeping other preserved content", () => {
    const cell: Cell = {
      col: 0,
      value: "ab",
      runs: [{ text: "a", font: { bold: true } }, { text: "b" }],
      sstIndex: 4,
      preservedAttrs: ' vm="2" ph="1"',
      preservedXml: "<extLst/>",
      preservedText: "<text:p>ab</text:p>",
    };
    const model = book([{ index: 0, cells: [cell] }]);
    setCellEntry(model, 0, { row: 0, col: 0 }, value("abc"));
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: "abc", preservedAttrs: ' ph="1"', preservedXml: "<extLst/>" });
  });

  it("removes preservedAttrs left empty by the vm drop", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: 1, preservedAttrs: " vm='1'" }] }]);
    setCellEntry(model, 0, { row: 0, col: 0 }, value(2));
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: 2 });
  });

  it("keeps the runs of an unchanged rich text (no change at all)", () => {
    const cell: Cell = { col: 0, value: "ab", runs: [{ text: "a", font: { bold: true } }, { text: "b" }], sstIndex: 4 };
    const model = book([{ index: 0, cells: [cell] }]);
    const sheets = model.sheets;
    setCellEntry(model, 0, { row: 0, col: 0 }, value("ab"));
    expect(model.sheets).toBe(sheets);
  });

  it("drops cm, dynamicArray, dataTableXml and formulaUnparsed with a changed formula, keeping the anchor's arrayRange", () => {
    const anchor: Cell = { col: 0, value: 1, formula: "SEQUENCE(2)", arrayRange: range(0, 0, 1, 0), dynamicArray: true, cm: 1, sstIndex: 3 };
    const model = book([
      { index: 0, cells: [anchor] },
      { index: 1, cells: [{ col: 0, value: 2 }] },
    ]);
    setCellEntry(model, 0, { row: 0, col: 0 }, { value: null, formula: "SEQUENCE(2)*2" });
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: null, formula: "SEQUENCE(2)*2", arrayRange: range(0, 0, 1, 0) });
    expect(cellAt(model, 1, 0)).toEqual({ col: 0, value: 2 });
  });

  it("drops the formula-bound fields when the formula is removed", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: 5, formula: "_xlfn.FOO()", formulaUnparsed: true, cm: 2 }] }]);
    setCellEntry(model, 0, { row: 0, col: 0 }, value(5));
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: 5 });
  });

  it("replaces a data table's formula with a literal typed over its anchor", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: 5, dataTableXml: '<f t="dataTable" ref="A1:B2" dt2D="0" dtr="0" r1="C1"/>' }] }]);
    setCellEntry(model, 0, { row: 0, col: 0 }, value(5));
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: 5 });
  });

  it("clears an array's members when a literal replaces its anchor", () => {
    const model = book(
      [
        { index: 0, cells: [{ col: 0, value: 1, formula: "B1:B3*2", arrayRange: range(0, 0, 2, 0) }] },
        { index: 1, cells: [{ col: 0, value: 2, styleId: 1 }] },
        { index: 2, cells: [{ col: 0, value: 3 }] },
        { index: 3, cells: [{ col: 0, value: 4 }] },
      ],
      { styles: [{}, { font: { bold: true } }] },
    );
    setCellEntry(model, 0, { row: 0, col: 0 }, value("x"));
    expect(sheetOf(model).rows).toEqual([
      { index: 0, cells: [{ col: 0, value: "x" }] },
      { index: 1, cells: [{ col: 0, value: null, styleId: 1 }] },
      { index: 3, cells: [{ col: 0, value: 4 }] },
    ]);
  });
});

describe("setCellEntry — styles", () => {
  it("applies a suggested format to a General cell through internStyle, once per distinct style", () => {
    const model = book([]);
    const styles = model.styles;
    setCellEntry(model, 0, { row: 1, col: 0 }, { value: 1200, numberFormat: "#,##0" });
    setCellEntry(model, 0, { row: 2, col: 0 }, { value: 7, numberFormat: "#,##0" });
    expect(model.styles).toHaveLength(2);
    expect(model.styles[0]).toBe(styles[0]);
    expect(model.styles[1]).toEqual({ ...styles[0], numberFormat: "#,##0" });
    expect(cellAt(model, 1, 0)).toEqual({ col: 0, value: 1200, styleId: 1 });
    expect(cellAt(model, 2, 0)).toEqual({ col: 0, value: 7, styleId: 1 });
  });

  it("carries a locale-dependent built-in id with its code", () => {
    const model = book([]);
    setCellEntry(model, 0, { row: 0, col: 0 }, { value: 46303, numberFormat: "m/d/yyyy", builtinFormatId: 14 });
    expect(model.styles[cellAt(model, 0, 0)!.styleId!]).toMatchObject({ numberFormat: "m/d/yyyy", builtinFormatId: 14 });
  });

  it("clears a General built-in id when a code alone is applied", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }] }], { styles: [{}, { builtinFormatId: 0 }] });
    setCellEntry(model, 0, { row: 0, col: 0 }, { value: 0.1, numberFormat: "0%" });
    expect(model.styles[cellAt(model, 0, 0)!.styleId!]).toEqual({ numberFormat: "0%" });
  });

  it("never changes a cell's format that is not General", () => {
    const styles: CellStyle[] = [{}, { numberFormat: "0.00", sourceXf: 1 }];
    const model = book([{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }] }], { styles });
    setCellEntry(model, 0, { row: 0, col: 0 }, { value: 46303, numberFormat: "m/d/yyyy", builtinFormatId: 14 });
    expect(model.styles).toBe(styles);
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: 46303, styleId: 1 });
  });

  it("keeps an inherited style inherited when it does not change (absent stays absent)", () => {
    const styles: CellStyle[] = [{}, { numberFormat: "0.0" }, { font: { italic: true } }];
    const model = book([{ index: 3, cells: [], styleId: 2 }], { styles, sheet: { columns: [{ min: 0, max: 0, styleId: 1 }] } });
    setCellEntry(model, 0, { row: 0, col: 0 }, { value: 1, numberFormat: "#,##0" });
    setCellEntry(model, 0, { row: 3, col: 5 }, value("x"));
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: 1 });
    expect(cellAt(model, 3, 5)).toEqual({ col: 5, value: "x" });
    expect(model.styles).toBe(styles);
  });

  it("styles from the row or column when the inherited style changes", () => {
    const styles: CellStyle[] = [{}, { font: { italic: true } }];
    const model = book([{ index: 0, cells: [], styleId: 1 }], { styles });
    setCellEntry(model, 0, { row: 0, col: 2 }, { value: 0.1, numberFormat: "0%" });
    expect(model.styles[cellAt(model, 0, 2)!.styleId!]).toEqual({ font: { italic: true }, numberFormat: "0%" });
  });

  it("keeps a styleId whose style is unchanged even when an equal style sits elsewhere in the table", () => {
    const styles: CellStyle[] = [{}, { font: { bold: true } }, { font: { bold: true } }];
    const model = book([{ index: 0, cells: [{ col: 0, value: 1, styleId: 2 }] }], { styles });
    setCellEntry(model, 0, { row: 0, col: 0 }, value(2));
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: 2, styleId: 2 });
  });

  it("sets quotePrefix with the entry and removes it otherwise", () => {
    const model = book([]);
    setCellEntry(model, 0, { row: 0, col: 0 }, { value: "0012", quotePrefix: true });
    const quoted = cellAt(model, 0, 0)!.styleId!;
    expect(model.styles[quoted]).toMatchObject({ quotePrefix: true });
    setCellEntry(model, 0, { row: 0, col: 0 }, value("12a"));
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: "12a" });
    expect(model.styles[0]!.quotePrefix).toBeUndefined();
  });

  it("turns wrapping on for a line break, sharing one interned style", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: "x", styleId: 1 }] }], {
      styles: [{}, { alignment: { horizontal: "center" }, sourceXf: 1 }],
    });
    setCellEntry(model, 0, { row: 0, col: 0 }, entryFromInput(parseInput("a\nb", SPREADSHEET_LOCALES["en-US"], "1900")));
    setCellEntry(model, 0, { row: 1, col: 0 }, { value: "c\nd", wrap: true });
    expect(model.styles).toHaveLength(4);
    expect(model.styles[2]).toEqual({ alignment: { horizontal: "center", wrap: true } });
    expect(model.styles[3]).toMatchObject({ alignment: { wrap: true } });
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: "a\nb", styleId: 2 });
    // A cell already wrapping keeps its style.
    setCellEntry(model, 0, { row: 0, col: 0 }, { value: "e\nf", wrap: true });
    expect(cellAt(model, 0, 0)!.styleId).toBe(2);
  });
});

describe("entry helpers", () => {
  it("maps parseInput's result, with wrap for a text holding a line break", () => {
    expect(entryFromInput({ value: 1200, numberFormat: "#,##0" })).toEqual({ value: 1200, numberFormat: "#,##0" });
    expect(entryFromInput({ value: "a\nb" })).toEqual({ value: "a\nb", wrap: true });
    expect(entryFromInput({ value: "a\nb" }, false)).toEqual({ value: "a\nb" });
    expect(entryFromInput({ value: null, formula: 'A1&"\n"' })).toEqual({ value: null, formula: 'A1&"\n"' });
    expect(entryFromInput({ value: "x", quotePrefix: true }, true)).toEqual({ value: "x", quotePrefix: true, wrap: true });
    expect(entryFromInput({ value: 46303, numberFormat: "m/d/yyyy", builtinFormatId: 14 })).toEqual({
      value: 46303,
      numberFormat: "m/d/yyyy",
      builtinFormatId: 14,
    });
  });

  it("compares entries field by field, errors by code", () => {
    expect(cellEntryEquals({ value: { error: "#N/A" } }, { value: { error: "#N/A" } })).toBe(true);
    expect(cellEntryEquals({ value: 1 }, { value: 1, numberFormat: "0" })).toBe(false);
    expect(cellEntryEquals({ value: null, formula: "A1" }, { value: null, formula: "A1" })).toBe(true);
  });

  it("tells a General format", () => {
    expect(isGeneralFormat({})).toBe(true);
    expect(isGeneralFormat({ numberFormat: "general" })).toBe(true);
    expect(isGeneralFormat({ builtinFormatId: 0 })).toBe(true);
    expect(isGeneralFormat({ builtinFormatId: 14, numberFormat: "m/d/yyyy" })).toBe(false);
    expect(isGeneralFormat({ numberFormat: "0" })).toBe(false);
  });
});

describe("fillWithEntry", () => {
  it("writes every target the entry with its formula shifted from the origin", () => {
    const model = book([]);
    fillWithEntry(model, 0, [range(0, 1, 2, 2)], { row: 0, col: 1 }, { value: null, formula: "A1*$A$1+SUM(A$1:A1)" });
    expect(cellAt(model, 0, 1)!.formula).toBe("A1*$A$1+SUM(A$1:A1)");
    expect(cellAt(model, 2, 1)!.formula).toBe("A3*$A$1+SUM(A$1:A3)");
    expect(cellAt(model, 1, 2)!.formula).toBe("B2*$A$1+SUM(B$1:B2)");
    expect(sheetOf(model).rows.flatMap((row) => row.cells)).toHaveLength(6);
  });

  it("shifts a reference off the grid to #REF!", () => {
    const model = book([]);
    fillWithEntry(model, 0, [range(0, 0, 1, 0)], { row: 1, col: 0 }, { value: null, formula: "A1+1" });
    expect(cellAt(model, 0, 0)!.formula).toBe("#REF!+1");
    expect(cellAt(model, 1, 0)!.formula).toBe("A1+1");
  });

  it("writes a merged range's anchor only, and keeps untouched rows", () => {
    const untouched: Row = { index: 9, cells: [{ col: 0, value: 1 }] };
    const model = book([untouched], { sheet: { merges: [range(0, 0, 1, 1)] } });
    fillWithEntry(model, 0, [range(0, 0, 2, 1)], { row: 0, col: 0 }, value("x"));
    expect(sheetOf(model).rows.map((row) => [row.index, row.cells.map((cell) => cell.col)])).toEqual([
      [0, [0]],
      [2, [0, 1]],
      [9, [0]],
    ]);
    expect(rowAt(model, 9)).toBe(untouched);
  });

  it("clips whole columns and rows to the used range, grown to include the origin", () => {
    const model = book([
      { index: 1, cells: [{ col: 1, value: 1 }] },
      { index: 3, cells: [{ col: 2, value: 1 }] },
    ]);
    fillWithEntry(model, 0, [range(0, 4, MAX_ROWS - 1, 4)], { row: 0, col: 4 }, value("c"));
    expect(sheetOf(model).rows.map((row) => row.index)).toEqual([0, 1, 2, 3]);
    expect(sheetOf(model).rows.every((row) => cellAt(model, row.index, 4)?.value === "c")).toBe(true);
    fillWithEntry(model, 0, [range(5, 0, 5, MAX_COLS - 1)], { row: 5, col: 0 }, value("r"));
    expect(rowAt(model, 5)!.cells.map((cell) => cell.col)).toEqual([0, 1, 2, 3, 4]);
  });

  it("fills only the origin of a whole column on an empty sheet", () => {
    const model = book([]);
    fillWithEntry(model, 0, [range(0, 2, MAX_ROWS - 1, 2)], { row: 3, col: 2 }, value(1));
    expect(sheetOf(model).rows).toEqual([{ index: 3, cells: [{ col: 2, value: 1 }] }]);
  });

  it("writes overlapping targets once and changes nothing when every target already holds the entry", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: "x" }, { col: 1, value: "x" }] }]);
    const sheets = model.sheets;
    fillWithEntry(model, 0, [range(0, 0, 0, 1), range(0, 1)], { row: 0, col: 0 }, value("x"));
    expect(model.sheets).toBe(sheets);
  });
});

describe("clearContents", () => {
  it("clears values and formulas, keeping styles, notes and other preserved attributes", () => {
    const model = book(
      [
        {
          index: 0,
          cells: [
            { col: 0, value: 1, styleId: 1 },
            { col: 1, value: "n", preservedXml: "<office:annotation/>", preservedText: "<text:p/>", sstIndex: 2 },
            { col: 2, value: 3, formula: "1+2", preservedAttrs: ' vm="1" ph="1"', cm: 1, dynamicArray: true, arrayRange: range(0, 2) },
            { col: 3, value: 4 },
          ],
        },
      ],
      { styles: [{}, { font: { bold: true } }] },
    );
    clearContents(model, 0, [range(0, 0, 0, 2)]);
    expect(rowAt(model, 0)!.cells).toEqual([
      { col: 0, value: null, styleId: 1 },
      { col: 1, value: null, preservedXml: "<office:annotation/>" },
      { col: 2, value: null, preservedAttrs: ' ph="1"' },
      { col: 3, value: 4 },
    ]);
  });

  it("touches only stored cells of whole columns and keeps untouched rows", () => {
    const keep: Row = { index: 4, cells: [{ col: 0, value: 1 }] };
    const model = book([{ index: 0, cells: [{ col: 1, value: 1 }] }, keep, { index: 900_000, cells: [{ col: 1, value: 2 }], hidden: true }]);
    clearContents(model, 0, [range(0, 1, MAX_ROWS - 1, 1)]);
    expect(sheetOf(model).rows).toEqual([keep, { index: 900_000, cells: [], hidden: true }]);
    expect(rowAt(model, 4)).toBe(keep);
  });

  it("changes nothing when nothing in the ranges holds a value", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: null, styleId: 1 }] }], { styles: [{}, { font: { bold: true } }] });
    const sheets = model.sheets;
    clearContents(model, 0, [range(0, 0, MAX_ROWS - 1, MAX_COLS - 1)]);
    expect(model.sheets).toBe(sheets);
  });

  it("clears a whole array, and an array's members with its anchor", () => {
    const rows = (): Row[] => [
      { index: 0, cells: [{ col: 0, value: 1, formula: "B1:B2", arrayRange: range(0, 0, 1, 0) }] },
      { index: 1, cells: [{ col: 0, value: 2 }] },
    ];
    const whole = book(rows());
    clearContents(whole, 0, [range(0, 0, 1, 0)]);
    expect(sheetOf(whole).rows).toEqual([]);
    const anchorOnly = book(rows());
    clearContents(anchorOnly, 0, [range(0, 0)]);
    expect(sheetOf(anchorOnly).rows).toEqual([]);
  });
});

describe("firstLockedCell", () => {
  const unlocked: CellStyle = { protection: { locked: false } };

  it("is null on an unprotected sheet", () => {
    expect(firstLockedCell(book([]), 0, [range(0, 0, 5, 5)])).toBeNull();
  });

  it("finds the first locked position row-major on a protected sheet", () => {
    const model = book(
      [
        { index: 0, cells: [{ col: 0, value: 1, styleId: 1 }, { col: 1, value: 1, styleId: 1 }] },
        { index: 1, cells: [{ col: 0, value: 1, styleId: 1 }] },
      ],
      { styles: [{}, unlocked], sheet: { protection: { sheet: true } } },
    );
    expect(firstLockedCell(model, 0, [range(0, 0)])).toBeNull();
    expect(firstLockedCell(model, 0, [range(0, 0, 0, 1)])).toBeNull();
    expect(firstLockedCell(model, 0, [range(0, 0, 0, 2)])).toEqual({ row: 0, col: 2 });
    expect(firstLockedCell(model, 0, [range(0, 0, 2, 1)])).toEqual({ row: 1, col: 1 });
    expect(firstLockedCell(model, 0, [range(0, 0), range(5, 5)])).toEqual({ row: 5, col: 5 });
  });

  it("judges whole columns, rows and the whole sheet by stored data alone", () => {
    const model = book(
      [
        { index: 2, cells: [{ col: 1, value: 1 }, { col: 3, value: 1, styleId: 0 }] },
        { index: 4, cells: [], styleId: 0 },
      ],
      { styles: [{}, unlocked], sheet: { protection: { sheet: true }, columns: [{ min: 1, max: 2, styleId: 1 }] } },
    );
    expect(firstLockedCell(model, 0, [range(0, 1, MAX_ROWS - 1, 2)])).toEqual({ row: 4, col: 1 });
    expect(firstLockedCell(model, 0, [range(0, 1, 3, 2)])).toBeNull();
    expect(firstLockedCell(model, 0, [range(2, 1, 2, MAX_COLS - 1)])).toEqual({ row: 2, col: 3 });
    expect(firstLockedCell(model, 0, [range(0, 0, MAX_ROWS - 1, MAX_COLS - 1)])).toEqual({ row: 0, col: 0 });
    expect(firstLockedCell(model, 0, [range(0, 1, 3, MAX_COLS - 1)])).toEqual({ row: 0, col: 3 });
  });

  it("answers null for a sheet whose default style is unlocked, whatever is selected", () => {
    const model = book([{ index: 7, cells: [{ col: 7, value: 1 }] }], { styles: [unlocked], sheet: { protection: { sheet: true } } });
    expect(firstLockedCell(model, 0, [range(0, 0, MAX_ROWS - 1, MAX_COLS - 1)])).toBeNull();
  });

  it("honours an unlocked row style and a locked cell inside it", () => {
    const model = book([{ index: 0, cells: [{ col: 4, value: 1, styleId: 0 }], styleId: 1 }], {
      styles: [{}, unlocked],
      sheet: { protection: { sheet: true } },
    });
    expect(firstLockedCell(model, 0, [range(0, 0, 0, 3)])).toBeNull();
    expect(firstLockedCell(model, 0, [range(0, 0, 0, MAX_COLS - 1)])).toEqual({ row: 0, col: 4 });
  });
});

describe("arrayRangeTouched", () => {
  const model = book([
    { index: 0, cells: [{ col: 0, value: 1, formula: "B1:B3*2", arrayRange: range(0, 0, 2, 1) }] },
    { index: 5, cells: [{ col: 5, value: 1, formula: "1", arrayRange: range(5, 5) }] },
  ]);
  const array = range(0, 0, 2, 1);

  it("refuses a member, or a part of the array, without the whole", () => {
    expect(arrayRangeTouched(model, 0, [range(1, 0)])).toEqual(array);
    expect(arrayRangeTouched(model, 0, [range(0, 0, 0, 1)])).toEqual(array);
    expect(arrayRangeTouched(model, 0, [range(0, 1, MAX_ROWS - 1, 1)])).toEqual(array);
  });

  it("allows the anchor alone, the whole array, and what does not touch it", () => {
    expect(arrayRangeTouched(model, 0, [range(0, 0)])).toBeNull();
    expect(arrayRangeTouched(model, 0, [array])).toBeNull();
    expect(arrayRangeTouched(model, 0, [range(0, 0, 9, 9)])).toBeNull();
    expect(arrayRangeTouched(model, 0, [range(0, 0, 2, 0), range(0, 1, 2, 1)])).toBeNull();
    expect(arrayRangeTouched(model, 0, [range(3, 0, 4, 4)])).toBeNull();
    expect(arrayRangeTouched(model, 0, [range(5, 5)])).toBeNull();
    expect(arrayRangeTouched(model, 0, [range(0, 0, MAX_ROWS - 1, MAX_COLS - 1)])).toBeNull();
  });
});

describe("SPEC-122 review fixes", () => {
  it("stores a changed style's interned id even when a row or column lends the same style, leaving only the default 0 absent", () => {
    const styles: CellStyle[] = [{}, { font: { bold: true } }, { font: { bold: true }, quotePrefix: true }];
    const model = book([{ index: 0, cells: [{ col: 0, value: "x", styleId: 2 }, { col: 1, value: "y", styleId: 3 }], styleId: 1 }], {
      styles: [...styles, { quotePrefix: true }],
    });
    setCellEntry(model, 0, { row: 0, col: 0 }, value("plain"));
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: "plain", styleId: 1 });
    const free = book([{ index: 0, cells: [{ col: 0, value: "x", styleId: 1 }] }], { styles: [{}, { quotePrefix: true }] });
    setCellEntry(free, 0, { row: 0, col: 0 }, value("plain"));
    expect(cellAt(free, 0, 0)).toEqual({ col: 0, value: "plain" });
  });

  it("drops the value-bound fields with a changed formula even when its cached value stays", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: 5, formula: "A2", preservedAttrs: ' vm="1"', preservedText: "<text:p>5</text:p>" }] }]);
    setCellEntry(model, 0, { row: 0, col: 0 }, { value: 5, formula: "A3" });
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: 5, formula: "A3" });
  });

  it("replaces a whole array with plain cells when a Fill covers it, keeping it when the anchor alone is written", () => {
    const rows = (): Row[] => [
      { index: 0, cells: [{ col: 0, value: 2, formula: "B1:B2*2", arrayRange: range(0, 0, 1, 0), dynamicArray: true, cm: 1 }] },
      { index: 1, cells: [{ col: 0, value: 4 }] },
    ];
    const whole = book(rows());
    fillWithEntry(whole, 0, [range(0, 0, 1, 0)], { row: 0, col: 0 }, { value: null, formula: "B1*2" });
    expect(sheetOf(whole).rows).toEqual([
      { index: 0, cells: [{ col: 0, value: null, formula: "B1*2" }] },
      { index: 1, cells: [{ col: 0, value: null, formula: "B2*2" }] },
    ]);
    const anchor = book(rows());
    fillWithEntry(anchor, 0, [range(0, 0), range(5, 5)], { row: 0, col: 0 }, { value: null, formula: "B1:B2*3" });
    expect(cellAt(anchor, 0, 0)).toEqual({ col: 0, value: null, formula: "B1:B2*3", arrayRange: range(0, 0, 1, 0) });
    expect(cellAt(anchor, 1, 0)).toEqual({ col: 0, value: 4 });
  });

  it("removes an emptied row whose only extras are layout caches or an ods row style, keeping one with an outline level", () => {
    const model = book([
      { index: 0, cells: [{ col: 0, value: 1 }], sourceAttrs: ' spans="1:3" x14ac:dyDescent="0.25"' },
      { index: 1, cells: [{ col: 0, value: 1 }], sourceStyleName: "ro1" },
      { index: 2, cells: [{ col: 0, value: 1 }], sourceAttrs: ' spans="1:3" outlineLevel="1"' },
    ]);
    clearContents(model, 0, [range(0, 0, 2, 0)]);
    expect(sheetOf(model).rows).toEqual([{ index: 2, cells: [], sourceAttrs: ' spans="1:3" outlineLevel="1"' }]);
  });

  it("judges General by the code a style renders with", () => {
    expect(isGeneralFormat({ builtinFormatId: 30 })).toBe(true); // a CJK built-in with no entry renders General
    expect(isGeneralFormat({ builtinFormatId: 30, numberFormat: "0.0" })).toBe(false);
    expect(isGeneralFormat({ builtinFormatId: 22, numberFormat: "m/d/yyyy h:mm" })).toBe(false);
  });
});
