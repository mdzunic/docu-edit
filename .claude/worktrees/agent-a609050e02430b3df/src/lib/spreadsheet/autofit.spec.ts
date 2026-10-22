// SPEC-122 "Row autofit": autofitRowHeightPt with fixedMeasurer — one line
// of the cell's font unwrapped (explicit line breaks included), greedy word-wrap
// at (column width − 5 px) when wrapped (explicit breaks and over-long words
// included; numbers never wrap), the rotated bounding box for 1–180, stacked
// text (255) as characters × line height, px → pt × 0.75 rounded UP to a
// 0.75 pt step, never below defaultRowHeightPt, the tallest cell winning and
// merged cells skipped; refitRows — only rows without customHeight holding a
// stored non-empty cell in the ranges (or a height of their own, so a
// cleared row shrinks back), heightPt absent when it equals the default,
// customHeight never set, untouched rows keeping their identity, the ods row
// style dropped with a changed height, the AUTOFIT_MAX_ROWS_PER_COMMAND
// cap (the first 10,000 rows only), and — given the rows before a command —
// only the rows it replaced or added refitting, none when it changed nothing.
import { describe, expect, it } from "vitest";
import type { GridRange } from "./address";
import { AUTOFIT_MAX_ROWS_PER_COMMAND, autofitRowHeightPt, fixedMeasurer, refitRows } from "./autofit";
import { MAX_ROWS, type Cell, type CellStyle, type ColumnSpan, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import type { FormatContext } from "./numberFormat/render";
import { createBlankWorkbook } from "./sheets";

const CTX: FormatContext = { locale: SPREADSHEET_LOCALES["en-US"], dateSystem: "1900" };
/** 10 px per character, 20 px per line. */
const MEASURER = fixedMeasurer(10, 20);

function book(rows: Row[], options: { styles?: CellStyle[]; columns?: ColumnSpan[]; sheet?: Partial<Sheet> } = {}): SpreadsheetModel {
  const model = createBlankWorkbook();
  const base = model.sheets[0]!;
  model.sheets = [{ ...base, rows, columns: options.columns ?? [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...options.sheet }];
  model.styles = options.styles ?? [{}];
  return model;
}

/** The height of row 0 holding `cells`. */
function heightOf(cells: Cell[], options: Parameters<typeof book>[1] = {}, measurer = MEASURER): number {
  const model = book([{ index: 0, cells }], options);
  return autofitRowHeightPt(model, model.sheets[0]!, 0, measurer, CTX);
}

const WRAP: CellStyle = { alignment: { wrap: true } };
/** "abcdef" (60 px, one line) in style `styleId`. */
const sixLetters = (styleId: number): Cell[] => [{ col: 0, value: "abcdef", styleId }];
/** A wrapped text two lines tall at the default 59 px: 30 pt. */
const twoLines = (col = 0): Cell => ({ col, value: "aaaaa aaaaa", styleId: 1 });
const rotated = (textRotation: number): CellStyle => ({ alignment: { textRotation } });

describe("fixedMeasurer", () => {
  it("measures every character at charPx and every line at linePx, whatever the font", () => {
    const measurer = fixedMeasurer(7, 18);
    const font = { family: "Calibri", sizePt: 30, bold: true, italic: false };
    expect(measurer.widthPx("abc", font)).toBe(21);
    expect(measurer.widthPx("", font)).toBe(0);
    expect(measurer.lineHeightPx(font)).toBe(18);
  });
});

describe("autofitRowHeightPt", () => {
  it("is one line of the cell's font for unwrapped text, however long", () => {
    expect(heightOf([{ col: 0, value: "a long line of text that never wraps" }], {}, fixedMeasurer(10, 24))).toBe(18);
  });

  it("is one line for unwrapped text with explicit line breaks too — a break makes a line only when the text wraps", () => {
    expect(heightOf([{ col: 0, value: "one\ntwo\nthree" }], {}, fixedMeasurer(10, 24))).toBe(18);
    expect(heightOf([{ col: 0, value: "one\ntwo\nthree", styleId: 1 }], { styles: [{}, WRAP] }, fixedMeasurer(10, 24))).toBe(54);
  });

  it("wraps greedily at the column width less 5 px", () => {
    const styles = [{}, WRAP];
    // 105 px column → 100 px of text: "aaaa bbbb" (90 px) fits, "cccc" starts a line.
    const columns: ColumnSpan[] = [{ min: 0, max: 0, widthPx: 105 }];
    expect(heightOf([{ col: 0, value: "aaaa bbbb cccc", styleId: 1 }], { styles, columns })).toBe(30); // 2 × 20 px
    // Exactly 100 px fits on one line at 105 px, not at 104 px (99 px of text).
    expect(heightOf([{ col: 0, value: "aaaaaaaaaa", styleId: 1 }], { styles, columns })).toBe(15);
    expect(heightOf([{ col: 0, value: "aaaaaaaaaa", styleId: 1 }], { styles, columns: [{ min: 0, max: 0, widthPx: 104 }] })).toBe(30);
  });

  it("wraps at the sheet's default width where no span sets one", () => {
    // 64 px − 5 → 59 px: five 10 px characters per line.
    expect(heightOf([{ col: 0, value: "abcde fghij klmno", styleId: 1 }], { styles: [{}, WRAP] })).toBe(45); // 3 × 20 px
  });

  it("keeps explicit line breaks inside wrapped text and breaks an over-long word between characters", () => {
    const styles = [{}, WRAP];
    const columns: ColumnSpan[] = [{ min: 0, max: 0, widthPx: 105 }];
    expect(heightOf([{ col: 0, value: "aa\nbb\ncc", styleId: 1 }], { styles, columns })).toBe(45); // 3 lines
    expect(heightOf([{ col: 0, value: "a".repeat(25), styleId: 1 }], { styles, columns })).toBe(45); // 10 + 10 + 5 characters
  });

  it("wraps justified and distributed text as the grid does, and never a number", () => {
    const columns: ColumnSpan[] = [{ min: 0, max: 0, widthPx: 105 }];
    const justified: CellStyle = { alignment: { horizontal: "justify" } };
    expect(heightOf([{ col: 0, value: "aaaa bbbb cccc", styleId: 1 }], { styles: [{}, justified], columns })).toBe(30);
    expect(heightOf([{ col: 0, value: 123_456_789_012_345, styleId: 1 }], { styles: [{}, WRAP], columns })).toBe(15);
  });

  it("takes the rotated bounding box for 1–180", () => {
    // "abcdef": 60 px wide, one 20 px line.
    const cell = sixLetters;
    const styles = [{}, rotated(90), rotated(45), rotated(180), rotated(135), rotated(30)];
    expect(heightOf(cell(1), { styles })).toBe(45); // 60 px
    // 60 sin 45° + 20 cos 45° = 56.57 px → 57 px → 42.75 pt.
    expect(heightOf(cell(2), { styles })).toBe(42.75);
    expect(heightOf(cell(3), { styles })).toBe(45); // 90° clockwise
    expect(heightOf(cell(4), { styles })).toBe(42.75); // 45° clockwise
    // 60 sin 30° + 20 cos 30° = 47.32 px → 48 px → 36 pt.
    expect(heightOf(cell(5), { styles })).toBe(36);
  });

  it("stacks text (255) as its longest line's characters × the line height", () => {
    const styles = [{}, rotated(255)];
    expect(heightOf([{ col: 0, value: "abcd", styleId: 1 }], { styles })).toBe(60); // 4 × 20 px
    expect(heightOf([{ col: 0, value: "ab\nabcde", styleId: 1 }], { styles })).toBe(75); // 5 × 20 px
  });

  it("rounds px → pt up to a 0.75 pt step", () => {
    const low = { sheet: { defaultRowHeightPt: 1 } };
    expect(heightOf([{ col: 0, value: "x" }], low, fixedMeasurer(10, 20))).toBe(15);
    expect(heightOf([{ col: 0, value: "x" }], low, fixedMeasurer(10, 20.01))).toBe(15.75);
    expect(heightOf([{ col: 0, value: "x" }], low, fixedMeasurer(10, 20.99))).toBe(15.75);
    expect(heightOf([{ col: 0, value: "x" }], low, fixedMeasurer(10, 21))).toBe(15.75);
    expect(heightOf([{ col: 0, value: "x" }], low, fixedMeasurer(10, 21.5))).toBe(16.5);
  });

  it("is never below the sheet's default height", () => {
    expect(heightOf([{ col: 0, value: "x" }], {}, fixedMeasurer(10, 8))).toBe(15);
    expect(heightOf([{ col: 0, value: "x" }], { sheet: { defaultRowHeightPt: 30 } })).toBe(30);
    const model = book([]);
    expect(autofitRowHeightPt(model, model.sheets[0]!, 7, MEASURER, CTX)).toBe(15); // no stored row
  });

  it("is the tallest cell of the row, skipping merged cells", () => {
    const styles = [{}, WRAP, rotated(90)];
    const cells: Cell[] = [
      { col: 0, value: "short" },
      { col: 1, value: "aaaaa aaaaa aaaaa", styleId: 1 }, // 3 lines at 59 px → 60 px
      { col: 2, value: "abc", styleId: 2 }, // 30 px
    ];
    expect(heightOf(cells, { styles })).toBe(45);
    expect(heightOf(cells, { styles, sheet: { merges: [{ top: 0, left: 1, bottom: 1, right: 1 }] } })).toBe(22.5);
  });

  it("measures the displayed text through the cell's number format", () => {
    // 0.5 as "50.00%" (6 characters) in a 90° rotation: 60 px.
    const styles = [{}, { numberFormat: "0.00%", alignment: { textRotation: 90 } }];
    expect(heightOf([{ col: 0, value: 0.5, styleId: 1 }], { styles })).toBe(45);
  });
});

describe("refitRows", () => {
  const all: GridRange = { top: 0, left: 0, bottom: MAX_ROWS - 1, right: 0 };
  const styles = [{}, WRAP];

  it("sets heightPt on rows without customHeight holding a non-empty cell in the ranges, never customHeight", () => {
    const custom: Row = { index: 1, cells: [twoLines()], heightPt: 50, customHeight: true };
    const outside: Row = { index: 2, cells: [twoLines(1)] };
    const model = book([{ index: 0, cells: [twoLines()] }, custom, outside], { styles });
    refitRows(model, model.sheets[0]!.id, [all], MEASURER, CTX);
    const rows = model.sheets[0]!.rows;
    expect(rows[0]).toEqual({ index: 0, cells: [twoLines()], heightPt: 30 });
    expect(rows[1]).toBe(custom);
    expect(rows[2]).toBe(outside); // its content is outside the ranges
  });

  it("leaves heightPt absent when it equals the default, and keeps unchanged rows and the model as they were", () => {
    const plain: Row = { index: 0, cells: [{ col: 0, value: "x" }] };
    const model = book([plain], { styles });
    const sheets = model.sheets;
    refitRows(model, model.sheets[0]!.id, [all], MEASURER, CTX);
    expect(model.sheets).toBe(sheets);
    expect(model.sheets[0]!.rows[0]).toBe(plain);
  });

  it("shrinks a row back to the default once its content is gone, dropping the stored height", () => {
    const model = book(
      [
        { index: 0, cells: [{ col: 0, value: "x" }], heightPt: 30 },
        { index: 1, cells: [], heightPt: 30 },
        { index: 2, cells: [{ col: 0, value: null, styleId: 1 }], heightPt: 30, sourceStyleName: "ro2" },
      ],
      { styles },
    );
    refitRows(model, model.sheets[0]!.id, [all], MEASURER, CTX);
    expect(model.sheets[0]!.rows).toEqual([
      { index: 0, cells: [{ col: 0, value: "x" }] },
      { index: 2, cells: [{ col: 0, value: null, styleId: 1 }] },
    ]);
  });

  it("copies on write: a new sheets array, sheet and rows array, untouched rows kept, the ods row style dropped", () => {
    const untouched: Row = { index: 5, cells: [{ col: 0, value: "x" }] };
    const model = book([{ index: 0, cells: [twoLines()], sourceStyleName: "ro1" }, untouched], { styles });
    const before = model.sheets[0]!;
    refitRows(model, before.id, [{ top: 0, left: 0, bottom: 0, right: 0 }], MEASURER, CTX);
    const after = model.sheets[0]!;
    expect(after).not.toBe(before);
    expect(after.rows).not.toBe(before.rows);
    expect(after.rows[0]).toEqual({ index: 0, cells: [twoLines()], heightPt: 30 });
    expect(after.rows[1]).toBe(untouched);
    expect(before.rows[0]).toEqual({ index: 0, cells: [twoLines()], sourceStyleName: "ro1" });
  });

  it("ignores an unknown sheet id and an empty range list", () => {
    const model = book([{ index: 0, cells: [twoLines()] }], { styles });
    const sheets = model.sheets;
    refitRows(model, "no-such-sheet", [all], MEASURER, CTX);
    refitRows(model, model.sheets[0]!.id, [], MEASURER, CTX);
    expect(model.sheets).toBe(sheets);
  });

  it("given the rows before the command, refits only the rows it replaced or added — none when it changed nothing", () => {
    // Heights a file gave its rows (16.5 pt), not what the measurer finds (one 20 px line: 15 pt).
    const kept: Row = { index: 0, cells: [{ col: 0, value: "x" }], heightPt: 16.5 };
    const edited: Row = { index: 1, cells: [{ col: 0, value: "y" }], heightPt: 16.5 };
    const model = book([kept, edited], { styles });
    const before = model.sheets[0]!.rows;
    const sheets = model.sheets;
    refitRows(model, model.sheets[0]!.id, [all], MEASURER, CTX, before);
    expect(model.sheets).toBe(sheets);
    // A command replaced row 1 and added row 3 (a wrapped two-line text).
    const replaced: Row = { ...edited, cells: [{ col: 0, value: "z" }] };
    const added: Row = { index: 3, cells: [twoLines()] };
    model.sheets = [{ ...model.sheets[0]!, rows: [kept, replaced, added] }];
    refitRows(model, model.sheets[0]!.id, [all], MEASURER, CTX, before);
    expect(model.sheets[0]!.rows).toEqual([kept, { index: 1, cells: [{ col: 0, value: "z" }] }, { ...added, heightPt: 30 }]);
    expect(model.sheets[0]!.rows[0]).toBe(kept);
  });

  it(`refits the first ${AUTOFIT_MAX_ROWS_PER_COMMAND} rows of one command only`, () => {
    expect(AUTOFIT_MAX_ROWS_PER_COMMAND).toBe(10_000);
    const rows: Row[] = Array.from({ length: AUTOFIT_MAX_ROWS_PER_COMMAND + 2 }, (_, index) => ({ index: index * 2, cells: [twoLines()] }));
    const model = book(rows, { styles });
    refitRows(model, model.sheets[0]!.id, [all], MEASURER, CTX);
    const refitted = model.sheets[0]!.rows;
    expect(refitted.filter((row) => row.heightPt === 30)).toHaveLength(AUTOFIT_MAX_ROWS_PER_COMMAND);
    expect(refitted[AUTOFIT_MAX_ROWS_PER_COMMAND - 1]!.heightPt).toBe(30);
    expect(refitted[AUTOFIT_MAX_ROWS_PER_COMMAND]).toBe(rows[AUTOFIT_MAX_ROWS_PER_COMMAND]);
    expect(refitted[AUTOFIT_MAX_ROWS_PER_COMMAND + 1]).toBe(rows[AUTOFIT_MAX_ROWS_PER_COMMAND + 1]);
  });
});
