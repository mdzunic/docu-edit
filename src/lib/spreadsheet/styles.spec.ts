// SPEC-117 "Model additions" / "Colours": style resolution is one function —
// effectiveStyleId takes the cell's styleId (an explicit 0 included), else
// its row's, else the covering ColumnSpan's (a binary search), else 0 — and
// cellStyle falls back to the default style for any id it cannot index.
// INDEXED_COLORS is ECMA-376 §18.8.27's 64-entry legacy palette, lowercase.
import { describe, expect, it } from "vitest";
import { cellStyle, columnSpanAt, defaultStyles, effectiveStyleId, INDEXED_COLORS, internStyle, styleKey } from "./styles";
import type { CellStyle, ColumnSpan, Row, Sheet } from "./model";

function sheetWith(partial: Partial<Sheet> = {}): Sheet {
  return {
    id: "sheet-1",
    name: "Sheet1",
    visibility: "visible",
    kind: "grid",
    rows: [],
    columns: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    merges: [],
    ...partial,
  };
}

describe("SPEC-117 effectiveStyleId: cell, then row, then column, then 0", () => {
  const sheet = sheetWith({
    columns: [
      { min: 0, max: 0, widthPx: 30 }, // a width, no style
      { min: 1, max: 2, styleId: 5 },
      { min: 4, max: 4, styleId: 6, hidden: true },
    ],
    rows: [
      {
        index: 0,
        cells: [
          { col: 0, value: "a" }, // no styleId
          { col: 1, value: "b", styleId: 3 },
          { col: 2, value: null, styleId: 0 }, // explicit 0 beats the column
        ],
      },
      {
        index: 1,
        styleId: 7, // customFormat row
        cells: [
          { col: 1, value: 1, styleId: 2 },
          { col: 3, value: 2 },
          { col: 4, value: null, styleId: 0 },
        ],
      },
      { index: 2, cells: [], heightPt: 30 }, // a row with no styleId
    ],
  });

  it("takes the cell's own styleId first, an explicit 0 included", () => {
    expect(effectiveStyleId(sheet, 0, 1)).toBe(3);
    expect(effectiveStyleId(sheet, 1, 1)).toBe(2);
    expect(effectiveStyleId(sheet, 0, 2)).toBe(0); // not the column's 5
    expect(effectiveStyleId(sheet, 1, 4)).toBe(0); // not the row's 7 nor the column's 6
  });

  it("else the row's styleId, over the column's", () => {
    expect(effectiveStyleId(sheet, 1, 3)).toBe(7); // a stored cell without a style
    expect(effectiveStyleId(sheet, 1, 2)).toBe(7); // no stored cell, column 2 styled 5
    expect(effectiveStyleId(sheet, 1, 10_000)).toBe(7);
  });

  it("else the column span covering the column", () => {
    expect(effectiveStyleId(sheet, 2, 1)).toBe(5);
    expect(effectiveStyleId(sheet, 2, 2)).toBe(5);
    expect(effectiveStyleId(sheet, 9, 4)).toBe(6); // an unstored row
    expect(effectiveStyleId(sheet, 0, 4)).toBe(6);
  });

  it("else 0 — including a span that carries only a width, and a cell with no styleId", () => {
    expect(effectiveStyleId(sheet, 0, 0)).toBe(0);
    expect(effectiveStyleId(sheet, 2, 0)).toBe(0);
    expect(effectiveStyleId(sheet, 2, 3)).toBe(0);
    expect(effectiveStyleId(sheet, 500, 500)).toBe(0);
    expect(effectiveStyleId(sheetWith(), 0, 0)).toBe(0);
  });

  it("finds the covering span among many — the same answer a scan gives", () => {
    const columns: ColumnSpan[] = [];
    for (let k = 0; k < 2000; k += 1) {
      // spans of 1–3 columns with gaps of 0–2, styled k + 1
      const min = columns.length === 0 ? 0 : columns.at(-1)!.max + 1 + (k % 3);
      columns.push({ min, max: min + (k % 3), styleId: k + 1 });
    }
    const big = sheetWith({ columns });
    const last = columns.at(-1)!.max;
    for (let col = 0; col <= last + 2; col += 1) {
      const scanned = columns.find((span) => span.min <= col && col <= span.max)?.styleId ?? 0;
      if (effectiveStyleId(big, 3, col) !== scanned) {
        expect.fail(`column ${col}: ${effectiveStyleId(big, 3, col)}, a scan says ${scanned}`);
      }
    }
  });

  it("resolves through rows and cells by binary search on large sheets", () => {
    const rows: Row[] = Array.from({ length: 5000 }, (_, k) => ({
      index: k * 3,
      styleId: k % 2 === 0 ? 11 : undefined,
      cells: [{ col: k, value: k, styleId: 20 + (k % 4) }],
    })).map((row) => (row.styleId === undefined ? { index: row.index, cells: row.cells } : row));
    const big = sheetWith({ rows, columns: [{ min: 0, max: 16_383, styleId: 9 }] });
    expect(effectiveStyleId(big, 300, 100)).toBe(20); // row 300 = k 100, its cell
    expect(effectiveStyleId(big, 300, 101)).toBe(11); // row style (k even)
    expect(effectiveStyleId(big, 303, 5)).toBe(9); // k 101 odd: no row style → column
    expect(effectiveStyleId(big, 301, 5)).toBe(9); // unstored row → column
  });
});

describe("SPEC-117 columnSpanAt: the sorted, non-overlapping spans", () => {
  const columns: ColumnSpan[] = [
    { min: 2, max: 4, widthPx: 10 },
    { min: 5, max: 5, widthPx: 20 },
    { min: 9, max: 12, widthPx: 30 },
  ];

  it("finds the span at its edges and inside", () => {
    expect(columnSpanAt(columns, 2)).toBe(columns[0]);
    expect(columnSpanAt(columns, 3)).toBe(columns[0]);
    expect(columnSpanAt(columns, 4)).toBe(columns[0]);
    expect(columnSpanAt(columns, 5)).toBe(columns[1]);
    expect(columnSpanAt(columns, 9)).toBe(columns[2]);
    expect(columnSpanAt(columns, 12)).toBe(columns[2]);
  });

  it("is undefined before, between and after the spans", () => {
    expect(columnSpanAt(columns, 0)).toBeUndefined();
    expect(columnSpanAt(columns, 1)).toBeUndefined();
    expect(columnSpanAt(columns, 6)).toBeUndefined();
    expect(columnSpanAt(columns, 8)).toBeUndefined();
    expect(columnSpanAt(columns, 13)).toBeUndefined();
    expect(columnSpanAt(columns, 16_383)).toBeUndefined();
  });

  it("is undefined with no spans", () => {
    expect(columnSpanAt([], 0)).toBeUndefined();
    expect(columnSpanAt(undefined, 0)).toBeUndefined();
  });
});

describe("SPEC-117 cellStyle: model.styles[id], or the default style", () => {
  const styles: CellStyle[] = [{}, { font: { bold: true } }, { fill: { kind: "solid", color: "#ffff00" } }];
  const model = { styles };

  it("returns the style at an in-range id", () => {
    expect(cellStyle(model, 0)).toBe(styles[0]);
    expect(cellStyle(model, 1)).toBe(styles[1]);
    expect(cellStyle(model, 2)).toBe(styles[2]);
  });

  it("falls back to styles[0] for an id it cannot index", () => {
    for (const id of [3, 99, -1, -0.5, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      expect(cellStyle(model, id)).toBe(styles[0]);
    }
  });

  it("the default style table is one empty Normal style", () => {
    expect(defaultStyles()).toEqual([{}]);
    expect(defaultStyles()).not.toBe(defaultStyles()); // a fresh table each time
    expect(cellStyle({ styles: defaultStyles() }, 4)).toEqual({});
  });
});

describe("SPEC-117 Colours: INDEXED_COLORS", () => {
  // The spec's table verbatim: each line is two runs of eight entries.
  const SPEC_TABLE = `
000000 ffffff ff0000 00ff00 0000ff ffff00 ff00ff 00ffff   000000 ffffff ff0000 00ff00 0000ff ffff00 ff00ff 00ffff
800000 008000 000080 808000 800080 008080 c0c0c0 808080   9999ff 993366 ffffcc ccffff 660066 ff8080 0066cc ccccff
000080 ff00ff ffff00 00ffff 800080 800000 008080 0000ff   00ccff ccffff ccffcc ffff99 99ccff ff99cc cc99ff ffcc99
3366ff 33cccc 99cc00 ffcc00 ff9900 ff6600 666699 969696   003366 339966 003300 333300 993300 993366 333399 333333
`;

  it("has 64 lowercase #rrggbb entries", () => {
    expect(INDEXED_COLORS).toHaveLength(64);
    for (const color of INDEXED_COLORS) {
      expect(color).toMatch(/^#[0-9a-f]{6}$/);
    }
  });

  it("matches the spec's table entry by entry", () => {
    const expected = SPEC_TABLE.trim()
      .split(/\s+/)
      .map((hex) => `#${hex}`);
    expect(expected).toHaveLength(64);
    expect([...INDEXED_COLORS]).toEqual(expected);
  });

  it("spot-checks well-known indices", () => {
    expect(INDEXED_COLORS[0]).toBe("#000000");
    expect(INDEXED_COLORS[1]).toBe("#ffffff");
    expect(INDEXED_COLORS[2]).toBe("#ff0000");
    expect(INDEXED_COLORS[8]).toBe("#000000");
    expect(INDEXED_COLORS[10]).toBe("#ff0000");
    expect(INDEXED_COLORS[13]).toBe("#ffff00");
    expect(INDEXED_COLORS[22]).toBe("#c0c0c0");
    expect(INDEXED_COLORS[23]).toBe("#808080");
    expect(INDEXED_COLORS[40]).toBe("#00ccff");
    expect(INDEXED_COLORS[55]).toBe("#969696");
    expect(INDEXED_COLORS[63]).toBe("#333333");
  });
});

describe("SPEC-122 internStyle: structural equality, copy-on-write appends", () => {
  it("returns an equal entry's index whatever the key order", () => {
    const model = { styles: [{}, { font: { bold: true as const, sizePt: 11 }, numberFormat: "0%" }] as CellStyle[] };
    const before = model.styles;
    expect(internStyle(model, { numberFormat: "0%", font: { sizePt: 11, bold: true } })).toBe(1);
    expect(model.styles).toBe(before);
  });

  it("ignores sourceXf and sourceStyleName, preferring a source-backed entry", () => {
    const model = { styles: [{}, { numberFormat: "0.00" }, { numberFormat: "0.00", sourceXf: 7 }] as CellStyle[] };
    expect(internStyle(model, { numberFormat: "0.00" })).toBe(2);
    expect(internStyle(model, { numberFormat: "0.00", sourceStyleName: "ce9" })).toBe(2);
    expect(styleKey({ numberFormat: "0.00", sourceXf: 3 })).toBe(styleKey({ numberFormat: "0.00" }));
  });

  it("appends a new style copy-on-write, without the source fields, and finds it again", () => {
    const first: CellStyle = {};
    const model = { styles: [first] as CellStyle[] };
    const before = model.styles;
    const id = internStyle(model, { alignment: { wrap: true }, sourceXf: 4 });
    expect(id).toBe(1);
    expect(model.styles).not.toBe(before);
    expect(before).toHaveLength(1);
    expect(model.styles[0]).toBe(first);
    expect(model.styles[1]).toEqual({ alignment: { wrap: true } });
    expect(internStyle(model, { alignment: { wrap: true } })).toBe(1);
    expect(model.styles).toHaveLength(2);
  });

  it("appends the style object itself when it names no source, and caches per styles array", () => {
    const model = { styles: [{}] as CellStyle[] };
    const style: CellStyle = { borders: { top: { style: "thin" } }, font: { italic: true } };
    expect(internStyle(model, style)).toBe(1);
    expect(model.styles[1]).toBe(style);
    const appended = model.styles;
    // Another array with the same entries has its own table and answers alike.
    const copy = { styles: [...appended] };
    expect(internStyle(copy, { font: { italic: true }, borders: { top: { style: "thin" } } })).toBe(1);
    expect(copy.styles).toHaveLength(2);
    expect(styleKey({ font: { sizePt: 11, bold: true } })).toBe(styleKey({ font: { bold: true, sizePt: 11 } }));
  });

  it("treats an absent field and an undefined one alike", () => {
    const model = { styles: [{}, { numberFormat: "@" }] as CellStyle[] };
    expect(internStyle(model, { numberFormat: "@", quotePrefix: undefined })).toBe(1);
    expect(internStyle(model, { quotePrefix: undefined })).toBe(0);
  });
});
