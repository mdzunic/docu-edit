import { describe, expect, it, vi } from "vitest";
import type { GridRange } from "./address";
import { clearContents, PROTECTED_CELL_MESSAGE, ARRAY_PART_MESSAGE } from "./cells";
import {
  applyNamedStyle,
  applyPattern,
  applyStylePatch,
  applyStylePatches,
  borderPresetPatches,
  capturePattern,
  clearRange,
  FORMAT_TOO_LARGE_MESSAGE,
  MAX_FORMAT_MATERIALISED,
  patchStyle,
  type StylePatch,
} from "./formatting";
import { MAX_COLS, MAX_ROWS, type Cell, type CellBorderEdge, type CellStyle, type ColumnSpan, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { createBlankWorkbook } from "./sheets";
import * as styles from "./styles";

vi.mock("./styles", async (importOriginal) => {
  const original = await importOriginal<typeof import("./styles")>();
  return { ...original, internStyle: vi.fn(original.internStyle) };
});

const SHEET = "sheet-1";

function book(rows: Row[], options: { styles?: CellStyle[]; sheet?: Partial<Sheet> } = {}): SpreadsheetModel {
  const model = createBlankWorkbook();
  model.sheets = [{ ...model.sheets[0]!, id: SHEET, rows, ...options.sheet }, { ...model.sheets[0]!, id: "other", name: "Other", rows: [] }];
  model.styles = options.styles ?? [{}];
  for (const sheet of model.sheets) {
    delete sheet.dirty;
  }
  return model;
}

const sheetOf = (model: SpreadsheetModel): Sheet => model.sheets[0]!;
const rowAt = (model: SpreadsheetModel, index: number): Row | undefined => sheetOf(model).rows.find((row) => row.index === index);
const cellAt = (model: SpreadsheetModel, row: number, col: number): Cell | undefined => rowAt(model, row)?.cells.find((cell) => cell.col === col);
const range = (top: number, left: number, bottom = top, right = left): GridRange => ({ top, left, bottom, right });
const column = (left: number, right = left): GridRange => ({ top: 0, left, bottom: MAX_ROWS - 1, right });
const rows = (top: number, bottom = top): GridRange => ({ top, left: 0, bottom, right: MAX_COLS - 1 });
const WHOLE: GridRange = { top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 };
const BOLD: StylePatch = { font: { bold: true } };
/** The style a position shows (cell → row → column → 0). */
const shown = (model: SpreadsheetModel, row: number, col: number): CellStyle => styles.cellStyle(model, styles.effectiveStyleId(sheetOf(model), row, col));
const storedCells = (model: SpreadsheetModel): number => sheetOf(model).rows.reduce((sum, row) => sum + row.cells.length, 0);

describe("SPEC-123 patchStyle", () => {
  it("is pure, merges key by key and returns the input itself when nothing changes", () => {
    const style: CellStyle = { font: { sizePt: 11, bold: true }, sourceXf: 3 };
    const frozen = structuredClone(style);
    expect(patchStyle(style, { font: { bold: true } })).toBe(style);
    expect(patchStyle(style, {})).toBe(style);
    const next = patchStyle(style, { font: { italic: true } });
    expect(next).toEqual({ font: { sizePt: 11, bold: true, italic: true } });
    expect(style).toEqual(frozen);
  });

  it("removes a field or a whole aspect with null and drops emptied aspects", () => {
    const style: CellStyle = { font: { bold: true }, fill: { kind: "solid", color: "#ff0000" }, alignment: { wrap: true, horizontal: "left" } };
    expect(patchStyle(style, { font: { bold: null } })).toEqual({ fill: style.fill, alignment: style.alignment });
    expect(patchStyle(style, { fill: null, alignment: null })).toEqual({ font: { bold: true } });
    expect(patchStyle(style, { alignment: { wrap: null } })).toEqual({ font: { bold: true }, fill: style.fill, alignment: { horizontal: "left" } });
    expect(patchStyle({ alignment: { indent: 2 } }, { alignment: { indent: 0 } })).toEqual({});
    expect(patchStyle({ borders: { top: { style: "thin" } } }, { borders: { top: null } })).toEqual({});
  });

  it("never carries sourceXf/sourceStyleName on a changed result", () => {
    const changed = patchStyle({ font: { bold: true }, sourceXf: 4, sourceStyleName: "ce3" }, { fill: { kind: "solid", color: "#00ff00" } });
    expect(changed).toEqual({ font: { bold: true }, fill: { kind: "solid", color: "#00ff00" } });
  });

  it("clears a built-in id with a plain code and treats General as no format", () => {
    const dated: CellStyle = { numberFormat: "m/d/yyyy", builtinFormatId: 14 };
    expect(patchStyle(dated, { numberFormat: "0.00" })).toEqual({ numberFormat: "0.00" });
    expect(patchStyle(dated, { numberFormat: "General" })).toEqual({});
    expect(patchStyle({}, { numberFormat: "m/d/yyyy", builtinFormatId: 14 })).toEqual(dated);
    expect(patchStyle(dated, { numberFormat: null })).toEqual({});
  });

  it("replaces border edges one by one and keeps namedStyle under direct formatting", () => {
    const edge: CellBorderEdge = { style: "thin" };
    const style: CellStyle = { borders: { top: edge, diagonal: { style: "thin", up: true } }, namedStyle: "Good" };
    expect(patchStyle(style, { borders: { bottom: { style: "double" }, top: null } })).toEqual({
      borders: { bottom: { style: "double" }, diagonal: { style: "thin", up: true } },
      namedStyle: "Good",
    });
  });
});

describe("SPEC-123 applyStylePatch — model invariants", () => {
  it("copies on write: a new rows array, new objects only for what changed, untouched rows keep identity", () => {
    const row0: Row = { index: 0, cells: [{ col: 0, value: 1 }, { col: 5, value: "x" }] };
    const row3: Row = { index: 3, cells: [{ col: 0, value: 2 }] };
    const model = book([row0, row3]);
    const before = sheetOf(model);
    expect(applyStylePatch(model, SHEET, [range(0, 0)], BOLD)).toEqual({ ok: true });
    expect(sheetOf(model)).not.toBe(before);
    expect(sheetOf(model).rows).not.toBe(before.rows);
    expect(rowAt(model, 3)).toBe(row3);
    expect(rowAt(model, 0)).not.toBe(row0);
    expect(rowAt(model, 0)!.cells[1]).toBe(row0.cells[1]);
    expect(row0.cells[0]).toEqual({ col: 0, value: 1 });
    expect(shown(model, 0, 0).font).toEqual({ bold: true });
  });

  it("marks only the edited sheet dirty and leaves the model untouched when nothing changes", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }] }], { styles: [{}, { font: { bold: true } }] });
    const sheets = model.sheets;
    applyStylePatch(model, SHEET, [range(0, 0)], BOLD);
    expect(model.sheets).toBe(sheets);
    applyStylePatch(model, SHEET, [range(0, 1)], BOLD);
    expect(sheetOf(model).dirty).toBe(true);
    expect(model.sheets[1]!.dirty).toBeUndefined();
  });

  it("stores an empty position only when its patched style differs from what it resolves to", () => {
    const model = book([], { styles: [{}, { font: { bold: true } }], sheet: { columns: [{ min: 2, max: 2, styleId: 1 }] } });
    applyStylePatch(model, SHEET, [range(0, 1, 1, 2)], BOLD);
    // B1:B2 materialise; C1:C2 already show bold through their column.
    expect(cellAt(model, 0, 1)).toEqual({ col: 1, value: null, styleId: 1 });
    expect(cellAt(model, 0, 2)).toBeUndefined();
    expect(storedCells(model)).toBe(2);
  });

  it("writes styleId explicitly when non-zero, and 0 only over a non-zero fallback", () => {
    const model = book([{ index: 0, styleId: 1, cells: [{ col: 0, value: "a" }] }, { index: 1, cells: [{ col: 0, value: "b", styleId: 1 }] }], {
      styles: [{}, { font: { bold: true } }],
    });
    applyStylePatch(model, SHEET, [range(0, 0, 1, 0)], { font: { bold: null } });
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: "a", styleId: 0 });
    expect(cellAt(model, 1, 0)).toEqual({ col: 0, value: "b" });
  });

  it("interns once per base style id within one command", () => {
    const intern = vi.mocked(styles.internStyle);
    const cells: Cell[] = Array.from({ length: 1000 }, (_, col) => ({ col, value: col }));
    const many: Row[] = Array.from({ length: 200 }, (_, index) => ({ index, cells }));
    const model = book(many);
    intern.mockClear();
    applyStylePatch(model, SHEET, [range(0, 0, 199, 999)], BOLD);
    expect(intern).toHaveBeenCalledTimes(1);
    expect(model.styles).toHaveLength(2);
    expect(shown(model, 199, 999).font).toEqual({ bold: true });
  });

  it("applies a font patch to every rich-text run, dropping what was bound to the old runs", () => {
    const rich: Cell = { col: 0, value: "ab", runs: [{ text: "a", font: { italic: true } }, { text: "b" }], sstIndex: 4 };
    const model = book([{ index: 0, cells: [rich] }]);
    applyStylePatch(model, SHEET, [range(0, 0)], BOLD);
    expect(cellAt(model, 0, 0)).toEqual({
      col: 0,
      value: "ab",
      styleId: 1,
      runs: [{ text: "a", font: { italic: true, bold: true } }, { text: "b", font: { bold: true } }],
    });
  });

  it("refuses a range on a protected sheet holding locked cells, unless formatting is permitted for its class", () => {
    const locked = book([{ index: 0, cells: [{ col: 0, value: 1 }] }], { sheet: { protection: { sheet: true } } });
    const sheets = locked.sheets;
    expect(applyStylePatch(locked, SHEET, [range(0, 0)], BOLD)).toEqual({ ok: false, message: PROTECTED_CELL_MESSAGE });
    expect(locked.sheets).toBe(sheets);
    const cells = book([], { sheet: { protection: { sheet: true, allowFormatCells: true } } });
    expect(applyStylePatch(cells, SHEET, [range(0, 0)], BOLD).ok).toBe(true);
    expect(applyStylePatch(cells, SHEET, [column(1)], BOLD)).toEqual({ ok: false, message: PROTECTED_CELL_MESSAGE });
    const columns = book([], { sheet: { protection: { sheet: true, allowFormatColumns: true } } });
    expect(applyStylePatch(columns, SHEET, [column(1)], BOLD).ok).toBe(true);
    expect(applyStylePatch(columns, SHEET, [rows(1)], BOLD).ok).toBe(false);
    const rowsAllowed = book([], { sheet: { protection: { sheet: true, allowFormatRows: true } } });
    expect(applyStylePatch(rowsAllowed, SHEET, [rows(1)], BOLD).ok).toBe(true);
    // Unlocked cells format on any protected sheet.
    const unlocked = book([{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }] }], {
      styles: [{}, { protection: { locked: false } }],
      sheet: { protection: { sheet: true } },
    });
    expect(applyStylePatch(unlocked, SHEET, [range(0, 0)], BOLD).ok).toBe(true);
  });
});

describe("SPEC-123 range classes", () => {
  it("whole columns split and create spans and touch only stored rows: column C of a 40-row sheet makes at most 40 cells", () => {
    const stored: Row[] = Array.from({ length: 40 }, (_, i) => ({ index: i * 1000, cells: [{ col: 0, value: i }], ...(i % 2 ? { styleId: 2 } : {}) }));
    const model = book(stored, {
      styles: [{}, { fill: { kind: "solid", color: "#eeeeee" } }, { font: { italic: true } }],
      sheet: { columns: [{ min: 1, max: 3, styleId: 1, widthPx: 80 }] },
    });
    const before = storedCells(model);
    applyStylePatch(model, SHEET, [column(2)], BOLD);
    const columns = sheetOf(model).columns;
    expect(columns.map(({ min, max }) => [min, max])).toEqual([
      [1, 1],
      [2, 2],
      [3, 3],
    ]);
    expect(columns[0]!.styleId).toBe(1);
    expect(columns[2]!.styleId).toBe(1);
    expect(model.styles[columns[1]!.styleId!]).toEqual({ fill: { kind: "solid", color: "#eeeeee" }, font: { bold: true } });
    expect(columns[1]!.widthPx).toBe(80);
    // Only the 20 styled rows need a cell: the row still wins there.
    expect(storedCells(model) - before).toBe(20);
    expect(storedCells(model) - before).toBeLessThanOrEqual(40);
    expect(shown(model, 1000, 2).font).toEqual({ italic: true, bold: true });
    expect(shown(model, 500_000, 2).font).toEqual({ bold: true });
    expect(shown(model, 500_000, 3).font).toBeUndefined();
  });

  it("a whole column with no span becomes a styled span", () => {
    const model = book([]);
    applyStylePatch(model, SHEET, [column(4, 5)], BOLD);
    expect(sheetOf(model).columns).toEqual([{ min: 4, max: 5, styleId: 1 }]);
    expect(sheetOf(model).rows).toEqual([]);
  });

  it("whole rows set Row.styleId and materialise the styled spans' positions", () => {
    const model = book([{ index: 2, cells: [{ col: 0, value: "a" }] }], {
      styles: [{}, { fill: { kind: "solid", color: "#eeeeee" } }],
      sheet: { columns: [{ min: 3, max: 4, styleId: 1 }] },
    });
    applyStylePatch(model, SHEET, [rows(2, 3)], BOLD);
    const bold = model.styles.findIndex((style) => styles.styleKey(style) === styles.styleKey({ font: { bold: true } }));
    expect(rowAt(model, 2)!.styleId).toBe(bold);
    expect(rowAt(model, 3)!.styleId).toBe(bold);
    expect(rowAt(model, 3)!.cells.map((cell) => cell.col)).toEqual([3, 4]);
    expect(shown(model, 3, 3)).toEqual({ fill: { kind: "solid", color: "#eeeeee" }, font: { bold: true } });
    expect(shown(model, 3, 100).font).toEqual({ bold: true });
    expect(shown(model, 4, 3).font).toBeUndefined();
  });

  it("a styled row keeps its own style id rewritten, materialising nothing", () => {
    const model = book([{ index: 0, styleId: 1, cells: [] }], { styles: [{}, { font: { italic: true } }], sheet: { columns: [{ min: 0, max: 0, styleId: 1 }] } });
    applyStylePatch(model, SHEET, [rows(0)], BOLD);
    expect(rowAt(model, 0)).toEqual({ index: 0, styleId: 2, cells: [] });
  });

  it("the whole sheet patches every span, every styled row and every stored cell, materialising nothing", () => {
    const model = book(
      [
        { index: 0, styleId: 1, cells: [{ col: 0, value: 1 }] },
        { index: 5, cells: [{ col: 2, value: 2, styleId: 2 }, { col: 9, value: 3 }] },
      ],
      { styles: [{}, { font: { italic: true } }, { fill: { kind: "solid", color: "#ff0000" } }], sheet: { columns: [{ min: 3, max: 3, widthPx: 90 }] } },
    );
    const before = storedCells(model);
    applyStylePatch(model, SHEET, [WHOLE], BOLD);
    expect(storedCells(model)).toBe(before);
    const columns = sheetOf(model).columns;
    expect(columns[0]!.min).toBe(0);
    expect(columns.at(-1)!.max).toBe(MAX_COLS - 1);
    for (const [row, col] of [
      [0, 0],
      [0, 7],
      [5, 2],
      [5, 9],
      [700_000, 3],
      [700_000, 16_000],
    ] as const) {
      expect(shown(model, row, col).font?.bold).toBe(true);
    }
    expect(shown(model, 5, 2).fill).toEqual({ kind: "solid", color: "#ff0000" });
    expect(columns.find((span) => span.min === 3)!.widthPx).toBe(90);
  });

  it("refuses past MAX_FORMAT_MATERIALISED before touching the model", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: 1 }] }]);
    const sheets = model.sheets;
    const styleTable = model.styles;
    const huge = range(0, 0, 9999, 100);
    expect((huge.bottom + 1) * (huge.right + 1)).toBeGreaterThan(MAX_FORMAT_MATERIALISED);
    expect(applyStylePatch(model, SHEET, [huge], BOLD)).toEqual({ ok: false, message: FORMAT_TOO_LARGE_MESSAGE });
    expect(model.sheets).toBe(sheets);
    expect(model.styles).toBe(styleTable);
    // A million rows under a styled column would materialise a cell each.
    const styled = book([], { styles: [{}, { font: { italic: true } }], sheet: { columns: [{ min: 0, max: 1, styleId: 1 }] } });
    expect(applyStylePatch(styled, SHEET, [rows(0, 600_000)], BOLD)).toEqual({ ok: false, message: FORMAT_TOO_LARGE_MESSAGE });
    // The same rows without styled spans cost one Row each — 600,001 fits.
    const plain = book([]);
    expect(applyStylePatch(plain, SHEET, [rows(0, 600_000)], BOLD).ok).toBe(true);
    expect(sheetOf(plain).rows).toHaveLength(600_001);
  });

  it("overlapping ranges are harmless", () => {
    const model = book([]);
    applyStylePatch(model, SHEET, [range(0, 0, 3, 3), range(2, 2, 5, 5), column(3)], BOLD);
    // A1:D4 and C3:F6 materialise 16 + 12 cells; column D's span then makes
    // D1:D6's style-only cells redundant, and they go.
    expect(storedCells(model)).toBe(16 + 12 - 6);
    expect(shown(model, 4, 4).font).toEqual({ bold: true });
    expect(cellAt(model, 0, 3)).toBeUndefined();
  });
});

describe("SPEC-123 borders", () => {
  const thin: CellBorderEdge = { style: "thin" };

  it("presets set the edges the table names, Thick outside drawing medium", () => {
    const model = book([]);
    applyStylePatches(model, SHEET, borderPresetPatches(range(1, 1, 3, 3), "outside", thin));
    expect(shown(model, 1, 1).borders).toEqual({ top: thin, left: thin });
    expect(shown(model, 2, 2).borders).toBeUndefined();
    expect(shown(model, 3, 3).borders).toEqual({ bottom: thin, right: thin });
    applyStylePatches(model, SHEET, borderPresetPatches(range(5, 5), "thick-outside", thin));
    const medium = { style: "medium" };
    expect(shown(model, 5, 5).borders).toEqual({ top: medium, right: medium, bottom: medium, left: medium });
    applyStylePatches(model, SHEET, borderPresetPatches(range(1, 1, 3, 3), "all", thin));
    expect(shown(model, 2, 2).borders).toEqual({ top: thin, right: thin, bottom: thin, left: thin });
    applyStylePatches(model, SHEET, borderPresetPatches(range(1, 1, 3, 3), "bottom", { style: "double" }));
    expect(shown(model, 3, 2).borders?.bottom).toEqual({ style: "double" });
    expect(shown(model, 2, 2).borders?.bottom).toEqual(thin);
  });

  it("No border removes every edge but the diagonals", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }] }], {
      styles: [{}, { borders: { top: thin, left: thin, diagonal: { style: "thin", up: true } } }],
    });
    applyStylePatches(model, SHEET, borderPresetPatches(range(0, 0), "none", thin));
    expect(shown(model, 0, 0).borders).toEqual({ diagonal: { style: "thin", up: true } });
  });

  it("clears the facing edge of a stored neighbour, never materialising one", () => {
    const model = book([{ index: 3, cells: [{ col: 1, value: "below", styleId: 1 }] }], {
      styles: [{}, { borders: { top: { style: "dashed" } } }],
      sheet: { columns: [] },
    });
    const entries = borderPresetPatches(range(0, 1, 2, 2), "bottom", thin);
    expect(entries.filter((entry) => entry.storedOnly)).toEqual([
      { range: range(3, 1, 3, 2), patch: { borders: { top: null } }, storedOnly: true },
    ]);
    applyStylePatches(model, SHEET, entries);
    expect(shown(model, 3, 1).borders).toBeUndefined();
    expect(cellAt(model, 3, 2)).toBeUndefined();
    expect(shown(model, 2, 1).borders).toEqual({ bottom: thin });
  });

  it("Outside on a whole column puts left/right on the span, top on row 1 and bottom on row 1,048,576", () => {
    const entries = borderPresetPatches(column(2), "outside", thin);
    const model = book([]);
    applyStylePatches(model, SHEET, entries);
    expect(sheetOf(model).columns).toHaveLength(1);
    expect(model.styles[sheetOf(model).columns[0]!.styleId!]!.borders).toEqual({ left: thin, right: thin });
    expect(shown(model, 0, 2).borders).toEqual({ left: thin, right: thin, top: thin });
    expect(shown(model, MAX_ROWS - 1, 2).borders).toEqual({ left: thin, right: thin, bottom: thin });
    expect(storedCells(model)).toBe(2);
  });
});

describe("SPEC-123 clearRange", () => {
  it("Clear contents is SPEC-122's Delete", () => {
    const rowsOf = (): Row[] => [{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }, { col: 1, value: "x", formula: "A1" }] }];
    const viaClear = book(rowsOf(), { styles: [{}, { font: { bold: true } }] });
    const viaDelete = book(rowsOf(), { styles: [{}, { font: { bold: true } }] });
    expect(clearRange(viaClear, SHEET, [range(0, 0, 0, 1)], "contents")).toEqual({ ok: true });
    clearContents(viaDelete, 0, [range(0, 0, 0, 1)]);
    expect(sheetOf(viaClear).rows).toEqual(sheetOf(viaDelete).rows);
  });

  it("Clear formats removes style ids, flattens runs and keeps the value", () => {
    const model = book(
      [
        { index: 0, styleId: 1, cells: [{ col: 0, value: "ab", styleId: 1, runs: [{ text: "a", font: { bold: true } }, { text: "b" }], sstIndex: 2 }] },
        { index: 1, cells: [{ col: 0, value: null, styleId: 1 }] },
      ],
      { styles: [{}, { font: { bold: true } }] },
    );
    clearRange(model, SHEET, [range(0, 0, 1, 0)], "formats");
    expect(sheetOf(model).rows).toEqual([{ index: 0, styleId: 1, cells: [{ col: 0, value: "ab" }] }]);
    clearRange(model, SHEET, [rows(0)], "formats");
    expect(sheetOf(model).rows).toEqual([{ index: 0, cells: [{ col: 0, value: "ab" }] }]);
  });

  it("Clear formats on whole columns removes the spans' styles; the whole sheet removes everything", () => {
    const model = book([{ index: 0, styleId: 1, cells: [{ col: 2, value: 1, styleId: 1 }] }], {
      styles: [{}, { font: { bold: true } }],
      sheet: { columns: [{ min: 1, max: 3, styleId: 1, widthPx: 70 }] },
    });
    clearRange(model, SHEET, [column(2)], "formats");
    expect(sheetOf(model).columns).toEqual([
      { min: 1, max: 1, styleId: 1, widthPx: 70 },
      { min: 2, max: 2, widthPx: 70 },
      { min: 3, max: 3, styleId: 1, widthPx: 70 },
    ]);
    expect(cellAt(model, 0, 2)).toEqual({ col: 2, value: 1 });
    clearRange(model, SHEET, [WHOLE], "formats");
    expect(sheetOf(model).columns.every((span) => span.styleId === undefined)).toBe(true);
    expect(rowAt(model, 0)!.styleId).toBeUndefined();
  });

  it("Clear all does both, and refuses part of an array like Delete", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }] }], { styles: [{}, { font: { bold: true } }] });
    clearRange(model, SHEET, [range(0, 0)], "all");
    expect(sheetOf(model).rows).toEqual([]);
    const array = book([{ index: 0, cells: [{ col: 0, value: 1, formula: "{1,2}", arrayRange: range(0, 0, 0, 1) }, { col: 1, value: 2 }] }]);
    expect(clearRange(array, SHEET, [range(0, 1)], "contents")).toEqual({ ok: false, message: ARRAY_PART_MESSAGE });
  });
});

describe("SPEC-123 applyNamedStyle", () => {
  it("replaces exactly the included aspects, sets namedStyle and appends a built-in on first use", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }] }], {
      styles: [{ font: { family: "Calibri", sizePt: 11 } }, { font: { bold: true, sizePt: 20 }, numberFormat: "0.00", alignment: { wrap: true } }],
    });
    expect(applyNamedStyle(model, SHEET, [range(0, 0)], "Good")).toEqual({ ok: true });
    expect(shown(model, 0, 0)).toEqual({
      font: { family: "Calibri", sizePt: 11, color: "#006100" },
      fill: { kind: "solid", color: "#c6efce" },
      numberFormat: "0.00",
      alignment: { wrap: true },
      namedStyle: "Good",
    });
    expect(model.namedStyles).toEqual([{ name: "Good", builtinId: 26, includes: ["font", "fill"], style: styles.builtinCellStyle("Good")!.style }]);
    applyNamedStyle(model, SHEET, [range(1, 0)], "Good");
    expect(model.namedStyles).toHaveLength(1);
  });

  it("renders and applies with the workbook's own definition, Normal being Clear formats", () => {
    const model = book([{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }] }], { styles: [{}, { font: { italic: true } }] });
    model.namedStyles = [{ name: "Good", builtinId: 26, includes: ["fill"], style: { fill: { kind: "solid", color: "#123456" } }, sourceXfId: 1 }];
    applyNamedStyle(model, SHEET, [range(0, 0)], "Good");
    expect(shown(model, 0, 0)).toEqual({ font: { italic: true }, fill: { kind: "solid", color: "#123456" }, namedStyle: "Good" });
    expect(model.namedStyles).toHaveLength(1);
    applyNamedStyle(model, SHEET, [range(0, 0)], "Normal");
    expect(cellAt(model, 0, 0)).toEqual({ col: 0, value: 1 });
    expect(applyNamedStyle(model, SHEET, [range(0, 0)], "Nope").ok).toBe(false);
  });

  it("direct formatting keeps namedStyle", () => {
    const model = book([]);
    applyNamedStyle(model, SHEET, [range(0, 0)], "Heading 4");
    applyStylePatch(model, SHEET, [range(0, 0)], { font: { italic: true } });
    expect(shown(model, 0, 0).namedStyle).toBe("Heading 4");
    expect(shown(model, 0, 0).font).toMatchObject({ bold: true, italic: true, color: { theme: "dk2" } });
  });
});

describe("SPEC-123 format painter patterns", () => {
  const model = (): SpreadsheetModel =>
    book(
      [
        { index: 0, cells: [{ col: 0, value: 1, styleId: 1 }, { col: 1, value: 2, styleId: 2 }] },
        { index: 1, cells: [{ col: 0, value: 3, styleId: 2 }] },
      ],
      { styles: [{}, { font: { bold: true } }, { font: { italic: true } }], sheet: { columns: [{ min: 6, max: 6, styleId: 1 }] } },
    );

  it("captures each position's style id", () => {
    expect(capturePattern(sheetOf(model()), range(0, 0, 1, 1))).toEqual({ rows: 2, cols: 2, ids: [1, 2, 2, 0] });
    expect(capturePattern(sheetOf(model()), column(5, 6))).toEqual({ rows: 1, cols: 2, ids: [0, 1], axis: "columns" });
    expect(capturePattern(sheetOf(model()), range(0, 0, 9999, 100))).toBeNull();
  });

  it("a single-cell target grows to the source's size; a larger one repeats; a smaller one takes what fits", () => {
    const grow = model();
    const pattern = capturePattern(sheetOf(grow), range(0, 0, 1, 1))!;
    applyPattern(grow, SHEET, range(4, 4), pattern);
    expect([shown(grow, 4, 4), shown(grow, 4, 5), shown(grow, 5, 4), shown(grow, 5, 5)]).toEqual([grow.styles[1], grow.styles[2], grow.styles[2], {}]);
    const repeat = model();
    applyPattern(repeat, SHEET, range(10, 0, 13, 0), capturePattern(sheetOf(repeat), range(0, 0, 1, 0))!);
    expect([10, 11, 12, 13].map((row) => styles.effectiveStyleId(sheetOf(repeat), row, 0))).toEqual([1, 2, 1, 2]);
    const fits = model();
    applyPattern(fits, SHEET, range(20, 0, 20, 1), capturePattern(sheetOf(fits), range(0, 0, 1, 1))!);
    expect([styles.effectiveStyleId(sheetOf(fits), 20, 0), styles.effectiveStyleId(sheetOf(fits), 20, 1)]).toEqual([1, 2]);
    expect(rowAt(fits, 21)).toBeUndefined();
  });

  it("a whole-column source paints whole columns through their spans", () => {
    const book1 = model();
    applyPattern(book1, SHEET, range(3, 9), capturePattern(sheetOf(book1), column(6))!);
    expect(sheetOf(book1).columns.find((span: ColumnSpan) => span.min === 9)).toEqual({ min: 9, max: 9, styleId: 1 });
  });
});
