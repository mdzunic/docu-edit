// SPEC-123 cell formatting — the unit tests of its formatting layer:
// patchStyle's semantics; applyStylePatch's model invariants (copy-on-write,
// sparseness, explicit ids, interning once per base style, the edited sheet
// alone dirty), every range class and its materialisation bound, the
// MAX_FORMAT_MATERIALISED cap, protection and its format permits, rich-text
// runs; the border presets and their neighbour clearing; clearRange;
// applyNamedStyle; the format painter's patterns; and numberCodes.ts — the
// number-format presets in all six locales with the toolbar's
// currency/percent/comma buttons and the number-format shortcuts, every
// "Decimal stepping" example plus the accounting and General cases, and the
// Format Cells code builders' round trip over every category's whole option
// set in every locale.
import { describe, expect, it, vi } from "vitest";
import type { ThemeColorRef } from "../core/theme";
import type { GridRange } from "./address";
import { ARRAY_PART_MESSAGE, clearContents, PROTECTED_CELL_MESSAGE } from "./cells";
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
import { MAX_COLS, MAX_ROWS, type Cell, type CellBorderEdge, type CellFill, type CellStyle, type CellValue, type ColumnSpan, type Row, type Sheet, type SpreadsheetModel } from "./model";
import {
  buildNumberCode,
  commaButtonFormat,
  currencyButtonFormat,
  CURRENCY_SYMBOLS,
  currencySymbolPosition,
  dateTypeCodes,
  describeNumberCode,
  FRACTION_TYPE_CODES,
  isValidFormatCode,
  LOCALE_CURRENCY,
  NUMBER_CATEGORIES,
  NUMBER_FORMAT_PRESETS,
  numberFormatChoiceOf,
  percentButtonFormat,
  presetIdOf,
  presetNumberFormat,
  shortcutNumberFormat,
  stepDecimals,
  TIME_TYPE_CODES,
  type NegativeStyle,
  type NumberCodeSpec,
  type NumberFormatChoice,
  type NumberPresetId,
  type NumberShortcutId,
} from "./numberCodes";
import { createBlankWorkbook } from "./sheets";
import * as styles from "./styles";
import { BUILTIN_NUMBER_FORMATS, LOCALE_DEPENDENT_BUILTINS } from "./numberFormat/builtin";
import { SPREADSHEET_LOCALE_TAGS, SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "./numberFormat/locale";
import { FormatCodeError, formatCodeSections } from "./numberFormat/parse";
import { formatValue, type FormatContext } from "./numberFormat/render";

const EN_US = SPREADSHEET_LOCALES["en-US"];
const DE_DE = SPREADSHEET_LOCALES["de-DE"];

const ID_41 = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)';
const ID_42 = '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)';
const ID_43 = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const ID_44 = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
const EURO_ACCOUNTING = '_-* #,##0.00 "€"_-;-* #,##0.00 "€"_-;_-* "-"?? "€"_-;_-@_-';
const NO_SYMBOL_ACCOUNTING = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-';

const NEGATIVES: readonly NegativeStyle[] = ["minus", "red", "parens", "redParens"];
const DECIMALS = Array.from({ length: 31 }, (_, decimals) => decimals);

type PresetTable = Record<NumberPresetId, NumberFormatChoice>;

/** The rows every locale shares; `rest` adds the locale's own. */
function presets(rest: Pick<PresetTable, "currency" | "accounting" | "long-date" | "time">): PresetTable {
  return {
    general: { numberFormat: null },
    number: { numberFormat: "0.00" },
    "short-date": { numberFormat: "m/d/yyyy", builtinFormatId: 14 },
    percentage: { numberFormat: "0.00%" },
    fraction: { numberFormat: "# ?/?" },
    scientific: { numberFormat: "0.00E+00" },
    text: { numberFormat: "@" },
    ...rest,
  };
}

/** The Reference's preset table. */
const PRESETS: Record<SpreadsheetLocaleTag, PresetTable> = {
  "en-US": presets({
    currency: { numberFormat: '"$"#,##0.00' },
    accounting: { numberFormat: ID_44, builtinFormatId: 44 },
    "long-date": { numberFormat: "dddd, mmmm d, yyyy" },
    time: { numberFormat: "h:mm:ss AM/PM" },
  }),
  "en-GB": presets({
    currency: { numberFormat: '"£"#,##0.00' },
    accounting: { numberFormat: '_-"£"* #,##0.00_-;-"£"* #,##0.00_-;_-"£"* "-"??_-;_-@_-' },
    "long-date": { numberFormat: "dddd, d mmmm yyyy" },
    time: { numberFormat: "h:mm:ss" },
  }),
  "de-DE": presets({
    currency: { numberFormat: '#,##0.00 "€"' },
    accounting: { numberFormat: EURO_ACCOUNTING },
    "long-date": { numberFormat: "dddd, d. mmmm yyyy" },
    time: { numberFormat: "h:mm:ss" },
  }),
  "fr-FR": presets({
    currency: { numberFormat: '#,##0.00 "€"' },
    accounting: { numberFormat: EURO_ACCOUNTING },
    "long-date": { numberFormat: "dddd d mmmm yyyy" },
    time: { numberFormat: "h:mm:ss" },
  }),
  "es-ES": presets({
    currency: { numberFormat: '#,##0.00 "€"' },
    accounting: { numberFormat: EURO_ACCOUNTING },
    "long-date": { numberFormat: 'dddd, d "de" mmmm "de" yyyy' },
    time: { numberFormat: "h:mm:ss" },
  }),
  "cs-CZ": presets({
    currency: { numberFormat: '#,##0.00 "Kč"' },
    accounting: { numberFormat: '_-* #,##0.00 "Kč"_-;-* #,##0.00 "Kč"_-;_-* "-"?? "Kč"_-;_-@_-' },
    "long-date": { numberFormat: "dddd d. mmmm yyyy" },
    time: { numberFormat: "h:mm:ss" },
  }),
};

const PRESET_IDS = NUMBER_FORMAT_PRESETS.map((preset) => preset.id);

function context(tag: SpreadsheetLocaleTag): FormatContext {
  return { locale: SPREADSHEET_LOCALES[tag], dateSystem: "1900" };
}

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
/** A value as a Vue `ref` hands it back: wrapped in a Proxy. */
const reactiveish = <T extends object>(value: T): T => new Proxy(value, {});

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

  // The model holds plain data only: IndexedDB structured-clones it, and a
  // Vue reactive Proxy — what a value kept in a `ref` comes back as — throws
  // DataCloneError there (and threw in patchStyle's own clone, so a colour
  // picked in Format Cells was never applied).
  it("copies patch values handed in as a Proxy into plain data", () => {
    const color = reactiveish<ThemeColorRef>({ theme: "accent1", lumMod: 0.6, lumOff: 0.4 });
    const patched = patchStyle(
      { font: { bold: true } },
      { font: reactiveish({ color }), fill: reactiveish<CellFill>({ kind: "solid", color }), borders: reactiveish({ top: reactiveish<CellBorderEdge>({ style: "thin", color }) }) },
    );
    expect(() => structuredClone(patched)).not.toThrow();
    const plain = { theme: "accent1", lumMod: 0.6, lumOff: 0.4 };
    expect(patched).toEqual({ font: { bold: true, color: plain }, fill: { kind: "solid", color: plain }, borders: { top: { style: "thin", color: plain } } });
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

const clearedRows = (): Row[] => [{ index: 0, cells: [{ col: 0, value: 1, styleId: 1 }, { col: 1, value: "x", formula: "A1" }] }];

describe("SPEC-123 clearRange", () => {
  it("Clear contents is SPEC-122's Delete", () => {
    const viaClear = book(clearedRows(), { styles: [{}, { font: { bold: true } }] });
    const viaDelete = book(clearedRows(), { styles: [{}, { font: { bold: true } }] });
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

describe("SPEC-123 number-format presets per locale", () => {
  it("lists the dropdown's eleven entries in order", () => {
    expect(NUMBER_FORMAT_PRESETS).toEqual([
      { id: "general", label: "General" },
      { id: "number", label: "Number" },
      { id: "currency", label: "Currency" },
      { id: "accounting", label: "Accounting" },
      { id: "short-date", label: "Short date" },
      { id: "long-date", label: "Long date" },
      { id: "time", label: "Time" },
      { id: "percentage", label: "Percentage" },
      { id: "fraction", label: "Fraction" },
      { id: "scientific", label: "Scientific" },
      { id: "text", label: "Text" },
    ]);
  });

  it("pins each locale's currency: symbol, side and negatives", () => {
    expect(LOCALE_CURRENCY).toEqual({
      "en-US": { symbol: "$", position: "before", negativeParens: true },
      "en-GB": { symbol: "£", position: "before", negativeParens: false },
      "de-DE": { symbol: "€", position: "after", negativeParens: false },
      "fr-FR": { symbol: "€", position: "after", negativeParens: false },
      "es-ES": { symbol: "€", position: "after", negativeParens: false },
      "cs-CZ": { symbol: "Kč", position: "after", negativeParens: false },
    });
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: every preset's code and built-in id", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    for (const id of PRESET_IDS) {
      expect(presetNumberFormat(id, locale), id).toStrictEqual(PRESETS[tag][id]);
      const code = presetNumberFormat(id, locale).numberFormat;
      if (code !== null) {
        expect(isValidFormatCode(code), code).toBe(true);
      }
    }
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: the currency, percent and comma buttons", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    expect(currencyButtonFormat(locale)).toStrictEqual(PRESETS[tag].accounting);
    expect(percentButtonFormat()).toStrictEqual({ numberFormat: "0%" });
    expect(commaButtonFormat(locale)).toStrictEqual(
      tag === "en-US" ? { numberFormat: ID_43, builtinFormatId: 43 } : { numberFormat: NO_SYMBOL_ACCOUNTING },
    );
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: the number-format shortcuts", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    const expected: Record<NumberShortcutId, NumberFormatChoice> = {
      general: { numberFormat: null },
      number: { numberFormat: "#,##0.00" },
      time: { numberFormat: "h:mm AM/PM" },
      date: { numberFormat: "d-mmm-yy" },
      currency: tag === "en-US" ? { numberFormat: '"$"#,##0.00_);\\("$"#,##0.00\\)', builtinFormatId: 7 } : PRESETS[tag].currency,
      percent: { numberFormat: "0%" },
      scientific: { numberFormat: "0.00E+00" },
    };
    for (const [id, choice] of Object.entries(expected)) {
      expect(shortcutNumberFormat(id as NumberShortcutId, locale), id).toStrictEqual(choice);
    }
  });

  it("numberFormatChoiceOf: General is null; a locale-dependent built-in's code carries its id", () => {
    for (const general of [undefined, null, "", "  ", "General", "general", "GENERAL"]) {
      expect(numberFormatChoiceOf(general), String(general)).toStrictEqual({ numberFormat: null });
    }
    for (const id of LOCALE_DEPENDENT_BUILTINS) {
      const code = BUILTIN_NUMBER_FORMATS[id]!;
      expect(numberFormatChoiceOf(code), code).toStrictEqual({ numberFormat: code, builtinFormatId: id });
    }
    for (const id of [1, 2, 3, 4, 9, 10, 11, 12, 15, 18, 19, 21, 37, 45, 48, 49]) {
      const code = BUILTIN_NUMBER_FORMATS[id]!;
      expect(numberFormatChoiceOf(code), code).toStrictEqual({ numberFormat: code });
    }
    expect(numberFormatChoiceOf('"€"#,##0.0')).toStrictEqual({ numberFormat: '"€"#,##0.0' });
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: presetIdOf names the preset a style's format is exactly", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    for (const id of PRESET_IDS) {
      const { numberFormat, builtinFormatId } = presetNumberFormat(id, locale);
      const style = numberFormat === null ? {} : builtinFormatId === undefined ? { numberFormat } : { numberFormat, builtinFormatId };
      expect(presetIdOf(style, locale), id).toBe(id);
    }
    expect(presetIdOf({ numberFormat: "General" }, locale)).toBe("general");
    expect(presetIdOf({ numberFormat: "0.000" }, locale)).toBeNull();
    // Short date is the locale-dependent id 14, never the bare code.
    expect(presetIdOf({ numberFormat: "m/d/yyyy" }, locale)).toBeNull();
  });

  it("presetIdOf: Excel's id 44 is Accounting only where it is the locale's Accounting", () => {
    const id44 = { numberFormat: ID_44, builtinFormatId: 44 };
    expect(presetIdOf(id44, EN_US)).toBe("accounting");
    expect(presetIdOf(id44, SPREADSHEET_LOCALES["en-GB"])).toBeNull();
    expect(presetIdOf(id44, DE_DE)).toBeNull();
  });

  it("isValidFormatCode is parseFormatCode not throwing", () => {
    for (const code of ["0.00", '"$"#,##0.00_);("$"#,##0.00)', "General", "@", "yyyy-mm-dd", "[Red]0.0", ID_44]) {
      expect(isValidFormatCode(code), code).toBe(true);
    }
    for (const code of ['"abc', "0;0;0;0;0", "[Purple]0", "0@", "0E+", "[Red"]) {
      expect(isValidFormatCode(code), code).toBe(false);
    }
  });

  it("renders through SPEC-118: en-US Currency of 1234.5 is $1,234.50, 0.0% of 0.125 is 12.5%", () => {
    expect(formatValue(1234.5, presetNumberFormat("currency", EN_US).numberFormat!, context("en-US")).text).toBe("$1,234.50");
    expect(formatValue(0.125, "0.0%", context("en-US")).text).toBe("12.5%");
    expect(formatValue(1234.5, presetNumberFormat("currency", SPREADSHEET_LOCALES["en-GB"]).numberFormat!, context("en-GB")).text).toBe("£1,234.50");
    expect(formatValue(1234.5, presetNumberFormat("currency", DE_DE).numberFormat!, context("de-DE")).text).toBe("1.234,50 €");
    expect(formatValue(-1234.5, presetNumberFormat("currency", SPREADSHEET_LOCALES["cs-CZ"]).numberFormat!, context("cs-CZ")).text).toBe(
      "-1 234,50 Kč",
    );
  });
});

describe("SPEC-123 stepDecimals (every Decimal stepping example)", () => {
  it.each<[code: string, dir: 1 | -1, stepped: string]>([
    ["0", 1, "0.0"],
    ["0.0", -1, "0"],
    ["#,##0.00", 1, "#,##0.000"],
    ["0%", 1, "0.0%"],
    ["0.00E+00", 1, "0.000E+00"],
    ['"$"#,##0.00_);("$"#,##0.00)', 1, '"$"#,##0.000_);("$"#,##0.000)'],
    ["0", -1, "0"],
    ["yyyy-mm-dd", 1, "yyyy-mm-dd"],
    ["yyyy-mm-dd", -1, "yyyy-mm-dd"],
  ])("the Reference's example: %s stepped %i is %s", (code, dir, stepped) => {
    expect(stepDecimals(code, 1234.5, dir)).toBe(stepped);
  });

  it.each<[value: CellValue, up: string, down: string]>([
    [1234.5, "0.00", "0"],
    [0.125, "0.0000", "0.00"],
    [-2.75, "0.000", "0.0"],
    [5, "0.0", "0"],
    [0, "0.0", "0"],
    [123456789012, "0.000000", "0.0000"], // "1.23457E+11": the mantissa's 5
    [0.0000000001234, "0.0000", "0.00"], // "1.234E-10"
    ["12.5", "0.0", "0"],
    [true, "0.0", "0"],
    [null, "0.0", "0"],
  ])("General on %j: +1 is %s, −1 is %s (the decimals generalText shows)", (value, up, down) => {
    for (const general of [undefined, "", "General", "general", "GENERAL"]) {
      expect(stepDecimals(general, value, 1)).toBe(up);
      expect(stepDecimals(general, value, -1)).toBe(down);
    }
  });

  it("steps id 44 down to id 42 through the one-decimal code, and back up", () => {
    const oneDecimal = '_("$"* #,##0.0_);_("$"* \\(#,##0.0\\);_("$"* "-"?_);_(@_)';
    expect(stepDecimals(ID_44, null, -1)).toBe(oneDecimal);
    expect(stepDecimals(oneDecimal, null, -1)).toBe(ID_42);
    expect(stepDecimals(ID_42, null, 1)).toBe(oneDecimal);
    expect(stepDecimals(oneDecimal, null, 1)).toBe(ID_44);
    expect(stepDecimals(ID_42, null, -1)).toBe(ID_42);
    expect(stepDecimals(ID_43, null, -1)).toBe('_(* #,##0.0_);_(* \\(#,##0.0\\);_(* "-"?_);_(@_)');
    expect(stepDecimals(ID_41, null, 1)).toBe('_(* #,##0.0_);_(* \\(#,##0.0\\);_(* "-"?_);_(@_)');
  });

  it("steps the symbol-after accounting template, its dash and symbol untouched", () => {
    const one = '_-* #,##0.0 "€"_-;-* #,##0.0 "€"_-;_-* "-"? "€"_-;_-@_-';
    const none = '_-* #,##0 "€"_-;-* #,##0 "€"_-;_-* "-" "€"_-;_-@_-';
    expect(stepDecimals(EURO_ACCOUNTING, 1, -1)).toBe(one);
    expect(stepDecimals(one, 1, -1)).toBe(none);
    expect(stepDecimals(none, 1, 1)).toBe(one);
    expect(stepDecimals(EURO_ACCOUNTING, 1, 1)).toBe('_-* #,##0.000 "€"_-;-* #,##0.000 "€"_-;_-* "-"??? "€"_-;_-@_-');
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: steps every code the builders make to the one with a decimal more or less", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    const mismatches: string[] = [];
    for (const spec of everyOptionSet(tag)) {
      if (spec.decimals === undefined || spec.decimals === 30) {
        continue;
      }
      const fewer = buildNumberCode(spec, locale)!;
      const more = buildNumberCode({ ...spec, decimals: spec.decimals + 1 }, locale)!;
      if (stepDecimals(fewer, 0, 1) !== more) {
        mismatches.push(`${fewer} +1 → ${stepDecimals(fewer, 0, 1)}, not ${more}`);
      }
      if (stepDecimals(more, 0, -1) !== fewer) {
        mismatches.push(`${more} −1 → ${stepDecimals(more, 0, -1)}, not ${fewer}`);
      }
    }
    expect(mismatches).toEqual([]);
  });

  it.each<[code: string, dir: 1 | -1, stepped: string]>([
    ['"0.0" 0', 1, '"0.0" 0.0'],
    ['0.00 "0.00"', -1, '0.0 "0.00"'],
    ["\\0#_0*00", 1, "\\0#_0*00.0"],
    ['"$"#,##0.00_);\\("$"#,##0.00\\)', -1, '"$"#,##0.0_);\\("$"#,##0.0\\)'],
    ["[>=100]0.0;[Red]0", 1, "[>=100]0.00;[Red]0.0"],
    ["[Color10]0.00", -1, "[Color10]0.0"],
    ["[$€-407] #,##0.00", -1, "[$€-407] #,##0.0"],
    ["#,##0,", 1, "#,##0.0,"],
    ['0.00;-0.00;"zero";@', 1, '0.000;-0.000;"zero";@'],
    ["[<1]0.00;yyyy-mm-dd", 1, "[<1]0.000;yyyy-mm-dd"],
  ])("edits only placeholders and points: %s stepped %i is %s", (code, dir, stepped) => {
    expect(stepDecimals(code, 1, dir)).toBe(stepped);
  });

  it.each<[code: string, dir: 1 | -1, stepped: string]>([
    ["##0.0E+0", -1, "##0E+0"],
    ["##0.0E+0", 1, "##0.00E+0"],
    ["0E+00", 1, "0.0E+00"],
    ["0.0E+00", -1, "0E+00"],
  ])("counts a scientific code's mantissa only: %s stepped %i is %s", (code, dir, stepped) => {
    expect(stepDecimals(code, 1, dir)).toBe(stepped);
  });

  it.each<[code: string, dir: 1 | -1, stepped: string]>([
    ["0.", 1, "0.0"],
    ["0.", -1, "0."], // no decimal placeholder to remove
    [".00", -1, ".0"],
    [".0", -1, "0"], // the last decimal stays as the integer placeholder
    ["???", 1, "???.0"], // `?` pads only an accounting zero section's dash
    ['"-"??', 1, '"-"??.0'], // … in the third section
    ['0.00;"-"', 1, '0.000;"-"'],
    ['0.00;-0.00;"-"', 1, '0.000;-0.000;"-"?'],
    ["0.0#", 1, "0.0#0"],
  ])("edge: %s stepped %i is %s", (code, dir, stepped) => {
    expect(stepDecimals(code, 1, dir)).toBe(stepped);
  });

  it.each(["yyyy-mm-dd", "h:mm:ss", "[h]:mm:ss", "mmss.0", "# ?/?", "# ??/100", "@", '"text"', "General;-General"])(
    "leaves %s as it is (no number section)",
    (code) => {
      expect(stepDecimals(code, 1, 1)).toBe(code);
      expect(stepDecimals(code, 1, -1)).toBe(code);
    },
  );

  it.each(['"abc', "0;0;0;0;0", "0@", "[Purple]0.0", "0.0E+"])("returns the invalid code %s unchanged", (code) => {
    expect(stepDecimals(code, 1, 1)).toBe(code);
    expect(stepDecimals(code, 1, -1)).toBe(code);
  });

  it("edits at formatCodeSections' offsets: each section's kind and its tokens where they are written", () => {
    const sections = formatCodeSections('"$"0.0;[Red]\\(0\\);yyyy;@');
    expect(sections.map((section) => section.kind)).toEqual(["number", "number", "date", "text"]);
    expect(sections[0]!.tokens.map(({ token, offset }) => [token.type, offset])).toEqual([
      ["literal", 0],
      ["digit", 3],
      ["point", 4],
      ["digit", 5],
    ]);
    expect(sections[1]!.tokens.map(({ offset }) => offset)).toEqual([12, 14, 15]); // [Red] is no token
    expect(sections[2]!.tokens).toEqual([{ token: { type: "date", part: "year", length: 4 }, offset: 18 }]);
    expect(formatCodeSections("0.0E")[0]!.tokens.map(({ token }) => token.type)).toEqual(["digit", "point", "digit"]); // the era `E` left out
    expect(() => formatCodeSections('"abc')).toThrow(FormatCodeError);
  });
});

/** Every category's whole option set in `tag`, written out here rather than
 * taken from numberCodes.ts — Currency with its None included. */
function everyOptionSet(tag: SpreadsheetLocaleTag): NumberCodeSpec[] {
  const locale = SPREADSHEET_LOCALES[tag];
  const symbols = ["", ...CURRENCY_SYMBOLS];
  return [
    { category: "general" },
    ...DECIMALS.flatMap((decimals) =>
      [false, true].flatMap((thousands) => NEGATIVES.map((negative): NumberCodeSpec => ({ category: "number", decimals, thousands, negative }))),
    ),
    ...DECIMALS.flatMap((decimals) =>
      NEGATIVES.flatMap((negative) => symbols.map((symbol): NumberCodeSpec => ({ category: "currency", decimals, negative, symbol }))),
    ),
    ...DECIMALS.flatMap((decimals) => symbols.map((symbol): NumberCodeSpec => ({ category: "accounting", decimals, symbol }))),
    ...dateTypeCodes(locale).map((type): NumberCodeSpec => ({ category: "date", type })),
    ...TIME_TYPE_CODES.map((type): NumberCodeSpec => ({ category: "time", type })),
    ...DECIMALS.map((decimals): NumberCodeSpec => ({ category: "percentage", decimals })),
    ...FRACTION_TYPE_CODES.map((type): NumberCodeSpec => ({ category: "fraction", type })),
    ...DECIMALS.map((decimals): NumberCodeSpec => ({ category: "scientific", decimals })),
    { category: "text" },
  ];
}

/** en-US's Number code for these options. */
function numberCode(decimals: number, thousands: boolean, negative: NegativeStyle): string | undefined {
  return buildNumberCode({ category: "number", decimals, thousands, negative }, EN_US);
}

/** A two-decimal Currency code for these options. */
function currencyCode(negative: NegativeStyle, symbol: string, locale = EN_US): string | undefined {
  return buildNumberCode({ category: "currency", decimals: 2, negative, symbol }, locale);
}

/** An Accounting code for these options. */
function accountingCode(decimals: number, symbol: string, tag: SpreadsheetLocaleTag = "en-US"): string | undefined {
  return buildNumberCode({ category: "accounting", decimals, symbol }, SPREADSHEET_LOCALES[tag]);
}

describe("SPEC-123 buildNumberCode / describeNumberCode round trip", () => {
  it("lists the categories, symbols and types of the Number tab", () => {
    expect(NUMBER_CATEGORIES.map((category) => category.label)).toEqual([
      "General",
      "Number",
      "Currency",
      "Accounting",
      "Date",
      "Time",
      "Percentage",
      "Fraction",
      "Scientific",
      "Text",
      "Custom",
    ]);
    expect(NUMBER_CATEGORIES.map((category) => category.id)).toEqual([
      "general",
      "number",
      "currency",
      "accounting",
      "date",
      "time",
      "percentage",
      "fraction",
      "scientific",
      "text",
      "custom",
    ]);
    expect(CURRENCY_SYMBOLS).toEqual(["$", "€", "£", "¥", "Kč", "CHF"]);
    expect(TIME_TYPE_CODES).toEqual(["h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss", "mm:ss", "[h]:mm:ss", "mmss.0"]);
    expect(FRACTION_TYPE_CODES).toEqual(["# ?/?", "# ??/??", "# ???/???", "# ?/2", "# ?/4", "# ?/8", "# ??/16", "# ?/10", "# ??/100"]);
    for (const tag of SPREADSHEET_LOCALE_TAGS) {
      expect(dateTypeCodes(SPREADSHEET_LOCALES[tag]), tag).toEqual([
        "m/d/yyyy",
        "d-mmm-yy",
        "d-mmm",
        "mmm-yy",
        "m/d/yyyy h:mm",
        PRESETS[tag]["long-date"].numberFormat,
        "yyyy-mm-dd",
        "mmmm yyyy",
      ]);
    }
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: each symbol's side — the locale's own on its side, $ £ ¥ before, € Kč CHF after", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    const sides = CURRENCY_SYMBOLS.map((symbol) => currencySymbolPosition(symbol, locale));
    expect(sides).toEqual(["before", "after", "before", "before", "after", "after"]);
    expect(currencySymbolPosition(LOCALE_CURRENCY[tag].symbol, locale)).toBe(LOCALE_CURRENCY[tag].position);
    expect(currencySymbolPosition("¤", locale)).toBe(LOCALE_CURRENCY[tag].position);
  });

  it("builds Number with each negative style, and Currency with the symbol in every section", () => {
    expect(numberCode(2, false, "minus")).toBe("0.00");
    expect(numberCode(0, false, "minus")).toBe("0");
    expect(numberCode(0, true, "minus")).toBe("#,##0");
    expect(numberCode(2, true, "red")).toBe("#,##0.00;[Red]#,##0.00");
    expect(numberCode(1, false, "parens")).toBe("0.0_);(0.0)");
    expect(numberCode(3, true, "redParens")).toBe("#,##0.000_);[Red](#,##0.000)");
    expect(currencyCode("parens", "$")).toBe('"$"#,##0.00_);("$"#,##0.00)');
    expect(currencyCode("red", "$")).toBe('"$"#,##0.00;[Red]"$"#,##0.00');
    expect(currencyCode("redParens", "€")).toBe('#,##0.00 "€"_);[Red](#,##0.00 "€")');
    expect(currencyCode("minus", "CHF", DE_DE)).toBe('#,##0.00 "CHF"');
    expect(currencyCode("minus", "¥", DE_DE)).toBe('"¥"#,##0.00');
    expect(currencyCode("minus", "")).toBe("#,##0.00");
    expect(buildNumberCode({ category: "currency" }, DE_DE)).toBe('#,##0.00 "€"'); // the locale's symbol, 2 decimals
  });

  it("builds Accounting from the locale's template: en-US's are Excel's ids 41–44 exactly", () => {
    expect(accountingCode(2, "$")).toBe(ID_44);
    expect(accountingCode(0, "$")).toBe(ID_42);
    expect(accountingCode(2, "")).toBe(ID_43);
    expect(accountingCode(0, "")).toBe(ID_41);
    expect(accountingCode(2, "€")).toBe('_(* #,##0.00 "€"_);_(* \\(#,##0.00 "€"\\);_(* "-"?? "€"_);_(@_)');
    expect(accountingCode(2, "€", "en-GB")).toBe(EURO_ACCOUNTING);
    expect(accountingCode(1, "$", "de-DE")).toBe('_-"$"* #,##0.0_-;-"$"* #,##0.0_-;_-"$"* "-"?_-;_-@_-');
    expect(accountingCode(0, "Kč", "cs-CZ")).toBe('_-* #,##0 "Kč"_-;-* #,##0 "Kč"_-;_-* "-" "Kč"_-;_-@_-');
    expect(accountingCode(2, "", "fr-FR")).toBe(NO_SYMBOL_ACCOUNTING);
  });

  it("builds the other categories, clamps decimals and falls back to a category's first type", () => {
    expect(buildNumberCode({ category: "general" }, EN_US)).toBeUndefined();
    expect(buildNumberCode({ category: "percentage", decimals: 0 }, EN_US)).toBe("0%");
    expect(buildNumberCode({ category: "percentage", decimals: 1 }, EN_US)).toBe("0.0%");
    expect(buildNumberCode({ category: "scientific", decimals: 0 }, EN_US)).toBe("0E+00");
    expect(buildNumberCode({ category: "scientific" }, EN_US)).toBe("0.00E+00");
    expect(buildNumberCode({ category: "text" }, EN_US)).toBe("@");
    expect(buildNumberCode({ category: "number", decimals: 31 }, EN_US)).toBe(`0.${"0".repeat(30)}`);
    expect(buildNumberCode({ category: "number", decimals: -3 }, EN_US)).toBe("0");
    expect(buildNumberCode({ category: "number", decimals: 2.7 }, EN_US)).toBe("0.00");
    expect(buildNumberCode({ category: "fraction", type: "# ?/8" }, EN_US)).toBe("# ?/8");
    expect(buildNumberCode({ category: "fraction", type: "# ?/3" }, EN_US)).toBe("# ?/?");
    expect(buildNumberCode({ category: "date" }, DE_DE)).toBe("m/d/yyyy");
    expect(buildNumberCode({ category: "date", type: "dddd, d. mmmm yyyy" }, DE_DE)).toBe("dddd, d. mmmm yyyy");
    expect(buildNumberCode({ category: "date", type: "dddd, d. mmmm yyyy" }, EN_US)).toBe("m/d/yyyy");
    expect(buildNumberCode({ category: "time", type: "hh:mm" }, EN_US)).toBe("h:mm AM/PM");
    expect(buildNumberCode({ category: "custom", type: '0.0 "kg"' }, EN_US)).toBe('0.0 "kg"');
    for (const general of [undefined, "", "General", "general"]) {
      expect(buildNumberCode({ category: "custom", type: general }, EN_US), String(general)).toBeUndefined();
    }
  });

  it.each(SPREADSHEET_LOCALE_TAGS)("%s: every option set builds a valid code that describes back to it", (tag) => {
    const locale = SPREADSHEET_LOCALES[tag];
    const specs = everyOptionSet(tag);
    const codes = specs.map((spec) => buildNumberCode(spec, locale));
    const described = codes.map((code) => describeNumberCode(code, locale));
    expect(described.map((spec) => buildNumberCode(spec, locale))).toEqual(codes);
    // The one alias: Currency with no symbol is Number with thousands.
    const expected = specs.map(
      (spec): NumberCodeSpec =>
        spec.category === "currency" && spec.symbol === ""
          ? { category: "number", decimals: spec.decimals!, thousands: true, negative: spec.negative! }
          : spec,
    );
    expect(described).toStrictEqual(expected);
    expect(codes.filter((code) => code !== undefined && !isValidFormatCode(code))).toEqual([]);
  });

  it("describes General as General and every code no option set builds as Custom", () => {
    for (const general of [undefined, "", "General", "general"]) {
      expect(describeNumberCode(general, EN_US), String(general)).toStrictEqual({ category: "general" });
    }
    const custom = [
      "0.0#",
      "##0.0E+0",
      BUILTIN_NUMBER_FORMATS[7]!, // Excel escapes the parentheses the builder writes bare
      BUILTIN_NUMBER_FORMATS[39]!,
      "[Red]0.00",
      '"€"#,##0.00', // € goes after the number
      "0.00 %",
      "yyyy-mm-dd hh:mm",
      `0.${"0".repeat(31)}`,
      "dddd, d mmmm yyyy", // en-GB's long date, not en-US's
    ];
    for (const code of custom) {
      expect(describeNumberCode(code, EN_US), code).toStrictEqual({ category: "custom", type: code });
      expect(buildNumberCode(describeNumberCode(code, EN_US), EN_US)).toBe(code);
    }
  });

  it("describes Excel's built-ins and the presets by their options", () => {
    expect(describeNumberCode(ID_44, EN_US)).toStrictEqual({ category: "accounting", decimals: 2, symbol: "$" });
    expect(describeNumberCode(ID_41, EN_US)).toStrictEqual({ category: "accounting", decimals: 0, symbol: "" });
    expect(describeNumberCode("#,##0.00", EN_US)).toStrictEqual({ category: "number", decimals: 2, thousands: true, negative: "minus" });
    expect(describeNumberCode("0%", DE_DE)).toStrictEqual({ category: "percentage", decimals: 0 });
    expect(describeNumberCode(EURO_ACCOUNTING, DE_DE)).toStrictEqual({ category: "accounting", decimals: 2, symbol: "€" });
    expect(describeNumberCode("h:mm:ss", DE_DE)).toStrictEqual({ category: "time", type: "h:mm:ss" });
    expect(describeNumberCode("@", DE_DE)).toStrictEqual({ category: "text" });
  });

  it("hands out a fresh spec each time", () => {
    const spec = describeNumberCode("0.00", EN_US);
    spec.decimals = 5;
    expect(describeNumberCode("0.00", EN_US)).toStrictEqual({ category: "number", decimals: 2, thousands: false, negative: "minus" });
  });
});
