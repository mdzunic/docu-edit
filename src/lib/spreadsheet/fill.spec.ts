// SPEC-125 "Series detection", "Double-click fill": every example of the
// series table as a unit test, reverse fills (up and left), the fill
// handle's modes (auto, the Ctrl/Alt toggle, and the chip's copy, series,
// formats only and without formatting), shrinking back inside the seed, the
// double-click extent, and fill down / fill right.
import { describe, expect, it } from "vitest";
import type { GridRange } from "./address";
import { getCell } from "./address";
import { PROTECTED_CELL_MESSAGE } from "./cells";
import { PASTE_MERGE_MESSAGE } from "./clipboard";
import { detectSeries, doubleClickFillTarget, fillDownRight, fillRange, type SeriesSeed } from "./fill";
import type { Cell, CellValue, Row, Sheet, SpreadsheetModel } from "./model";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { dateToSerial, serialToParts } from "./serial";

const EN_US = SPREADSHEET_LOCALES["en-US"];
const DATE = "m/d/yyyy";
const TIME = "h:mm";

function sheet(fields: Partial<Sheet> = {}): Sheet {
  return { id: "s1", name: "Sheet1", visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...fields };
}

function model(s: Sheet, fields: Partial<SpreadsheetModel> = {}): SpreadsheetModel {
  return { kind: "spreadsheet", id: "book", title: "T", format: "xlsx", readerVersion: 8, sheets: [s], activeSheetIndex: 0, dateSystem: "1900", opaque: [], styles: [{}], defaultStyleId: 0, ...fields };
}

const value = (col: number, v: Cell["value"], extra: Partial<Cell> = {}): Cell => ({ col, value: v, ...extra });
const row = (index: number, cells: Cell[], extra: Partial<Row> = {}): Row => ({ index, cells, ...extra });
const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });
const cellOf = (m: SpreadsheetModel, r: number, c: number): Cell | undefined => getCell(m.sheets[0]!, r, c);
const date = (y: number, mo: number, d: number): number => dateToSerial(y, mo, d, "1900");
const iso = (serial: CellValue): string => {
  const p = serialToParts(serial as number, "1900");
  return `${p.year}-${String(p.month).padStart(2, "0")}-${String(p.day).padStart(2, "0")}`;
};

/** The continuation of a seed line: the values at indices n, n+1, … */
function next(seed: SeriesSeed[], count: number, options: { toggle?: boolean; locale?: keyof typeof SPREADSHEET_LOCALES } = {}): CellValue[] {
  const series = detectSeries(seed, SPREADSHEET_LOCALES[options.locale ?? "en-US"]);
  const extend = series.kind !== "repeat" && (options.toggle ? !series.byDefault : series.byDefault);
  return Array.from({ length: count }, (_, k) => {
    const i = seed.length + k;
    return extend ? series.at(i).value : seed[i % seed.length]!.value;
  });
}

const nums = (...values: number[]): SeriesSeed[] => values.map((v) => ({ value: v }));
const dates = (...values: number[]): SeriesSeed[] => values.map((v) => ({ value: v, code: DATE }));
const texts = (...values: string[]): SeriesSeed[] => values.map((v) => ({ value: v }));

describe("detectSeries — every example of the series table", () => {
  it("5 → 5, 5 (toggle: 6, 7)", () => {
    expect(next(nums(5), 2)).toEqual([5, 5]);
    expect(next(nums(5), 2, { toggle: true })).toEqual([6, 7]);
  });

  it("1, 3 → 5, 7", () => {
    expect(next(nums(1, 3), 2)).toEqual([5, 7]);
  });

  it("1, 2, 4 → least squares 5.33333333333333, 6.83333333333333", () => {
    expect(next(nums(1, 2, 4), 2)).toEqual([5.33333333333333, 6.83333333333333]);
  });

  it("0.1, 0.2 → 0.3, 0.4 (15 significant digits, never 0.30000000000000004)", () => {
    expect(next(nums(0.1, 0.2), 2)).toEqual([0.3, 0.4]);
  });

  it("2026-01-31 (a date) → 2026-02-01, … (toggle: copy)", () => {
    expect(next(dates(date(2026, 1, 31)), 2).map(iso)).toEqual(["2026-02-01", "2026-02-02"]);
    expect(next(dates(date(2026, 1, 31)), 2, { toggle: true }).map(iso)).toEqual(["2026-01-31", "2026-01-31"]);
  });

  it("2026-01-31, 2026-02-28 → month-ends 2026-03-31, 2026-04-30", () => {
    expect(next(dates(date(2026, 1, 31), date(2026, 2, 28)), 2).map(iso)).toEqual(["2026-03-31", "2026-04-30"]);
  });

  it("2026-01-15, 2026-02-15 → 2026-03-15 (EDATE from the first seed)", () => {
    expect(next(dates(date(2026, 1, 15), date(2026, 2, 15)), 2).map(iso)).toEqual(["2026-03-15", "2026-04-15"]);
    // Clamped to a short month's end, then back to the 31st.
    expect(next(dates(date(2026, 1, 31), date(2026, 3, 31)), 2).map(iso)).toEqual(["2026-05-31", "2026-07-31"]);
  });

  it("2024-03-01, 2025-03-01 → 2026-03-01", () => {
    expect(next(dates(date(2024, 3, 1), date(2025, 3, 1)), 2).map(iso)).toEqual(["2026-03-01", "2027-03-01"]);
  });

  it("2026-01-01, 2026-01-08 → weekly 2026-01-15", () => {
    expect(next(dates(date(2026, 1, 1), date(2026, 1, 8)), 2).map(iso)).toEqual(["2026-01-15", "2026-01-22"]);
  });

  it("9:00 (a time-only format) → 10:00, 11:00", () => {
    expect(next([{ value: 9 / 24, code: TIME }], 2).map((v) => Math.round((v as number) * 24 * 60))).toEqual([600, 660]);
  });

  it("Mon / MONDAY / monday → Tue / TUESDAY / tuesday …, wrapping", () => {
    expect(next(texts("Mon"), 7)).toEqual(["Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Mon"]);
    expect(next(texts("MONDAY"), 2)).toEqual(["TUESDAY", "WEDNESDAY"]);
    expect(next(texts("monday"), 2)).toEqual(["tuesday", "wednesday"]);
  });

  it("Jan, Apr → Jul, Oct, Jan", () => {
    expect(next(texts("Jan", "Apr"), 3)).toEqual(["Jul", "Oct", "Jan"]);
  });

  it("Montag (de-DE locale) → Dienstag", () => {
    expect(next(texts("Montag"), 1, { locale: "de-DE" })).toEqual(["Dienstag"]);
  });

  it("Q3 2026 / Qtr 4 → Q4 2026, Q1 2027 / Qtr 1", () => {
    expect(next(texts("Q3 2026"), 2)).toEqual(["Q4 2026", "Q1 2027"]);
    expect(next(texts("Qtr 4"), 1)).toEqual(["Qtr 1"]);
    expect(next(texts("Quarter 1"), 1)).toEqual(["Quarter 2"]);
  });

  it("Item 9 / Item 009 / Item 1, Item 3 → Item 10 / Item 010 / Item 5", () => {
    expect(next(texts("Item 9"), 1)).toEqual(["Item 10"]);
    expect(next(texts("Item 009"), 1)).toEqual(["Item 010"]);
    expect(next(texts("Item 1", "Item 3"), 1)).toEqual(["Item 5"]);
  });

  it("1st → 2nd, 3rd, 4th, …, 11th, 12th, 13th, 21st", () => {
    const ordinals = next(texts("1st"), 20);
    expect(ordinals.slice(0, 3)).toEqual(["2nd", "3rd", "4th"]);
    expect(ordinals.slice(9, 12)).toEqual(["11th", "12th", "13th"]);
    expect(ordinals[19]).toBe("21st");
  });

  it("=A1*2 at B1 → =A2*2 at B2 (shifted by the distance)", () => {
    const series = detectSeries([{ value: 2, formula: "A1*2" }], EN_US);
    expect(series.at(1)).toEqual({ value: 2, formula: "A2*2" });
    expect(detectSeries([{ value: 2, formula: "A1*2" }], EN_US, { axis: "col" }).at(2)).toEqual({ value: 2, formula: "C1*2" });
  });

  it("TRUE, #N/A, abc, Item 1, Thing 2, mixed kinds → the seed pattern repeats", () => {
    expect(next([{ value: true }], 2)).toEqual([true, true]);
    expect(next([{ value: { error: "#N/A" } }], 1)).toEqual([{ error: "#N/A" }]);
    expect(next(texts("abc"), 2)).toEqual(["abc", "abc"]);
    expect(next(texts("Item 1", "Thing 2"), 2)).toEqual(["Item 1", "Thing 2"]);
    expect(next([{ value: 1 }, { value: "a" }], 2)).toEqual([1, "a"]);
    expect(detectSeries(texts("Mon", "Wed", "Thu"), EN_US).kind).toBe("repeat"); // the step must be constant
  });
});

describe("fillRange", () => {
  /** B2 = 1 (bold), B3 = 3 (italic); D2 = "Mon"; F2 = =B2*2. */
  function book(): SpreadsheetModel {
    return model(
      sheet({ rows: [row(1, [value(1, 1, { styleId: 1 }), value(3, "Mon"), value(5, 2, { formula: "B2*2" })]), row(2, [value(1, 3, { styleId: 2 })])] }),
      { styles: [{}, { font: { bold: true } }, { font: { italic: true } }] },
    );
  }

  it("extends each seed line down its series, styles taken k mod n", () => {
    const m = book();
    expect(fillRange(m, "s1", range(1, 1, 2, 1), range(1, 1, 5, 1), "auto")).toEqual({ ok: true, applied: "series", range: range(1, 1, 5, 1) });
    expect([3, 4, 5].map((r) => cellOf(m, r, 1))).toEqual([value(1, 5, { styleId: 1 }), value(1, 7, { styleId: 2 }), value(1, 9, { styleId: 1 })]);
    expect(m.sheets[0]!.dirty).toBe(true);
  });

  it("extends backwards filling up and left", () => {
    const m = book();
    fillRange(m, "s1", range(1, 1, 2, 1), range(0, 1, 2, 1), "auto");
    expect(cellOf(m, 0, 1)).toEqual(value(1, -1, { styleId: 2 }));
    fillRange(m, "s1", range(1, 3, 1, 3), range(1, 1, 1, 3), "auto");
    expect([cellOf(m, 1, 2)?.value, cellOf(m, 1, 1)?.value]).toEqual(["Sun", "Sat"]);
  });

  it("toggles copy ↔ series (Ctrl / Alt at release)", () => {
    const one = model(sheet({ rows: [row(0, [value(0, 5)])] }));
    fillRange(one, "s1", range(0, 0, 0, 0), range(0, 0, 2, 0), "auto");
    expect([1, 2].map((r) => cellOf(one, r, 0)?.value)).toEqual([5, 5]);
    expect(fillRange(one, "s1", range(0, 0, 0, 0), range(0, 0, 2, 0), "toggle")).toMatchObject({ applied: "series" });
    expect([1, 2].map((r) => cellOf(one, r, 0)?.value)).toEqual([6, 7]);
    const m = book();
    expect(fillRange(m, "s1", range(1, 1, 2, 1), range(1, 1, 4, 1), "toggle")).toMatchObject({ applied: "copy" });
    expect([3, 4].map((r) => cellOf(m, r, 1)?.value)).toEqual([1, 3]);
  });

  it("the chip's modes: copy cells, fill series, formatting only, without formatting", () => {
    const copy = book();
    fillRange(copy, "s1", range(1, 3, 1, 3), range(1, 3, 3, 3), "copy");
    expect([2, 3].map((r) => cellOf(copy, r, 3)?.value)).toEqual(["Mon", "Mon"]);
    const series = book();
    fillRange(series, "s1", range(1, 3, 1, 3), range(1, 3, 3, 3), "series");
    expect([2, 3].map((r) => cellOf(series, r, 3)?.value)).toEqual(["Tue", "Wed"]);
    const formats = book();
    formats.sheets[0]!.rows.push(row(3, [value(1, "keep")]));
    expect(fillRange(formats, "s1", range(1, 1, 2, 1), range(1, 1, 3, 1), "formats")).toMatchObject({ applied: "formats" });
    expect(cellOf(formats, 3, 1)).toEqual(value(1, "keep", { styleId: 1 }));
    const plain = book();
    plain.sheets[0]!.rows.push(row(3, [value(1, "old", { styleId: 2 })]));
    expect(fillRange(plain, "s1", range(1, 1, 2, 1), range(1, 1, 3, 1), "no-formats")).toMatchObject({ applied: "no-formats" });
    expect(cellOf(plain, 3, 1)).toEqual(value(1, 5, { styleId: 2 }));
  });

  it("shifts formulas by their distance in every mode", () => {
    const m = book();
    fillRange(m, "s1", range(1, 5, 1, 5), range(1, 5, 3, 5), "series");
    expect([2, 3].map((r) => cellOf(m, r, 5)?.formula)).toEqual(["B3*2", "B4*2"]);
    fillRange(m, "s1", range(1, 5, 1, 5), range(1, 5, 1, 7), "copy");
    expect(cellOf(m, 1, 7)?.formula).toBe("D2*2");
  });

  it("dragging back inside the seed clears the contents (not the styles) of the cells left behind", () => {
    const m = book();
    expect(fillRange(m, "s1", range(1, 1, 2, 1), range(1, 1, 1, 1), "auto")).toMatchObject({ ok: true, cleared: true });
    expect(cellOf(m, 1, 1)?.value).toBe(1);
    expect(cellOf(m, 2, 1)).toEqual(value(1, null, { styleId: 2 }));
  });

  it("repeats the seed's merges and refuses a merge crossing the band, or a locked cell", () => {
    const merged = model(sheet({ rows: [row(0, [value(0, "a")])], merges: [range(0, 0, 0, 1)] }));
    fillRange(merged, "s1", range(0, 0, 0, 1), range(0, 0, 2, 1), "auto");
    expect(merged.sheets[0]!.merges).toEqual([range(0, 0, 0, 1), range(1, 0, 1, 1), range(2, 0, 2, 1)]);
    const crossing = model(sheet({ rows: [row(0, [value(0, 1)])], merges: [range(2, 0, 3, 1)] }));
    expect(fillRange(crossing, "s1", range(0, 0, 0, 0), range(0, 0, 2, 0), "auto")).toEqual({ ok: false, message: PASTE_MERGE_MESSAGE });
    const locked = model(sheet({ rows: [row(0, [value(0, 1)])], protection: { sheet: true } }));
    expect(fillRange(locked, "s1", range(0, 0, 0, 0), range(0, 0, 2, 0), "auto")).toEqual({ ok: false, message: PROTECTED_CELL_MESSAGE });
  });
});

describe("double-click fill — the extent rule", () => {
  /** A1:A20 hold 1…20; B1, the seed, =A1*2. */
  const column = (): Sheet => sheet({ rows: Array.from({ length: 20 }, (_, r) => row(r, [value(0, r + 1), ...(r === 0 ? [value(1, 2, { formula: "A1*2" })] : [])])) });

  it("follows the left column's contiguous data", () => {
    expect(doubleClickFillTarget(column(), range(0, 1, 0, 1))).toEqual(range(0, 1, 19, 1));
  });

  it("follows the right column when the left one is empty below the selection", () => {
    const s = sheet({ rows: Array.from({ length: 5 }, (_, r) => row(r, [value(3, r), ...(r === 0 ? [value(2, "x")] : [])])) });
    expect(doubleClickFillTarget(s, range(0, 2, 0, 2))).toEqual(range(0, 2, 4, 2));
  });

  it("stops before the first non-empty cell below the selection in its own columns", () => {
    const s = column();
    s.rows[9] = row(9, [value(0, 10), value(1, "stop")]);
    expect(doubleClickFillTarget(s, range(0, 1, 0, 1))).toEqual(range(0, 1, 8, 1));
  });

  it("does nothing without adjacent data", () => {
    expect(doubleClickFillTarget(sheet({ rows: [row(0, [value(5, 1)])] }), range(0, 5, 0, 5))).toBeNull();
  });

  it("fills those rows through fillRange", () => {
    const m = model(column());
    const target = doubleClickFillTarget(m.sheets[0]!, range(0, 1, 0, 1))!;
    fillRange(m, "s1", range(0, 1, 0, 1), target, "auto");
    expect(cellOf(m, 19, 1)?.formula).toBe("A20*2");
  });
});

describe("fill down / fill right", () => {
  it("copies each column's top cell into the rest of the selection, formulas shifted and styles copied, never a series", () => {
    const m = model(sheet({ rows: [row(0, [value(0, 1, { styleId: 1 }), value(1, 0, { formula: "A1+1" })])] }), { styles: [{}, { font: { bold: true } }] });
    expect(fillDownRight(m, "s1", [range(0, 0, 3, 1)], "down")).toMatchObject({ ok: true });
    expect([1, 2, 3].map((r) => cellOf(m, r, 0))).toEqual([1, 2, 3].map(() => value(0, 1, { styleId: 1 })));
    expect(cellOf(m, 3, 1)?.formula).toBe("A4+1");
  });

  it("copies each row's left cell to the right", () => {
    const m = model(sheet({ rows: [row(0, [value(0, "Mon")]), row(1, [value(0, 2, { formula: "A1*2" })])] }));
    fillDownRight(m, "s1", [range(0, 0, 1, 2)], "right");
    expect([1, 2].map((c) => cellOf(m, 0, c)?.value)).toEqual(["Mon", "Mon"]);
    expect(cellOf(m, 1, 2)?.formula).toBe("C1*2");
  });

  it("a one-row (one-column) selection copies from the row above (the column to the left)", () => {
    const m = model(sheet({ rows: [row(0, [value(0, 7), value(1, 8)])] }));
    fillDownRight(m, "s1", [range(1, 0, 1, 1)], "down");
    expect([cellOf(m, 1, 0)?.value, cellOf(m, 1, 1)?.value]).toEqual([7, 8]);
    fillDownRight(m, "s1", [range(0, 2, 1, 2)], "right");
    expect([cellOf(m, 0, 2)?.value, cellOf(m, 1, 2)?.value]).toEqual([8, 8]);
  });
});
