// SPEC-133 "Table styles": Excel's 60 built-in names; each row of the
// family-template table asserted cell by cell (header / body and stripes /
// edges / totals) for every member, with its colour slot (member 1 dk1,
// members 2–7 accent1–accent6); the neutral fallback (Dark 8–11, unknown and
// workbook-defined names) and None; tableCellLook's positions (header and
// totals only when present, stripes over the visible data rows only, first/
// last column bold only when on, nothing outside the ref); and styleOverLook's
// precedence — the cell's own properties win, Normal's count as unset.
import { describe, expect, it } from "vitest";
import { BUILT_IN_THEMES, DEFAULT_THEME, resolveColor, type ThemeColorRef, type ThemeColorSlot } from "../core/theme";
import type { GridRange } from "./address";
import type { CellBorderEdge, CellFill, CellStyle, Row, SheetTable, SheetTableStyle } from "./model";
import { BUILT_IN_TABLE_STYLE_NAMES, styleOverLook, tableCellLook, tableStyleLook } from "./tableStyles";

const SLOTS: ThemeColorSlot[] = ["dk1", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
const W: ThemeColorRef = { theme: "lt1" };
const K: ThemeColorRef = { theme: "dk1" };
const t80 = (slot: ThemeColorSlot): ThemeColorRef => ({ theme: slot, lumMod: 0.2, lumOff: 0.8 });
const t60 = (slot: ThemeColorSlot): ThemeColorRef => ({ theme: slot, lumMod: 0.4, lumOff: 0.6 });
const t40 = (slot: ThemeColorSlot): ThemeColorRef => ({ theme: slot, lumMod: 0.6, lumOff: 0.4 });
const thin = (color: ThemeColorRef | string): CellBorderEdge => ({ style: "thin", color });
const double = (color: ThemeColorRef): CellBorderEdge => ({ style: "double", color });
const solid = (color: ThemeColorRef | string): CellFill => ({ kind: "solid", color });

/** B2:E7 — header row 2, data rows 3–6, totals row 7 (0-based rows 1, 2–5, 6). */
const REF: GridRange = { top: 1, left: 1, bottom: 6, right: 4 };
const HEADER = 1;
const FIRST_STRIPE = 2;
const SECOND_STRIPE = 3;
const TOTALS = 6;
const INNER = 2; // a column inside the outline

interface TableOptions {
  ref?: GridRange;
  headerRow?: boolean;
  totalsRow?: boolean;
  style?: Partial<SheetTableStyle>;
}

/** A table styled `name`, banded rows on and the other options off. */
function tableOf(name: string | undefined, options: TableOptions = {}): SheetTable {
  return {
    id: "t1",
    name: "Sales",
    ref: options.ref ?? REF,
    headerRow: options.headerRow ?? true,
    totalsRow: options.totalsRow ?? true,
    columns: [{ name: "A" }, { name: "B" }, { name: "C" }, { name: "D" }],
    style: { name, showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false, ...options.style },
  };
}

const NO_ROWS = { rows: [] as Row[] };

/** The look of (row, col) in a table styled `name` on a sheet with no hidden rows. */
function at(name: string | undefined, row: number, col: number, options: TableOptions = {}): CellStyle | undefined {
  return tableCellLook(NO_ROWS, tableOf(name, options), row, col, DEFAULT_THEME);
}

/** A family's seven names with their colour slot. */
function members(kind: string, first: number): { name: string; slot: ThemeColorSlot }[] {
  return SLOTS.map((slot, i) => ({ name: `TableStyle${kind}${first + i}`, slot }));
}

/** Every value under a `color` key, anywhere in `value`. */
function colorsIn(value: unknown): unknown[] {
  if (typeof value !== "object" || value === null) {
    return [];
  }
  return Object.entries(value).flatMap(([key, inner]) => (key === "color" ? [inner] : colorsIn(inner)));
}

describe("BUILT_IN_TABLE_STYLE_NAMES", () => {
  it("lists Excel's 60 built-ins in gallery order: Light 1–21, Medium 1–28, Dark 1–11", () => {
    const expected = [
      ...Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`),
      ...Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`),
      ...Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`),
    ];
    expect(BUILT_IN_TABLE_STYLE_NAMES).toEqual(expected);
  });

  it("gives every built-in a look that paints its header, totals and first/last column", () => {
    for (const name of BUILT_IN_TABLE_STYLE_NAMES) {
      const look = tableStyleLook(name, DEFAULT_THEME);
      expect(look.headerRow, name).toBeDefined();
      expect(look.totalsRow, name).toBeDefined();
      expect(look.firstColumn, name).toEqual({ font: { bold: true } });
      expect(look.lastColumn, name).toEqual({ font: { bold: true } });
      expect(at(name, HEADER, INNER), name).toBeDefined();
      expect(at(name, TOTALS, INNER), name).toBeDefined();
    }
  });

  it("keeps every colour a theme reference, so a look follows the workbook theme", () => {
    for (const name of BUILT_IN_TABLE_STYLE_NAMES) {
      for (const color of colorsIn(tableStyleLook(name, DEFAULT_THEME))) {
        expect(color, name).toEqual(expect.objectContaining({ theme: expect.any(String) }));
      }
    }
    const header = at("TableStyleMedium2", HEADER, INNER)!.fill as { color: ThemeColorRef };
    expect(resolveColor(header.color, DEFAULT_THEME)).toBe(DEFAULT_THEME.colors.accent1);
    const facet = BUILT_IN_THEMES.find((theme) => theme.name === "Facet")!;
    expect(resolveColor(header.color, facet)).toBe(facet.colors.accent1);
  });
});

describe("tableStyleLook — the family templates, row by row", () => {
  it.each(members("Light", 1))("$name: header bold over thin C; stripes T80; table top and bottom thin C; totals bold under double C", ({ name, slot }) => {
    const c: ThemeColorRef = { theme: slot };
    expect(at(name, HEADER, INNER)).toEqual({ font: { bold: true }, borders: { top: thin(c), bottom: thin(c) } });
    expect(at(name, FIRST_STRIPE, INNER)).toEqual({ fill: solid(t80(slot)) });
    expect(at(name, SECOND_STRIPE, INNER)).toBeUndefined();
    expect(at(name, SECOND_STRIPE, 1, { style: { showColumnStripes: true } })).toEqual({ fill: solid(t80(slot)) });
    expect(at(name, SECOND_STRIPE, INNER, { style: { showColumnStripes: true } })).toBeUndefined();
    expect(at(name, TOTALS, INNER)).toEqual({ font: { bold: true }, borders: { top: double(c), bottom: thin(c) } });
  });

  it.each(members("Light", 8))("$name: header fill C, bold W; no body fill, stripes framed thin C; outline thin C; totals bold under double C", ({ name, slot }) => {
    const c: ThemeColorRef = { theme: slot };
    const header = { fill: solid(c), font: { bold: true, color: W } };
    expect(at(name, HEADER, INNER)).toEqual({ ...header, borders: { top: thin(c) } });
    expect(at(name, HEADER, 1)).toEqual({ ...header, borders: { top: thin(c), left: thin(c) } });
    expect(at(name, FIRST_STRIPE, INNER)).toEqual({ borders: { top: thin(c), bottom: thin(c) } });
    expect(at(name, SECOND_STRIPE, INNER)).toBeUndefined();
    expect(at(name, SECOND_STRIPE, 4)).toEqual({ borders: { right: thin(c) } });
    expect(at(name, FIRST_STRIPE, INNER, { style: { showRowStripes: false } })).toBeUndefined();
    // A column stripe is framed left and right.
    expect(at(name, SECOND_STRIPE, 3, { style: { showColumnStripes: true } })).toEqual({ borders: { left: thin(c), right: thin(c) } });
    expect(at(name, TOTALS, INNER)).toEqual({ font: { bold: true }, borders: { top: double(c), bottom: thin(c) } });
  });

  it.each(members("Light", 15))("$name: header bold over thin C; stripes T80; every cell thin C; totals bold under double C", ({ name, slot }) => {
    const c: ThemeColorRef = { theme: slot };
    const grid = { top: thin(c), right: thin(c), bottom: thin(c), left: thin(c) };
    expect(at(name, HEADER, INNER)).toEqual({ font: { bold: true }, borders: grid });
    expect(at(name, FIRST_STRIPE, INNER)).toEqual({ fill: solid(t80(slot)), borders: grid });
    expect(at(name, SECOND_STRIPE, 1)).toEqual({ borders: grid });
    expect(at(name, SECOND_STRIPE, 1, { style: { showColumnStripes: true } })).toEqual({ fill: solid(t80(slot)), borders: grid });
    expect(at(name, TOTALS, INNER)).toEqual({ font: { bold: true }, borders: { ...grid, top: double(c) } });
  });

  for (const [first, headerFill] of [
    [1, "C"],
    [15, "K"],
  ] as const) {
    it.each(members("Medium", first))(`$name: header fill ${headerFill}, bold W; stripes T80, inner horizontals thin T40; outline thin T40; totals bold under double C`, ({ name, slot }) => {
      const c: ThemeColorRef = { theme: slot };
      const line = thin(t40(slot));
      expect(at(name, HEADER, INNER)).toEqual({ fill: solid(headerFill === "K" ? K : c), font: { bold: true, color: W }, borders: { top: line, bottom: line } });
      expect(at(name, FIRST_STRIPE, INNER)).toEqual({ fill: solid(t80(slot)), borders: { top: line, bottom: line } });
      expect(at(name, SECOND_STRIPE, INNER)).toEqual({ borders: { top: line, bottom: line } });
      expect(at(name, SECOND_STRIPE, 1)).toEqual({ borders: { top: line, bottom: line, left: line } });
      expect(at(name, SECOND_STRIPE, 4)).toEqual({ borders: { top: line, bottom: line, right: line } });
      expect(at(name, TOTALS, INNER)).toEqual({ font: { bold: true }, borders: { top: double(c), bottom: line } });
    });
  }

  for (const [first, innerLine] of [
    [8, "W"],
    [22, "T40"],
  ] as const) {
    it.each(members("Medium", first))(`$name: header and totals fill C, bold W; body T80, stripes T60, inner lines ${innerLine}; no outline`, ({ name, slot }) => {
      const c: ThemeColorRef = { theme: slot };
      const line = thin(innerLine === "W" ? W : t40(slot));
      const strong = { fill: solid(c), font: { bold: true, color: W } };
      expect(at(name, HEADER, INNER)).toEqual({ ...strong, borders: { right: line, bottom: line, left: line } });
      expect(at(name, FIRST_STRIPE, INNER)).toEqual({ fill: solid(t60(slot)), borders: { top: line, right: line, bottom: line, left: line } });
      expect(at(name, SECOND_STRIPE, INNER)).toEqual({ fill: solid(t80(slot)), borders: { top: line, right: line, bottom: line, left: line } });
      expect(at(name, SECOND_STRIPE, 1)).toEqual({ fill: solid(t80(slot)), borders: { top: line, right: line, bottom: line } });
      expect(at(name, SECOND_STRIPE, 1, { style: { showColumnStripes: true } })?.fill).toEqual(solid(t60(slot)));
      expect(at(name, TOTALS, INNER)).toEqual({ ...strong, borders: { top: line, right: line, left: line } });
    });
  }

  it.each(members("Dark", 1))("$name: header and totals fill K, bold W; body C darkened 25% with W text, stripes C; no lines", ({ name, slot }) => {
    const strong = { fill: solid(K), font: { bold: true, color: W } };
    expect(at(name, HEADER, INNER)).toEqual(strong);
    expect(at(name, FIRST_STRIPE, INNER)).toEqual({ fill: solid({ theme: slot }), font: { color: W } });
    expect(at(name, SECOND_STRIPE, 1)).toEqual({ fill: solid({ theme: slot, lumMod: 0.75 }), font: { color: W } });
    expect(at(name, SECOND_STRIPE, 1, { style: { showColumnStripes: true } })).toEqual({ fill: solid({ theme: slot }), font: { color: W } });
    expect(at(name, TOTALS, INNER)).toEqual(strong);
  });

  it.each(["TableStyleDark8", "TableStyleDark9", "TableStyleDark10", "TableStyleDark11", "Invoice table", "TableStyleMedium29", "TableStyleLight0", "tablestylemedium2"])(
    "%s: the neutral fallback — header bold over thin K, totals bold, nothing else",
    (name) => {
      expect(at(name, HEADER, INNER)).toEqual({ font: { bold: true }, borders: { bottom: thin(K) } });
      expect(at(name, FIRST_STRIPE, INNER, { style: { showColumnStripes: true } })).toBeUndefined();
      expect(at(name, SECOND_STRIPE, 1)).toBeUndefined();
      expect(at(name, TOTALS, INNER)).toEqual({ font: { bold: true } });
      expect(at(name, SECOND_STRIPE, 1, { style: { showFirstColumn: true } })).toEqual({ font: { bold: true } });
    },
  );

  it("paints nothing for no style (None)", () => {
    expect(tableStyleLook(undefined, DEFAULT_THEME)).toEqual({});
    const all = { style: { showRowStripes: true, showColumnStripes: true, showFirstColumn: true, showLastColumn: true } };
    for (let row = REF.top; row <= REF.bottom; row += 1) {
      for (let col = REF.left; col <= REF.right; col += 1) {
        expect(at(undefined, row, col, all), `${row}:${col}`).toBeUndefined();
      }
    }
  });
});

describe("tableCellLook", () => {
  const NAME = "TableStyleMedium2";
  const line = thin(t40("accent1"));

  it("is undefined outside the table's ref", () => {
    for (const [row, col] of [
      [0, INNER],
      [7, INNER],
      [3, 0],
      [3, 5],
    ] as const) {
      expect(at(NAME, row, col), `${row}:${col}`).toBeUndefined();
    }
  });

  it("gives the header and totals parts only to a table that has those rows", () => {
    const bare = { headerRow: false, totalsRow: false };
    // ref.top is then the first data row (a first stripe, on the outline)…
    expect(at(NAME, HEADER, INNER, bare)).toEqual({ fill: solid(t80("accent1")), borders: { top: line, bottom: line } });
    // …and ref.bottom the last (the sixth data row: a second stripe).
    expect(at(NAME, TOTALS, INNER, bare)).toEqual({ borders: { top: line, bottom: line } });
  });

  it("alternates row stripes over the visible data rows only, the first visible one the first stripe", () => {
    const ref = { top: 0, left: 0, bottom: 8, right: 1 }; // a header row, data rows 1–8
    const table = tableOf(NAME, { ref, totalsRow: false });
    const rows: Row[] = [
      { index: 2, cells: [], hidden: true },
      { index: 4, cells: [{ col: 0, value: 4 }], filterHidden: true },
    ];
    const striped = (sheet: { rows: Row[] }, row: number): boolean => tableCellLook(sheet, table, row, 1, DEFAULT_THEME)?.fill !== undefined;
    expect([1, 3, 5, 6, 7, 8].map((row) => striped({ rows }, row))).toEqual([true, false, true, false, true, false]);
    // A new rows array (the model's copy-on-write) is a new answer.
    expect([1, 2, 3, 4, 5].map((row) => striped({ rows: [] }, row))).toEqual([true, false, true, false, true]);
    // A hidden first data row hands the first stripe to the next visible one.
    const firstHidden = { rows: [{ index: 1, cells: [], hidden: true as const }] };
    expect([2, 3, 4].map((row) => striped(firstHidden, row))).toEqual([true, false, true]);
  });

  it("paints no stripes when banded rows are off", () => {
    expect(at(NAME, FIRST_STRIPE, INNER, { style: { showRowStripes: false } })).toEqual({ borders: { top: line, bottom: line } });
  });

  it("alternates column stripes by column from the table's first", () => {
    const fills = [1, 2, 3, 4].map((col) => at(NAME, SECOND_STRIPE, col, { style: { showColumnStripes: true } })?.fill);
    expect(fills).toEqual([solid(t80("accent1")), undefined, solid(t80("accent1")), undefined]);
  });

  it("bolds the first and last column of the data body only when their options are on", () => {
    expect(at(NAME, SECOND_STRIPE, 1)?.font).toBeUndefined();
    expect(at(NAME, SECOND_STRIPE, 4)?.font).toBeUndefined();
    expect(at(NAME, SECOND_STRIPE, 1, { style: { showFirstColumn: true } })?.font).toEqual({ bold: true });
    expect(at(NAME, SECOND_STRIPE, INNER, { style: { showFirstColumn: true } })?.font).toBeUndefined();
    expect(at(NAME, SECOND_STRIPE, 4, { style: { showLastColumn: true } })?.font).toEqual({ bold: true });
    expect(at(NAME, SECOND_STRIPE, 4, { style: { showFirstColumn: true } })?.font).toBeUndefined();
    // The header row keeps the header part.
    expect(at(NAME, HEADER, 1, { style: { showFirstColumn: true } })?.font).toEqual({ bold: true, color: W });
  });

  it("returns one shared object per style and position class", () => {
    expect(at(NAME, SECOND_STRIPE, 2)).toBe(at(NAME, SECOND_STRIPE, 3));
    // Data rows 2–9: rows 3 and 7 are second stripes, 4 and 6 first ones.
    const tall = { ref: { ...REF, bottom: 10 } };
    expect(at(NAME, 3, INNER, tall)).toBe(at(NAME, 7, INNER, tall));
    expect(at(NAME, 4, INNER, tall)).toBe(at(NAME, 6, 3, tall));
    expect(at(NAME, 4, INNER, tall)).not.toBe(at(NAME, 3, INNER, tall));
  });
});

/** `value` frozen all the way down — a mutation would throw. */
function deepFreeze<T>(value: T): T {
  if (typeof value === "object" && value !== null) {
    Object.values(value).forEach(deepFreeze);
    Object.freeze(value);
  }
  return value;
}

describe("styleOverLook", () => {
  const RED = "#ff0000";
  const LOOK: CellStyle = {
    font: { bold: true, color: W },
    fill: solid({ theme: "accent1" }),
    borders: { top: thin(t40("accent1")), bottom: double({ theme: "accent1" }) },
  };
  /** A blank workbook's Normal style (SPEC-121's createBlankWorkbook). */
  const NORMAL: CellStyle = { font: { family: { theme: "minor" }, sizePt: 11, color: { theme: "dk1" } }, namedStyle: "Normal" };

  it("returns the cell's own style when there is no look", () => {
    const own: CellStyle = { fill: solid(RED) };
    expect(styleOverLook(undefined, own)).toBe(own);
  });

  it("lets the look show wherever the cell sets nothing", () => {
    expect(styleOverLook(LOOK, {})).toEqual(LOOK);
  });

  it("lets the cell's own fill, font properties and border edges win one by one", () => {
    const own: CellStyle = { font: { italic: true, color: RED }, fill: solid("#ffff00"), borders: { top: thin(RED) } };
    expect(styleOverLook(LOOK, own)).toEqual({
      font: { bold: true, italic: true, color: RED },
      fill: solid("#ffff00"),
      borders: { top: thin(RED), bottom: double({ theme: "accent1" }) },
    });
    // The cell's bold beside the look's colour.
    expect(styleOverLook(LOOK, { font: { bold: true } }).font).toEqual({ bold: true, color: W });
  });

  it("keeps the cell's other properties", () => {
    const own: CellStyle = { numberFormat: "0.00", alignment: { horizontal: "center" }, protection: { locked: false }, borders: { diagonal: { style: "thin", up: true } } };
    expect(styleOverLook(LOOK, own)).toEqual({ ...LOOK, ...own, borders: { ...LOOK.borders, diagonal: { style: "thin", up: true } } });
  });

  it("counts a value inherited from Normal as unset, so the look's colour shows over a blank workbook's dk1", () => {
    expect(styleOverLook(LOOK, NORMAL, NORMAL).font).toEqual({ family: { theme: "minor" }, sizePt: 11, bold: true, color: W });
    // A colour of the cell's own still wins, as does one Normal never set.
    expect(styleOverLook(LOOK, { ...NORMAL, font: { ...NORMAL.font, color: RED } }, NORMAL).font?.color).toBe(RED);
    expect(styleOverLook(LOOK, { font: { color: RED } }, {}).font?.color).toBe(RED);
    // The source's tint spelling is not a difference.
    expect(styleOverLook(LOOK, { font: { color: { theme: "dk1", xlsxTint: "0" } as ThemeColorRef } }, NORMAL).font?.color).toEqual(W);
    // Without Normal every present property is the cell's own.
    expect(styleOverLook(LOOK, NORMAL).font?.color).toEqual({ theme: "dk1" });
  });

  it("returns the cell's own style itself when the look changes nothing, and mutates neither input", () => {
    const own = deepFreeze<CellStyle>({ font: { bold: true, color: RED }, fill: solid(RED), borders: { top: thin(RED), bottom: thin(RED) } });
    const look = deepFreeze<CellStyle>(structuredClone(LOOK));
    expect(styleOverLook(look, own)).toBe(own);
    const plain = deepFreeze<CellStyle>({ font: { italic: true } });
    const merged = styleOverLook(look, plain);
    expect(merged).not.toBe(plain);
    expect(plain).toEqual({ font: { italic: true } });
    expect(look).toEqual(LOOK);
  });
});
