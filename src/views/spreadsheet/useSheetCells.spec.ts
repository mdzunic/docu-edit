// SPEC-117's cell layer as useSheetCells assembles it (AC-31, AC-33, AC-35,
// AC-45): a merged range is one element sized to the range — hosted by the
// first row of it that shows when its anchor row is hidden or scrolled off,
// and still one element (plus an aria-hidden copy that only paints) when it
// crosses a freeze line, since the frozen panes scroll independently — and a
// string spills into its empty neighbours whenever it has room, never
// gated on an estimate of its width. Mounted over the real metrics and
// viewport (jsdom has no layout, so the viewport is its 1280 × 720 fallback:
// 64 px columns and 20 px rows at 100%). SPEC-119: with Show formulas on, a
// formula cell shows "=" + formula left-aligned whatever its style says, an
// array member its anchor's "{=…}", and a constant its General text with no
// number format.
import { afterEach, describe, expect, it } from "vitest";
import { defineComponent, h, ref, type ComputedRef } from "vue";
import { mount } from "@vue/test-utils";
import SheetCellLayer from "../../components/spreadsheet/SheetCellLayer.vue";
import type { GridRange } from "../../lib/spreadsheet/address";
import type { Cell, CellStyle, Row, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { singleCell } from "../../lib/spreadsheet/selection";
import { cellAttributes } from "./sheetLayerProps";
import { useSheetCells, type BandItem, type CellItem } from "./useSheetCells";
import { useSheetMetrics } from "./useSheetMetrics";
import { useSheetViewport, type SheetViewport } from "./useSheetViewport";

const cleanups: Array<() => void> = [];
afterEach(() => {
  for (const cleanup of cleanups.splice(0)) {
    cleanup();
  }
});

/** A sheet from [row, col, value] triples (0-based), plus layout extras. */
function sheetOf(cells: ReadonlyArray<readonly [number, number, Cell["value"]]>, extra: Partial<Sheet> = {}): Sheet {
  const rows: Row[] = [];
  for (const [index, col, value] of cells.toSorted((a, b) => a[0] - b[0] || a[1] - b[1])) {
    const last = rows.at(-1);
    if (last?.index === index) {
      last.cells.push({ col, value });
    } else {
      rows.push({ index, cells: [{ col, value }] });
    }
  }
  return { conditionalFormats: [], validations: [], id: "s1", name: "Sheet1", visibility: "visible", kind: "grid", rows, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...extra };
}

/** `sheet` with row `index` hidden. */
function hideRow(sheet: Sheet, index: number): Sheet {
  const rows = sheet.rows.filter((row) => row.index !== index);
  const existing = sheet.rows.find((row) => row.index === index);
  rows.push({ ...(existing ?? { index, cells: [] }), hidden: true });
  return { ...sheet, rows: rows.toSorted((a, b) => a.index - b.index) };
}

function bookOf(sheet: Sheet, styles: CellStyle[] = [{}]): SpreadsheetModel {
  return {
    names: [],
    dxfs: [],
    kind: "spreadsheet",
    id: "book",
    title: "Book",
    format: "xlsx",
    readerVersion: 2,
    sheets: [sheet],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles,
    defaultStyleId: 0,
  };
}

/** The composable over `sheet` at `zoom`, mounted so the viewport's
 * lifecycle hooks have an instance. */
function setup(sheet: Sheet, options: { zoom?: number; styles?: CellStyle[]; showFormulas?: boolean } = {}) {
  const model = bookOf(sheet, options.styles);
  let bands!: ComputedRef<BandItem[]>;
  let viewport!: SheetViewport;
  const wrapper = mount(
    defineComponent({
      setup() {
        const sheetMetrics = useSheetMetrics(() => sheet);
        viewport = useSheetViewport(ref(null), {
          sheet: () => sheet,
          metrics: () => sheetMetrics.metrics.value,
          frozen: () => sheetMetrics.frozen.value,
          zoom: () => options.zoom ?? 1,
          initialTopLeft: () => undefined,
        });
        bands = useSheetCells({
          model: () => model,
          sheet: () => sheet,
          metrics: () => sheetMetrics.metrics.value,
          merges: () => sheetMetrics.merges.value,
          selection: () => singleCell({ row: 0, col: 0 }),
          viewport,
          showFormulas: () => options.showFormulas === true,
        }).bands;
        return () => h(SheetCellLayer, { bands: bands.value, frozenWidth: viewport.frozenW.value });
      },
    }),
  );
  cleanups.push(() => wrapper.unmount());
  return { bands, viewport, wrapper };
}

interface Placed {
  band: BandItem["kind"];
  region: "frozen" | "scrolled";
  row: number;
  cell: CellItem;
}

/** Every item the bands mount, with where it is mounted. */
function placed(bands: BandItem[]): Placed[] {
  return bands.flatMap((band) =>
    band.rows.flatMap((row) => [
      ...row.frozen.map((cell): Placed => ({ band: band.kind, region: "frozen", row: row.index, cell })),
      ...row.scrolled.map((cell): Placed => ({ band: band.kind, region: "scrolled", row: row.index, cell })),
    ]),
  );
}

/** The items drawn for `address`: its elements and its paint-only copies. */
function drawn(bands: BandItem[], address: string): { elements: Placed[]; copies: Placed[] } {
  const all = placed(bands).filter((item) => item.cell.ref === address);
  return { elements: all.filter((item) => !item.cell.paintOnly), copies: all.filter((item) => item.cell.paintOnly) };
}

const box = (cell: CellItem) => ({ left: cell.left, top: cell.top, width: cell.width, height: cell.height });
const textBox = (cell: CellItem) => ({ left: cell.textStyle.left, width: cell.textStyle.width });
const merge = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });

describe("useSheetCells — merges", () => {
  it("renders a merge whose anchor row is hidden once, hosted by the first row of it that shows", () => {
    // B4:D6 with row 4 hidden: rows 5–6 show the merge, 2 × 20 px tall.
    const sheet = hideRow(sheetOf([[3, 1, "Merged"]], { merges: [merge(3, 1, 5, 3)] }), 3);
    const { bands } = setup(sheet);
    const { elements, copies } = drawn(bands.value, "B4");
    expect(copies).toEqual([]);
    expect(elements).toHaveLength(1);
    const [only] = elements;
    expect(only!.row).toBe(4);
    expect(box(only!.cell)).toEqual({ left: 64, top: 0, width: 3 * 64, height: 2 * 20 });
    expect(only!.cell.text).toBe("Merged");
    expect([only!.cell.rowSpan, only!.cell.colSpan]).toEqual([3, 3]);
    // Its member cells render nothing of their own.
    for (const address of ["B5", "C5", "D6"]) {
      expect(drawn(bands.value, address).elements).toEqual([]);
    }
  });

  it("mounts that merge zoomed far out too, where only stored cells, merges and the active cell mount", () => {
    // A value in column 301 widens the extent, so 10% zoom shows over
    // 10,000 positions: the merge and the active cell A1 mount, nothing else.
    const sheet = hideRow(sheetOf([[3, 1, "Merged"], [0, 300, 1]], { merges: [merge(3, 1, 5, 3)] }), 3);
    const { bands } = setup(sheet, { zoom: 0.1 });
    expect(placed(bands.value).map((item) => item.cell.ref).toSorted()).toEqual(["A1", "B4"]);
    expect(drawn(bands.value, "B4").elements.map((item) => item.row)).toEqual([4]);
  });

  it("hosts a merge whose anchor row is scrolled off by the first row the window renders", () => {
    const sheet = sheetOf([[1, 1, "Tall"]], { merges: [merge(1, 1, 49, 2)] });
    const { bands, viewport } = setup(sheet);
    viewport.scrollToOffsets(0, 600); // row 31 at the top; the window starts 8 rows above
    const { elements } = drawn(bands.value, "B2");
    expect(elements).toHaveLength(1);
    expect(elements[0]!.row).toBe(22);
    expect(box(elements[0]!.cell)).toEqual({ left: 64, top: 20 - 22 * 20, width: 2 * 64, height: 49 * 20 });
  });

  it("gives a merge crossing the frozen rows one element, in the frozen band, and a paint-only copy below", () => {
    // Rows 1–2 frozen; A2:A4 merged.
    const sheet = sheetOf([[1, 0, "Tall"]], { merges: [merge(1, 0, 3, 0)], frozen: { rows: 2, cols: 0 } });
    const { bands, viewport } = setup(sheet);
    const { elements, copies } = drawn(bands.value, "A2");
    expect(elements.map(({ band, region, row }) => ({ band, region, row }))).toEqual([{ band: "frozen", region: "scrolled", row: 1 }]);
    expect(box(elements[0]!.cell)).toEqual({ left: 0, top: 0, width: 64, height: 60 });
    expect(copies.map(({ band, region, row }) => ({ band, region, row }))).toEqual([{ band: "main", region: "scrolled", row: 2 }]);
    // The copy paints the same range, from 20 px above its host row.
    expect(box(copies[0]!.cell)).toEqual({ left: 0, top: -20, width: 64, height: 60 });
    expect(copies[0]!.cell.text).toBe("Tall");

    // Scrolled well past it, only the frozen band still draws it.
    viewport.scrollToOffsets(0, 2000);
    expect(drawn(bands.value, "A2").elements.map((item) => item.band)).toEqual(["frozen"]);
    expect(drawn(bands.value, "A2").copies).toEqual([]);
  });

  it("gives a merge crossing the frozen columns one element, in the frozen columns, and a paint-only copy beside it", () => {
    // Column A frozen; the title A1:H1 runs on into the scrolled columns.
    const sheet = sheetOf([[0, 0, "Title"]], { merges: [merge(0, 0, 0, 7)], frozen: { rows: 0, cols: 1 } });
    const { bands } = setup(sheet);
    const { elements, copies } = drawn(bands.value, "A1");
    expect(elements.map(({ region, row }) => ({ region, row }))).toEqual([{ region: "frozen", row: 0 }]);
    expect(box(elements[0]!.cell)).toEqual({ left: 0, top: 0, width: 8 * 64, height: 20 });
    // The scrolled region starts at the freeze line, 64 px in.
    expect(copies.map(({ region, row }) => ({ region, row }))).toEqual([{ region: "scrolled", row: 0 }]);
    expect(box(copies[0]!.cell)).toEqual({ left: -64, top: 0, width: 8 * 64, height: 20 });
  });

  it("keeps one element per merge, with one id and data-ref, when it crosses both freeze lines", () => {
    const sheet = sheetOf([[0, 0, "Corner"]], { merges: [merge(0, 0, 4, 2)], frozen: { rows: 2, cols: 1 } });
    const { bands, wrapper } = setup(sheet);
    const { elements, copies } = drawn(bands.value, "A1");
    expect(elements.map(({ band, region }) => `${band}/${region}`)).toEqual(["frozen/frozen"]);
    expect(copies.map(({ band, region }) => `${band}/${region}`).toSorted()).toEqual(["frozen/scrolled", "main/frozen", "main/scrolled"]);

    // In the DOM: one gridcell for A1 — the id aria-activedescendant names —
    // and three aria-hidden copies that carry no identity.
    expect(wrapper.findAll('[data-testid="sheet-cell"][data-ref="A1"]')).toHaveLength(1);
    expect(wrapper.findAll("#sheet-cell-A1")).toHaveLength(1);
    const paintOnly = wrapper.findAll(".sheet-cell.is-paint-only");
    expect(paintOnly).toHaveLength(3);
    for (const copy of paintOnly) {
      expect(copy.attributes("aria-hidden")).toBe("true");
      for (const name of ["id", "role", "data-testid", "data-ref", "aria-rowindex", "aria-selected"]) {
        expect(copy.attributes(name), name).toBeUndefined();
      }
      expect(copy.text()).toBe("Corner");
    }
  });

  it("hides the gridlines inside an unfilled merge but not its own right and bottom ones", () => {
    const sheet = sheetOf([[4, 2, "Merged title"]], { merges: [merge(4, 2, 4, 4)] });
    const { bands } = setup(sheet);
    const style = drawn(bands.value, "C5").elements[0]!.cell.boxStyle;
    expect(style["background-image"]).toBe("linear-gradient(#ffffff, #ffffff)");
    expect(style["background-size"]).toBe("calc(100% - 1px) calc(100% - 1px)");
    expect(style["background-repeat"]).toBe("no-repeat");
  });

  it("keeps a filled merge's own fill", () => {
    const sheet: Sheet = { ...sheetOf([], { merges: [merge(4, 2, 4, 4)] }), rows: [{ index: 4, cells: [{ col: 2, value: "Merged", styleId: 1 }] }] };
    const { bands } = setup(sheet, { styles: [{}, { fill: { kind: "solid", color: "#ffff00" } }] });
    const style = drawn(bands.value, "C5").elements[0]!.cell.boxStyle;
    expect(style["background-color"]).toBe("#ffff00");
    expect(style["background-image"]).toBeUndefined();
  });
});

describe("useSheetCells — overflow", () => {
  it("spills a string the width estimate would put just inside its cell into empty neighbours", () => {
    // 8 characters × MDW 7 + padding = 60 px by the estimate, but capitals
    // and bold text render wider than 64 px: the text box takes the room.
    const sheet = sheetOf([
      [0, 0, "MMMMMMMM"],
      [1, 0, "WARNING!"],
      [2, 0, "WARNING!"],
      [2, 1, 5],
    ]);
    const { bands } = setup(sheet);
    const a1 = drawn(bands.value, "A1").elements[0]!.cell;
    expect(a1.raised).toBe(true);
    expect(a1.textStyle.left).toBe("0px");
    expect(Number.parseFloat(a1.textStyle.width!)).toBeGreaterThan(64);
    expect(drawn(bands.value, "A2").elements[0]!.cell.raised).toBe(true);
    // A neighbour holding a value blocks it: the text stays in its cell.
    const a3 = drawn(bands.value, "A3").elements[0]!.cell;
    expect(a3.raised).toBe(false);
    expect(textBox(a3)).toEqual({ left: "0px", width: "64px" });
  });

  it("widens the text box on the side the alignment spills to", () => {
    // C1 right-aligned: it spills left over A1:B1; E1 centred: it takes the
    // smaller side's room both ways (D1 to its left, F1… to its right).
    const sheet: Sheet = {
      ...sheetOf([]),
      rows: [{ index: 0, cells: [{ col: 2, value: "Right", styleId: 1 }, { col: 4, value: "Centre", styleId: 2 }] }],
    };
    const { bands } = setup(sheet, { styles: [{}, { alignment: { horizontal: "right" } }, { alignment: { horizontal: "center" } }] });
    expect(textBox(drawn(bands.value, "C1").elements[0]!.cell)).toEqual({ left: "-128px", width: "192px" });
    expect(textBox(drawn(bands.value, "E1").elements[0]!.cell)).toEqual({ left: "-64px", width: "192px" });
  });

  it("never spills a merge, a number or wrapped text", () => {
    const sheet: Sheet = {
      ...sheetOf([], { merges: [{ top: 0, left: 0, bottom: 0, right: 1 }] }),
      rows: [{ index: 0, cells: [{ col: 0, value: "Merged and long" }, { col: 3, value: 123456789 }, { col: 6, value: "Wrapped and long", styleId: 1 }] }],
    };
    const { bands } = setup(sheet, { styles: [{}, { alignment: { wrap: true } }] });
    for (const address of ["A1", "D1", "G1"]) {
      expect(drawn(bands.value, address).elements[0]!.cell.raised, address).toBe(false);
    }
  });
});

describe("cellAttributes", () => {
  it("gives a cell its id, gridcell role, test id, data-ref and 1-based ARIA indices", () => {
    const sheet = sheetOf([[4, 2, "Merged title"]], { merges: [{ top: 4, left: 2, bottom: 4, right: 4 }] });
    const { bands } = setup(sheet);
    expect(cellAttributes(drawn(bands.value, "C5").elements[0]!.cell)).toEqual({
      id: "sheet-cell-C5",
      role: "gridcell",
      "data-testid": "sheet-cell",
      "data-ref": "C5",
      "aria-rowindex": 6,
      "aria-colindex": 4,
      "aria-rowspan": undefined,
      "aria-colspan": 3,
      "aria-selected": "false",
    });
  });

  it("gives a merge's paint-only copy nothing but aria-hidden", () => {
    const sheet = sheetOf([[1, 0, "Tall"]], { merges: [{ top: 1, left: 0, bottom: 3, right: 0 }], frozen: { rows: 2, cols: 0 } });
    const { bands } = setup(sheet);
    expect(cellAttributes(drawn(bands.value, "A2").copies[0]!.cell)).toEqual({ "aria-hidden": "true" });
  });
});

describe("useSheetCells — Show formulas (SPEC-119)", () => {
  /** A1 a right-aligned "0.00" formula cell, B1 the same style on a
   * constant, C1 a date serial, A2:A3 an array (A3 a member), D1 text. */
  function formulaSheet(): Sheet {
    const rows: Row[] = [
      {
        index: 0,
        cells: [
          { col: 0, value: 60, formula: "SUM(B2:B4)", styleId: 1 },
          { col: 1, value: 1234.5, styleId: 1 },
          { col: 2, value: 46303, styleId: 2 },
          { col: 3, value: "text" },
        ],
      },
      { index: 1, cells: [{ col: 0, value: 6, formula: "B2:B3*2", arrayRange: { top: 1, left: 0, bottom: 2, right: 0 } }] },
      { index: 2, cells: [{ col: 0, value: 4 }] },
    ];
    return { ...sheetOf([]), rows };
  }
  const STYLES: CellStyle[] = [{}, { numberFormat: "0.00", alignment: { horizontal: "right" } }, { numberFormat: "m/d/yyyy", builtinFormatId: 14 }];

  const text = (bands: BandItem[], address: string): CellItem => drawn(bands, address).elements[0]!.cell;

  it("shows each formula as '=' + formula, left-aligned as text, and each constant's General text", () => {
    const { bands } = setup(formulaSheet(), { styles: STYLES, showFormulas: true });
    const a1 = text(bands.value, "A1");
    expect(a1.text).toBe("=SUM(B2:B4)");
    expect(a1.boxStyle["text-align"]).toBe("left");
    expect(a1.kindClass).toBe("is-text");
    expect(text(bands.value, "A2").text).toBe("{=B2:B3*2}");
    expect(text(bands.value, "A3").text).toBe("{=B2:B3*2}");
    expect(text(bands.value, "B1").text).toBe("1234.5");
    expect(text(bands.value, "B1").boxStyle["text-align"]).toBe("right");
    expect(text(bands.value, "C1").text).toBe("46303");
    expect(text(bands.value, "D1").text).toBe("text");
  });

  it("shows the formatted values when it is off", () => {
    const { bands } = setup(formulaSheet(), { styles: STYLES });
    expect(text(bands.value, "A1").text).toBe("60.00");
    expect(text(bands.value, "A1").boxStyle["text-align"]).toBe("right");
    expect(text(bands.value, "B1").text).toBe("1234.50");
    expect(text(bands.value, "C1").text).toMatch(/^#+$/); // a date never shrinks: too narrow, it is #
    expect(text(bands.value, "A3").text).toBe("4");
  });
});
