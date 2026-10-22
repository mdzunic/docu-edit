// SPEC-137 "The grid's ceiling" (AC-47–AC-51, AC-66), on the mounted grid:
// cells are keyed by viewport slot, so a scroll step patches the same
// elements — the ones that stay keep their nodes, and the slot the leaving
// row vacated re-renders the entering row's address; never more than
// GRID_MAX_RENDERED_CELLS `sheet-cell` elements nor GRID_MAX_DOM_NODES
// elements under `spreadsheet-grid`, at 1280 × 720 and 1920 × 1080 from
// 100% down to 10% — on a dense sheet with every cell bordered, the worst
// case for the node budget; and the overview switch: zoomed far out the
// `sheet-overview-canvas` paints the sheet and only the active cell keeps an
// element, and back at 100% the canvas goes and the DOM cells return. jsdom
// has no layout, so the scroller's client size is stubbed, and no 2D
// context, so getContext answers null (the painter then paints nothing).
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defineComponent, h, nextTick, ref } from "vue";
import { mount, type VueWrapper } from "@vue/test-utils";
import type { CellStyle, Row, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SheetViewFields } from "../../views/spreadsheet/useSheetViewState";
import { GRID_MAX_DOM_NODES, GRID_MAX_RENDERED_CELLS } from "../../views/spreadsheet/useSheetCells";
import { useSheetMetrics } from "../../views/spreadsheet/useSheetMetrics";
import { useSheetSelection } from "../../views/spreadsheet/useSheetSelection";
import { useSheetZoom } from "../../views/spreadsheet/useSheetZoom";
import SheetGrid from "./SheetGrid.vue";

const size = { width: 1280, height: 720 };
const wrappers: VueWrapper[] = [];

beforeEach(() => {
  for (const [name, axis] of [["clientWidth", "width"], ["clientHeight", "height"]] as const) {
    Object.defineProperty(HTMLElement.prototype, name, {
      configurable: true,
      get(this: HTMLElement) {
        return this.dataset.testid === "spreadsheet-grid" ? size[axis] : 0;
      },
    });
  }
  vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockReturnValue(null);
});

afterEach(() => {
  for (const wrapper of wrappers.splice(0)) {
    wrapper.unmount();
  }
  vi.restoreAllMocks();
});

/** `rows` × `cols` of values (numbers and text alternating), every cell in
 * style `styleId`. */
function denseSheet(rows: number, cols: number, styleId = 0): Sheet {
  const out: Row[] = [];
  for (let r = 0; r < rows; r += 1) {
    out.push({ index: r, cells: Array.from({ length: cols }, (_, c) => ({ col: c, value: c % 2 === 0 ? r * cols + c : `t${r}`, ...(styleId ? { styleId } : {}) })) });
  }
  return { conditionalFormats: [], validations: [], id: "s1", name: "Sheet1", visibility: "visible", kind: "grid", rows: out, columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] };
}

const THIN = { style: "thin" as const };
const STYLES: CellStyle[] = [{}, { borders: { top: THIN, bottom: THIN, left: THIN, right: THIN } }];

function bookOf(sheet: Sheet): SpreadsheetModel {
  return { names: [], dxfs: [], kind: "spreadsheet", id: "book", title: "Book", format: "xlsx", readerVersion: 2, sheets: [sheet], activeSheetIndex: 0, dateSystem: "1900", opaque: [], styles: STYLES, defaultStyleId: 0 };
}

/** SheetGrid over `sheet` with the view-level state it is handed. */
function mountGrid(sheet: Sheet, zoomPercent = 100) {
  const model = bookOf(sheet);
  const active = ref<SheetViewFields>({ zoomPercent });
  const viewState = {
    active,
    activeSheet: ref(sheet),
    updateActive(patch: SheetViewFields) {
      active.value = { ...active.value, ...patch };
    },
  };
  const wrapper = mount(
    defineComponent({
      setup() {
        const sheetMetrics = useSheetMetrics(() => sheet);
        const selection = useSheetSelection({ sheet: () => sheet, sheetMetrics, viewState, model: () => model });
        const zoom = useSheetZoom({ viewState });
        const context = { model, sheetMetrics, selection, zoom, viewState, formulas: {} };
        return () => h(SheetGrid, { sheet, context });
      },
    }),
    { attachTo: document.body },
  );
  wrappers.push(wrapper);
  const grid = (): HTMLElement => wrapper.get('[data-testid="spreadsheet-grid"]').element as HTMLElement;
  return {
    grid,
    cells: () => [...grid().querySelectorAll<HTMLElement>(".sheet-cell")],
    elements: () => grid().querySelectorAll("*").length + 1,
    canvas: () => grid().querySelector('[data-testid="sheet-overview-canvas"]'),
    async zoomTo(percent: number) {
      viewState.updateActive({ zoomPercent: percent });
      await settle();
    },
    async scrollTo(top: number) {
      Object.defineProperty(grid(), "scrollTop", { configurable: true, writable: true, value: top });
      grid().dispatchEvent(new Event("scroll"));
      await settle();
    },
  };
}

async function settle(): Promise<void> {
  await nextTick();
  await nextTick();
}

describe("SheetGrid — slot keys (AC-47)", () => {
  it("patches the same cell elements across a scroll step: the staying cells keep their nodes, the vacated slot shows the entering row", async () => {
    const view = mountGrid(denseSheet(400, 30));
    await view.scrollTo(400); // row 21 at the top
    const before = new Map(view.cells().map((element) => [element.dataset.ref!, element]));
    const columnA = [...before.keys()].flatMap((ref) => (/^A\d+$/.test(ref) ? [Number(ref.slice(1))] : []));
    const firstRow = Math.min(...columnA);
    const lastRow = Math.max(...columnA);
    const leaving = before.get(`A${firstRow}`)!;
    const staying = before.get(`A${firstRow + 5}`)!;

    await view.scrollTo(420); // one row down
    const after = view.cells();
    expect(after).toHaveLength(before.size);
    // No element was mounted or unmounted: the same nodes, patched.
    expect(new Set(after)).toEqual(new Set(before.values()));
    expect(staying.dataset.ref).toBe(`A${firstRow + 5}`);
    expect(leaving.isConnected).toBe(true);
    expect(leaving.dataset.ref).toBe(`A${lastRow + 1}`);
    expect(leaving.id).toBe(`sheet-cell-A${lastRow + 1}`);
    expect(leaving.getAttribute("aria-rowindex")).toBe(String(lastRow + 2));
    expect(after.some((element) => element.dataset.ref === `A${firstRow}`)).toBe(false);
  });
});

describe("SheetGrid — the ceiling (AC-48, AC-51)", () => {
  for (const viewport of [{ width: 1280, height: 720 }, { width: 1920, height: 1080 }]) {
    it(`holds ≤ ${GRID_MAX_RENDERED_CELLS} cells and < ${GRID_MAX_DOM_NODES} elements at every zoom, ${viewport.width} × ${viewport.height}`, async () => {
      Object.assign(size, viewport);
      const view = mountGrid(denseSheet(700, 140, 1));
      for (const percent of [100, 75, 50, 25, 10]) {
        await view.zoomTo(percent);
        expect(view.cells().length, `${percent}%`).toBeLessThanOrEqual(GRID_MAX_RENDERED_CELLS);
        expect(view.elements(), `${percent}%`).toBeLessThan(GRID_MAX_DOM_NODES);
      }
      Object.assign(size, { width: 1280, height: 720 });
    });
  }
});

describe("SheetGrid — the overview (AC-49, AC-50)", () => {
  it("switches to the canvas far out — only the active cell keeps an element — and back to DOM cells at 100%", async () => {
    const view = mountGrid(denseSheet(700, 40));
    expect(view.canvas()).toBeNull();
    expect(view.cells().length).toBeGreaterThan(1000);
    await view.zoomTo(10);
    expect(view.canvas()).not.toBeNull();
    expect(view.cells().map((element) => element.dataset.ref)).toEqual(["A1"]);
    expect(view.grid().getAttribute("aria-activedescendant")).toBe("sheet-cell-A1");
    // The headers are painted, not DOM, in overview.
    expect(view.grid().querySelectorAll('[data-testid="sheet-col-header"]')).toHaveLength(0);
    await view.zoomTo(100);
    expect(view.canvas()).toBeNull();
    expect(view.cells().length).toBeGreaterThan(1000);
    expect(view.grid().querySelectorAll('[data-testid="sheet-col-header"]').length).toBeGreaterThan(10);
  });

  it("carries data-calc-state, idle when recalculation is not wired", () => {
    const view = mountGrid(denseSheet(10, 5));
    expect(view.grid().getAttribute("data-calc-state")).toBe("idle");
  });
});
