// SPEC-131 through the real workbook view at /sheets/:id, on a hand-built book
// (the readers are SPEC-131's other half): the drawing layer between the
// paper and the overlays, one layer per pane with a drawing straddling the
// freeze line mounted in both; a press selecting a picture (handles, focus on
// it, the cell selection unchanged), an arrow nudging it as one "Move image"
// step, Escape handing focus back to the grid; the context menu's Send to
// Back; and Alt Text… renaming it. SpreadsheetView.validation.spec.ts's
// harness: fresh modules per test, idb mocked, the view attached so focus is
// real, the grid given a 1280 × 720 viewport.
import { flushPromises, mount } from "@vue/test-utils";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createMemoryHistory, createRouter } from "vue-router";
import { EMU_PER_PX } from "../lib/spreadsheet/drawings";
import type { Sheet, SpreadsheetModel } from "../lib/spreadsheet/model";
import { SPREADSHEET_READER_VERSION } from "../lib/spreadsheet/open";

const cleanups: Array<() => void> = [];

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
  vi.spyOn(Element.prototype, "clientWidth", "get").mockReturnValue(1280);
  vi.spyOn(Element.prototype, "clientHeight", "get").mockReturnValue(720);
  vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockReturnValue(null);
});

afterEach(() => {
  for (const cleanup of cleanups.splice(0)) {
    cleanup();
  }
  vi.doUnmock("../lib/idb");
  vi.restoreAllMocks();
  document.body.innerHTML = "";
});

/** Two frozen rows (40 px); a picture at C4 (100 × 50 px) and a rectangle
 * from B1 to D4 straddling the freeze line, on top. */
function drawingBook(): SpreadsheetModel {
  const sheet: Sheet = {
    conditionalFormats: [],
    id: crypto.randomUUID(),
    name: "Sheet1",
    visibility: "visible",
    kind: "grid",
    rows: [],
    columns: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    merges: [],
    validations: [],
    frozen: { rows: 2, cols: 0 },
    drawings: [
      {
        id: "dr1",
        name: "Picture 1",
        anchor: { kind: "twoCell", editAs: "oneCell", from: { col: 2, colOffEmu: 0, row: 3, rowOffEmu: 0 }, to: { col: 4, colOffEmu: 0, row: 6, rowOffEmu: 0 }, extEmu: { cx: 100 * EMU_PER_PX, cy: 50 * EMU_PER_PX } },
        content: { kind: "image", mediaId: "xl/media/image1.png" },
      },
      {
        id: "dr2",
        name: "Rectangle 2",
        anchor: { kind: "twoCell", from: { col: 1, colOffEmu: 0, row: 0, rowOffEmu: 0 }, to: { col: 3, colOffEmu: 0, row: 3, rowOffEmu: 0 } },
        content: { kind: "shape", visual: { geometry: "rect" } },
      },
    ],
  };
  return {
    names: [],
    dxfs: [],
    kind: "spreadsheet",
    id: crypto.randomUUID(),
    title: "Drawings",
    format: "xlsx",
    readerVersion: SPREADSHEET_READER_VERSION,
    sheets: [sheet],
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    media: [{ mediaId: "xl/media/image1.png", contentType: "image/png", data: "iVBORw0KGgo=" }],
    styles: [{}],
    defaultStyleId: 0,
  };
}

async function mountBook() {
  const store = await import("../lib/documentStore");
  const { routes } = await import("../router");
  const { default: SpreadsheetView } = await import("./SpreadsheetView.vue");
  const book = drawingBook();
  store.storeSpreadsheet(book);
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(`/sheets/${book.id}`);
  const wrapper = mount(SpreadsheetView, { global: { plugins: [router] }, attachTo: document.body });
  cleanups.push(() => wrapper.unmount());
  await flushPromises();
  return wrapper;
}

type View = Awaited<ReturnType<typeof mountBook>>;

const drawingsNamed = (wrapper: View, name: string) => wrapper.findAll(`[data-testid="sheet-drawing"][data-drawing-name="${name}"]`);

async function key(target: Element, init: KeyboardEventInit): Promise<void> {
  target.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init }));
  await flushPromises();
}

describe("SPEC-131 the drawing layer in the workbook view", () => {
  it("paints one layer per pane, the straddling rectangle in both, the picture as its image", async () => {
    const wrapper = await mountBook();
    const grid = wrapper.get('[data-testid="spreadsheet-grid"]');
    expect(grid.findAll('[data-testid="sheet-drawing-layer"]').map((layer) => layer.attributes("data-pane"))).toEqual(["top", "main"]);
    const rectangle = drawingsNamed(wrapper, "Rectangle 2");
    expect(rectangle.map((copy) => copy.attributes("data-primary"))).toEqual(["true", "false"]);
    expect(rectangle[0]!.attributes("style")).toContain("left: 64px; top: 0px; width: 128px; height: 60px");
    // The main pane starts at the freeze line (40 px): the lower part sits 40 px up in it.
    expect(rectangle[1]!.attributes("style")).toContain("left: 64px; top: -40px");
    const [picture] = drawingsNamed(wrapper, "Picture 1");
    expect(picture!.attributes("role")).toBe("img");
    expect(picture!.get('[data-testid="sheet-drawing-image"]').attributes("src")).toBe("data:image/png;base64,iVBORw0KGgo=");
  });

  it("a press selects the picture with its handles and focus, the cell selection unchanged; an arrow nudges it as one 'Move image' step; Escape returns focus to the grid", async () => {
    const wrapper = await mountBook();
    const grid = wrapper.get('[data-testid="spreadsheet-grid"]');
    const active = grid.attributes("aria-activedescendant");
    const picture = () => drawingsNamed(wrapper, "Picture 1")[0]!;
    picture().element.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true, cancelable: true, button: 0 }));
    await flushPromises();
    expect(picture().attributes("aria-selected")).toBe("true");
    expect(picture().findAll('[data-testid^="sheet-drawing-handle-"]')).toHaveLength(8);
    expect(document.activeElement).toBe(picture().element);
    expect(grid.attributes("aria-activedescendant")).toBe(active);
    const left = picture().attributes("style");
    await key(picture().element, { key: "ArrowRight" });
    expect(picture().attributes("style")).not.toBe(left);
    expect(picture().attributes("style")).toContain("left: 129px");
    expect(wrapper.get('[data-testid="sheet-menu-undo"]').text()).toContain("Undo Move image");
    await key(picture().element, { key: "Escape" });
    expect(picture().attributes("aria-selected")).toBe("false");
    expect(grid.element.contains(document.activeElement)).toBe(true);
    expect(grid.attributes("aria-activedescendant")).toBe(active);
  });

  it("the context menu's Send to Back reorders the layer, and Alt Text… renames the drawing", async () => {
    const wrapper = await mountBook();
    const rectangle = () => drawingsNamed(wrapper, "Rectangle 2")[0]!;
    rectangle().element.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: 100, clientY: 50 }));
    await flushPromises();
    await wrapper.get('[data-testid="sheet-drawing-back"]').trigger("click");
    await flushPromises();
    const main = wrapper.findAll('[data-testid="sheet-drawing-layer"]')[1]!;
    expect(main.findAll('[data-testid="sheet-drawing"]').map((item) => item.attributes("data-drawing-name"))).toEqual(["Rectangle 2", "Picture 1"]);
    rectangle().element.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: 100, clientY: 50 }));
    await flushPromises();
    await wrapper.get('[data-testid="sheet-drawing-alt-text"]').trigger("click");
    await flushPromises();
    await wrapper.get('[data-testid="sheet-drawing-alt-text-input"]').setValue("Totals box");
    await wrapper.get('[data-testid="sheet-drawing-alt-text-ok"]').trigger("click");
    await flushPromises();
    expect(rectangle().attributes("aria-label")).toBe("Totals box");
    expect(wrapper.get('[data-testid="sheet-menu-undo"]').text()).toContain("Undo Alt text");
  });
});
