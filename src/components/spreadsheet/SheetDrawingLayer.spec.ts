// SPEC-131 the sheet drawing layer: each painted drawing's element (roles,
// data-kind, names, aria-selected, handles), what it paints (an image with
// its crop, a shape's canvas and text, group members, the labelled
// placeholders, a foreign drawing's face), its events, and the frozen-pane
// split — a drawing straddling a freeze line mounted in both panes at the
// right offsets, one primary copy.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mount } from "@vue/test-utils";
import type { SheetDrawing, SheetDrawingContent } from "../../lib/spreadsheet/model";
import { unsupportedChartVisual } from "../../lib/spreadsheet/charts";
import { drawingPanes, layoutDrawingPanes, resizedRect, type PaneLayout, type ShownDrawing } from "../../views/spreadsheet/sheetDrawingLayout";
import SheetDrawingLayer from "./SheetDrawingLayer.vue";

beforeEach(() => {
  vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockReturnValue(null);
});
afterEach(() => vi.restoreAllMocks());

const anchor = { kind: "absolute" as const, posEmu: { x: 0, y: 0 }, extEmu: { cx: 0, cy: 0 } };

function drawing(id: string, name: string, content: SheetDrawingContent, extra: Partial<SheetDrawing> = {}): SheetDrawing {
  return { id, name, anchor, content, ...extra };
}

const imageSrc = (mediaId: string) => (mediaId === "xl/media/image1.png" ? "data:image/png;base64,AAAA" : "");

const NO_FREEZE: PaneLayout = { frozenW: 0, frozenH: 0, viewCols: 800, viewRows: 600, sx: 0, sy: 0 };

function mountPane(shown: ShownDrawing[], selectedId: string | null = null, layout: PaneLayout = NO_FREEZE, index = 0) {
  const pane = layoutDrawingPanes(shown, layout, selectedId)[index]!;
  return mount(SheetDrawingLayer, { props: { pane, zoom: 1, imageSrc } });
}

const rect = (left: number, top: number, width: number, height: number) => ({ left, top, width, height });

describe("SPEC-131 SheetDrawingLayer", () => {
  it("paints each drawing as a focusable sheet-drawing: img for an image, group otherwise, named by alt text else name, with data-kind", () => {
    const wrapper = mountPane([
      { drawing: drawing("dr1", "Picture 1", { kind: "image", mediaId: "xl/media/image1.png", crop: { l: 0.25, t: 0, r: 0.25, b: 0.5 } }, { altText: "Company logo" }), rect: rect(10, 20, 100, 50) },
      { drawing: drawing("dr2", "Rounded Rectangle 2", { kind: "shape", visual: { geometry: "roundRect" }, text: { paragraphs: [{ text: "Assumptions", align: "center" }], anchor: "middle", bold: true } }), rect: rect(200, 20, 120, 60) },
      { drawing: drawing("dr3", "Group 3", { kind: "group", members: [{ x: 0, y: 0, w: 0.5, h: 1, content: { kind: "shape", visual: { geometry: "rect" } } }, { x: 0.5, y: 0, w: 0.5, h: 1, content: { kind: "shape", visual: { geometry: "ellipse" } } }] }), rect: rect(10, 100, 80, 40) },
    ]);
    expect(wrapper.get('[data-testid="sheet-drawing-layer"]').attributes("data-pane")).toBe("main");
    const items = wrapper.findAll('[data-testid="sheet-drawing"]');
    expect(items.map((item) => [item.attributes("role"), item.attributes("aria-label"), item.attributes("data-kind"), item.attributes("data-drawing-name"), item.attributes("tabindex")])).toEqual([
      ["img", "Company logo", "image", "Picture 1", "0"],
      ["group", "Rounded Rectangle 2", "shape", "Rounded Rectangle 2", "0"],
      ["group", "Group 3", "group", "Group 3", "0"],
    ]);
    expect(items[0]!.attributes("style")).toContain("left: 10px; top: 20px; width: 100px; height: 50px");
    // The crop keeps the middle half across and the top half down: the
    // image is drawn at twice the box each way, shifted a quarter left.
    const img = items[0]!.get('[data-testid="sheet-drawing-image"]');
    expect(img.attributes("src")).toBe("data:image/png;base64,AAAA");
    expect(img.attributes("style")).toContain("left: -50px; top: 0px; width: 200px; height: 100px");
    expect(items[1]!.find('[data-testid="sheet-drawing-canvas"]').exists()).toBe(true);
    expect(items[1]!.get('[data-testid="sheet-drawing-text"]').text()).toBe("Assumptions");
    expect(items[1]!.get('[data-testid="sheet-drawing-text"]').attributes("style")).toContain("font-weight: bold");
    expect(items[2]!.findAll('[data-testid="sheet-drawing-member"]')).toHaveLength(2);
  });

  it("paints charts, face-less foreign drawings and undisplayable pictures as labelled placeholders, and a foreign face as its face", () => {
    const wrapper = mountPane([
      { drawing: drawing("dr1", "Chart 1", { kind: "chart", visual: unsupportedChartVisual(""), sourcePart: "xl/charts/chart1.xml" }), rect: rect(0, 0, 100, 60) },
      { drawing: drawing("dr2", "Check Box 2", { kind: "foreign", label: "Form control" }), rect: rect(0, 70, 100, 20) },
      { drawing: drawing("dr3", "Picture 3", { kind: "image", mediaId: "xl/media/image2.emf" }), rect: rect(0, 100, 50, 50) },
      { drawing: drawing("dr4", "Slicer 4", { kind: "foreign", label: "Slicer", face: { kind: "shape", visual: { geometry: "rect" }, text: { paragraphs: [{ text: "Region" }], anchor: "top" } } }), rect: rect(0, 160, 100, 80) },
    ]);
    const items = wrapper.findAll('[data-testid="sheet-drawing"]');
    expect(items.slice(0, 3).map((item) => item.get('[data-testid="sheet-drawing-placeholder"]').text())).toEqual(["Chart", "Form control", "Picture"]);
    expect(items.map((item) => item.attributes("data-kind"))).toEqual(["chart", "foreign", "image", "foreign"]);
    expect(items[3]!.find('[data-testid="sheet-drawing-placeholder"]').exists()).toBe(false);
    expect(items[3]!.get('[data-testid="sheet-drawing-text"]').text()).toBe("Region");
  });

  it("marks the selected drawing aria-selected with its eight handles, and emits its presses, keys and menu", async () => {
    const shown = [{ drawing: drawing("dr1", "Rectangle 1", { kind: "shape", visual: { geometry: "rect" } }, { rotationDeg: 30 }), rect: rect(40, 40, 80, 40) }];
    const wrapper = mountPane(shown, "dr1");
    const item = wrapper.get('[data-testid="sheet-drawing"]');
    expect(item.attributes("aria-selected")).toBe("true");
    expect(item.attributes("style")).toContain("rotate(30deg)");
    expect(wrapper.findAll('[data-testid^="sheet-drawing-handle-"]').map((handle) => handle.attributes("data-testid"))).toEqual(
      ["nw", "n", "ne", "e", "se", "s", "sw", "w"].map((at) => `sheet-drawing-handle-${at}`),
    );
    await item.trigger("pointerdown");
    expect(wrapper.emitted("press")?.[0]?.[0]).toBe("dr1");
    await wrapper.get('[data-testid="sheet-drawing-handle-se"]').trigger("pointerdown");
    expect(wrapper.emitted("handle-press")?.[0]?.slice(0, 2)).toEqual(["dr1", "se"]);
    await item.trigger("keydown", { key: "Delete" });
    expect(wrapper.emitted("key")?.[0]?.[0]).toBe("dr1");
    await item.trigger("contextmenu", { clientX: 15, clientY: 25 });
    expect(wrapper.emitted("menu")?.[0]).toEqual(["dr1", 15, 25]);
    const unselected = mountPane(shown, null);
    expect(unselected.get('[data-testid="sheet-drawing"]').attributes("aria-selected")).toBe("false");
    expect(unselected.find('[data-testid="sheet-drawing-handle-se"]').exists()).toBe(false);
  });

  it("with frozen panes mounts a drawing in every pane it meets, split at the freeze line, the scrolled copy offset by the scroll", () => {
    // Two frozen rows (40 px) and one frozen column (64 px), scrolled 100 px
    // down and 30 px across; the drawing spans y 20–80 in column B onwards.
    const layout: PaneLayout = { frozenW: 64, frozenH: 40, viewCols: 800, viewRows: 600, sx: 30, sy: 100 };
    expect(drawingPanes(layout).map((pane) => [pane.id, pane.left, pane.top, pane.originX, pane.originY])).toEqual([
      ["corner", 0, 0, 0, 0],
      ["top", 64, 0, 94, 0],
      ["left", 0, 40, 0, 140],
      ["main", 64, 40, 94, 140],
    ]);
    const shown = [{ drawing: drawing("dr1", "Rectangle 1", { kind: "shape", visual: { geometry: "rect" } }), rect: rect(100, 20, 200, 60) }];
    const panes = layoutDrawingPanes(shown, layout, "dr1");
    const copies = panes.flatMap((pane) => pane.items.map((item) => ({ pane: pane.id, x: item.x, y: item.y, primary: item.primary })));
    expect(copies).toEqual([
      { pane: "top", x: 6, y: 20, primary: true },
      { pane: "main", x: 6, y: -120, primary: false },
    ]);
    const top = mountPane(shown, "dr1", layout, 1);
    expect(top.get('[data-testid="sheet-drawing-layer"]').attributes("style")).toContain("left: 64px; top: 0px; width: 736px; height: 40px");
    expect(top.get('[data-testid="sheet-drawing"]').attributes("data-primary")).toBe("true");
    expect(top.find('[data-testid="sheet-drawing-handle-nw"]').exists()).toBe(true);
    const main = mountPane(shown, "dr1", layout, 3);
    const copy = main.get('[data-testid="sheet-drawing"]');
    expect([copy.attributes("data-primary"), copy.attributes("aria-hidden"), copy.attributes("tabindex")]).toEqual(["false", "true", "-1"]);
    expect(main.find('[data-testid="sheet-drawing-handle-nw"]').exists()).toBe(false);
    // Far outside the main pane's window plus one screen: not mounted.
    const far = layoutDrawingPanes([{ drawing: shown[0]!.drawing, rect: rect(100, 5000, 50, 50) }], layout, null);
    expect(far.map((pane) => pane.items.length)).toEqual([0, 0, 0, 0]);
  });

  it("resizes from a handle with the 8 px minimum, keeping the aspect on a corner when asked", () => {
    const start = rect(100, 100, 200, 100);
    expect(resizedRect(start, "e", 50, 0, false)).toEqual(rect(100, 100, 250, 100));
    expect(resizedRect(start, "nw", 300, 300, false)).toEqual(rect(292, 192, 8, 8));
    expect(resizedRect(start, "se", 100, 0, true)).toEqual(rect(100, 100, 300, 150));
    expect(resizedRect(start, "nw", -150, -150, false)).toEqual(rect(0, 0, 300, 200));
  });
});
