// SPEC-131 useSheetDrawings: selection, the keyboard (nudge, Shift nudge, the
// grid's top-left clamp, Delete, Escape, Tab cycling), the z-order and alt-text
// commands, Insert ▸ Image and Shapes — each command ONE labelled undo step
// that undo restores — over a hand-built model and the real undo stack.
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef, triggerRef, type EffectScope } from "vue";
import { IMAGE_TOO_LARGE_MESSAGE } from "../../lib/core/media";
import { drawingRectPx, EMU_PER_PX } from "../../lib/spreadsheet/drawings";
import { buildGridMetrics } from "../../lib/spreadsheet/geometry";
import type { SheetDrawing, SpreadsheetFormat, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { useSheetDrawings } from "./useSheetDrawings";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

const scopes: EffectScope[] = [];
afterEach(() => scopes.splice(0).forEach((scope) => scope.stop()));

const px = (n: number): number => n * EMU_PER_PX;

/** A picture at B2 (twoCell editAs oneCell, 100 × 50 px), a shape C5:E8, a
 * hidden shape, and a shape pinned at the grid's top-left. */
function sampleDrawings(): SheetDrawing[] {
  return [
    {
      id: "dr1",
      name: "Picture 1",
      anchor: { kind: "twoCell", editAs: "oneCell", from: { col: 1, colOffEmu: 0, row: 1, rowOffEmu: 0 }, to: { col: 2, colOffEmu: 0, row: 4, rowOffEmu: 0 }, extEmu: { cx: px(100), cy: px(50) } },
      content: { kind: "image", mediaId: "xl/media/image1.png" },
    },
    {
      id: "dr2",
      name: "Rectangle 2",
      anchor: { kind: "twoCell", from: { col: 2, colOffEmu: 0, row: 4, rowOffEmu: 0 }, to: { col: 4, colOffEmu: 0, row: 7, rowOffEmu: 0 } },
      content: { kind: "shape", visual: { geometry: "rect" } },
    },
    {
      id: "dr3",
      name: "Hidden 3",
      hidden: true,
      anchor: { kind: "oneCell", from: { col: 0, colOffEmu: 0, row: 0, rowOffEmu: 0 }, extEmu: { cx: px(10), cy: px(10) } },
      content: { kind: "shape", visual: { geometry: "ellipse" } },
    },
    {
      id: "dr4",
      name: "Oval 4",
      anchor: { kind: "oneCell", from: { col: 0, colOffEmu: 0, row: 0, rowOffEmu: 0 }, extEmu: { cx: px(20), cy: px(20) } },
      content: { kind: "shape", visual: { geometry: "ellipse" } },
    },
  ];
}

function harness(drawings: SheetDrawing[] = sampleDrawings(), format: SpreadsheetFormat = "xlsx") {
  const model: SpreadsheetModel = createBlankWorkbook();
  model.format = format;
  model.sheets = [{ ...model.sheets[0]!, drawings }];
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const activeSheet = computed(() => workbook.value?.sheets[0] ?? null);
  const selection = ref({ active: { row: 2, col: 1 }, ranges: [{ top: 2, left: 1, bottom: 2, right: 1 }], activeRangeIndex: 0 });
  const busy = ref(false);
  const toasts: string[] = [];
  const hooks = { focusGrid: vi.fn(), focusDrawing: vi.fn(), visiblePx: () => ({ width: 1000, height: 600 }), reveal: vi.fn() };
  const scope = effectScope();
  scopes.push(scope);
  const { undo, drawings: controller } = scope.run(() => {
    const stack = useSpreadsheetUndo({ model: workbook, recalc: (_keys, done) => done([]), onApplied: () => triggerRef(workbook), activateSheet: () => {} });
    const sheetDrawings = useSheetDrawings({
      book: { workbook, activeSheet },
      undo: stack,
      selection: { selection: computed(() => selection.value) },
      sheetMetrics: { metrics: computed(() => buildGridMetrics(activeSheet.value!)) },
      syncInto: vi.fn(),
      busy: () => busy.value,
      toast: (message) => toasts.push(message),
      loadImage: async () => null,
    });
    return { undo: stack, drawings: sheetDrawings };
  })!;
  controller.bindGrid(hooks);
  const sheet = () => workbook.value!.sheets[0]!;
  const rectOf = (id: string) => drawingRectPx(sheet().drawings!.find((d) => d.id === id)!.anchor, buildGridMetrics(sheet()));
  return { drawings: controller, undo, sheet, rectOf, hooks, toasts, busy, workbook };
}

function key(k: string, init: KeyboardEventInit = {}): KeyboardEvent {
  return new KeyboardEvent("keydown", { key: k, cancelable: true, ...init });
}

/** A PNG's signature and IHDR, enough for decodeImageDimensions. */
function pngBytes(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(33);
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82]);
  new DataView(bytes.buffer).setUint32(16, width);
  new DataView(bytes.buffer).setUint32(20, height);
  return bytes;
}

describe("SPEC-131 useSheetDrawings", () => {
  it("paints the non-hidden drawings back to front and selects one, focusing its element; a hidden drawing cannot be selected", async () => {
    const h = harness();
    expect(h.drawings.shown.value.map((entry) => entry.drawing.id)).toEqual(["dr1", "dr2", "dr4"]);
    h.drawings.select("dr2");
    expect(h.drawings.selectedId.value).toBe("dr2");
    await nextTick();
    expect(h.hooks.focusDrawing).toHaveBeenCalledWith("dr2");
    h.drawings.select("dr3");
    expect(h.drawings.selectedId.value).toBeNull();
  });

  it("arrows nudge 1 px and Shift+arrows 10 px, each one 'Move …' step; a nudge past the grid's top-left changes nothing", () => {
    const h = harness();
    const start = h.rectOf("dr1");
    h.drawings.select("dr1");
    const right = key("ArrowRight");
    h.drawings.onDrawingKeydown("dr1", right);
    expect(right.defaultPrevented).toBe(true);
    expect(h.rectOf("dr1").left).toBeCloseTo(start.left + 1, 3);
    expect(h.undo.undoLabel.value).toBe("Move image");
    h.drawings.onDrawingKeydown("dr1", key("ArrowDown", { shiftKey: true }));
    expect(h.rectOf("dr1").top).toBeCloseTo(start.top + 10, 3);
    expect(h.rectOf("dr1").width).toBeCloseTo(100, 3);
    h.undo.undo();
    h.undo.undo();
    expect(h.rectOf("dr1")).toEqual(start);
    // dr4 sits at A1: left and up are clamped, so nothing is recorded.
    h.drawings.select("dr4");
    h.drawings.onDrawingKeydown("dr4", key("ArrowLeft"));
    h.drawings.onDrawingKeydown("dr4", key("ArrowUp", { shiftKey: true }));
    expect(h.rectOf("dr4").left).toBe(0);
    expect(h.undo.undoLabel.value).toBe("");
  });

  it("Delete removes the selected drawing as one 'Delete …' step and returns focus to the grid; undo restores it", () => {
    const h = harness();
    const before = h.sheet().drawings;
    h.drawings.select("dr2");
    h.drawings.onDrawingKeydown("dr2", key("Delete"));
    expect(h.sheet().drawings!.map((d) => d.id)).toEqual(["dr1", "dr3", "dr4"]);
    expect(h.sheet().drawingsDirty).toBe(true);
    expect(h.undo.undoLabel.value).toBe("Delete shape");
    expect(h.drawings.selectedId.value).toBeNull();
    expect(h.hooks.focusGrid).toHaveBeenCalled();
    h.undo.undo();
    expect(h.sheet().drawings).toBe(before);
  });

  it("Escape deselects and returns focus to the grid; Tab and Shift+Tab cycle the painted drawings in z-order, wrapping", () => {
    const h = harness();
    h.drawings.select("dr1");
    const order: (string | null)[] = [];
    for (let i = 0; i < 3; i += 1) {
      h.drawings.onDrawingKeydown(h.drawings.selectedId.value!, key("Tab"));
      order.push(h.drawings.selectedId.value);
    }
    expect(order).toEqual(["dr2", "dr4", "dr1"]);
    h.drawings.onDrawingKeydown("dr1", key("Tab", { shiftKey: true }));
    expect(h.drawings.selectedId.value).toBe("dr4");
    expect(h.hooks.reveal).toHaveBeenCalled();
    h.drawings.onDrawingKeydown("dr4", key("Escape"));
    expect(h.drawings.selectedId.value).toBeNull();
    expect(h.hooks.focusGrid).toHaveBeenCalled();
    expect(h.undo.canUndo.value).toBe(false);
  });

  it("the z-order moves reorder Sheet.drawings, each its own step; a move that cannot move records nothing", () => {
    const h = harness();
    h.drawings.select("dr1");
    expect(h.drawings.order.value).toEqual({ selected: true, atFront: false, atBack: true });
    h.drawings.runAction("front");
    expect(h.sheet().drawings!.map((d) => d.id)).toEqual(["dr2", "dr3", "dr4", "dr1"]);
    expect(h.undo.undoLabel.value).toBe("Bring to front");
    h.drawings.runAction("forward");
    expect(h.undo.undoLabel.value).toBe("Bring to front");
    h.drawings.runAction("backward");
    expect(h.sheet().drawings!.map((d) => d.id)).toEqual(["dr2", "dr3", "dr1", "dr4"]);
    expect(h.undo.undoLabel.value).toBe("Send backward");
    h.drawings.reorder("back");
    expect(h.sheet().drawings!.map((d) => d.id)).toEqual(["dr1", "dr2", "dr3", "dr4"]);
    expect(h.undo.undoLabel.value).toBe("Send to back");
    h.drawings.reorder("forward");
    expect(h.undo.undoLabel.value).toBe("Bring forward");
    h.undo.undo();
    expect(h.sheet().drawings!.map((d) => d.id)).toEqual(["dr1", "dr2", "dr3", "dr4"]);
  });

  it("Alt Text… opens with the current text; OK writes it as one 'Alt text' step and an empty field removes it", () => {
    const h = harness();
    h.drawings.openContextMenu("dr2", 10, 20);
    expect(h.drawings.contextMenu.value).toEqual({ x: 10, y: 20 });
    expect(h.drawings.selectedId.value).toBe("dr2");
    h.drawings.runAction("alt-text");
    expect(h.drawings.contextMenu.value).toBeNull();
    expect(h.drawings.altText.value).toEqual({ id: "dr2", initial: "", name: "Rectangle 2" });
    h.drawings.confirmAltText("Quarterly total");
    expect(h.sheet().drawings![1]!.altText).toBe("Quarterly total");
    expect(h.undo.undoLabel.value).toBe("Alt text");
    h.drawings.openAltText();
    expect(h.drawings.altText.value?.initial).toBe("Quarterly total");
    h.drawings.confirmAltText("");
    expect("altText" in h.sheet().drawings![1]!).toBe(false);
    h.undo.undo();
    expect(h.sheet().drawings![1]!.altText).toBe("Quarterly total");
  });

  it("a gesture's move and resize are one 'Move …' / 'Resize …' step each, re-anchored with the drawing's behaviour", () => {
    const h = harness();
    h.drawings.setRect("dr2", { ...h.rectOf("dr2"), left: 300, top: 200 }, "Move");
    expect(h.undo.undoLabel.value).toBe("Move shape");
    expect(h.sheet().drawings![1]!.anchor.kind).toBe("twoCell");
    expect(h.rectOf("dr2")).toMatchObject({ left: 300, top: 200 });
    h.drawings.setRect("dr1", { ...h.rectOf("dr1"), width: 200, height: 100 }, "Resize");
    expect(h.undo.undoLabel.value).toBe("Resize image");
    expect(h.sheet().drawings![0]!.anchor).toMatchObject({ kind: "twoCell", editAs: "oneCell", extEmu: { cx: px(200), cy: px(100) } });
  });

  it("Insert ▸ Image puts the picture at the active cell as one 'Insert image' step, selected; an over-limit file toasts and changes nothing", async () => {
    const h = harness([]);
    const bytes = pngBytes(300, 150);
    expect(await h.drawings.insertImageFile({ type: "image/png", size: bytes.length, arrayBuffer: async () => bytes.buffer })).toBe(true);
    const [picture] = h.sheet().drawings!;
    expect(picture).toMatchObject({ name: "Picture 1", content: { kind: "image", mediaId: "xl/media/image1.png" }, anchor: { kind: "twoCell", editAs: "oneCell" } });
    expect(picture!.anchor.from).toMatchObject({ col: 1, row: 2 });
    expect(h.workbook.value!.media).toEqual([expect.objectContaining({ mediaId: "xl/media/image1.png", contentType: "image/png" })]);
    expect(h.rectOf(picture!.id)).toMatchObject({ width: 300, height: 150 });
    expect(h.undo.undoLabel.value).toBe("Insert image");
    expect(h.drawings.selectedId.value).toBe(picture!.id);
    const huge = { type: "image/png", size: 11 * 1024 * 1024, arrayBuffer: vi.fn() };
    expect(await h.drawings.insertImageFile(huge)).toBe(false);
    expect(h.toasts).toEqual([IMAGE_TOO_LARGE_MESSAGE]);
    expect(huge.arrayBuffer).not.toHaveBeenCalled();
    expect(h.sheet().drawings).toHaveLength(1);
  });

  it("Insert ▸ Shapes arms a geometry that Escape disarms; a drawn box lands as one 'Insert shape' step, selected", () => {
    const h = harness([]);
    h.drawings.openGallery();
    expect(h.drawings.galleryOpen.value).toBe(true);
    h.drawings.pickShape("ellipse");
    expect(h.drawings.galleryOpen.value).toBe(false);
    expect(h.drawings.armed.value).toBe("ellipse");
    window.dispatchEvent(key("Escape"));
    expect(h.drawings.armed.value).toBeNull();
    h.drawings.pickShape("rect");
    h.drawings.drawShape("rect", { left: 70, top: 30, width: 120, height: 60 });
    expect(h.drawings.armed.value).toBeNull();
    const [shape] = h.sheet().drawings!;
    expect(shape).toMatchObject({ name: "Rectangle 1", anchor: { kind: "twoCell", editAs: "twoCell" }, content: { kind: "shape", visual: { geometry: "rect" } } });
    expect(h.undo.undoLabel.value).toBe("Insert shape");
    expect(h.drawings.selectedId.value).toBe(shape!.id);
  });

  it("offers no insertion on a CSV workbook, and nothing runs while a cell is being edited", async () => {
    const csv = harness([], "csv");
    expect(csv.drawings.canInsert.value).toBe(false);
    csv.drawings.openGallery();
    expect(csv.drawings.galleryOpen.value).toBe(false);
    const bytes = pngBytes(10, 10);
    expect(await csv.drawings.insertImageFile({ type: "image/png", size: bytes.length, arrayBuffer: async () => bytes.buffer })).toBe(false);
    const h = harness();
    h.busy.value = true;
    h.drawings.select("dr1");
    h.drawings.remove("dr1");
    expect(h.drawings.selectedId.value).toBeNull();
    expect(h.sheet().drawings).toHaveLength(4);
  });
});
