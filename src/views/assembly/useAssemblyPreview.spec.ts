// SPEC-109: the assembly preview — the sheet at fit-to-window; a press that
// picks the topmost placed item; overlay items dragged on the canvas (Shift
// keeping one axis) and nudged with the arrow keys (1 mm, Shift: 1 cm); and
// the preview drawn by the very composeToCanvas the export calls, at its own
// smaller pixel size — no second drawing path.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { effectScope, nextTick } from "vue";
import { flushPromises } from "@vue/test-utils";
import type { DecodedImage } from "../../lib/media/renderEdits";
import {
  constrainDrag,
  hitTest,
  nudgeFor,
  PREVIEW_FALLBACK_WIDTH_PX,
  PREVIEW_MAX_EDGE_PX,
  previewDisplaySize,
  previewPixelSize,
} from "./useAssemblyPreview";
import type { PlacedItem } from "../../lib/core/assemblyLayout";

/** A 10 cm square placed at `xCm`. */
function squareAt(index: number, xCm: number): PlacedItem {
  return { index, xCm, yCm: 0, widthCm: 10, heightCm: 10 } as PlacedItem;
}

function context() {
  return { save() {}, restore() {}, setTransform() {}, clearRect() {}, fillRect() {}, drawImage: vi.fn() };
}

function pointer(xPx: number, yPx: number, shiftKey = false) {
  return { clientX: xPx, clientY: yPx, button: 0, pointerId: 1, shiftKey, preventDefault() {} } as unknown as PointerEvent;
}

function key(name: string, shiftKey = false) {
  return { key: name, shiftKey, altKey: false, ctrlKey: false, metaKey: false, preventDefault: vi.fn() } as unknown as KeyboardEvent;
}

describe("the preview's arithmetic", () => {
  const A4_LANDSCAPE = { widthCm: 29.7, heightCm: 21, content: { xCm: 1, yCm: 1, widthCm: 27.7, heightCm: 19 } };

  it("fits the sheet in the window, either axis limiting", () => {
    expect(previewDisplaySize(A4_LANDSCAPE, { widthPx: 600, heightPx: 600 })).toEqual({ widthPx: 600, heightPx: 424 });
    expect(previewDisplaySize(A4_LANDSCAPE, { widthPx: 1000, heightPx: 300 })).toEqual({ widthPx: 424, heightPx: 300 });
    // Unmeasured: a fixed width at the sheet's aspect.
    expect(previewDisplaySize(A4_LANDSCAPE, { widthPx: 0, heightPx: 0 })).toEqual({ widthPx: PREVIEW_FALLBACK_WIDTH_PX, heightPx: 339 });
  });

  it("draws at the device pixel ratio, capped", () => {
    expect(previewPixelSize({ widthPx: 600, heightPx: 424 }, 2)).toEqual({ widthPx: 1200, heightPx: 848 });
    const capped = previewPixelSize({ widthPx: 1800, heightPx: 1273 }, 2);
    expect(Math.max(capped.widthPx, capped.heightPx)).toBe(PREVIEW_MAX_EDGE_PX);
  });

  it("picks the topmost item under a point", () => {
    const placed = [squareAt(0, 0), squareAt(1, 5)];
    expect(hitTest(placed, 7, 5)?.index).toBe(1);
    expect(hitTest(placed, 2, 5)?.index).toBe(0);
    expect(hitTest(placed, 20, 5)).toBeNull();
  });

  it("locks a Shift-drag to the axis it moved along most", () => {
    expect(constrainDrag(3, 1, false)).toEqual({ dxCm: 3, dyCm: 1 });
    expect(constrainDrag(3, 1, true)).toEqual({ dxCm: 3, dyCm: 0 });
    expect(constrainDrag(-1, -4, true)).toEqual({ dxCm: 0, dyCm: -4 });
  });

  it("nudges by 1 mm, or 1 cm with Shift", () => {
    expect(nudgeFor("ArrowRight", false)).toEqual({ dxCm: 0.1, dyCm: 0 });
    expect(nudgeFor("ArrowUp", true)).toEqual({ dxCm: 0, dyCm: -1 });
    expect(nudgeFor("Enter", false)).toBeNull();
  });
});

describe("the preview controller", () => {
  beforeEach(() => {
    vi.resetModules();
    localStorage.clear();
    vi.doMock("../../lib/idb", () => ({
      loadAllRecords: vi.fn().mockResolvedValue([]),
      putRecord: vi.fn().mockResolvedValue(undefined),
      deleteRecord: vi.fn().mockResolvedValue(undefined),
    }));
  });

  afterEach(() => {
    vi.doUnmock("../../lib/idb");
    vi.doUnmock("../../lib/media/composeCanvas");
    vi.restoreAllMocks();
  });

  const decode = vi.fn(async (): Promise<DecodedImage> => ({ image: {} as CanvasImageSource, widthPx: 100, heightPx: 100, release: () => {} }));

  async function setup() {
    const calls: Array<{ width: number; height: number }> = [];
    vi.doMock("../../lib/media/composeCanvas", async (importOriginal) => {
      const original = await importOriginal<typeof import("../../lib/media/composeCanvas")>();
      return {
        ...original,
        composeToCanvas: vi.fn((...args: Parameters<typeof original.composeToCanvas>) => {
          calls.push({ width: args[0].width, height: args[0].height });
          return original.composeToCanvas(...args);
        }),
      };
    });
    const store = await import("../../lib/documentStore");
    const { assembleMediaModel } = await import("../../lib/media/open");
    const fixtures = await import("../../lib/media/fixtures/build");
    const { useAssembly } = await import("./useAssembly");
    const { useAssemblyPreview } = await import("./useAssemblyPreview");
    // Two 100 × 100 squares.
    const bytes = await fixtures.buildPngFixtureBytes();
    const ids: string[] = [];
    for (const title of ["back", "front"]) {
      const model = await assembleMediaModel(title, "png", bytes);
      model.image = { widthPx: 100, heightPx: 100 };
      store.storeMedia(model);
      ids.push(model.id);
    }
    const steps = {
      createCanvas: vi.fn((width: number, height: number) => ({ width, height, getContext: () => context() }) as unknown as HTMLCanvasElement),
      decode,
      encode: vi.fn(async () => bytes),
      between: vi.fn(async () => {}),
      release: vi.fn(),
    };
    const scope = effectScope();
    const { assembly, preview } = scope.run(() => {
      const controller = useAssembly({ steps });
      return { assembly: controller, preview: useAssemblyPreview({ assembly: controller, decode }) };
    })!;
    assembly.tray.add(ids);
    // A4 landscape, 1 cm margins all round: a 27.7 × 19 cm content box,
    // shown 297 × 210 px — 10 px a centimetre.
    assembly.sheet.setOrientation("landscape");
    assembly.sheet.setMargin("top", "1");
    const canvas = document.createElement("canvas");
    vi.spyOn(canvas, "getContext").mockReturnValue(context() as unknown as CanvasRenderingContext2D);
    vi.spyOn(canvas, "getBoundingClientRect").mockReturnValue({ left: 0, top: 0, width: 297, height: 210 } as DOMRect);
    preview.attachCanvas(canvas);
    await flushPromises();
    return { assembly, preview, canvas, calls, steps };
  }

  it("draws through composeToCanvas — the function the export calls — at its own size", async () => {
    const { assembly, calls, preview } = await setup();
    expect(preview.renders.value).toBeGreaterThan(0);
    const previewSize = calls.at(-1);
    expect(previewSize).toEqual(preview.pixels.value ? { width: preview.pixels.value.widthPx, height: preview.pixels.value.heightPx } : null);
    await assembly.out.save();
    // The export: the same function, at sheetPixelSize (A4 landscape, 150 dpi).
    expect(calls.at(-1)).toEqual({ width: 1754, height: 1240 });
  });

  it("selects the item under a press; in a grid nothing moves", async () => {
    const { assembly, preview } = await setup();
    assembly.setColumns("2");
    const [, right] = assembly.placed.value;
    const before = JSON.stringify(assembly.placed.value);
    preview.onPointerDown(pointer((right.xCm + 1) * 10, (right.yCm + 1) * 10));
    expect(assembly.tray.selected.value?.title).toBe("front");
    preview.onPointerMove(pointer(10, 10));
    preview.onPointerUp(pointer(10, 10));
    expect(JSON.stringify(assembly.placed.value)).toBe(before);
  });

  it("drags the topmost overlay item, Shift keeping it on one axis", async () => {
    const { assembly, preview } = await setup();
    assembly.setLayout("overlay");
    // Both items are contained, centred: the press hits the top one.
    preview.onPointerDown(pointer(148, 105));
    expect(assembly.tray.selected.value?.title).toBe("front");
    // 27.7 px right and 19 px down: 10% of the content box each way.
    preview.onPointerMove(pointer(148 + 27.7, 105 + 19));
    expect(assembly.tray.entries.value[1].transform).toMatchObject({ xPercent: 60, yPercent: 60 });
    // Shift: only the larger component survives.
    preview.onPointerMove(pointer(148 + 27.7, 105 + 9.5, true));
    expect(assembly.tray.entries.value[1].transform).toMatchObject({ xPercent: 60, yPercent: 50 });
    preview.onPointerUp(pointer(0, 0));
    // After release, moving does nothing.
    preview.onPointerMove(pointer(0, 0));
    expect(assembly.tray.entries.value[1].transform).toMatchObject({ xPercent: 60, yPercent: 50 });
    expect(assembly.tray.entries.value[0].transform).toMatchObject({ xPercent: 50, yPercent: 50 });
  });

  it("nudges the selected overlay item with the arrow keys", async () => {
    const { assembly, preview } = await setup();
    assembly.setLayout("overlay");
    assembly.select(assembly.tray.entries.value[0].key);
    const right = key("ArrowRight");
    preview.onKeydown(right);
    expect(right.preventDefault).toHaveBeenCalled();
    const [item] = assembly.placed.value;
    // 1 mm of a 27.7 cm box.
    expect(assembly.tray.entries.value[0].transform.xPercent).toBe(50.361);
    preview.onKeydown(key("ArrowDown", true));
    expect(assembly.tray.entries.value[0].transform.yPercent).toBe(55.263);
    await nextTick();
    expect(assembly.placed.value[0].yCm).toBeGreaterThan(item.yCm);
    // Outside overlay, arrows are not the preview's.
    assembly.setLayout("grid");
    const ignored = key("ArrowLeft");
    preview.onKeydown(ignored);
    expect(ignored.preventDefault).not.toHaveBeenCalled();
  });
});
