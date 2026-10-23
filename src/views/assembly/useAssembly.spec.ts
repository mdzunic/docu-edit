// SPEC-109: the assembly workspace's controller, run in an effect scope over
// a real (IndexedDB-mocked) document store. The tray — reorder by drop and by
// keyboard, duplicate, remove, the 100-item limit; the tray order as the
// layout order; the persisted plan and its restoration, with the notice for
// an item deleted or trashed since; the refusal paths (the canvas ceiling and
// MAX_OUTPUT_EDGE_PX before any work, JPEG and transparency); cancellation
// leaving no file and no record; the produced image saved as a NEW document
// read back from its bytes; and every source deep-frozen throughout.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { effectScope, nextTick } from "vue";
import type { MediaModel } from "../../lib/media/model";
import type { DecodedImage } from "../../lib/media/renderEdits";

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  vi.doMock("../../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
});

// Every workspace a test opens is disposed after it. A live one's debounced
// plan write fires into whichever test is running 250 ms later, which then
// restores that plan's sheet — a Custom 200 × 20 cm one, say.
const scopes: (() => void)[] = [];

afterEach(() => {
  scopes.splice(0).forEach((stop) => stop());
  vi.doUnmock("../../lib/idb");
  vi.restoreAllMocks();
  vi.useRealTimers();
});

function deepFreeze<T>(value: T): T {
  if (value && typeof value === "object" && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const inner of Object.values(value)) {
      deepFreeze(inner);
    }
  }
  return value;
}

/** A 2-D context stub: enough for composeToCanvas, recording draws. */
function stubContext(draws: unknown[][]) {
  return {
    save() {},
    restore() {},
    setTransform() {},
    clearRect() {},
    fillRect(...args: unknown[]) {
      draws.push(["fillRect", ...args]);
    },
    drawImage(...args: unknown[]) {
      draws.push(["drawImage", ...args]);
    },
  };
}

async function setup(stepOverrides: Record<string, unknown> = {}) {
  const store = await import("../../lib/documentStore");
  const { assembleMediaModel } = await import("../../lib/media/open");
  const fixtures = await import("../../lib/media/fixtures/build");
  const { useAssembly } = await import("./useAssembly");
  const persistence = await import("./useAssemblyPersistence");
  const { loadPreferences } = await import("../../lib/preferences");
  const { MAX_ASSEMBLY_ITEMS } = await import("../../lib/core/assemblyLayout");
  const png = await fixtures.buildPngFixtureBytes();
  const jpeg = fixtures.buildJpegFixtureBytes();

  async function seed(title: string, format: "png" | "jpeg" = "png"): Promise<MediaModel> {
    const model = await assembleMediaModel(title, format, format === "png" ? png : jpeg);
    store.storeMedia(model);
    return model;
  }

  const draws: unknown[][] = [];
  const canvases: Array<{ width: number; height: number }> = [];
  const decoded: string[] = [];
  const steps = {
    createCanvas: vi.fn((width: number, height: number) => {
      const canvas = { width, height, getContext: () => stubContext(draws) };
      canvases.push(canvas);
      return canvas as unknown as HTMLCanvasElement;
    }),
    decode: vi.fn(async (part: { mediaId: string }): Promise<DecodedImage> => {
      decoded.push(part.mediaId);
      return { image: { part: part.mediaId } as unknown as CanvasImageSource, widthPx: 64, heightPx: 40, release: () => {} };
    }),
    // The "encoded" output is the PNG fixture: the summary must come from it.
    encode: vi.fn(async () => png),
    canvasCeiling: vi.fn(() => 268_435_456),
    between: vi.fn(async () => {}),
    download: vi.fn(),
    release: vi.fn((canvas: { width: number; height: number }) => {
      canvas.width = 0;
      canvas.height = 0;
    }),
    ...stepOverrides,
  };

  function open() {
    const scope = effectScope();
    scopes.push(() => scope.stop());
    const assembly = scope.run(() => useAssembly({ steps }))!;
    return { assembly, stop: () => scope.stop() };
  }

  return { store, seed, open, steps, draws, canvases, decoded, persistence, loadPreferences, MAX_ASSEMBLY_ITEMS, png };
}

const titles = (assembly: { tray: { entries: { value: Array<{ title: string }> } } }) =>
  assembly.tray.entries.value.map((entry) => entry.title);

describe("the tray", () => {
  it("holds library images in the order chosen, and that order is the layout order", async () => {
    const { seed, open, store } = await setup();
    const a = await seed("a");
    const b = await seed("b", "jpeg");
    store.getOrCreateDocument("doc-1");
    const { assembly } = open();
    expect(assembly.tray.add([b.id, "doc-1", "missing", a.id])).toBe(2);
    expect(titles(assembly)).toEqual(["b", "a"]);
    expect(assembly.tray.entries.value.map((entry) => [entry.widthPx, entry.heightPx])).toEqual([
      [160, 120],
      [64, 40],
    ]);
    assembly.setColumns("2");
    // Cell 1 holds the first tray entry; cell 2 the second.
    const [first, second] = assembly.placed.value;
    expect([first.index, second.index]).toEqual([0, 1]);
    expect(first.xCm).toBeLessThan(second.xCm);
    expect(assembly.items.value.map((item) => item.id)).toEqual([b.id, a.id]);
  });

  it("reorders by keyboard steps and by a drop, announcing each move", async () => {
    const { seed, open } = await setup();
    const ids = [(await seed("a")).id, (await seed("b")).id, (await seed("c")).id];
    const { assembly } = open();
    assembly.tray.add(ids);
    const [a, , c] = assembly.tray.entries.value;
    expect(assembly.tray.move(a.key, 1)).toBe(true);
    expect(titles(assembly)).toEqual(["b", "a", "c"]);
    expect(assembly.tray.announcement.value).toBe("“a” moved to position 2 of 3.");
    expect(assembly.tray.move(c.key, 1)).toBe(false);
    expect(assembly.tray.announcement.value).toBe("“c” is already last.");
    expect(assembly.tray.moveTo(c.key, 0)).toBe(true);
    expect(titles(assembly)).toEqual(["c", "b", "a"]);
    expect(assembly.tray.announcement.value).toBe("“c” moved to position 1 of 3.");
    // The layout follows.
    expect(assembly.items.value.map((item) => item.id)).toEqual([ids[2], ids[1], ids[0]]);
  });

  it("duplicates an entry under its own key, and removes one", async () => {
    const { seed, open } = await setup();
    const a = await seed("a");
    const { assembly } = open();
    assembly.tray.add([a.id]);
    const [entry] = assembly.tray.entries.value;
    expect(assembly.tray.duplicate(entry.key)).toBe(true);
    const [, copy] = assembly.tray.entries.value;
    expect(copy.id).toBe(a.id);
    expect(copy.key).not.toBe(entry.key);
    assembly.tray.remove(entry.key);
    expect(assembly.tray.entries.value.map((item) => item.key)).toEqual([copy.key]);
    expect(assembly.tray.announcement.value).toBe("Removed “a”.");
  });

  it("refuses past 100 entries, adding as many as fit", async () => {
    const { seed, open, MAX_ASSEMBLY_ITEMS } = await setup();
    const a = await seed("a");
    const { assembly } = open();
    expect(assembly.tray.add(Array.from({ length: MAX_ASSEMBLY_ITEMS + 3 }, () => a.id))).toBe(100);
    expect(assembly.tray.entries.value).toHaveLength(100);
    expect(assembly.tray.refusal.value).toBe("A tray holds at most 100 images. 3 images were not added.");
    expect(assembly.tray.duplicate(assembly.tray.entries.value[0].key)).toBe(false);
    expect(assembly.tray.refusal.value).toBe("A tray holds at most 100 images.");
    expect(assembly.tray.entries.value).toHaveLength(100);
  });
});

describe("the persisted plan", () => {
  it("persists ids, layout, transforms, sheet and output under one preferences key — never pixels", async () => {
    vi.useFakeTimers();
    const { seed, open, loadPreferences, persistence } = await setup();
    const a = await seed("shot");
    const { assembly } = open();
    assembly.tray.add([a.id]);
    assembly.setLayout("overlay");
    assembly.select(assembly.tray.entries.value[0].key);
    assembly.setTransformField("opacity", "40");
    assembly.sheet.setOrientation("landscape");
    assembly.setFormat("jpeg");
    assembly.setTitle("Pair");
    expect(loadPreferences().assemblyPlan).toBeNull();
    vi.advanceTimersByTime(persistence.PERSIST_DELAY_MS);
    const plan = loadPreferences().assemblyPlan;
    expect(plan?.items).toEqual([{ id: a.id, title: "shot", transform: { xPercent: 50, yPercent: 50, scalePercent: 100, opacity: 40, blend: "normal" } }]);
    expect(plan?.layout.layout).toBe("overlay");
    expect(plan?.sheet?.orientation).toBe("landscape");
    expect(plan?.output).toMatchObject({ format: "jpeg", title: "Pair", quality: 90 });
    const stored = localStorage.getItem("docu-edit:preferences") ?? "";
    expect(stored).toContain('"assemblyPlan"');
    expect(stored).not.toMatch(/widthPx|base64|data:image/);
  });

  it("restores the plan on return", async () => {
    const { seed, open } = await setup();
    const a = await seed("a");
    const b = await seed("b");
    const first = open();
    first.assembly.tray.add([a.id, b.id]);
    first.assembly.tray.move(first.assembly.tray.entries.value[0].key, 1);
    first.assembly.setColumns("2");
    first.assembly.sheet.setMarginPreset("Narrow");
    first.stop();
    const again = open().assembly;
    expect(titles(again)).toEqual(["b", "a"]);
    expect(again.layout.columns).toBe(2);
    expect(again.sheet.marginPreset.value).toBe("Narrow");
    expect(again.notice.value).toBe("");
  });

  it("drops an item deleted or trashed since, with a one-line notice naming it", async () => {
    const { seed, open, store, loadPreferences } = await setup();
    const a = await seed("kept");
    const b = await seed("gone");
    const c = await seed("binned");
    const first = open();
    first.assembly.tray.add([a.id, b.id, c.id]);
    first.stop();
    store.deleteDocument(b.id);
    store.trashDocument(c.id);
    const { assembly } = open();
    expect(titles(assembly)).toEqual(["kept"]);
    expect(assembly.notice.value).toBe("“gone”, “binned” are no longer in the library and were removed from the tray.");
    expect(loadPreferences().assemblyPlan?.items.map((item) => item.id)).toEqual([a.id]);
  });

  it("starts afresh from 'Combine…', saying when the selection was cut to 100", async () => {
    const { seed, open, persistence } = await setup();
    const a = await seed("a");
    const b = await seed("b");
    persistence.startAssembly([b.id, "not-an-image", a.id]);
    expect(titles(open().assembly)).toEqual(["b", "a"]);
    persistence.startAssembly(Array.from({ length: 101 }, () => a.id));
    const { assembly } = open();
    expect(assembly.tray.entries.value).toHaveLength(100);
    expect(assembly.notice.value).toBe("A tray holds at most 100 images; the first 100 of 101 were added.");
  });

  it("writes a pending change when the page is hidden (a reload or a closed tab)", async () => {
    vi.useFakeTimers();
    const { seed, open, loadPreferences } = await setup();
    const a = await seed("a");
    open().assembly.tray.add([a.id]);
    expect(loadPreferences().assemblyPlan).toBeNull();
    window.dispatchEvent(new Event("pagehide"));
    expect(loadPreferences().assemblyPlan?.items.map((item) => item.id)).toEqual([a.id]);
  });

  it("clears the stored plan on Start over", async () => {
    const { seed, open, loadPreferences } = await setup();
    const a = await seed("a");
    const { assembly } = open();
    assembly.tray.add([a.id]);
    assembly.setLayout("vertical");
    assembly.flush();
    expect(loadPreferences().assemblyPlan).not.toBeNull();
    assembly.reset();
    await nextTick();
    expect(loadPreferences().assemblyPlan).toBeNull();
    expect(assembly.tray.entries.value).toEqual([]);
    expect(assembly.layout.layout).toBe("grid");
  });
});

describe("the refusal paths — before any work", () => {
  it("refuses an output over the measured canvas ceiling, naming it and the dpi that fits", async () => {
    const { seed, open, steps, store } = await setup({ canvasCeiling: vi.fn(() => 4_000_000) });
    const a = await seed("a");
    const { assembly } = open();
    assembly.tray.add([a.id]);
    assembly.sheet.setOrientation("landscape");
    assembly.sheet.setDpi(300);
    expect(assembly.out.refusal.value).toBe(
      "At 300 dpi the image would be 3508 × 2480 px (8.7 megapixels), more than this browser's canvas limit of 4.0 megapixels. At 150 dpi it would fit.",
    );
    expect(assembly.out.fixDpi.value).toBe(150);
    const before = store.listDocuments().length;
    expect(await assembly.out.save()).toBeNull();
    expect(await assembly.out.download()).toBe(false);
    expect(steps.decode).not.toHaveBeenCalled();
    expect(steps.createCanvas).not.toHaveBeenCalled();
    expect(steps.download).not.toHaveBeenCalled();
    expect(store.listDocuments()).toHaveLength(before);
    expect(assembly.sheet.setDpi(150)).toBe(true);
    expect(assembly.out.refusal.value).toBe("");
  });

  it("refuses either side over 16384 px", async () => {
    const { seed, open } = await setup();
    const a = await seed("a");
    const { assembly } = open();
    assembly.tray.add([a.id]);
    assembly.sheet.setPaper("Custom");
    assembly.sheet.setCustomSize("width", "200");
    assembly.sheet.setCustomSize("height", "20");
    assembly.sheet.setDpi(300);
    expect(assembly.out.refusal.value).toContain("either side can be at most 16384 px. At 150 dpi it would fit.");
    expect(assembly.out.canProduce.value).toBe(false);
  });

  it("refuses an image whose size could not be read, naming it", async () => {
    const { seed, open } = await setup();
    const a = await seed("mystery");
    const { assembly } = open();
    assembly.tray.add([a.id]);
    assembly.tray.replace([{ ...assembly.tray.entries.value[0], widthPx: 0, heightPx: 0 }]);
    expect(assembly.out.refusal.value).toBe("The size of “mystery” could not be read; remove it from the tray.");
    expect(assembly.out.canProduce.value).toBe(false);
  });

  it("offers no dpi that cannot help a sheet fitted to its content", async () => {
    const { seed, open } = await setup();
    const a = await seed("wide");
    const { assembly } = open();
    assembly.tray.add([a.id]);
    assembly.tray.replace([{ ...assembly.tray.entries.value[0], widthPx: 20000, heightPx: 3000 }]);
    assembly.sheet.setMarginPreset("Narrow");
    expect(assembly.sheet.setPaper("Fit to content")).toBe(true);
    expect(assembly.out.refusal.value).toContain("It would not fit even at 72 dpi");
    expect(assembly.out.fixDpi.value).toBeNull();
  });

  it("refuses an empty tray", async () => {
    const { open } = await setup();
    const { assembly } = open();
    expect(assembly.placed.value).toEqual([]);
    expect(assembly.out.refusal.value).toBe("Add at least one image to the tray.");
  });

  it("allows PNG's transparent sheet, and fills it with the chosen colour for JPEG, saying so first", async () => {
    const { seed, open } = await setup();
    const a = await seed("a");
    const { assembly } = open();
    assembly.tray.add([a.id]);
    expect(assembly.sheet.setup.value.background).toBe("transparent");
    expect(assembly.background.value).toBe("transparent");
    expect(assembly.out.backgroundNotice.value).toBe("");
    assembly.setFormat("jpeg");
    expect(assembly.out.backgroundNotice.value).toContain("JPEG has no transparency");
    expect(assembly.background.value).toBe("#ffffff");
    assembly.setMatte("#336699");
    expect(assembly.background.value).toBe("#336699");
    expect(assembly.setQuality("0")).toBe(false);
    expect(assembly.setQuality("75")).toBe(true);
    expect(assembly.output.quality).toBe(75);
  });
});

describe("producing the output", () => {
  it("saves a NEW PNG document whose summary is read from the produced bytes, numbered on a second save", async () => {
    const { seed, open, store, steps, canvases, draws } = await setup();
    const a = deepFreeze(await seed("a"));
    const b = deepFreeze(await seed("b", "jpeg"));
    const recordsBefore = [a.id, b.id].map((id) => ({ ...store.getStoredRecord(id) }));
    const { assembly } = open();
    assembly.tray.add([a.id, b.id]);
    assembly.setColumns("2");
    const progress: string[] = [];
    steps.between.mockImplementation(async () => {
      progress.push(assembly.out.progress.value);
    });
    const saved = await assembly.out.save();
    expect(saved).toMatchObject({ kind: "media", format: "png", contentType: "image/png", title: "Combined image" });
    // The encoder's bytes (the 64 × 40 fixture here) decide the summary.
    expect(saved?.image).toEqual({ widthPx: 64, heightPx: 40 });
    expect(store.getMedia(saved!.id)).toBe(saved);
    // At the sheet's pixel size: A4 portrait at 150 dpi.
    expect(steps.createCanvas).toHaveBeenCalledWith(1240, 1754);
    expect(draws.filter((draw) => draw[0] === "drawImage")).toHaveLength(2);
    expect(steps.encode).toHaveBeenCalledWith(expect.anything(), "image/png", undefined);
    // One row of two: one decode yield between the sources, then a draw.
    expect(progress).toContain("Item 0 of 2");
    expect(assembly.out.progress.value).toBe("");
    // The output canvas was released.
    expect(canvases[0]).toMatchObject({ width: 0, height: 0 });
    // The sources: untouched — same records, same models.
    expect([a.id, b.id].map((id) => ({ ...store.getStoredRecord(id) }))).toEqual(recordsBefore);
    const again = await assembly.out.save();
    expect(again?.title).toBe("Combined image (2)");
    expect(store.listDocuments().filter((entry) => entry.format === "png")).toHaveLength(3);
  });

  it("decodes each source once for the session, however often it is drawn", async () => {
    const { createSourceCache } = await import("../../lib/media/composeCanvas");
    const raw = vi.fn(async (): Promise<DecodedImage> => ({ image: {} as CanvasImageSource, widthPx: 64, heightPx: 40, release: () => {} }));
    const cache = createSourceCache({ decode: raw, rasteriseSvg: vi.fn() });
    const { seed, open } = await setup({ decode: cache.decodeSource });
    const a = await seed("a");
    const { assembly } = open();
    assembly.tray.add([a.id]);
    assembly.tray.duplicate(assembly.tray.entries.value[0].key);
    await assembly.out.save();
    await assembly.out.download();
    expect(raw).toHaveBeenCalledTimes(1);
  });

  it("encodes JPEG at the chosen quality and downloads through triggerDownload's step, storing nothing", async () => {
    const { seed, open, store, steps } = await setup();
    const a = deepFreeze(await seed("a"));
    const { assembly } = open();
    assembly.tray.add([a.id]);
    assembly.setFormat("jpeg");
    assembly.setQuality("80");
    assembly.setTitle("Before and after");
    const before = store.listDocuments().length;
    expect(await assembly.out.download()).toBe(true);
    expect(steps.encode).toHaveBeenCalledWith(expect.anything(), "image/jpeg", 0.8);
    expect(steps.download).toHaveBeenCalledWith(expect.any(Uint8Array), "Before and after.jpeg", "image/jpeg");
    expect(store.listDocuments()).toHaveLength(before);
  });

  it("stops at the next boundary when cancelled, leaving no file and no library record", async () => {
    const { seed, open, store, steps, canvases } = await setup();
    const ids = [(await seed("a")).id, (await seed("b")).id, (await seed("c")).id];
    const { assembly } = open();
    assembly.tray.add(ids);
    assembly.setLayout("vertical");
    let boundaries = 0;
    steps.between.mockImplementation(async () => {
      boundaries += 1;
      // Cancel during the drawing (after the sources are decoded).
      if (boundaries === 3) {
        expect(assembly.out.busy.value).toBe(true);
        assembly.out.cancel();
      }
    });
    const before = store.listDocuments().length;
    expect(await assembly.out.save()).toBeNull();
    expect(steps.encode).not.toHaveBeenCalled();
    expect(store.listDocuments()).toHaveLength(before);
    expect(assembly.out.failure.value).toBe("");
    expect(assembly.out.busy.value).toBe(false);
    expect(canvases[0]).toMatchObject({ width: 0, height: 0 });
    // A cancelled download hands nothing to the browser.
    boundaries = 0;
    expect(await assembly.out.download()).toBe(false);
    expect(steps.download).not.toHaveBeenCalled();
  });

  it("stops when the workspace is left mid-way, writing nothing", async () => {
    const { seed, open, store, steps } = await setup();
    const ids = [(await seed("a")).id, (await seed("b")).id];
    const { assembly, stop } = open();
    assembly.tray.add(ids);
    assembly.setLayout("vertical");
    steps.between.mockImplementation(async () => {
      stop();
    });
    const before = store.listDocuments().length;
    expect(await assembly.out.save()).toBeNull();
    expect(steps.encode).not.toHaveBeenCalled();
    expect(store.listDocuments()).toHaveLength(before);
  });

  it("reports an unreadable source by name and writes nothing", async () => {
    const { seed, open, store } = await setup({ decode: vi.fn().mockRejectedValue(new Error("bad")) });
    const a = await seed("broken");
    const { assembly } = open();
    assembly.tray.add([a.id]);
    const before = store.listDocuments().length;
    expect(await assembly.out.save()).toBeNull();
    expect(assembly.out.failure.value).toBe("Could not save the image: “broken” could not be read.");
    expect(store.listDocuments()).toHaveLength(before);
  });
});

describe("Fit to content", () => {
  it("lays the content out in the column count the sheet was sized for", async () => {
    const { seed, open } = await setup();
    const ids = [(await seed("s1")).id, (await seed("s2")).id, (await seed("s3")).id];
    const { assembly } = open();
    assembly.tray.add(ids);
    // Three 1920 × 1080 screenshots, auto columns: sized 2 × 2 (3840 × 2160).
    assembly.tray.replace(assembly.tray.entries.value.map((entry) => ({ ...entry, widthPx: 1920, heightPx: 1080 })));
    expect(assembly.sheet.setPaper("Fit to content")).toBe(true);
    expect(assembly.sheet.itemPx.value).toEqual({ widthPx: 3840, heightPx: 2160 });
    expect(assembly.plan.value.columns).toBe(2);
    const [first, second, third] = assembly.placed.value;
    expect(second.yCm).toBe(first.yCm);
    expect(third.yCm).toBeGreaterThan(first.yCm);
    // Each item fills its cell along one side and nearly all of it — the
    // 0.2 cm gaps are all that is not image.
    const content = assembly.geometry.value!.content;
    expect(first.cell.widthCm).toBeCloseTo((content.widthCm - 0.2) / 2, 3);
    expect(Math.max(first.widthCm / first.cell.widthCm, first.heightCm / first.cell.heightCm)).toBeCloseTo(1, 3);
    expect((first.widthCm * first.heightCm) / (first.cell.widthCm * first.cell.heightCm)).toBeGreaterThan(0.99);
    // On a paper size, "auto" decides on that paper's own box.
    assembly.sheet.setPaper("A4");
    expect(assembly.plan.value.columns).toBe("auto");
  });
});

describe("overlay editing", () => {
  it("edits the selected item's transform from the inspector, and nudges it by 1 mm / 1 cm", async () => {
    const { seed, open } = await setup();
    const a = await seed("a");
    const { assembly } = open();
    assembly.tray.add([a.id, a.id]);
    const [, top] = assembly.tray.entries.value;
    expect(assembly.overlayItem.value).toBeNull();
    assembly.setLayout("overlay");
    assembly.select(top.key);
    expect(assembly.overlayItem.value?.key).toBe(top.key);
    expect(assembly.setTransformField("scalePercent", "50")).toBe(true);
    expect(assembly.setTransformField("blend", "screen")).toBe(true);
    expect(assembly.setTransformField("blend", "lighten")).toBe(false);
    expect(assembly.setTransformField("xPercent", "abc")).toBe(false);
    const placed = assembly.placed.value[1];
    expect([placed.blend, placed.widthCm]).toEqual(["screen", 7.96]);
    // A4 portrait, Normal margins: a 15.92 cm wide content box.
    assembly.nudge(top.key, 0.1, 0);
    expect(assembly.tray.entries.value[1].transform.xPercent).toBe(50.628);
    assembly.nudge(top.key, 0, 1);
    expect(assembly.tray.entries.value[1].transform.yPercent).toBe(54.062);
    // Back to front: the selected (second) item is drawn last.
    expect(assembly.placed.value.map((item) => item.index)).toEqual([0, 1]);
  });
});
