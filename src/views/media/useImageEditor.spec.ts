// SPEC-112: the image editor's session, without a rendering engine (the
// browser-bound steps are fakes that record what they were asked to draw):
// undo granularity per gesture, the redaction notice's presence, flatten's
// confirmation and irreversibility, the SVG raster base, the canvas-ceiling
// refusal, the outputs' write-through and failure paths — and no edit path
// mutating a deep-frozen source.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { effectScope, shallowRef, type EffectScope } from "vue";
import type { ImageEditSteps } from "../../lib/media/renderEdits";
import type { MediaModel } from "../../lib/media/model";

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
  vi.restoreAllMocks();
});

/** Deep-freezes a value in place (a frozen model: any mutation throws). */
function deepFreeze<T>(value: T): T {
  if (value && typeof value === "object" && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const child of Object.values(value)) {
      deepFreeze(child);
    }
  }
  return value;
}

/** Browser-free steps: decode reports a size, render returns a canvas-shaped
 * record of its target, encode returns recognisable bytes. */
function fakeSteps(overrides: Partial<ImageEditSteps> = {}) {
  let rasterSize: { widthPx: number; heightPx: number } | null = null;
  const steps: ImageEditSteps = {
    decode: vi.fn(async (_bytes: Uint8Array, contentType: string) => {
      const size = contentType === "image/png" && rasterSize ? rasterSize : { widthPx: 160, heightPx: 120 };
      return { image: {} as CanvasImageSource, ...size, release: () => {} };
    }),
    preview: (decoded) => ({ image: decoded.image, naturalPx: { widthPx: decoded.widthPx, heightPx: decoded.heightPx } }),
    render: vi.fn(async (_source, _edits, target) => ({ width: target.widthPx, height: target.heightPx }) as HTMLCanvasElement),
    encode: vi.fn(async (image: HTMLCanvasElement) => new Uint8Array([0xee, image.width, image.height])),
    rasteriseSvg: vi.fn(async (_bytes: Uint8Array, size?: { widthPx: number; heightPx: number }) => {
      rasterSize = size ?? { widthPx: 300, heightPx: 150 };
      return { data: new Uint8Array([0x89, 0x50]), ...rasterSize };
    }),
    canvasCeiling: () => 4096 * 4096,
    download: vi.fn(),
    release: vi.fn(),
    ...overrides,
  };
  return steps;
}

let scope: EffectScope | null = null;
afterEach(() => {
  scope?.stop();
  scope = null;
});

async function setup(options: { format?: "jpeg" | "svg" | "gif"; svgBytes?: Uint8Array; steps?: ImageEditSteps; freeze?: boolean } = {}) {
  const store = await import("../../lib/documentStore");
  const { assembleMediaModel } = await import("../../lib/media/open");
  const fixtures = await import("../../lib/media/fixtures/build");
  const { useImageEditor, REDACTION_NOTICE } = await import("./useImageEditor");
  const { IMAGE_UNDO_LIMIT } = await import("./useEditUndo");
  const format = options.format ?? "jpeg";
  const bytes =
    format === "svg"
      ? (options.svgBytes ?? fixtures.buildSvgFixtureBytes())
      : format === "gif"
        ? fixtures.buildGifFixtureBytes()
        : fixtures.buildJpegFixtureBytes();
  const model = await assembleMediaModel("photo", format, bytes);
  store.storeMedia(model);
  if (options.freeze) {
    deepFreeze(model);
  }
  const media = shallowRef<MediaModel | null>(model);
  const steps = options.steps ?? fakeSteps();
  scope = effectScope();
  const editor = scope.run(() =>
    useImageEditor({
      media,
      steps,
      onStored: () => {
        media.value = store.getMedia(model.id) ?? null;
      },
    }),
  )!;
  return { store, model, media, steps, editor, REDACTION_NOTICE, IMAGE_UNDO_LIMIT };
}

describe("useImageEditor — entering", () => {
  it("decodes the base once, builds the preview and starts from the saved stack", async () => {
    const { editor, steps } = await setup();
    await editor.enter();
    expect(editor.active.value).toBe(true);
    expect(editor.phase.value).toBe("ready");
    expect(editor.naturalPx.value).toEqual({ widthPx: 160, heightPx: 120 });
    expect(editor.history.value).toEqual([]);
    expect(editor.geometry.value).toMatchObject({ widthPx: 160, heightPx: 120 });
    expect(steps.decode).toHaveBeenCalledTimes(1);
  });

  it("refuses up front an image over the canvas ceiling, naming the limit", async () => {
    const steps = fakeSteps({ canvasCeiling: () => 10_000 });
    const { editor } = await setup({ steps });
    await editor.enter();
    expect(editor.phase.value).toBe("refused");
    expect(editor.refusal.value).toContain("160 × 120 px");
    expect(editor.refusal.value).toContain("0.0 megapixels");
    expect(steps.decode).not.toHaveBeenCalled();
  });

  it("says an animated GIF is edited as its first frame before it starts", async () => {
    const { editor } = await setup({ format: "gif" });
    await editor.enter();
    expect(editor.isAnimatedGif.value).toBe(true);
  });
});

describe("useImageEditor — undo granularity", () => {
  it("records one entry per completed gesture: a slider drag, a crop drag, a drawn redaction", async () => {
    const { editor } = await setup();
    await editor.enter();

    for (const value of [5, 10, 20, 30]) {
      editor.commands.adjustInput("brightness", value);
    }
    editor.commands.release();
    expect(editor.history.value).toEqual([{ kind: "adjust", brightness: 30 }]);
    expect(editor.undo.undoLabel.value).toBe("Adjust brightness");

    editor.crop.startDrag("se");
    for (const dx of [-5, -10, -20, -40]) {
      editor.crop.dragBy(dx, 0);
    }
    editor.crop.endDrag();
    expect(editor.crop.announcement.value).toBe("0, 0, 120 × 120");
    expect(editor.undo.undoLabel.value).toBe("Crop");

    editor.objects.startDraw("redact", 10, 10);
    editor.objects.drawTo(20, 15);
    editor.objects.drawTo(40, 30);
    editor.objects.endDraw();
    expect(editor.undo.undoLabel.value).toBe("Redact");

    // Three gestures, three steps — however many moves each made.
    editor.undoStep();
    expect(editor.stack.value.some((edit) => edit.kind === "redact")).toBe(false);
    editor.undoStep();
    expect(editor.crop.announcement.value).toBe("0, 0, 160 × 120");
    editor.undoStep();
    expect(editor.history.value).toEqual([]);
    expect(editor.undo.canUndo.value).toBe(false);
    // And redo walks them back.
    editor.redoStep();
    editor.redoStep();
    editor.redoStep();
    expect(editor.stack.value.map((edit) => edit.kind)).toEqual(["crop", "adjust", "redact"]);
  });

  it("makes each key press on a crop handle one step, 1 px or 10 px with Shift", async () => {
    const { editor } = await setup();
    await editor.enter();
    editor.crop.setAspect("16:9");
    expect(editor.crop.announcement.value).toBe("0, 15, 160 × 90");
    // The preset holds the aspect: the bottom edge up one row is 89 rows,
    // and 89 · 16/9 = 158.2 → 158 columns.
    editor.crop.handleKey("s", new KeyboardEvent("keydown", { key: "ArrowUp" }));
    expect(editor.crop.announcement.value).toBe("0, 15, 158 × 89");
    editor.undoStep();
    expect(editor.crop.announcement.value).toBe("0, 15, 160 × 90");
    // Free: the edge alone moves.
    editor.crop.setAspect("free");
    editor.crop.handleKey("s", new KeyboardEvent("keydown", { key: "ArrowUp" }));
    expect(editor.crop.announcement.value).toBe("0, 15, 160 × 89");
    editor.crop.handleKey("s", new KeyboardEvent("keydown", { key: "ArrowUp", shiftKey: true }));
    expect(editor.crop.announcement.value).toBe("0, 15, 160 × 79");
    editor.crop.handleKey("move", new KeyboardEvent("keydown", { key: "ArrowRight", shiftKey: true }));
    expect(editor.crop.announcement.value).toBe("0, 15, 160 × 79");
    editor.crop.handleKey("move", new KeyboardEvent("keydown", { key: "ArrowDown" }));
    expect(editor.crop.announcement.value).toBe("0, 16, 160 × 79");
    editor.undoStep();
    editor.undoStep();
    expect(editor.crop.announcement.value).toBe("0, 15, 160 × 89");
    expect(editor.undo.undoLabel.value).toBe("Crop");
  });

  it("records BEFORE each mutation and commits after it (the SPEC-080 order)", async () => {
    const { editor } = await setup();
    await editor.enter();
    const order: string[] = [];
    const record = editor.undo.record;
    const commit = editor.undo.commit;
    vi.spyOn(editor.undo, "record").mockImplementation((label) => {
      order.push(`record ${label} at ${editor.history.value.length}`);
      record(label);
    });
    vi.spyOn(editor.undo, "commit").mockImplementation(() => {
      order.push(`commit at ${editor.history.value.length}`);
      commit();
    });
    editor.commands.rotate(1);
    editor.objects.add("annotate");
    expect(order).toEqual(["record Rotate right at 0", "commit at 1", "record Add arrow at 1", "commit at 2"]);
  });

  it("drops a click that draws nothing, without an entry", async () => {
    const { editor } = await setup();
    await editor.enter();
    editor.objects.startDraw("redact", 50, 50);
    editor.objects.endDraw();
    expect(editor.history.value).toEqual([]);
    expect(editor.undo.canUndo.value).toBe(false);
  });

  it("caps the stack at IMAGE_UNDO_LIMIT, dropping the oldest", async () => {
    const { editor, IMAGE_UNDO_LIMIT } = await setup();
    await editor.enter();
    for (let i = 0; i < IMAGE_UNDO_LIMIT + 20; i += 1) {
      editor.commands.flip("horizontal");
    }
    let undone = 0;
    while (editor.undo.canUndo.value) {
      editor.undoStep();
      undone += 1;
    }
    expect(IMAGE_UNDO_LIMIT).toBe(100);
    expect(undone).toBe(100);
  });

  it("wires SPEC-015's shortcuts and Cmd/Ctrl+[ / ] while editing", async () => {
    const { editor } = await setup();
    await editor.enter();
    const key = (init: KeyboardEventInit) => editor.handleKeydown(new KeyboardEvent("keydown", init));
    expect(key({ key: "]", ctrlKey: true })).toBe(true);
    expect(editor.history.value).toEqual([{ kind: "rotate", quarterTurns: 1 }]);
    expect(key({ key: "[", ctrlKey: true })).toBe(true);
    expect(editor.history.value).toEqual([]);
    expect(key({ key: "z", ctrlKey: true })).toBe(true);
    expect(editor.history.value).toEqual([{ kind: "rotate", quarterTurns: 1 }]);
    expect(key({ key: "z", ctrlKey: true, shiftKey: true })).toBe(true);
    expect(editor.history.value).toEqual([]);
    // Escape leaves — after asking, since there are unsaved edits.
    key({ key: "]", ctrlKey: true });
    expect(key({ key: "Escape" })).toBe(true);
    expect(editor.discardPending.value).toBe(true);
    editor.exit();
    expect(editor.active.value).toBe(false);
  });
});

describe("useImageEditor — objects from the keyboard", () => {
  it("adds, nudges (1 px, Shift 10 px), resizes (Alt) and deletes, one step per key", async () => {
    const { editor } = await setup();
    await editor.enter();
    editor.objects.add("redact");
    const index = editor.history.value.length - 1;
    // Placed at the middle quarter of the 160 × 120 output.
    expect(editor.history.value[index]).toMatchObject({ x: 60, y: 45, width: 40, height: 30, color: "#000000" });
    const key = (init: KeyboardEventInit) => editor.objects.handleKey(index, new KeyboardEvent("keydown", init));
    key({ key: "ArrowRight" });
    key({ key: "ArrowDown", shiftKey: true });
    key({ key: "ArrowLeft", altKey: true });
    expect(editor.history.value[index]).toMatchObject({ x: 61, y: 55, width: 39, height: 30 });
    expect(editor.undo.undoLabel.value).toBe("Resize redaction");
    editor.undoStep();
    expect(editor.undo.undoLabel.value).toBe("Move redaction");
    expect(key({ key: "Delete" })).toBe(true);
    expect(editor.history.value).toEqual([]);
    expect(editor.undo.undoLabel.value).toBe("Delete redaction");
  });
});

describe("useImageEditor — the redaction notice", () => {
  it("is present exactly while the stack holds a redaction", async () => {
    const { editor, REDACTION_NOTICE } = await setup();
    expect(REDACTION_NOTICE).toBe(
      "The original is still stored. Choose Flatten and discard original to remove the hidden pixels.",
    );
    await editor.enter();
    expect(editor.hasRedaction.value).toBe(false);
    editor.objects.add("annotate");
    expect(editor.hasRedaction.value).toBe(false);
    editor.objects.add("redact");
    expect(editor.hasRedaction.value).toBe(true);
    editor.undoStep();
    expect(editor.hasRedaction.value).toBe(false);
    editor.redoStep();
    expect(editor.hasRedaction.value).toBe(true);
    const index = editor.history.value.findIndex((edit) => edit.kind === "redact");
    editor.objects.remove(index);
    expect(editor.hasRedaction.value).toBe(false);
  });
});

describe("useImageEditor — flatten", () => {
  it("asks first, then replaces the source's bytes, clears the stack, and no undo brings the original back", async () => {
    const { editor, store, model, steps } = await setup();
    const originalData = model.media[0].data;
    await editor.enter();
    // Nothing to flatten: the dialog does not open.
    editor.actions.requestFlatten();
    expect(editor.actions.flattenPending.value).toBe(false);

    editor.objects.add("redact");
    editor.actions.requestFlatten();
    expect(editor.actions.flattenPending.value).toBe(true);
    // Asking changed nothing.
    expect(store.getMedia(model.id)).toBe(model);
    editor.actions.cancelFlatten();
    expect(editor.actions.flattenPending.value).toBe(false);
    expect(store.getMedia(model.id)).toBe(model);

    editor.actions.requestFlatten();
    expect(await editor.actions.confirmFlatten()).toBe(true);
    const flattened = store.getMedia(model.id)!;
    expect(flattened).not.toBe(model);
    expect(flattened.sourceMediaId).toBe(model.sourceMediaId);
    expect(flattened.media).toHaveLength(1);
    expect(flattened.media[0].mediaId).toBe(model.sourceMediaId);
    expect(flattened.media[0].data).not.toBe(originalData);
    expect(flattened.edits).toBeUndefined();
    expect(flattened.renderMediaId).toBeUndefined();
    expect(flattened.contentType).toBe("image/jpeg");
    // Rendered at full resolution, chunked.
    expect(steps.render).toHaveBeenLastCalledWith(expect.anything(), [expect.objectContaining({ kind: "redact" })], { widthPx: 160, heightPx: 120 }, expect.objectContaining({ chunked: true }));

    // The session restarted on the new source: nothing to undo into.
    expect(editor.history.value).toEqual([]);
    expect(editor.undo.canUndo.value).toBe(false);
    editor.undoStep();
    expect(editor.history.value).toEqual([]);
    expect(store.getMedia(model.id)).toBe(flattened);
    expect(JSON.stringify(store.getMedia(model.id))).not.toContain(originalData);
  });
});

describe("useImageEditor — the outputs", () => {
  it("Save stores the stack and a render, leaving the source part untouched", async () => {
    const { editor, store, model } = await setup();
    await editor.enter();
    editor.crop.setAspect("16:9");
    editor.commands.rotate(1);
    editor.commands.setAdjust("brightness", 30);
    expect(await editor.actions.save()).toBe(true);
    const saved = store.getMedia(model.id)!;
    expect(saved.sourceMediaId).toBe(model.sourceMediaId);
    expect(saved.media[0]).toBe(model.media[0]);
    expect(saved.edits).toEqual([
      { kind: "crop", x: 0, y: 15, width: 160, height: 90 },
      { kind: "rotate", quarterTurns: 1 },
      { kind: "adjust", brightness: 30 },
    ]);
    const render = saved.media.find((part) => part.mediaId === saved.renderMediaId);
    expect(render?.contentType).toBe("image/jpeg");
    expect(editor.dirty.value).toBe(false);
    expect(editor.actions.notice.value).toBe("Saved.");
  });

  it("Save as a copy creates a new document with the same source and stack", async () => {
    const { editor, store, model } = await setup();
    await editor.enter();
    editor.commands.rotate(-1);
    expect(await editor.actions.saveCopy()).toBe(true);
    const copy = store.getMedia(editor.actions.copyId.value)!;
    expect(copy.id).not.toBe(model.id);
    expect(copy.title).toBe("photo (copy)");
    expect(copy.media.find((part) => part.mediaId === copy.sourceMediaId)?.data).toBe(model.media[0].data);
    expect(copy.edits).toEqual([{ kind: "rotate", quarterTurns: 3 }]);
    // The original is as it was.
    expect(store.getMedia(model.id)).toBe(model);
  });

  it("Export downloads a flattened PNG or JPEG", async () => {
    const { editor, steps } = await setup();
    await editor.enter();
    editor.commands.rotate(1);
    editor.actions.openExport();
    editor.actions.exportFormat.value = "jpeg";
    expect(editor.actions.exportBackgroundNotice()).toContain("JPEG images have no transparency");
    expect(await editor.actions.exportImage()).toBe(true);
    expect(steps.encode).toHaveBeenLastCalledWith(expect.anything(), "image/jpeg", { quality: 0.9, background: "#ffffff" });
    expect(steps.download).toHaveBeenCalledWith(expect.any(Uint8Array), "photo.jpeg", "image/jpeg");
    expect(editor.actions.exportOpen.value).toBe(false);
  });

  it("reports a failure and mutates nothing", async () => {
    const steps = fakeSteps({ encode: vi.fn().mockRejectedValue(new Error("encoder broke")) });
    const { editor, store, model } = await setup({ steps });
    await editor.enter();
    editor.commands.rotate(1);
    expect(await editor.actions.save()).toBe(false);
    expect(await editor.actions.saveCopy()).toBe(false);
    editor.objects.add("redact");
    editor.actions.requestFlatten();
    expect(await editor.actions.confirmFlatten()).toBe(false);
    expect(editor.actions.failure.value).toBe("Could not flatten the image: encoder broke");
    expect(store.getMedia(model.id)).toBe(model);
    expect(store.listDocuments()).toHaveLength(1);
    expect(editor.actions.busy.value).toBe(false);
  });

  it("Revert to original names what it discards and clears the stack, undoably", async () => {
    const { editor } = await setup();
    await editor.enter();
    editor.crop.setAspect("1:1");
    editor.commands.rotate(1);
    editor.requestRevert();
    expect(editor.revertPending.value).toBe(true);
    expect(editor.revertSummary.value).toBe("Crop, Rotate right");
    editor.confirmRevert();
    expect(editor.history.value).toEqual([]);
    expect(editor.undo.undoLabel.value).toBe("Revert to original");
    editor.undoStep();
    expect(editor.history.value).toHaveLength(2);
  });

  it("removes one history row", async () => {
    const { editor } = await setup();
    await editor.enter();
    editor.crop.setAspect("1:1");
    editor.commands.rotate(1);
    editor.removeEntry(0);
    expect(editor.history.value).toEqual([{ kind: "rotate", quarterTurns: 1 }]);
    expect(editor.undo.undoLabel.value).toBe("Remove crop");
  });
});

describe("useImageEditor — SVG", () => {
  it("rasterises once, on the first edit, at the natural size; the raster becomes the base and the SVG stays the source", async () => {
    const { editor, store, model, steps } = await setup({ format: "svg" });
    expect(model.image).toEqual({ widthPx: 96, heightPx: 72 });
    await editor.enter();
    expect(editor.isSvg.value).toBe(true);
    expect(steps.rasteriseSvg).toHaveBeenCalledTimes(1);
    expect(steps.rasteriseSvg).toHaveBeenCalledWith(expect.any(Uint8Array), { widthPx: 96, heightPx: 72 });
    expect(editor.naturalPx.value).toEqual({ widthPx: 96, heightPx: 72 });
    editor.commands.rotate(1);
    expect(await editor.actions.save()).toBe(true);
    const saved = store.getMedia(model.id)!;
    expect(saved.sourceMediaId).toBe(model.sourceMediaId);
    expect(saved.media.find((part) => part.mediaId === saved.sourceMediaId)?.contentType).toBe("image/svg+xml");
    expect(saved.media.find((part) => part.mediaId === saved.baseMediaId)?.contentType).toBe("image/png");
    expect(saved.media.find((part) => part.mediaId === saved.renderMediaId)?.contentType).toBe("image/png");
    // Re-entering edits the stored raster: no second rasterisation.
    editor.exit();
    await editor.enter();
    expect(steps.rasteriseSvg).toHaveBeenCalledTimes(1);
    expect(editor.history.value).toEqual([{ kind: "rotate", quarterTurns: 1 }]);
  });

  it("asks for a size when the SVG declares none", async () => {
    const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"><rect width="5" height="5"/></svg>');
    const { editor, steps } = await setup({ format: "svg", svgBytes: svg });
    await editor.enter();
    expect(editor.phase.value).toBe("size");
    editor.svgSize.value = { widthPx: 400, heightPx: 200 };
    await editor.startSvg();
    expect(steps.rasteriseSvg).toHaveBeenCalledWith(expect.any(Uint8Array), { widthPx: 400, heightPx: 200 });
    expect(editor.phase.value).toBe("ready");
    expect(editor.naturalPx.value).toEqual({ widthPx: 400, heightPx: 200 });
  });
});

describe("useImageEditor — no edit path mutates the source", () => {
  it("works over a deep-frozen model: edit, save, copy, export, revert and flatten", async () => {
    const { editor, store, model } = await setup({ freeze: true });
    const snapshot = JSON.stringify(model);
    await editor.enter();
    editor.crop.setAspect("4:3");
    editor.commands.rotate(1);
    editor.commands.flip("vertical");
    editor.commands.setResizePercent(50);
    editor.commands.setAdjust("contrast", 20);
    editor.objects.add("redact");
    editor.objects.add("annotate");
    expect(await editor.actions.save()).toBe(true);
    const saved = store.getMedia(model.id)!;
    deepFreeze(saved);
    expect(await editor.actions.saveCopy()).toBe(true);
    expect(await editor.actions.exportImage()).toBe(true);
    editor.requestRevert();
    editor.confirmRevert();
    editor.objects.add("redact");
    editor.actions.requestFlatten();
    expect(await editor.actions.confirmFlatten()).toBe(true);
    expect(editor.actions.failure.value).toBe("");
    // The frozen original is exactly what it was.
    expect(JSON.stringify(model)).toBe(snapshot);
  });
});
