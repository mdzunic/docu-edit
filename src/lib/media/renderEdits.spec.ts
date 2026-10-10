// SPEC-112: the shared renderer, against a recording 2-D context that keeps
// a real pixel buffer — so the arithmetic handed to the canvas is asserted
// exactly (the transform, the target size, whole-pixel redaction fills) and
// the rendered PIXELS are asserted by sampling, never the encoded bytes.
// Also: the preview buffer's cap, the single-decode cache, and the encoder's
// JPEG background.
import { describe, expect, it, vi } from "vitest";
import type { ImageEdit } from "../core/imageEdits";
import {
  createDecodeCache,
  drawAnnotation,
  EDIT_PREVIEW_MAX_EDGE_PX,
  encodeRender,
  previewBufferSize,
  previewSource,
  previewTargetSize,
  renderEditedImage,
  type DecodedImage,
} from "./renderEdits";

type Call = [string, ...unknown[]];

/** A canvas whose 2-D context records every call and holds real RGBA
 * pixels: drawImage paints the whole target `base`, fillRect paints its
 * rectangle in the current (hex) fillStyle. */
function recordingCanvas(width: number, height: number, base = [100, 150, 200, 255]) {
  const calls: Call[] = [];
  const pixels = new Uint8ClampedArray(width * height * 4);
  const hex = (value: string) => [1, 3, 5].map((i) => Number.parseInt(value.slice(i, i + 2), 16));
  const ctx: Record<string, unknown> = {
    fillStyle: "#000000",
    save: () => calls.push(["save"]),
    restore: () => calls.push(["restore"]),
    setTransform: (...args: unknown[]) => calls.push(["setTransform", ...args]),
    drawImage: (...args: unknown[]) => {
      calls.push(["drawImage", ...args]);
      for (let i = 0; i < pixels.length; i += 4) pixels.set(base, i);
    },
    getImageData: (x: number, y: number, w: number, h: number) => {
      calls.push(["getImageData", x, y, w, h]);
      return { data: pixels.slice((y * width + x) * 4, ((y + h) * width + x) * 4), width: w, height: h };
    },
    putImageData: (image: { data: Uint8ClampedArray }, x: number, y: number) => {
      calls.push(["putImageData", x, y]);
      pixels.set(image.data, (y * width + x) * 4);
    },
    fillRect: (x: number, y: number, w: number, h: number) => {
      calls.push(["fillRect", ctx.fillStyle, x, y, w, h]);
      const [r, g, b] = hex(String(ctx.fillStyle));
      for (let row = y; row < y + h; row += 1) {
        for (let col = x; col < x + w; col += 1) pixels.set([r, g, b, 255], (row * width + col) * 4);
      }
    },
  };
  for (const name of ["beginPath", "moveTo", "lineTo", "closePath", "stroke", "fill", "strokeRect", "ellipse", "fillText"]) {
    ctx[name] = (...args: unknown[]) => calls.push([name, ...args]);
  }
  const canvas = {
    width,
    height,
    getContext: () => ctx,
    toBlob: (done: (blob: Blob) => void, type: string, quality?: number) => {
      calls.push(["toBlob", type, quality]);
      // jsdom's Blob has no arrayBuffer(); every target browser's does.
      done({ type, arrayBuffer: async () => new Uint8Array([1, 2, 3]).buffer } as unknown as Blob);
    },
  };
  const sample = (x: number, y: number) => Array.from(pixels.slice((y * width + x) * 4, (y * width + x) * 4 + 4));
  return { canvas: canvas as unknown as HTMLCanvasElement, calls, sample };
}

/** A factory remembering every canvas it made. */
function factory(base?: number[]) {
  const made: ReturnType<typeof recordingCanvas>[] = [];
  return {
    made,
    create: (width: number, height: number) => {
      const canvas = recordingCanvas(width, height, base);
      made.push(canvas);
      return canvas.canvas;
    },
  };
}

const IMAGE = { tag: "decoded image" } as unknown as CanvasImageSource;
const SOURCE = { image: IMAGE, naturalPx: { widthPx: 160, heightPx: 120 } };

describe("renderEditedImage — the arithmetic handed to the canvas", () => {
  it("draws the identity stack as one drawImage of the natural image", async () => {
    const f = factory();
    const out = await renderEditedImage(SOURCE, [], { widthPx: 160, heightPx: 120 }, { createCanvas: f.create });
    expect([out.width, out.height]).toEqual([160, 120]);
    const calls = f.made[0].calls;
    expect(calls).toContainEqual(["setTransform", 1, 0, 0, 1, 0, 0]);
    expect(calls.filter((call) => call[0] === "drawImage")).toEqual([["drawImage", IMAGE, 0, 0, 160, 120]]);
    expect(calls.some((call) => call[0] === "getImageData")).toBe(false);
  });

  it("maps crop + rotate through one transform, scaled to the preview's size", async () => {
    const f = factory();
    const stack: ImageEdit[] = [
      { kind: "rotate", quarterTurns: 1 },
      { kind: "crop", x: 0, y: 15, width: 160, height: 90 },
    ];
    // The output is 90 × 160; the preview renders it at half size.
    await renderEditedImage(SOURCE, stack, { widthPx: 45, heightPx: 80 }, { createCanvas: f.create });
    const transform = f.made[0].calls.find((call) => call[0] === "setTransform" && call[1] !== 1);
    // turn-right ∘ translate(0, −15) = (0, 1, −1, 0, 105, 0); halved.
    expect(transform).toEqual(["setTransform", 0, 0.5, -0.5, 0, 52.5, 0]);
    expect([f.made[0].canvas.width, f.made[0].canvas.height]).toEqual([45, 80]);
  });

  it("samples nearest-neighbour only when the resize asks for it", async () => {
    const f = factory();
    const ctx = () => f.made.at(-1)!.canvas.getContext("2d") as unknown as Record<string, unknown>;
    await renderEditedImage(SOURCE, [{ kind: "resize", widthPx: 32, heightPx: 24, resample: "nearest" }], { widthPx: 32, heightPx: 24 }, { createCanvas: f.create });
    expect(ctx().imageSmoothingEnabled).toBe(false);
    await renderEditedImage(SOURCE, [{ kind: "resize", widthPx: 32, heightPx: 24, resample: "bilinear" }], { widthPx: 32, heightPx: 24 }, { createCanvas: f.create });
    expect(ctx().imageSmoothingEnabled).toBe(true);
  });
});

describe("renderEditedImage — pixels, by sampling", () => {
  it("adjusts through core/imageAdjust's maths, alpha untouched", async () => {
    const f = factory([100, 150, 200, 255]);
    await renderEditedImage(SOURCE, [{ kind: "adjust", brightness: 50 }], { widthPx: 16, heightPx: 12 }, { createCanvas: f.create });
    // 100 + 127.5 = 227.5 → 228; the rest clamp at 255.
    expect(f.made[0].sample(5, 5)).toEqual([228, 255, 255, 255]);
  });

  it("covers a redaction with whole, opaque pixels snapped outward", async () => {
    const f = factory([100, 150, 200, 255]);
    const stack: ImageEdit[] = [
      { kind: "adjust", invert: true },
      { kind: "redact", x: 10, y: 10, width: 15, height: 15, color: "#000000" },
    ];
    // At 0.3 scale the cover spans 3..7.5 → snapped to 3..8 on both axes.
    await renderEditedImage(SOURCE, stack, { widthPx: 48, heightPx: 36 }, { createCanvas: f.create });
    const canvas = f.made[0];
    expect(canvas.calls.find((call) => call[0] === "fillRect")).toEqual(["fillRect", "#000000", 3, 3, 5, 5]);
    expect(canvas.sample(3, 3)).toEqual([0, 0, 0, 255]);
    expect(canvas.sample(7, 7)).toEqual([0, 0, 0, 255]);
    // Outside it: the inverted picture (255 − 100, 255 − 150, 255 − 200).
    expect(canvas.sample(8, 8)).toEqual([155, 105, 55, 255]);
    expect(canvas.sample(2, 2)).toEqual([155, 105, 55, 255]);
  });

  it("runs the stages in the stack's order, yielding between them at full resolution", async () => {
    const f = factory();
    const stages: string[] = [];
    const stack: ImageEdit[] = [
      { kind: "annotate", shape: { shape: "line", x1: 0, y1: 0, x2: 10, y2: 10 }, color: "#ff0000", strokeWidth: 2, opacity: 1 },
      { kind: "redact", x: 0, y: 0, width: 4, height: 4, color: "#000000" },
      { kind: "adjust", contrast: 10 },
      { kind: "flip", axis: "horizontal" },
    ];
    await renderEditedImage(SOURCE, stack, { widthPx: 2048, heightPx: 1200 }, {
      createCanvas: f.create,
      chunked: true,
      onStage: (stage) => stages.push(stage),
    });
    expect(stages).toEqual(["geometry", "adjust", "redact", "annotate"]);
    // 2048 columns → 512-row bands → three bands over 1200 rows.
    expect(f.made[0].calls.filter((call) => call[0] === "getImageData").map((call) => call.slice(1))).toEqual([
      [0, 0, 2048, 512],
      [0, 512, 2048, 512],
      [0, 1024, 2048, 176],
    ]);
  });

  it("never touches the source or the stack it was given", async () => {
    const f = factory();
    const stack = Object.freeze([Object.freeze({ kind: "redact", x: 1, y: 1, width: 2, height: 2, color: "#00ff00" })]) as readonly ImageEdit[];
    const source = Object.freeze({ image: IMAGE, naturalPx: Object.freeze({ widthPx: 160, heightPx: 120 }) });
    await expect(renderEditedImage(source, stack, { widthPx: 16, heightPx: 12 }, { createCanvas: f.create })).resolves.toBeDefined();
  });
});

describe("drawAnnotation", () => {
  it("scales an arrow from output pixels and adds its head", () => {
    const { canvas, calls } = recordingCanvas(10, 10);
    const ctx = canvas.getContext("2d") as unknown as CanvasRenderingContext2D;
    drawAnnotation(ctx, { kind: "annotate", shape: { shape: "arrow", x1: 0, y1: 0, x2: 100, y2: 0 }, color: "#ff0000", strokeWidth: 4, opacity: 0.5 }, 0.5, 0.5);
    expect(calls).toContainEqual(["moveTo", 0, 0]);
    expect(calls).toContainEqual(["lineTo", 50, 0]);
    expect(calls.filter((call) => call[0] === "fill")).toHaveLength(1);
    expect((ctx as unknown as Record<string, unknown>).globalAlpha).toBe(0.5);
    expect((ctx as unknown as Record<string, unknown>).lineWidth).toBe(2);
  });
});

describe("preview sizing", () => {
  it("caps the preview buffer at EDIT_PREVIEW_MAX_EDGE_PX, never upscaling", () => {
    expect(EDIT_PREVIEW_MAX_EDGE_PX).toBe(2048);
    expect(previewBufferSize({ widthPx: 8000, heightPx: 6000 })).toEqual({ widthPx: 2048, heightPx: 1536 });
    expect(previewBufferSize({ widthPx: 3000, heightPx: 6000 })).toEqual({ widthPx: 1024, heightPx: 2048 });
    expect(previewBufferSize({ widthPx: 160, heightPx: 120 })).toEqual({ widthPx: 160, heightPx: 120 });
  });

  it("fits the output into the stage, floored, within the cap", () => {
    expect(previewTargetSize({ widthPx: 90, heightPx: 160 }, { widthPx: 800, heightPx: 600 })).toEqual({ widthPx: 337, heightPx: 600 });
    expect(previewTargetSize({ widthPx: 8000, heightPx: 2000 }, { widthPx: 5000, heightPx: 5000 })).toEqual({ widthPx: 2048, heightPx: 512 });
  });

  it("builds a downscaled buffer for a large decode, addressed in natural pixels", () => {
    const f = factory();
    const decoded: DecodedImage = { image: IMAGE, widthPx: 4096, heightPx: 3072, release: () => {} };
    const source = previewSource(decoded, EDIT_PREVIEW_MAX_EDGE_PX, f.create);
    expect(source.naturalPx).toEqual({ widthPx: 4096, heightPx: 3072 });
    expect([f.made[0].canvas.width, f.made[0].canvas.height]).toEqual([2048, 1536]);
    expect(f.made[0].calls).toContainEqual(["drawImage", IMAGE, 0, 0, 2048, 1536]);
    // A decode within the cap is used as it is.
    const small = previewSource({ ...decoded, widthPx: 160, heightPx: 120 }, EDIT_PREVIEW_MAX_EDGE_PX, f.create);
    expect(small.image).toBe(IMAGE);
    expect(f.made).toHaveLength(1);
  });
});

describe("createDecodeCache", () => {
  it("decodes each part once, however often it is read, and retries after a failure", async () => {
    const decode = vi.fn(async () => ({ image: IMAGE, widthPx: 1, heightPx: 1, release: vi.fn() }));
    const cache = createDecodeCache(decode);
    const bytes = vi.fn(() => new Uint8Array([1]));
    const [a, b] = await Promise.all([cache.get("part-1", bytes, "image/png"), cache.get("part-1", bytes, "image/png")]);
    expect(a).toBe(b);
    await cache.get("part-1", bytes, "image/png");
    expect(decode).toHaveBeenCalledTimes(1);
    expect(bytes).toHaveBeenCalledTimes(1);

    const failing = vi.fn().mockRejectedValueOnce(new Error("bad")).mockResolvedValue({ image: IMAGE, widthPx: 1, heightPx: 1, release: () => {} });
    const retrying = createDecodeCache(failing);
    await expect(retrying.get("p", bytes, "image/png")).rejects.toThrow("bad");
    await expect(retrying.get("p", bytes, "image/png")).resolves.toMatchObject({ widthPx: 1 });
    expect(failing).toHaveBeenCalledTimes(2);

    cache.clear();
    await Promise.resolve();
    expect(a.release).toHaveBeenCalled();
  });
});

describe("encodeRender", () => {
  it("paints a JPEG's background before the picture, and passes PNG through", async () => {
    const f = factory();
    const rendered = recordingCanvas(4, 3);
    const jpeg = await encodeRender(rendered.canvas, "image/jpeg", { quality: 0.8, background: "#123456" }, f.create);
    expect(jpeg).toEqual(new Uint8Array([1, 2, 3]));
    const flattened = f.made[0];
    expect(flattened.calls[0]).toEqual(["fillRect", "#123456", 0, 0, 4, 3]);
    expect(flattened.calls[1]).toEqual(["drawImage", rendered.canvas, 0, 0]);
    expect(flattened.calls).toContainEqual(["toBlob", "image/jpeg", 0.8]);
    // The temporary canvas is released.
    expect([flattened.canvas.width, flattened.canvas.height]).toEqual([0, 0]);

    await encodeRender(rendered.canvas, "image/png", {}, f.create);
    expect(f.made).toHaveLength(1);
    expect(rendered.calls).toContainEqual(["toBlob", "image/png", undefined]);
  });
});
