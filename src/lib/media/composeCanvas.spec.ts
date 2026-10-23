// SPEC-109: the assembly renderer, under a stub 2-D context that keeps a
// real RGBA buffer (jsdom has no canvas): drawImage samples the source
// nearest-neighbour, fillRect fills, both composite source-over with
// globalAlpha (and multiply, to prove the blend reaches the canvas). The
// rendered output is asserted the way SPEC-109 says it may be — its pixel
// dimensions, the alpha of a background sample and the colour at the centre
// of each placed item, read back with getImageData — never as bytes. Also:
// the source-rectangle arithmetic for cover/actual from planAssembly's own
// output, chunking and cancellation at chunk boundaries, the decode cache's
// single-decode property and SVG rasterisation at the placed size.
import { describe, expect, it, vi } from "vitest";
import { ASSEMBLY_PLAN_DEFAULTS, DEFAULT_OVERLAY_TRANSFORM, planAssembly, type AssemblyPlan, type PlacedItem } from "../core/assemblyLayout";
import type { MediaPart } from "../core/media";
import type { SheetGeometry } from "../core/sheetLayout";
import {
  composeToCanvas,
  createSourceCache,
  drawingChunks,
  encodeCanvas,
  MAX_SVG_RASTER_PIXELS,
  prepareSources,
  sourceKey,
  svgRasterSize,
  toPixelRect,
  type ComposeSource,
} from "./composeCanvas";
import type { DecodedImage } from "./renderEdits";

type Rgba = [number, number, number, number];

/** A fake decoded image: `pixel(x, y)` in its own pixels. */
interface FakeImage {
  width: number;
  height: number;
  pixel: (x: number, y: number) => Rgba;
}

function solid(color: Rgba, width = 100, height = 100): FakeImage {
  return { width, height, pixel: () => color };
}

function hex(value: string): Rgba {
  return [1, 3, 5].map((i) => Number.parseInt(value.slice(i, i + 2), 16)).concat(255) as Rgba;
}

function blendChannel(mode: string, backdrop: number, source: number): number {
  if (mode === "multiply") {
    return (backdrop * source) / 255;
  }
  if (mode === "screen") {
    return backdrop + source - (backdrop * source) / 255;
  }
  return source;
}

/** A canvas whose 2-D context paints into a real RGBA buffer and records
 * the compositing state of every paint. */
function pixelCanvas(width: number, height: number) {
  const pixels = new Uint8ClampedArray(width * height * 4);
  const paints: Array<{ op: string; alpha: number; composite: string }> = [];
  const stack: Array<{ alpha: number; composite: string; fill: string }> = [];
  const ctx = {
    globalAlpha: 1,
    globalCompositeOperation: "source-over",
    fillStyle: "#000000",
    imageSmoothingEnabled: false,
    imageSmoothingQuality: "low",
    save() {
      stack.push({ alpha: ctx.globalAlpha, composite: ctx.globalCompositeOperation, fill: ctx.fillStyle });
    },
    restore() {
      const top = stack.pop();
      if (top) {
        ctx.globalAlpha = top.alpha;
        ctx.globalCompositeOperation = top.composite;
        ctx.fillStyle = top.fill;
      }
    },
    setTransform() {},
    clearRect(x: number, y: number, w: number, h: number) {
      for (let row = y; row < y + h; row += 1) {
        for (let col = x; col < x + w; col += 1) {
          pixels.set([0, 0, 0, 0], (row * width + col) * 4);
        }
      }
    },
    paint(col: number, row: number, color: Rgba) {
      if (col < 0 || row < 0 || col >= width || row >= height) {
        return;
      }
      const at = (row * width + col) * 4;
      const alpha = (color[3] / 255) * ctx.globalAlpha;
      const backdropAlpha = pixels[at + 3] / 255;
      const outAlpha = alpha + backdropAlpha * (1 - alpha);
      for (let c = 0; c < 3; c += 1) {
        const mixed = backdropAlpha > 0 ? blendChannel(ctx.globalCompositeOperation, pixels[at + c], color[c]) : color[c];
        const value = outAlpha > 0 ? (mixed * alpha + pixels[at + c] * backdropAlpha * (1 - alpha)) / outAlpha : 0;
        pixels[at + c] = Math.round(value);
      }
      pixels[at + 3] = Math.round(outAlpha * 255);
    },
    fillRect(x: number, y: number, w: number, h: number) {
      paints.push({ op: "fillRect", alpha: ctx.globalAlpha, composite: ctx.globalCompositeOperation });
      const color = hex(ctx.fillStyle);
      for (let row = y; row < y + h; row += 1) {
        for (let col = x; col < x + w; col += 1) {
          ctx.paint(col, row, color);
        }
      }
    },
    drawImage(image: FakeImage, sx: number, sy: number, sw: number, sh: number, dx: number, dy: number, dw: number, dh: number) {
      paints.push({ op: "drawImage", alpha: ctx.globalAlpha, composite: ctx.globalCompositeOperation });
      for (let row = dy; row < dy + dh; row += 1) {
        for (let col = dx; col < dx + dw; col += 1) {
          const u = Math.floor(sx + ((col + 0.5 - dx) / dw) * sw);
          const v = Math.floor(sy + ((row + 0.5 - dy) / dh) * sh);
          ctx.paint(col, row, image.pixel(u, v));
        }
      }
    },
    getImageData(x: number, y: number, w: number, h: number) {
      const data = new Uint8ClampedArray(w * h * 4);
      for (let row = 0; row < h; row += 1) {
        data.set(pixels.slice(((y + row) * width + x) * 4, ((y + row) * width + x + w) * 4), row * w * 4);
      }
      return { data, width: w, height: h };
    },
  };
  const canvas = { width, height, getContext: () => ctx };
  const sample = (x: number, y: number): number[] => Array.from(ctx.getImageData(Math.floor(x), Math.floor(y), 1, 1).data);
  return { canvas: canvas as unknown as HTMLCanvasElement, sample, paints };
}

/** The canvas pixel at the centre of a placed item. */
function centreOf(item: PlacedItem, sx: number, sy: number): [number, number] {
  const rect = toPixelRect(item, sx, sy);
  return [rect.x + rect.width / 2, rect.y + rect.height / 2];
}

function sourcesOf(images: FakeImage[], itemSizes: Array<[number, number]>): Map<number, ComposeSource> {
  return new Map(
    images.map((image, i) => [
      i,
      { image: image as unknown as CanvasImageSource, scaleX: image.width / itemSizes[i][0], scaleY: image.height / itemSizes[i][1] },
    ]),
  );
}

/** A4 landscape-ish: 30 × 20 cm, 1 cm margins. */
const SHEET: SheetGeometry = { widthCm: 30, heightCm: 20, content: { xCm: 1, yCm: 1, widthCm: 28, heightCm: 18 } };

function plan(overrides: Partial<AssemblyPlan> = {}): AssemblyPlan {
  return { ...ASSEMBLY_PLAN_DEFAULTS, ...overrides };
}

const RED: Rgba = [220, 30, 30, 255];
const BLUE: Rgba = [30, 30, 220, 255];

describe("composeToCanvas — sampled pixels", () => {
  const items = [
    { id: "a", widthPx: 100, heightPx: 100 },
    { id: "b", widthPx: 200, heightPx: 100 },
  ];
  const placed = planAssembly(items, plan({ columns: 2 }), SHEET);
  const sources = sourcesOf([solid(RED), solid(BLUE, 200, 100)], [[100, 100], [200, 100]]);

  it("draws at the canvas's pixel size, transparent background, each item's colour at its centre", async () => {
    const { canvas, sample } = pixelCanvas(300, 200);
    expect(await composeToCanvas(canvas, placed, { geometry: SHEET, background: "transparent" }, sources)).toBe(true);
    expect([canvas.width, canvas.height]).toEqual([300, 200]);
    // A background sample (inside the margin) is fully transparent.
    expect(sample(2, 2)[3]).toBe(0);
    expect(sample(150, 199)[3]).toBe(0);
    const [ax, ay] = centreOf(placed[0], 10, 10);
    const [bx, by] = centreOf(placed[1], 10, 10);
    expect(sample(ax, ay)).toEqual(RED);
    expect(sample(bx, by)).toEqual(BLUE);
  });

  it("paints an opaque sheet background", async () => {
    const { canvas, sample } = pixelCanvas(300, 200);
    await composeToCanvas(canvas, placed, { geometry: SHEET, background: "#ffffff" }, sources);
    expect(sample(2, 2)).toEqual([255, 255, 255, 255]);
  });

  it("is the same picture at two sizes — the preview and the export are one function", async () => {
    const small = pixelCanvas(150, 100);
    const large = pixelCanvas(600, 400);
    await composeToCanvas(small.canvas, placed, { geometry: SHEET, background: "#ffffff" }, sources);
    await composeToCanvas(large.canvas, placed, { geometry: SHEET, background: "#ffffff" }, sources);
    for (const item of placed) {
      const [sx, sy] = centreOf(item, 5, 5);
      const [lx, ly] = centreOf(item, 20, 20);
      expect(small.sample(sx, sy)).toEqual(large.sample(lx, ly));
    }
    expect(small.sample(1, 1)).toEqual(large.sample(4, 4));
  });

  it("letterboxes a contain fit in the cell background", async () => {
    const boxed = planAssembly(items, plan({ columns: 2, cellBackground: "#00ff00" }), SHEET);
    const { canvas, sample } = pixelCanvas(300, 200);
    await composeToCanvas(canvas, boxed, { geometry: SHEET, background: "transparent" }, sources);
    // Cell 2 is 13.9 × 18 cm; the 2:1 item fills 13.9 × 6.95 across its
    // middle — above and below it, the cell's green.
    const cell = toPixelRect(boxed[1].cell, 10, 10);
    expect(sample(cell.x + cell.width / 2, cell.y + 3)).toEqual([0, 255, 0, 255]);
    expect(sample(...centreOf(boxed[1], 10, 10))).toEqual(BLUE);
    // The gap between the cells stays the sheet's (transparent).
    expect(sample(cell.x - 1, cell.y + 10)[3]).toBe(0);
  });

  it("crops a cover fit to the source rectangle planAssembly computed", async () => {
    // Left half red, right half blue; covering a tall cell keeps the middle.
    const halves: FakeImage = { width: 200, height: 100, pixel: (x) => (x < 100 ? RED : BLUE) };
    const covered = planAssembly([{ id: "h", widthPx: 200, heightPx: 100 }], plan({ columns: 2, fit: "cover", align: "left" }), SHEET);
    expect(covered[0].sourcePx).toEqual({ x: 0, y: 0, width: 77.222, height: 100 });
    const { canvas, sample } = pixelCanvas(300, 200);
    await composeToCanvas(canvas, covered, { geometry: SHEET, background: "transparent" }, sourcesOf([halves], [[200, 100]]));
    const rect = toPixelRect(covered[0], 10, 10);
    // Aligned left: only red survives, all the way across.
    expect(sample(rect.x + 1, rect.y + 50)).toEqual(RED);
    expect(sample(rect.x + rect.width - 1, rect.y + 50)).toEqual(RED);
  });

  it("hands drawImage the source rectangle scaled onto the decoded image (actual fit, 2× source)", async () => {
    const drawImage = vi.fn();
    const ctx = {
      save() {},
      restore() {},
      setTransform() {},
      clearRect() {},
      fillRect() {},
      drawImage,
    };
    const canvas = { width: 300, height: 200, getContext: () => ctx } as unknown as HTMLCanvasElement;
    // 1000 × 500 px at 72 dpi is 35.278 × 17.639 cm: clipped to the 13.9 cm
    // cell, 13.9 / 35.278 of its 1000 px survive.
    const actual = planAssembly([{ id: "big", widthPx: 1000, heightPx: 500 }], plan({ columns: 2, fit: "actual", align: "top-left" }), SHEET, 72);
    expect(actual[0].sourcePx).toEqual({ x: 0, y: 0, width: 394.013, height: 500 });
    const image = { tag: "decoded" } as unknown as CanvasImageSource;
    await composeToCanvas(canvas, actual, { geometry: SHEET, background: "transparent" }, new Map([[0, { image, scaleX: 2, scaleY: 2 }]]));
    expect(drawImage).toHaveBeenCalledWith(image, 0, 0, 788.026, 1000, 10, 10, 139, 176);
  });

  it("draws overlay items back to front, with opacity and blend", async () => {
    const overlay = planAssembly(
      [
        { id: "back", widthPx: 100, heightPx: 100 },
        { id: "front", widthPx: 100, heightPx: 100, transform: { ...DEFAULT_OVERLAY_TRANSFORM, scalePercent: 50, opacity: 50 } },
        { id: "mul", widthPx: 100, heightPx: 100, transform: { ...DEFAULT_OVERLAY_TRANSFORM, xPercent: 30, scalePercent: 20, blend: "multiply" } },
      ],
      plan({ layout: "overlay" }),
      SHEET,
    );
    const { canvas, sample, paints } = pixelCanvas(300, 200);
    await composeToCanvas(
      canvas,
      overlay,
      { geometry: SHEET, background: "#ffffff" },
      sourcesOf([solid(RED), solid(BLUE), solid([128, 128, 128, 255])], [[100, 100], [100, 100], [100, 100]]),
    );
    // The centre: blue at half opacity over red.
    expect(sample(150, 100)).toEqual([125, 30, 125, 255]);
    // Off the front item, the back item shows.
    const back = toPixelRect(overlay[0], 10, 10);
    expect(sample(back.x + 3, back.y + 3)).toEqual(RED);
    // The multiplied grey item (left of the front one, over the red) darkens
    // what is under it rather than covering it.
    const mul = toPixelRect(overlay[2], 10, 10);
    expect(sample(mul.x + mul.width / 2, mul.y + mul.height / 2)).toEqual([110, 15, 15, 255]);
    expect(paints.filter((paint) => paint.op === "drawImage").map((paint) => [paint.alpha, paint.composite])).toEqual([
      [1, "source-over"],
      [0.5, "source-over"],
      [1, "multiply"],
    ]);
  });

  it("returns false and draws nothing without a 2-D context", async () => {
    const canvas = { width: 10, height: 10, getContext: () => null } as unknown as HTMLCanvasElement;
    expect(await composeToCanvas(canvas, placed, { geometry: SHEET, background: "#ffffff" }, sources)).toBe(false);
  });
});

describe("chunking and cancellation", () => {
  const nine = Array.from({ length: 9 }, (_, i) => ({ id: `i${i}`, widthPx: 100, heightPx: 100 }));

  it("draws row by row, yielding between rows and reporting progress", async () => {
    const placed = planAssembly(nine, plan({ columns: 3 }), SHEET);
    expect(drawingChunks(placed).map((chunk) => chunk.map((item) => item.index))).toEqual([[0, 1, 2], [3, 4, 5], [6, 7, 8]]);
    const between = vi.fn(async () => {});
    const progress: string[] = [];
    const { canvas } = pixelCanvas(60, 40);
    await composeToCanvas(canvas, placed, { geometry: SHEET, background: "#ffffff" }, new Map(), {
      between,
      onProgress: (done, total) => progress.push(`Item ${done} of ${total}`),
    });
    expect(between).toHaveBeenCalledTimes(2);
    expect(progress).toEqual(["Item 0 of 9", "Item 3 of 9", "Item 6 of 9", "Item 9 of 9"]);
  });

  it("draws overlay item by item", () => {
    const placed = planAssembly(nine.slice(0, 3), plan({ layout: "overlay" }), SHEET);
    expect(drawingChunks(placed)).toHaveLength(3);
  });

  it("stops at the next chunk boundary when cancelled", async () => {
    const placed = planAssembly(nine, plan({ columns: 3 }), SHEET);
    let stop = false;
    const progress: number[] = [];
    const { canvas, paints } = pixelCanvas(60, 40);
    const sources = new Map(placed.map((item) => [item.index, { image: solid(RED) as unknown as CanvasImageSource, scaleX: 1, scaleY: 1 }]));
    const finished = await composeToCanvas(canvas, placed, { geometry: SHEET, background: "transparent" }, sources, {
      between: async () => {
        stop = true;
      },
      cancelled: () => stop,
      onProgress: (done) => progress.push(done),
    });
    expect(finished).toBe(false);
    expect(progress).toEqual([0, 3]);
    expect(paints.filter((paint) => paint.op === "drawImage")).toHaveLength(3);
  });
});

function decoded(widthPx: number, heightPx: number): DecodedImage {
  return { image: { widthPx } as unknown as CanvasImageSource, widthPx, heightPx, release: vi.fn() };
}

describe("sources", () => {
  const png: MediaPart = { mediaId: "png-1", contentType: "image/png", data: "AQID" };
  const svg: MediaPart = { mediaId: "svg-1", contentType: "image/svg+xml", data: "PHN2Zy8+" };

  it("decodes a part once, however many items and renders read it", async () => {
    const decode = vi.fn(async () => decoded(64, 40));
    const cache = createSourceCache({ decode, rasteriseSvg: vi.fn() });
    const first = await cache.decodeSource(png);
    const again = await cache.decodeSource(png);
    expect(again).toBe(first);
    expect(decode).toHaveBeenCalledTimes(1);
    expect(decode).toHaveBeenCalledWith(new Uint8Array([1, 2, 3]), "image/png");
    expect(cache.size()).toBe(1);
  });

  it("forgets a failed decode so a retry decodes again", async () => {
    const decode = vi.fn().mockRejectedValueOnce(new Error("bad")).mockResolvedValue(decoded(1, 1));
    const cache = createSourceCache({ decode, rasteriseSvg: vi.fn() });
    await expect(cache.decodeSource(png)).rejects.toThrow("bad");
    await cache.decodeSource(png);
    expect(decode).toHaveBeenCalledTimes(2);
  });

  it("rasterises an SVG at the placed pixel size, cached per size", async () => {
    const rasteriseSvg = vi.fn(async (_bytes: Uint8Array, size?: { widthPx: number; heightPx: number }) => ({
      data: new Uint8Array([9]),
      widthPx: size?.widthPx ?? 0,
      heightPx: size?.heightPx ?? 0,
    }));
    const decode = vi.fn(async () => decoded(1200, 600));
    const cache = createSourceCache({ decode, rasteriseSvg });
    await cache.decodeSource(svg, { widthPx: 1200, heightPx: 600 });
    await cache.decodeSource(svg, { widthPx: 1200, heightPx: 600 });
    expect(rasteriseSvg).toHaveBeenCalledTimes(1);
    expect(rasteriseSvg.mock.calls[0][1]).toEqual({ widthPx: 1200, heightPx: 600 });
    expect(decode).toHaveBeenCalledWith(new Uint8Array([9]), "image/png");
    expect(sourceKey(svg, { widthPx: 1200, heightPx: 600 })).toBe("svg-1@1200x600");
    expect(sourceKey(png, { widthPx: 1200, heightPx: 600 })).toBe("png-1");
  });

  it("bounds an SVG's raster, aspect kept, and keeps only the latest two sizes of it", async () => {
    // An overlay logo scaled tenfold on an A4 sheet at 150 dpi.
    const huge = svgRasterSize({ widthPx: 12400, heightPx: 17540 });
    expect(huge.widthPx * huge.heightPx).toBeLessThanOrEqual(MAX_SVG_RASTER_PIXELS);
    expect(huge.widthPx / huge.heightPx).toBeCloseTo(12400 / 17540, 3);
    expect(svgRasterSize({ widthPx: 40000, heightPx: 10 })).toEqual({ widthPx: 16384, heightPx: 4 });
    expect(svgRasterSize({ widthPx: 1390, heightPx: 695 })).toEqual({ widthPx: 1390, heightPx: 695 });

    const rasteriseSvg = vi.fn(async (_bytes: Uint8Array, size?: { widthPx: number; heightPx: number }) => ({
      data: new Uint8Array([9]),
      widthPx: size?.widthPx ?? 0,
      heightPx: size?.heightPx ?? 0,
    }));
    const cache = createSourceCache({ decode: vi.fn(async () => decoded(1, 1)), rasteriseSvg });
    for (const widthPx of [100, 200, 300]) {
      await cache.decodeSource(svg, { widthPx, heightPx: widthPx });
    }
    expect(cache.size()).toBe(2);
    // The newest two are served from the cache; the oldest decodes again.
    await cache.decodeSource(svg, { widthPx: 300, heightPx: 300 });
    await cache.decodeSource(svg, { widthPx: 200, heightPx: 200 });
    expect(rasteriseSvg).toHaveBeenCalledTimes(3);
    await cache.decodeSource(svg, { widthPx: 100, heightPx: 100 });
    expect(rasteriseSvg).toHaveBeenCalledTimes(4);
  });

  it("prepares an SVG source at the size its item is drawn, scaled onto the item's pixels", async () => {
    // A 300 × 150 SVG contained in a 13.9 × 18 cm cell at 100 px/cm: drawn
    // 1390 × 695 px, so it is rasterised at exactly that size.
    const placed = planAssembly([{ id: "logo", widthPx: 300, heightPx: 150 }], plan({ columns: 2 }), SHEET);
    const decode = vi.fn(async (_part: MediaPart, size?: { widthPx: number; heightPx: number }) => decoded(size?.widthPx ?? 300, size?.heightPx ?? 150));
    const { sources, failed } = await prepareSources(placed, [{ part: svg, widthPx: 300, heightPx: 150 }], 100, 100, decode);
    expect(decode).toHaveBeenCalledWith(svg, { widthPx: 1390, heightPx: 695 });
    expect(failed).toEqual([]);
    expect(sources.get(0)).toMatchObject({ scaleX: 1390 / 300, scaleY: 695 / 150 });
  });

  it("decodes a duplicated item's part once and records a failure without throwing", async () => {
    const placed = planAssembly(
      Array.from({ length: 3 }, (_, i) => ({ id: `d${i}`, widthPx: 64, heightPx: 40 })),
      plan(),
      SHEET,
    );
    const cache = createSourceCache({ decode: vi.fn(async () => decoded(64, 40)), rasteriseSvg: vi.fn() });
    const broken: MediaPart = { mediaId: "broken", contentType: "image/png", data: "AA==" };
    const decode = (part: MediaPart) => (part === broken ? Promise.reject(new Error("no")) : cache.decodeSource(part));
    const input = { part: png, widthPx: 64, heightPx: 40 };
    const result = await prepareSources(placed, [input, input, { ...input, part: broken }], 10, 10, decode);
    expect(cache.size()).toBe(1);
    expect([...result.sources.keys()]).toEqual([0, 1]);
    expect(result.failed).toEqual([2]);
  });

  it("stops preparing when cancelled", async () => {
    const placed = planAssembly(Array.from({ length: 3 }, (_, i) => ({ id: `c${i}`, widthPx: 1, heightPx: 1 })), plan(), SHEET);
    const input = { part: png, widthPx: 1, heightPx: 1 };
    const result = await prepareSources(placed, [input, input, input], 1, 1, async () => decoded(1, 1), { cancelled: () => true });
    expect(result.cancelled).toBe(true);
    expect(result.sources.size).toBe(0);
  });
});

describe("encodeCanvas", () => {
  it("encodes through toBlob with the type and quality", async () => {
    const toBlob = vi.fn((done: (blob: Blob) => void) => {
      done({ arrayBuffer: async () => new Uint8Array([7, 8]).buffer } as unknown as Blob);
    });
    const canvas = { toBlob } as unknown as HTMLCanvasElement;
    expect(await encodeCanvas(canvas, "image/jpeg", 0.9)).toEqual(new Uint8Array([7, 8]));
    expect(toBlob.mock.calls[0].slice(1)).toEqual(["image/jpeg", 0.9]);
  });

  it("rejects when the browser cannot encode", async () => {
    const canvas = { toBlob: (done: (blob: Blob | null) => void) => done(null) } as unknown as HTMLCanvasElement;
    await expect(encodeCanvas(canvas, "image/png")).rejects.toThrow("could not be encoded");
  });
});
