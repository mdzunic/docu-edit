// SPEC-112 (SPEC-109's probe): the canvas ceiling is measured — allocate,
// paint one pixel, read it back, halve on failure — cached for the session,
// skipped for areas every browser holds, and a refusal names the limit.
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  CANVAS_PROBE_START_PIXELS,
  canvasCeilingFor,
  canvasCeilingRefusal,
  GUARANTEED_CANVAS_PIXELS,
  probeMaxCanvasPixels,
  resetCanvasProbe,
} from "./canvasCeiling";

afterEach(() => resetCanvasProbe());

/** Canvases that really hold pixels only up to `limit` — beyond it they
 * read back blank, as refusing browsers do. */
function limitedFactory(limit: number) {
  const sizes: number[] = [];
  const create = (width: number, height: number) => {
    sizes.push(width * height);
    let painted = false;
    const canvas = {
      width,
      height,
      getContext: () => ({
        fillStyle: "",
        fillRect: () => {
          painted = width * height <= limit;
        },
        getImageData: () => ({ data: painted ? new Uint8ClampedArray([1, 2, 3, 255]) : new Uint8ClampedArray(4) }),
      }),
    };
    return canvas as unknown as HTMLCanvasElement;
  };
  return { create, sizes };
}

describe("probeMaxCanvasPixels", () => {
  it("halves until a canvas holds its pixel, then caches the answer", () => {
    const f = limitedFactory(70_000_000);
    const ceiling = probeMaxCanvasPixels(f.create);
    // 268 435 456 → 134 217 728 → 67 108 864 (holds).
    expect(ceiling).toBe(CANVAS_PROBE_START_PIXELS / 4);
    expect(f.sizes).toHaveLength(3);
    expect(probeMaxCanvasPixels(f.create)).toBe(ceiling);
    expect(f.sizes).toHaveLength(3);
  });

  it("never reports less than the guaranteed floor", () => {
    expect(probeMaxCanvasPixels(limitedFactory(0).create)).toBe(GUARANTEED_CANVAS_PIXELS);
  });

  it("does not probe for an area every browser holds", () => {
    const probe = vi.fn(() => 1);
    expect(canvasCeilingFor(4000 * 3000, probe)).toBe(GUARANTEED_CANVAS_PIXELS);
    expect(probe).not.toHaveBeenCalled();
    expect(canvasCeilingFor(8000 * 6000, probe)).toBe(1);
  });
});

describe("canvasCeilingRefusal", () => {
  it("names the image's size and the limit", () => {
    expect(canvasCeilingRefusal({ widthPx: 100, heightPx: 100 }, 20_000)).toBeNull();
    expect(canvasCeilingRefusal({ widthPx: 20_000, heightPx: 15_000 }, 268_435_456)).toBe(
      "This image is 20000 × 15000 px (300.0 megapixels), larger than this browser can edit: its canvas limit is 268.4 megapixels.",
    );
  });
});
