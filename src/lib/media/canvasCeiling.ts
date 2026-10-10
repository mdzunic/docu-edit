// The canvas ceiling, measured rather than assumed (SPEC-109's
// `probeMaxCanvasPixels`, first needed here by the image editor, SPEC-112).
// Browsers disagree about the largest canvas they will allocate, and the
// ones that refuse do it by handing back a blank surface rather than
// throwing — so a too-large edit would silently save a white image. The
// probe allocates a candidate area, paints one pixel in its far corner,
// reads it back, and halves on failure down to a floor every target browser
// supports; the answer is cached for the session.
//
// Browser-bound (canvas), the same exception core/svgToPng.ts takes; no Vue.
import type { PixelSize } from "../core/imageEdits";

/** An area every target browser allocates (iOS Safari's documented
 * 16,777,216-pixel cap): work at or under it needs no probe. */
export const GUARANTEED_CANVAS_PIXELS = 4096 * 4096;

/** Where the probe starts: Chromium's own area cap. */
export const CANVAS_PROBE_START_PIXELS = 16384 * 16384;

export type CanvasFactory = (widthPx: number, heightPx: number) => HTMLCanvasElement;

const browserCanvas: CanvasFactory = (widthPx, heightPx) => {
  const canvas = document.createElement("canvas");
  canvas.width = widthPx;
  canvas.height = heightPx;
  return canvas;
};

let probed: number | null = null;

/** Whether a square canvas of `areaPx` really holds pixels. */
function canAllocate(areaPx: number, create: CanvasFactory): boolean {
  const side = Math.floor(Math.sqrt(areaPx));
  let canvas: HTMLCanvasElement | null = null;
  try {
    canvas = create(side, side);
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      return false;
    }
    ctx.fillStyle = "#010203";
    ctx.fillRect(side - 1, side - 1, 1, 1);
    const pixel = ctx.getImageData(side - 1, side - 1, 1, 1).data;
    return pixel[0] === 1 && pixel[1] === 2 && pixel[2] === 3 && pixel[3] === 255;
  } catch {
    return false;
  } finally {
    if (canvas) {
      // Release the backing store at once rather than at the next GC.
      canvas.width = 0;
      canvas.height = 0;
    }
  }
}

/** The largest canvas area (pixels) this browser verifiably allocates —
 * probed once per session, then cached. */
export function probeMaxCanvasPixels(create: CanvasFactory = browserCanvas): number {
  if (probed !== null) {
    return probed;
  }
  let area = CANVAS_PROBE_START_PIXELS;
  while (area > GUARANTEED_CANVAS_PIXELS && !canAllocate(area, create)) {
    area = Math.floor(area / 2);
  }
  probed = Math.max(area, GUARANTEED_CANVAS_PIXELS);
  return probed;
}

/** Forgets the cached probe (tests). */
export function resetCanvasProbe(): void {
  probed = null;
}

/** The ceiling to hold `areaPx` against: an area under the guaranteed floor
 * needs no probe; anything larger is measured. */
export function canvasCeilingFor(areaPx: number, probe: () => number = probeMaxCanvasPixels): number {
  return areaPx <= GUARANTEED_CANVAS_PIXELS ? GUARANTEED_CANVAS_PIXELS : probe();
}

/** "12.6 megapixels". */
export function megapixelsLabel(areaPx: number): string {
  return `${(areaPx / 1_000_000).toFixed(1)} megapixels`;
}

/** Why an image of `sizePx` cannot be edited under `ceilingPx`, naming the
 * limit — or null when it fits. */
export function canvasCeilingRefusal(sizePx: PixelSize, ceilingPx: number): string | null {
  const area = sizePx.widthPx * sizePx.heightPx;
  if (area <= ceilingPx) {
    return null;
  }
  return (
    `This image is ${sizePx.widthPx} × ${sizePx.heightPx} px (${megapixelsLabel(area)}), ` +
    `larger than this browser can edit: its canvas limit is ${megapixelsLabel(ceilingPx)}.`
  );
}
