// Rasterises an SVG to PNG at insert time for a `.docx`-bound document
// (SPEC-024 AC-2): `.docx` has no SVG picture representation, so the
// original vector is converted once, on insert, and never retained — unlike
// every other reader/writer module in `lib/`, this one is deliberately not
// framework-free: `HTMLCanvasElement`/`Image` decoding has no meaningful
// pure-JS equivalent, and it only ever runs from the browser-only insert
// flow (never from the parse/serialise pipeline reader.ts/writer.ts share
// with Vitest+jsdom, which cannot decode images).
import { decodeImageDimensions } from "./imageDimensions";

export class SvgConversionError extends Error {}

/** SVG's own default intrinsic size when neither width/height nor a usable
 * viewBox is present (CSS Images spec's replaced-element default). */
const DEFAULT_SVG_WIDTH_PX = 300;
const DEFAULT_SVG_HEIGHT_PX = 150;

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.addEventListener("load", () => resolve(img));
    img.addEventListener("error", () => reject(new SvgConversionError("That image could not be read")));
    img.src = url;
  });
}

export interface ConvertedPng {
  data: Uint8Array;
  widthPx: number;
  heightPx: number;
}

/** Converts SVG bytes to PNG bytes, rendered at `size` when given (SPEC-108
 * rasterises at the size a sheet places the SVG at), else at the SVG's own
 * declared pixel size (its `width`/`height`, falling back to its `viewBox`,
 * falling back to the CSS default 300x150 — the same order
 * decodeImageDimensions resolves for the *unconverted* SVG's own natural
 * size). Throws
 * SvgConversionError for malformed SVG or when canvas is unavailable — the
 * caller (the insert flow) shows this as an error and inserts nothing,
 * matching AC-33/34's reject-cleanly pattern for an oversized image. */
export async function convertSvgToPng(
  svgBytes: Uint8Array,
  size?: { widthPx: number; heightPx: number },
): Promise<ConvertedPng> {
  const dims = size ?? decodeImageDimensions(svgBytes, "image/svg+xml");
  const widthPx = Math.max(1, Math.round(dims?.widthPx ?? DEFAULT_SVG_WIDTH_PX));
  const heightPx = Math.max(1, Math.round(dims?.heightPx ?? DEFAULT_SVG_HEIGHT_PX));

  const blob = new Blob([svgBytes], { type: "image/svg+xml" });
  const url = URL.createObjectURL(blob);
  let canvas: HTMLCanvasElement | null = null;
  try {
    const image = await loadImage(url);
    const target = document.createElement("canvas");
    canvas = target;
    target.width = widthPx;
    target.height = heightPx;
    const ctx = target.getContext("2d");
    if (!ctx) {
      throw new SvgConversionError("This browser cannot convert SVG images");
    }
    ctx.drawImage(image, 0, 0, widthPx, heightPx);
    const pngBlob = await new Promise<Blob>((resolve, reject) => {
      target.toBlob((result) => {
        if (result) {
          resolve(result);
        } else {
          reject(new SvgConversionError("That image could not be converted"));
        }
      }, "image/png");
    });
    const buffer = await pngBlob.arrayBuffer();
    return { data: new Uint8Array(buffer), widthPx, heightPx };
  } finally {
    URL.revokeObjectURL(url);
    if (canvas) {
      // Release the backing store now rather than at the next GC (SPEC-109).
      canvas.width = 0;
      canvas.height = 0;
    }
  }
}
