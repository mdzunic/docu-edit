// Paints one already-extracted page (a `PdfPageInput`, SPEC-034's writer
// input) onto a canvas and returns PNG bytes (SPEC-104's deck image export).
// The raster is drawn from the very input the PDF is built from — text with
// `fillText` in the bundled face the PDF embeds (`faceCssFont`), images with
// `drawImage`, rects with `fillRect`/`strokeRect`, vector paths as canvas
// paths — so a PNG and a PDF of the same sheet page agree, and nothing is a
// screenshot of the editor.
//
// Browser-bound like core/svgToPng.ts (it needs a canvas, image decoding and
// `canvas.toBlob`); the painting itself is `paintPdfPage` over a 2D context,
// which the unit tests drive with a recording stand-in.
import { rectFillRgb, DEFAULT_RECT_STROKE_PT, type PdfPageInput } from "./pdfBuild";
import { faceCssFont } from "./pdfFonts";

/** The drawable a decoded page image becomes. */
export type RasterImage = CanvasImageSource;

/** The subset of CanvasRenderingContext2D `paintPdfPage` uses. */
export type RasterContext = Pick<
  CanvasRenderingContext2D,
  | "save"
  | "restore"
  | "fillRect"
  | "strokeRect"
  | "fillText"
  | "drawImage"
  | "beginPath"
  | "moveTo"
  | "lineTo"
  | "bezierCurveTo"
  | "closePath"
  | "fill"
  | "stroke"
  | "setLineDash"
> & {
  fillStyle: string | CanvasGradient | CanvasPattern;
  strokeStyle: string | CanvasGradient | CanvasPattern;
  lineWidth: number;
  font: string;
  lineCap: CanvasLineCap;
  lineJoin: CanvasLineJoin;
};

function channel(value: number): number {
  return Math.round(Math.min(1, Math.max(0, value)) * 255);
}

function cssRgb(rgb: readonly number[]): string {
  return `rgb(${channel(rgb[0])}, ${channel(rgb[1])}, ${channel(rgb[2])})`;
}

const CAPS: CanvasLineCap[] = ["butt", "round", "square"];
const JOINS: CanvasLineJoin[] = ["miter", "round", "bevel"];

/** Paints `page` at `scale` device px per point, in PDF's own paint order:
 * white paper, under-the-text rects, images and paths, the text, then the
 * on-top rects, images and paths. `images` maps an image item's `src` to its
 * decoded drawable; an item missing from it is skipped, as buildPdf skips an
 * undecodable image. */
export function paintPdfPage(
  ctx: RasterContext,
  page: PdfPageInput,
  scale: number,
  images: ReadonlyMap<string, RasterImage>,
): void {
  ctx.save();
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, page.widthPt * scale, page.heightPt * scale);

  const paintRects = (onTop: boolean) => {
    for (const rect of page.rects ?? []) {
      if (Boolean(rect.onTop) !== onTop) {
        continue;
      }
      const x = rect.xPt * scale;
      const y = rect.yPt * scale;
      const width = rect.widthPt * scale;
      const height = rect.heightPt * scale;
      const fill = rectFillRgb(rect.fill);
      if (fill) {
        ctx.fillStyle = cssRgb(fill);
        ctx.fillRect(x, y, width, height);
      }
      if (rect.stroke) {
        ctx.strokeStyle = cssRgb(rect.stroke);
        ctx.lineWidth = Math.max(0, rect.strokeWidthPt ?? DEFAULT_RECT_STROKE_PT) * scale;
        ctx.setLineDash([]);
        ctx.strokeRect(x, y, width, height);
      }
    }
  };

  const paintImages = (onTop: boolean) => {
    for (const item of page.images) {
      const image = images.get(item.src);
      if (!image || Boolean(item.onTop) !== onTop) {
        continue;
      }
      ctx.drawImage(image, item.xPt * scale, item.yPt * scale, item.widthPt * scale, item.heightPt * scale);
    }
  };

  const paintPaths = (onTop: boolean) => {
    for (const item of page.paths ?? []) {
      if (Boolean(item.onTop) !== onTop || (!item.stroke && !item.fill)) {
        continue;
      }
      ctx.beginPath();
      for (const subpath of item.subpaths) {
        ctx.moveTo(subpath.start.xPt * scale, subpath.start.yPt * scale);
        for (const segment of subpath.segments) {
          if (segment.c1XPt !== undefined && segment.c2XPt !== undefined) {
            ctx.bezierCurveTo(
              segment.c1XPt * scale,
              (segment.c1YPt ?? 0) * scale,
              segment.c2XPt * scale,
              (segment.c2YPt ?? 0) * scale,
              segment.toXPt * scale,
              segment.toYPt * scale,
            );
          } else {
            ctx.lineTo(segment.toXPt * scale, segment.toYPt * scale);
          }
        }
        if (subpath.closed) {
          ctx.closePath();
        }
      }
      if (item.fill) {
        ctx.fillStyle = cssRgb(item.fill);
        ctx.fill();
      }
      if (item.stroke) {
        ctx.strokeStyle = cssRgb(item.stroke.color);
        ctx.lineWidth = Math.max(0, item.stroke.widthPt) * scale;
        ctx.setLineDash((item.stroke.dashPt ?? []).map((dash) => dash * scale));
        ctx.lineCap = CAPS[item.stroke.cap ?? 0];
        ctx.lineJoin = JOINS[item.stroke.join ?? 0];
        ctx.stroke();
      }
    }
  };

  paintRects(false);
  paintImages(false);
  paintPaths(false);
  for (const item of page.items) {
    if (item.text.length === 0) {
      continue;
    }
    ctx.font = faceCssFont(item.face, item.sizePt * scale);
    ctx.fillStyle = cssRgb(item.color);
    // yPt is the baseline, which is where fillText's default alphabetic
    // baseline puts the glyphs.
    ctx.fillText(item.text, item.xPt * scale, item.yPt * scale);
  }
  paintRects(true);
  paintImages(true);
  paintPaths(true);
  ctx.restore();
}

async function decodeImage(src: string): Promise<RasterImage | null> {
  const image = new Image();
  image.src = src;
  try {
    await image.decode();
    return image;
  } catch {
    return null; // Undecodable: skipped, like buildPdf skips it.
  }
}

/** Waits for every bundled face the page's text uses, so `fillText` does not
 * fall back to a generic face while a registered one is still loading. */
async function loadFonts(page: PdfPageInput, scale: number): Promise<void> {
  if (typeof document === "undefined" || !document.fonts) {
    return;
  }
  const fonts = new Set(page.items.map((item) => faceCssFont(item.face, item.sizePt * scale)));
  await Promise.all(Array.from(fonts, (font) => document.fonts.load(font).catch(() => [])));
}

/** Renders `page` to PNG bytes `widthPx` pixels wide — an absolute width,
 * never a multiplier: the caller owns the pixel size (and its clamp). The
 * height follows the page's aspect. The canvas is released (`width = 0`) once
 * encoded, so a long export never holds more than one page's pixels. */
export async function rasterisePdfPage(page: PdfPageInput, widthPx: number): Promise<Uint8Array> {
  const width = Math.max(1, Math.round(widthPx));
  const scale = width / page.widthPt;
  const height = Math.max(1, Math.round(page.heightPt * scale));
  const sources = Array.from(new Set(page.images.map((item) => item.src)));
  const decoded = await Promise.all(sources.map(async (src) => [src, await decodeImage(src)] as const));
  const images = new Map<string, RasterImage>();
  for (const [src, image] of decoded) {
    if (image) {
      images.set(src, image);
    }
  }
  await loadFonts(page, scale);
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  try {
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      throw new Error("this browser cannot draw images");
    }
    paintPdfPage(ctx, page, scale, images);
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, "image/png"));
    if (!blob) {
      throw new Error("the image could not be encoded");
    }
    return new Uint8Array(await blob.arrayBuffer());
  } finally {
    canvas.width = 0;
    canvas.height = 0;
  }
}
