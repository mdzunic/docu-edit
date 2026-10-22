// The image editor's one renderer (SPEC-112): `renderEditedImage` draws a
// source through an edit stack at any pixel size. The on-screen preview
// calls it at screen scale over a downscaled buffer (capped at
// EDIT_PREVIEW_MAX_EDGE_PX), and Save, Save as a copy, Export and Flatten
// call it at full resolution over the full decode — so nothing can be true
// of one and false of the other, the rule SPEC-109 sets for composition.
//
// The stages, in the stack's fixed order:
// 1. geometry — straighten, crop, rotate, flip and resize are ONE drawImage
//    through core/imageEdits' cumulative transform;
// 2. adjust — core/imageAdjust's pixel functions over getImageData, in row
//    bands, never a CSS filter;
// 3. redact — opaque fills, snapped OUTWARD to whole pixels so no partly
//    covered edge pixel keeps a trace of what it hides;
// 4. annotate — the shapes, at the edit's colour, width and opacity.
// Full-resolution work yields to the main thread between stages (and adjust
// bands) so a large image never freezes the page.
//
// Browser-bound (canvas, image decoding), the same exception
// core/svgToPng.ts takes; no Vue. The canvas factory is injectable so the
// stage arithmetic is unit-testable against a recording context.
import {
  normaliseStack,
  resolveEditGeometry,
  multiplyMatrix,
  scaleMatrix,
  type ImageEdit,
  type PixelSize,
} from "../core/imageEdits";
import { applyAdjustments } from "../core/imageAdjust";
import { yieldToMainThread } from "../core/idleScheduler";
import { convertSvgToPng, type ConvertedPng } from "../core/svgToPng";
import { triggerBlobDownload } from "../core/triggerDownload";
import { canvasCeilingFor, type CanvasFactory } from "./canvasCeiling";

/** The preview buffer's longest edge: a 48-megapixel photo previews from a
 * 2048-pixel copy, never from the full decode. */
export const EDIT_PREVIEW_MAX_EDGE_PX = 2048;

/** The rows one adjust band covers at most this many pixels. */
const ADJUST_BAND_PIXELS = 1 << 20;

/** What the renderer returns — a canvas, drawable and encodable. */
export type ImageBitmapLike = HTMLCanvasElement;

/** A drawable plus the natural size the stack's coordinates are in: the
 * drawable may be a smaller preview buffer of that image. */
export interface EditSource {
  image: CanvasImageSource;
  naturalPx: PixelSize;
}

export type RenderStage = "geometry" | "adjust" | "redact" | "annotate";

export interface RenderOptions {
  createCanvas?: CanvasFactory;
  /** Full-resolution work: yield between stages and adjust bands. */
  chunked?: boolean;
  onStage?: (stage: RenderStage) => void;
}

const browserCanvas: CanvasFactory = (widthPx, heightPx) => {
  const canvas = document.createElement("canvas");
  canvas.width = widthPx;
  canvas.height = heightPx;
  return canvas;
};

/** The preview buffer's size: capped at `maxEdgePx` on the longer edge,
 * never upscaled, aspect preserved, at least 1 × 1. */
export function previewBufferSize(naturalPx: PixelSize, maxEdgePx = EDIT_PREVIEW_MAX_EDGE_PX): PixelSize {
  const width = Math.max(1, Math.round(naturalPx.widthPx));
  const height = Math.max(1, Math.round(naturalPx.heightPx));
  const longest = Math.max(width, height);
  if (longest <= maxEdgePx) {
    return { widthPx: width, heightPx: height };
  }
  const scale = maxEdgePx / longest;
  return { widthPx: Math.max(1, Math.round(width * scale)), heightPx: Math.max(1, Math.round(height * scale)) };
}

/** The size the preview renders an `outputPx` result at: fitted into the
 * stage box (aspect preserved, floored so it never overhangs) and capped at
 * `maxEdgePx`. */
export function previewTargetSize(outputPx: PixelSize, boxPx: PixelSize, maxEdgePx = EDIT_PREVIEW_MAX_EDGE_PX): PixelSize {
  const width = Math.max(1, outputPx.widthPx);
  const height = Math.max(1, outputPx.heightPx);
  const boxWidth = Math.max(1, Math.min(boxPx.widthPx, maxEdgePx));
  const boxHeight = Math.max(1, Math.min(boxPx.heightPx, maxEdgePx));
  const scale = Math.min(boxWidth / width, boxHeight / height);
  return { widthPx: Math.max(1, Math.floor(width * scale)), heightPx: Math.max(1, Math.floor(height * scale)) };
}

async function stage(options: RenderOptions, name: RenderStage): Promise<void> {
  if (options.chunked) {
    await yieldToMainThread();
  }
  options.onStage?.(name);
}

function context(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    throw new Error("This browser cannot edit images.");
  }
  return ctx;
}

/** Draws one annotation scaled from output pixels by (sx, sy). */
export function drawAnnotation(
  ctx: CanvasRenderingContext2D,
  edit: Extract<ImageEdit, { kind: "annotate" }>,
  sx: number,
  sy: number,
): void {
  const scale = (sx + sy) / 2;
  const shape = edit.shape;
  ctx.save();
  ctx.globalAlpha = edit.opacity;
  ctx.strokeStyle = edit.color;
  ctx.fillStyle = edit.color;
  ctx.lineWidth = Math.max(0.5, edit.strokeWidth * scale);
  ctx.lineCap = "round";
  ctx.lineJoin = "round";
  switch (shape.shape) {
    case "line":
    case "arrow": {
      const x1 = shape.x1 * sx;
      const y1 = shape.y1 * sy;
      const x2 = shape.x2 * sx;
      const y2 = shape.y2 * sy;
      ctx.beginPath();
      ctx.moveTo(x1, y1);
      ctx.lineTo(x2, y2);
      ctx.stroke();
      if (shape.shape === "arrow") {
        const angle = Math.atan2(y2 - y1, x2 - x1);
        const head = Math.max(8, edit.strokeWidth * 4) * scale;
        ctx.beginPath();
        ctx.moveTo(x2, y2);
        ctx.lineTo(x2 - head * Math.cos(angle - Math.PI / 6), y2 - head * Math.sin(angle - Math.PI / 6));
        ctx.lineTo(x2 - head * Math.cos(angle + Math.PI / 6), y2 - head * Math.sin(angle + Math.PI / 6));
        ctx.closePath();
        ctx.fill();
      }
      break;
    }
    case "rectangle":
      ctx.strokeRect(shape.x * sx, shape.y * sy, shape.width * sx, shape.height * sy);
      break;
    case "ellipse":
      ctx.beginPath();
      ctx.ellipse(
        (shape.x + shape.width / 2) * sx,
        (shape.y + shape.height / 2) * sy,
        (shape.width / 2) * sx,
        (shape.height / 2) * sy,
        0,
        0,
        Math.PI * 2,
      );
      ctx.stroke();
      break;
    case "highlight":
      // A marker, not a cover: multiplied over the picture so what is under
      // it stays legible.
      ctx.globalCompositeOperation = "multiply";
      ctx.fillRect(shape.x * sx, shape.y * sy, shape.width * sx, shape.height * sy);
      break;
    case "freehand": {
      ctx.beginPath();
      shape.points.forEach((p, i) => (i === 0 ? ctx.moveTo(p.x * sx, p.y * sy) : ctx.lineTo(p.x * sx, p.y * sy)));
      if (shape.points.length === 1) {
        ctx.lineTo(shape.points[0].x * sx + 0.01, shape.points[0].y * sy);
      }
      ctx.stroke();
      break;
    }
    case "text":
      ctx.font = `${shape.fontSize * scale}px sans-serif`;
      ctx.textBaseline = "top";
      ctx.fillText(shape.text, shape.x * sx, shape.y * sy);
      break;
    default:
      break;
  }
  ctx.restore();
}

/** Renders `source` through `edits` at `targetPx` (the output's own size at
 * full resolution, a fitted size for the preview). Never touches the
 * source; returns a fresh canvas the caller encodes, draws or releases. */
export async function renderEditedImage(
  source: EditSource,
  edits: readonly ImageEdit[],
  targetPx: PixelSize,
  options: RenderOptions = {},
): Promise<ImageBitmapLike> {
  const stack = normaliseStack(edits);
  const geometry = resolveEditGeometry(stack, source.naturalPx);
  const width = Math.max(1, Math.round(targetPx.widthPx));
  const height = Math.max(1, Math.round(targetPx.heightPx));
  const canvas = (options.createCanvas ?? browserCanvas)(width, height);
  const ctx = context(canvas);
  const sx = width / geometry.widthPx;
  const sy = height / geometry.heightPx;

  await stage(options, "geometry");
  const resize = stack.find((edit) => edit.kind === "resize");
  const m = multiplyMatrix(scaleMatrix(sx, sy), geometry.transform);
  ctx.save();
  ctx.imageSmoothingEnabled = !(resize && resize.resample === "nearest");
  ctx.imageSmoothingQuality = "high";
  ctx.setTransform(m.a, m.b, m.c, m.d, m.e, m.f);
  ctx.drawImage(source.image, 0, 0, source.naturalPx.widthPx, source.naturalPx.heightPx);
  ctx.restore();
  ctx.setTransform(1, 0, 0, 1, 0, 0);

  const adjust = stack.find((edit) => edit.kind === "adjust");
  if (adjust) {
    await stage(options, "adjust");
    const band = Math.max(1, Math.floor(ADJUST_BAND_PIXELS / width));
    for (let y = 0; y < height; y += band) {
      const rows = Math.min(band, height - y);
      const image = ctx.getImageData(0, y, width, rows);
      applyAdjustments(image.data, adjust, image.data);
      ctx.putImageData(image, 0, y);
      if (options.chunked && y + rows < height) {
        await yieldToMainThread();
      }
    }
  }

  const redactions = stack.filter((edit) => edit.kind === "redact");
  if (redactions.length > 0) {
    await stage(options, "redact");
    for (const edit of redactions) {
      const x0 = Math.max(0, Math.floor(edit.x * sx));
      const y0 = Math.max(0, Math.floor(edit.y * sy));
      const x1 = Math.min(width, Math.ceil((edit.x + edit.width) * sx));
      const y1 = Math.min(height, Math.ceil((edit.y + edit.height) * sy));
      if (x1 > x0 && y1 > y0) {
        ctx.fillStyle = edit.color;
        ctx.fillRect(x0, y0, x1 - x0, y1 - y0);
      }
    }
  }

  const annotations = stack.filter((edit) => edit.kind === "annotate");
  if (annotations.length > 0) {
    await stage(options, "annotate");
    for (const edit of annotations) {
      drawAnnotation(ctx, edit, sx, sy);
    }
  }
  return canvas;
}

// --- decoding, encoding and the editor's browser steps ----------------------

/** A decoded raster and its natural size. */
export interface DecodedImage {
  image: CanvasImageSource;
  widthPx: number;
  heightPx: number;
  release: () => void;
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.addEventListener("load", () => resolve(img));
    img.addEventListener("error", () => reject(new Error("The image could not be read.")));
    img.src = url;
  });
}

/** Decodes raster bytes (an animated GIF decodes as its first frame). */
export async function decodeImageBytes(bytes: Uint8Array, contentType: string): Promise<DecodedImage> {
  const blob = new Blob([bytes as unknown as BlobPart], { type: contentType });
  if (typeof createImageBitmap === "function") {
    try {
      const bitmap = await createImageBitmap(blob);
      return { image: bitmap, widthPx: bitmap.width, heightPx: bitmap.height, release: () => bitmap.close() };
    } catch {
      // Fall through to an <img>, which some engines decode where
      // createImageBitmap refuses.
    }
  }
  const url = URL.createObjectURL(blob);
  try {
    const img = await loadImage(url);
    return { image: img, widthPx: img.naturalWidth, heightPx: img.naturalHeight, release: () => {} };
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** A decode cache keyed by media id: one decode per part per session, however
 * many previews and exports read it. A failed decode is forgotten so a
 * retry decodes again. */
export function createDecodeCache(decode: (bytes: Uint8Array, contentType: string) => Promise<DecodedImage>) {
  const entries = new Map<string, Promise<DecodedImage>>();
  return {
    get(mediaId: string, bytes: () => Uint8Array, contentType: string): Promise<DecodedImage> {
      let entry = entries.get(mediaId);
      if (!entry) {
        entry = decode(bytes(), contentType);
        entries.set(mediaId, entry);
        entry.catch(() => entries.delete(mediaId));
      }
      return entry;
    },
    clear(): void {
      for (const entry of entries.values()) {
        entry.then((decoded) => decoded.release()).catch(() => {});
      }
      entries.clear();
    },
  };
}

/** The preview's source: the decode itself when it is within the cap, else
 * a downscaled copy — either way addressed in the natural size's pixels. */
export function previewSource(
  decoded: DecodedImage,
  maxEdgePx = EDIT_PREVIEW_MAX_EDGE_PX,
  createCanvas: CanvasFactory = browserCanvas,
): EditSource {
  const naturalPx = { widthPx: decoded.widthPx, heightPx: decoded.heightPx };
  const size = previewBufferSize(naturalPx, maxEdgePx);
  if (size.widthPx === naturalPx.widthPx && size.heightPx === naturalPx.heightPx) {
    return { image: decoded.image, naturalPx };
  }
  const canvas = createCanvas(size.widthPx, size.heightPx);
  const ctx = context(canvas);
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = "high";
  ctx.drawImage(decoded.image, 0, 0, size.widthPx, size.heightPx);
  return { image: canvas, naturalPx };
}

export interface EncodeOptions {
  /** JPEG quality, 0..1. */
  quality?: number;
  /** "#rrggbb" painted behind a JPEG, which cannot hold alpha. */
  background?: string;
}

/** Encodes a rendered canvas as PNG or JPEG bytes. */
export async function encodeRender(
  image: ImageBitmapLike,
  contentType: string,
  options: EncodeOptions = {},
  createCanvas: CanvasFactory = browserCanvas,
): Promise<Uint8Array> {
  let canvas = image;
  if (contentType === "image/jpeg") {
    canvas = createCanvas(image.width, image.height);
    const ctx = context(canvas);
    ctx.fillStyle = options.background ?? "#ffffff";
    ctx.fillRect(0, 0, image.width, image.height);
    ctx.drawImage(image, 0, 0);
  }
  try {
    const blob = await new Promise<Blob>((resolve, reject) => {
      canvas.toBlob(
        (result) => (result ? resolve(result) : reject(new Error("The image could not be encoded."))),
        contentType,
        options.quality,
      );
    });
    return new Uint8Array(await blob.arrayBuffer());
  } finally {
    if (canvas !== image) {
      releaseCanvas(canvas);
    }
  }
}

/** Frees a canvas's backing store now rather than at the next GC. */
export function releaseCanvas(canvas: { width: number; height: number }): void {
  canvas.width = 0;
  canvas.height = 0;
}

/** The browser-bound steps the editor runs — injectable so the session's
 * logic is unit-testable without a rendering engine. */
export interface ImageEditSteps {
  decode: (bytes: Uint8Array, contentType: string) => Promise<DecodedImage>;
  preview: (decoded: DecodedImage) => EditSource;
  render: (source: EditSource, edits: readonly ImageEdit[], targetPx: PixelSize, options?: RenderOptions) => Promise<ImageBitmapLike>;
  encode: (image: ImageBitmapLike, contentType: string, options?: EncodeOptions) => Promise<Uint8Array>;
  rasteriseSvg: (bytes: Uint8Array, size?: PixelSize) => Promise<ConvertedPng>;
  /** The canvas ceiling an area is held against (probed only when needed). */
  canvasCeiling: (areaPx: number) => number;
  download: (bytes: Uint8Array, fileName: string, contentType: string) => void;
  release: (image: ImageBitmapLike) => void;
}

export const BROWSER_EDIT_STEPS: ImageEditSteps = {
  decode: decodeImageBytes,
  preview: (decoded) => previewSource(decoded),
  render: renderEditedImage,
  encode: (image, contentType, options) => encodeRender(image, contentType, options),
  rasteriseSvg: convertSvgToPng,
  canvasCeiling: (areaPx) => canvasCeilingFor(areaPx),
  download: triggerBlobDownload,
  release: releaseCanvas,
};
