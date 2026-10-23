// The assembly renderer (SPEC-109): ONE function draws a planned assembly
// onto a canvas, and both the workspace's on-screen preview and the exported
// PNG/JPEG call it — the preview at its fit-to-window size, the export at
// SPEC-108's sheetPixelSize — so there is no second drawing path that could
// disagree with what the user saw. It draws the sheet background, then per
// placed item: save, set globalAlpha and globalCompositeOperation, fill the
// cell background, drawImage(source, sourcePx…, destCm→px…), restore.
// Drawing is chunked by the layout's rows (item by item for overlay), with an
// awaited `between` hook and a cancel check at every chunk boundary, so a
// large export never holds the main thread and stops cleanly.
//
// Sources are decoded through one session cache keyed by the part's media id
// (decodeSource) — a duplicated tray item, a re-render and the export all
// read the same decode. An SVG part is rasterised (core/svgToPng) at the
// pixel size it is drawn at, so a vector logo is sharp at 600 dpi.
//
// Browser-bound (canvas + image decoding), the same exception
// core/svgToPng.ts and media/renderEdits.ts take; no Vue, no store.
import type { BlendMode, PlacedItem } from "../core/assemblyLayout";
import type { PixelSize } from "../core/imageEdits";
import type { MediaPart } from "../core/media";
import { TRANSPARENT, type RectCm, type SheetGeometry } from "../core/sheetLayout";
import { convertSvgToPng, type ConvertedPng } from "../core/svgToPng";
import { mediaPartBytes } from "./download";
import { decodeImageBytes, type DecodedImage } from "./renderEdits";

/** The sheet as the renderer needs it: its geometry and the background
 * colour actually painted ("#rrggbb", or "transparent" for none). */
export interface ComposeSheet {
  geometry: SheetGeometry;
  background: string;
}

/** A decoded source for one placed item: what drawImage reads, and how many
 * of its pixels make one pixel of the item's nominal size (1 for a raster;
 * more for an SVG rasterised at a large placed size). */
export interface ComposeSource {
  image: CanvasImageSource;
  scaleX: number;
  scaleY: number;
}

/** Sources by PlacedItem.index; an index without one draws its cell only. */
export type ComposeSources = ReadonlyMap<number, ComposeSource>;

export interface ComposeOptions {
  /** Awaited between drawing chunks — yieldToMainThread for the export. */
  between?: () => Promise<void>;
  /** Checked at every chunk boundary; true stops the drawing there. */
  cancelled?: () => boolean;
  /** Called with the items drawn so far, after each chunk. */
  onProgress?: (done: number, total: number) => void;
}

const COMPOSITE: Record<BlendMode, GlobalCompositeOperation> = {
  normal: "source-over",
  multiply: "multiply",
  screen: "screen",
  overlay: "overlay",
};

/** A centimetre rectangle in whole canvas pixels at `sx` × `sy` px per cm.
 * Each edge is rounded on its own, so neighbouring cells meet exactly. */
export function toPixelRect(rect: RectCm, sx: number, sy: number): { x: number; y: number; width: number; height: number } {
  const x = Math.round(rect.xCm * sx);
  const y = Math.round(rect.yCm * sy);
  return {
    x,
    y,
    width: Math.max(0, Math.round((rect.xCm + rect.widthCm) * sx) - x),
    height: Math.max(0, Math.round((rect.yCm + rect.heightCm) * sy) - y),
  };
}

/** The drawing chunks: placed items grouped by their layout row, rows in
 * order, tray order within a row. */
export function drawingChunks(placed: readonly PlacedItem[]): PlacedItem[][] {
  const byRow = new Map<number, PlacedItem[]>();
  for (const item of [...placed].sort((a, b) => a.row - b.row || a.index - b.index)) {
    const chunk = byRow.get(item.row) ?? [];
    chunk.push(item);
    byRow.set(item.row, chunk);
  }
  return [...byRow.values()];
}

function drawPlaced(ctx: CanvasRenderingContext2D, item: PlacedItem, source: ComposeSource | undefined, sx: number, sy: number): void {
  ctx.save();
  ctx.globalAlpha = Math.min(1, Math.max(0, item.opacity));
  ctx.globalCompositeOperation = COMPOSITE[item.blend] ?? "source-over";
  if (item.cellBackground !== TRANSPARENT) {
    const cell = toPixelRect(item.cell, sx, sy);
    ctx.fillStyle = item.cellBackground;
    ctx.fillRect(cell.x, cell.y, cell.width, cell.height);
  }
  const dest = toPixelRect(item, sx, sy);
  const crop = item.sourcePx;
  if (source && dest.width > 0 && dest.height > 0 && crop.width > 0 && crop.height > 0) {
    ctx.drawImage(
      source.image,
      crop.x * source.scaleX,
      crop.y * source.scaleY,
      crop.width * source.scaleX,
      crop.height * source.scaleY,
      dest.x,
      dest.y,
      dest.width,
      dest.height,
    );
  }
  ctx.restore();
}

/** Draws `placed` on `canvas` at the canvas's own pixel size: the whole
 * sheet maps onto the whole canvas. Resolves true when every item was
 * drawn; false when `cancelled` stopped it at a chunk boundary, or the
 * canvas has no 2-D context (nothing is drawn then). */
export async function composeToCanvas(
  canvas: HTMLCanvasElement,
  placed: readonly PlacedItem[],
  sheet: ComposeSheet,
  sources: ComposeSources,
  options: ComposeOptions = {},
): Promise<boolean> {
  const ctx = canvas.getContext("2d");
  if (!ctx || sheet.geometry.widthCm <= 0 || sheet.geometry.heightCm <= 0) {
    return false;
  }
  const sx = canvas.width / sheet.geometry.widthCm;
  const sy = canvas.height / sheet.geometry.heightCm;
  ctx.save();
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.globalAlpha = 1;
  ctx.globalCompositeOperation = "source-over";
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = "high";
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  if (sheet.background !== TRANSPARENT) {
    ctx.fillStyle = sheet.background;
    ctx.fillRect(0, 0, canvas.width, canvas.height);
  }
  ctx.restore();
  const total = placed.length;
  let done = 0;
  options.onProgress?.(0, total);
  for (const [i, chunk] of drawingChunks(placed).entries()) {
    if (i > 0 && options.between) {
      await options.between();
    }
    if (options.cancelled?.()) {
      return false;
    }
    for (const item of chunk) {
      drawPlaced(ctx, item, sources.get(item.index), sx, sy);
      done += 1;
    }
    options.onProgress?.(done, total);
  }
  return true;
}

// --- sources ------------------------------------------------------------------

export const SVG_CONTENT_TYPE = "image/svg+xml";

/** The browser-bound decoding steps, injectable for tests. */
export interface SourceSteps {
  decode: (bytes: Uint8Array, contentType: string) => Promise<DecodedImage>;
  /** Rasterises at `size`, or at the SVG's own declared size without one. */
  rasteriseSvg: (bytes: Uint8Array, size?: PixelSize) => Promise<ConvertedPng>;
}

export const BROWSER_SOURCE_STEPS: SourceSteps = {
  decode: decodeImageBytes,
  rasteriseSvg: convertSvgToPng,
};

/** What a part is cached under: its media id, and for an SVG the pixel
 * size it is rasterised at. */
export function sourceKey(part: MediaPart, placedPx?: PixelSize): string {
  return part.contentType === SVG_CONTENT_TYPE && placedPx
    ? `${part.mediaId}@${placedPx.widthPx}x${placedPx.heightPx}`
    : part.mediaId;
}

export interface SourceCache {
  /** The decoded part — decoded once, then served from the cache. */
  decodeSource: (part: MediaPart, placedPx?: PixelSize) => Promise<DecodedImage>;
  /** How many decodes the cache holds. */
  size: () => number;
  /** Releases every decode. */
  clear: () => void;
}

/** A decode cache: one decode per key; a failed decode is forgotten so a
 * retry decodes again. Reading a part never changes it. */
export function createSourceCache(steps: SourceSteps = BROWSER_SOURCE_STEPS): SourceCache {
  const entries = new Map<string, Promise<DecodedImage>>();
  return {
    decodeSource(part, placedPx) {
      const key = sourceKey(part, placedPx);
      let entry = entries.get(key);
      if (!entry) {
        const bytes = mediaPartBytes(part);
        entry =
          part.contentType === SVG_CONTENT_TYPE
            ? steps
                .rasteriseSvg(bytes, placedPx)
                .then((png) => steps.decode(png.data, "image/png"))
            : steps.decode(bytes, part.contentType);
        entries.set(key, entry);
        entry.catch(() => entries.delete(key));
      }
      return entry;
    },
    size: () => entries.size,
    clear() {
      for (const entry of entries.values()) {
        entry.then((decoded) => decoded.release()).catch(() => {});
      }
      entries.clear();
    },
  };
}

/** The session's one source cache. */
const sessionSources = createSourceCache();

/** Decodes a media part through the session cache (an SVG at `placedPx`). */
export function decodeSource(part: MediaPart, placedPx?: PixelSize): Promise<DecodedImage> {
  return sessionSources.decodeSource(part, placedPx);
}

/** One tray item as the renderer reads it: the part holding its current
 * image, and the pixel size its PlacedItem.sourcePx is measured in. */
export interface ComposeInput {
  part: MediaPart;
  widthPx: number;
  heightPx: number;
}

export interface PreparedSources {
  sources: Map<number, ComposeSource>;
  /** Indexes whose source could not be decoded. */
  failed: number[];
  /** True when `cancelled` stopped the preparation. */
  cancelled: boolean;
}

/** Decodes the source of every placed item (`inputs` by tray index) for a
 * canvas at `sx` × `sy` px per cm. An SVG is rasterised at the whole pixel
 * size its item is drawn at — the visible rectangle grown back to the
 * item's full extent when a cropping fit cut part of it away. */
export async function prepareSources(
  placed: readonly PlacedItem[],
  inputs: readonly (ComposeInput | null)[],
  sx: number,
  sy: number,
  decode: (part: MediaPart, placedPx?: PixelSize) => Promise<DecodedImage> = decodeSource,
  options: Pick<ComposeOptions, "between" | "cancelled"> = {},
): Promise<PreparedSources> {
  const sources = new Map<number, ComposeSource>();
  const failed: number[] = [];
  for (const [i, item] of placed.entries()) {
    const input = inputs[item.index];
    if (!input || item.sourcePx.width <= 0 || item.sourcePx.height <= 0) {
      continue;
    }
    if (i > 0 && options.between) {
      await options.between();
    }
    if (options.cancelled?.()) {
      return { sources, failed, cancelled: true };
    }
    const dest = toPixelRect(item, sx, sy);
    const placedPx =
      input.part.contentType === SVG_CONTENT_TYPE
        ? {
            widthPx: Math.max(1, Math.round((dest.width * input.widthPx) / item.sourcePx.width)),
            heightPx: Math.max(1, Math.round((dest.height * input.heightPx) / item.sourcePx.height)),
          }
        : undefined;
    try {
      const decoded = await decode(input.part, placedPx);
      sources.set(item.index, {
        image: decoded.image,
        scaleX: input.widthPx > 0 ? decoded.widthPx / input.widthPx : 1,
        scaleY: input.heightPx > 0 ? decoded.heightPx / input.heightPx : 1,
      });
    } catch {
      failed.push(item.index);
    }
  }
  return { sources, failed, cancelled: false };
}

/** Encodes a composed canvas as PNG or JPEG bytes (quality 0..1, JPEG only). */
export async function encodeCanvas(canvas: HTMLCanvasElement, contentType: string, quality?: number): Promise<Uint8Array> {
  const blob = await new Promise<Blob>((resolve, reject) => {
    canvas.toBlob(
      (result) => (result ? resolve(result) : reject(new Error("The image could not be encoded."))),
      contentType,
      quality,
    );
  });
  return new Uint8Array(await blob.arrayBuffer());
}
