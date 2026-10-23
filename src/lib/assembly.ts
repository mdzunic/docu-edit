// The assembly workspace's output and tray rules (SPEC-109) — the
// framework-free half of /assemble's composables: how big the output is and
// whether this browser can make it (MAX_OUTPUT_EDGE_PX on either axis, the
// measured canvas ceiling on the area — with the dpi that WOULD fit, so the
// fix is one click), the live estimate, the colour JPEG paints where the
// sheet asked for transparency, the tray's spoken reorder announcements and
// the restored-tray notice. The geometry is core/sheetLayout's and the
// layout core/assemblyLayout's; nothing here imports Vue.
import { SHEET_DPI_CHOICES, sheetPixelSize, TRANSPARENT, type SheetGeometry } from "./core/sheetLayout";
import { megapixelsLabel } from "./media/canvasCeiling";
import { mediaFileExtension } from "./media/model";
import { formatFileSize } from "./format";
import { DEFAULT_ASSEMBLY_TITLE, type AssemblyFormat } from "./assemblyPersist";

/** Either side of the output, independently, is at most this many pixels. */
export const MAX_OUTPUT_EDGE_PX = 16384;

export const ASSEMBLY_CONTENT_TYPES: Record<AssemblyFormat, string> = {
  png: "image/png",
  jpeg: "image/jpeg",
};

export const ASSEMBLY_FORMAT_LABELS: Record<AssemblyFormat, string> = {
  png: "PNG",
  jpeg: "JPEG",
};

export interface OutputPixels {
  widthPx: number;
  heightPx: number;
}

/** Whether `pixels` fit both bounds: each side, and the area. */
export function outputFits(pixels: OutputPixels, ceilingPx: number): boolean {
  return (
    pixels.widthPx > 0 &&
    pixels.heightPx > 0 &&
    pixels.widthPx <= MAX_OUTPUT_EDGE_PX &&
    pixels.heightPx <= MAX_OUTPUT_EDGE_PX &&
    pixels.widthPx * pixels.heightPx <= ceilingPx
  );
}

/** The highest SHEET_DPI_CHOICES step at which `geometry` fits, or null
 * when not even the lowest does. */
export function fittingDpi(geometry: SheetGeometry, ceilingPx: number): number | null {
  const fitting = SHEET_DPI_CHOICES.filter((dpi) => outputFits(sheetPixelSize(geometry, dpi), ceilingPx));
  return fitting.length > 0 ? Math.max(...fitting) : null;
}

/** Why the output cannot be made — naming the limit it breaks and the dpi
 * that would fit — or null when it can. Checked before any decoding. */
export function outputRefusal(geometry: SheetGeometry, dpi: number, ceilingPx: number): string | null {
  const pixels = sheetPixelSize(geometry, dpi);
  if (outputFits(pixels, ceilingPx)) {
    return null;
  }
  const size = `${pixels.widthPx} × ${pixels.heightPx} px`;
  const fits = fittingDpi(geometry, ceilingPx);
  const fix =
    fits !== null
      ? `At ${fits} dpi it would fit.`
      : `It would not fit even at ${Math.min(...SHEET_DPI_CHOICES)} dpi; choose a smaller sheet.`;
  if (pixels.widthPx > MAX_OUTPUT_EDGE_PX || pixels.heightPx > MAX_OUTPUT_EDGE_PX) {
    return `At ${dpi} dpi the image would be ${size}; either side can be at most ${MAX_OUTPUT_EDGE_PX} px. ${fix}`;
  }
  const area = pixels.widthPx * pixels.heightPx;
  return (
    `At ${dpi} dpi the image would be ${size} (${megapixelsLabel(area)}), more than this browser's ` +
    `canvas limit of ${megapixelsLabel(ceilingPx)}. ${fix}`
  );
}

/** Approximate output bytes. PNG: never more than four bytes a pixel plus
 * the file's overhead (the bound SPEC-108's estimate states). JPEG: a
 * typical photograph's bytes a pixel at that quality, from 0.05 at 1 to
 * 0.65 at 100. */
export function estimateOutputBytes(pixels: OutputPixels, format: AssemblyFormat, quality: number): number {
  const area = Math.max(0, pixels.widthPx) * Math.max(0, pixels.heightPx);
  if (format === "png") {
    return area * 4 + 2048;
  }
  const q = Math.min(100, Math.max(1, quality)) / 100;
  return Math.round(area * (0.05 + 0.6 * q * q)) + 2048;
}

/** "1754 × 1240 px at 150 dpi · PNG up to about 8.3 MB", with the measured
 * canvas limit appended once it has been probed. */
export function outputEstimate(
  pixels: OutputPixels,
  dpi: number,
  format: AssemblyFormat,
  quality: number,
  measuredCeilingPx: number | null,
): string {
  const bytes = formatFileSize(estimateOutputBytes(pixels, format, quality));
  const label = format === "png" ? `PNG up to about ${bytes}` : `JPEG about ${bytes}`;
  const limit = measuredCeilingPx !== null ? ` · canvas limit ${megapixelsLabel(measuredCeilingPx)}` : "";
  return `${pixels.widthPx} × ${pixels.heightPx} px at ${dpi} dpi · ${label}${limit}`;
}

/** The colour the output paints behind everything: the sheet's own, except
 * that JPEG has no alpha — a "transparent" sheet falls back to `matte`. */
export function outputBackground(sheetBackground: string, format: AssemblyFormat, matte: string): string {
  return format === "jpeg" && sheetBackground === TRANSPARENT ? matte : sheetBackground;
}

/** The notice shown before the user commits when JPEG cannot honour the
 * sheet's transparency ("" when there is nothing to say). */
export function jpegBackgroundNotice(sheetBackground: string, format: AssemblyFormat): string {
  return format === "jpeg" && sheetBackground === TRANSPARENT
    ? "JPEG has no transparency, so the transparent background will be filled with the colour chosen here."
    : "";
}

/** The file a download is saved as: "<title>.png" / "<title>.jpeg". */
export function outputFileName(title: string, format: AssemblyFormat): string {
  const trimmed = title.trim();
  return `${trimmed.length > 0 ? trimmed : DEFAULT_ASSEMBLY_TITLE}.${mediaFileExtension(format)}`;
}

/** `list` with the entry at `from` moved to `to` (both clamped). */
export function moveEntry<T>(list: readonly T[], from: number, to: number): T[] {
  const next = [...list];
  if (from < 0 || from >= next.length) {
    return next;
  }
  const [entry] = next.splice(from, 1);
  next.splice(Math.min(Math.max(0, to), next.length), 0, entry);
  return next;
}

/** The live region's words for a reorder: "“shot” moved to position 2 of 3." */
export function trayMoveAnnouncement(title: string, position: number, total: number): string {
  return `“${title}” moved to position ${position} of ${total}.`;
}

/** The one-line notice for items a restored tray dropped. */
export function droppedItemsNotice(titles: readonly string[]): string {
  if (titles.length === 0) {
    return "";
  }
  const named = titles.map((title) => `“${title || "Untitled"}”`).join(", ");
  return titles.length === 1
    ? `${named} is no longer in the library and was removed from the tray.`
    : `${named} are no longer in the library and were removed from the tray.`;
}
