// The framework-free rules of deck print/export (SPEC-104): the limits, the
// PNG pixel size and its clamp, download names, the per-page outline
// headings and the rotation note. useDeckExport wires these to the dialog and
// the off-screen sheet; nothing here touches the DOM or mutates the model.
import { cmToPx } from "../core/units";
import { PT_PER_CM } from "../core/pdfExtract";
import type { PdfHeading, PdfRectItem } from "../core/pdfBuild";
import { DEFAULT_PRESENTATION_TITLE, type PresentationModel, type SlideFrame } from "./model";
import { slideOutline } from "./slideModel";
import {
  NOTE_LINE_COLOR,
  NOTE_LINE_WIDTH_PT,
  SLIDE_FRAME_WIDTH_PT,
  type SheetPage,
} from "./printLayout";

/** Pages mounted at once while exporting: mount → extract → unmount, rolling. */
export const EXPORT_WINDOW_PAGES = 4;
/** Printing mounts every page at once, so a print is capped. */
export const PRINT_PAGE_LIMIT = 250;
/** The longest edge a PNG export may have; larger scales clamp down. */
export const MAX_PNG_EDGE_PX = 8192;
/** The most slides one image export writes. */
export const MAX_PNG_EXPORT_SLIDES = 200;
/** How long a mounted sheet may take to settle (images decoded, fonts
 * loaded) before the export or print goes ahead anyway — SPEC-034's
 * realizeAllPages precedent. */
export const SHEET_SETTLE_DEADLINE_MS = 8000;
/** The PNG scale choices, 1× being the slide's own CSS-pixel size. */
export const PNG_SCALES: readonly number[] = [1, 2, 4];

export const PNG_MIME_TYPE = "image/png";
export const ZIP_MIME_TYPE = "application/zip";

/** A PNG export's pixel size: `widthPx = Math.round(cmToPx(slideWidthCm) *
 * scale)` with the height at the slide's aspect, clamped so the longer edge
 * never passes MAX_PNG_EDGE_PX. A 16:9 deck (33.867 × 19.05 cm) is exactly
 * 1280×720 / 2560×1440 / 5120×2880 at 1×/2×/4×. */
export function pngPixelSize(
  slideWidthCm: number,
  slideHeightCm: number,
  scale: number,
): { widthPx: number; heightPx: number; clamped: boolean } {
  let widthPx = Math.round(cmToPx(slideWidthCm) * scale);
  let heightPx = Math.round(widthPx * (slideHeightCm / slideWidthCm));
  const longer = Math.max(widthPx, heightPx);
  const clamped = longer > MAX_PNG_EDGE_PX;
  if (clamped) {
    const factor = MAX_PNG_EDGE_PX / longer;
    widthPx = Math.min(MAX_PNG_EDGE_PX, Math.round(widthPx * factor));
    heightPx = Math.min(MAX_PNG_EDGE_PX, Math.round(widthPx * (slideHeightCm / slideWidthCm)));
  }
  return { widthPx: Math.max(1, widthPx), heightPx: Math.max(1, heightPx), clamped };
}

/** The image tab's estimate line, naming the clamp when one applied. */
export function pngEstimate(count: number, size: { widthPx: number; heightPx: number; clamped: boolean }): string {
  const images = count === 1 ? "1 image" : `${count} images`;
  const pixels = `${size.widthPx} × ${size.heightPx} px${count === 1 ? "" : " each"}`;
  const clamp = size.clamped ? ` — clamped to ${MAX_PNG_EDGE_PX} px on the longer edge` : "";
  return `${images}, ${pixels}${clamp}`;
}

function fileTitle(title: string): string {
  const trimmed = title.trim();
  return trimmed.length > 0 ? trimmed : DEFAULT_PRESENTATION_TITLE;
}

/** `{title}-slide-{NN}.png`: the 1-based slide number zero-padded to the
 * deck's own digit count, two digits at least. */
export function slidePngFileName(title: string, slideNumber: number, slideCount: number): string {
  const digits = Math.max(2, String(Math.max(1, slideCount)).length);
  return `${fileTitle(title)}-slide-${String(slideNumber).padStart(digits, "0")}.png`;
}

/** `{title}-slides.zip`, the whole range's images. */
export function slidesZipFileName(title: string): string {
  return `${fileTitle(title)}-slides.zip`;
}

/** The refusal message for a print over PRINT_PAGE_LIMIT pages, or "". */
export function printRefusal(pageCount: number): string {
  return pageCount > PRINT_PAGE_LIMIT
    ? `This prints ${pageCount} pages — printing is limited to ${PRINT_PAGE_LIMIT} pages. Choose a smaller range.`
    : "";
}

/** The refusal message for an image export over MAX_PNG_EXPORT_SLIDES, or "". */
export function imagesRefusal(slideCount: number): string {
  return slideCount > MAX_PNG_EXPORT_SLIDES
    ? `Image export is limited to ${MAX_PNG_EXPORT_SLIDES} slides at a time. Choose a smaller range.`
    : "";
}

/** One level-1 outline entry per slide a sheet page places — the slide's
 * title, else "Slide n" — so the PDF gets bookmarks (a notes continuation
 * page places no slide and adds none). */
export function pageHeadings(model: Pick<PresentationModel, "slides">, page: SheetPage): PdfHeading[] {
  return page.slots.map((slot) => {
    const slide = model.slides[slot.slideIndex];
    const title = slide ? slideOutline(slide).title : undefined;
    return { level: 1, title: title ?? `Slide ${slot.slideIndex + 1}`, yPt: slot.slide.yCm * PT_PER_CM };
  });
}

function hasRotatedFrame(frames: readonly SlideFrame[]): boolean {
  return frames.some(
    (frame) => (frame.rotationDeg ?? 0) % 360 !== 0 || (frame.kind === "group" && hasRotatedFrame(frame.children)),
  );
}

/** Whether any of `slides` holds a frame with a rotation — whose text prints
 * unrotated at its axis-aligned box (the dialog's rotation note). */
export function rangeHasRotatedFrame(model: Pick<PresentationModel, "slides">, slides: readonly number[]): boolean {
  return slides.some((index) => hasRotatedFrame(model.slides[index]?.frames ?? []));
}

/** A sheet page's own vector chrome, exactly as the plan lays it out (not
 * read back off the DOM, where border widths snap to device pixels): each
 * framed slot's 0.75 pt black outline — on top, so no slide content covers
 * it — and each ruled note line as a 0.5 pt #808080 bar centred on its row. */
export function sheetPageRects(page: SheetPage): PdfRectItem[] {
  const rects: PdfRectItem[] = [];
  for (const slot of page.slots) {
    if (slot.framed) {
      rects.push({
        xPt: slot.slide.xCm * PT_PER_CM,
        yPt: slot.slide.yCm * PT_PER_CM,
        widthPt: slot.slide.widthCm * PT_PER_CM,
        heightPt: slot.slide.heightCm * PT_PER_CM,
        stroke: [0, 0, 0],
        strokeWidthPt: SLIDE_FRAME_WIDTH_PT,
        onTop: true,
      });
    }
    const lines = slot.noteLines;
    for (const yCm of lines?.yCm ?? []) {
      rects.push({
        xPt: (lines?.xCm ?? 0) * PT_PER_CM,
        yPt: yCm * PT_PER_CM - NOTE_LINE_WIDTH_PT / 2,
        widthPt: (lines?.widthCm ?? 0) * PT_PER_CM,
        heightPt: NOTE_LINE_WIDTH_PT,
        fill: NOTE_LINE_COLOR,
      });
    }
  }
  return rects;
}
