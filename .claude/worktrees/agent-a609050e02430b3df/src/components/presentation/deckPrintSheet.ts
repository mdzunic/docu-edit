// The DOM side of DeckPrintSheet.vue (SPEC-104), typed here because the .vue
// script is espree-linted: how the export reads the sheet's markup — which
// of its elements the PDF extractor treats as images, canvases and rects —
// how a mounted page is found, and when a mounted sheet has settled.
import type { ExtractOptions } from "../../lib/core/pdfExtract";
import { PX_PER_CM } from "../../lib/core/units";
import { SHEET_SETTLE_DEADLINE_MS } from "../../lib/presentation/deckExport";
import type { PresentationModel } from "../../lib/presentation/model";
import {
  NOTE_LINE_COLOR,
  NOTE_LINE_WIDTH_PT,
  NOTES_FONT_SIZE_PT,
  NOTES_LINE_HEIGHT,
  PAGE_NUMBER_SIZE_PT,
  SLIDE_FRAME_COLOR,
  SLIDE_FRAME_WIDTH_PT,
  type SheetPage,
} from "../../lib/presentation/printLayout";
import { stageFields, type StageField } from "./slideRender";

/** What the extractor reads off a slide sheet: slide pictures, media posters
 * and picture bullets; shape and chart canvases; and every slide box whose
 * background or border paints — slide backgrounds, text-frame fills, table
 * cells, foreign-frame boxes, media plates. (The sheet's own handout frames
 * and note lines join as exact rects from the plan — `sheetPageRects`.) */
export const SHEET_EXTRACT_OPTIONS: ExtractOptions = {
  imageSelector: "img.frame-image, img.media-poster, img.bullet-image",
  canvasSelector: "canvas.shape-canvas, canvas.chart-canvas",
  rectSelector: ".stage-surface, .text-frame, .slide-table-cell, .foreign-frame, .media-plate",
};

/** The mounted sheet (teleported to <body>), if any. */
export function sheetElement(): HTMLElement | null {
  return document.querySelector<HTMLElement>(".deck-print-sheet");
}

/** The mounted page planned at `index`, if it is in the window. */
export function sheetPageElement(index: number): HTMLElement | null {
  return document.querySelector<HTMLElement>(`.deck-print-sheet [data-page-index="${index}"]`);
}

/** Waits until a mounted sheet has settled — its images decoded and the
 * document's fonts loaded — or SHEET_SETTLE_DEADLINE_MS passed (the SPEC-034
 * realizeAllPages precedent: a slow image never blocks an export forever). */
export async function settleSheet(root: HTMLElement): Promise<void> {
  const images = Array.from(root.querySelectorAll("img")).map((img) =>
    typeof img.decode === "function" ? img.decode().catch(() => undefined) : undefined,
  );
  const fonts = document.fonts ? document.fonts.ready.then(() => undefined) : undefined;
  let timer = 0;
  const deadline = new Promise<void>((resolve) => {
    timer = window.setTimeout(resolve, SHEET_SETTLE_DEADLINE_MS);
  });
  try {
    await Promise.race([Promise.all([...images, fonts]), deadline]);
  } finally {
    window.clearTimeout(timer);
  }
}

// --- the sheet's view model -------------------------------------------------

type Style = Record<string, string>;

export interface SheetSlotView {
  key: string;
  slideIndex: number;
  /** The slide picture's box on the page. */
  style: Style;
  /** The SlideStage's `scale`: the slot's own in print, 1 in export. */
  stageScale: number;
  fields: StageField[];
}

export interface SheetPageView {
  index: number;
  style: Style;
  slots: SheetSlotView[];
  /** "Frame slides": each outline's box, straddling its slide's edge. */
  frames: { key: string; style: Style }[];
  noteLines: { key: string; style: Style }[];
  notes: { style: Style; paragraphs: string[] } | null;
  pageNumber: { text: string; style: Style } | null;
}

const PX_PER_PT = 96 / 72;
const CM_PER_PT = 2.54 / 72;

/** The sheet is paper, not app chrome: white in every app theme, black ink,
 * and its own text (notes, page numbers) in the bundled sans the PDF embeds.
 * Inline like the slide renderer's model colours, never a UI token. */
export const SHEET_PAPER_STYLE: Style = { background: "#ffffff", color: "#000000" };
const SHEET_TEXT_STYLE: Style = { fontFamily: '"Liberation Sans", Arial, sans-serif', color: "#000000" };

/** Typed pass-throughs for the .vue script (espree bars the casts there). */
export function asSheetPages(value: unknown): SheetPage[] {
  return (value ?? []) as SheetPage[];
}
export function asDeck(value: unknown): PresentationModel {
  return value as PresentationModel;
}

/** Lays each mounted page out in CSS px. In "print" mode a sheet cm is a CSS
 * cm and each slide stage is scaled into its slot. In "export" mode the page
 * is magnified by 1 / its slots' scale so every slide renders at scale 1 with
 * no transform anywhere — `extractPageContent` maps the magnified page back
 * onto the paper size, and reads font sizes that are right only when no
 * transform sits between the text and the page. Every length (font sizes and
 * rule widths included) is multiplied by the same factor, so both modes lay
 * out identically, only at different sizes. */
export function sheetPageViews(
  deck: PresentationModel,
  pages: readonly SheetPage[],
  mode: "print" | "export",
  instant: Date,
): SheetPageView[] {
  return pages.map((page) => {
    const magnify = mode === "export" ? 1 / (page.slots[0]?.scale || 1) : 1;
    const px = (cm: number) => `${cm * PX_PER_CM * magnify}px`;
    const ptPx = (pt: number) => `${pt * PX_PER_PT * magnify}px`;
    const box = (rect: { xCm: number; yCm: number; widthCm: number; heightCm: number }): Style => ({
      left: px(rect.xCm),
      top: px(rect.yCm),
      width: px(rect.widthCm),
      height: px(rect.heightCm),
    });
    const frameCm = SLIDE_FRAME_WIDTH_PT * CM_PER_PT;
    return {
      index: page.index,
      style: { ...SHEET_PAPER_STYLE, width: px(page.widthCm), height: px(page.heightCm) },
      slots: page.slots.map((slot, n) => ({
        key: `${page.index}-${n}-${slot.slideIndex}`,
        slideIndex: slot.slideIndex,
        style: box(slot.slide),
        stageScale: slot.scale * magnify,
        fields: stageFields(deck, slot.slideIndex, instant),
      })),
      frames: page.slots.flatMap((slot, n) =>
        slot.framed
          ? [
              {
                key: `${n}`,
                style: {
                  ...box({
                    xCm: slot.slide.xCm - frameCm / 2,
                    yCm: slot.slide.yCm - frameCm / 2,
                    widthCm: slot.slide.widthCm + frameCm,
                    heightCm: slot.slide.heightCm + frameCm,
                  }),
                  border: `${ptPx(SLIDE_FRAME_WIDTH_PT)} solid ${SLIDE_FRAME_COLOR}`,
                },
              },
            ]
          : [],
      ),
      noteLines: page.slots.flatMap((slot, n) =>
        (slot.noteLines?.yCm ?? []).map((yCm, row) => ({
          key: `${n}-${row}`,
          style: {
            left: px(slot.noteLines?.xCm ?? 0),
            top: px(yCm - (NOTE_LINE_WIDTH_PT * CM_PER_PT) / 2),
            width: px(slot.noteLines?.widthCm ?? 0),
            borderTop: `${ptPx(NOTE_LINE_WIDTH_PT)} solid ${NOTE_LINE_COLOR}`,
          },
        })),
      ),
      notes: page.notes
        ? {
            style: {
              ...box(page.notes.box),
              ...SHEET_TEXT_STYLE,
              fontSize: ptPx(NOTES_FONT_SIZE_PT),
              lineHeight: String(NOTES_LINE_HEIGHT),
            },
            paragraphs: page.notes.paragraphs,
          }
        : null,
      pageNumber:
        page.pageNumber && page.pageNumberBox
          ? { text: page.pageNumber, style: { ...box(page.pageNumberBox), ...SHEET_TEXT_STYLE, fontSize: ptPx(PAGE_NUMBER_SIZE_PT) } }
          : null,
    };
  });
}
