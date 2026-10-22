// Deck print & export (SPEC-104) — the one wiring call PresentationView gains.
// Owns the print dialog's choices (persisted for the session here, not in
// lib/preferences), the off-screen export sheet's mounted window, and the
// three runs: browser print (every planned page mounted, `window.print()`
// under an injected @page rule), PDF (each page mounted → settled →
// extractPageContent → unmounted, then fonts/images loaded exactly as the
// word processor's PDF export loads them and buildPdf), and PNG (the same
// pages at the Slides layout, each rasterised with core/pdfRaster). The
// render never touches the editor stage, and nothing here mutates the model:
// the sheet reads it, the planners in lib/presentation are pure.
import { computed, nextTick, ref, watch, type Ref } from "vue";
import { DEFAULT_PRESENTATION_TITLE, type PresentationModel } from "../../lib/presentation/model";
import { buildPdf, PDF_MIME_TYPE, pdfFileName, type PdfPageInput } from "../../lib/core/pdfBuild";
import { extractPageContent, PT_PER_CM } from "../../lib/core/pdfExtract";
import { rasterisePdfPage } from "../../lib/core/pdfRaster";
import { zip } from "../../lib/core/zipWriter";
import { triggerBlobDownload } from "../../lib/core/triggerDownload";
import { yieldToMainThread } from "../../lib/core/idleScheduler";
import { loadPdfFonts, loadPdfImages, registerRasterFaces } from "../../lib/pdfAssets";
import {
  DEFAULT_DECK_PRINT_OPTIONS,
  defaultFrameSlides,
  layoutInfo,
  noteLinesApply,
  planSheet,
  resolveSlideRange,
  sheetPageSize,
  type DeckPrintLayout,
  type DeckPrintOptions,
  type DeckPrintPaper,
  type DeckPrintRange,
  type SheetPage,
} from "../../lib/presentation/printLayout";
import {
  EXPORT_WINDOW_PAGES,
  PNG_MIME_TYPE,
  ZIP_MIME_TYPE,
  imagesRefusal,
  pageHeadings,
  pngEstimate,
  pngPixelSize,
  printRefusal,
  rangeHasRotatedFrame,
  sheetPageRects,
  slidePngFileName,
  slidesZipFileName,
} from "../../lib/presentation/deckExport";
import type { Orientation } from "../../lib/core/pageSetup";
import { SHEET_EXTRACT_OPTIONS, settleSheet, sheetElement, sheetPageElement } from "../../components/presentation/deckPrintSheet";

export type DeckExportTab = "pdf" | "images";
/** "print": real-size pages with the slides scaled into their slots;
 * "export": each page magnified so its slides render at scale 1 with no
 * transform — the extractor reads computed font sizes, which a transform
 * would leave unscaled. */
export type DeckSheetMode = "print" | "export";

export const DECK_EXPORT_ERROR_PREFIX = "Couldn't export the deck";

export interface UseDeckExportDeps {
  model: Ref<PresentationModel | null>;
  currentIndex: Ref<number>;
  slideTotal: Ref<number>;
  reportError: (message: string) => void;
}

interface Choices {
  tab: DeckExportTab;
  layout: DeckPrintLayout;
  paper: DeckPrintPaper;
  orientation: Orientation;
  range: DeckPrintRange;
  rangeText: string;
  noteLines: boolean;
  frameSlides: boolean;
  includeHidden: boolean;
  scale: number;
}

/** The dialog's choices for this browsing session (every view instance). */
const session: Choices = {
  tab: "pdf",
  layout: DEFAULT_DECK_PRINT_OPTIONS.layout,
  paper: DEFAULT_DECK_PRINT_OPTIONS.paper,
  orientation: DEFAULT_DECK_PRINT_OPTIONS.orientation,
  range: DEFAULT_DECK_PRINT_OPTIONS.range,
  rangeText: "",
  noteLines: DEFAULT_DECK_PRINT_OPTIONS.noteLines,
  frameSlides: defaultFrameSlides(DEFAULT_DECK_PRINT_OPTIONS.layout),
  includeHidden: false,
  scale: 2,
};

export function useDeckExport(deps: UseDeckExportDeps) {
  const isOpen = ref(false);
  const tab = ref<DeckExportTab>(session.tab);
  const layout = ref<DeckPrintLayout>(session.layout);
  const paper = ref<DeckPrintPaper>(session.paper);
  const orientation = ref<Orientation>(session.orientation);
  const range = ref<DeckPrintRange>(session.range);
  const rangeText = ref(session.rangeText);
  const noteLines = ref(session.noteLines);
  const frameSlides = ref(session.frameSlides);
  const includeHidden = ref(session.includeHidden);
  const scale = ref(session.scale);
  const busy = ref(false);
  const progress = ref({ done: 0, total: 0 });
  /** A refusal (page/slide limits) the dialog shows until the next run. */
  const notice = ref("");
  /** The sheet's mounted pages and how they render. */
  const sheetPages = ref<SheetPage[]>([]);
  const sheetMode = ref<DeckSheetMode>("export");
  let abortRequested = false;

  // "Frame slides" defaults on for every layout but Slides.
  watch(layout, (value) => {
    frameSlides.value = defaultFrameSlides(value);
  });
  watch([tab, layout, paper, orientation, range, rangeText, noteLines, frameSlides, includeHidden, scale], () => {
    Object.assign(session, {
      tab: tab.value,
      layout: layout.value,
      paper: paper.value,
      orientation: orientation.value,
      range: range.value,
      rangeText: rangeText.value,
      noteLines: noteLines.value,
      frameSlides: frameSlides.value,
      includeHidden: includeHidden.value,
      scale: scale.value,
    });
    notice.value = "";
  });

  const options = computed<DeckPrintOptions>(() => ({
    layout: layout.value,
    paper: paper.value,
    orientation: orientation.value,
    range: range.value,
    rangeText: rangeText.value,
    currentIndex: deps.currentIndex.value,
    includeHidden: includeHidden.value,
    noteLines: noteLines.value,
    frameSlides: frameSlides.value,
  }));
  /** The image export's options: the same range at the Slides layout. */
  const imageOptions = computed<DeckPrintOptions>(() => ({ ...options.value, layout: "slides", frameSlides: false }));

  const resolved = computed(() =>
    deps.model.value ? resolveSlideRange(deps.model.value, options.value) : { slides: [], error: "" },
  );
  const rangeError = computed(() => resolved.value.error ?? "");
  const pdfPlan = computed(() => (deps.model.value ? planSheet(deps.model.value, options.value) : []));
  const pageCount = computed(() => pdfPlan.value.length);
  const hasHidden = computed(() => deps.model.value?.slides.some((slide) => slide.hidden === true) ?? false);
  const noteLinesEnabled = computed(() => layoutInfo(layout.value).noteLines === "optional");
  const noteLinesChecked = computed(() => noteLinesApply(layout.value, noteLines.value));
  const rotationWarning = computed(() =>
    deps.model.value ? rangeHasRotatedFrame(deps.model.value, resolved.value.slides) : false,
  );
  const pngSize = computed(() => {
    const model = deps.model.value;
    return model ? pngPixelSize(model.slideWidthCm, model.slideHeightCm, scale.value) : { widthPx: 0, heightPx: 0, clamped: false };
  });
  const estimate = computed(() => pngEstimate(resolved.value.slides.length, pngSize.value));
  const canRun = computed(() => !busy.value && rangeError.value.length === 0 && resolved.value.slides.length > 0);

  function open(which: DeckExportTab = "pdf"): void {
    if (!deps.model.value || deps.slideTotal.value === 0) return;
    tab.value = which;
    notice.value = "";
    isOpen.value = true;
  }

  function abort(): void {
    abortRequested = true;
  }

  function close(): void {
    abort();
    isOpen.value = false;
  }

  /** Cmd/Ctrl+P inside /slides/:id opens the dialog on the PDF tab. */
  function onPrintShortcut(event: KeyboardEvent): void {
    if ((event.metaKey || event.ctrlKey) && !event.shiftKey && !event.altKey && event.key.toLowerCase() === "p") {
      event.preventDefault();
      open("pdf");
    }
  }

  function fail(error: unknown): void {
    const reason = error instanceof Error && error.message.length > 0 ? error.message : "an unexpected error occurred";
    deps.reportError(`${DECK_EXPORT_ERROR_PREFIX} — ${reason}`);
  }

  /** Mounts, settles, extracts and unmounts each planned page in a rolling
   * window of EXPORT_WINDOW_PAGES, yielding between pages. Returns false when
   * aborted — at a page boundary, before anything downloads. */
  async function runSheet(
    model: PresentationModel,
    plan: SheetPage[],
    onPage: (input: PdfPageInput, page: SheetPage) => Promise<void>,
  ): Promise<boolean> {
    abortRequested = false;
    sheetMode.value = "export";
    progress.value = { done: 0, total: plan.length };
    try {
      for (const [index, page] of plan.entries()) {
        if (abortRequested) return false;
        sheetPages.value = plan.slice(index, index + EXPORT_WINDOW_PAGES);
        await nextTick();
        if (abortRequested) return false; // e.g. the dialog closed, taking the sheet
        const element = sheetPageElement(page.index);
        if (!element) throw new Error("the export sheet did not render");
        await settleSheet(element);
        if (abortRequested) return false;
        const input = extractPageContent(element, page.widthCm * PT_PER_CM, page.heightCm * PT_PER_CM, SHEET_EXTRACT_OPTIONS);
        input.headings = pageHeadings(model, page);
        input.rects = [...(input.rects ?? []), ...sheetPageRects(page)];
        await onPage(input, page);
        progress.value = { done: index + 1, total: plan.length };
        await yieldToMainThread();
      }
      return !abortRequested;
    } finally {
      sheetPages.value = [];
    }
  }

  async function run(work: (model: PresentationModel) => Promise<void>): Promise<void> {
    const model = deps.model.value;
    if (!model || busy.value) return;
    busy.value = true;
    notice.value = "";
    try {
      await work(model);
    } catch (error) {
      fail(error);
    } finally {
      busy.value = false;
      progress.value = { done: 0, total: 0 };
    }
  }

  /** "PDF": real pages, not screenshots — searchable text, bookmarks per
   * slide, `/Title` from the deck title. */
  function exportPdf(): Promise<void> {
    return run(async (model) => {
      const plan = pdfPlan.value;
      if (plan.length === 0) return;
      const pages: PdfPageInput[] = [];
      if (!(await runSheet(model, plan, async (input) => void pages.push(input)))) return;
      const [fonts, images] = await Promise.all([loadPdfFonts(pages), loadPdfImages(pages)]);
      const bytes = await buildPdf({ title: model.title, pages, fonts, images });
      if (abortRequested) return;
      triggerBlobDownload(bytes, pdfFileName(model.title, DEFAULT_PRESENTATION_TITLE), PDF_MIME_TYPE);
    });
  }

  /** PNGs of `plan`'s pages (Slides layout): one slide downloads as its own
   * `{title}-slide-NN.png`, more as `{title}-slides.zip` (stored — PNG is
   * already compressed). */
  async function exportPngs(model: PresentationModel, plan: SheetPage[]): Promise<void> {
    const refusal = imagesRefusal(plan.length);
    if (refusal) {
      notice.value = refusal;
      return;
    }
    const { widthPx } = pngSize.value;
    const files: { name: string; data: Uint8Array }[] = [];
    const done = await runSheet(model, plan, async (input, page) => {
      await registerRasterFaces(await loadPdfFonts([input]));
      const slideIndex = page.slots[0]?.slideIndex ?? page.index;
      files.push({
        name: slidePngFileName(model.title, slideIndex + 1, model.slides.length),
        data: await rasterisePdfPage(input, widthPx),
      });
    });
    if (!done || files.length === 0) return;
    if (files.length === 1) {
      triggerBlobDownload(files[0].data, files[0].name, PNG_MIME_TYPE);
      return;
    }
    const archive = await zip(files.map((file) => ({ ...file, store: true })));
    if (abortRequested) return;
    triggerBlobDownload(archive, slidesZipFileName(model.title), ZIP_MIME_TYPE);
  }

  /** "Images": the chosen range at the Slides layout, at the chosen scale. */
  function exportImages(): Promise<void> {
    return run((model) => exportPngs(model, planSheet(model, imageOptions.value)));
  }

  /** One slide (the current one by default) as a PNG at the chosen scale,
   * hidden or not. */
  function exportSlidePng(index = deps.currentIndex.value): Promise<void> {
    return run((model) =>
      exportPngs(model, planSheet(model, { ...imageOptions.value, range: "current", currentIndex: index, includeHidden: true })),
    );
  }

  /** "Print…": the browser prints the sheet — only the sheet shows under
   * @media print (body.is-deck-print), on pages sized by an injected @page
   * rule: the deck's own slide size for Slides, the chosen paper otherwise. */
  function print(): Promise<void> {
    return run(async (model) => {
      const plan = pdfPlan.value;
      const refusal = printRefusal(plan.length);
      if (refusal || plan.length === 0) {
        notice.value = refusal;
        return;
      }
      const size = sheetPageSize({ ...options.value, slideWidthCm: model.slideWidthCm, slideHeightCm: model.slideHeightCm });
      const style = document.createElement("style");
      style.dataset.deckPrint = "";
      style.textContent = `@page { size: ${size.widthCm}cm ${size.heightCm}cm; margin: 0; }`;
      sheetMode.value = "print";
      sheetPages.value = plan;
      try {
        await nextTick();
        document.body.classList.add("is-deck-print");
        document.head.append(style);
        const sheet = sheetElement();
        if (sheet) await settleSheet(sheet);
        window.print();
      } finally {
        document.body.classList.remove("is-deck-print");
        style.remove();
        sheetPages.value = [];
        sheetMode.value = "export";
      }
    });
  }

  return {
    /** The deck the sheet renders (read-only). */
    deck: deps.model,
    isOpen,
    tab,
    layout,
    paper,
    orientation,
    range,
    rangeText,
    rangeError,
    noteLines,
    noteLinesEnabled,
    noteLinesChecked,
    frameSlides,
    includeHidden,
    hasHidden,
    scale,
    estimate,
    pngSize,
    pageCount,
    rotationWarning,
    progress,
    busy,
    canRun,
    notice,
    sheetPages,
    sheetMode,
    open,
    print,
    exportPdf,
    exportImages,
    exportSlidePng,
    abort,
    close,
    onPrintShortcut,
  };
}

export type DeckExport = ReturnType<typeof useDeckExport>;
