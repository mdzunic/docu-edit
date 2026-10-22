// SPEC-104 — the deck print/export controller: range + hidden resolution, the
// rolling mount window and its progress ticks, abort at a page boundary (no
// download, nothing mutated), the PRINT_PAGE_LIMIT refusal and the
// MAX_PNG_EDGE_PX clamp. The DOM-bound seams — the mounted sheet, the
// extractor, the rasteriser, font/image loading and the download — are
// stand-ins; buildPdf is real. The model is deep-frozen throughout.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { effectScope, ref, shallowRef, watch } from "vue";
import type { PresentationModel, Slide } from "../../lib/presentation/model";
import { MAX_PNG_EDGE_PX, PRINT_PAGE_LIMIT } from "../../lib/presentation/deckExport";
import { useDeckExport } from "./useDeckExport";

const h = vi.hoisted(() => ({
  downloads: [] as { bytes: Uint8Array; name: string; type: string }[],
  rasterWidths: [] as number[],
  extracted: [] as number[],
  onSettle: null as null | ((count: number) => void),
  settles: 0,
}));

vi.mock("../../lib/core/triggerDownload", () => ({
  triggerBlobDownload: (bytes: Uint8Array, name: string, type: string) => {
    h.downloads.push({ bytes, name, type });
  },
}));
vi.mock("../../components/presentation/deckPrintSheet", () => ({
  SHEET_EXTRACT_OPTIONS: {},
  sheetElement: () => document.createElement("div"),
  sheetPageElement: (index: number) => {
    const el = document.createElement("div");
    el.dataset.pageIndex = String(index);
    return el;
  },
  settleSheet: async () => {
    h.settles += 1;
    h.onSettle?.(h.settles);
  },
}));
vi.mock("../../lib/core/pdfExtract", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../lib/core/pdfExtract")>();
  return {
    ...actual,
    extractPageContent: (el: HTMLElement, widthPt: number, heightPt: number) => {
      h.extracted.push(Number(el.dataset.pageIndex));
      return { widthPt, heightPt, items: [], links: [], headings: [], images: [], paths: [] };
    },
  };
});
vi.mock("../../lib/pdfAssets", () => ({
  loadPdfFonts: async () => new Map(),
  loadPdfImages: async () => new Map(),
  registerRasterFaces: async () => undefined,
}));
vi.mock("../../lib/core/pdfRaster", () => ({
  rasterisePdfPage: async (_page: unknown, widthPx: number) => {
    h.rasterWidths.push(widthPx);
    return new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
  },
}));

function deepFreeze<T>(value: T): T {
  if (value && typeof value === "object" && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const child of Object.values(value)) deepFreeze(child);
  }
  return value;
}

function deck(count: number, hidden: number[] = [], widthCm = 33.867, heightCm = 19.05): PresentationModel {
  const slides: Slide[] = Array.from({ length: count }, (_, index) => {
    const slide: Slide = { id: `s${index}`, frames: [] };
    if (hidden.includes(index)) slide.hidden = true;
    return slide;
  });
  return deepFreeze({
    kind: "presentation",
    id: "d",
    title: "Pitch",
    format: "pptx",
    slides,
    slideWidthCm: widthCm,
    slideHeightCm: heightCm,
  } as PresentationModel);
}

function harness(model: PresentationModel, currentIndex = 0) {
  const errors: string[] = [];
  const scope = effectScope();
  const c = scope.run(() =>
    useDeckExport({
      model: shallowRef<PresentationModel | null>(model),
      currentIndex: ref(currentIndex),
      slideTotal: ref(model.slides.length),
      reportError: (message) => errors.push(message),
    }),
  )!;
  // Fresh choices for every test (they persist for the session by design).
  c.layout.value = "slides";
  c.paper.value = "A4";
  c.orientation.value = "portrait";
  c.range.value = "all";
  c.rangeText.value = "";
  c.includeHidden.value = false;
  c.noteLines.value = false;
  c.scale.value = 2;
  c.tab.value = "pdf";
  const windows: number[] = [];
  const ticks: string[] = [];
  scope.run(() => {
    watch(c.sheetPages, (pages) => windows.push(pages.length), { flush: "sync" });
    watch(c.progress, (p) => ticks.push(`${p.done}/${p.total}`), { flush: "sync" });
  });
  return { c, errors, windows, ticks, stop: () => scope.stop() };
}

function pageTreeCount(bytes: Uint8Array): number {
  const match = /\/Type \/Pages \/Kids \[[^\]]*\] \/Count (\d+)/.exec(new TextDecoder("latin1").decode(bytes));
  return match ? Number(match[1]) : -1;
}

beforeEach(() => {
  h.downloads.length = 0;
  h.rasterWidths.length = 0;
  h.extracted.length = 0;
  h.onSettle = null;
  h.settles = 0;
});

afterEach(() => vi.restoreAllMocks());

describe("useDeckExport dialog state", () => {
  it("opens on the requested tab; Cmd/Ctrl+P opens it on the PDF tab", () => {
    const { c } = harness(deck(3));
    c.open("images");
    expect([c.isOpen.value, c.tab.value]).toEqual([true, "images"]);
    c.close();
    const event = new KeyboardEvent("keydown", { key: "p", ctrlKey: true, cancelable: true });
    c.onPrintShortcut(event);
    expect([c.isOpen.value, c.tab.value, event.defaultPrevented]).toEqual([true, "pdf", true]);
    c.close();
    c.onPrintShortcut(new KeyboardEvent("keydown", { key: "p", ctrlKey: true, shiftKey: true }));
    expect(c.isOpen.value).toBe(false);
  });

  it("defaults Frame slides on for every layout but Slides", async () => {
    const { c } = harness(deck(3));
    c.layout.value = "handout-4";
    await Promise.resolve();
    expect(c.frameSlides.value).toBe(true);
    c.layout.value = "slides";
    await Promise.resolve();
    expect(c.frameSlides.value).toBe(false);
  });

  it("resolves the range and the hidden filter into a page count and range error", () => {
    const { c } = harness(deck(3, [1]), 1);
    c.layout.value = "slides";
    expect(c.pageCount.value).toBe(2);
    c.includeHidden.value = true;
    expect(c.pageCount.value).toBe(3);
    c.includeHidden.value = false;
    c.range.value = "current";
    expect(c.rangeError.value).toContain("hidden");
    expect(c.canRun.value).toBe(false);
    c.range.value = "custom";
    c.rangeText.value = "9";
    expect(c.rangeError.value).toBe("Slides go from 1 to 3");
    c.rangeText.value = "1, 3";
    expect(c.rangeError.value).toBe("");
    expect(c.canRun.value).toBe(true);
    expect(c.hasHidden.value).toBe(true);
    expect(harness(deck(2)).c.hasHidden.value).toBe(false);
  });

  it("enables note lines only for handouts 1/2/4 and forces them on for 3-up", () => {
    const { c } = harness(deck(3));
    const states = (["slides", "notes", "handout-1", "handout-2", "handout-3", "handout-4", "handout-6", "handout-9"] as const).map((layout) => {
      c.layout.value = layout;
      return [layout, c.noteLinesEnabled.value, c.noteLinesChecked.value];
    });
    expect(states).toEqual([
      ["slides", false, false],
      ["notes", false, false],
      ["handout-1", true, false],
      ["handout-2", true, false],
      ["handout-3", false, true],
      ["handout-4", true, false],
      ["handout-6", false, false],
      ["handout-9", false, false],
    ]);
  });
});

describe("useDeckExport PDF export", () => {
  it("mounts, extracts and unmounts pages in a rolling window of 4, ticking progress, then downloads one PDF", async () => {
    const model = deck(6);
    const { c, windows, ticks } = harness(model);
    await c.exportPdf();
    expect(h.extracted).toEqual([0, 1, 2, 3, 4, 5]);
    expect(Math.max(...windows)).toBeLessThanOrEqual(4);
    expect(windows).toEqual([4, 4, 4, 3, 2, 1, 0]);
    expect(ticks).toEqual(["0/6", "1/6", "2/6", "3/6", "4/6", "5/6", "6/6", "0/0"]);
    expect(h.downloads).toHaveLength(1);
    const [download] = h.downloads;
    expect([download.name, download.type]).toEqual(["Pitch.pdf", "application/pdf"]);
    expect(new TextDecoder("latin1").decode(download.bytes.subarray(0, 8))).toBe("%PDF-1.7");
    expect(pageTreeCount(download.bytes)).toBe(6);
    // Bookmarks: one per slide, "Slide n" for untitled slides.
    expect(new TextDecoder("latin1").decode(download.bytes)).toContain("/Title (Slide 6)");
    expect(c.busy.value).toBe(false);
    expect(c.sheetPages.value).toEqual([]);
  });

  it("abort stops at the next page boundary: no download, nothing mounted, the model untouched", async () => {
    const model = deck(6);
    const before = JSON.stringify(model);
    const { c } = harness(model);
    h.onSettle = (count) => {
      if (count === 2) c.abort();
    };
    await c.exportPdf();
    expect(h.extracted).toEqual([0]);
    expect(h.downloads).toEqual([]);
    expect(c.sheetPages.value).toEqual([]);
    expect(c.busy.value).toBe(false);
    expect(JSON.stringify(model)).toBe(before);
  });
});

describe("useDeckExport print", () => {
  it("refuses a print over PRINT_PAGE_LIMIT pages with a message naming the limit", async () => {
    const print = vi.spyOn(window, "print").mockImplementation(() => undefined);
    const { c } = harness(deck(PRINT_PAGE_LIMIT + 1));
    await c.print();
    expect(print).not.toHaveBeenCalled();
    expect(c.notice.value).toContain(String(PRINT_PAGE_LIMIT));
  });

  it("marks body.is-deck-print and injects an @page rule sized to the deck for Slides, then restores", async () => {
    const seen: string[] = [];
    vi.spyOn(window, "print").mockImplementation(() => {
      seen.push(String(document.body.classList.contains("is-deck-print")));
      seen.push(document.querySelector("style[data-deck-print]")?.textContent ?? "");
    });
    const { c } = harness(deck(2, [], 25.4, 19.05));
    await c.print();
    expect(seen).toEqual(["true", "@page { size: 25.4cm 19.05cm; margin: 0; }"]);
    expect(document.body.classList.contains("is-deck-print")).toBe(false);
    expect(document.querySelector("style[data-deck-print]")).toBeNull();
    expect(c.sheetPages.value).toEqual([]);
  });

  it("uses the chosen paper for the other layouts", async () => {
    let rule = "";
    vi.spyOn(window, "print").mockImplementation(() => {
      rule = document.querySelector("style[data-deck-print]")?.textContent ?? "";
    });
    const { c } = harness(deck(2));
    c.layout.value = "handout-6";
    c.paper.value = "Letter";
    c.orientation.value = "landscape";
    await c.print();
    expect(rule).toBe("@page { size: 27.94cm 21.59cm; margin: 0; }");
  });
});

describe("useDeckExport image export", () => {
  it("one slide downloads as {title}-slide-NN.png at round(cmToPx(width) × scale)", async () => {
    const { c } = harness(deck(3));
    c.range.value = "custom";
    c.rangeText.value = "1";
    await c.exportImages();
    expect(h.rasterWidths).toEqual([2560]);
    expect(h.downloads.map((d) => [d.name, d.type])).toEqual([["Pitch-slide-01.png", "image/png"]]);
  });

  it("the whole range downloads as {title}-slides.zip with stored entries", async () => {
    const { c } = harness(deck(3, [1]));
    c.scale.value = 1;
    await c.exportImages();
    expect(h.rasterWidths).toEqual([1280, 1280]);
    const [archive] = h.downloads;
    expect([archive.name, archive.type]).toEqual(["Pitch-slides.zip", "application/zip"]);
    const { unzip } = await import("../../lib/core/zip");
    expect([...(await unzip(archive.bytes)).keys()]).toEqual(["Pitch-slide-01.png", "Pitch-slide-03.png"]);
  });

  it("clamps the raster width so the longer edge stays within MAX_PNG_EDGE_PX", async () => {
    const { c } = harness(deck(1, [], 142.24, 19.05));
    c.scale.value = 4;
    expect(c.estimate.value).toContain("clamped");
    await c.exportImages();
    expect(h.rasterWidths).toEqual([MAX_PNG_EDGE_PX]);
  });

  it("exportSlidePng exports the current slide even when it is hidden", async () => {
    const { c } = harness(deck(3, [1]), 1);
    await c.exportSlidePng();
    expect(h.downloads.map((d) => d.name)).toEqual(["Pitch-slide-02.png"]);
  });

  it("refuses more than 200 slides at a time", async () => {
    const { c } = harness(deck(201));
    await c.exportImages();
    expect(h.downloads).toEqual([]);
    expect(c.notice.value).toContain("200");
  });
});
