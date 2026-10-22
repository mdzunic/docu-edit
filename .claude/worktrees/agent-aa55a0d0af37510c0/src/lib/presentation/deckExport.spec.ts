// SPEC-104 — the framework-free rules of deck print/export: the PNG pixel
// size and its clamp, download names, the limits, per-page bookmarks, the
// rotation note and the plan-derived frame/note-line rects.
import { describe, expect, it } from "vitest";
import type { PresentationModel, Slide, SlideFrame } from "./model";
import {
  MAX_PNG_EDGE_PX,
  MAX_PNG_EXPORT_SLIDES,
  PRINT_PAGE_LIMIT,
  imagesRefusal,
  pageHeadings,
  pngEstimate,
  pngPixelSize,
  printRefusal,
  rangeHasRotatedFrame,
  sheetPageRects,
  slidePngFileName,
  slidesZipFileName,
} from "./deckExport";
import { DEFAULT_DECK_PRINT_OPTIONS, planSheet } from "./printLayout";
import { PT_PER_CM } from "../core/pdfExtract";

function textFrame(text: string, extra: Partial<SlideFrame> = {}): SlideFrame {
  return {
    id: crypto.randomUUID(),
    kind: "text",
    role: "title",
    xCm: 1,
    yCm: 1,
    widthCm: 10,
    heightCm: 2,
    paragraphs: [{ runs: [{ text }] }],
    ...extra,
  } as SlideFrame;
}

function deck(slides: Slide[]): PresentationModel {
  return Object.freeze({
    kind: "presentation",
    id: "d",
    title: "Deck",
    format: "pptx",
    slides,
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
  }) as PresentationModel;
}

describe("pngPixelSize", () => {
  it("is exactly 1280×720 / 2560×1440 / 5120×2880 for a 16:9 deck at 1×/2×/4×", () => {
    expect(pngPixelSize(33.867, 19.05, 1)).toEqual({ widthPx: 1280, heightPx: 720, clamped: false });
    expect(pngPixelSize(33.867, 19.05, 2)).toEqual({ widthPx: 2560, heightPx: 1440, clamped: false });
    expect(pngPixelSize(33.867, 19.05, 4)).toEqual({ widthPx: 5120, heightPx: 2880, clamped: false });
  });

  it("clamps the longer edge to MAX_PNG_EDGE_PX before the raster call", () => {
    // A 142 cm-wide banner deck at 4× would be ~21 000 px wide.
    const size = pngPixelSize(142.24, 19.05, 4);
    expect(size.clamped).toBe(true);
    expect(size.widthPx).toBe(MAX_PNG_EDGE_PX);
    expect(size.heightPx).toBe(Math.round(MAX_PNG_EDGE_PX * (19.05 / 142.24)));
    // A tall portrait deck clamps on its height.
    const tall = pngPixelSize(19.05, 142.24, 4);
    expect(tall.clamped).toBe(true);
    expect(tall.heightPx).toBeLessThanOrEqual(MAX_PNG_EDGE_PX);
    expect(Math.max(tall.widthPx, tall.heightPx)).toBeGreaterThan(MAX_PNG_EDGE_PX - 4);
  });

  it("the estimate names the size and the clamp", () => {
    expect(pngEstimate(1, pngPixelSize(33.867, 19.05, 2))).toBe("1 image, 2560 × 1440 px");
    expect(pngEstimate(3, pngPixelSize(33.867, 19.05, 1))).toBe("3 images, 1280 × 720 px each");
    expect(pngEstimate(1, pngPixelSize(142.24, 19.05, 4))).toContain(`clamped to ${MAX_PNG_EDGE_PX} px`);
  });
});

describe("download names", () => {
  it("zero-pads the slide number to the deck's digit count, two digits at least", () => {
    expect(slidePngFileName("Pitch", 1, 3)).toBe("Pitch-slide-01.png");
    expect(slidePngFileName("Pitch", 7, 120)).toBe("Pitch-slide-007.png");
    expect(slidePngFileName("  ", 2, 9)).toBe("Untitled presentation-slide-02.png");
  });

  it("names the range archive after the deck", () => {
    expect(slidesZipFileName("Pitch")).toBe("Pitch-slides.zip");
  });
});

describe("limits", () => {
  it("refuses a print over PRINT_PAGE_LIMIT pages with a message naming the limit", () => {
    expect(printRefusal(PRINT_PAGE_LIMIT)).toBe("");
    expect(printRefusal(PRINT_PAGE_LIMIT + 1)).toContain(String(PRINT_PAGE_LIMIT));
  });

  it("refuses an image export over MAX_PNG_EXPORT_SLIDES", () => {
    expect(imagesRefusal(MAX_PNG_EXPORT_SLIDES)).toBe("");
    expect(imagesRefusal(MAX_PNG_EXPORT_SLIDES + 1)).toContain(String(MAX_PNG_EXPORT_SLIDES));
  });
});

describe("pageHeadings", () => {
  it("emits one level-1 heading per placed slide — its title, else \"Slide n\"", () => {
    const model = deck([
      { id: "a", frames: [textFrame("Welcome")] },
      { id: "b", frames: [] },
      { id: "c", frames: [textFrame("Close")] },
    ]);
    const [page] = planSheet(model, { ...DEFAULT_DECK_PRINT_OPTIONS, layout: "handout-4" });
    expect(pageHeadings(model, page)).toEqual([
      { level: 1, title: "Welcome", yPt: page.slots[0].slide.yCm * PT_PER_CM },
      { level: 1, title: "Slide 2", yPt: page.slots[1].slide.yCm * PT_PER_CM },
      { level: 1, title: "Close", yPt: page.slots[2].slide.yCm * PT_PER_CM },
    ]);
  });
});

describe("rangeHasRotatedFrame", () => {
  it("finds a rotated frame in the range, groups included", () => {
    const model = deck([
      { id: "a", frames: [textFrame("Plain")] },
      { id: "b", frames: [textFrame("Tilted", { rotationDeg: 30 })] },
      {
        id: "c",
        frames: [{ id: "g", kind: "group", xCm: 0, yCm: 0, widthCm: 5, heightCm: 5, children: [textFrame("In", { rotationDeg: 90 })] } as SlideFrame],
      },
    ]);
    expect(rangeHasRotatedFrame(model, [0])).toBe(false);
    expect(rangeHasRotatedFrame(model, [0, 1])).toBe(true);
    expect(rangeHasRotatedFrame(model, [2])).toBe(true);
  });
});

describe("sheetPageRects", () => {
  it("outlines each framed slot with a 0.75 pt black stroke on top, and rules each note line at 0.5 pt #808080", () => {
    const model = deck([{ id: "a", frames: [] }]);
    const [page] = planSheet(model, { ...DEFAULT_DECK_PRINT_OPTIONS, layout: "handout-1", noteLines: true, frameSlides: true });
    const rects = sheetPageRects(page);
    const slot = page.slots[0];
    expect(rects[0]).toEqual({
      xPt: slot.slide.xCm * PT_PER_CM,
      yPt: slot.slide.yCm * PT_PER_CM,
      widthPt: slot.slide.widthCm * PT_PER_CM,
      heightPt: slot.slide.heightCm * PT_PER_CM,
      stroke: [0, 0, 0],
      strokeWidthPt: 0.75,
      onTop: true,
    });
    const lines = rects.slice(1);
    expect(lines).toHaveLength(slot.noteLines!.yCm.length);
    for (const [index, line] of lines.entries()) {
      expect(line.fill).toBe("#808080");
      expect(line.heightPt).toBe(0.5);
      expect(line.yPt + 0.25).toBeCloseTo(slot.noteLines!.yCm[index] * PT_PER_CM, 6);
    }
  });

  it("is empty for an unframed Slides page", () => {
    const [page] = planSheet(deck([{ id: "a", frames: [] }]), DEFAULT_DECK_PRINT_OPTIONS);
    expect(sheetPageRects(page)).toEqual([]);
  });
});
