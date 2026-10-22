// SPEC-104 — the deck print sheet planner: slot geometry for every layout,
// note-line rows, page counts, range resolution and the hidden filter. The
// model is deep-frozen throughout: planning a print mutates nothing.
import { describe, expect, it } from "vitest";
import type { PresentationModel, Slide } from "./model";
import {
  DECK_PRINT_LAYOUTS,
  DEFAULT_DECK_PRINT_OPTIONS,
  HANDOUT_GUTTER_CM,
  HANDOUT_MARGIN_CM,
  NOTE_LINE_SPACING_CM,
  defaultFrameSlides,
  noteLinesApply,
  paginateNotes,
  planSheet,
  plannedSlides,
  resolveSlideRange,
  sheetPageSize,
  type DeckPrintOptions,
  type SheetSlot,
} from "./printLayout";

function deepFreeze<T>(value: T): T {
  if (value && typeof value === "object" && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const child of Object.values(value)) {
      deepFreeze(child);
    }
  }
  return value;
}

/** A frozen 16:9 deck of `count` slides; `hidden` lists 0-based hidden ones. */
function deck(count: number, hidden: number[] = [], notes: Record<number, string[]> = {}): PresentationModel {
  const slides: Slide[] = Array.from({ length: count }, (_, index) => {
    const slide: Slide = { id: `s${index}`, frames: [] };
    if (hidden.includes(index)) slide.hidden = true;
    if (notes[index]) slide.notes = notes[index];
    return slide;
  });
  return deepFreeze({
    kind: "presentation",
    id: "d",
    title: "Deck",
    format: "pptx",
    slides,
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
  } as PresentationModel);
}

function options(overrides: Partial<DeckPrintOptions> = {}): DeckPrintOptions {
  return { ...DEFAULT_DECK_PRINT_OPTIONS, ...overrides };
}

const close = (actual: number, expected: number) => expect(actual).toBeCloseTo(expected, 4);

function expectInside(slot: SheetSlot): void {
  expect(slot.slide.xCm).toBeGreaterThanOrEqual(slot.cell.xCm - 1e-6);
  expect(slot.slide.yCm).toBeGreaterThanOrEqual(slot.cell.yCm - 1e-6);
  expect(slot.slide.xCm + slot.slide.widthCm).toBeLessThanOrEqual(slot.cell.xCm + slot.cell.widthCm + 1e-6);
  expect(slot.slide.yCm + slot.slide.heightCm).toBeLessThanOrEqual(slot.cell.yCm + slot.cell.heightCm + 1e-6);
  // Native aspect.
  close(slot.slide.widthCm / slot.slide.heightCm, 33.867 / 19.05);
}

describe("DECK_PRINT_LAYOUTS", () => {
  it("lists Slides, Notes pages and the six handouts with their grids", () => {
    expect(DECK_PRINT_LAYOUTS.map((l) => [l.value, l.label, l.columns, l.rows])).toEqual([
      ["slides", "Slides", 1, 1],
      ["notes", "Notes pages", 1, 1],
      ["handout-1", "Handout 1 per page", 1, 1],
      ["handout-2", "Handout 2 per page", 1, 2],
      ["handout-3", "Handout 3 per page", 1, 3],
      ["handout-4", "Handout 4 per page", 2, 2],
      ["handout-6", "Handout 6 per page", 2, 3],
      ["handout-9", "Handout 9 per page", 3, 3],
    ]);
  });

  it("note lines: optional for 1/2/4-up, always for 3-up, never for 6/9-up, notes or slides", () => {
    for (const layout of ["handout-1", "handout-2", "handout-4"] as const) {
      expect([noteLinesApply(layout, true), noteLinesApply(layout, false)]).toEqual([true, false]);
    }
    expect([noteLinesApply("handout-3", true), noteLinesApply("handout-3", false)]).toEqual([true, true]);
    for (const layout of ["handout-6", "handout-9", "notes", "slides"] as const) {
      expect(noteLinesApply(layout, true)).toBe(false);
    }
  });

  it("frames slides by default on every layout but Slides", () => {
    expect(DECK_PRINT_LAYOUTS.map((l) => defaultFrameSlides(l.value))).toEqual([false, true, true, true, true, true, true, true]);
  });
});

describe("sheetPageSize", () => {
  const size = { slideWidthCm: 25.4, slideHeightCm: 19.05 };
  it("is the deck's own slide size for the Slides layout", () => {
    expect(sheetPageSize({ ...size, layout: "slides", paper: "A4", orientation: "portrait" })).toEqual({ widthCm: 25.4, heightCm: 19.05 });
  });

  it("is the chosen paper in the chosen orientation otherwise (A4 default)", () => {
    expect(sheetPageSize({ ...size, layout: "handout-4", paper: "A4", orientation: "portrait" })).toEqual({ widthCm: 21, heightCm: 29.7 });
    expect(sheetPageSize({ ...size, layout: "notes", paper: "Letter", orientation: "landscape" })).toEqual({ widthCm: 27.94, heightCm: 21.59 });
  });
});

describe("resolveSlideRange", () => {
  const model = deck(5, [1, 3]);

  it("all: every visible slide, hidden ones filtered unless included", () => {
    expect(resolveSlideRange(model, options()).slides).toEqual([0, 2, 4]);
    expect(resolveSlideRange(model, options({ includeHidden: true })).slides).toEqual([0, 1, 2, 3, 4]);
  });

  it("current: the editor's current slide", () => {
    expect(resolveSlideRange(model, options({ range: "current", currentIndex: 2 })).slides).toEqual([2]);
    const hidden = resolveSlideRange(model, options({ range: "current", currentIndex: 1 }));
    expect(hidden.slides).toEqual([]);
    expect(hidden.error).toContain("hidden");
    expect(resolveSlideRange(model, options({ range: "current", currentIndex: 1, includeHidden: true })).slides).toEqual([1]);
  });

  it("custom: slide numbers parsed by core/pageRange with the slide noun", () => {
    expect(resolveSlideRange(model, options({ range: "custom", rangeText: "1-3, 5" })).slides).toEqual([0, 2, 4]);
    expect(resolveSlideRange(model, options({ range: "custom", rangeText: "9" })).error).toBe("Slides go from 1 to 5");
    expect(resolveSlideRange(model, options({ range: "custom", rangeText: "x" })).error).toContain("is not a slide number");
    expect(resolveSlideRange(model, options({ range: "custom", rangeText: "2, 4" })).error).toContain("hidden");
  });

  it("an empty deck has nothing to print", () => {
    expect(resolveSlideRange(deck(0), options()).error).toBeTruthy();
  });
});

describe("planSheet geometry", () => {
  it("slides: one full-bleed page per slide, sized to the deck, no page number", () => {
    const pages = planSheet(deck(3, [1]), options({ layout: "slides" }));
    expect(pages).toHaveLength(2);
    for (const page of pages) {
      expect([page.widthCm, page.heightCm]).toEqual([33.867, 19.05]);
      expect(page.slots).toHaveLength(1);
      expect(page.slots[0].slide).toEqual({ xCm: 0, yCm: 0, widthCm: 33.867, heightCm: 19.05 });
      expect(page.slots[0].scale).toBe(1);
      expect(page.pageNumber).toBeUndefined();
    }
    expect(plannedSlides(pages)).toEqual([0, 2]);
  });

  it.each([
    ["handout-1", 1, 1],
    ["handout-2", 1, 2],
    ["handout-3", 1, 3],
    ["handout-4", 2, 2],
    ["handout-6", 2, 3],
    ["handout-9", 3, 3],
  ] as const)("%s: a %i × %i grid inside the margins, gutters between cells", (layout, columns, rows) => {
    const perPage = columns * rows;
    const pages = planSheet(deck(perPage + 1), options({ layout, frameSlides: true }));
    expect(pages).toHaveLength(2); // one overflow slide starts a second page
    const [first] = pages;
    expect([first.widthCm, first.heightCm]).toEqual([21, 29.7]);
    expect(first.slots).toHaveLength(perPage);
    const cellWidth = (21 - 2 * HANDOUT_MARGIN_CM - (columns - 1) * HANDOUT_GUTTER_CM) / columns;
    const cellHeight = (29.7 - 2 * HANDOUT_MARGIN_CM - (rows - 1) * HANDOUT_GUTTER_CM) / rows;
    first.slots.forEach((slot, offset) => {
      close(slot.cell.xCm, HANDOUT_MARGIN_CM + (offset % columns) * (cellWidth + HANDOUT_GUTTER_CM));
      close(slot.cell.yCm, HANDOUT_MARGIN_CM + Math.floor(offset / columns) * (cellHeight + HANDOUT_GUTTER_CM));
      close(slot.cell.widthCm, cellWidth);
      close(slot.cell.heightCm, cellHeight);
      expectInside(slot);
      expect(slot.framed).toBe(true);
      // Centred in its slide area.
      if (!slot.noteLines) {
        close(slot.slide.xCm + slot.slide.widthCm / 2, slot.cell.xCm + slot.cell.widthCm / 2);
        close(slot.slide.yCm + slot.slide.heightCm / 2, slot.cell.yCm + slot.cell.heightCm / 2);
      }
    });
    expect(pages.map((page) => page.pageNumber)).toEqual(["1", "2"]);
    expect(pages[1].slots.map((slot) => slot.slideIndex)).toEqual([perPage]);
  });

  it("handout-4 of three visible slides (the fixture's two + one) is one page", () => {
    expect(planSheet(deck(3, [1]), options({ layout: "handout-4" }))).toHaveLength(1);
    expect(planSheet(deck(3, [1]), options({ layout: "handout-4", includeHidden: true }))[0].slots).toHaveLength(3);
  });

  it("note lines split the slot 55/45 — slide left, rules right, 0.85 cm apart", () => {
    const [page] = planSheet(deck(1), options({ layout: "handout-1", noteLines: true }));
    const slot = page.slots[0];
    expect(slot.noteLines).toBeDefined();
    const lines = slot.noteLines!;
    expect(slot.slide.xCm + slot.slide.widthCm).toBeLessThanOrEqual(slot.cell.xCm + slot.cell.widthCm * 0.55 + 1e-6);
    close(lines.xCm, slot.cell.xCm + slot.cell.widthCm * 0.55 + HANDOUT_GUTTER_CM / 2);
    close(lines.xCm + lines.widthCm, slot.cell.xCm + slot.cell.widthCm);
    expect(lines.yCm.length).toBe(Math.floor(slot.slide.heightCm / NOTE_LINE_SPACING_CM + 1e-9));
    lines.yCm.forEach((y, index) => close(y, slot.slide.yCm + (index + 1) * NOTE_LINE_SPACING_CM));
    expect(Math.max(...lines.yCm)).toBeLessThanOrEqual(slot.slide.yCm + slot.slide.heightCm + 1e-6);
  });

  it("3-up always carries note lines; 6/9-up never do", () => {
    expect(planSheet(deck(3), options({ layout: "handout-3", noteLines: false }))[0].slots[0].noteLines).toBeDefined();
    expect(planSheet(deck(6), options({ layout: "handout-6", noteLines: true }))[0].slots[0].noteLines).toBeUndefined();
    expect(planSheet(deck(9), options({ layout: "handout-9", noteLines: true }))[0].slots[0].noteLines).toBeUndefined();
  });

  it("notes: the slide fitted into the top half, its notes below, a page number", () => {
    const model = deck(2, [], { 0: ["First point", "Second point"] });
    const pages = planSheet(model, options({ layout: "notes" }));
    expect(pages).toHaveLength(2);
    const [page] = pages;
    const slot = page.slots[0];
    expectInside(slot);
    expect(slot.slide.yCm + slot.slide.heightCm).toBeLessThanOrEqual(29.7 / 2);
    expect(page.notes?.paragraphs).toEqual(["First point", "Second point"]);
    expect(page.notes?.box.yCm).toBeGreaterThanOrEqual(29.7 / 2);
    expect(page.notes?.continued).toBe(false);
    expect(page.pageNumber).toBe("1");
    expect(page.pageNumberBox?.yCm).toBeCloseTo(29.7 - HANDOUT_MARGIN_CM, 6);
    // A slide without notes still gets its page, with an empty notes block.
    expect(pages[1].notes?.paragraphs).toEqual([]);
  });

  it("notes overflow onto continuation pages that repeat the notes only", () => {
    const long = Array.from({ length: 60 }, (_, i) => `Talking point number ${i + 1} with a little more detail to say.`);
    const pages = planSheet(deck(1, [], { 0: long }), options({ layout: "notes" }));
    expect(pages.length).toBeGreaterThan(1);
    expect(pages[0].slots).toHaveLength(1);
    for (const page of pages.slice(1)) {
      expect(page.slots).toEqual([]);
      expect(page.notes?.continued).toBe(true);
      expect(page.notes?.box.yCm).toBe(HANDOUT_MARGIN_CM);
    }
    expect(pages.flatMap((page) => page.notes?.paragraphs ?? [])).toEqual(long);
    expect(pages.map((page) => page.pageNumber)).toEqual(pages.map((_, i) => String(i + 1)));
  });

  it("an empty range plans no pages", () => {
    expect(planSheet(deck(2, [0, 1]), options({ layout: "handout-4" }))).toEqual([]);
  });
});

describe("paginateNotes", () => {
  it("keeps paragraphs whole when they fit, and splits one that straddles a break", () => {
    expect(paginateNotes(["a", "b"], 10, 5, 5)).toEqual([["a", "b"]]);
    expect(paginateNotes(["aaaa bbbb cccc"], 4, 2, 5)).toEqual([["aaaa bbbb"], ["cccc"]]);
    expect(paginateNotes([""], 10, 1, 1)).toEqual([[""]]);
    expect(paginateNotes([], 10, 1, 1)).toEqual([[]]);
  });
});
