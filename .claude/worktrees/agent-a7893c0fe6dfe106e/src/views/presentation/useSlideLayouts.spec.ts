// SPEC-089 layout gallery session: cards (deck layouts + synthesised
// fallbacks), open modes, apply/new/reset sequencing (record → mutate → commit),
// the kept-in-place report, and adopting a synthesised fallback into the deck.
import { computed, ref } from "vue";
import { describe, expect, it, vi } from "vitest";
import { useSlideLayouts } from "./useSlideLayouts";
import { instantiateStandardLayout, STANDARD_LAYOUTS, standardLayoutKey } from "../../lib/presentation/layouts";
import type { PresentationModel, Slide } from "../../lib/presentation/model";

function setup(slide: Slide, layouts = [] as PresentationModel["layouts"]) {
  const model = ref<PresentationModel>({
    kind: "presentation",
    id: "m",
    title: "T",
    format: "pptx",
    slides: [slide],
    layouts,
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
  });
  const currentSlide = computed(() => model.value.slides[0] ?? null);
  const currentIndex = ref(0);
  const record = vi.fn();
  const commit = vi.fn();
  const addSlide = vi.fn();
  const c = useSlideLayouts({ model, currentSlide, currentIndex, record, commit, addSlide });
  return { model, c, record, commit, addSlide };
}

const titleAndBody = instantiateStandardLayout(
  STANDARD_LAYOUTS.find((d) => d.kind === "titleAndBody")!,
  33.867,
  19.05,
);

describe("cards", () => {
  it("lists the deck's own layouts then synthesised fallbacks for missing kinds", () => {
    const deckLayout = { key: "ppt/slideLayouts/slideLayout1.xml", kind: "title" as const, name: "Deck Title", slots: [] };
    const { c } = setup({ id: "s", frames: [] }, [deckLayout]);
    const cards = c.cards.value;
    expect(cards[0].key).toBe(deckLayout.key);
    // Every standard kind except the deck's own "title" appears as a fallback.
    const fallbackKinds = cards.slice(1).map((l) => l.kind);
    expect(fallbackKinds).not.toContain("title");
    expect(fallbackKinds).toContain("twoContent");
    expect(cards.slice(1).every((l) => l.key.startsWith("standard:"))).toBe(true);
  });
});

describe("open modes", () => {
  it("openForChange seeds the current slide's layoutKey", () => {
    const { c } = setup({ id: "s", frames: [], layoutKey: "standard:titleAndBody" }, [titleAndBody]);
    c.openForChange();
    expect(c.open.value).toBe(true);
    expect(c.mode.value).toBe("change");
    expect(c.currentKey.value).toBe("standard:titleAndBody");
  });
  it("openForNew starts on the first card", () => {
    const { c } = setup({ id: "s", frames: [] }, [titleAndBody]);
    c.openForNew();
    expect(c.mode.value).toBe("new");
    expect(c.currentKey.value).toBe(titleAndBody.key);
  });
});

describe("apply (change mode)", () => {
  it("records one entry, mutates, commits, and closes when everything placed", () => {
    const slide: Slide = {
      id: "s",
      frames: [
        { id: "t", kind: "text", role: "title", phType: "title", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, paragraphs: [] },
      ],
    };
    const { c, record, commit } = setup(slide, [titleAndBody]);
    c.openForChange();
    c.choose(titleAndBody.key);
    c.apply();
    expect(record).toHaveBeenCalledWith("Change slide layout");
    expect(commit).toHaveBeenCalledTimes(1);
    expect(slide.layoutKey).toBe(titleAndBody.key);
    expect(c.open.value).toBe(false); // nothing kept in place → auto-closed
  });

  it("surfaces the report and stays open when a frame is kept in place", () => {
    const slide: Slide = {
      id: "s",
      frames: [
        { id: "t", kind: "text", role: "title", phType: "title", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, paragraphs: [] },
        { id: "img", kind: "image", mediaId: "m1", xCm: 5, yCm: 5, widthCm: 3, heightCm: 3 },
      ],
    };
    const { c } = setup(slide, [titleAndBody]);
    c.openForChange();
    c.choose(titleAndBody.key);
    c.apply();
    expect(c.report.value?.keptInPlace).toContain("img");
    expect(c.open.value).toBe(true);
    c.dismissReport();
    expect(c.report.value).toBeNull();
    expect(c.open.value).toBe(false);
  });

  it("adopts a synthesised fallback into the deck catalogue before applying", () => {
    const slide: Slide = { id: "s", frames: [] };
    const { model, c } = setup(slide, []);
    c.openForChange();
    c.choose(standardLayoutKey("sectionHeader"));
    c.apply();
    expect((model.value.layouts ?? []).some((l) => l.key === standardLayoutKey("sectionHeader"))).toBe(true);
    expect(slide.layoutKey).toBe(standardLayoutKey("sectionHeader"));
  });
});

describe("apply (new mode) + reset", () => {
  it("new mode delegates to addSlide with the chosen key and closes", () => {
    const { c, addSlide } = setup({ id: "s", frames: [] }, [titleAndBody]);
    c.openForNew();
    c.choose(titleAndBody.key);
    c.apply();
    expect(addSlide).toHaveBeenCalledWith(0, titleAndBody.key);
    expect(c.open.value).toBe(false);
  });

  it("canReset reflects a resolvable layout, and resetSlide records + commits", () => {
    const slide: Slide = {
      id: "s",
      layoutKey: titleAndBody.key,
      frames: [
        { id: "t", kind: "text", role: "title", phType: "title", xCm: 9, yCm: 9, widthCm: 1, heightCm: 1, paragraphs: [] },
      ],
    };
    const { c, record, commit } = setup(slide, [titleAndBody]);
    expect(c.canReset.value).toBe(true);
    c.resetSlide();
    expect(record).toHaveBeenCalledWith("Reset slide");
    expect(commit).toHaveBeenCalled();
    expect((slide.frames[0] as { xCm: number }).xCm).toBeCloseTo(titleAndBody.slots[0].xCm, 3);
  });

  it("canReset is false when the slide has no resolvable layout", () => {
    const { c } = setup({ id: "s", frames: [] }, []);
    expect(c.canReset.value).toBe(false);
  });
});
