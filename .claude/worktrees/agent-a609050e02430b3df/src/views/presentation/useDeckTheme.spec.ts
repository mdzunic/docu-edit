// SPEC-094 AC: the deck theme controller — gallery order and marking, apply as
// one "Apply theme" undo entry, hover preview never persisting, customise +
// reset, and the odp/opt-out notes.
import { beforeEach, describe, expect, it, vi } from "vitest";
import { shallowRef } from "vue";
import { BUILT_IN_THEMES, DEFAULT_THEME } from "../../lib/core/theme";
import type { PresentationModel, Slide } from "../../lib/presentation/model";
import { useDeckTheme, type DeckThemeDeps } from "./useDeckTheme";

function slide(overrides: Partial<Slide> = {}): Slide {
  return { id: crypto.randomUUID(), frames: [], ...overrides };
}

function harness(overrides: Partial<PresentationModel> = {}) {
  const model = shallowRef<PresentationModel | null>({
    kind: "presentation",
    id: "d",
    title: "Deck",
    format: "pptx",
    slides: [slide()],
    slideWidthCm: 25,
    slideHeightCm: 14,
    theme: { ...structuredClone(DEFAULT_THEME), partName: "ppt/theme/theme1.xml" },
    ...overrides,
  });
  const record = vi.fn();
  const commit = vi.fn();
  const deps: DeckThemeDeps = { model, record, commit };
  return { model, record, commit, controller: useDeckTheme(deps) };
}

beforeEach(() => {
  localStorage.clear();
});

describe("open state (the presentationThemeOpen preference)", () => {
  it("starts closed and toggles through the preference", () => {
    const { controller } = harness();
    expect(controller.open.value).toBe(false);
    controller.toggle();
    expect(controller.open.value).toBe(true);
    controller.toggle();
    expect(controller.open.value).toBe(false);
  });
});

describe("gallery (AC-41/42)", () => {
  it("lists the current theme first and marked, then the six built-ins", () => {
    const { controller } = harness();
    const cards = controller.gallery.value;
    expect(cards).toHaveLength(7);
    expect(cards[0].current).toBe(true);
    expect(cards[0].name).toBe("Office");
    expect(cards.slice(1).map((card) => card.name)).toEqual(BUILT_IN_THEMES.map((theme) => theme.name));
    expect(cards.slice(1).every((card) => !card.current)).toBe(true);
  });

  it("each card draws its own six accents", () => {
    const { controller } = harness();
    expect(controller.gallery.value[2].accents).toEqual([
      BUILT_IN_THEMES[1].colors.accent1,
      BUILT_IN_THEMES[1].colors.accent2,
      BUILT_IN_THEMES[1].colors.accent3,
      BUILT_IN_THEMES[1].colors.accent4,
      BUILT_IN_THEMES[1].colors.accent5,
      BUILT_IN_THEMES[1].colors.accent6,
    ]);
  });
});

describe("apply (AC-34)", () => {
  it("applies the selected built-in as ONE 'Apply theme' undo entry", () => {
    const { controller, model, record, commit } = harness();
    controller.select("builtin-1");
    controller.apply();
    expect(record).toHaveBeenCalledTimes(1);
    expect(record).toHaveBeenCalledWith("Apply theme");
    expect(commit).toHaveBeenCalledTimes(1);
    expect(model.value?.theme?.name).toBe("Facet");
    expect(model.value?.theme?.dirty).toBe(true);
    expect(model.value?.theme?.partName).toBe("ppt/theme/theme1.xml");
  });

  it("re-applying the current card is a no-op", () => {
    const { controller, record } = harness();
    controller.apply("current");
    expect(record).not.toHaveBeenCalled();
  });
});

describe("hover preview (never persisted)", () => {
  it("overrides displayTheme while hovering and clears on leave", () => {
    const { controller, model, commit } = harness();
    expect(controller.displayTheme.value?.name).toBe("Office");
    controller.hoverPreview("builtin-2");
    expect(controller.displayTheme.value?.name).toBe("Ion");
    expect(model.value?.theme?.name).toBe("Office"); // the model never moved
    expect(commit).not.toHaveBeenCalled();
    controller.clearPreview();
    expect(controller.displayTheme.value?.name).toBe("Office");
  });
});

describe("customise + reset (AC-43)", () => {
  it("edits a slot and a font, then reset restores the pristine theme", () => {
    const { controller, model, record } = harness();
    controller.setSchemeColor("accent1", "#123456");
    expect(model.value?.theme?.colors.accent1).toBe("#123456");
    expect(controller.isCustom.value).toBe(true);
    controller.setFont("major", "Georgia");
    expect(model.value?.theme?.fonts.major).toBe("Georgia");
    expect(controller.canReset.value).toBe(true);
    controller.reset();
    expect(record).toHaveBeenLastCalledWith("Reset theme");
    expect(model.value?.theme?.colors.accent1).toBe(DEFAULT_THEME.colors.accent1);
    expect(model.value?.theme?.fonts.major).toBe(DEFAULT_THEME.fonts.major);
    expect(controller.canReset.value).toBe(false);
  });

  it("rejects a malformed colour and a blank font", () => {
    const { controller, record } = harness();
    controller.setSchemeColor("accent1", "red");
    controller.setFont("minor", "   ");
    expect(record).not.toHaveBeenCalled();
  });
});

describe("honesty notes (AC-24/32)", () => {
  it("shows the loss note for odp decks only", () => {
    expect(harness({ format: "odp" }).controller.lossNote.value).toMatch(/theme/i);
    expect(harness().controller.lossNote.value).toBeNull();
  });

  it("names slides whose p:clrMapOvr cannot be parsed as opted out", () => {
    const { controller } = harness({
      slides: [
        slide(),
        slide({ preserved: { "p:clrMapOvr": "<p:clrMapOvr><p:garbage/></p:clrMapOvr>" } }),
        slide({ preserved: { "p:clrMapOvr": "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>" } }),
      ],
    });
    expect(controller.optedOutSlides.value).toEqual([2]);
  });
});
