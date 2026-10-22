// SPEC-090 — the deck-setup controller: preset/orientation/custom interplay,
// unit-aware size parsing and rejection, one undo entry per Apply, and Cancel
// (drafts are local, so nothing changes until Apply).
import { describe, expect, it, vi } from "vitest";
import { ref } from "vue";
import { useDeckSetup } from "./useDeckSetup";
import type { PresentationModel, Slide, SlideFrame } from "../../lib/presentation/model";

function frame(): SlideFrame {
  return { id: "t1", kind: "text", xCm: 2, yCm: 3, widthCm: 10, heightCm: 4, paragraphs: [{ runs: [{ text: "hi", sizePt: 20 }] }] };
}

function harness(over: Partial<PresentationModel> = {}) {
  const slide: Slide = { id: "s1", frames: [frame()] };
  const model = ref<PresentationModel | null>({
    kind: "presentation",
    id: "d1",
    title: "Deck",
    format: "pptx",
    slides: [slide],
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
    ...over,
  });
  const record = vi.fn();
  const commit = vi.fn();
  const controller = useDeckSetup({ model, currentIndex: ref(0), record, commit, reportError: vi.fn() });
  return { controller, model, record, commit };
}

describe("useDeckSetup open + init", () => {
  it("preselects the preset, orientation and fields from the deck", () => {
    const { controller } = harness();
    controller.open();
    expect(controller.isOpen.value).toBe(true);
    expect(controller.preset.value).toBe("16:9");
    expect(controller.orientation.value).toBe("landscape");
    expect(controller.widthText.value).toBe("33.87");
    expect(controller.heightText.value).toBe("19.05");
    expect(controller.firstSlideNumber.value).toBe(1);
    expect(controller.activeTab.value).toBe("size");
  });

  it("shows the deck's stored first slide number", () => {
    const { controller } = harness({ firstSlideNumber: 5 });
    controller.open();
    expect(controller.firstSlideNumber.value).toBe(5);
  });
});

describe("useDeckSetup size interplay", () => {
  it("picking a preset refills width and height", () => {
    const { controller } = harness();
    controller.open();
    controller.setPreset("4:3");
    expect(controller.preset.value).toBe("4:3");
    expect(controller.widthText.value).toBe("25.4");
    expect(controller.heightText.value).toBe("19.05");
  });

  it("orientation swaps width and height", () => {
    const { controller } = harness();
    controller.open();
    controller.setOrientation("portrait");
    expect(controller.widthText.value).toBe("19.05");
    expect(controller.heightText.value).toBe("33.87");
  });

  it("typing in a field switches the preset to Custom", () => {
    const { controller } = harness();
    controller.open();
    controller.setWidth("20");
    expect(controller.preset.value).toBe("Custom");
  });

  it("parses other units (10in) and clears the error", () => {
    const { controller } = harness();
    controller.open();
    controller.setWidth("10in");
    expect(controller.sizeError.value).toBe("");
    expect(controller.canApply.value).toBe(true);
  });

  it("rejects an out-of-range size, showing the error and disabling Apply", () => {
    const { controller } = harness();
    controller.open();
    controller.setWidth("500"); // > 142.24 cm
    expect(controller.sizeError.value).not.toBe("");
    expect(controller.canApply.value).toBe(false);
  });

  it("rejects an unparseable size", () => {
    const { controller } = harness();
    controller.open();
    controller.setHeight("abc");
    expect(controller.canApply.value).toBe(false);
  });
});

describe("useDeckSetup first slide number", () => {
  it("disables Apply for a non-integer or out-of-range value", () => {
    const { controller } = harness();
    controller.open();
    controller.setFirstSlideNumber(0);
    expect(controller.canApply.value).toBe(false);
    controller.setFirstSlideNumber(10000);
    expect(controller.canApply.value).toBe(false);
    controller.setFirstSlideNumber(2.5);
    expect(controller.canApply.value).toBe(false);
    controller.setFirstSlideNumber(3);
    expect(controller.canApply.value).toBe(true);
  });
});

describe("useDeckSetup Apply", () => {
  it("size Apply records one 'Deck setup' undo entry and commits once", () => {
    const { controller, model, record, commit } = harness();
    controller.open();
    controller.setPreset("4:3");
    controller.setRescaleMode("scale");
    controller.apply();
    expect(record).toHaveBeenCalledTimes(1);
    expect(record).toHaveBeenCalledWith("Deck setup");
    expect(commit).toHaveBeenCalledTimes(1);
    expect(model.value!.slideWidthCm).toBe(25.4);
    expect(model.value!.deckDirty).toBe(true);
    expect(model.value!.slides[0].dirty).toBe(true);
    expect(controller.isOpen.value).toBe(false);
  });

  it("keep mode leaves geometry untouched and warns on overflow", () => {
    const { controller, model } = harness();
    controller.open();
    // Shrink to the minimum: the frame sits at y=3, below a 2.54cm-tall slide.
    controller.setWidth("2.54");
    controller.setHeight("2.54");
    controller.setRescaleMode("keep");
    expect(controller.overflowCount.value).toBeGreaterThan(0);
    controller.apply();
    expect(model.value!.slides[0].dirty).toBeUndefined();
    expect(model.value!.slides[0].frames[0].widthCm).toBe(10);
  });

  it("properties Apply records one 'Deck properties' entry and writes the draft back", () => {
    const { controller, model, record } = harness();
    controller.open();
    controller.setTab("properties");
    controller.setProperty("author", "Ada");
    controller.setProperty("subject", "Q3");
    controller.apply();
    expect(record).toHaveBeenCalledWith("Deck properties");
    expect(model.value!.properties?.author).toBe("Ada");
    expect(model.value!.properties?.subject).toBe("Q3");
    expect(model.value!.deckDirty).toBe(true);
  });

  it("Cancel changes nothing", () => {
    const { controller, model, record, commit } = harness();
    controller.open();
    controller.setPreset("4:3");
    controller.close();
    expect(record).not.toHaveBeenCalled();
    expect(commit).not.toHaveBeenCalled();
    expect(model.value!.slideWidthCm).toBe(33.867);
    expect(model.value!.deckDirty).toBeUndefined();
  });
});

describe("useDeckSetup odp hint", () => {
  it("isOdp reflects the deck format", () => {
    expect(harness({ format: "odp" }).controller.isOdp.value).toBe(true);
    expect(harness({ format: "pptx" }).controller.isOdp.value).toBe(false);
  });
});
