// SPEC-093: the Header & Footer controller. Tested without mounting — refs work
// standalone (the useDeckSetup.spec pattern).
import { computed, ref } from "vue";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { PresentationModel, Slide } from "../../lib/presentation/model";
import { setClock } from "../../lib/clock";
import { useHeaderFooter } from "./useHeaderFooter";

function slide(overrides: Partial<Slide> = {}): Slide {
  return { id: crypto.randomUUID(), frames: [], ...overrides };
}

function harness(model: PresentationModel, index = 0) {
  const modelRef = ref<PresentationModel | null>(model);
  const currentIndex = ref(index);
  const currentSlide = computed(() => modelRef.value?.slides[currentIndex.value] ?? null);
  const record = vi.fn();
  const commit = vi.fn();
  const controller = useHeaderFooter({ model: modelRef, currentSlide, currentIndex, record, commit });
  return { controller, record, commit, model };
}

function baseModel(overrides: Partial<PresentationModel> = {}): PresentationModel {
  return {
    kind: "presentation",
    id: "deck",
    title: "Deck",
    format: "pptx",
    slides: [slide(), slide()],
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
    ...overrides,
  };
}

afterEach(() => setClock(null));

describe("useHeaderFooter", () => {
  it("opens seeding the draft from the deck values + current slide", () => {
    const model = baseModel({ headerFooter: { footerText: "Confidential", dateFormat: "yyyy-MM-dd", showOnTitleSlide: false } });
    model.slides[1].fields = [{ kind: "footer", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1 }];
    const { controller } = harness(model, 1);
    expect(controller.open.value).toBe(false);
    controller.show();
    expect(controller.open.value).toBe(true);
    expect(controller.footer.value).toBe(true);
    expect(controller.footerText.value).toBe("Confidential");
    expect(controller.dateFormat.value).toBe("yyyy-MM-dd");
    expect(controller.showOnTitleSlide.value).toBe(false);
    expect(controller.slideNumber.value).toBe(false);
  });

  it("previews the six date formats against today's instant", () => {
    setClock(new Date(2026, 4, 14));
    const { controller } = harness(baseModel());
    const options = controller.dateFormatOptions.value;
    expect(options).toHaveLength(6);
    expect(options[0]).toEqual({ value: "d MMMM yyyy", label: "14 May 2026" });
    expect(options.find((o) => o.value === "yyyy-MM-dd")?.label).toBe("2026-05-14");
  });

  it("flags an out-of-subset date on the current slide", () => {
    const model = baseModel();
    model.slides[0].fields = [{ kind: "date", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, wireType: "datetime13" }];
    const { controller } = harness(model, 0);
    expect(controller.datePreserved.value).toBe(true);
  });

  it("Apply records, applies to this slide, commits, and closes", () => {
    const { controller, record, commit, model } = harness(baseModel(), 1);
    controller.show();
    controller.footer.value = true;
    controller.footerText.value = "Draft";
    controller.slideNumber.value = true;
    controller.apply();
    expect(record).toHaveBeenCalledWith("Header and footer");
    expect(commit).toHaveBeenCalledOnce();
    expect(controller.open.value).toBe(false);
    expect(model.headerFooter?.footerText).toBe("Draft");
    expect(model.slides[0].fields).toBeUndefined();
    expect(model.slides[1].fields?.map((f) => f.kind).toSorted()).toEqual(["footer", "slideNumber"]);
    expect(model.headerFooterDirty).toBe(true);
  });

  it("Apply to all writes every slide with the 'apply to all' undo label", () => {
    const { controller, record, model } = harness(baseModel(), 0);
    controller.show();
    controller.footer.value = true;
    controller.footerText.value = "Confidential";
    controller.applyToAll();
    expect(record).toHaveBeenCalledWith("Apply header and footer to all slides");
    for (const s of model.slides) {
      expect(s.fields?.map((f) => f.kind)).toEqual(["footer"]);
    }
  });

  it("Cancel changes nothing", () => {
    const { controller, record, commit, model } = harness(baseModel(), 0);
    controller.show();
    controller.footer.value = true;
    controller.footerText.value = "Nope";
    controller.cancel();
    expect(controller.open.value).toBe(false);
    expect(record).not.toHaveBeenCalled();
    expect(commit).not.toHaveBeenCalled();
    expect(model.headerFooter).toBeUndefined();
    expect(model.slides[0].fields).toBeUndefined();
  });
});
