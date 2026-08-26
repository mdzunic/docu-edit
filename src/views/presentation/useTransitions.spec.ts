import { beforeEach, describe, expect, it, vi } from "vitest";
import { computed, ref } from "vue";
import type { PresentationModel, Slide } from "../../lib/presentation/model";
import { useTransitions, type TransitionsDeps } from "./useTransitions";

function slide(overrides: Partial<Slide> = {}): Slide {
  return { id: crypto.randomUUID(), frames: [], ...overrides };
}

function harness(slides: Slide[], format: "pptx" | "odp" = "pptx") {
  const model = ref<PresentationModel | null>({
    kind: "presentation",
    id: "d",
    title: "Deck",
    format,
    slides,
    slideWidthCm: 25,
    slideHeightCm: 14,
  });
  const currentIndex = ref(0);
  const currentSlide = computed(() => model.value?.slides[currentIndex.value] ?? null);
  const record = vi.fn();
  const commit = vi.fn();
  const deps: TransitionsDeps = { model, currentSlide, currentIndex, record, commit };
  return { deps, model, currentIndex, record, commit, controller: useTransitions(deps) };
}

beforeEach(() => {
  localStorage.clear();
});

describe("two None states (AC-3)", () => {
  it("a plain slide shows None and no preserved note", () => {
    const { controller } = harness([slide()]);
    expect(controller.selectedType.value).toBe("none");
    expect(controller.isPreserved.value).toBe(false);
    expect(controller.canApplyAll.value).toBe(false);
  });

  it("a preserved out-of-subset slide shows None WITH the note", () => {
    const { controller } = harness([slide({ preserved: { "p:transition": "<p:transition><p:morph/></p:transition>" } })]);
    expect(controller.selectedType.value).toBe("none");
    expect(controller.isPreserved.value).toBe(true);
    expect(controller.canApplyAll.value).toBe(false); // unset model
  });
});

describe("choosing an effect (AC-7)", () => {
  it("sets the transition, records one undo entry, and persists", () => {
    const { controller, model, record, commit } = harness([slide()]);
    controller.setEffect("fade");
    expect(model.value!.slides[0].transition).toEqual({ type: "fade" });
    expect(record).toHaveBeenCalledTimes(1);
    expect(commit).toHaveBeenCalledTimes(1);
  });

  it("keeps duration/advance fields when switching effect", () => {
    const { controller, model } = harness([slide({ transition: { type: "fade", durationMs: 1000, advanceAfterMs: 4000 } })]);
    controller.setEffect("push");
    expect(model.value!.slides[0].transition).toEqual({
      type: "push",
      direction: "left",
      durationMs: 1000,
      advanceAfterMs: 4000,
    });
  });
});

describe("preserved-replacement confirmation (AC-8)", () => {
  it("gates the first change and applies only on confirm", () => {
    const { controller, model, record, commit } = harness([
      slide({ preserved: { "p:transition": "<p:transition><p:morph/></p:transition>" } }),
    ]);
    controller.setEffect("fade");
    // Nothing changed yet — the confirmation is pending.
    expect(controller.pendingReplace.value).not.toBeNull();
    expect(model.value!.slides[0].transition).toBeUndefined();
    expect(model.value!.slides[0].preserved?.["p:transition"]).toBeDefined();
    expect(record).not.toHaveBeenCalled();

    controller.confirmReplace();
    expect(controller.pendingReplace.value).toBeNull();
    expect(model.value!.slides[0].transition).toEqual({ type: "fade" });
    expect(model.value!.slides[0].preserved).toBeUndefined();
    expect(record).toHaveBeenCalledTimes(1);
    expect(commit).toHaveBeenCalledTimes(1);
  });

  it("cancelling changes nothing", () => {
    const { controller, model, record } = harness([
      slide({ preserved: { "p:transition": "<p:transition><p:morph/></p:transition>" } }),
    ]);
    controller.setEffect("none");
    controller.cancelReplace();
    expect(controller.pendingReplace.value).toBeNull();
    expect(model.value!.slides[0].transition).toBeUndefined();
    expect(model.value!.slides[0].preserved?.["p:transition"]).toBeDefined();
    expect(record).not.toHaveBeenCalled();
  });
});

describe("Apply to all (AC-9/AC-10)", () => {
  it("is disabled while the current slide's transition is unset (AC-9)", () => {
    const { controller } = harness([slide(), slide()]);
    expect(controller.canApplyAll.value).toBe(false);
    controller.applyToAll(); // no-op
    expect(controller.pendingApplyAllCount.value).toBe(0);
  });

  it("copies to every slide with no confirmation when none are preserved", () => {
    const { controller, model, record, commit } = harness([
      slide({ transition: { type: "wipe", direction: "up", durationMs: 900 } }),
      slide(),
      slide(),
    ]);
    controller.applyToAll();
    expect(controller.pendingApplyAllCount.value).toBe(0);
    for (const s of model.value!.slides) {
      expect(s.transition).toEqual({ type: "wipe", direction: "up", durationMs: 900 });
    }
    expect(record).toHaveBeenCalledTimes(1); // one undo entry
    expect(commit).toHaveBeenCalledTimes(1); // persists once
  });

  it("gates behind a batched confirmation naming the preserved count, then applies (AC-10)", () => {
    const { controller, model, record, commit } = harness([
      slide({ transition: { type: "fade" } }),
      slide({ preserved: { "p:transition": "<p:transition><p:morph/></p:transition>" } }),
      slide({ preserved: { "p:transition": "<p:transition><p:zoom/></p:transition>" } }),
    ]);
    controller.applyToAll();
    expect(controller.pendingApplyAllCount.value).toBe(2);
    // Nothing applied until confirmed.
    expect(model.value!.slides[1].transition).toBeUndefined();
    expect(record).not.toHaveBeenCalled();

    controller.confirmApplyAll();
    expect(controller.pendingApplyAllCount.value).toBe(0);
    for (const s of model.value!.slides) {
      expect(s.transition).toEqual({ type: "fade" });
      expect(s.preserved).toBeUndefined();
    }
    expect(record).toHaveBeenCalledTimes(1);
    expect(commit).toHaveBeenCalledTimes(1);
  });

  it("cancelling the batched confirmation changes nothing (AC-10d)", () => {
    const { controller, model, record } = harness([
      slide({ transition: { type: "fade" } }),
      slide({ preserved: { "p:transition": "<p:transition><p:morph/></p:transition>" } }),
    ]);
    controller.applyToAll();
    controller.cancelApplyAll();
    expect(controller.pendingApplyAllCount.value).toBe(0);
    expect(model.value!.slides[1].transition).toBeUndefined();
    expect(model.value!.slides[1].preserved?.["p:transition"]).toBeDefined();
    expect(record).not.toHaveBeenCalled();
  });
});

describe("field edits", () => {
  it("duration is in seconds and clamps to the model bounds", () => {
    const { controller, model } = harness([slide({ transition: { type: "fade" } })]);
    controller.setDurationSeconds(2); // 2s -> 2000ms
    expect(model.value!.slides[0].transition?.durationMs).toBe(2000);
    controller.setDurationSeconds(0.01); // clamps up to 100ms
    expect(model.value!.slides[0].transition?.durationMs).toBe(100);
  });

  it("advance-on-click stores only the false case", () => {
    const { controller, model } = harness([slide({ transition: { type: "fade" } })]);
    controller.setAdvanceOnClick(false);
    expect(model.value!.slides[0].transition?.advanceOnClick).toBe(false);
    controller.setAdvanceOnClick(true);
    expect(model.value!.slides[0].transition?.advanceOnClick).toBeUndefined();
  });

  it("auto-advance toggles the field with a seeded default", () => {
    const { controller, model } = harness([slide({ transition: { type: "fade" } })]);
    controller.setAutoAdvanceEnabled(true);
    expect(model.value!.slides[0].transition?.advanceAfterMs).toBe(3000);
    controller.setAutoAdvanceSeconds(6);
    expect(model.value!.slides[0].transition?.advanceAfterMs).toBe(6000);
    controller.setAutoAdvanceEnabled(false);
    expect(model.value!.slides[0].transition?.advanceAfterMs).toBeUndefined();
  });
});

describe("preview (AC-19)", () => {
  it("previews from the previous slide, or from nothing on slide 1", () => {
    const { controller, model, currentIndex } = harness([
      slide({ transition: { type: "fade" } }),
      slide({ transition: { type: "push", direction: "left" } }),
    ]);
    controller.preview();
    expect(controller.previewState.value?.outgoing).toBeUndefined(); // slide 1 → from black
    controller.onPreviewDone();
    expect(controller.previewState.value).toBeNull();

    currentIndex.value = 1;
    controller.preview();
    expect(controller.previewState.value?.outgoing).toBe(model.value!.slides[0]);
    expect(controller.previewState.value?.incoming.transition).toEqual({ type: "push", direction: "left" });
  });
});
