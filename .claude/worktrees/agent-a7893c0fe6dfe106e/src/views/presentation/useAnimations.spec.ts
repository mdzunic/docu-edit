import { beforeEach, describe, expect, it, vi } from "vitest";
import { computed, ref } from "vue";
import type { PresentationModel, Slide, SlideFrame } from "../../lib/presentation/model";
import { useAnimations, type AnimationsDeps } from "./useAnimations";

function frame(id: string, kind: SlideFrame["kind"] = "text"): SlideFrame {
  const base = { id, xCm: 0, yCm: 0, widthCm: 5, heightCm: 2 };
  if (kind === "shape") return { ...base, kind: "shape", geometry: "rect" };
  if (kind === "image") return { ...base, kind: "image", mediaId: "m" };
  if (kind === "group") return { ...base, kind: "group", children: [], xml: "<g/>" };
  return { ...base, kind: "text", paragraphs: [] };
}

function harness(slide: Slide, format: "pptx" | "odp" = "pptx") {
  const model = ref<PresentationModel | null>({
    kind: "presentation",
    id: "d",
    title: "Deck",
    format,
    slides: [slide],
    slideWidthCm: 25,
    slideHeightCm: 14,
  });
  const currentIndex = ref(0);
  const currentSlide = computed(() => model.value?.slides[currentIndex.value] ?? null);
  const selectedFrameId = ref<string | null>(null);
  const record = vi.fn();
  const commit = vi.fn();
  const selectFrame = vi.fn((id: string | null) => (selectedFrameId.value = id));
  const deps: AnimationsDeps = { model, currentSlide, currentIndex, record, commit, selectedFrameId, selectFrame };
  return { deps, model, selectedFrameId, record, commit, selectFrame, c: useAnimations(deps) };
}

beforeEach(() => localStorage.clear());

describe("useAnimations", () => {
  it("Add is disabled with helper text when no frame is selected (AC-8)", () => {
    const { c } = harness({ id: "s", frames: [frame("f1")] });
    expect(c.canAdd.value).toBe(false);
    expect(c.addHelp.value).toContain("Select a frame");
  });

  it("Add is disabled for a non-animatable frame kind", () => {
    const h = harness({ id: "s", frames: [frame("g", "group")] });
    h.selectedFrameId.value = "g";
    expect(h.c.canAdd.value).toBe(false);
    expect(h.c.addHelp.value).toContain("can't be animated");
  });

  it("adds a fade-in/click animation, records + commits, selects the new row (AC-16)", () => {
    const h = harness({ id: "s", frames: [frame("f1"), frame("f2", "shape")] });
    h.selectedFrameId.value = "f1";
    h.c.add();
    expect(h.model.value!.slides[0].animations).toEqual([{ frameId: "f1", effect: "fadeIn", trigger: "click" }]);
    expect(h.record).toHaveBeenCalledTimes(1);
    expect(h.commit).toHaveBeenCalledTimes(1);
    expect(h.c.selectedIndex.value).toBe(0);
  });

  it("normalises row 1 trigger to click and locks its trigger control (AC-3)", () => {
    const h = harness({
      id: "s",
      frames: [frame("f1")],
      animations: [{ frameId: "f1", effect: "flyIn", trigger: "afterPrevious" }],
    });
    expect(h.c.rows.value[0].trigger).toBe("click");
    expect(h.c.triggerLocked.value).toBe(true);
    h.c.setTrigger("withPrevious"); // ignored while locked
    expect(h.model.value!.slides[0].animations![0].trigger).toBe("afterPrevious");
  });

  it("selecting a row highlights its target frame (AC-14)", () => {
    const h = harness({
      id: "s",
      frames: [frame("f1"), frame("f2", "shape")],
      animations: [
        { frameId: "f1", effect: "fadeIn", trigger: "click" },
        { frameId: "f2", effect: "flyIn", trigger: "click" },
      ],
    });
    h.c.selectRow(1);
    expect(h.selectFrame).toHaveBeenCalledWith("f2");
    expect(h.c.selectedIndex.value).toBe(1);
  });

  it("reorders and removes with one gesture each", () => {
    const h = harness({
      id: "s",
      frames: [frame("f1"), frame("f2", "shape")],
      animations: [
        { frameId: "f1", effect: "fadeIn", trigger: "click" },
        { frameId: "f2", effect: "flyIn", trigger: "click" },
      ],
    });
    h.c.selectRow(1);
    h.c.moveUp();
    expect(h.model.value!.slides[0].animations!.map((a) => a.frameId)).toEqual(["f2", "f1"]);
    expect(h.c.selectedIndex.value).toBe(0);
    h.c.remove();
    expect(h.model.value!.slides[0].animations!.map((a) => a.frameId)).toEqual(["f1"]);
  });

  it("gates Add behind confirm when the slide has preserved timing (AC-17)", () => {
    const h = harness({ id: "s", frames: [frame("f1")], preserved: { "p:timing": "<p:timing/>" } });
    h.selectedFrameId.value = "f1";
    h.c.add();
    // Nothing added yet — the confirm is pending.
    expect(h.c.pendingReplaceFrameId.value).toBe("f1");
    expect(h.model.value!.slides[0].animations).toBeUndefined();
    h.c.confirmReplace();
    expect(h.model.value!.slides[0].animations).toHaveLength(1);
    expect(h.model.value!.slides[0].preserved?.["p:timing"]).toBeUndefined();
  });

  it("cancel leaves preserved timing and adds nothing", () => {
    const h = harness({ id: "s", frames: [frame("f1")], preserved: { "p:timing": "<p:timing/>" } });
    h.selectedFrameId.value = "f1";
    h.c.add();
    h.c.cancelReplace();
    expect(h.c.pendingReplaceFrameId.value).toBeNull();
    expect(h.model.value!.slides[0].animations).toBeUndefined();
    expect(h.model.value!.slides[0].preserved?.["p:timing"]).toBe("<p:timing/>");
  });
});
