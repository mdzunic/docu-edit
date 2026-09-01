// SPEC-096: multi-selection semantics — this file's first spec. The set is
// z-order stable and top-level only; selectedFrameId stays the single-frame
// Ref (holding selectedFrameIds[0], the anchor) so useAnimations' and
// useSlideTextEditing's deps are untouched; a slide change clears everything.
import { computed, shallowRef } from "vue";
import { describe, expect, it } from "vitest";
import type { Slide } from "../../lib/presentation/model";
import { useFrameSelection } from "./useFrameSelection";

function slideWith(ids: string[]): Slide {
  return {
    id: `slide-${ids.join("-")}`,
    frames: ids.map((id) => ({
      id,
      kind: "shape" as const,
      geometry: "rect" as const,
      xCm: 1,
      yCm: 1,
      widthCm: 2,
      heightCm: 2,
    })),
  };
}

function setup(ids = ["a", "b", "c"]) {
  const slide = shallowRef<Slide | null>(slideWith(ids));
  const currentSlide = computed(() => slide.value);
  const selection = useFrameSelection({ currentSlide });
  return { slide, selection };
}

describe("useFrameSelection multi-selection (SPEC-096 AC-1)", () => {
  it("select() collapses to a single frame; selectedFrameId mirrors [0]", () => {
    const { selection } = setup();
    selection.select("b");
    expect(selection.selectedFrameIds.value).toEqual(["b"]);
    expect(selection.selectedFrameId.value).toBe("b");
    expect(selection.selectedFrame.value?.id).toBe("b");
    expect(selection.anchorFrame.value?.id).toBe("b");
  });

  it("toggle adds and removes frames, keeping the set z-order stable", () => {
    const { selection } = setup();
    selection.select("c");
    selection.toggle("a");
    // z-order (slide order) stable: a before c.
    expect(selection.selectedFrameIds.value).toEqual(["a", "c"]);
    selection.toggle("a");
    expect(selection.selectedFrameIds.value).toEqual(["c"]);
    expect(selection.isSelected("c")).toBe(true);
    expect(selection.isSelected("a")).toBe(false);
  });

  it("selectMany replaces (or with add=true extends) the set, deduplicated in z-order", () => {
    const { selection } = setup();
    selection.selectMany(["c", "a"]);
    expect(selection.selectedFrameIds.value).toEqual(["a", "c"]);
    selection.selectMany(["b"], true);
    expect(selection.selectedFrameIds.value).toEqual(["a", "b", "c"]);
    selection.selectMany(["nope"]);
    expect(selection.selectedFrameIds.value).toEqual([]);
  });

  it("selectAll selects every frame on the slide", () => {
    const { selection } = setup();
    selection.selectAll();
    expect(selection.selectedFrameIds.value).toEqual(["a", "b", "c"]);
  });

  it("writing selectedFrameId collapses the whole set (the SPEC-081 contract)", () => {
    const { selection } = setup();
    selection.selectAll();
    selection.selectedFrameId.value = "b";
    expect(selection.selectedFrameIds.value).toEqual(["b"]);
    selection.selectedFrameId.value = null;
    expect(selection.selectedFrameIds.value).toEqual([]);
  });

  it("a slide change clears the whole set", async () => {
    const { slide, selection } = setup();
    selection.selectAll();
    slide.value = slideWith(["x", "y"]);
    // The watch is synchronous enough after a tick.
    await Promise.resolve();
    expect(selection.selectedFrameIds.value).toEqual([]);
  });

  it("cycle still walks single frames in z-order", () => {
    const { selection } = setup();
    selection.cycle(1);
    expect(selection.selectedFrameIds.value).toEqual(["a"]);
    selection.cycle(1);
    expect(selection.selectedFrameIds.value).toEqual(["b"]);
    selection.cycle(-1);
    expect(selection.selectedFrameIds.value).toEqual(["a"]);
  });
});
