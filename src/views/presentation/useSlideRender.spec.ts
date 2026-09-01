// SPEC-096 — the shallowRef re-render trick must reach INSIDE groups: group
// commands (rotate, resize) mutate children in place, so refresh() has to
// re-identify the child objects too or the stage renders them stale (a
// shallow top-level spread let exactly that through QA once).
import { shallowRef } from "vue";
import { describe, expect, it } from "vitest";
import { createBlankPresentation } from "../../lib/presentation/slides";
import type { GroupFrame, PresentationModel } from "../../lib/presentation/model";
import { useSlideRender } from "./useSlideRender";

function setup() {
  const model = shallowRef<PresentationModel | null>(createBlankPresentation());
  const group: GroupFrame = {
    id: "g1",
    xCm: 2,
    yCm: 2,
    widthCm: 8,
    heightCm: 6,
    kind: "group",
    children: [
      { id: "c1", xCm: 3, yCm: 3, widthCm: 2, heightCm: 1, kind: "shape", geometry: "rect" },
      {
        id: "g2",
        xCm: 6,
        yCm: 4,
        widthCm: 3,
        heightCm: 3,
        kind: "group",
        children: [{ id: "c2", xCm: 6.5, yCm: 4.5, widthCm: 1, heightCm: 1, kind: "shape", geometry: "rect" }],
      },
    ],
  };
  model.value!.slides[0].frames.push(group);
  const render = useSlideRender({
    model,
    currentIndex: shallowRef(0),
    recomputeScale: () => {},
  });
  return { model, render };
}

describe("useSlideRender refresh (SPEC-096)", () => {
  it("re-identifies the slide, its frames AND group descendants, preserving values", () => {
    const { model, render } = setup();
    const slide0 = model.value!.slides[0];
    const group0 = slide0.frames.find((f) => f.id === "g1") as GroupFrame;
    const child0 = group0.children[0];
    const nested0 = group0.children[1] as GroupFrame;
    const nestedChild0 = nested0.children[0];

    render.refresh();

    const slide1 = model.value!.slides[0];
    const group1 = slide1.frames.find((f) => f.id === "g1") as GroupFrame;
    const nested1 = group1.children[1] as GroupFrame;
    expect(slide1).not.toBe(slide0);
    expect(group1).not.toBe(group0);
    expect(group1.children[0]).not.toBe(child0);
    expect(nested1).not.toBe(nested0);
    expect(nested1.children[0]).not.toBe(nestedChild0);
    // Values ride along untouched.
    expect(group1.children[0]).toEqual(child0);
    expect(nested1.children[0]).toEqual(nestedChild0);
  });
});
