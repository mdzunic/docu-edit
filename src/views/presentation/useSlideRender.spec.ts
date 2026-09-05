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

function withMedia() {
  const model = shallowRef<PresentationModel | null>(createBlankPresentation());
  model.value!.media = [{ mediaId: "ppt/media/media1.mp4", contentType: "video/mp4", data: "AAAA" }];
  const render = useSlideRender({ model, currentIndex: shallowRef(0), recomputeScale: () => {} });
  return { model, render };
}

describe("useSlideRender commit media gate (SPEC-095 AC-12/14)", () => {
  it("does NOT re-slice model.media on a plain commit (a nudge/keystroke)", () => {
    const { model, render } = withMedia();
    const before = model.value!.media;
    render.commit();
    expect(model.value!.media).toBe(before);
  });

  it("re-slices model.media after bumpMedia (a just-inserted image or clip renders)", () => {
    const { model, render } = withMedia();
    const before = model.value!.media;
    render.bumpMedia();
    render.commit();
    expect(model.value!.media).not.toBe(before);
    expect(model.value!.media).toEqual(before);
  });

  it("clears the gate after each commit (the next plain commit does not re-slice)", () => {
    const { model, render } = withMedia();
    render.bumpMedia();
    render.commit();
    const afterBump = model.value!.media;
    render.commit();
    expect(model.value!.media).toBe(afterBump);
  });
});

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
