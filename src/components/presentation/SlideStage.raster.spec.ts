// SPEC-104 — the export sheet's render knobs on the shared slide stage:
// `rasterScale` reaches paintShapeIntoCanvas as its existing fifth (zoom)
// argument, so an exported shape's canvas carries devicePixelRatio ×
// rasterScale pixels; `flatRotation` drops the frame's CSS rotation (its text
// prints unrotated at the axis-aligned box) and bakes a shape's rotation into
// its canvas instead. Both default off, leaving the editor render unchanged.
import { afterEach, describe, expect, it, vi } from "vitest";
import { mount } from "@vue/test-utils";

const paints = vi.hoisted(() => [] as { zoom: unknown; rotationDeg: unknown }[]);

vi.mock("../../lib/core/shapeCanvas", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../lib/core/shapeCanvas")>();
  return {
    ...actual,
    paintShapeIntoCanvas: (_canvas: HTMLCanvasElement, shape: { rotationDeg?: number }, _w: number, _h: number, zoom?: number) => {
      paints.push({ zoom, rotationDeg: shape.rotationDeg });
    },
  };
});

// jsdom plays no media; let the stage believe the clip is playable so the
// idle poster and its play badge render as they do in a browser.
vi.mock("./mediaRender", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./mediaRender")>();
  return { ...actual, createCanPlayProbe: () => () => true };
});

const { default: SlideStage } = await import("./SlideStage.vue");

const slide = {
  id: "s",
  frames: [
    { id: "sh", kind: "shape", geometry: "rect", xCm: 1, yCm: 1, widthCm: 4, heightCm: 2, rotationDeg: 30 },
    { id: "tx", kind: "text", xCm: 6, yCm: 1, widthCm: 4, heightCm: 2, rotationDeg: 30, paragraphs: [{ runs: [{ text: "Hi" }] }] },
  ],
};

function mountStage(props: Record<string, unknown> = {}) {
  return mount(SlideStage, { props: { slide, widthCm: 33.867, heightCm: 19.05, interactive: false, ...props } });
}

afterEach(() => {
  paints.length = 0;
});

describe("SlideStage rasterScale / flatRotation (SPEC-104)", () => {
  it("defaults leave the editor render as it was: zoom 1, CSS rotation, an unrotated canvas", () => {
    const wrapper = mountStage();
    expect(paints).toEqual([{ zoom: 1, rotationDeg: undefined }]);
    const frames = wrapper.findAll(".frame").map((frame) => (frame.element as HTMLElement).style.transform);
    expect(frames).toEqual(["rotate(30deg)", "rotate(30deg)"]);
  });

  it("forwards rasterScale to paintShapeIntoCanvas's fifth argument", () => {
    mountStage({ rasterScale: 2 });
    expect(paints).toEqual([{ zoom: 2, rotationDeg: undefined }]);
  });

  it("flatRotation renders frames at their axis-aligned box and bakes shape rotation into the canvas", () => {
    const wrapper = mountStage({ rasterScale: 2, flatRotation: true });
    expect(paints).toEqual([{ zoom: 2, rotationDeg: 30 }]);
    const frames = wrapper.findAll(".frame").map((frame) => (frame.element as HTMLElement).style.transform);
    expect(frames).toEqual(["", ""]);
  });

  it("screen-only chrome carries .print-exclude: the placeholder prompt", () => {
    const prompt = mount(SlideStage, {
      props: {
        slide: { id: "p", frames: [{ id: "t", kind: "text", role: "title", phType: "title", xCm: 1, yCm: 1, widthCm: 4, heightCm: 2, paragraphs: [] }] },
        widthCm: 33.867,
        heightCm: 19.05,
      },
    });
    expect(prompt.get('[data-testid="slide-placeholder-prompt"]').classes()).toContain("print-exclude");
  });

  it("a media frame exports as its poster image, its play badge excluded (SPEC-095's MediaFrame)", () => {
    const media = {
      id: "m",
      kind: "media",
      media: "video",
      mediaId: "clip.mp4",
      posterMediaId: "poster.png",
      contentType: "video/mp4",
      xCm: 1,
      yCm: 1,
      widthCm: 16,
      heightCm: 9,
    };
    const wrapper = mount(SlideStage, {
      props: {
        slide: { id: "v", frames: [media] },
        widthCm: 33.867,
        heightCm: 19.05,
        interactive: false,
        media: [{ mediaId: "poster.png", contentType: "image/png", data: "AA==" }],
      },
    });
    expect(wrapper.find("img.media-poster").exists()).toBe(true);
    expect(wrapper.get('[data-testid="media-play"]').classes()).toContain("print-exclude");
  });
});
