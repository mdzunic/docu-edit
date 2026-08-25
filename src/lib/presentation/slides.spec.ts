// SPEC-080 — the framework-free slide commands: blank-deck shape, add /
// duplicate / delete / move selection + dirty flags, and stripRetention.
import { describe, expect, it } from "vitest";
import {
  addSlide,
  createBlankPresentation,
  deleteSlide,
  duplicateSlide,
  moveSlide,
  stripRetention,
  unconvertibleFrameCount,
} from "./slides";
import { DEFAULT_PRESENTATION_TITLE, type PresentationModel, type TextFrame } from "./model";
import { readPresentationBytes } from "./open";
import { buildPptxFixtureBytes } from "./fixtures/build";

function blank(): PresentationModel {
  return createBlankPresentation();
}

describe("createBlankPresentation", () => {
  it("is a 16:9 pptx deck with one title+subtitle slide and no retained package (AC-2/3)", () => {
    const model = blank();
    expect(model.kind).toBe("presentation");
    expect(model.format).toBe("pptx");
    expect(model.title).toBe(DEFAULT_PRESENTATION_TITLE);
    expect(model.slideWidthCm).toBeCloseTo(33.867, 3);
    expect(model.slideHeightCm).toBeCloseTo(19.05, 3);
    expect(model.opaque).toBeUndefined();
    expect(model.structureDirty).toBeUndefined();
    expect(model.slides).toHaveLength(1);

    const [title, subtitle] = model.slides[0].frames as TextFrame[];
    expect(title.role).toBe("title");
    expect(title.phType).toBe("title");
    expect(subtitle.role).toBe("subtitle");
    expect(subtitle.phType).toBe("subTitle");
    // Title y 32% h 18%, subtitle y 52% h 12% (the Reference).
    expect(title.yCm).toBeCloseTo(0.32 * 19.05, 3);
    expect(title.heightCm).toBeCloseTo(0.18 * 19.05, 3);
    expect(subtitle.yCm).toBeCloseTo(0.52 * 19.05, 3);
    expect(title.paragraphs).toEqual([]);
  });

  it("mints a distinct id each time", () => {
    expect(blank().id).not.toBe(blank().id);
  });
});

describe("addSlide", () => {
  it("inserts a title+body slide after the current one and returns its index (AC-25/26)", () => {
    const model = blank();
    const index = addSlide(model, 0);
    expect(index).toBe(1);
    expect(model.slides).toHaveLength(2);
    expect(model.structureDirty).toBe(true);
    const [title, body] = model.slides[1].frames as TextFrame[];
    expect(title.role).toBe("title");
    expect(body.role).toBe("body");
    expect(body.phType).toBe("body");
    expect(body.phIdx).toBe(1);
    expect(body.yCm).toBeCloseTo(0.3 * model.slideHeightCm, 3);
  });
});

describe("duplicateSlide", () => {
  it("inserts an exact copy after its source, fresh ids, shared partName (AC-29/30/31/32)", async () => {
    const model = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Deck");
    const source = model.slides[1];
    const index = duplicateSlide(model, 1);
    expect(index).toBe(2);
    const copy = model.slides[2];
    expect(model.slides).toHaveLength(5);
    expect(model.structureDirty).toBe(true);
    // Same frames, fresh frame + slide ids.
    expect(copy.id).not.toBe(source.id);
    expect(copy.frames.map((f) => f.id)).not.toEqual(source.frames.map((f) => f.id));
    expect(copy.frames.map((f) => ({ ...f, id: "x" }))).toEqual(source.frames.map((f) => ({ ...f, id: "x" })));
    // partName shared with the source — the pptx writer copies the source bytes.
    expect(copy.partName).toBe(source.partName);
  });
});

describe("deleteSlide", () => {
  it("removes the slide and returns the nearest neighbour, down to zero (AC-34/36)", () => {
    const model = blank();
    addSlide(model, 0);
    addSlide(model, 1); // 3 slides
    expect(deleteSlide(model, 1)).toBe(1); // middle → same index (now the old third)
    expect(model.slides).toHaveLength(2);
    expect(deleteSlide(model, 1)).toBe(0); // last → previous
    expect(deleteSlide(model, 0)).toBe(0); // only → 0 (empty)
    expect(model.slides).toHaveLength(0);
  });
});

describe("moveSlide", () => {
  it("moves and clamps, keeping the moved slide's new index (AC-37)", () => {
    const model = blank();
    addSlide(model, 0);
    addSlide(model, 1); // ids a,b,c
    const [a, b, c] = model.slides;
    expect(moveSlide(model, 2, 0)).toBe(0);
    expect(model.slides).toEqual([c, a, b]);
    expect(model.structureDirty).toBe(true);
    expect(moveSlide(model, 0, 99)).toBe(2); // clamp
    expect(model.slides).toEqual([a, b, c]);
  });
});

describe("stripRetention", () => {
  it("drops every retained-package field so the writers rebuild from the model (AC-15)", async () => {
    const model = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Deck");
    model.structureDirty = true;
    const stripped = stripRetention(model);
    expect(stripped.opaque).toBeUndefined();
    expect(stripped.odpContentEnvelope).toBeUndefined();
    expect(stripped.structureDirty).toBeUndefined();
    for (const slide of stripped.slides) {
      expect(slide.partName).toBeUndefined();
      expect(slide.sourceXml).toBeUndefined();
      expect(slide.preserved).toBeUndefined();
    }
    // The source model is untouched (a copy, not a mutation).
    expect(model.opaque).toBeDefined();
    expect(model.slides[0].partName).toBeDefined();
  });
});

describe("unconvertibleFrameCount", () => {
  it("is 0 for a placeholder-only deck and counts non-text frames otherwise (AC-13 notice)", async () => {
    expect(unconvertibleFrameCount(blank())).toBe(0);
    const model = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Deck");
    // The fixture's slide 4 carries a shape, an image, a group and a foreign frame.
    expect(unconvertibleFrameCount(model)).toBeGreaterThan(0);
  });
});
