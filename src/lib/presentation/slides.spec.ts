// SPEC-080 — the framework-free slide commands: blank-deck shape, add /
// duplicate / delete / move selection + dirty flags, and stripRetention.
import { describe, expect, it } from "vitest";
import {
  addSlide,
  createBlankPresentation,
  deleteSlide,
  duplicateSlide,
  moveSlide,
  r3,
  stripRetention,
  titleBodyFrames,
  titleSlideFrames,
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

// SPEC-106: the placeholder pairs are exported so generated decks take the
// editor's geometry — extracted with no behaviour change, so each helper's
// frames are exactly what the command it came from builds (ids aside).
function withoutIds(frames: TextFrame[]): Omit<TextFrame, "id">[] {
  return frames.map(({ id: _id, ...rest }) => rest);
}

describe("placeholder frame exports (SPEC-106)", () => {
  it("titleSlideFrames is createBlankPresentation's title + subtitle pair", () => {
    const model = blank();
    const frames = titleSlideFrames(model.slideWidthCm, model.slideHeightCm);
    expect(withoutIds(frames)).toEqual(withoutIds(model.slides[0].frames as TextFrame[]));
    expect(new Set(frames.map((frame) => frame.id)).size).toBe(2);
  });

  it("titleBodyFrames is addSlide's title + body pair, at any surface size", () => {
    const model = blank();
    model.slideWidthCm = 25.4;
    addSlide(model, 0);
    expect(withoutIds(titleBodyFrames(25.4, 19.05))).toEqual(withoutIds(model.slides[1].frames as TextFrame[]));
    const [title, body] = titleBodyFrames(25.4, 19.05);
    expect([title.role, body.role, body.phIdx]).toEqual(["title", "body", 1]);
    expect(body.xCm).toBe(r3(0.08 * 25.4));
    expect(body.heightCm).toBe(r3(0.6 * 19.05));
  });

  it("r3 rounds centimetres to three decimals", () => {
    expect(r3(2.70936)).toBe(2.709);
    expect(r3(0.0005)).toBe(0.001);
    expect(r3(12)).toBe(12);
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

  it("keeps layouts and layoutKey — the catalogue is model data, not retention (SPEC-089 AC-21)", async () => {
    const model = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Deck");
    expect((model.layouts?.length ?? 0)).toBeGreaterThan(0);
    const stripped = stripRetention(model);
    expect(stripped.layouts).toEqual(model.layouts);
    expect(stripped.slides[0].layoutKey).toBe(model.slides[0].layoutKey);
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
