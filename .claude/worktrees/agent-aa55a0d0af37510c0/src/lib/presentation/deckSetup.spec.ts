import { describe, expect, it } from "vitest";
import { applyDeckProperties, applyDeckSize, framesOutsideSlide, setFirstSlideNumber } from "./deckSetup";
import type { PresentationModel, Slide, SlideFrame } from "./model";

function textFrame(over: Partial<Extract<SlideFrame, { kind: "text" }>> = {}): SlideFrame {
  return {
    id: "t1",
    kind: "text",
    xCm: 2,
    yCm: 3,
    widthCm: 10,
    heightCm: 4,
    paragraphs: [{ runs: [{ text: "hello", sizePt: 20 }] }],
    ...over,
  };
}

function model(over: Partial<PresentationModel> = {}, frames: SlideFrame[] = [textFrame()]): PresentationModel {
  const slide: Slide = { id: "s1", frames };
  return {
    kind: "presentation",
    id: "d1",
    title: "Deck",
    format: "pptx",
    slides: [slide],
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
    ...over,
  };
}

describe("applyDeckSize scale (SPEC-090)", () => {
  it("scales and centres a frame; s = min(newW/oldW, newH/oldH)", () => {
    // old 33.867×19.05 → new 25.4×19.05. s = min(0.75.., 1) = 0.75 (25.4/33.867).
    const m = model();
    applyDeckSize(m, { widthCm: 25.4, heightCm: 19.05 }, "scale");
    const s = 25.4 / 33.867;
    const dx = (25.4 - 33.867 * s) / 2; // ≈ 0
    const dy = (19.05 - 19.05 * s) / 2;
    const f = m.slides[0].frames[0];
    expect(f.xCm).toBeCloseTo(2 * s + dx, 3);
    expect(f.yCm).toBeCloseTo(3 * s + dy, 3);
    expect(f.widthCm).toBeCloseTo(10 * s, 3);
    expect(f.heightCm).toBeCloseTo(4 * s, 3);
  });

  it("centres when the aspect ratio changes (height-limited scale adds horizontal dx)", () => {
    // old 25.4×19.05 → new 33.867×19.05. s = min(1.33, 1) = 1. dx = (33.867-25.4)/2.
    const m = model({ slideWidthCm: 25.4, slideHeightCm: 19.05 });
    applyDeckSize(m, { widthCm: 33.867, heightCm: 19.05 }, "scale");
    const dx = (33.867 - 25.4) / 2;
    const f = m.slides[0].frames[0];
    expect(f.xCm).toBeCloseTo(2 + dx, 3);
    expect(f.yCm).toBeCloseTo(3, 3);
    expect(f.widthCm).toBeCloseTo(10, 3);
  });

  it("scales group children in place (no second centring offset)", () => {
    const child = textFrame({ id: "c1", xCm: 4, yCm: 4, widthCm: 5, heightCm: 2 });
    const group: SlideFrame = { id: "g1", kind: "group", xCm: 1, yCm: 1, widthCm: 10, heightCm: 10, children: [child], xml: "<x/>" };
    const m = model({}, [group]);
    applyDeckSize(m, { widthCm: 25.4, heightCm: 19.05 }, "scale");
    const s = 25.4 / 33.867;
    const dx = (25.4 - 33.867 * s) / 2;
    const g = m.slides[0].frames[0] as Extract<SlideFrame, { kind: "group" }>;
    // group own bounds carry dx/dy
    expect(g.xCm).toBeCloseTo(1 * s + dx, 3);
    // child scales but WITHOUT dx/dy (absorbed by the group's move)
    expect(g.children[0].xCm).toBeCloseTo(4 * s, 3);
    expect(g.children[0].widthCm).toBeCloseTo(5 * s, 3);
    // SPEC-096: scaling is a child-geometry change, so the verbatim xml is
    // dropped and the serialisers rebuild the group at the scaled geometry
    // (present = verbatim, absent = authored). Before SPEC-096 the xml was
    // kept — and re-emitted stale at the old scale.
    expect(g.xml).toBeUndefined();
  });

  it("scales foreign frame bounds while leaving its xml verbatim", () => {
    const foreign: SlideFrame = { id: "f1", kind: "foreign", xCm: 2, yCm: 2, widthCm: 8, heightCm: 6, xml: "<obj/>", label: "Object" };
    const m = model({}, [foreign]);
    applyDeckSize(m, { widthCm: 25.4, heightCm: 19.05 }, "scale");
    const s = 25.4 / 33.867;
    const f = m.slides[0].frames[0] as Extract<SlideFrame, { kind: "foreign" }>;
    expect(f.widthCm).toBeCloseTo(8 * s, 3);
    expect(f.xml).toBe("<obj/>");
  });

  it("scales run sizePt in 0.5 pt steps and materialises the default when absent", () => {
    const m = model({}, [textFrame({ paragraphs: [{ runs: [{ text: "a" }, { text: "b", sizePt: 20 }] }] })]);
    // scale up by 2 (old 10×10 → new 20×20 with same aspect)
    applyDeckSize({ ...m, slideWidthCm: 10, slideHeightCm: 10 } as PresentationModel, { widthCm: 20, heightCm: 20 }, "scale");
    // rebuild explicitly to control size
    const m2 = model({ slideWidthCm: 10, slideHeightCm: 10 }, [textFrame({ paragraphs: [{ runs: [{ text: "a" }, { text: "b", sizePt: 20 }] }] })]);
    applyDeckSize(m2, { widthCm: 20, heightCm: 20 }, "scale");
    const runs = (m2.slides[0].frames[0] as Extract<SlideFrame, { kind: "text" }>).paragraphs[0].runs;
    expect(runs[0].sizePt).toBe(36); // 18 default × 2
    expect(runs[1].sizePt).toBe(40); // 20 × 2
  });

  it("clamps scaled font size to 6..96", () => {
    const big = model({ slideWidthCm: 10, slideHeightCm: 10 }, [textFrame({ paragraphs: [{ runs: [{ text: "a", sizePt: 80 }] }] })]);
    applyDeckSize(big, { widthCm: 20, heightCm: 20 }, "scale"); // ×2 → 160 clamped to 96
    expect((big.slides[0].frames[0] as Extract<SlideFrame, { kind: "text" }>).paragraphs[0].runs[0].sizePt).toBe(96);

    const small = model({ slideWidthCm: 20, slideHeightCm: 20 }, [textFrame({ paragraphs: [{ runs: [{ text: "a", sizePt: 8 }] }] })]);
    applyDeckSize(small, { widthCm: 5, heightCm: 5 }, "scale"); // ×0.25 → 2 clamped to 6
    expect((small.slides[0].frames[0] as Extract<SlideFrame, { kind: "text" }>).paragraphs[0].runs[0].sizePt).toBe(6);
  });

  it("sets deckDirty and marks each slide edited", () => {
    const m = model();
    applyDeckSize(m, { widthCm: 25.4, heightCm: 19.05 }, "scale");
    expect(m.deckDirty).toBe(true);
    expect(m.slides[0].dirty).toBe(true);
    expect(m.slideWidthCm).toBe(25.4);
    expect(m.slideHeightCm).toBe(19.05);
  });

  it("deletes odp sourceXml when scaling (markSlideEdited)", () => {
    const m = model({ format: "odp" });
    m.slides[0].sourceXml = "<draw:page/>";
    applyDeckSize(m, { widthCm: 25.4, heightCm: 19.05 }, "scale");
    expect(m.slides[0].sourceXml).toBeUndefined();
  });
});

describe("applyDeckSize keep (SPEC-090)", () => {
  it("changes only the deck size, leaving frames and slides untouched", () => {
    const m = model();
    const before = { ...m.slides[0].frames[0] };
    applyDeckSize(m, { widthCm: 25.4, heightCm: 19.05 }, "keep");
    expect(m.slides[0].frames[0]).toEqual(before);
    expect(m.slides[0].dirty).toBeUndefined();
    expect(m.deckDirty).toBe(true);
    expect(m.slideWidthCm).toBe(25.4);
  });
});

describe("framesOutsideSlide (SPEC-090)", () => {
  it("counts frames whose bounds fall entirely outside the slide", () => {
    const inside = textFrame({ id: "a", xCm: 1, yCm: 1, widthCm: 5, heightCm: 5 });
    const outside = textFrame({ id: "b", xCm: 30, yCm: 1, widthCm: 5, heightCm: 5 });
    const m = model({ slideWidthCm: 25.4, slideHeightCm: 19.05 }, [inside, outside]);
    expect(framesOutsideSlide(m)).toBe(1);
  });

  it("does not count a frame that partially overlaps the slide", () => {
    const straddling = textFrame({ id: "a", xCm: -2, yCm: 1, widthCm: 5, heightCm: 5 });
    const m = model({ slideWidthCm: 25.4, slideHeightCm: 19.05 }, [straddling]);
    expect(framesOutsideSlide(m)).toBe(0);
  });
});

describe("setFirstSlideNumber (SPEC-090)", () => {
  it("stores a non-default origin and sets deckDirty", () => {
    const m = model();
    setFirstSlideNumber(m, 5);
    expect(m.firstSlideNumber).toBe(5);
    expect(m.deckDirty).toBe(true);
  });

  it("clears the field for the default 1", () => {
    const m = model({ firstSlideNumber: 5 });
    setFirstSlideNumber(m, 1);
    expect(m.firstSlideNumber).toBeUndefined();
    expect(m.deckDirty).toBe(true);
  });
});

describe("applyDeckProperties (SPEC-090)", () => {
  it("replaces properties and sets deckDirty", () => {
    const m = model();
    applyDeckProperties(m, { author: "Ada", created: "2026-01-01T00:00:00Z" });
    expect(m.properties).toEqual({ author: "Ada", created: "2026-01-01T00:00:00Z" });
    expect(m.deckDirty).toBe(true);
  });
});

describe("plain-JSON invariant (SPEC-090 AC-1)", () => {
  it("a deck-setup mutated model still round-trips through JSON", () => {
    const m = model({ firstSlideNumber: 3, deckDirty: true, properties: { author: "Ada" } });
    applyDeckSize(m, { widthCm: 25.4, heightCm: 19.05 }, "scale");
    expect(JSON.parse(JSON.stringify(m))).toEqual(m);
  });
});
