// SPEC-089 layout rules: classification of every pptx type plus the exotic
// fallback, standard-slot maths, the three-pass placeholder match (exact-idx
// collision and a surplus placeholder), reset, and addSlide with/without a key.
import { describe, expect, it } from "vitest";
import {
  applySlideLayout,
  classifyFromSlots,
  classifyLayout,
  instantiateStandardLayout,
  layoutFor,
  resetSlideLayout,
  STANDARD_LAYOUTS,
  standardLayoutKey,
  synthesiseStandardLayouts,
} from "./layouts";
import { addSlide } from "./slides";
import type { LayoutSlot, PresentationModel, Slide, SlideLayout, TextFrame } from "./model";

function textFrame(partial: Partial<TextFrame> & Pick<TextFrame, "id">): TextFrame {
  return {
    xCm: 0,
    yCm: 0,
    widthCm: 1,
    heightCm: 1,
    kind: "text",
    paragraphs: [],
    ...partial,
  };
}

function model(slides: Slide[], layouts: SlideLayout[]): PresentationModel {
  return {
    kind: "presentation",
    id: "m",
    title: "T",
    format: "pptx",
    slides,
    layouts,
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
  };
}

describe("classifyLayout: every pptx type + the exotic fallback", () => {
  it("maps each known @type", () => {
    expect(classifyLayout("title", [])).toBe("title");
    expect(classifyLayout("obj", [])).toBe("titleAndBody");
    expect(classifyLayout("tx", [])).toBe("titleAndBody");
    expect(classifyLayout("secHead", [])).toBe("sectionHeader");
    expect(classifyLayout("twoObj", [])).toBe("twoContent");
    expect(classifyLayout("twoTxTwoObj", [])).toBe("twoContent");
    expect(classifyLayout("titleOnly", [])).toBe("titleOnly");
    expect(classifyLayout("blank", [])).toBe("blank");
  });

  it("classifies an unrecognised type as other (e.g. cust)", () => {
    expect(classifyLayout("cust", [])).toBe("other");
    expect(classifyLayout("vertTx", [])).toBe("other");
  });
});

describe("classifyFromSlots: no declared @type (odp)", () => {
  const slot = (role: LayoutSlot["role"]): LayoutSlot => ({ role, xCm: 0, yCm: 0, widthCm: 1, heightCm: 1 });
  it("infers kinds from slot roles", () => {
    expect(classifyFromSlots([])).toBe("blank");
    expect(classifyFromSlots([slot("title")])).toBe("titleOnly");
    expect(classifyFromSlots([slot("title"), slot("subtitle")])).toBe("title");
    expect(classifyFromSlots([slot("title"), slot("body")])).toBe("titleAndBody");
    expect(classifyFromSlots([slot("title"), slot("body"), slot("body")])).toBe("twoContent");
  });
  it("classifies an unrecognised placeholder object (odp graphic) as other", () => {
    expect(classifyFromSlots([slot("title"), slot("other")])).toBe("other");
  });
});

describe("standard-slot maths", () => {
  it("instantiates fractions at the deck size, rounded to 3dp", () => {
    const def = STANDARD_LAYOUTS.find((d) => d.kind === "titleAndBody")!;
    const layout = instantiateStandardLayout(def, 33.867, 19.05);
    expect(layout.key).toBe("standard:titleAndBody");
    const [title, body] = layout.slots;
    expect(title).toMatchObject({ role: "title", phType: "title" });
    expect(title.xCm).toBeCloseTo(0.08 * 33.867, 3);
    expect(body).toMatchObject({ role: "body", phType: "body", phIdx: 1 });
    expect(body.heightCm).toBeCloseTo(0.6 * 19.05, 3);
  });
  it("synthesises the full six-layout set", () => {
    const set = synthesiseStandardLayouts(33.867, 19.05);
    expect(set.map((l) => l.kind)).toEqual([
      "title",
      "titleAndBody",
      "sectionHeader",
      "twoContent",
      "titleOnly",
      "blank",
    ]);
    expect(set.find((l) => l.kind === "blank")!.slots).toEqual([]);
  });
});

describe("applySlideLayout: three-pass matching", () => {
  it("matches exact phType+idx first, then role, and mints unmatched slots", () => {
    const slide: Slide = {
      id: "s",
      frames: [
        textFrame({ id: "title", role: "title", phType: "title", paragraphs: [{ runs: [{ text: "Hi" }] }] }),
        textFrame({ id: "body1", role: "body", phType: "body", phIdx: 1, paragraphs: [{ runs: [{ text: "A" }] }] }),
      ],
    };
    const twoContent = instantiateStandardLayout(STANDARD_LAYOUTS.find((d) => d.kind === "twoContent")!, 33.867, 19.05);
    const m = model([slide], [twoContent]);
    const report = applySlideLayout(m, 0, twoContent.key);
    expect(report.applied).toBe(true);
    expect(report.placed).toBe(2); // title + body1
    expect(report.created).toBe(1); // body idx 2 minted
    // Title text survives, geometry snapped to the slot.
    const title = slide.frames.find((f) => f.id === "title") as TextFrame;
    expect(title.paragraphs[0].runs[0].text).toBe("Hi");
    expect(title.yCm).toBeCloseTo(twoContent.slots[0].yCm, 3);
    // Two body placeholders now present.
    const bodies = slide.frames.filter((f): f is TextFrame => f.kind === "text" && f.role === "body");
    expect(bodies).toHaveLength(2);
    expect(slide.layoutKey).toBe(twoContent.key);
    expect(slide.dirty).toBe(true);
  });

  it("handles an exact-idx collision (two body idx 1 frames): exact wins, the other is re-mapped by pass 3", () => {
    const slide: Slide = {
      id: "s",
      frames: [
        textFrame({ id: "b1a", role: "body", phType: "body", phIdx: 1 }),
        textFrame({ id: "b1b", role: "body", phType: "body", phIdx: 1 }),
      ],
    };
    const titleAndBody = instantiateStandardLayout(
      STANDARD_LAYOUTS.find((d) => d.kind === "titleAndBody")!,
      33.867,
      19.05,
    );
    const m = model([slide], [titleAndBody]);
    const report = applySlideLayout(m, 0, titleAndBody.key);
    // body slot takes b1a exactly (idx 1); title slot mops up the surplus b1b in
    // pass 3, so nothing is minted and nothing is left in place.
    expect(report.placed).toBe(2);
    expect(report.created).toBe(0);
    expect(report.keptInPlace).toEqual([]);
    const b1a = slide.frames.find((f) => f.id === "b1a") as TextFrame;
    const b1b = slide.frames.find((f) => f.id === "b1b") as TextFrame;
    expect(b1a.role).toBe("body");
    expect(b1a.phIdx).toBe(1);
    expect(b1b.role).toBe("title"); // re-mapped into the title slot
  });

  it("keeps a genuine surplus placeholder in place (more frames than slots)", () => {
    const slide: Slide = {
      id: "s",
      frames: [
        textFrame({ id: "title", role: "title", phType: "title" }),
        textFrame({ id: "extra", role: "body", phType: "body", phIdx: 1 }),
      ],
    };
    const titleOnly = instantiateStandardLayout(STANDARD_LAYOUTS.find((d) => d.kind === "titleOnly")!, 33.867, 19.05);
    const report = applySlideLayout(model([slide], [titleOnly]), 0, titleOnly.key);
    expect(report.placed).toBe(1);
    expect(report.created).toBe(0);
    expect(report.keptInPlace).toEqual(["extra"]);
  });

  it("keeps non-placeholder frames in place and never covers a picture", () => {
    const slide: Slide = {
      id: "s",
      frames: [
        textFrame({ id: "title", role: "title", phType: "title" }),
        { id: "img", kind: "image", xCm: 5, yCm: 5, widthCm: 3, heightCm: 3, mediaId: "m1" },
      ],
    };
    const twoContent = instantiateStandardLayout(STANDARD_LAYOUTS.find((d) => d.kind === "twoContent")!, 33.867, 19.05);
    const report = applySlideLayout(model([slide], [twoContent]), 0, twoContent.key);
    expect(report.keptInPlace).toContain("img");
    // Minted placeholders splice behind the image (lower z-order).
    const imgIndex = slide.frames.findIndex((f) => f.id === "img");
    const mintedBody = slide.frames.filter((f): f is TextFrame => f.kind === "text" && f.role === "body");
    expect(mintedBody.length).toBe(2);
    expect(slide.frames.indexOf(mintedBody[0])).toBeLessThan(imgIndex);
  });

  it("returns applied:false for an unknown layoutKey", () => {
    const slide: Slide = { id: "s", frames: [] };
    expect(applySlideLayout(model([slide], []), 0, "nope").applied).toBe(false);
  });
});

describe("resetSlideLayout", () => {
  it("snaps exact-match placeholders back, minting nothing", () => {
    const titleAndBody = instantiateStandardLayout(
      STANDARD_LAYOUTS.find((d) => d.kind === "titleAndBody")!,
      33.867,
      19.05,
    );
    const slide: Slide = {
      id: "s",
      layoutKey: titleAndBody.key,
      frames: [textFrame({ id: "t", role: "title", phType: "title", xCm: 99, yCm: 99 })],
    };
    const m = model([slide], [titleAndBody]);
    const report = resetSlideLayout(m, 0);
    expect(report.created).toBe(0);
    expect(report.placed).toBe(1);
    const t = slide.frames[0] as TextFrame;
    expect(t.xCm).toBeCloseTo(titleAndBody.slots[0].xCm, 3);
    expect(t.yCm).toBeCloseTo(titleAndBody.slots[0].yCm, 3);
  });

  it("is inert (applied:false) when the slide has no resolvable layout", () => {
    const slide: Slide = { id: "s", frames: [] };
    expect(resetSlideLayout(model([slide], []), 0).applied).toBe(false);
  });
});

describe("layoutFor + addSlide", () => {
  it("resolves a slide's layout, or undefined for an unknown key", () => {
    const layout = synthesiseStandardLayouts(33.867, 19.05)[0];
    const m = model([{ id: "s", frames: [], layoutKey: layout.key }], [layout]);
    expect(layoutFor(m, m.slides[0])).toBe(layout);
    expect(layoutFor(m, { id: "x", frames: [], layoutKey: "ghost" })).toBeUndefined();
  });

  it("addSlide without a key keeps SPEC-080's title + body slide", () => {
    const m = model([{ id: "s", frames: [] }], []);
    const at = addSlide(m, 0);
    const frames = m.slides[at].frames as TextFrame[];
    expect(frames.map((f) => f.role)).toEqual(["title", "body"]);
    expect(m.slides[at].layoutKey).toBeUndefined();
  });

  it("addSlide with a key mints one empty frame per slot and sets layoutKey", () => {
    const section = instantiateStandardLayout(STANDARD_LAYOUTS.find((d) => d.kind === "sectionHeader")!, 33.867, 19.05);
    const m = model([{ id: "s", frames: [] }], [section]);
    const at = addSlide(m, 0, standardLayoutKey("sectionHeader"));
    const frames = m.slides[at].frames as TextFrame[];
    expect(frames).toHaveLength(section.slots.length);
    expect(frames.every((f) => f.paragraphs.length === 0)).toBe(true);
    expect(m.slides[at].layoutKey).toBe(section.key);
  });
});
