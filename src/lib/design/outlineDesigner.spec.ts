// SPEC-115 — the model-free designer's rule table, exactly: over a document
// with nested headings, a document with none, an empty document and a
// 300-heading document (bounded). Pure and deterministic: the same source
// gives the same plan, and the source is never touched.
import { describe, expect, it } from "vitest";
import { MAX_PLANNED_SLIDES, validatePlan } from "./plan";
import { cutWords, planFromOutline, type OutlineBlock, type OutlineSource } from "./outlineDesigner";

const h = (level: number, text: string): OutlineBlock => ({ kind: "heading", level, text });
const p = (text: string, listLevel?: number): OutlineBlock =>
  listLevel === undefined ? { kind: "paragraph", text } : { kind: "paragraph", text, listLevel };

function deepFreeze<T>(value: T): T {
  if (value !== null && typeof value === "object") {
    Object.values(value).forEach(deepFreeze);
    Object.freeze(value);
  }
  return value;
}

const NESTED: OutlineSource = {
  title: "Rotterdam migration",
  blocks: [
    h(1, "Where we are"),
    p("A board update on the move of every workload to the Rotterdam site."),
    h(2, "Status"),
    p("Phase one finished in August, two weeks ahead of the plan agreed in the spring."),
    p("Phase two is underway", 0),
    p("Network cut-over booked", 1),
    h(3, "Risks"),
    p("Vendor lead times"),
    h(4, "Mitigation"),
    h(3, "Costs"),
    h(2, "Next steps"),
    h(1, "Appendix"),
    p("Glossary of terms"),
    p(""),
    h(1, "Questions"),
  ],
};

describe("planFromOutline — the rule table over a document with nested headings", () => {
  const plan = planFromOutline(deepFreeze(structuredClone(NESTED)));

  it("is exactly this plan", () => {
    expect(plan).toEqual({
      title: "Rotterdam migration",
      theme: "Office",
      size: "16:9",
      slides: [
        // Rule 1: the title slide, subtitled by the paragraph after the first H1.
        { layout: "title", title: "Rotterdam migration", bullets: ["A board update on the move of every workload to the Rotterdam…"] },
        // Rule 2: an H1 with H2 children is a section slide (its subtitle
        // paragraph was used by the title slide, so it is not repeated).
        { layout: "sectionHeader", title: "Where we are", bullets: [] },
        // Rules 3 and 4: an H2's paragraphs and list items (≤ 6, ≤ 12 words),
        // deeper headings folded one level per depth, their text one below.
        {
          layout: "titleAndBody",
          title: "Status",
          bullets: [
            "Phase one finished in August, two weeks ahead of the plan agreed…",
            "Phase two is underway",
            { text: "Network cut-over booked", level: 1 },
            { text: "Risks", level: 1 },
            { text: "Vendor lead times", level: 2 },
            { text: "Mitigation", level: 2 },
          ],
        },
        // Rule 5: an H2 with no text is titleOnly.
        { layout: "titleOnly", title: "Next steps", bullets: [] },
        // Rule 2: an H1 with no H2 children is a content slide.
        { layout: "titleAndBody", title: "Appendix", bullets: ["Glossary of terms"] },
        { layout: "titleOnly", title: "Questions", bullets: [] },
      ],
    });
  });

  it("caps a slide at six bullets", () => {
    expect(plan.slides[2].bullets).toHaveLength(6);
  });

  it("is a plan validatePlan accepts with nothing to change", () => {
    expect(validatePlan(plan)).toEqual({ plan, dropped: [] });
  });

  it("is deterministic", () => {
    expect(planFromOutline(NESTED)).toEqual(plan);
  });
});

describe("planFromOutline — the other documents", () => {
  it("a document with no headings: the title slide, then its text on one slide titled with the document's title", () => {
    const source: OutlineSource = {
      title: "Notes",
      blocks: ["One.", "Two.", "Three.", "Four.", "Five.", "Six.", "Seven."].map((text) => p(text)),
    };
    expect(planFromOutline(deepFreeze(source))).toEqual({
      title: "Notes",
      theme: "Office",
      size: "16:9",
      slides: [
        { layout: "title", title: "Notes", bullets: [] },
        { layout: "titleAndBody", title: "Notes", bullets: ["One.", "Two.", "Three.", "Four.", "Five.", "Six."] },
      ],
    });
  });

  it("an empty document: one title slide, with the default title when it has none", () => {
    expect(planFromOutline({ title: "", blocks: [] })).toEqual({
      title: "Untitled presentation",
      theme: "Office",
      size: "16:9",
      slides: [{ layout: "title", title: "Untitled presentation", bullets: [] }],
    });
  });

  it("a 300-heading document is bounded at the plan's slide limit, or at maxSlides", () => {
    const blocks = Array.from({ length: 300 }, (_, i) => [h(2, `Topic ${i + 1}`), p(`Detail ${i + 1}`)]).flat();
    const source = deepFreeze({ title: "Long", blocks });
    const plan = planFromOutline(source);
    expect(plan.slides).toHaveLength(MAX_PLANNED_SLIDES);
    expect(plan.slides[1]).toEqual({ layout: "titleAndBody", title: "Topic 1", bullets: ["Detail 1"] });
    expect(plan.slides.at(-1)?.title).toBe(`Topic ${MAX_PLANNED_SLIDES - 1}`);
    expect(planFromOutline(source, { maxSlides: 5 }).slides.map((slide) => slide.title)).toEqual([
      "Long",
      "Topic 1",
      "Topic 2",
      "Topic 3",
      "Topic 4",
    ]);
    expect(planFromOutline(source, { maxSlides: 500 }).slides).toHaveLength(MAX_PLANNED_SLIDES);
    expect(validatePlan(plan).dropped).toEqual([]);
  });

  it("rule 6: the source deck's theme (a built-in, in any case), else the default; size 16:9", () => {
    expect(planFromOutline({ title: "D", theme: "wisp", blocks: [] }).theme).toBe("Wisp");
    expect(planFromOutline({ title: "D", theme: "House style", blocks: [] }).theme).toBe("Office");
    expect(planFromOutline({ title: "D", theme: "Ion", blocks: [] }).size).toBe("16:9");
  });

  it("an H1 whose next block is a heading gives the title slide no subtitle", () => {
    const plan = planFromOutline({ title: "D", blocks: [h(1, "A"), h(2, "B"), p("b")] });
    expect(plan.slides[0].bullets).toEqual([]);
    expect(plan.slides.map((slide) => slide.layout)).toEqual(["title", "sectionHeader", "titleAndBody"]);
  });
});

describe("cutWords", () => {
  it("cuts at 12 words on a word boundary with an ellipsis, collapsing whitespace", () => {
    expect(cutWords("one two  three")).toBe("one two three");
    expect(cutWords("a b c d e f g h i j k l")).toBe("a b c d e f g h i j k l");
    expect(cutWords("a b c d e f g h i j k l m")).toBe("a b c d e f g h i j k l…");
  });

  it("keeps a bullet within the plan's 200 characters", () => {
    expect(Array.from(cutWords("x".repeat(300)))).toHaveLength(200);
  });
});
