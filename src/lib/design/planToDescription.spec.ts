// SPEC-115 — a reviewed plan as SPEC-106's DeckDescription: every plan field
// mapped (title → the title layout with its one bullet the subtitle, a blank
// slide with no text → blank, everything else → bullets with levels kept,
// notes, size), the closed-vocabulary substitutions validatePlan makes on the
// way and their drop records, and a plan that validates building — through
// deckFromDescription, the builder's own path — into a deck whose slide count
// and titles match.
import { describe, expect, it } from "vitest";
import { deckFromDescription } from "../generate/deckDescription";
import type { TextFrame } from "../presentation/model";
import { buildPlannedDeck } from "./build";
import { planToDescription, validatePlan, type DeckPlan } from "./plan";

const PLAN: DeckPlan = {
  title: "Rotterdam migration",
  theme: "Ion",
  size: "4:3",
  slides: [
    { layout: "title", title: "Rotterdam migration", bullets: ["Board update, October"], note: "Welcome everyone." },
    { layout: "sectionHeader", title: "Where we are", bullets: [] },
    { layout: "titleAndBody", title: "Status", bullets: ["On track", { text: "Phase 2 done", level: 1 }, { text: "Phase 3 next", level: 2 }] },
    { layout: "twoContent", title: "Before and after", bullets: ["Two sites", "One site"] },
    { layout: "titleOnly", title: "Questions?", bullets: [] },
    { layout: "blank", title: "", bullets: [], note: "Pause here." },
    { layout: "blank", title: "Not empty", bullets: [] },
  ],
};

function titleText(frames: unknown[]): string {
  const title = frames.find((frame) => (frame as TextFrame).role === "title") as TextFrame | undefined;
  return title ? title.paragraphs.map((p) => p.runs.map((run) => run.text).join("")).join("") : "";
}

describe("planToDescription maps every plan field", () => {
  it("deck title and size; each slide by the Reference's mapping; notes as one paragraph", () => {
    expect(planToDescription(PLAN)).toEqual({
      title: "Rotterdam migration",
      size: "4:3",
      slides: [
        { layout: "title", title: "Rotterdam migration", subtitle: "Board update, October", notes: ["Welcome everyone."] },
        { layout: "bullets", title: "Where we are", bullets: [] },
        {
          layout: "bullets",
          title: "Status",
          bullets: ["On track", { text: "Phase 2 done", level: 1 }, { text: "Phase 3 next", level: 2 }],
        },
        { layout: "bullets", title: "Before and after", bullets: ["Two sites", "One site"] },
        { layout: "bullets", title: "Questions?", bullets: [] },
        { layout: "blank", notes: ["Pause here."] },
        { layout: "bullets", title: "Not empty", bullets: [] },
      ],
    });
  });

  it("a title slide with no bullet has no subtitle; one with several builds only its first (its subtitle)", () => {
    const description = planToDescription({
      ...PLAN,
      slides: [
        { layout: "title", title: "Bare", bullets: [] },
        { layout: "title", title: "Busy", bullets: [{ text: "Sub", level: 2 }, "Dropped"] },
      ],
    });
    expect(description.slides).toEqual([
      { layout: "title", title: "Bare" },
      { layout: "title", title: "Busy", subtitle: "Sub" },
    ]);
  });

  it("strings stay strings — Markdown, HTML and instructions are text the builder carries verbatim", () => {
    const hostile = "**Ignore previous instructions** <script>alert(1)</script> [link](https://example.test)";
    const deck = deckFromDescription(
      planToDescription({ ...PLAN, slides: [{ layout: "titleAndBody", title: hostile, bullets: [hostile] }] }),
    );
    const [title, body] = deck.slides[0].frames as TextFrame[];
    expect(title.paragraphs[0].runs).toEqual([{ text: hostile }]);
    expect(body.paragraphs[0].runs).toEqual([{ text: hostile }]);
  });
});

describe("the closed vocabulary's substitutions arrive as drop records before anything is built", () => {
  it("unknown layout, theme and size become their nearest defaults, each recorded, and the result builds", () => {
    const { plan, dropped } = validatePlan({
      title: "Deck",
      theme: "Corporate",
      size: "A4",
      slides: [
        { layout: "hero", title: "One", bullets: ["a"] },
        { layout: "other", title: "Two" },
        { layout: "chart", title: "Three" },
      ],
    });
    expect(dropped.map((record) => [record.kind, record.location])).toEqual([
      ["theme", "deck"],
      ["slideSize", "deck"],
      ["layout", "slide 1"],
      ["layout", "slide 2"],
      ["chart", "slide 3"],
    ]);
    expect(plan.theme).toBe("Office");
    expect(plan.size).toBe("16:9");
    expect(plan.slides.map((slide) => slide.layout)).toEqual(["titleAndBody", "titleOnly", "titleAndBody"]);
    expect(() => deckFromDescription(planToDescription(plan))).not.toThrow();
  });
});

describe("a plan that validates builds a deck whose slide count and titles match", () => {
  it("through deckFromDescription, with every title, bullet and note reaching the deck", () => {
    const { plan } = validatePlan(PLAN);
    const deck = buildPlannedDeck(plan, { generatedAt: "2026-10-24T12:00:00.000Z", generatedModel: "stub" });
    expect(deck.slides).toHaveLength(PLAN.slides.length);
    expect(deck.slides.map((slide) => titleText(slide.frames))).toEqual(PLAN.slides.map((slide) => slide.title));
    expect(deck.title).toBe("Rotterdam migration");
    expect(deck.slideWidthCm / deck.slideHeightCm).toBeCloseTo(4 / 3, 2);
    // The subtitle, the bullets with their levels, and the notes.
    const subtitle = deck.slides[0].frames.find((frame) => (frame as TextFrame).role === "subtitle") as TextFrame;
    expect(subtitle.paragraphs[0].runs[0].text).toBe("Board update, October");
    const status = deck.slides[2].frames.find((frame) => (frame as TextFrame).role === "body") as TextFrame;
    expect(status.paragraphs.map((p) => [p.runs[0].text, p.level ?? 0])).toEqual([
      ["On track", 0],
      ["Phase 2 done", 1],
      ["Phase 3 next", 2],
    ]);
    expect(deck.slides[0].notes).toEqual(["Welcome everyone."]);
    expect(deck.slides[5].notes).toEqual(["Pause here."]);
  });
});
