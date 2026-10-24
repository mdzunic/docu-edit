// SPEC-115 — the plan validators. validatePlan defaults rather than rejects
// wherever a default exists (each closed-vocabulary substitution and each cut
// recorded), and rejects only what is not a plan at all; validateRedesign
// never rejects a reply for its slide count — it normalises every reply to
// exactly the open deck's slides, each index once, with a drop record for
// every entry it discarded, appended or kept.
import { describe, expect, it } from "vitest";
import { BUILT_IN_THEMES } from "../core/theme";
import { STANDARD_LAYOUTS } from "../presentation/layouts";
import { DocumentGenerationError } from "../generate/errors";
import type { DeckDropRecord } from "../generate/deckDrops";
import {
  MAX_PLANNED_SLIDES,
  PLAN_LAYOUT_KINDS,
  PLAN_THEME_NAMES,
  normaliseThemeName,
  validatePlan,
  validateRedesign,
  type DeckRedesign,
} from "./plan";

const kinds = (dropped: DeckDropRecord[]) => dropped.map((record) => record.kind);

function expectEveryIndexOnce(redesign: DeckRedesign, slideCount: number): void {
  expect(redesign.slides).toHaveLength(slideCount);
  expect(redesign.slides.map((entry) => entry.source).toSorted((a, b) => a - b)).toEqual(
    Array.from({ length: slideCount }, (_, i) => i),
  );
}

describe("the closed vocabulary comes from the product's registries", () => {
  it("layout kinds are SPEC-089's standard set — the SlideLayoutKind union minus other", () => {
    expect(PLAN_LAYOUT_KINDS).toEqual(STANDARD_LAYOUTS.map((def) => def.kind));
    expect(PLAN_LAYOUT_KINDS).toEqual(["title", "titleAndBody", "sectionHeader", "twoContent", "titleOnly", "blank"]);
    expect(PLAN_LAYOUT_KINDS).not.toContain("other");
  });

  it("theme names are SPEC-094's BUILT_IN_THEMES", () => {
    expect(PLAN_THEME_NAMES).toEqual(BUILT_IN_THEMES.map((theme) => theme.name));
  });

  it("a theme name differing only in case or surrounding whitespace normalises to the registry spelling", () => {
    expect(normaliseThemeName("  retrospect ")).toBe("Retrospect");
    expect(normaliseThemeName("ION")).toBe("Ion");
    expect(normaliseThemeName("Ionic")).toBeUndefined();
    expect(normaliseThemeName(3)).toBeUndefined();
  });
});

describe("validatePlan", () => {
  const good = {
    title: "Rotterdam migration",
    theme: "Facet",
    size: "4:3",
    slides: [
      { layout: "title", title: "Rotterdam migration", bullets: ["Board update"] },
      { layout: "titleAndBody", title: "Status", bullets: ["On track", { text: "Phase 2 done", level: 1 }], note: "Lead with the date." },
      { layout: "sectionHeader", title: "Risks" },
    ],
  };

  it("accepts a well-formed plan unchanged, with nothing dropped", () => {
    const { plan, dropped } = validatePlan(good);
    expect(dropped).toEqual([]);
    expect(plan).toEqual({
      title: "Rotterdam migration",
      theme: "Facet",
      size: "4:3",
      slides: [
        { layout: "title", title: "Rotterdam migration", bullets: ["Board update"] },
        { layout: "titleAndBody", title: "Status", bullets: ["On track", { text: "Phase 2 done", level: 1 }], note: "Lead with the date." },
        { layout: "sectionHeader", title: "Risks", bullets: [] },
      ],
    });
  });

  it("an unknown layout (including other) becomes titleAndBody with bullets, titleOnly without — recorded", () => {
    const { plan, dropped } = validatePlan({
      ...good,
      slides: [
        { layout: "comparison", title: "A", bullets: ["x"] },
        { layout: "other", title: "B" },
        { title: "C" },
      ],
    });
    expect(plan.slides.map((slide) => slide.layout)).toEqual(["titleAndBody", "titleOnly", "titleOnly"]);
    expect(dropped).toEqual([
      { kind: "layout", location: "slide 1", reason: 'unknown layout "comparison"; used titleAndBody' },
      { kind: "layout", location: "slide 2", reason: 'unknown layout "other"; used titleOnly' },
      { kind: "layout", location: "slide 3", reason: "no layout named; used titleOnly" },
    ]);
  });

  it("an unknown theme becomes Office, a case-different one is accepted silently", () => {
    const unknown = validatePlan({ ...good, theme: "Corporate Blue" });
    expect(unknown.plan.theme).toBe("Office");
    expect(unknown.dropped).toEqual([{ kind: "theme", location: "deck", reason: 'unknown theme "Corporate Blue"; used Office' }]);

    const cased = validatePlan({ ...good, theme: " wisp " });
    expect(cased.plan.theme).toBe("Wisp");
    expect(cased.dropped).toEqual([]);
  });

  it("an unknown or missing size becomes 16:9, recorded", () => {
    const unknown = validatePlan({ ...good, size: "21:9" });
    expect(unknown.plan.size).toBe("16:9");
    expect(kinds(unknown.dropped)).toEqual(["slideSize"]);
    const { size: _size, ...withoutSize } = good;
    void _size;
    const missing = validatePlan(withoutSize);
    expect(missing.plan.size).toBe("16:9");
    expect(missing.dropped).toEqual([{ kind: "slideSize", location: "deck", reason: "no slide size named; used 16:9" }]);
  });

  it("defaults: a missing deck title, slide title and bullets; null counts as absent", () => {
    const { plan } = validatePlan({ theme: "Office", size: "16:9", slides: [{ layout: "titleOnly", title: null, note: null }] });
    expect(plan.title).toBe("Untitled presentation");
    expect(plan.slides[0]).toEqual({ layout: "titleOnly", title: "", bullets: [] });
  });

  it("keeps at most one bullet on a title slide, recording each further one", () => {
    const { plan, dropped } = validatePlan({
      ...good,
      slides: [{ layout: "title", title: "Hello", bullets: ["Subtitle", "Second", "Third"] }],
    });
    expect(plan.slides[0].bullets).toEqual(["Subtitle"]);
    expect(dropped).toEqual([
      { kind: "bullet", location: "slide 1", reason: "a title slide keeps one line under its title; “Second” dropped" },
      { kind: "bullet", location: "slide 1", reason: "a title slide keeps one line under its title; “Third” dropped" },
    ]);
  });

  it("keeps the first 8 bullets, cuts a bullet at 200 characters and clamps a level — each recorded", () => {
    const long = "x".repeat(250);
    const bullets = [long, { text: "deep", level: 12 }, "3", "4", "5", "6", "7", "8", "nine", "ten"];
    const { plan, dropped } = validatePlan({ ...good, slides: [{ layout: "titleAndBody", title: "Many", bullets }] });
    expect(plan.slides[0].bullets).toHaveLength(8);
    expect(plan.slides[0].bullets[0]).toBe("x".repeat(200));
    expect(plan.slides[0].bullets[1]).toEqual({ text: "deep", level: 8 });
    expect(dropped.map((record) => record.reason)).toEqual([
      "bullet 1 cut at 200 characters",
      "bullet 2's level 12 is outside 0–8; used 8",
      "a slide holds at most 8 bullets; “nine” dropped",
      "a slide holds at most 8 bullets; “ten” dropped",
    ]);
    expect(new Set(kinds(dropped))).toEqual(new Set(["bullet"]));
  });

  it(`cuts slides past ${MAX_PLANNED_SLIDES}, recorded once`, () => {
    const slides = Array.from({ length: 75 }, (_, i) => ({ layout: "titleOnly", title: `S${i + 1}` }));
    const { plan, dropped } = validatePlan({ ...good, slides });
    expect(plan.slides).toHaveLength(MAX_PLANNED_SLIDES);
    expect(plan.slides.at(-1)?.title).toBe("S60");
    expect(dropped).toEqual([
      { kind: "slideCount", location: "deck", reason: "75 slides planned; a plan holds at most 60, so slides 61–75 were cut" },
    ]);
  });

  it("a chart or table slide becomes a titleAndBody slide with SPEC-106's chart/table drop", () => {
    const { plan, dropped } = validatePlan({
      ...good,
      slides: [
        { layout: "chart", title: "Revenue", bullets: ["Q1 10"] },
        { layout: "titleOnly", title: "Grid", table: { rows: 3 } },
      ],
    });
    expect(plan.slides.map((slide) => slide.layout)).toEqual(["titleAndBody", "titleAndBody"]);
    expect(kinds(dropped)).toEqual(["chart", "table"]);
  });

  it("ignores unknown keys", () => {
    const { plan, dropped } = validatePlan({ ...good, colour: "blue", slides: [{ layout: "blank", title: "", mood: "calm" }] });
    expect(plan.slides[0]).toEqual({ layout: "blank", title: "", bullets: [] });
    expect(dropped).toEqual([]);
  });

  it.each([
    ["a non-object", [], "plan: expected a JSON object"],
    ["null", null, "plan: expected a JSON object"],
    ["missing slides", { title: "x" }, "slides: missing; expected a non-empty array of slides"],
    ["slides not an array", { slides: {} }, "slides: expected a non-empty array of slides"],
    ["empty slides", { slides: [] }, "slides: expected a non-empty array of slides"],
    ["a slide that is not an object", { slides: [{ title: "a" }, "b"] }, "slides[1]: expected an object"],
    ["a non-string slide title", { slides: [{ title: 4 }] }, "slides[0].title: expected a string"],
    ["a non-string deck title", { title: ["x"], slides: [{ title: "a" }] }, "title: expected a string"],
    ["a non-string bullet", { slides: [{ title: "a", bullets: ["ok", 7] }] }, "slides[0].bullets[1]: expected a string or { text, level }"],
    ["a non-string bullet text", { slides: [{ title: "a", bullets: [{ text: 7 }] }] }, "slides[0].bullets[0].text: expected a string"],
    ["a non-string note", { slides: [{ title: "a", note: { text: "n" } }] }, "slides[0].note: expected a string"],
  ])("rejects %s with a DocumentGenerationError naming the path", (_name, value, message) => {
    expect(() => validatePlan(value)).toThrow(DocumentGenerationError);
    expect(() => validatePlan(value)).toThrow(message);
  });
});

describe("validateRedesign (AC-86)", () => {
  it("a well-formed reply is kept as it is", () => {
    const { redesign, dropped } = validateRedesign(
      { theme: "Ion", slides: [{ source: 2, layout: "title" }, { source: 0, layout: null }, { source: 1, layout: "titleOnly" }] },
      3,
    );
    expect(redesign).toEqual({
      theme: "Ion",
      slides: [
        { source: 2, layout: "title" },
        { source: 0, layout: null },
        { source: 1, layout: "titleOnly" },
      ],
    });
    expect(dropped).toEqual([]);
  });

  it("fewer slides than the deck: the omitted ones are appended in original order, layout kept", () => {
    const { redesign, dropped } = validateRedesign({ slides: [{ source: 2 }, { source: 0, layout: "sectionHeader" }] }, 4);
    expectEveryIndexOnce(redesign, 4);
    expect(redesign.slides).toEqual([
      { source: 2, layout: null },
      { source: 0, layout: "sectionHeader" },
      { source: 1, layout: null },
      { source: 3, layout: null },
    ]);
    expect(dropped).toEqual([
      { kind: "slideOrder", location: "slide 2", reason: "not placed by the model; kept, after the placed slides, with its current layout" },
      { kind: "slideOrder", location: "slide 4", reason: "not placed by the model; kept, after the placed slides, with its current layout" },
    ]);
  });

  it("more slides than the deck: entries naming no slide of it are discarded", () => {
    const reply = { slides: [0, 1, 2, 3, 4].map((source) => ({ source })) };
    const { redesign, dropped } = validateRedesign(reply, 3);
    expectEveryIndexOnce(redesign, 3);
    expect(redesign.slides.map((entry) => entry.source)).toEqual([0, 1, 2]);
    expect(kinds(dropped)).toEqual(["slideOrder", "slideOrder"]);
    expect(dropped[0].reason).toBe("entry 4 names slide index 3, which is not a slide of this deck; ignored");
  });

  it("a duplicate source keeps its first entry and discards the repeat", () => {
    const { redesign, dropped } = validateRedesign({ slides: [{ source: 1, layout: "blank" }, { source: 1, layout: "title" }, { source: 0 }] }, 2);
    expectEveryIndexOnce(redesign, 2);
    expect(redesign.slides).toEqual([
      { source: 1, layout: "blank" },
      { source: 0, layout: null },
    ]);
    expect(dropped).toEqual([{ kind: "slideOrder", location: "slide 2", reason: "named more than once; entry 2 ignored" }]);
  });

  it("an out-of-range, negative, non-integer or non-number source is discarded", () => {
    const { redesign, dropped } = validateRedesign(
      { slides: [{ source: 9 }, { source: -1 }, { source: 1.5 }, { source: "0" }, { layout: "title" }, "x", { source: 1 }] },
      2,
    );
    expectEveryIndexOnce(redesign, 2);
    expect(redesign.slides).toEqual([
      { source: 1, layout: null },
      { source: 0, layout: null },
    ]);
    expect(kinds(dropped)).toEqual(["slideOrder", "slideOrder", "slideOrder", "slideOrder", "slideOrder", "slideOrder", "slideOrder"]);
    expect(dropped.slice(0, 6).map((record) => record.location)).toEqual(["deck", "deck", "deck", "deck", "deck", "deck"]);
  });

  it("an unknown layout (or other) becomes Keep current, recorded against its slide", () => {
    const { redesign, dropped } = validateRedesign({ slides: [{ source: 0, layout: "hero" }, { source: 1, layout: "other" }] }, 2);
    expectEveryIndexOnce(redesign, 2);
    expect(redesign.slides).toEqual([
      { source: 0, layout: null },
      { source: 1, layout: null },
    ]);
    expect(dropped).toEqual([
      { kind: "layout", location: "slide 1", reason: 'unknown layout "hero"; kept the current layout' },
      { kind: "layout", location: "slide 2", reason: 'unknown layout "other"; kept the current layout' },
    ]);
  });

  it("an unknown theme becomes Keep current; absent or null keeps without a record; a case-different name is accepted", () => {
    const unknown = validateRedesign({ theme: "Neon", slides: [{ source: 0 }] }, 1);
    expectEveryIndexOnce(unknown.redesign, 1);
    expect(unknown.redesign.theme).toBeNull();
    expect(unknown.dropped).toEqual([{ kind: "theme", location: "deck", reason: 'unknown theme "Neon"; kept the current theme' }]);
    expect(validateRedesign({ theme: null, slides: [] }, 1).redesign.theme).toBeNull();
    expect(validateRedesign({ slides: [] }, 1).dropped.map((record) => record.kind)).toEqual(["slideOrder"]);
    expect(validateRedesign({ theme: "SLICE", slides: [] }, 1).redesign.theme).toBe("Slice");
  });

  it("naming 2 slides of a 3-slide deck, or naming 5, yields exactly 3 rows with each slide once (AC-13)", () => {
    const two = validateRedesign({ slides: [{ source: 2 }, { source: 0 }] }, 3).redesign;
    expect(two.slides.map((entry) => entry.source)).toEqual([2, 0, 1]);
    const five = validateRedesign({ slides: [{ source: 4 }, { source: 1 }, { source: 1 }, { source: 0 }, { source: 3 }] }, 3).redesign;
    expect(five.slides.map((entry) => entry.source)).toEqual([1, 0, 2]);
  });

  it("rejects — the repair path — only a non-object or a reply whose slides is not an array", () => {
    expect(() => validateRedesign("nope", 3)).toThrow("redesign: expected a JSON object");
    expect(() => validateRedesign([], 3)).toThrow(DocumentGenerationError);
    expect(() => validateRedesign({ theme: "Ion" }, 3)).toThrow("slides: missing; expected an array of { source, layout }");
    expect(() => validateRedesign({ slides: { 0: 1 } }, 3)).toThrow("slides: expected an array of { source, layout }");
  });
});
