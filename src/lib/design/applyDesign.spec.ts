// SPEC-115 — the designer's model mutations. applyPlanDesign sets a built
// deck's theme and every slide's kind (resolved to the deck's own layout, or
// the synthesised standard:{kind} one added to its catalogue) and removes the
// empty placeholders a layout kept in place; applyRedesign reorders, re-lays
// out and re-themes the open deck without adding, deleting or rewriting a
// slide; isRedesignNoop is what lets the view record nothing when nothing
// changes.
import { describe, expect, it } from "vitest";
import { BUILT_IN_THEMES } from "../core/theme";
import { layoutFor, synthesiseStandardLayouts } from "../presentation/layouts";
import type { PresentationModel, Slide, SlideLayout, TextFrame } from "../presentation/model";
import { createBlankPresentation, titleBodyFrames } from "../presentation/slides";
import {
  applyPlanDesign,
  applyRedesign,
  isRedesignNoop,
  resolveLayoutKey,
  seedLayoutKind,
  slideTitleText,
} from "./applyDesign";
import { buildPlannedDeck } from "./build";
import { PLAN_LAYOUT_KINDS, type DeckPlan, type PlanLayoutKind } from "./plan";

const PROVENANCE = { generatedAt: "2026-10-24T12:00:00.000Z", generatedModel: "stub" };

function planWith(kinds: PlanLayoutKind[]): DeckPlan {
  return {
    title: "Kinds",
    theme: "Retrospect",
    size: "16:9",
    slides: kinds.map((layout, i) => ({ layout, title: `Slide ${i + 1}`, bullets: layout === "title" ? ["Sub"] : i % 2 ? ["Point"] : [] })),
  };
}

function textOf(frame: TextFrame): string {
  return frame.paragraphs.map((p) => p.runs.map((run) => run.text).join("")).join("\n");
}

/** A deck of three titled slides with notes and a picture-like free text box. */
function threeSlideDeck(): PresentationModel {
  const deck = createBlankPresentation();
  deck.slides = ["Alpha", "Beta", "Gamma"].map((title, i): Slide => {
    const [titleFrame, body] = titleBodyFrames(deck.slideWidthCm, deck.slideHeightCm);
    titleFrame.paragraphs = [{ runs: [{ text: title }] }];
    body.paragraphs = [{ runs: [{ text: `${title} body` }] }];
    const box: TextFrame = { id: `box-${i}`, kind: "text", xCm: 1, yCm: 1, widthCm: 3, heightCm: 1, paragraphs: [{ runs: [{ text: `note box ${i}` }] }] };
    return { id: `s${i}`, frames: [titleFrame, body, box], notes: [`${title} notes`] };
  });
  return deck;
}

function snapshotText(deck: PresentationModel) {
  return deck.slides.map((slide) => ({
    id: slide.id,
    notes: slide.notes,
    text: slide.frames.filter((f): f is TextFrame => f.kind === "text").map(textOf),
    free: slide.frames.filter((f) => f.kind !== "text" || f.role === undefined),
  }));
}

describe("applyPlanDesign — every built slide has a layoutFor result of exactly its planned kind (AC-45)", () => {
  it("each of the six kinds, on a deck with no catalogue: synthesised standard layouts are added", () => {
    const deck = buildPlannedDeck(planWith([...PLAN_LAYOUT_KINDS]), PROVENANCE);
    expect(deck.slides.map((slide) => layoutFor(deck, slide)?.kind)).toEqual([...PLAN_LAYOUT_KINDS]);
    expect(deck.slides.map((slide) => slide.layoutKey)).toEqual(PLAN_LAYOUT_KINDS.map((kind) => `standard:${kind}`));
    const standard = synthesiseStandardLayouts(deck.slideWidthCm, deck.slideHeightCm);
    expect(deck.layouts).toEqual(PLAN_LAYOUT_KINDS.map((kind) => standard.find((layout) => layout.kind === kind)));
    // Plain model mutations: nothing is marked edited on a deck with no package.
    expect(deck.slides.some((slide) => slide.dirty)).toBe(false);
  });

  it("a kind resolves to the deck's own first layout of that kind before any synthesised one", () => {
    const deck = createBlankPresentation();
    const own: SlideLayout = { key: "ppt/slideLayouts/custom7.xml", kind: "titleOnly", name: "Headline", slots: [] };
    const second: SlideLayout = { ...own, key: "ppt/slideLayouts/custom8.xml" };
    deck.layouts = [own, second];
    expect(resolveLayoutKey(deck, "titleOnly")).toBe(own.key);
    expect(deck.layouts).toHaveLength(2);
    expect(resolveLayoutKey(deck, "blank")).toBe("standard:blank");
    expect(deck.layouts.map((layout) => layout.key)).toEqual([own.key, second.key, "standard:blank"]);
    // Resolving again reuses the layout it added.
    expect(resolveLayoutKey(deck, "blank")).toBe("standard:blank");
    expect(deck.layouts).toHaveLength(3);
  });

  it("removes a placeholder the layout kept in place when it carries no text — a titleOnly slide has no empty body box", () => {
    const deck = buildPlannedDeck(
      { title: "T", theme: "Office", size: "16:9", slides: [{ layout: "titleOnly", title: "Only a title", bullets: [] }] },
      PROVENANCE,
    );
    const frames = deck.slides[0].frames as TextFrame[];
    expect(frames.map((frame) => frame.role)).toEqual(["title"]);
    expect(textOf(frames[0])).toBe("Only a title");
  });

  it("keeps a placeholder that carries text even when the layout has no slot for it (blank with a title)", () => {
    const deck = buildPlannedDeck(
      { title: "T", theme: "Office", size: "16:9", slides: [{ layout: "blank", title: "Kept", bullets: ["Also kept"] }] },
      PROVENANCE,
    );
    expect(layoutFor(deck, deck.slides[0])?.kind).toBe("blank");
    expect((deck.slides[0].frames as TextFrame[]).map(textOf)).toEqual(["Kept", "Also kept"]);
  });

  it("sets the plan's theme with applyTheme", () => {
    const deck = buildPlannedDeck(planWith(["title"]), PROVENANCE);
    const retrospect = BUILT_IN_THEMES.find((theme) => theme.name === "Retrospect")!;
    expect(deck.theme).toMatchObject({ name: "Retrospect", colors: retrospect.colors, fonts: retrospect.fonts });
  });

  it("a slide the plan has no row for is left alone", () => {
    const deck = createBlankPresentation();
    applyPlanDesign(deck, { title: "T", theme: "Office", size: "16:9", slides: [] });
    expect(deck.slides[0].layoutKey).toBeUndefined();
  });
});

describe("applyRedesign — order, layouts and theme only (AC-5, AC-20..22)", () => {
  it("moves the slides into the reviewed order and leaves count, text, notes and non-placeholder frames identical", () => {
    const deck = threeSlideDeck();
    const before = snapshotText(deck);
    applyRedesign(deck, {
      theme: null,
      slides: [
        { source: 2, layout: null },
        { source: 0, layout: null },
        { source: 1, layout: null },
      ],
    });
    expect(deck.slides.map((slide) => slide.id)).toEqual(["s2", "s0", "s1"]);
    expect(deck.slides).toHaveLength(3);
    const byId = new Map(snapshotText(deck).map((entry) => [entry.id, entry]));
    for (const entry of before) {
      expect(byId.get(entry.id)).toEqual(entry);
    }
    expect(deck.structureDirty).toBe(true);
  });

  it("applies a row's kind only where it differs from the slide's current kind, and a chosen theme", () => {
    const deck = threeSlideDeck();
    const before = snapshotText(deck);
    deck.layouts = synthesiseStandardLayouts(deck.slideWidthCm, deck.slideHeightCm);
    deck.slides[1].layoutKey = "standard:titleAndBody";
    applyRedesign(deck, {
      theme: "Slice",
      slides: [
        { source: 0, layout: "sectionHeader" },
        { source: 1, layout: "titleAndBody" },
        { source: 2, layout: "titleOnly" },
      ],
    });
    expect(deck.slides.map((slide) => layoutFor(deck, slide)?.kind)).toEqual(["sectionHeader", "titleAndBody", "titleOnly"]);
    // Slide 2 already had its kind: not re-applied, so not marked edited.
    expect(deck.slides[1].dirty).toBeUndefined();
    expect(deck.theme?.name).toBe("Slice");
    // Still no text, note or free frame changed — a titleOnly slide keeps its
    // body text (a placeholder with text is never removed).
    expect(snapshotText(deck)).toEqual(before);
  });

  it("seeds a row with the slide's current kind when it is one of the six, else Keep current", () => {
    const deck = threeSlideDeck();
    deck.layouts = [
      ...synthesiseStandardLayouts(deck.slideWidthCm, deck.slideHeightCm),
      { key: "custom", kind: "other", name: "Custom", slots: [] },
    ];
    deck.slides[0].layoutKey = "standard:sectionHeader";
    deck.slides[1].layoutKey = "custom";
    expect(deck.slides.map((slide) => seedLayoutKind(deck, slide))).toEqual(["sectionHeader", null, null]);
    expect(deck.slides.map(slideTitleText)).toEqual(["Alpha", "Beta", "Gamma"]);
  });
});

describe("isRedesignNoop (AC-24)", () => {
  it("is true for the deck's own order, kept or unchanged kinds, and a kept or current theme", () => {
    const deck = threeSlideDeck();
    deck.layouts = synthesiseStandardLayouts(deck.slideWidthCm, deck.slideHeightCm);
    deck.slides[0].layoutKey = "standard:title";
    const identity = [0, 1, 2].map((source) => ({ source, layout: null }));
    expect(isRedesignNoop(deck, { theme: null, slides: identity })).toBe(true);
    expect(isRedesignNoop(deck, { theme: "Office", slides: [{ source: 0, layout: "title" }, ...identity.slice(1)] })).toBe(true);
    expect(isRedesignNoop(deck, { theme: "Ion", slides: identity })).toBe(false);
    expect(isRedesignNoop(deck, { theme: null, slides: [identity[1], identity[0], identity[2]] })).toBe(false);
    expect(isRedesignNoop(deck, { theme: null, slides: [{ source: 0, layout: "blank" }, ...identity.slice(1)] })).toBe(false);
  });
});
