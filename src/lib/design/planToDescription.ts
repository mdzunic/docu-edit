// A reviewed DeckPlan as SPEC-106's DeckDescription (SPEC-115). The designer
// adds no serialisation: the plan becomes a plain JSON description, which
// `deckFromDescription` validates and builds through the same path
// `PresentationBuilder.fromDescription` takes — so a plan can only produce a
// deck SPEC-106's API would have produced from the same JSON. Text stays text:
// every title, bullet and note is a plain string field, which slideRuns.ts
// turns into SlideRun text (never Markdown, never HTML).
//
// `DeckDescription` has four layouts and no layout kind or theme, so the
// mapping is coarse on purpose and `applyPlanDesign` (applyDesign.ts) sets
// the plan's theme and per-slide kinds on the built model afterwards:
//   title                                → "title", its one bullet the subtitle
//   blank, with no title and no bullets  → "blank"
//   every other slide                    → "bullets", levels kept
// A title slide's bullets past the first are not built — validatePlan has
// already recorded each as a drop, and the plan step lists them before Build.
import type { DeckDescription, DeckSlideDescription } from "../generate/deckDescription";
import type { BulletRow } from "../generate/deckSlides";
import type { DeckPlan, PlannedBullet, PlannedSlide } from "./plan";

function bulletRow(bullet: PlannedBullet): BulletRow {
  return typeof bullet === "string" ? bullet : { text: bullet.text, level: bullet.level };
}

function notesOf(slide: PlannedSlide): { notes?: string[] } {
  return slide.note !== undefined && slide.note.trim() !== "" ? { notes: [slide.note] } : {};
}

function slideDescription(slide: PlannedSlide): DeckSlideDescription {
  if (slide.layout === "title") {
    const subtitle = slide.bullets[0];
    return {
      layout: "title",
      title: slide.title,
      ...(subtitle === undefined ? {} : { subtitle: typeof subtitle === "string" ? subtitle : subtitle.text }),
      ...notesOf(slide),
    };
  }
  if (slide.layout === "blank" && slide.title.trim() === "" && slide.bullets.length === 0) {
    return { layout: "blank", ...notesOf(slide) };
  }
  return { layout: "bullets", title: slide.title, bullets: slide.bullets.map(bulletRow), ...notesOf(slide) };
}

export function planToDescription(plan: DeckPlan): DeckDescription {
  return {
    title: plan.title,
    size: plan.size,
    slides: plan.slides.map(slideDescription),
  };
}
