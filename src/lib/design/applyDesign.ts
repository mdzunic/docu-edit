// The deck designer's model mutations (SPEC-115) — plain model edits through
// SPEC-089's `applySlideLayout`, SPEC-094's `applyTheme` and SPEC-080's
// `moveSlide`; no serialisation. `applyPlanDesign` finishes a freshly built
// deck (DeckDescription has no theme or layout kind); `applyRedesign`
// rearranges the open deck, which the view wraps in one record → commit().
//
// A planned kind resolves to the deck's own layout: the first entry of
// `model.layouts` with that kind, else the `standard:{kind}` layout that
// `synthesiseStandardLayouts` produces at the deck's size — added to the
// catalogue as SPEC-089's gallery adds it — so a slide given a kind always has
// a `layoutFor` result of exactly that kind, and a model naming a layout the
// product does not have can never produce a deck that references one.
import { BUILT_IN_THEMES, DEFAULT_THEME, type DocumentTheme } from "../core/theme";
import { applySlideLayout, layoutFor, synthesiseStandardLayouts } from "../presentation/layouts";
import type { PresentationModel, Slide, SlideFrame, SlideLayoutKind } from "../presentation/model";
import { moveSlide } from "../presentation/slides";
import { applyTheme, deckThemeOf } from "../presentation/theme";
import { isPlanLayoutKind, type DeckPlan, type DeckRedesign, type PlanLayoutKind } from "./plan";
import type { RedesignSlideSummary } from "./prompt";

/** The longest a slide title is quoted to the model (redesign requests). */
export const MAX_REDESIGN_TITLE_CHARS = 200;

/** The built-in theme with this registry name (Office for anything else —
 * validation has already normalised the name). */
export function builtInTheme(name: string): DocumentTheme {
  return BUILT_IN_THEMES.find((theme) => theme.name === name) ?? DEFAULT_THEME;
}

/** The kind of the layout a slide uses, or undefined when its `layoutKey`
 * resolves to nothing. */
export function currentLayoutKind(model: PresentationModel, slide: Slide): SlideLayoutKind | undefined {
  return layoutFor(model, slide)?.kind;
}

/** What a redesign row starts as: the slide's current kind when it is one of
 * the six, else null ("Keep current"). */
export function seedLayoutKind(model: PresentationModel, slide: Slide): PlanLayoutKind | null {
  const kind = currentLayoutKind(model, slide);
  return isPlanLayoutKind(kind) ? kind : null;
}

/** The slide's title-placeholder text ("" when it has none), paragraphs
 * joined by a space. */
export function slideTitleText(slide: Slide): string {
  const frame = slide.frames.find((f) => f.kind === "text" && f.role === "title");
  if (!frame || frame.kind !== "text") {
    return "";
  }
  return frame.paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(""))
    .join(" ")
    .trim();
}

/** The open deck's slides as a redesign request lists them: each one's
 * 0-based index, its title-placeholder text cut at MAX_REDESIGN_TITLE_CHARS
 * ("" when none) and its current layout kind ("other" for one outside the six
 * or none) — never body text, notes or any other frame. */
export function summariseSlides(model: PresentationModel): RedesignSlideSummary[] {
  return model.slides.map((slide, index) => ({
    index,
    title: Array.from(slideTitleText(slide)).slice(0, MAX_REDESIGN_TITLE_CHARS).join(""),
    layout: currentLayoutKind(model, slide) ?? "other",
  }));
}

/** The layout key a kind resolves to on this deck, adding the synthesised
 * standard layout to `model.layouts` when the deck has none of that kind. */
export function resolveLayoutKey(model: PresentationModel, kind: PlanLayoutKind): string {
  const own = (model.layouts ?? []).find((layout) => layout.kind === kind);
  if (own) {
    return own.key;
  }
  const standard = synthesiseStandardLayouts(model.slideWidthCm, model.slideHeightCm).find((layout) => layout.kind === kind);
  if (!standard) {
    throw new Error(`no standard layout of kind ${kind}`);
  }
  model.layouts = [...(model.layouts ?? []), standard];
  return standard.key;
}

function isEmptyPlaceholder(frame: SlideFrame): boolean {
  return (
    frame.kind === "text" &&
    frame.role !== undefined &&
    !frame.paragraphs.some((paragraph) => paragraph.runs.some((run) => run.text.length > 0))
  );
}

/** Applies `kind` to slide `index` and removes every placeholder the layout
 * kept in place that carries no text — so a titleOnly slide has no empty body
 * box. A frame an animation targets is left alone. */
function applyKind(model: PresentationModel, index: number, kind: PlanLayoutKind): void {
  const report = applySlideLayout(model, index, resolveLayoutKey(model, kind));
  if (!report.applied) {
    return;
  }
  const slide = model.slides[index];
  const kept = new Set(report.keptInPlace);
  const animated = new Set((slide.animations ?? []).map((animation) => animation.frameId));
  slide.frames = slide.frames.filter((frame) => !(kept.has(frame.id) && !animated.has(frame.id) && isEmptyPlaceholder(frame)));
}

/** Sets a freshly built deck's theme and each slide's kind (slide i takes
 * plan slide i — `planToDescription` maps one to one). */
export function applyPlanDesign(model: PresentationModel, plan: DeckPlan): void {
  applyTheme(model, builtInTheme(plan.theme));
  plan.slides.forEach((planned, index) => {
    if (index < model.slides.length) {
      applyKind(model, index, planned.layout);
    }
  });
  if (!model.opaque) {
    // As the builder's own settle(): `dirty` only means something for a deck
    // with a retained package to rebuild against, and a built deck has none.
    for (const slide of model.slides) {
      delete slide.dirty;
    }
  }
}

/** Whether applying `redesign` would change nothing: the order is the
 * deck's own, no row's kind differs from its slide's, and the theme is kept
 * (or is the deck's current one). */
export function isRedesignNoop(model: PresentationModel, redesign: DeckRedesign): boolean {
  const sameOrder = redesign.slides.every((entry, index) => entry.source === index);
  const sameKinds = redesign.slides.every(
    (entry) => entry.layout === null || entry.layout === currentLayoutKind(model, model.slides[entry.source]),
  );
  const sameTheme = redesign.theme === null || redesign.theme === deckThemeOf(model).name;
  return sameOrder && sameKinds && sameTheme;
}

/** Moves the slides into the reviewed order, applies each row's kind where it
 * differs from the slide's current one, and applies a chosen theme. Never
 * adds, deletes or rewrites a slide: every slide's text, notes and
 * non-placeholder frames are untouched. `redesign` must be validated against
 * this deck (one entry per slide, every index once). */
export function applyRedesign(model: PresentationModel, redesign: DeckRedesign): void {
  const original = model.slides.slice();
  redesign.slides.forEach((entry, target) => {
    const from = model.slides.indexOf(original[entry.source]);
    if (from !== target) {
      moveSlide(model, from, target);
    }
  });
  redesign.slides.forEach((entry, index) => {
    if (entry.layout !== null && entry.layout !== currentLayoutKind(model, model.slides[index])) {
      applyKind(model, index, entry.layout);
    }
  });
  if (redesign.theme !== null && redesign.theme !== deckThemeOf(model).name) {
    applyTheme(model, builtInTheme(redesign.theme));
  }
}
