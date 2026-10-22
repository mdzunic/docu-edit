// The slide-transition editing commands (SPEC-085), kept in the framework-free
// lib layer so the panel composable stays thin and the two-None-states / apply-
// to-all / preserved-replacement rules are unit-testable without a component.
// The panel decides WHEN to run these (and gates the destructive ones behind
// confirmations); this module decides WHAT they do to the model.
import type { PresentationModel, Slide, SlideTransition } from "./model";
import { markSlideEdited } from "./text";

/** Whether a slide carries a preserved out-of-subset transition — the honest
 * "original transition kept" state (Decisions 1.i). pptx keeps the source XML
 * under `p:transition`; odp marks it with `draw:transition` (its bytes survive
 * via the page's retained drawing-page style). Distinct from a plain slide
 * (no transition at all) and from an explicit `type:"none"`. */
export function hasPreservedTransition(slide: Slide): boolean {
  return !!(slide.preserved && (slide.preserved["p:transition"] || slide.preserved["draw:transition"]));
}

/** Applies a transition to one slide: sets the model value (a copy), drops any
 * preserved out-of-subset transition it replaces, and marks it edited so the
 * writers rebuild it (odp also drops `sourceXml`). */
export function applyTransitionToSlide(
  slide: Slide,
  transition: SlideTransition,
  format: PresentationModel["format"],
): void {
  slide.transition = { ...transition };
  if (slide.preserved) {
    delete slide.preserved["p:transition"];
    delete slide.preserved["draw:transition"];
    if (Object.keys(slide.preserved).length === 0) {
      delete slide.preserved;
    }
  }
  markSlideEdited(slide, format);
}

/** Copies a transition onto every slide in the deck (including the source),
 * replacing preserved out-of-subset transitions on the way (Apply to all,
 * AC-10). */
export function applyTransitionToAll(model: PresentationModel, transition: SlideTransition): void {
  for (const slide of model.slides) {
    applyTransitionToSlide(slide, transition, model.format);
  }
}

/** How many slides (of the whole deck) hold a preserved out-of-subset
 * transition an Apply-to-all would replace — the count the batched
 * confirmation names (shown only when > 0). */
export function slidesWithPreservedTransition(model: PresentationModel): number {
  return model.slides.filter(hasPreservedTransition).length;
}
