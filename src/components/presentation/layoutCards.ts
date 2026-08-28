// Typed transforms for LayoutGalleryDialog.vue (SPEC-089). The dialog's SFC is
// linted by espree (runtime `defineProps` only, no TS), so its card view-models
// and the apply-report text live here as plain TypeScript.
import type { LayoutApplyReport } from "../../lib/presentation/layouts";
import { newSlideFramesForLayout } from "../../lib/presentation/layouts";
import type { Slide, SlideLayout } from "../../lib/presentation/model";
import { frameAccessibleName } from "../../views/presentation/useFrameManipulation";

/** One gallery card: the layout, a synthesised preview slide of its slots (for
 * the `<SlideStage>` thumbnail), and whether it is a synthesised fallback the
 * deck does not yet carry ("will be added to this deck"). */
export interface LayoutCard {
  key: string;
  name: string;
  isNew: boolean;
  previewSlide: Slide;
}

/** A synthesised preview `Slide` of a layout's slots — one empty placeholder
 * frame per slot, rendered read-only by `<SlideStage :interactive="false">`. */
function previewSlideFor(layout: SlideLayout): Slide {
  return { id: `preview-${layout.key}`, frames: newSlideFramesForLayout(layout) };
}

/** The card view-models for a set of catalogue layouts. A `standard:`-keyed
 * layout is a not-yet-authored fallback (the "new" badge). */
export function toLayoutCards(layouts: SlideLayout[]): LayoutCard[] {
  return layouts.map((layout) => ({
    key: layout.key,
    name: layout.name,
    isNew: layout.key.startsWith("standard:"),
    previewSlide: previewSlideFor(layout),
  }));
}

/** The accessible names of the frames an apply could not place, for the report
 * (`frameAccessibleName`, e.g. "Placeholder", "Image"). */
export function keptInPlaceNames(report: LayoutApplyReport, slide: Slide | null): string[] {
  if (!slide) {
    return [];
  }
  return report.keptInPlace
    .map((id) => slide.frames.find((f) => f.id === id))
    .filter((f): f is NonNullable<typeof f> => f !== undefined)
    .map((f) => frameAccessibleName(f));
}

/** The report headline, e.g. "2 items kept in place". */
export function reportHeadline(report: LayoutApplyReport): string {
  const n = report.keptInPlace.length;
  return `${n} ${n === 1 ? "item" : "items"} kept in place`;
}
