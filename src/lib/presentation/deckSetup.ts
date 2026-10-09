// The framework-free deck-setup command layer (SPEC-090): the slide-size
// rescale maths, the first-slide-number origin, and the deck-metadata write-
// back that the "Deck setup" dialog's controller sequences (record → mutate →
// commit). Each command marks the deck `deckDirty` so the writers rebuild the
// deck-level parts (pptx `p:sldSz`/`firstSlideNum`/`core.xml`, odp page-layout
// properties + `meta.xml`); a `"scale"` resize also dirties every slide so the
// per-format writers re-emit their geometry.
//
// Layering rule (presentation/boundary.spec.ts): imports only presentation
// siblings and src/lib/core.
import type { PageSize } from "../core/pageSetup";
import type { DocumentProperties } from "../core/documentProperties";
import type { PresentationModel, Slide, SlideFrame, SlideParagraph } from "./model";
import { DEFAULT_SIZE_PT, markSlideEdited } from "./text";

/** How a resize treats existing content. `"scale"` scales-and-centres every
 * frame (and its text) uniformly; `"keep"` leaves all geometry alone. */
export type RescaleMode = "scale" | "keep";

/** Font-size clamp for a scaled run (0.5 pt steps). The bounds match
 * `PresentationToolbar.vue`'s `PRESENTATION_FONT_SIZE_MIN`/`MAX`, restated here
 * because `lib/presentation` may not import that component. */
const SCALED_SIZE_MIN_PT = 6;
const SCALED_SIZE_MAX_PT = 96;

/** cm rounded to 3 decimals — the repo's `r3` convention (SPEC-082), so a
 * build → re-parse round-trip stays within ±0.01 cm. */
function r3(value: number): number {
  return Math.round(value * 1000) / 1000;
}

/** Applies a new slide size to the deck. `"scale"` scales every frame's bounds
 * and text uniformly by `s = min(newW/oldW, newH/oldH)` and centres the result;
 * `"keep"` changes only the deck size. Both set `deckDirty`; re-applying the
 * same size is a geometry no-op (SPEC-090). */
export function applyDeckSize(model: PresentationModel, size: PageSize, mode: RescaleMode): void {
  const oldW = model.slideWidthCm;
  const oldH = model.slideHeightCm;
  if (mode === "scale") {
    const s = Math.min(size.widthCm / oldW, size.heightCm / oldH);
    const dx = (size.widthCm - oldW * s) / 2;
    const dy = (size.heightCm - oldH * s) / 2;
    for (const slide of model.slides) {
      for (const frame of slide.frames) {
        scaleFrame(frame, s, dx, dy);
      }
      markSlideEdited(slide, model.format);
    }
  }
  model.slideWidthCm = r3(size.widthCm);
  model.slideHeightCm = r3(size.heightCm);
  model.deckDirty = true;
}

/** A copy of `slide` with every frame's bounds (group children included) and
 * text scaled by `factor` and translated by the offsets — the per-slide half
 * of the rescale maths, shared with SPEC-099's slide import (a slide arriving
 * from a deck of another size). Pure: the argument is not mutated. */
export function scaleSlideGeometry(slide: Slide, factor: number, offsetXCm: number, offsetYCm: number): Slide {
  const copy = structuredClone(slide);
  for (const frame of copy.frames) {
    scaleFrame(frame, factor, offsetXCm, offsetYCm);
  }
  return copy;
}

/** Scales one frame's bounds and text, recursing into group children (already
 * slide-space). Group/foreign `xml` is NOT rewritten — only the frame's own
 * bounds move; the writers re-emit the verbatim markup placed by those bounds. */
function scaleFrame(frame: SlideFrame, s: number, dx: number, dy: number): void {
  frame.xCm = r3(frame.xCm * s + dx);
  frame.yCm = r3(frame.yCm * s + dy);
  frame.widthCm = r3(frame.widthCm * s);
  frame.heightCm = r3(frame.heightCm * s);
  if (frame.kind === "text") {
    scaleParagraphs(frame.paragraphs, s);
  } else if (frame.kind === "shape") {
    if (frame.paragraphs) {
      scaleParagraphs(frame.paragraphs, s);
    }
  } else if (frame.kind === "group") {
    // Children carry their own slide-space bounds; their centring offset is
    // absorbed by the group's own move, so nested frames scale WITHOUT a
    // second dx/dy translation. Scaling is a child-geometry change, so a
    // verbatim group's xml is dropped and the serialisers rebuild it
    // (SPEC-096's discipline: present = verbatim, absent = authored).
    delete frame.xml;
    for (const child of frame.children) {
      scaleFrame(child, s, 0, 0);
    }
  }
}

/** Scales every run's effective `sizePt` by `s`, materialising the default
 * (18) when absent, in 0.5 pt steps clamped to 6..96. */
function scaleParagraphs(paragraphs: SlideParagraph[], s: number): void {
  for (const paragraph of paragraphs) {
    for (const run of paragraph.runs) {
      const base = run.sizePt ?? DEFAULT_SIZE_PT;
      run.sizePt = clamp(Math.round(base * s * 2) / 2, SCALED_SIZE_MIN_PT, SCALED_SIZE_MAX_PT);
    }
  }
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

/** How many frames' (rotated) bounds fall entirely outside the slide rect
 * `0..slideWidthCm × 0..slideHeightCm` — the "keep" mode overflow warning.
 * Group children are not counted individually (the group frame stands in). */
export function framesOutsideSlide(model: PresentationModel): number {
  let count = 0;
  for (const slide of model.slides) {
    for (const frame of slide.frames) {
      if (isFrameOutside(frame, model.slideWidthCm, model.slideHeightCm)) {
        count += 1;
      }
    }
  }
  return count;
}

function isFrameOutside(frame: SlideFrame, slideWidthCm: number, slideHeightCm: number): boolean {
  const { minX, minY, maxX, maxY } = rotatedBounds(frame);
  // Outside iff the frame's axis-aligned (rotated) bounding box does not
  // overlap the slide rectangle at all.
  return !(minX < slideWidthCm && maxX > 0 && minY < slideHeightCm && maxY > 0);
}

interface Bounds {
  minX: number;
  minY: number;
  maxX: number;
  maxY: number;
}

function rotatedBounds(frame: SlideFrame): Bounds {
  const { xCm, yCm, widthCm, heightCm } = frame;
  const rotation = frame.rotationDeg ?? 0;
  if (rotation === 0) {
    return { minX: xCm, minY: yCm, maxX: xCm + widthCm, maxY: yCm + heightCm };
  }
  const cx = xCm + widthCm / 2;
  const cy = yCm + heightCm / 2;
  const rad = (rotation * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const corners = [
    [xCm, yCm],
    [xCm + widthCm, yCm],
    [xCm, yCm + heightCm],
    [xCm + widthCm, yCm + heightCm],
  ];
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const [px, py] of corners) {
    const rx = cx + (px - cx) * cos - (py - cy) * sin;
    const ry = cy + (px - cx) * sin + (py - cy) * cos;
    minX = Math.min(minX, rx);
    minY = Math.min(minY, ry);
    maxX = Math.max(maxX, rx);
    maxY = Math.max(maxY, ry);
  }
  return { minX, minY, maxX, maxY };
}

/** Sets the slide-number origin (1..9999). `1` (the default) clears the field
 * so only a value ≠ 1 is stored; either way the deck goes `deckDirty` (so an
 * odp deck whose only change is this still rewrites its deck-level parts, even
 * though odp does not serialise the number). */
export function setFirstSlideNumber(model: PresentationModel, n: number): void {
  if (n === 1) {
    delete model.firstSlideNumber;
  } else {
    model.firstSlideNumber = n;
  }
  model.deckDirty = true;
}

/** Replaces the deck's metadata with `properties` (the dialog's draft, which
 * carries the untouched created/modified/custom/extra fields straight back) and
 * marks the deck `deckDirty`. */
export function applyDeckProperties(model: PresentationModel, properties: DocumentProperties): void {
  model.properties = properties;
  model.deckDirty = true;
}
