// The first presentation mutations (SPEC-080), framework-free like every
// other presentation/ module: create a blank deck and add / duplicate /
// delete / move slides. Each structural command mutates `model.slides` in
// place, marks the model `structureDirty` so the writers rebuild the
// structural parts, and returns the index the view should select. New and
// duplicated slides are born WITHOUT `partName`/`sourceXml` retention (new
// slides) or SHARING the source's `partName` (duplicates) — the substitution
// writers key off exactly that to decide "serialise fresh" vs "copy the
// source part's bytes verbatim" (see pptx/odp writer rebuild rules). Text and
// object editing stay out of scope (SPEC-081/082); a blank deck's
// placeholders render their empty state but cannot be typed into yet.
import {
  DEFAULT_PRESENTATION_TITLE,
  DEFAULT_SLIDE_HEIGHT_CM,
  DEFAULT_SLIDE_WIDTH_CM,
  type PlaceholderRole,
  type PresentationModel,
  type Slide,
  type SlideFrame,
  type TextFrame,
} from "./model";
import { retargetAnimations } from "./animations";
import { newSlideFramesForLayout } from "./layouts";

/** cm rounded to 3 decimals — a hair finer than the eye and far coarser than
 * the EMU/ODF length precision the writers round to, so a placeholder's
 * geometry survives a build → re-parse round trip unchanged. */
function r3(value: number): number {
  return Math.round(value * 1000) / 1000;
}

/** A placeholder text frame at a fraction of the slide surface. Roles map to
 * the pptx `phType` the writers emit: title→`title`, subtitle→`subTitle`,
 * body→`body` (with `phIdx: 1`). Born empty — SPEC-081 fills the text. */
function placeholder(
  role: PlaceholderRole,
  phType: string,
  phIdx: number | undefined,
  xFrac: number,
  yFrac: number,
  wFrac: number,
  hFrac: number,
  slideWidthCm: number,
  slideHeightCm: number,
): TextFrame {
  const frame: TextFrame = {
    id: crypto.randomUUID(),
    xCm: r3(xFrac * slideWidthCm),
    yCm: r3(yFrac * slideHeightCm),
    widthCm: r3(wFrac * slideWidthCm),
    heightCm: r3(hFrac * slideHeightCm),
    kind: "text",
    role,
    phType,
    paragraphs: [],
  };
  if (phIdx !== undefined) {
    frame.phIdx = phIdx;
  }
  return frame;
}

/** A fresh, from-scratch blank deck (SPEC-080): 16:9 widescreen, `pptx`
 * format, one slide with a centred title and subtitle placeholder, no
 * retained package. Carrying `"pptx"` from birth is what lets the library
 * badge/route it as a presentation with no new `LibraryFormat`. */
export function createBlankPresentation(): PresentationModel {
  const w = DEFAULT_SLIDE_WIDTH_CM;
  const h = DEFAULT_SLIDE_HEIGHT_CM;
  const slide: Slide = {
    id: crypto.randomUUID(),
    frames: [
      // Title y 32% h 18%, subtitle y 52% h 12%, both x 8% w 84% (the Reference).
      placeholder("title", "title", undefined, 0.08, 0.32, 0.84, 0.18, w, h),
      placeholder("subtitle", "subTitle", undefined, 0.08, 0.52, 0.84, 0.12, w, h),
    ],
  };
  return {
    kind: "presentation",
    id: crypto.randomUUID(),
    title: DEFAULT_PRESENTATION_TITLE,
    format: "pptx",
    slides: [slide],
    slideWidthCm: w,
    slideHeightCm: h,
  };
}

/** A new slide inserted AFTER `afterIndex`. Without `layoutKey`, a title + body
 * slide per the Reference geometry (title x 8% y 8% w 84% h 16%; body x 8% y 30%
 * w 84% h 60%) — SPEC-080's unchanged behaviour. With a `layoutKey` naming a
 * catalogue entry (SPEC-089 AC-8), the slide is born with one empty `TextFrame`
 * per slot of that layout and its `layoutKey` set. Returns the new slide's index
 * (which becomes current). */
export function addSlide(model: PresentationModel, afterIndex: number, layoutKey?: string): number {
  const w = model.slideWidthCm;
  const h = model.slideHeightCm;
  const layout = layoutKey ? (model.layouts ?? []).find((l) => l.key === layoutKey) : undefined;
  const slide: Slide = layout
    ? { id: crypto.randomUUID(), frames: newSlideFramesForLayout(layout), layoutKey: layout.key }
    : {
        id: crypto.randomUUID(),
        frames: [
          placeholder("title", "title", undefined, 0.08, 0.08, 0.84, 0.16, w, h),
          placeholder("body", "body", 1, 0.08, 0.3, 0.84, 0.6, w, h),
        ],
      };
  const at = Math.min(Math.max(afterIndex + 1, 0), model.slides.length);
  model.slides.splice(at, 0, slide);
  model.structureDirty = true;
  return at;
}

/** Re-mints every frame's id (and group children's) so a duplicated slide's
 * frames are independent selection/undo targets (AC-30), recording each
 * old→new id in `idMap` so animation targets can follow (SPEC-086 AC-19). */
function remintFrameIds(frames: SlideFrame[], idMap: Map<string, string>): void {
  for (const frame of frames) {
    const next = crypto.randomUUID();
    idMap.set(frame.id, next);
    frame.id = next;
    if (frame.kind === "group") {
      remintFrameIds(frame.children, idMap);
    }
  }
}

/** An exact copy of slide `index`, inserted immediately after it (AC-29..33).
 * The copy keeps the source's frames (fresh ids), preserved XML and — for a
 * deck opened from a file — its `partName`/`sourceXml`: sharing the source
 * `partName` is the signal the pptx writer copies the source part's bytes
 * verbatim under a fresh name, and the retained `sourceXml` is the odp page
 * body the odp writer re-emits (with its `draw:name` uniqued). Returns the
 * copy's index. */
export function duplicateSlide(model: PresentationModel, index: number): number {
  const source = model.slides[index];
  if (!source) {
    return index;
  }
  const copy: Slide = structuredClone(source);
  copy.id = crypto.randomUUID();
  const idMap = new Map<string, string>();
  remintFrameIds(copy.frames, idMap);
  // Keep the copy's animations, retargeted onto its fresh frame ids (AC-19).
  copy.animations = retargetAnimations(copy.animations, idMap);
  model.slides.splice(index + 1, 0, copy);
  model.structureDirty = true;
  return index + 1;
}

/** Removes slide `index` (allowed down to zero slides). Returns the index of
 * the nearest remaining neighbour to make current (0 when the deck is now
 * empty; the view shows its empty state). */
export function deleteSlide(model: PresentationModel, index: number): number {
  if (index < 0 || index >= model.slides.length) {
    return index;
  }
  model.slides.splice(index, 1);
  model.structureDirty = true;
  return model.slides.length === 0 ? 0 : Math.min(index, model.slides.length - 1);
}

/** Moves the slide at `from` to `to` (clamped), keeping it selected. Returns
 * the moved slide's new index. */
export function moveSlide(model: PresentationModel, from: number, to: number): number {
  if (from < 0 || from >= model.slides.length) {
    return from;
  }
  const target = Math.min(Math.max(to, 0), model.slides.length - 1);
  if (target === from) {
    return from;
  }
  const [moved] = model.slides.splice(from, 1);
  model.slides.splice(target, 0, moved);
  model.structureDirty = true;
  return target;
}

/** The number of frames a cross-format "Download a copy as…" cannot carry
 * through the model (SPEC-080): anything the from-scratch writers cannot yet
 * serialise — every non-placeholder-text frame (images, shapes, groups,
 * foreign objects). Blank/new decks hold only text placeholders, so this is
 * 0 for them; SPEC-081+ shrink it as the frame serialiser grows. */
export function unconvertibleFrameCount(model: PresentationModel): number {
  let count = 0;
  for (const slide of model.slides) {
    for (const frame of slide.frames) {
      if (frame.kind !== "text") {
        count += 1;
      }
    }
  }
  return count;
}

/** A copy of the model with every retained-package field removed
 * (`opaque`, `odpContentEnvelope`, and per-slide `partName`/`sourceXml`/
 * `preserved`/`dirty`), so the writers rebuild every part from the model
 * alone (SPEC-080). The cross-format download path passes this so a
 * `.pptx`-born deck saved as `.odp` never leaks pptx parts and vice versa.
 * Pure — the stored deck keeps its format and retention. */
export function stripRetention(model: PresentationModel): PresentationModel {
  const { opaque: _opaque, odpContentEnvelope: _envelope, ...rest } = model;
  const copy: PresentationModel = structuredClone(rest);
  for (const slide of copy.slides) {
    delete slide.partName;
    delete slide.sourceXml;
    delete slide.preserved;
    delete slide.dirty;
  }
  delete copy.structureDirty;
  return copy;
}

export { DEFAULT_PRESENTATION_TITLE };
