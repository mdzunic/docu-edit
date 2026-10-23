// The slide-stage re-render plumbing, lifted out of PresentationView so the
// view stays a wiring hub within its size budget. The presentation model is
// plain (non-reactive) data behind a shallowRef, so an in-place mutation (a
// structural command, a frame move, a style change) is invisible to the
// renderer until the slide and its frames get fresh object identities. These
// three helpers own that trick: `bumpCurrentSlide` re-identifies the current
// slide, `refresh` re-renders mid-gesture, and `commit` re-renders a finished
// edit and reports it through `onCommit` — since SPEC-103 the deck's autosave
// (drafts on SPEC-010's cadence, commits at its commit points) rather than a
// write-through per gesture. Every editing composable in the view (undo,
// slide/frame commands, notes, transitions, animations) shares them.
import { nextTick, triggerRef, type Ref, type ShallowRef } from "vue";
import type { PresentationModel, SlideFrame } from "../../lib/presentation/model";

/** A fresh object identity for a frame AND its group descendants — a group
 * command (rotate, resize, ungroup-undo) mutates the children in place, and a
 * shallow spread would hand the stage the same child objects, leaving them
 * rendered stale (SPEC-096). */
function reidentifyFrame(frame: SlideFrame): SlideFrame {
  return frame.kind === "group" ? { ...frame, children: frame.children.map(reidentifyFrame) } : { ...frame };
}

export interface SlideRenderDeps {
  model: ShallowRef<PresentationModel | null>;
  currentIndex: Ref<number>;
  /** Recompute the fit-to-workspace scale after a commit (a deleted slide can
   * change the surface). */
  recomputeScale: () => void;
  /** Called at the end of every `commit()` — the deck's edit signal
   * (useDeckSaving's noteEdit: autosave's dirty flag and the version-history
   * cadence). `refresh()` never calls it, so a mid-drag preview arms nothing
   * (SPEC-103). */
  onCommit?: () => void;
}

export function useSlideRender(deps: SlideRenderDeps) {
  // SPEC-095: `model.media` is one IndexedDB record whose write latch skips only
  // while the array reference is unchanged, so re-slicing it on every commit
  // rewrites a 20 MB clip whole on every nudge/keystroke. This flag gates the
  // re-slice: a media-mutating command sets it via `bumpMedia()` before its
  // `commit()`, and only then is the media reference (and its store record)
  // refreshed. Cleared each commit. SPEC-103 adopts the same gate for the
  // deck's drafts: with `opaque` split out too, every draft tick after a
  // session's first writes the shell alone.
  let mediaDirty = false;

  /** Marks `model.media` as changed so the next `commit()` re-slices it (and the
   * store persists the record). Every path that mutates `model.media` — an image
   * insert (SPEC-082), a media insert, "Replace poster…" — must call this before
   * `commit()`, or a just-inserted clip would not render (AC-13). */
  function bumpMedia(): void {
    mediaDirty = true;
  }

  /** Replaces the current slide + its frames with fresh object identities so
   * the stage re-renders their geometry/style (the shallowRef trick). */
  function bumpCurrentSlide(): void {
    const model = deps.model.value;
    const slide = model?.slides[deps.currentIndex.value];
    if (!model || !slide) return;
    model.slides[deps.currentIndex.value] = { ...slide, frames: slide.frames.map(reidentifyFrame) };
  }

  /** Re-render the stage without reporting an edit — the live preview during
   * a drag or resize gesture; `commit` reports it once on release. */
  function refresh(): void {
    const model = deps.model.value;
    if (!model) return;
    bumpCurrentSlide();
    model.slides = model.slides.slice();
    triggerRef(deps.model);
  }

  /** Re-render the shallowRef after a finished mutation and report the edit
   * (`onCommit`). The selection is clamped in case a delete dropped it past
   * the tail. */
  function commit(): void {
    const model = deps.model.value;
    if (!model) return;
    if (deps.currentIndex.value > model.slides.length - 1) {
      deps.currentIndex.value = Math.max(0, model.slides.length - 1);
    }
    // A fresh array reference so the filmstrip/stage props re-render (commands
    // splice in place); the model object identity is kept so the store record
    // and undo snapshots stay in step.
    bumpCurrentSlide();
    model.slides = model.slides.slice();
    // A fresh media reference ONLY when a media-mutating command asked for it
    // (bumpMedia), so the stage's media resolver rebuilds and a just-inserted
    // image/clip renders immediately — while a plain nudge/keystroke leaves the
    // reference (and the store's media record) untouched (SPEC-095 AC-12).
    if (mediaDirty && model.media) {
      model.media = model.media.slice();
    }
    mediaDirty = false;
    triggerRef(deps.model);
    void nextTick(deps.recomputeScale);
    deps.onCommit?.();
  }

  return {
    commit,
    refresh,
    bumpCurrentSlide,
    bumpMedia,
  };
}
