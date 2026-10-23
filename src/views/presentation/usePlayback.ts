// The playback step machine (SPEC-084) — the NORMATIVE advance pipeline for
// downstream specs. It owns "where are we in the show": a `PlaybackStep`
// ({ slideIndex, buildIndex }) plus an `atEnd` flag for the end screen. This
// spec has no builds (builds are [] until SPEC-086) and no transitions
// (SPEC-085), so advance/back move whole slides; those specs wrap this same
// machine — SPEC-085 delays `current` visually inside transitions without
// changing its logic, SPEC-086 populates builds so advance walks them before
// the next slide, and SPEC-088 reimplements this machine in the exported
// file's script. Keeping the machine pure (a step counter over a reactive
// slide count) is what lets all three reuse it. No DOM, no router: the view
// wires those to `current`/`atEnd`.
import { ref, type Ref } from "vue";

/** A position in the show: which slide, and which build within it. `buildIndex`
 * is always 0 in this spec (no builds yet) but is part of the normative shape
 * SPEC-086 fills in. */
export interface PlaybackStep {
  slideIndex: number;
  buildIndex: number;
}

export interface UsePlaybackDeps {
  /** The deck's slide count (reactive). Zero-slide decks never reach playback
   * — the Present button is disabled — but the machine still clamps safely. */
  slideCount: Ref<number>;
  /** `prefers-reduced-motion`, read by the view and passed through so the
   * pipeline exposes it to SPEC-085/086 (this spec has no motion to reduce,
   * but the plumbing must exist). Defaults to a never-reduced ref. */
  reducedMotion?: Ref<boolean>;
  /** How many build groups slide `i` has (SPEC-086). Advance walks these before
   * moving to the next slide; back lands on a slide fully-built. Absent = a
   * deck with no builds (every slide reports 0), preserving SPEC-084 behaviour. */
  buildCountAt?: (slideIndex: number) => number;
  /** The starting slide index (0-based); clamped into range. A hidden start
   * is honoured — a deep link to a hidden slide still shows it (SPEC-104). */
  start?: number;
  /** Whether slide `i` is hidden (SPEC-104): advance/back/goTo step over a
   * hidden slide, and the end screen follows the last VISIBLE slide. Absent
   * = nothing is hidden, preserving SPEC-084 behaviour. */
  isHidden?: (slideIndex: number) => boolean;
}

export interface UsePlayback {
  current: Ref<PlaybackStep>;
  atEnd: Ref<boolean>;
  reducedMotion: Ref<boolean>;
  advance: () => void;
  back: () => void;
  goTo: (slideIndex: number) => void;
  /** SPEC-100: writes the position directly — the ONLY way a follower's
   * machine is written (from the owner's broadcast state). */
  setStep: (step: PlaybackStep, atEnd?: boolean) => void;
}

export function usePlayback(deps: UsePlaybackDeps): UsePlayback {
  const reducedMotion = deps.reducedMotion ?? ref(false);

  function buildCount(slideIndex: number): number {
    return deps.buildCountAt ? Math.max(0, deps.buildCountAt(slideIndex)) : 0;
  }
  function lastIndex(): number {
    return Math.max(0, deps.slideCount.value - 1);
  }
  function clampIndex(index: number): number {
    if (!Number.isFinite(index)) return 0;
    return Math.min(Math.max(Math.trunc(index), 0), lastIndex());
  }
  function hidden(slideIndex: number): boolean {
    return deps.isHidden ? deps.isHidden(slideIndex) : false;
  }
  /** The first visible slide after `slideIndex`, or -1 when none is left. */
  function nextVisible(slideIndex: number): number {
    for (let index = slideIndex + 1; index <= lastIndex(); index += 1) {
      if (!hidden(index)) return index;
    }
    return -1;
  }
  /** The last visible slide before `slideIndex`, or -1 when none is left. */
  function previousVisible(slideIndex: number): number {
    for (let index = Math.min(slideIndex, lastIndex() + 1) - 1; index >= 0; index -= 1) {
      if (!hidden(index)) return index;
    }
    return -1;
  }

  const current = ref<PlaybackStep>({ slideIndex: clampIndex(deps.start ?? 0), buildIndex: 0 });
  const atEnd = ref(false);

  /** Next build group on the slide (SPEC-086), else next slide, else the end
   * screen. Idempotent once `atEnd` — the view turns a further advance into an
   * exit. */
  function advance(): void {
    if (atEnd.value) return;
    const { slideIndex, buildIndex } = current.value;
    if (buildIndex < buildCount(slideIndex)) {
      current.value = { slideIndex, buildIndex: buildIndex + 1 };
      return;
    }
    const next = nextVisible(slideIndex);
    if (next < 0) {
      atEnd.value = true;
      return;
    }
    current.value = { slideIndex: next, buildIndex: 0 };
  }

  /** Previous slide, landed fully-built (builds collapse on back — SPEC-084's
   * rule, AC-22). From the end screen it steps back onto the last (visible)
   * slide. */
  function back(): void {
    if (atEnd.value) {
      atEnd.value = false;
      const last = hidden(lastIndex()) ? previousVisible(lastIndex()) : lastIndex();
      const target = last < 0 ? current.value.slideIndex : last;
      current.value = { slideIndex: target, buildIndex: buildCount(target) };
      return;
    }
    const prev = previousVisible(current.value.slideIndex);
    if (prev < 0) return;
    current.value = { slideIndex: prev, buildIndex: buildCount(prev) };
  }

  /** Jump straight to a slide (Home/End, number-then-Enter), clearing the end
   * screen. Out-of-range indices clamp; a hidden target steps over to the
   * next visible slide (the previous one when none follows). */
  function goTo(slideIndex: number): void {
    atEnd.value = false;
    let target = clampIndex(slideIndex);
    if (hidden(target)) {
      const next = nextVisible(target);
      const prev = next < 0 ? previousVisible(target) : -1;
      target = next >= 0 ? next : prev >= 0 ? prev : target;
    }
    current.value = { slideIndex: target, buildIndex: 0 };
  }

  /** Lands exactly on `step` (SPEC-100): the slide index clamps exactly as
   * goTo clamps it, and the build index into 0..that slide's build count. It
   * is a plain write, not a walk — a hidden slide is NOT stepped over, because
   * the owner whose state this mirrors already resolved hidden slides (and a
   * deep link to one shows it). An unchanged position keeps the same `current`
   * object, so its watchers see no step. */
  function setStep(step: PlaybackStep, end = false): void {
    atEnd.value = end;
    const slideIndex = clampIndex(step.slideIndex);
    const rawBuild = Number.isFinite(step.buildIndex) ? Math.trunc(step.buildIndex) : 0;
    const buildIndex = Math.min(Math.max(rawBuild, 0), buildCount(slideIndex));
    if (current.value.slideIndex !== slideIndex || current.value.buildIndex !== buildIndex) {
      current.value = { slideIndex, buildIndex };
    }
  }

  return { current, atEnd, reducedMotion, advance, back, goTo, setStep };
}
