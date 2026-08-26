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
  /** The starting slide index (0-based); clamped into range. */
  start?: number;
}

export interface UsePlayback {
  current: Ref<PlaybackStep>;
  atEnd: Ref<boolean>;
  reducedMotion: Ref<boolean>;
  advance: () => void;
  back: () => void;
  goTo: (slideIndex: number) => void;
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
    if (slideIndex >= lastIndex()) {
      atEnd.value = true;
      return;
    }
    current.value = { slideIndex: slideIndex + 1, buildIndex: 0 };
  }

  /** Previous slide, landed fully-built (builds collapse on back — SPEC-084's
   * rule, AC-22). From the end screen it steps back onto the last slide. */
  function back(): void {
    if (atEnd.value) {
      atEnd.value = false;
      current.value = { slideIndex: lastIndex(), buildIndex: buildCount(lastIndex()) };
      return;
    }
    if (current.value.slideIndex <= 0) return;
    const prev = current.value.slideIndex - 1;
    current.value = { slideIndex: prev, buildIndex: buildCount(prev) };
  }

  /** Jump straight to a slide (Home/End, number-then-Enter), clearing the end
   * screen. Out-of-range indices clamp. */
  function goTo(slideIndex: number): void {
    atEnd.value = false;
    current.value = { slideIndex: clampIndex(slideIndex), buildIndex: 0 };
  }

  return { current, atEnd, reducedMotion, advance, back, goTo };
}
