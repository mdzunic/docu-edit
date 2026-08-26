// The slide-transition editing session (SPEC-085) — the presentation view's
// wiring for the Transition side panel: the persisted open/closed state (a
// `presentationTransitionsOpen` preference), the current slide's transition
// fields, the field-edit commands (each one undo entry + persist), the two
// confirmation flows (replacing a preserved out-of-subset transition, and the
// batched Apply-to-all), and the on-stage Preview. The heavy lifting — the
// two-None-states and preserved-replacement rules — lives in the framework-free
// lib module; this composable only sequences record → mutate → commit and holds
// the transient confirmation/preview UI state.
import { computed, ref, shallowRef, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide, SlideTransition, TransitionDirection, TransitionType } from "../../lib/presentation/model";
import {
  DEFAULT_TRANSITION_DURATION_MS,
  MAX_TRANSITION_DURATION_MS,
  MIN_TRANSITION_DURATION_MS,
  usesDirection,
} from "../../lib/presentation/transitions";
import {
  applyTransitionToAll,
  applyTransitionToSlide,
  hasPreservedTransition,
  slidesWithPreservedTransition,
} from "../../lib/presentation/slideTransition";
import { usePreferences } from "../../composables/usePreferences";

/** The auto-advance delay a freshly-ticked "Auto-advance" checkbox seeds. */
export const DEFAULT_AUTO_ADVANCE_MS = 3000;

export interface TransitionsDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
}

interface PreviewState {
  outgoing: Slide | undefined;
  incoming: Slide;
  transition: SlideTransition;
}

export function useTransitions(deps: TransitionsDeps) {
  const { preferences, setPreference } = usePreferences();

  const open = computed(() => preferences.presentationTransitionsOpen);
  function toggle(): void {
    setPreference("presentationTransitionsOpen", !open.value);
  }

  const transition = computed<SlideTransition | null>(() => deps.currentSlide.value?.transition ?? null);
  /** True when the slide holds a preserved out-of-subset transition (the honest
   * "original transition kept" note; Decisions 1.i, AC-3). */
  const isPreserved = computed(() => {
    const slide = deps.currentSlide.value;
    return !!slide && !slide.transition && hasPreservedTransition(slide);
  });

  const selectedType = computed<TransitionType | "none">(() => transition.value?.type ?? "none");
  const direction = computed<TransitionDirection>(() => transition.value?.direction ?? "left");
  const directionEnabled = computed(() => usesDirection(selectedType.value));
  const durationSeconds = computed(() => (transition.value?.durationMs ?? DEFAULT_TRANSITION_DURATION_MS) / 1000);
  const advanceOnClick = computed(() => transition.value?.advanceOnClick !== false);
  const autoAdvanceEnabled = computed(() => transition.value?.advanceAfterMs !== undefined);
  const autoAdvanceSeconds = computed(() => (transition.value?.advanceAfterMs ?? DEFAULT_AUTO_ADVANCE_MS) / 1000);

  /** Apply-to-all is enabled iff the current slide has a set transition
   * (any effect or explicit none) — never while unset (AC-9). */
  const canApplyAll = computed(() => !!deps.currentSlide.value?.transition);

  // --- confirmation state ----------------------------------------------------
  const pendingReplace = shallowRef<{ label: string; next: SlideTransition } | null>(null);
  /** > 0 while the batched Apply-to-all confirmation is shown; the count of
   * slides whose original transition would be replaced. */
  const pendingApplyAllCount = ref(0);

  function baseTransition(): SlideTransition {
    return deps.currentSlide.value?.transition ? { ...deps.currentSlide.value.transition } : { type: "none" };
  }

  /** Routes a field change: a slide with a preserved out-of-subset transition
   * gates the FIRST change behind an inline confirmation (AC-8); otherwise it
   * commits straight away. */
  function change(label: string, next: SlideTransition): void {
    const slide = deps.currentSlide.value;
    if (!slide || !deps.model.value) {
      return;
    }
    if (hasPreservedTransition(slide) && !slide.transition) {
      pendingReplace.value = { label, next };
      return;
    }
    commitTransition(label, next);
  }

  function commitTransition(label: string, next: SlideTransition): void {
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!slide || !model) {
      return;
    }
    deps.record(label);
    applyTransitionToSlide(slide, next, model.format);
    deps.commit();
  }

  function confirmReplace(): void {
    const pending = pendingReplace.value;
    pendingReplace.value = null;
    if (pending) {
      commitTransition(pending.label, pending.next);
    }
  }
  function cancelReplace(): void {
    pendingReplace.value = null;
  }

  // --- field commands --------------------------------------------------------
  function setEffect(type: TransitionType | "none"): void {
    const base = baseTransition();
    const next: SlideTransition = { type };
    if (usesDirection(type)) {
      next.direction = base.direction ?? "left";
    }
    if (base.durationMs !== undefined) next.durationMs = base.durationMs;
    if (base.advanceOnClick === false) next.advanceOnClick = false;
    if (base.advanceAfterMs !== undefined) next.advanceAfterMs = base.advanceAfterMs;
    change("Transition effect", next);
  }
  function setDirection(dir: TransitionDirection): void {
    change("Transition direction", { ...baseTransition(), direction: dir });
  }
  function setDurationSeconds(seconds: number): void {
    const ms = Math.min(
      MAX_TRANSITION_DURATION_MS,
      Math.max(MIN_TRANSITION_DURATION_MS, Math.round(seconds * 1000)),
    );
    change("Transition duration", { ...baseTransition(), durationMs: ms });
  }
  function setAdvanceOnClick(value: boolean): void {
    const next = baseTransition();
    if (value) {
      delete next.advanceOnClick;
    } else {
      next.advanceOnClick = false;
    }
    change("Transition advance on click", next);
  }
  function setAutoAdvanceEnabled(value: boolean): void {
    const base = baseTransition();
    const next = { ...base };
    if (value) {
      next.advanceAfterMs = base.advanceAfterMs ?? DEFAULT_AUTO_ADVANCE_MS;
    } else {
      delete next.advanceAfterMs;
    }
    change("Transition auto-advance", next);
  }
  function setAutoAdvanceSeconds(seconds: number): void {
    const ms = Math.max(0, Math.round(seconds * 1000));
    change("Transition auto-advance", { ...baseTransition(), advanceAfterMs: ms });
  }

  // --- apply to all ----------------------------------------------------------
  function applyToAll(): void {
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!slide || !slide.transition || !model) {
      return;
    }
    const count = slidesWithPreservedTransition(model);
    if (count > 0) {
      pendingApplyAllCount.value = count;
      return;
    }
    performApplyAll();
  }
  function performApplyAll(): void {
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!slide || !slide.transition || !model) {
      return;
    }
    deps.record("Apply transition to all slides");
    applyTransitionToAll(model, slide.transition);
    pendingApplyAllCount.value = 0;
    deps.commit();
  }
  function confirmApplyAll(): void {
    performApplyAll();
  }
  function cancelApplyAll(): void {
    pendingApplyAllCount.value = 0;
  }

  // --- preview ---------------------------------------------------------------
  const previewState = shallowRef<PreviewState | null>(null);
  function preview(): void {
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!slide || !model) {
      return;
    }
    const index = deps.currentIndex.value;
    const outgoing = index > 0 ? model.slides[index - 1] : undefined;
    previewState.value = { outgoing, incoming: slide, transition: slide.transition ?? { type: "none" } };
  }
  function onPreviewDone(): void {
    previewState.value = null;
  }

  return {
    open,
    toggle,
    transition,
    isPreserved,
    selectedType,
    direction,
    directionEnabled,
    durationSeconds,
    advanceOnClick,
    autoAdvanceEnabled,
    autoAdvanceSeconds,
    canApplyAll,
    pendingReplace,
    pendingApplyAllCount,
    setEffect,
    setDirection,
    setDurationSeconds,
    setAdvanceOnClick,
    setAutoAdvanceEnabled,
    setAutoAdvanceSeconds,
    confirmReplace,
    cancelReplace,
    applyToAll,
    confirmApplyAll,
    cancelApplyAll,
    preview,
    previewState,
    onPreviewDone,
  };
}
