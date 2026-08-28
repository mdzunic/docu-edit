// The layout gallery session (SPEC-089) — the presentation view's wiring for
// the modal layout picker and Reset slide. Unlike the transitions/animations
// side panels, the gallery is a dialog, so its open state is a plain ref (no
// `preferences` key). This composable holds the open/mode state, exposes the
// deck's own layouts plus the synthesised standard fallbacks as `cards` (a
// fresh array per the useAnimations reactivity rule), and sequences
// record → mutate → commit for apply / new-slide / reset, surfacing the
// "N items kept in place" report an apply could not fully place.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide, SlideLayout } from "../../lib/presentation/model";
import {
  applySlideLayout,
  layoutFor,
  resetSlideLayout,
  synthesiseStandardLayouts,
  type LayoutApplyReport,
} from "../../lib/presentation/layouts";

export type LayoutGalleryMode = "change" | "new";

export interface SlideLayoutsDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
  /** `useSlideCommands.add` — inserts a new slide (optionally from a layout). */
  addSlide: (afterIndex?: number, layoutKey?: string) => void;
}

export function useSlideLayouts(deps: SlideLayoutsDeps) {
  const open = ref(false);
  const mode = ref<LayoutGalleryMode>("change");
  const chosenKey = ref<string | null>(null);
  const report = ref<LayoutApplyReport | null>(null);

  // Fresh array each read (the useAnimations rule): the deck's own layouts, then
  // a synthesised standard layout for each standard kind the deck lacks, marked
  // (by its `standard:` key) as "will be added to this deck".
  const cards = computed<SlideLayout[]>(() => {
    const model = deps.model.value;
    if (!model) {
      return [];
    }
    const own = (model.layouts ?? []).slice();
    const ownKinds = new Set(own.map((l) => l.kind));
    const fallbacks = synthesiseStandardLayouts(model.slideWidthCm, model.slideHeightCm).filter(
      (l) => !ownKinds.has(l.kind),
    );
    return [...own, ...fallbacks];
  });

  const currentKey = computed<string | null>(() => chosenKey.value);

  const canReset = computed(() => {
    const model = deps.model.value;
    const slide = deps.currentSlide.value;
    return !!model && !!slide && layoutFor(model, slide) !== undefined;
  });

  function openForChange(): void {
    mode.value = "change";
    chosenKey.value = deps.currentSlide.value?.layoutKey ?? cards.value[0]?.key ?? null;
    report.value = null;
    open.value = true;
  }
  function openForNew(): void {
    mode.value = "new";
    chosenKey.value = cards.value[0]?.key ?? null;
    report.value = null;
    open.value = true;
  }
  function close(): void {
    open.value = false;
  }
  function choose(key: string): void {
    chosenKey.value = key;
  }

  function ensureLayout(model: PresentationModel, key: string): void {
    // A synthesised fallback the user picks must join the catalogue so
    // applySlideLayout/addSlide resolve it and the writers author its part.
    if ((model.layouts ?? []).some((l) => l.key === key)) {
      return;
    }
    const fallback = cards.value.find((l) => l.key === key);
    if (fallback) {
      model.layouts = [...(model.layouts ?? []), fallback];
    }
  }

  function apply(): void {
    const model = deps.model.value;
    const key = chosenKey.value;
    if (!model || !key) {
      return;
    }
    if (mode.value === "new") {
      ensureLayout(model, key);
      deps.addSlide(deps.currentIndex.value, key);
      close();
      return;
    }
    deps.record("Change slide layout");
    ensureLayout(model, key);
    const result = applySlideLayout(model, deps.currentIndex.value, key);
    deps.commit();
    if (result.applied && result.keptInPlace.length > 0) {
      report.value = result;
    } else {
      report.value = null;
      close();
    }
  }

  function resetSlide(): void {
    const model = deps.model.value;
    if (!model || !canReset.value) {
      return;
    }
    deps.record("Reset slide");
    const result = resetSlideLayout(model, deps.currentIndex.value);
    deps.commit();
    if (result.applied && result.keptInPlace.length > 0) {
      report.value = result;
    } else {
      report.value = null;
      close();
    }
  }

  function dismissReport(): void {
    report.value = null;
    close();
  }

  // Passthroughs the gallery dialog needs (its props are only { controller,
  // slideNumber }): the deck size for thumbnails and the current slide for the
  // report's frame names.
  const widthCm = computed(() => deps.model.value?.slideWidthCm ?? 0);
  const heightCm = computed(() => deps.model.value?.slideHeightCm ?? 0);
  const currentSlide = deps.currentSlide;

  return {
    open,
    mode,
    openForChange,
    openForNew,
    close,
    cards,
    currentKey,
    choose,
    apply,
    resetSlide,
    canReset,
    report,
    dismissReport,
    widthCm,
    heightCm,
    currentSlide,
  };
}
