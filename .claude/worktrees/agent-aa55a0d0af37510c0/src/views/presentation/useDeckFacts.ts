// The deck's derived facts — the title, badge and format the header shows, the
// slide count and current slide, the live region's "Slide 3 of 9", the media
// list and each slide's resolved field band. All read off the deck's shallowRef
// so a commit (triggerRef) recomputes them live.
//
// Lifted out of PresentationView so the view stays a wiring hub within its line
// budget (SPEC-090's rule: when it fails, extract). Nothing here decides
// anything — it is the shape the deck is read in, which is why it is a
// composable of computeds rather than a set of handlers.
import { computed, type ComputedRef, type Ref, type ShallowRef } from "vue";
import { formatBadgeLabel, presentationFormatLabel } from "../../lib/format";
import { stageFields } from "../../components/presentation/slideRender";
import { now } from "../../lib/clock";
import type { PresentationModel, Slide } from "../../lib/presentation/model";

export interface DeckFactsDeps {
  presentation: ShallowRef<PresentationModel | null>;
  currentIndex: Ref<number>;
}

export function useDeckFacts(deps: DeckFactsDeps) {
  const { presentation, currentIndex } = deps;

  const title = computed(() => presentation.value?.title ?? "No document");
  const badge = computed(() => (presentation.value ? formatBadgeLabel(presentation.value.format) : ""));
  const formatName = computed(() =>
    presentation.value ? presentationFormatLabel(presentation.value.format) : "",
  );
  const slideTotal = computed(() => presentation.value?.slides.length ?? 0);
  const slideCountLabel = computed(() =>
    slideTotal.value === 1 ? "1 slide" : `${slideTotal.value} slides`,
  );

  const currentSlide: ComputedRef<Slide | null> = computed(
    () => presentation.value?.slides[currentIndex.value] ?? null,
  );
  const liveLabel = computed(() =>
    presentation.value ? `Slide ${currentIndex.value + 1} of ${slideTotal.value}` : "",
  );
  const deckMedia = computed(() => presentation.value?.media ?? []);

  // SPEC-093: each slide's resolved field band, recomputed on every commit so
  // footers/numbers/dates stay live (AC-27/30/31).
  const deckFields = computed(() => {
    const model = presentation.value;
    return model ? model.slides.map((_, index) => stageFields(model, index, now())) : [];
  });

  function selectSlide(index = 0): void {
    if (index >= 0 && index < slideTotal.value) {
      currentIndex.value = index;
    }
  }

  return {
    title,
    badge,
    formatName,
    slideTotal,
    slideCountLabel,
    currentSlide,
    liveLabel,
    deckMedia,
    deckFields,
    selectSlide,
  };
}
