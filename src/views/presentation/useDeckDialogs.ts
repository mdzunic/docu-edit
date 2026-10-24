// The deck-level dialog controllers, wired together (SPEC-098's size-budget
// extraction — the usePresentationLoad move the spec names had already landed
// as SPEC-096's useDeckLoad, so this is its sanctioned equivalent, the
// useEditingPanels precedent): slide layouts + Reset slide (SPEC-089), deck
// setup (SPEC-090), header & footer (SPEC-093), print & export (SPEC-104) and
// Slide › Redesign… (SPEC-115). They take the same record/commit pair and are
// the controllers PresentationDialogs mounts; the view hands them on whole.
import type { ComputedRef, Ref } from "vue";
import type { PresentationModel, Slide } from "../../lib/presentation/model";
import { useSlideLayouts } from "./useSlideLayouts";
import { useDeckSetup } from "./useDeckSetup";
import { useHeaderFooter } from "./useHeaderFooter";
import { useDeckExport } from "./useDeckExport";
import { useDeckRedesign } from "./useDeckRedesign";

export interface DeckDialogsDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  slideTotal: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
  /** Adds a slide (the layout gallery's "new slide with this layout"). */
  addSlide: (afterIndex?: number, layoutKey?: string) => void;
  reportError: (message: string) => void;
}

export function useDeckDialogs(deps: DeckDialogsDeps) {
  const { model, currentSlide, currentIndex, slideTotal, record, commit, addSlide, reportError } = deps;
  return {
    layouts: useSlideLayouts({ model, currentSlide, currentIndex, record, commit, addSlide }),
    deckSetup: useDeckSetup({ model, currentIndex, record, commit, reportError }),
    headerFooter: useHeaderFooter({ model, currentSlide, currentIndex, record, commit }),
    deckExport: useDeckExport({ model, currentIndex, slideTotal, reportError }),
    // Open state only; the designer loads behind its experiment's gates.
    redesign: useDeckRedesign({ model, currentIndex, record, commit }),
  };
}
