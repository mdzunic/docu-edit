// The open state of a deck's "Redesign…" (SPEC-115) — wiring only. The
// designer itself (views/design: the dialog, its session and lib/design) is
// loaded through the deck-designer experiment's gates, never from here: the
// slides menu bar's gate renders the menu item, which the menu routes to
// `show`, and PresentationDialogs' gate loads DeckDesignerDialog in redesign
// mode with `host` — the view's model and its record/commit — while `open`.
import { computed, ref, type Ref } from "vue";
import type { PresentationModel } from "../../lib/presentation/model";

export interface DeckRedesignDeps {
  model: Ref<PresentationModel | null>;
  currentIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
}

export function useDeckRedesign(deps: DeckRedesignDeps) {
  const open = ref(false);
  /** A deck with no slides has nothing to arrange. */
  const canRedesign = computed(() => (deps.model.value?.slides.length ?? 0) > 0);

  function show(): void {
    if (canRedesign.value) {
      open.value = true;
    }
  }

  function close(): void {
    open.value = false;
  }

  const host = { model: deps.model, currentIndex: deps.currentIndex, record: deps.record, commit: deps.commit };

  return { open, canRedesign, show, close, host };
}
