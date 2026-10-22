// The deck-navigation block, extracted from PresentationView (SPEC-099) so
// the view can take the slide sorter's wiring within its size budget: the
// undo-wrapped slide commands, the deck's derived facts (title, counts,
// current slide, media, field bands, the live "Slide 3 of 9"), the "Add
// slide" action and the filmstrip's command/reorder dispatch — everything the
// filmstrip, the sorter's card menu and the stage need to move around the
// deck. It composes the composables that already owned those pieces
// (useSlideCommands, useDeckFacts, useFilmstripCommands) rather than
// replacing them.
import type { Ref, ShallowRef } from "vue";
import type { PresentationModel } from "../../lib/presentation/model";
import { useDeckFacts } from "./useDeckFacts";
import { useFilmstripCommands } from "./useFilmstripCommands";
import { useSlideCommands } from "./useSlideCommands";

export interface DeckNavigationDeps {
  presentation: ShallowRef<PresentationModel | null>;
  currentIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
  /** The layout gallery's two entry points — called lazily, so the gallery
   * (which itself adds slides through these commands) can be built after. */
  openLayoutForNew: () => void;
  openLayoutForChange: () => void;
}

export function useDeckNavigation(deps: DeckNavigationDeps) {
  const { presentation, currentIndex } = deps;
  const slideCommands = useSlideCommands({ model: presentation, currentIndex, record: deps.record, commit: deps.commit });
  const facts = useDeckFacts({ presentation, currentIndex });

  /** "Add slide": a new slide after the current one. */
  function addSlideAction(): void {
    slideCommands.add(currentIndex.value);
  }

  const filmstrip = useFilmstripCommands({
    commands: slideCommands,
    selectSlide: facts.selectSlide,
    openLayoutForNew: () => deps.openLayoutForNew(),
    openLayoutForChange: () => deps.openLayoutForChange(),
  });

  return {
    ...facts,
    slideCommands,
    addSlideAction,
    onFilmstripCommand: filmstrip.onCommand,
    onFilmstripReorder: filmstrip.onReorder,
  };
}
