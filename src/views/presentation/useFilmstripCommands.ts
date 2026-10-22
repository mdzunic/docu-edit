// Dispatches the SlideFilmstrip's structural intents (SPEC-080) to the slide
// commands + layout gallery, extracted from PresentationView so the view stays
// a thin wiring hub (PresentationView.size.spec.ts). The filmstrip emits one
// `command` event carrying an action + index; this maps each action onto the
// already-undo-wrapped SlideCommands, selecting the target slide first for the
// layout actions so the gallery opens on it.
import type { SlideCommands } from "./useSlideCommands";

export interface FilmstripCommandsDeps {
  commands: SlideCommands;
  /** Make `index` the current slide (the layout gallery opens on it). */
  selectSlide: (index: number) => void;
  /** Open the layout gallery for a brand-new slide after the selected one. */
  openLayoutForNew: () => void;
  /** Open the layout gallery to change the selected slide's layout. */
  openLayoutForChange: () => void;
}

export function useFilmstripCommands(deps: FilmstripCommandsDeps) {
  function onCommand(payload = { action: "", index: 0 }): void {
    const { action, index } = payload;
    if (action === "new") {
      deps.commands.add(index);
    } else if (action === "new-with-layout") {
      deps.selectSlide(index);
      deps.openLayoutForNew();
    } else if (action === "duplicate") {
      deps.commands.duplicate(index);
    } else if (action === "layout") {
      deps.selectSlide(index);
      deps.openLayoutForChange();
    } else if (action === "delete") {
      deps.commands.remove(index);
    } else if (action === "move-up") {
      deps.commands.moveUp(index);
    } else if (action === "move-down") {
      deps.commands.moveDown(index);
    } else if (action === "hide") {
      deps.commands.toggleHidden(index);
    }
  }

  function onReorder(payload = { from: 0, to: 0 }): void {
    deps.commands.reorder(payload.from, payload.to);
  }

  return { onCommand, onReorder };
}
