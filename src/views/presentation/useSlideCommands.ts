// Wires the filmstrip's structural intents (SPEC-080) to the framework-free
// slide commands: record the pre-command state on the undo stack, run the
// command, select the slide it returns, then re-render + report through
// `commit`. Lives beside usePresentationUndo in the presentation composable
// directory so PresentationView stays a thin wiring hub. Since SPEC-103 it
// also owns the filmstrip's `command`/`reorder` dispatch (onFilmstripCommand/
// onFilmstripReorder), which had lived in its own useFilmstripCommands.
import type { Ref } from "vue";
import type { PresentationModel } from "../../lib/presentation/model";
import { addSlide, deleteSlide, duplicateSlide, moveSlide, setSlideHidden } from "../../lib/presentation/slides";

export interface SlideCommandsDeps {
  model: Ref<PresentationModel | null>;
  currentIndex: Ref<number>;
  /** Snapshot the pre-command state for undo (usePresentationUndo.record). */
  record: (label?: string) => void;
  /** Re-render the shallowRef and report the edit (useSlideRender.commit). */
  commit: () => void;
  /** Make `index` the current slide (the layout gallery opens on it). */
  selectSlide?: (index: number) => void;
  /** The layout gallery's two entry points — called lazily, so the gallery
   * (which itself adds slides through these commands) can be built after. */
  openLayoutForNew?: () => void;
  openLayoutForChange?: () => void;
}

export interface SlideCommands {
  /** Inserts a new slide after `afterIndex`; with `layoutKey` the slide is born
   * from that catalogue layout's slots (SPEC-089), else SPEC-080's title+body. */
  add: (afterIndex?: number, layoutKey?: string) => void;
  duplicate: (index?: number) => void;
  remove: (index?: number) => void;
  moveUp: (index?: number) => void;
  moveDown: (index?: number) => void;
  reorder: (from: number, to: number) => void;
  /** Hides a shown slide or shows a hidden one (SPEC-104) — one undo entry,
   * "Hide slide" / "Show slide". The slide stays selected. */
  toggleHidden: (index?: number) => void;
  /** The filmstrip's (and the sorter card menu's) one `command` event — an
   * action plus a slide index — mapped onto the commands above, selecting the
   * slide first for the layout actions so the gallery opens on it. */
  onFilmstripCommand: (payload?: { action: string; index: number }) => void;
  /** The filmstrip's drag-reorder. */
  onFilmstripReorder: (payload?: { from: number; to: number }) => void;
}

export function useSlideCommands(deps: SlideCommandsDeps): SlideCommands {
  function run(label: string, apply: (model: PresentationModel) => number): void {
    const model = deps.model.value;
    if (!model) {
      return;
    }
    deps.record(label);
    deps.currentIndex.value = apply(model);
    deps.commit();
  }

  const commands: SlideCommands = {
    add: (afterIndex = deps.currentIndex.value, layoutKey?: string) =>
      run("New slide", (model) => addSlide(model, afterIndex, layoutKey)),
    duplicate: (index = deps.currentIndex.value) => run("Duplicate slide", (model) => duplicateSlide(model, index)),
    remove: (index = deps.currentIndex.value) => run("Delete slide", (model) => deleteSlide(model, index)),
    moveUp: (index = deps.currentIndex.value) => run("Move slide up", (model) => moveSlide(model, index, index - 1)),
    moveDown: (index = deps.currentIndex.value) => run("Move slide down", (model) => moveSlide(model, index, index + 1)),
    reorder: (from, to) => run("Move slide", (model) => moveSlide(model, from, to)),
    toggleHidden: (index = deps.currentIndex.value) => {
      const slide = deps.model.value?.slides[index];
      if (!slide) {
        return;
      }
      const hidden = slide.hidden === true;
      run(hidden ? "Show slide" : "Hide slide", (model) => setSlideHidden(model, index, !hidden));
    },
    onFilmstripCommand: (payload = { action: "", index: 0 }) => {
      const { action, index } = payload;
      if (action === "new") {
        commands.add(index);
      } else if (action === "new-with-layout") {
        deps.selectSlide?.(index);
        deps.openLayoutForNew?.();
      } else if (action === "duplicate") {
        commands.duplicate(index);
      } else if (action === "layout") {
        deps.selectSlide?.(index);
        deps.openLayoutForChange?.();
      } else if (action === "delete") {
        commands.remove(index);
      } else if (action === "move-up") {
        commands.moveUp(index);
      } else if (action === "move-down") {
        commands.moveDown(index);
      } else if (action === "hide") {
        commands.toggleHidden(index);
      }
    },
    onFilmstripReorder: (payload = { from: 0, to: 0 }) => commands.reorder(payload.from, payload.to),
  };
  return commands;
}
