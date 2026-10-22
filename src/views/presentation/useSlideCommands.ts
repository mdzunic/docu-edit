// Wires the filmstrip's structural intents (SPEC-080) to the framework-free
// slide commands: record the pre-command state on the undo stack, run the
// command, select the slide it returns, then re-render + persist through
// `commit`. Lives beside usePresentationUndo in the presentation composable
// directory so PresentationView stays a thin wiring hub.
import type { Ref } from "vue";
import type { PresentationModel } from "../../lib/presentation/model";
import { addSlide, deleteSlide, duplicateSlide, moveSlide, setSlideHidden } from "../../lib/presentation/slides";

export interface SlideCommandsDeps {
  model: Ref<PresentationModel | null>;
  currentIndex: Ref<number>;
  /** Snapshot the pre-command state for undo (usePresentationUndo.record). */
  record: (label?: string) => void;
  /** Re-render the shallowRef and write the deck through (touchDocument). */
  commit: () => void;
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

  return {
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
  };
}
