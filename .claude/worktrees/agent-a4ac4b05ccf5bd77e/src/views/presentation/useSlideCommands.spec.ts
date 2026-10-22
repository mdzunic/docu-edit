// SPEC-104 — the filmstrip's Hide slide / Show slide command: one undo entry
// recorded BEFORE the mutation (SPEC-080's record → mutate → commit order),
// the slide marked edited for its format, and the deck persisted through
// commit. The other slide commands are covered by the view specs.
import { describe, expect, it } from "vitest";
import { ref, shallowRef } from "vue";
import { createBlankPresentation, addSlide } from "../../lib/presentation/slides";
import type { PresentationFormat, PresentationModel } from "../../lib/presentation/model";
import { usePresentationUndo } from "./usePresentationUndo";
import { useSlideCommands } from "./useSlideCommands";
import { useFilmstripCommands } from "./useFilmstripCommands";

function harness(format: PresentationFormat = "pptx") {
  const deck = createBlankPresentation();
  deck.format = format;
  addSlide(deck, 0);
  addSlide(deck, 1);
  for (const slide of deck.slides) {
    slide.sourceXml = "<draw:page/>";
  }
  const model = shallowRef<PresentationModel | null>(deck);
  const currentIndex = ref(0);
  const order: string[] = [];
  let commits = 0;
  const undo = usePresentationUndo({ model, onApply: () => (commits += 1) });
  const commands = useSlideCommands({
    model,
    currentIndex,
    record: (label) => {
      order.push(`record ${label} hidden=${model.value?.slides[1].hidden === true}`);
      undo.record(label);
    },
    commit: () => {
      order.push(`commit hidden=${model.value?.slides[1].hidden === true}`);
      commits += 1;
    },
  });
  return { model, currentIndex, order, undo, commands, commits: () => commits };
}

describe("useSlideCommands.toggleHidden (SPEC-104)", () => {
  it("records \"Hide slide\" before mutating, then commits; the slide is marked edited", () => {
    const { model, order, commands, commits } = harness();
    commands.toggleHidden(1);
    expect(order).toEqual(["record Hide slide hidden=false", "commit hidden=true"]);
    expect(model.value?.slides[1].hidden).toBe(true);
    expect(model.value?.slides[1].dirty).toBe(true);
    expect(model.value?.structureDirty).toBe(true); // addSlide's, untouched by the toggle
    expect(commits()).toBe(1);
  });

  it("records \"Show slide\" on a hidden slide and deletes the flag", () => {
    const { model, order, commands } = harness();
    model.value!.slides[1].hidden = true;
    commands.toggleHidden(1);
    expect(order[0]).toBe("record Show slide hidden=true");
    expect("hidden" in model.value!.slides[1]).toBe(false);
  });

  it("is one undo step: undo restores the previous state, redo re-applies it", () => {
    const { model, undo, commands } = harness();
    commands.toggleHidden(1);
    expect(undo.canUndo.value).toBe(true);
    undo.undo();
    expect(model.value?.slides[1].hidden).toBeUndefined();
    undo.redo();
    expect(model.value?.slides[1].hidden).toBe(true);
  });

  it("odp: marking edited drops the page's verbatim source so it rebuilds", () => {
    const { model, commands } = harness("odp");
    commands.toggleHidden(2);
    expect(model.value?.slides[2].sourceXml).toBeUndefined();
    expect(model.value?.slides[2].dirty).toBe(true);
  });

  it("keeps the toggled slide current", () => {
    const { currentIndex, commands } = harness();
    commands.toggleHidden(2);
    expect(currentIndex.value).toBe(2);
  });

  it("the filmstrip's hide intent routes to it", () => {
    const { model, commands } = harness();
    const filmstrip = useFilmstripCommands({
      commands,
      selectSlide: () => undefined,
      openLayoutForNew: () => undefined,
      openLayoutForChange: () => undefined,
    });
    filmstrip.onCommand({ action: "hide", index: 0 });
    expect(model.value?.slides[0].hidden).toBe(true);
  });
});
