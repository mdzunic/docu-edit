// SPEC-099 AC-4 — the shared slide multi-selection: click, Cmd/Ctrl+click,
// Shift+click, Cmd/Ctrl+A, Escape, arrows/Home/End with Shift extending, and
// collapsing onto a current slide moved from elsewhere.
import { describe, expect, it } from "vitest";
import { nextTick, ref } from "vue";
import { useSlideSelection } from "./useSlideSelection";

function harness(total = 6) {
  const currentIndex = ref(0);
  const count = ref(total);
  const selection = useSlideSelection({
    count: () => count.value,
    currentIndex,
    select: (index) => {
      currentIndex.value = index;
    },
  });
  return { currentIndex, count, selection };
}

function keydown(key: string, init: KeyboardEventInit = {}): KeyboardEvent {
  return new KeyboardEvent("keydown", { key, ...init });
}

describe("useSlideSelection", () => {
  it("starts as just the current slide", () => {
    const { selection } = harness();
    expect(selection.indexes.value).toEqual([0]);
  });

  it("a click selects one slide and makes it current", () => {
    const { selection, currentIndex } = harness();
    selection.click(3);
    expect(selection.indexes.value).toEqual([3]);
    expect(currentIndex.value).toBe(3);
  });

  it("Cmd/Ctrl+click toggles a slide in and out", () => {
    const { selection, currentIndex } = harness();
    selection.click(1);
    selection.click(4, { ctrlKey: true });
    expect(selection.indexes.value).toEqual([1, 4]);
    expect(currentIndex.value).toBe(4);
    selection.click(1, { metaKey: true });
    expect(selection.indexes.value).toEqual([4]);
    // The last selected slide cannot be toggled away.
    selection.click(4, { metaKey: true });
    expect(selection.indexes.value).toEqual([4]);
  });

  it("Shift+click extends from the anchor", () => {
    const { selection, currentIndex } = harness();
    selection.click(1);
    selection.click(3, { shiftKey: true });
    expect(selection.indexes.value).toEqual([1, 2, 3]);
    expect(currentIndex.value).toBe(3);
    selection.click(0, { shiftKey: true });
    expect(selection.indexes.value).toEqual([0, 1]);
  });

  it("Cmd/Ctrl+A selects all and Escape collapses back to the current slide", () => {
    const { selection } = harness();
    selection.click(2);
    expect(selection.onKeydown(keydown("a", { ctrlKey: true }))).toBe(true);
    expect(selection.indexes.value).toEqual([0, 1, 2, 3, 4, 5]);
    expect(selection.onKeydown(keydown("Escape"))).toBe(true);
    expect(selection.indexes.value).toEqual([2]);
    // A single selection leaves Escape to others.
    expect(selection.onKeydown(keydown("Escape"))).toBe(false);
  });

  it("arrows/Home/End move the roving selection; Shift+arrow extends", () => {
    const { selection, currentIndex } = harness();
    selection.click(2);
    selection.onKeydown(keydown("ArrowRight"));
    expect(selection.indexes.value).toEqual([3]);
    selection.onKeydown(keydown("ArrowDown", { shiftKey: true }));
    expect(selection.indexes.value).toEqual([3, 4]);
    selection.onKeydown(keydown("ArrowUp", { shiftKey: true }));
    selection.onKeydown(keydown("ArrowUp", { shiftKey: true }));
    expect(selection.indexes.value).toEqual([2, 3]);
    selection.onKeydown(keydown("End"));
    expect(selection.indexes.value).toEqual([5]);
    selection.onKeydown(keydown("Home", { shiftKey: true }));
    expect(selection.indexes.value).toEqual([0, 1, 2, 3, 4, 5]);
    expect(currentIndex.value).toBe(0);
    expect(selection.onKeydown(keydown("x"))).toBe(false);
  });

  it("collapses onto a current slide moved from elsewhere", async () => {
    const { selection, currentIndex } = harness();
    selection.click(1);
    selection.click(2, { shiftKey: true });
    currentIndex.value = 5;
    await nextTick();
    expect(selection.indexes.value).toEqual([5]);
  });

  it("selectRange picks a command's result; indexes past the end fall away", () => {
    const { selection, count, currentIndex } = harness();
    selection.selectRange([3, 4]);
    expect(selection.indexes.value).toEqual([3, 4]);
    expect(currentIndex.value).toBe(3);
    count.value = 4;
    expect(selection.indexes.value).toEqual([3]);
  });
});
