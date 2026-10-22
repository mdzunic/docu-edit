import { effectScope, ref } from "vue";
import { afterEach, describe, expect, it, vi } from "vitest";
import { useCaretScroll } from "./useCaretScroll";
import type { CaretPosition } from "../../lib/textEditing";

// caretClientRect needs real layout jsdom does not have, so the caret's
// on-screen rect is the one seam these tests stub.
const caretRect = { top: 0, bottom: 0, left: 0, right: 0 } as { top: number; bottom: number; left: number; right: number };
vi.mock("../../lib/caretGeometry", () => ({
  caretClientRect: () => caretRect,
}));

function makeScrollContainer(rect: { top: number; bottom: number; left: number; right: number }) {
  const el = {
    scrollTop: 0,
    scrollLeft: 0,
    getBoundingClientRect: () =>
      ({ ...rect, width: rect.right - rect.left, height: rect.bottom - rect.top }) as DOMRect,
  };
  return el as unknown as Element & { scrollTop: number; scrollLeft: number };
}

function mount(container: Element | null) {
  const contentRoot = ref<Element | null>(document.createElement("div"));
  const caret: CaretPosition = { section: 0, block: 0, offset: 0 };
  const scope = effectScope();
  let view!: ReturnType<typeof useCaretScroll>;
  scope.run(() => {
    view = useCaretScroll({
      contentRoot,
      getCaret: () => caret,
      getScrollContainer: () => container,
    });
  });
  return { view, stop: () => scope.stop() };
}

describe("useCaretScroll", () => {
  afterEach(() => {
    caretRect.top = 0;
    caretRect.bottom = 0;
    caretRect.left = 0;
    caretRect.right = 0;
  });

  it("scrolls horizontally to frame a caret in another grid column (AC-37)", async () => {
    // A viewport 0..1000 wide/tall; the caret sits far to the right (another
    // grid column), off the right edge.
    const container = makeScrollContainer({ top: 0, bottom: 1000, left: 0, right: 1000 });
    caretRect.top = 100;
    caretRect.bottom = 120;
    caretRect.left = 1600;
    caretRect.right = 1620;
    const h = mount(container);
    await h.view.scrollModelCaretIntoView({ section: 0, block: 5, offset: 0 });
    // Framed a third of the way in: 1600 - 0 - 1000/3.
    expect(container.scrollLeft).toBeCloseTo(1600 - 1000 / 3, 5);
    // Vertically already on screen (100..120 within 0..1000), so no v-scroll.
    expect(container.scrollTop).toBe(0);
    h.stop();
  });

  it("leaves horizontal scroll alone when the caret is already in view", async () => {
    const container = makeScrollContainer({ top: 0, bottom: 1000, left: 0, right: 1000 });
    caretRect.top = 100;
    caretRect.bottom = 120;
    caretRect.left = 400;
    caretRect.right = 420;
    const h = mount(container);
    await h.view.scrollModelCaretIntoView({ section: 0, block: 0, offset: 0 });
    expect(container.scrollLeft).toBe(0);
    expect(container.scrollTop).toBe(0);
    h.stop();
  });
});
