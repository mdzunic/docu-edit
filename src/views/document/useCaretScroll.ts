// Keeping the model caret on screen (SPEC-011/025): two scroll behaviours the
// editor backbone hands out through EditorContext. Both read the live caret
// geometry off the rendered canvas and nudge the scroll container — no model
// state, so they live outside DocumentView as a small, testable unit.
//
// The scroll container and caret are plain (non-ref) editor state read live
// through getters, the same live-read contract EditorContext uses for them.
import { nextTick, type Ref } from "vue";
import { caretClientRect } from "../../lib/caretGeometry";
import type { CaretPosition } from "../../lib/textEditing";

export interface CaretScrollDeps {
  /** The contenteditable canvas root the caret geometry is measured against. */
  contentRoot: Ref<Element | null>;
  /** The live caret — read fresh at call time, never snapshotted. */
  getCaret: () => CaretPosition | null;
  /** The scrollable `<main>` ancestor, once mounted. */
  getScrollContainer: () => Element | null;
}

export function useCaretScroll(deps: CaretScrollDeps) {
  const { contentRoot, getCaret, getScrollContainer } = deps;

  /** Brings a model position into view and settles it a third of the way down
   * the canvas — a deliberate move to somewhere else in the document (a
   * bookmark jump, the page a break just started), where landing on the
   * viewport's very edge would show the destination without its context.
   * `always` skips the "only when it is off-screen" test, for callers that
   * know the caret changed page and want that page framed either way. */
  async function scrollModelCaretIntoView(position: CaretPosition | null = null, always = false): Promise<void> {
    await nextTick();
    const root = contentRoot.value;
    const scrollContainer = getScrollContainer();
    if (!root || !position || !scrollContainer) {
      return;
    }
    const rect = caretClientRect(root, position);
    if (!rect) {
      return;
    }
    const containerRect = scrollContainer.getBoundingClientRect();
    if (always || rect.top < containerRect.top || rect.bottom > containerRect.bottom) {
      scrollContainer.scrollTop += rect.top - containerRect.top - containerRect.height / 3;
    }
    // In the Print-layout page grid a caret can move to another column, so frame
    // it horizontally too, mirroring the vertical path above (SPEC-063 AC-37).
    if (always || rect.left < containerRect.left || rect.right > containerRect.right) {
      scrollContainer.scrollLeft += rect.left - containerRect.left - containerRect.width / 3;
    }
  }

  /** Keeps the caret on screen after an edit, moving the view as little as
   * possible. An edit must never leave the user typing somewhere they cannot
   * see — filling a page pushes the caret past the fold, and until this ran
   * the canvas simply stayed put — but scrolling further than the line just
   * written would lose their place in its own way, so this nudges rather than
   * reframes (contrast scrollModelCaretIntoView above). */
  async function keepCaretVisible(): Promise<void> {
    await nextTick();
    const root = contentRoot.value;
    const caret = getCaret();
    const scrollContainer = getScrollContainer();
    if (!root || !caret || !scrollContainer) {
      return;
    }
    const rect = caretClientRect(root, caret);
    if (!rect) {
      return;
    }
    const view = scrollContainer.getBoundingClientRect();
    const pad = 24;
    if (rect.top < view.top + pad) {
      scrollContainer.scrollTop -= view.top + pad - rect.top;
    } else if (rect.bottom > view.bottom - pad) {
      scrollContainer.scrollTop += rect.bottom - (view.bottom - pad);
    }
  }

  return { scrollModelCaretIntoView, keepCaretVisible };
}
