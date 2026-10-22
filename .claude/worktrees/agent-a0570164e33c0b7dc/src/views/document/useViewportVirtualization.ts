// --- page virtualization (SPEC-028 AC-49/50) --------------------------------
// Mounting every block of a 200-page document is what used to blow the
// settle budget: each mid-pass relayout shifted page boundaries, moving
// most paragraphs to a different DocPage, and Vue unmounted + remounted
// nearly the whole canvas several times per pass. Beyond a small page
// count only pages that matter render their content — everything near the
// viewport, the caret's / selection anchor's / current find match's page,
// and pages already realized whose layout hasn't changed (sticky, so DOM
// nodes on earlier pages survive edits elsewhere — AC-52/53) — while the
// rest render as correctly-sized empty page boxes. Their blocks are
// measured off-DOM by the pagination controller's canvas metrics, which is
// why virtualization requires `canMeasureOffscreen`.
import { computed, nextTick, ref, watch } from "vue";
import type { Ref } from "vue";
import { PX_PER_CM } from "../../lib/core/units";
import type { PageLayout } from "../../lib/pagination";
import type { PaginationController } from "../../composables/usePagination";
import type { EditorContext } from "./context";

export interface ViewportVirtualizationDeps extends EditorContext {
  pagination: PaginationController;
  paginationPages: Ref<PageLayout[]>;
  /** usePrintPreview's realize-everything gate (its `forceRealizeAll`):
   * print preview, printing and PDF export pause virtualization. */
  printForceRealize: Ref<boolean>;
  /** Reactive read of the view's `pageCount` computed — a getter over a
   * computed, like useOutlinePane's getCaretPosition. */
  getPageCount: () => number;
}

/** Documents up to this many pages always realize fully — the exact
 * pre-virtualization behaviour for ordinary documents. */
const VIRTUALIZE_MIN_PAGES = 12;
/** DocPage's inter-page margin (AC-2's visible gap), part of page-top math. */
const PAGE_GAP_PX = 12;
/** Extra realization margin around the viewport, px. */
const VIEWPORT_OVERSCAN_PX = 400;

export function useViewportVirtualization(deps: ViewportVirtualizationDeps) {
  const {
    pageRef,
    getCaret,
    getScrollContainer,
    onSessionClose,
    pagination,
    paginationPages,
    printForceRealize,
    getPageCount,
  } = deps;

  // Starts narrow: a large document's first render (opening lands before any
  // scroll/mount refresh) must realize only the top pages, not all of them —
  // refreshViewportRange corrects the range from the real viewport.
  const viewportPageRange = ref({ start: 0, end: 2 });
  /** Page-layout objects that have rendered their content once: they keep
   * rendering while reconciliation preserves them (an unchanged page never
   * loses its DOM), and drop out naturally when an edit replaces the page. */
  const realizedSticky = new WeakSet();
  /** The selection anchor's page (−1 for a collapsed caret): both ends of a
   * selection must stay rendered or the DOM selection would collapse. */
  const anchorPage = ref(-1);
  /** The caret's 1-based page (the status bar's "Page N of M"), always
   * kept realized so the caret's paragraph stays in the DOM. */
  const caretPage = ref(1);

  const virtualizing = computed(
    () =>
      // Print preview / printing / PDF export need every page's real content
      // in the DOM (SPEC-034) — virtualization pauses while they run.
      !printForceRealize.value &&
      pagination.canMeasureOffscreen && paginationPages.value.length > VIRTUALIZE_MIN_PAGES,
  );

  function refreshViewportRange() {
    const container = getScrollContainer();
    const pagesEl = pageRef.value;
    const layout = paginationPages.value;
    if (!container || !pagesEl || layout.length === 0) {
      return;
    }
    const pagesTop =
      pagesEl.getBoundingClientRect().top -
      container.getBoundingClientRect().top +
      container.scrollTop;
    const viewTop = container.scrollTop - pagesTop - VIEWPORT_OVERSCAN_PX - container.clientHeight;
    const viewBottom =
      container.scrollTop - pagesTop + 2 * container.clientHeight + VIEWPORT_OVERSCAN_PX;
    let top = 0;
    let start = layout.length - 1;
    let end = 0;
    for (let index = 0; index < layout.length; index += 1) {
      const bottom = top + layout[index].pageSize.heightCm * PX_PER_CM;
      if (bottom >= viewTop && top <= viewBottom) {
        if (index < start) {
          start = index;
        }
        if (index > end) {
          end = index;
        }
      }
      top = bottom + PAGE_GAP_PX;
    }
    if (start > end) {
      start = 0;
      end = Math.min(2, layout.length - 1);
    }
    const current = viewportPageRange.value;
    if (current.start !== start || current.end !== end) {
      viewportPageRange.value = { start, end };
    }
  }

  let viewportRefreshHandle = 0;
  function scheduleViewportRefresh() {
    if (viewportRefreshHandle !== 0 || typeof window.requestAnimationFrame !== "function") {
      return;
    }
    viewportRefreshHandle = window.requestAnimationFrame(() => {
      viewportRefreshHandle = 0;
      refreshViewportRange();
    });
  }
  watch(paginationPages, scheduleViewportRefresh);

  // Newly realized pages render blocks that may only have estimated metrics
  // (tables, images — canvas metrics cover paragraphs only): re-schedule so
  // the controller measures them off their fresh DOM.
  watch([viewportPageRange, caretPage], () => {
    if (!virtualizing.value) {
      return;
    }
    // Realization changed, the model didn't: measure the freshly rendered
    // blocks without forcing a synchronous relayout of an unchanged layout
    // (SPEC-047 AC-16 — see scheduleMeasure's own comment).
    void nextTick().then(() => pagination.scheduleMeasure());
  });

  function refreshCaretPage() {
    const caret = getCaret();
    if (caret) {
      caretPage.value = pagination.pageOfCaret(caret) + 1;
    } else if (caretPage.value > getPageCount()) {
      caretPage.value = getPageCount();
    }
  }
  watch(paginationPages, refreshCaretPage);

  // The next document opens scrolled to the top: realize only its head
  // until the viewport refresh sees the real scroll position.
  onSessionClose(() => {
    viewportPageRange.value = { start: 0, end: 2 };
    anchorPage.value = -1;
    caretPage.value = 1;
  });

  return {
    viewportPageRange,
    anchorPage,
    caretPage,
    realizedSticky,
    virtualizing,
    refreshViewportRange,
    scheduleViewportRefresh,
    refreshCaretPage,
  };
}
