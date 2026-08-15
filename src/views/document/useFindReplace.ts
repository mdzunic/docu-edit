// Find and replace (SPEC-017). The pure logic (compiling a query, locating
// matches, rewriting the model) lives in lib/findReplace.ts; this composable
// owns the DOM-facing half: the find bar's reactive state, the on-screen
// highlight overlay, scrolling the current match into view, and applying
// replaces through the session's applyEdit path so undo/autosave treat them
// identically to every other edit.
//
// The performance contract (AC-30, AC-31 — first results within 300 ms and no
// typing block on a 200-page document) shapes the design: match *counting* is
// a cheap string scan, but *measuring* a DOM rectangle for every match is not,
// so the overlay only measures matches whose paragraph is on screen (plus the
// current one), and a document edit while the bar is open reschedules that
// work on a debounce/rAF instead of running it synchronously in the keystroke
// path.
import { computed, ref, shallowRef, watch } from "vue";
import type { Ref } from "vue";
import {
  compileQuery,
  findMatches,
  indexAtOrAfter,
  isQueryError,
  matchSelection,
  noHighlightRects,
  noMatches,
  replaceMatch,
  replaceMatches,
} from "../../lib/findReplace";
import type { FindMatch, HighlightRect } from "../../lib/findReplace";
import { noIdleHandle, scheduleIdleWork } from "../../lib/core/idleScheduler";
import {
  BLOCK_SELECTOR,
  domPointForCaret,
  setDomCaret,
} from "../../lib/domCaret";
import { caretSelection, documentStart, noSelection } from "../../lib/commands";
import type { ModelSelection } from "../../lib/commands";
import { noCaret } from "../../lib/textEditing";
import type { CaretPosition } from "../../lib/textEditing";
import type { PageLayout } from "../../lib/pagination";
import type { EditorContext } from "./context";

export interface FindReplaceDeps extends EditorContext {
  /** Whether Find itself may open/search (SPEC-047 AC-27: find still works
   * over a too-large read-only document). `editable` stays the gate for the
   * mutating actions below it (Replace/Replace all) — this is strictly
   * broader than `editable`, never narrower. */
  canSearch: Readonly<Ref<boolean>>;
  paginationPages: Ref<PageLayout[]>;
  pageOfCaret: (caret: CaretPosition) => number;
  updateCaretStatus: (range?: { start: CaretPosition; end: CaretPosition }) => void;
  setCaret: (caret: CaretPosition | null) => void;
  /** Called when the engine wants the search re-run after an edit. */
  registerRefresh: (fn: () => void) => void;
}

export function useFindReplace(deps: FindReplaceDeps) {
  const {
    doc,
    contentRoot,
    pageRef,
    editable,
    canSearch,
    paginationPages,
    pageOfCaret,
    applyEdit,
    updateCaretStatus,
    addToast,
    getCaret,
    setCaret,
    getScrollContainer,
    registerRefresh,
    onSessionClose,
    onUnmount,
    onScroll,
  } = deps;

  const findVisible = ref(false);
  const findExpanded = ref(false);
  const findQuery = ref("");
  const findReplaceText = ref("");
  const findMatchCase = ref(false);
  const findWholeWords = ref(false);
  const findUseRegex = ref(false);
  const findError = ref("");
  const findHint = ref("");
  const findCurrentIndex = ref(-1);
  // Shallow refs: both are wholesale-replaced, never mutated in place, and can
  // grow large (thousands of matches) — a deep proxy per entry would waste.
  const findMatchesRef = shallowRef<FindMatch[]>(noMatches());
  const findHighlightRects = shallowRef<HighlightRect[]>(noHighlightRects());

  // The caret to restore when Escape closes the bar (AC-13), a per-block index
  // of matches for viewport-limited measuring, and the scheduling handles that
  // keep highlight work off the keystroke path. Plain variables — nothing
  // renders from them.
  let caretBeforeFind = noCaret();
  let findMatchesByBlock = new Map<string, FindMatch[]>();
  let findIdleHandle = noIdleHandle();
  let findDebounceTimer = 0;
  let findHintTimer = 0;

  /** The current find match's page: realized so the highlight overlay and
   * scroll-into-view can measure it (SPEC-017). */
  const findMatchPage = computed(() => {
    void paginationPages.value;
    const match = findMatchesRef.value[findCurrentIndex.value];
    if (!match) {
      return -1;
    }
    return pageOfCaret({ section: match.section, block: match.block, offset: match.start });
  });

  function currentFindAnchor(): CaretPosition | null {
    const matches = findMatchesRef.value;
    const index = findCurrentIndex.value;
    if (index >= 0 && index < matches.length) {
      return matchSelection(matches[index]).anchor;
    }
    return getCaret();
  }

  function currentFindSelection(): ModelSelection | null {
    const matches = findMatchesRef.value;
    const index = findCurrentIndex.value;
    if (index >= 0 && index < matches.length) {
      return matchSelection(matches[index]);
    }
    const caret = getCaret();
    return caret ? caretSelection(caret) : noSelection();
  }

  function currentFindOptions() {
    return {
      matchCase: findMatchCase.value,
      wholeWords: findWholeWords.value,
      regex: findUseRegex.value,
    };
  }

  function openFind(expand = false) {
    if (!doc.value || !canSearch.value) {
      return;
    }
    if (!findVisible.value) {
      // Snapshot the caret only on a fresh open, so Escape can put it back
      // exactly where it was (AC-13) even after navigation moved the selection.
      caretBeforeFind = getCaret();
    }
    findVisible.value = true;
    if (expand) {
      findExpanded.value = true;
    }
    runSearch();
  }

  /** Escape / the close button (AC-4, AC-11, AC-12): hide the bar, clear the
   * highlights, and — for Escape — return the caret to where it was (AC-13). */
  function closeFind(restoreCaret = false) {
    if (!findVisible.value) {
      return;
    }
    findVisible.value = false;
    findExpanded.value = false;
    findError.value = "";
    findHint.value = "";
    setFindMatches(noMatches(), -1);
    findHighlightRects.value = noHighlightRects();
    cancelFindScheduled();
    const root = contentRoot.value as HTMLElement | null;
    if (root && canSearch.value) {
      root.focus();
      if (restoreCaret && caretBeforeFind) {
        setCaret(caretBeforeFind);
        setDomCaret(root, caretBeforeFind);
        updateCaretStatus({ start: caretBeforeFind, end: caretBeforeFind });
      }
    }
  }

  function toggleFindExpand() {
    findExpanded.value = !findExpanded.value;
  }

  /** Tears the bar down when the document session ends (navigation/unmount).
   * The three search options are deliberately kept — they persist for the
   * session (AC-23), across documents too. */
  function resetFindState() {
    findVisible.value = false;
    findExpanded.value = false;
    findQuery.value = "";
    findReplaceText.value = "";
    findError.value = "";
    findHint.value = "";
    setFindMatches(noMatches(), -1);
    findHighlightRects.value = noHighlightRects();
    cancelFindScheduled();
  }

  function setFindMatches(matches: FindMatch[], index: number) {
    findMatchesRef.value = matches;
    findCurrentIndex.value = index;
    const byBlock = new Map<string, FindMatch[]>();
    for (const match of matches) {
      const key = `${match.section}:${match.block}`;
      const list = byBlock.get(key);
      if (list) {
        list.push(match);
      } else {
        byBlock.set(key, [match]);
      }
    }
    findMatchesByBlock = byBlock;
  }

  /** Recomputes matches for the current query/options (AC-5). Cheap enough to
   * run synchronously even on a large document (AC-30); only DOM measurement
   * is deferred. `preserveAnchor` keeps the current match near a model
   * position across a change. */
  function runSearch(preserveAnchor: CaretPosition | null = null) {
    const model = doc.value;
    if (!model) {
      return;
    }
    findHint.value = "";
    const compiled = compileQuery(findQuery.value, currentFindOptions());
    if (compiled === null) {
      findError.value = "";
      setFindMatches(noMatches(), -1);
      scheduleFindMeasure();
      return;
    }
    if (isQueryError(compiled)) {
      // AC-24, AC-25: an invalid regex shows an inline error and does not run.
      findError.value = compiled.error;
      setFindMatches(noMatches(), -1);
      scheduleFindMeasure();
      return;
    }
    findError.value = "";
    const matches = findMatches(model, compiled);
    let index = matches.length > 0 ? 0 : -1;
    if (preserveAnchor && matches.length > 0) {
      index = indexAtOrAfter(matches, preserveAnchor);
    }
    setFindMatches(matches, index);
    scheduleFindMeasure();
    if (index >= 0) {
      scheduleScrollCurrentIntoView();
    }
  }

  function findNext() {
    stepFind(1);
  }

  function findPrev() {
    stepFind(-1);
  }

  /** Enter / Shift+Enter and the prev/next buttons (AC-3, AC-8, AC-9),
   * wrapping at either end with a transient hint (AC-10). */
  function stepFind(direction = 1) {
    const matches = findMatchesRef.value;
    if (matches.length === 0) {
      return;
    }
    let index = findCurrentIndex.value;
    if (index < 0) {
      index = 0;
    } else {
      const moved = index + direction;
      if (moved < 0) {
        index = matches.length - 1;
        showFindHint("Wrapped to bottom");
      } else if (moved >= matches.length) {
        index = 0;
        showFindHint("Wrapped to top");
      } else {
        index = moved;
      }
    }
    findCurrentIndex.value = index;
    scheduleFindMeasure();
    scheduleScrollCurrentIntoView();
  }

  function showFindHint(text = "") {
    findHint.value = text;
    window.clearTimeout(findHintTimer);
    findHintTimer = window.setTimeout(() => {
      findHint.value = "";
    }, 1500);
  }

  /** The "Replace" button (AC-16): replaces the current match, taking its
   * first run's formatting (AC-26), then advances to the next match (AC-8). */
  function replaceCurrent() {
    const model = doc.value;
    const matches = findMatchesRef.value;
    const index = findCurrentIndex.value;
    if (!model || !editable.value || findError.value !== "" || index < 0 || index >= matches.length) {
      return;
    }
    const match = matches[index];
    const before = matchSelection(match);
    const anchor: CaretPosition = {
      section: match.section,
      block: match.block,
      offset: match.start + findReplaceText.value.length,
    };
    applyEdit(replaceMatch(model, match, findReplaceText.value), before, "replace");
    refreshFindAfterEdit(anchor);
  }

  /** The "Replace all" button (AC-17): one edit for every match (AC-19's
   * single undo step), reporting the count (AC-18). */
  function replaceAll() {
    const model = doc.value;
    if (!model || !editable.value || findQuery.value === "" || findError.value !== "") {
      return;
    }
    const compiled = compileQuery(findQuery.value, currentFindOptions());
    if (compiled === null || isQueryError(compiled)) {
      return;
    }
    const matches = findMatches(model, compiled);
    const count = matches.length;
    if (count > 0) {
      const before = currentFindSelection() ?? caretSelection(documentStart());
      applyEdit(replaceMatches(model, matches, findReplaceText.value), before, "replace all");
    }
    addToast(`Replaced ${count} occurrences`);
    refreshFindAfterEdit(noCaret());
  }

  /** Recomputes matches after a model edit (a replace, or typing in the canvas
   * with the bar open), keeping the current match near `anchor`. */
  function refreshFindAfterEdit(anchor: CaretPosition | null = null) {
    const model = doc.value;
    if (!model || !findVisible.value) {
      return;
    }
    const compiled = compileQuery(findQuery.value, currentFindOptions());
    if (compiled === null || isQueryError(compiled)) {
      setFindMatches(noMatches(), -1);
      scheduleFindMeasure();
      return;
    }
    const matches = findMatches(model, compiled);
    const focus = anchor ?? currentFindAnchor();
    const index = focus
      ? indexAtOrAfter(matches, focus)
      : matches.length > 0
        ? 0
        : -1;
    setFindMatches(matches, index);
    scheduleFindMeasure();
    if (index >= 0) {
      scheduleScrollCurrentIntoView();
    }
  }

  /** Called from every model-edit path while the bar is open. Debounced so a
   * burst of keystrokes reschedules a single recompute rather than rescanning
   * and re-measuring on every character (AC-31). */
  function scheduleFindRefresh() {
    if (!findVisible.value) {
      return;
    }
    window.clearTimeout(findDebounceTimer);
    findDebounceTimer = window.setTimeout(() => {
      refreshFindAfterEdit(currentFindAnchor());
    }, 120);
  }

  function scheduleFindMeasure() {
    if (findIdleHandle) {
      return;
    }
    // Highlight measurement forces a full reflow — the most expensive read in
    // the find path, and the match COUNT never depends on it. Idle scheduling
    // keeps it out of the keystroke-to-first-result window (AC-30/31).
    findIdleHandle = scheduleIdleWork(
      () => {
        findIdleHandle = null;
        measureFindHighlights();
      },
      { timeoutMs: 150, fallbackBudgetMs: 8, maxBudgetMs: 16 },
    );
  }

  function scheduleScrollCurrentIntoView() {
    window.requestAnimationFrame(scrollCurrentMatchIntoView);
  }

  function cancelFindScheduled() {
    if (findIdleHandle) {
      findIdleHandle.cancel();
      findIdleHandle = null;
    }
    window.clearTimeout(findDebounceTimer);
    window.clearTimeout(findHintTimer);
  }

  function findViewportBounds() {
    const scrollContainer = getScrollContainer();
    if (scrollContainer) {
      const rect = scrollContainer.getBoundingClientRect();
      return { top: rect.top, bottom: rect.bottom };
    }
    return { top: 0, bottom: window.innerHeight };
  }

  /** The set of "section:block" keys for paragraphs currently on (or just off)
   * screen — the only ones whose matches are worth measuring. */
  function visibleFindBlockKeys(root: HTMLElement, bounds: { top: number; bottom: number }) {
    const keys = new Set<string>();
    const margin = 200;
    const pageElements = root.querySelectorAll('[data-testid="document-page"]');
    const scopes: HTMLElement[] = [];
    if (pageElements.length > 0) {
      for (const pageElement of Array.from(pageElements)) {
        const rect = (pageElement as HTMLElement).getBoundingClientRect();
        if (rect.bottom < bounds.top - margin || rect.top > bounds.bottom + margin) {
          continue;
        }
        scopes.push(pageElement as HTMLElement);
      }
    } else {
      scopes.push(root);
    }
    for (const scope of scopes) {
      for (const element of Array.from(scope.querySelectorAll(BLOCK_SELECTOR))) {
        const rect = (element as HTMLElement).getBoundingClientRect();
        if (rect.bottom < bounds.top - margin || rect.top > bounds.bottom + margin) {
          continue;
        }
        keys.add(`${element.getAttribute("data-section")}:${element.getAttribute("data-block")}`);
      }
    }
    return keys;
  }

  function matchClientRects(root: HTMLElement, match: FindMatch): DOMRect[] {
    const startPoint = domPointForCaret(root, {
      section: match.section,
      block: match.block,
      offset: match.start,
    });
    const endPoint = domPointForCaret(root, {
      section: match.section,
      block: match.block,
      offset: match.end,
    });
    if (!startPoint || !endPoint) {
      return [];
    }
    const range = root.ownerDocument.createRange();
    try {
      range.setStart(startPoint.node, startPoint.offset);
      range.setEnd(endPoint.node, endPoint.offset);
    } catch {
      return [];
    }
    return Array.from(range.getClientRects());
  }

  /** Rebuilds the overlay rectangles, measuring only on-screen matches (plus
   * the current one). Runs off the keystroke path via rAF. */
  function measureFindHighlights() {
    const root = contentRoot.value as HTMLElement | null;
    const pageEl = pageRef.value;
    const matches = findMatchesRef.value;
    if (!findVisible.value || !root || !pageEl || matches.length === 0) {
      findHighlightRects.value = noHighlightRects();
      return;
    }
    const pageRect = pageEl.getBoundingClientRect();
    const visibleKeys = visibleFindBlockKeys(root, findViewportBounds());
    const currentIndex = findCurrentIndex.value;
    const rects: HighlightRect[] = noHighlightRects();
    for (const [key, blockMatches] of findMatchesByBlock) {
      if (!visibleKeys.has(key)) {
        continue;
      }
      for (const match of blockMatches) {
        pushMatchRects(rects, root, pageRect, match, false, currentIndex, matches);
      }
    }
    // The current match is always drawn distinctly, even if it fell outside
    // the measured band before its scroll-into-view lands (AC-6).
    if (currentIndex >= 0 && currentIndex < matches.length) {
      const current = matches[currentIndex];
      if (!visibleKeys.has(`${current.section}:${current.block}`)) {
        pushMatchRects(rects, root, pageRect, current, true, currentIndex, matches);
      }
    }
    findHighlightRects.value = rects;
  }

  function pushMatchRects(
    rects: HighlightRect[],
    root: HTMLElement,
    pageRect: DOMRect,
    match: FindMatch,
    forceCurrent = false,
    currentIndex = -1,
    matches: FindMatch[],
  ) {
    const isCurrent = forceCurrent || matches[currentIndex] === match;
    for (const rect of matchClientRects(root, match)) {
      rects.push({
        left: rect.left - pageRect.left,
        top: rect.top - pageRect.top,
        width: rect.width,
        height: rect.height,
        current: isCurrent,
      });
    }
  }

  /** Scrolls the current match into view (AC-7). */
  function scrollCurrentMatchIntoView() {
    const root = contentRoot.value as HTMLElement | null;
    const matches = findMatchesRef.value;
    const index = findCurrentIndex.value;
    const scrollContainer = getScrollContainer();
    if (!root || !scrollContainer || index < 0 || index >= matches.length) {
      return;
    }
    const rects = matchClientRects(root, matches[index]);
    if (rects.length === 0) {
      return;
    }
    const rect = rects[0];
    const view = scrollContainer.getBoundingClientRect();
    const pad = 60;
    if (rect.top < view.top + pad) {
      scrollContainer.scrollTop -= view.top + pad - rect.top;
    } else if (rect.bottom > view.bottom - pad) {
      scrollContainer.scrollTop += rect.bottom - (view.bottom - pad);
    }
  }

  // Re-run the search whenever the query or an option changes while the bar is
  // open (AC-5, AC-20..23).
  watch([findQuery, findMatchCase, findWholeWords, findUseRegex], () => {
    if (findVisible.value) {
      runSearch(currentFindAnchor());
    }
  });

  registerRefresh(scheduleFindRefresh);
  onSessionClose(resetFindState);
  onUnmount(cancelFindScheduled);
  onScroll(() => {
    if (findVisible.value) {
      scheduleFindMeasure();
    }
  });

  return {
    findVisible,
    findExpanded,
    findQuery,
    findReplaceText,
    findMatchCase,
    findWholeWords,
    findUseRegex,
    findError,
    findHint,
    findCurrentIndex,
    findMatchesRef,
    findHighlightRects,
    findMatchPage,
    openFind,
    closeFind,
    toggleFindExpand,
    findNext,
    findPrev,
    replaceCurrent,
    replaceAll,
    scheduleFindMeasure,
  };
}
