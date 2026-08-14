// Live word count and the statistics dialog (SPEC-036). The document-wide
// total and the statistics dialog both compute via a worker (SPEC-047
// AC-12: wordCountClient.ts) rather than the per-block WeakMap cache this
// composable used to own directly — a worker round trip loses cross-call
// object identity anyway (postMessage's structured clone mints fresh
// objects every call), so there is nothing left for a local cache to hit;
// the worker recomputes the walk fresh each time instead. The document-wide
// total recomputes on a throttle (AC-2: never delays input); the
// selection-aware count recomputes synchronously on every selection change
// (DocumentView's updateCaretStatus calls setSelection alongside its other
// per-selection state) — bounded by selection size, not document size, so
// it stays on the main thread rather than paying a worker round trip for
// something already cheap.
import { computed, ref, watch } from "vue";
import type { Ref } from "vue";
import { isCollapsed, type ModelSelection } from "../../lib/commands";
import type { CaretPosition } from "../../lib/textEditing";
import type { DocumentModel } from "../../lib/documentModel";
import { emptyCountStats, selectionWordCount, type CountStats } from "../../lib/wordCount";
import { computeDocumentStats } from "../../lib/workers/wordCountClient";

/** Coalesces bursts of recomputation into at most one run per window, with
 * both a leading run (the first change after a quiet period lands right
 * away) and a trailing run (the latest state after a burst is never
 * dropped) — the standard throttle shape, picked over a debounce because
 * AC-1 asks for a count that keeps updating *while* the user types, not
 * one that waits for them to stop. */
function createThrottle(run: () => void, windowMs: number) {
  let lastRun = 0;
  let timer: ReturnType<typeof setTimeout> | null = null;
  function schedule() {
    if (timer !== null) {
      return;
    }
    const elapsed = Date.now() - lastRun;
    const delay = Math.max(0, windowMs - elapsed);
    timer = setTimeout(() => {
      timer = null;
      lastRun = Date.now();
      run();
    }, delay);
  }
  function cancel() {
    if (timer !== null) {
      clearTimeout(timer);
      timer = null;
    }
  }
  return { schedule, cancel };
}

const THROTTLE_MS = 300;

export interface WordCountDeps {
  doc: Ref<DocumentModel | null>;
  onUnmount: (cb: () => void) => void;
}

export function useWordCount(deps: WordCountDeps) {
  const { doc, onUnmount } = deps;

  const liveWordCount = ref(0);
  /** The visible/hidden status text's last announced word count — a
   * screen-reader-facing region only re-announces when its text actually
   * changes, so binding it to a value that only updates on a real change
   * (rather than every throttle tick) is what makes the announcement
   * "material" (AC-28) without any separate threshold bookkeeping. */
  const announcement = ref("0 words");

  // Guards against an in-flight worker reply landing after a later request
  // has already superseded it (a burst of edits can have more than one
  // count in flight at once) — only the most recently issued request's
  // reply is ever applied.
  let liveRequestToken = 0;

  async function recomputeLiveCount() {
    const model = doc.value;
    const token = ++liveRequestToken;
    const stats = model ? await computeDocumentStats(model, { includeNotes: false }) : emptyCountStats();
    if (token !== liveRequestToken) {
      return;
    }
    const words = stats.words;
    if (words === liveWordCount.value) {
      return;
    }
    liveWordCount.value = words;
    announcement.value = `${words} ${words === 1 ? "word" : "words"}`;
  }

  const throttle = createThrottle(() => {
    void recomputeLiveCount();
  }, THROTTLE_MS);

  // The first count (opening/creating a document) is visible immediately
  // rather than waiting out the throttle window; every later change goes
  // through the throttle.
  void recomputeLiveCount();
  watch(doc, () => {
    throttle.schedule();
  });

  onUnmount(() => {
    throttle.cancel();
  });

  // --- selection-aware count (AC-3) ----------------------------------------

  const selectionActive = ref(false);
  const selectionWords = ref(0);

  /** Called from DocumentView's updateCaretStatus alongside its other
   * per-selection state refreshes — the single place every caret/selection
   * change already flows through. */
  function setSelection(start: CaretPosition | null, end: CaretPosition | null) {
    const model = doc.value;
    if (!model || !start || !end) {
      selectionActive.value = false;
      selectionWords.value = 0;
      return;
    }
    const selection: ModelSelection = { anchor: start, focus: end };
    if (isCollapsed(selection)) {
      selectionActive.value = false;
      selectionWords.value = 0;
      return;
    }
    selectionActive.value = true;
    selectionWords.value = selectionWordCount(model, selection);
  }

  /** The status bar's text (AC-1, AC-3): "N of M words" while a selection
   * is active, else the document's own total. Literal "words" (not
   * singular/plural) in the selection form, matching the AC's own wording. */
  const statusBarText = computed(() => {
    const total = liveWordCount.value;
    if (selectionActive.value) {
      return `${selectionWords.value} of ${total} words`;
    }
    return `${total} ${total === 1 ? "word" : "words"}`;
  });

  // --- statistics dialog (AC-4..12) ----------------------------------------

  const statisticsOpen = ref(false);
  /** "Include footnotes and endnotes" (AC-11); resets to off each time the
   * dialog reopens, matching TocOptionsDialog/PageSetupDialog's own
   * per-open-prefill convention rather than persisting across sessions. */
  const includeNotes = ref(false);
  const dialogStats = ref<CountStats>(emptyCountStats());

  let dialogRequestToken = 0;

  async function refreshDialogStats() {
    const model = doc.value;
    const token = ++dialogRequestToken;
    const stats = model ? await computeDocumentStats(model, { includeNotes: includeNotes.value }) : emptyCountStats();
    if (token !== dialogRequestToken) {
      return;
    }
    dialogStats.value = stats;
  }

  watch([doc, includeNotes], () => {
    if (statisticsOpen.value) {
      void refreshDialogStats();
    }
  });

  function openStatistics() {
    includeNotes.value = false;
    dialogStats.value = emptyCountStats();
    statisticsOpen.value = true;
    void refreshDialogStats();
  }

  function closeStatistics() {
    statisticsOpen.value = false;
  }

  return {
    liveWordCount,
    statusBarText,
    announcement,
    setSelection,
    statisticsOpen,
    includeNotes,
    dialogStats,
    /** Exposed for dialogStats' other consumer (SPEC-041's Properties
     * dialog Statistics tab): before the worker move dialogStats was a
     * computed that tracked `doc` on its own, so merely rendering it was
     * enough — now anything that shows it must ask for a refresh. */
    refreshDialogStats,
    openStatistics,
    closeStatistics,
  };
}
