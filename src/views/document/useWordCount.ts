// Live word count and the statistics dialog (SPEC-036). Owns a per-block
// count cache the same way usePagination.ts owns its measurement cache (a
// plain `WeakMap<Block, CountStats>` local to this composable instance) so
// a keystroke re-tokenizes only the paragraph it touched (AC-26). The
// document-wide total recomputes on a throttle (AC-2: never delays input);
// the selection-aware count recomputes synchronously on every selection
// change (DocumentView's updateCaretStatus calls setSelection alongside its
// other per-selection state, cheap since a selection is normally small).
import { computed, ref, watch } from "vue";
import type { Ref } from "vue";
import { isCollapsed, type ModelSelection } from "../../lib/commands";
import type { CaretPosition } from "../../lib/textEditing";
import type { Block, DocumentModel } from "../../lib/documentModel";
import {
  documentBodyStats,
  documentStats,
  emptyCountStats,
  selectionWordCount,
  type CountStats,
} from "../../lib/wordCount";

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

  const cache = new WeakMap<Block, CountStats>();

  const liveWordCount = ref(0);
  /** The visible/hidden status text's last announced word count — a
   * screen-reader-facing region only re-announces when its text actually
   * changes, so binding it to a value that only updates on a real change
   * (rather than every throttle tick) is what makes the announcement
   * "material" (AC-28) without any separate threshold bookkeeping. */
  const announcement = ref("0 words");

  function recomputeLiveCount() {
    const model = doc.value;
    const words = model ? documentBodyStats(model, cache).words : 0;
    if (words === liveWordCount.value) {
      return;
    }
    liveWordCount.value = words;
    announcement.value = `${words} ${words === 1 ? "word" : "words"}`;
  }

  const throttle = createThrottle(recomputeLiveCount, THROTTLE_MS);

  // The first count (opening/creating a document) is visible immediately
  // rather than waiting out the throttle window; every later change goes
  // through the throttle.
  recomputeLiveCount();
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

  const dialogStats = computed<CountStats>(() => {
    const model = doc.value;
    if (!model) {
      return emptyCountStats();
    }
    return documentStats(model, { includeNotes: includeNotes.value }, cache);
  });

  function openStatistics() {
    includeNotes.value = false;
    statisticsOpen.value = true;
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
    openStatistics,
    closeStatistics,
  };
}
