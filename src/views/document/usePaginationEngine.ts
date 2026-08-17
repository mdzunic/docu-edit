// --- the pagination engine (SPEC-028, SPEC-031, SPEC-047) -------------------
// The controller's session lifecycle, as opposed to its algorithm (which lives
// in composables/usePagination.ts) and its render derivation (usePaginationView
// next door): construct the controller over this session's document, decide
// WHEN it re-lays-out, and expose the handful of layout readouts the rest of
// the view derives from — the page list, the busy/settled flags, the open
// progress indicator, and the displayed page numbers.
//
// Note numbering lives here too, and not in useNotes, because the two are
// mutually recursive: a footnote's label depends on which page its reference
// landed on (`footnoteRefPages`, "restart each page"), and the layout depends
// on the label — a mark's rendered width differs per number. usePagination
// reads the labels through a callback so the cycle resolves lazily, and the
// invalidation watcher below closes it: a label that CHANGES must drop the
// measurements taken with the old one. Splitting the two apart would put that
// cycle across two files with nothing naming it.
//
// The deps are deliberately thin — doc, the canvas root and the outline's
// hidden set. Everything else in the view consumes what this returns.
import { computed, onBeforeUnmount, onMounted, ref, watch } from "vue";
import type { Ref } from "vue";
import { usePagination } from "../../composables/usePagination";
import { collectNoteRefs, noteLabels, noteText } from "../../lib/notes";
import { computePageNumbers } from "../../lib/pageNumbering";
import type { EditorContext } from "./context";

export interface PaginationEngineDeps extends EditorContext {
  /** useOutlinePane's collapsed-heading hidden set (SPEC-037 AC-16): the
   * blocks a collapsed heading hides are laid out as if they were not there.
   * Wired from the composable above this one in DocumentView. */
  hiddenBodyBlocks: Readonly<Ref<ReadonlySet<string>>>;
}

export function usePaginationEngine(deps: PaginationEngineDeps) {
  const { doc, documentId, contentRoot, hiddenBodyBlocks } = deps;

  // The controller owns measurement; every doc.value replacement re-lays-out
  // synchronously (flush: "sync") so the rendered pages always agree with the
  // new model's block indices before Vue patches the canvas.
  const pagination = usePagination(
    () => doc.value,
    () => contentRoot.value as HTMLElement | null,
    (noteId) => noteLabelsMap.value.get(noteId) ?? "",
    () => hiddenBodyBlocks.value,
  );
  // A body-heading collapse/expand changes only the hidden set, not the model,
  // so re-lay-out explicitly when it changes (AC-16). Empty-to-empty is not a
  // change: the computed mints a fresh Set object whenever `doc` is replaced
  // (opening a document), and relaying out for that fused a second full
  // relayout into the open mount task (SPEC-047 AC-16) for a hidden set that
  // is still "nothing hidden".
  watch(hiddenBodyBlocks, (now, prev) => {
    if (now.size === 0 && prev !== undefined && prev.size === 0) {
      return;
    }
    pagination.schedule();
  });
  const paginationPages = pagination.pages;
  const paginationBusy = pagination.busy;
  const paginationState = pagination.state;

  // --- open progress indicator (SPEC-047 AC-3..5): shown from this
  // document's mount until pagination first reports settled, distinct from
  // paginationBusy above (which reappears on every later edit-triggered
  // relayout too — this one only ever fires once per document open). ------
  const documentOpening = ref(true);
  watch(documentId, () => {
    documentOpening.value = true;
  });
  watch(paginationState, (state) => {
    if (state === "clean") {
      documentOpening.value = false;
    }
  });
  // The "never needs to measure" fallback lives in the initial-layout kick
  // below (see the doc watcher): a document whose pagination never
  // transitions through "measuring" gives the state watcher above no change
  // to observe, so the kick checks once, right after the first layout ran.

  // --- note numbering (SPEC-031 AC-7..9, AC-17..20) ---------------------------
  // Displayed note numbers, derived fresh from mark order and the numbering
  // settings on every model change — never persisted, like list numbering.
  // Footnotes thread the per-page assignment through for "restart each page".
  const noteLabelsMap = computed(() => {
    const model = doc.value;
    if (!model) {
      return new Map<string, string>();
    }
    const refPages = pagination.footnoteRefPages.value;
    const labels = noteLabels(model, "footnote", collectNoteRefs(model, "footnote"), (ref, index) => refPages[index] ?? 0);
    for (const [noteId, label] of noteLabels(model, "endnote", collectNoteRefs(model, "endnote"))) {
      labels.set(noteId, label);
    }
    return labels;
  });

  /** Reference-mark tooltips (AC-14): each note's plain text. */
  const noteTooltipsMap = computed(() => {
    const model = doc.value;
    const tooltips = new Map<string, string>();
    for (const note of model?.notes ?? []) {
      tooltips.set(note.noteId, noteText(note));
    }
    return tooltips;
  });

  // A mark's rendered label width differs per number, and the measurement
  // cache is keyed by block object identity — so when an existing note's
  // label CHANGES (an earlier note inserted/deleted renumbers it), the
  // paragraphs holding those marks must re-measure. New ids (a fresh insert)
  // need none: nothing measured them yet.
  let knownNoteLabels = new Map<string, string>();
  watch(noteLabelsMap, (labels) => {
    let changed = false;
    for (const [noteId, label] of labels) {
      const known = knownNoteLabels.get(noteId);
      if (known !== undefined && known !== label) {
        changed = true;
        break;
      }
    }
    knownNoteLabels = new Map(labels);
    if (changed) {
      pagination.invalidateAll();
    }
  });

  watch(doc, () => pagination.schedule(), { flush: "sync", immediate: true });
  // The controller starts un-primed: every layout request up to here (the
  // immediate watcher above, document.fonts.ready, anything else that fires
  // during setup/mount) is only RECORDED, so the whole estimate pass, first
  // page render and viewport refresh stop fusing into the component's own
  // mount task — the largest main-thread block of the open phase (SPEC-047
  // AC-16). prime() runs the recorded layout in its own frame task; the
  // rAF-less fallback keeps the old synchronous mount timing for
  // non-browser environments.
  let initialLayoutFrame = 0;
  let initialLayoutTimer = 0;
  let initialLayoutIdle = 0;
  onMounted(() => {
    const kick = () => {
      pagination.prime();
      // A document whose pagination never needs to measure never transitions
      // through "measuring", so the paginationState watcher above never
      // observes a change — checked once here, right after the first layout.
      if (paginationState.value === "clean") {
        documentOpening.value = false;
      }
    };
    // canMeasureOffscreen doubles as the "real rendering environment" probe
    // (it is false exactly where canvas measurement — and painting — do not
    // exist, i.e. jsdom): only a painting browser gains anything from
    // deferring the first layout past the mount task, and the unit suites
    // interact with the canvas synchronously after mount. requestIdleCallback
    // (not rAF, not a bare timer) is the deferral that actually separates the
    // tasks: a same-frame rAF callback runs in the mount task's own rendering
    // update, and a timer can beat the frame — both fuse the layout right
    // back into the block it was deferred out of. The timeout bounds how long
    // a busy frame can hold the first layout back.
    if (pagination.canMeasureOffscreen && typeof window.requestAnimationFrame === "function") {
      if (typeof window.requestIdleCallback === "function") {
        initialLayoutIdle = window.requestIdleCallback(kick, { timeout: 100 });
      } else {
        initialLayoutFrame = window.requestAnimationFrame(() => {
          initialLayoutFrame = 0;
          initialLayoutTimer = window.setTimeout(kick, 0);
        });
      }
    } else {
      kick();
    }
  });
  onBeforeUnmount(() => {
    if (initialLayoutFrame !== 0 && typeof window.cancelAnimationFrame === "function") {
      window.cancelAnimationFrame(initialLayoutFrame);
    }
    if (initialLayoutTimer !== 0) {
      window.clearTimeout(initialLayoutTimer);
    }
    if (initialLayoutIdle !== 0 && typeof window.cancelIdleCallback === "function") {
      window.cancelIdleCallback(initialLayoutIdle);
    }
    pagination.stop();
  });

  const pageCount = computed(() => Math.max(1, paginationPages.value.length));
  // The displayed page number of every laid-out page (SPEC-029 AC-20..23),
  // recomputed whenever pagination or a section's numbering options change.
  const pageNumbers = computed(() => {
    const model = doc.value;
    return model ? computePageNumbers(paginationPages.value, model) : [];
  });

  return {
    pagination,
    paginationPages,
    paginationBusy,
    paginationState,
    documentOpening,
    noteLabelsMap,
    noteTooltipsMap,
    pageCount,
    pageNumbers,
  };
}
