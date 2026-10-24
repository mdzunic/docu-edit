// Find and replace in the deck (SPEC-098 "View wiring"): the state behind
// SPEC-017's FindBar — generalised by its scope props, never forked — plus the
// grouped results list (DeckSearchResults.vue, rows from deckSearchRows.ts)
// and the stage highlights (SlideProofingLayer.vue). lib/presentation/
// deckSearch.ts decides what matches, in what order, and what a replace
// writes; deckSearchWalk.ts runs the walk.
//
// - Keys (window): Cmd/Ctrl+F opens the bar, Cmd/Ctrl+H with the replace row
//   (find.open / find.openReplace through findShortcut; preventDefault, so
//   the browser's find never opens over the deck). The bar's Escape closes it
//   and hands focus back to the stage.
// - The walk: the query and every option debounce FIND_DEBOUNCE_MS, and the
//   performance clock (`walkTiming`) starts when that fires; an edit to the
//   deck re-walks after the same debounce, keeping the current hit.
// - Navigation: a click or Enter on a result, or find-next/prev (wrapping
//   with "Wrapped to top"/"Wrapped to bottom"), makes a hit current: its
//   slide becomes current, its frame is selected and a text session opens on
//   it with the match selected; a notes hit opens the notes panel and selects
//   the range in its textarea. Stepping from the bar hands focus back to the
//   bar, so a second Enter steps again rather than typing over the match.
//   Exactly one slide render, no mutation, no undo entry.
// - Replace: the current hit (one undo entry, "Replace"), then on to the next;
//   Replace all: every hit as ONE undo entry ("Replace all"), reported as
//   "Replaced N occurrences". A running text session ends first, so its DOM
//   can never write stale text back over a replacement.
// - Every option persists for the view session; closing clears the hits.
import { computed, nextTick, onScopeDispose, ref, shallowRef, watch, type Ref } from "vue";
import { compileQuery, isQueryError, type CompiledQuery } from "../../lib/core/findQuery";
import { findShortcut } from "../../lib/shortcuts";
import type { PresentationModel } from "../../lib/presentation/model";
import { DECK_SCOPE_LABEL, deckScopeTitle, findInDeck, isHitCurrent, replaceAllInDeck, replaceHit, type DeckHit } from "../../lib/presentation/deckSearch";
import { measureStageRange, type HighlightRect } from "../../components/presentation/proofingRects";
import { createDeckWalk, type DeckWalkDeps, type WalkTiming } from "./deckSearchWalk";
import { deckResultGroups } from "./deckSearchRows";
import type { StageSearchHooks } from "./stageSearchHooks";

/** The input debounce; the performance clock starts when it fires. */
export const FIND_DEBOUNCE_MS = 150;
export const WRAPPED_TO_TOP = "Wrapped to top";
export const WRAPPED_TO_BOTTOM = "Wrapped to bottom";
export const SEARCHING_HINT = "Searching…";
export const FIND_UNDO_LABELS = { replace: "Replace", replaceAll: "Replace all" } as const;

export interface DeckSearchDeps extends StageSearchHooks, DeckWalkDeps {
  model: Ref<PresentationModel | null>;
  currentIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
}

/** Two hits starting at the same place (slide, target, paragraph, offset). */
const samePlace = (a: DeckHit, b: DeckHit): boolean =>
  a.slideIndex === b.slideIndex && a.scope === b.scope && a.frameId === b.frameId && a.row === b.row &&
  a.column === b.column && a.para === b.para && a.start === b.start;
const sameHit = (a: DeckHit, b: DeckHit): boolean => samePlace(a, b) && a.end === b.end;

/** Fresh identities for every touched slide and its frames so the stage and
 * the filmstrip re-render them (commit() re-slices the array). */
function reidentify(model: PresentationModel, touched: DeckHit[]): void {
  for (const slideIndex of new Set(touched.map((hit) => hit.slideIndex))) {
    const slide = model.slides[slideIndex];
    if (slide) model.slides[slideIndex] = { ...slide, frames: slide.frames.map((frame) => ({ ...frame })) };
  }
}

export function useDeckSearch(deps: DeckSearchDeps) {
  const open = ref(false);
  const expanded = ref(false);
  const query = ref("");
  const replaceText = ref("");
  const matchCase = ref(false);
  const wholeWords = ref(false);
  const useRegex = ref(false);
  /** The walk's hits so far — a fresh array on every change. */
  const hits = shallowRef<DeckHit[]>([]);
  const currentHitIndex = ref(-1);
  /** A sliced walk is running (the deck-search-progress line). */
  const indexing = ref(false);
  /** A transient line: a wrap or a replace report. */
  const message = ref("");
  const highlightRects = shallowRef<HighlightRect[]>([]);
  const walkTiming = shallowRef<WalkTiming | null>(null);

  const compiled = computed(() => compileQuery(query.value, { matchCase: matchCase.value, wholeWords: wholeWords.value, regex: useRegex.value }));
  const error = computed(() => (compiled.value !== null && isQueryError(compiled.value) ? compiled.value.error : ""));
  const runnable = (): CompiledQuery | null => (compiled.value !== null && !isQueryError(compiled.value) ? compiled.value : null);

  // --- the walk --------------------------------------------------------------
  const walker = createDeckWalk(deps);
  let debounce = 0;

  function cancelWalk(): void {
    if (debounce) clearTimeout(debounce);
    debounce = 0;
    walker.cancel();
    indexing.value = false;
  }

  /** Walks now (the debounce fired or the bar opened); the hit that was
   * current stays current when the new walk still finds it. An invalid or
   * empty query runs nothing. */
  function walk(): void {
    const previous = hits.value[currentHitIndex.value] ?? null;
    cancelWalk();
    const model = deps.model.value;
    const q = runnable();
    if (!open.value || !model || !q) {
      hits.value = [];
      currentHitIndex.value = -1;
      return;
    }
    walker.start(model, q, (found, done, timing) => {
      hits.value = found;
      indexing.value = !done;
      walkTiming.value = timing;
      currentHitIndex.value = done && previous ? found.findIndex((hit) => sameHit(hit, previous)) : -1;
    });
  }

  /** Debounces a walk (typing, an option, an edit to the deck). */
  function soon(): void {
    if (!open.value) return;
    if (debounce) clearTimeout(debounce);
    debounce = window.setTimeout(() => {
      debounce = 0;
      walk();
    }, FIND_DEBOUNCE_MS);
  }

  /** Completes a pending or running walk at once — a step or a replace needs
   * every hit of the query now. */
  function flush(): void {
    if (!debounce && !indexing.value) return;
    const previous = hits.value[currentHitIndex.value] ?? null;
    cancelWalk();
    const model = deps.model.value;
    const q = runnable();
    hits.value = open.value && model && q ? findInDeck(model, q) : [];
    currentHitIndex.value = previous ? hits.value.findIndex((hit) => sameHit(hit, previous)) : -1;
  }

  watch([query, matchCase, wholeWords, useRegex], () => {
    message.value = "";
    currentHitIndex.value = -1;
    soon();
  });
  // Every commit re-slices `slides`, so an edit anywhere re-walks.
  watch(() => deps.model.value?.slides, soon);

  // --- the stage highlights ----------------------------------------------------
  function measure(): void {
    const surface = deps.surface();
    if (!open.value || !surface) {
      highlightRects.value = [];
      return;
    }
    // Inside a focused session the browser's own selection shows the match.
    const active = surface.ownerDocument.activeElement;
    const sessionFocused = active instanceof HTMLElement && active.isContentEditable && surface.contains(active);
    const editing = deps.editingFrameId();
    const rects: HighlightRect[] = [];
    hits.value.forEach((hit, index) => {
      if (hit.slideIndex !== deps.currentIndex.value || hit.scope === "notes" || (sessionFocused && hit.frameId === editing)) return;
      for (const rect of measureStageRange(surface, hit)) rects.push({ ...rect, current: index === currentHitIndex.value });
    });
    highlightRects.value = rects;
  }
  /** Re-measures after the next render (a slide change, a stage rescale). */
  const remeasure = (): void => void nextTick(measure);
  watch([hits, currentHitIndex, deps.currentIndex, open, () => deps.editingFrameId()], remeasure);
  const onFocusIn = (): void => void (open.value && remeasure());
  if (typeof window !== "undefined") window.addEventListener("focusin", onFocusIn);

  // --- navigation ----------------------------------------------------------------
  async function activate(hit: DeckHit, fromBar: boolean): Promise<void> {
    const returnFocus = fromBar && typeof document !== "undefined" ? document.activeElement : null;
    deps.stopEditing();
    deps.currentIndex.value = hit.slideIndex;
    // The slide renders (dropping its stale selection) before we select.
    await nextTick();
    if (hit.scope === "notes") {
      await deps.openNotes(hit);
    } else if (hit.frameId) {
      deps.selectFrame(hit.frameId);
      deps.startEditing(hit);
    }
    await nextTick();
    if (returnFocus instanceof HTMLElement && returnFocus.isConnected) returnFocus.focus();
    remeasure();
  }

  /** Makes hit `index` current and shows it (a result's click or Enter). */
  function goToHit(index: number, fromBar = false): void {
    const hit = hits.value[index];
    if (!hit) return;
    if (!fromBar) message.value = "";
    currentHitIndex.value = index;
    void activate(hit, fromBar);
  }

  function step(direction: 1 | -1): void {
    flush();
    const list = hits.value;
    if (list.length === 0 || error.value !== "") return;
    let index = currentHitIndex.value + direction;
    if (currentHitIndex.value < 0) {
      // From the current slide: its first hit forward, the last before it back.
      const here = list.findIndex((hit) => hit.slideIndex >= deps.currentIndex.value);
      index = (here >= 0 ? here : list.length) - (direction > 0 ? 0 : 1);
    }
    message.value = index >= list.length ? WRAPPED_TO_TOP : index < 0 ? WRAPPED_TO_BOTTOM : "";
    goToHit(index >= list.length ? 0 : index < 0 ? list.length - 1 : index, true);
  }
  const next = (): void => step(1);
  const prev = (): void => step(-1);

  // --- replace -------------------------------------------------------------------
  function replaceCurrent(): void {
    flush();
    const model = deps.model.value;
    const q = runnable();
    const index = currentHitIndex.value;
    const hit = hits.value[index];
    if (!model || !q || !hit) return;
    deps.stopEditing();
    if (!isHitCurrent(model, hit)) {
      walk();
      return;
    }
    deps.record(FIND_UNDO_LABELS.replace);
    replaceHit(model, hit, replaceText.value);
    reidentify(model, [hit]);
    deps.commit();
    cancelWalk();
    hits.value = findInDeck(model, q);
    // Hits before the replaced one are unchanged; the replacement itself,
    // when it matches the query again, is stepped over.
    const at = hits.value[index];
    const following = at && samePlace(at, hit) ? index + 1 : index;
    if (hits.value.length === 0) {
      currentHitIndex.value = -1;
      message.value = "";
      return;
    }
    message.value = following >= hits.value.length ? WRAPPED_TO_TOP : "";
    goToHit(following >= hits.value.length ? 0 : following, true);
  }

  function replaceAll(): void {
    const model = deps.model.value;
    const q = runnable();
    if (!model || !q) return;
    deps.stopEditing();
    cancelWalk();
    const all = findInDeck(model, q);
    let count = 0;
    if (all.length > 0) {
      deps.record(FIND_UNDO_LABELS.replaceAll);
      count = replaceAllInDeck(model, all, replaceText.value);
      reidentify(model, all);
      deps.commit();
    }
    hits.value = findInDeck(model, q);
    currentHitIndex.value = -1;
    message.value = `Replaced ${count} occurrences`;
  }

  // --- the bar ---------------------------------------------------------------------
  function openFind(withReplace = false): void {
    if (withReplace) expanded.value = true;
    if (!open.value) {
      open.value = true;
      walk();
    } else if (!withReplace && typeof document !== "undefined") {
      // Already open: back to the query, selected (the bar focuses it only on opening).
      void nextTick(() => {
        const field = document.querySelector<HTMLInputElement>('[data-testid="find-query"]');
        field?.focus();
        field?.select();
      });
    }
  }

  function closeFind(): void {
    open.value = false;
    cancelWalk();
    hits.value = [];
    currentHitIndex.value = -1;
    message.value = "";
    highlightRects.value = [];
    deps.focusStage();
  }

  const toggle = (): void => (open.value ? closeFind() : openFind());
  const toggleExpand = (): void => void (expanded.value = !expanded.value);

  function onGlobalKeydown(event: KeyboardEvent): void {
    const id = findShortcut(event)?.id;
    if (id !== "find.open" && id !== "find.openReplace") return;
    if (event.target instanceof Element && event.target.closest('[role="dialog"]')) return;
    event.preventDefault();
    openFind(id === "find.openReplace");
  }

  // --- derived state ------------------------------------------------------------------
  const hitCount = computed(() => hits.value.length);
  const scopeTitle = computed(() => deckScopeTitle(deps.model.value));
  const hint = computed(() => message.value || (indexing.value ? SEARCHING_HINT : ""));
  const groups = computed(() => deckResultGroups(deps.model.value, hits.value));

  /** FindBar's props, bound with v-bind. */
  const bar = computed(() => ({
    visible: open.value,
    expanded: expanded.value,
    query: query.value,
    replaceText: replaceText.value,
    matchCase: matchCase.value,
    wholeWords: wholeWords.value,
    useRegex: useRegex.value,
    matchCount: hitCount.value,
    currentIndex: currentHitIndex.value,
    error: error.value,
    hint: hint.value,
    scopeLabel: DECK_SCOPE_LABEL,
    scopeTitle: scopeTitle.value,
    placeholder: "Find in presentation",
  }));

  onScopeDispose(() => {
    cancelWalk();
    if (typeof window !== "undefined") window.removeEventListener("focusin", onFocusIn);
  });

  return {
    open,
    expanded,
    query,
    replaceText,
    matchCase,
    wholeWords,
    useRegex,
    error,
    hint,
    hits,
    groups,
    currentHitIndex,
    hitCount,
    scopeLabel: DECK_SCOPE_LABEL,
    scopeTitle,
    indexing,
    highlightRects,
    walkTiming,
    bar,
    /** FindBar's events, bound with v-on. */
    barEvents: {
      "update:query": (text: string): void => void (query.value = text),
      "update:replaceText": (text: string): void => void (replaceText.value = text),
      "update:matchCase": (on: boolean): void => void (matchCase.value = on),
      "update:wholeWords": (on: boolean): void => void (wholeWords.value = on),
      "update:useRegex": (on: boolean): void => void (useRegex.value = on),
      "toggle-expand": toggleExpand,
      next,
      prev,
      replace: replaceCurrent,
      "replace-all": replaceAll,
      close: closeFind,
    },
    openFind,
    closeFind,
    toggle,
    toggleExpand,
    next,
    prev,
    goToHit,
    replaceCurrent,
    replaceAll,
    remeasure,
    onGlobalKeydown,
  };
}

export type DeckSearch = ReturnType<typeof useDeckSearch>;

/** Narrows a component's runtime-typed `controller` prop (the .vue parser
 * bars casts). */
export function asDeckSearch(value: unknown): DeckSearch {
  return value as DeckSearch;
}
