// Background spell checking (SPEC-035). Mirrors useFindReplace.ts's shape
// (idle-scheduled measurement, viewport-limited work, an overlay rect list
// DocumentView paints) but the "search" here is a per-visible-paragraph
// dictionary scan instead of a cheap whole-document string match — so, per
// AC-16.1/AC-16.2, checking itself (not just DOM measurement) is scoped to
// the visible region plus a margin, and results are cached per paragraph so
// re-scrolling past already-checked text costs nothing.
//
// Scope: body paragraphs only (model.sections[i].blocks), matching what
// callers can address with a plain {section, block, offset} CaretPosition
// (no `cell` — table-cell/header/footer/note paragraphs are out of scope
// for this pass, a deliberate limitation noted in the implementation
// summary rather than a partial/half-built attempt at covering them).
import { shallowRef, watch } from "vue";
import type { Ref } from "vue";
import type { DocumentModel, Paragraph } from "../../lib/documentModel";
import { effectiveLanguage, isDictionaryAvailable, normalizeLanguageTag } from "../../lib/language";
import { loadDictionary } from "../../lib/spellcheck/dictionaries";
import type { Dictionary } from "../../lib/spellcheck/dictionary";
import { isCodeStyledFont } from "../../lib/spellcheck/skip";
import { getSuggestions } from "../../lib/spellcheck/suggestions";
import { checkTextsInWorker } from "../../lib/workers/spellCheckClient";
import {
  addCustomWord,
  customDictionaryEntries,
  hasCustomWord,
  removeCustomWord,
  subscribeCustomDictionary,
  type CustomDictionaryEntry,
} from "../../lib/customDictionary";
import { noIdleHandle, scheduleIdleWork } from "../../lib/idleScheduler";
import { BLOCK_SELECTOR, domPointForCaret } from "../../lib/domCaret";

/** A flagged word, addressed the same way a bookmark/find-match is: a
 * paragraph offset range into its concatenated run text. */
export interface SpellCheckIssue {
  section: number;
  block: number;
  start: number;
  end: number;
  word: string;
  lang: string;
}

/** One on-screen underline strip (SPEC-035 AC-1): positioned at the bottom
 * of a word's glyph box, page-relative, the same coordinate space
 * useFindReplace's HighlightRect uses. Carries the issue it renders so
 * DocumentView's right-click hit test (SPEC-035 AC-2) can hand the whole
 * issue to the suggestions menu straight from the rect it matched. */
export interface SpellUnderlineRect {
  left: number;
  top: number;
  width: number;
  height: number;
  section: number;
  block: number;
  start: number;
  end: number;
  word: string;
  lang: string;
}

/** The spell context menu's state (SPEC-035 AC-2, AC-3): a typed empty
 * seed for DocumentView.vue's `ref()`, the same convention as
 * lib/findReplace.ts's noHighlightRects — DocumentView.vue's script is
 * espree-parsed (see its own notes on `asCharacterFormatKey`), so it cannot
 * write the array-of-strings annotation `suggestions` needs inline. */
export interface SpellContextMenuState {
  visible: boolean;
  left: number;
  top: number;
  section: number;
  block: number;
  start: number;
  end: number;
  word: string;
  lang: string;
  suggestions: string[];
}

export function noSpellContextMenu(): SpellContextMenuState {
  return {
    visible: false,
    left: 0,
    top: 0,
    section: 0,
    block: 0,
    start: 0,
    end: 0,
    word: "",
    lang: "",
    suggestions: [],
  };
}

/** Typed empty seed for a SpellUnderlineRect default parameter in
 * DocumentView.vue's espree-parsed script (same convention as
 * noSpellContextMenu). */
export function noSpellUnderlineRect(): SpellUnderlineRect {
  return { left: 0, top: 0, width: 0, height: 0, section: 0, block: 0, start: 0, end: 0, word: "", lang: "" };
}

// SPEC-035 AC-16.1: the visible region plus a 500px margin above/below.
const VIEWPORT_MARGIN_PX = 500;

/** Identifies one specific occurrence for "Ignore" (SPEC-035 AC-5) — a
 * module-level pure function (captures nothing from useSpellCheck's
 * closure). */
function occurrenceKey(section: number, block: number, start: number, word: string): string {
  return `${section}:${block}:${start}:${word.toLowerCase()}`;
}

export interface SpellCheckDeps {
  doc: Ref<DocumentModel | null>;
  contentRoot: Ref<Element | null>;
  pageRef: Ref<Element | null>;
  getScrollContainer: () => Element | null;
  onScroll: (cb: () => void) => void;
  onSessionClose: (cb: () => void) => void;
  onUnmount: (cb: () => void) => void;
}

export function useSpellCheck(deps: SpellCheckDeps) {
  const { doc, contentRoot, pageRef, getScrollContainer, onScroll, onSessionClose, onUnmount } = deps;

  const spellUnderlineRects = shallowRef<SpellUnderlineRect[]>([]);
  // The custom dictionary's live entries (SPEC-035 AC-8.1, AC-8.2), for the
  // settings dialog — refreshed by the same subscription that invalidates
  // the issue cache below.
  const customDictionaryList = shallowRef<CustomDictionaryEntry[]>([...customDictionaryEntries()]);

  // Per-paragraph issue cache ("section:block" -> issues), cleared wholesale
  // on every model change — cheap (a Map.clear()); only the visible subset
  // is ever recomputed (AC-16.1), so document size never matters here.
  let issueCache = new Map<string, SpellCheckIssue[]>();
  // Bumped by every cache invalidation below (never by runSpellCheck itself
  // — it only reads this at the start of a pass). A worker round trip means
  // a check started against one document/language state can resolve after
  // a *later* invalidation already changed what's true — the two
  // invalidateIssueCache*/spellCheckToken helpers make that resolution a
  // guaranteed no-op instead of silently overwriting fresh results with a
  // stale answer (SPEC-047's async spell-check rewrite surfaced this: a run
  // checked against German, then re-checked against English, could have the
  // German worker reply land second and win).
  let spellCheckToken = 0;

  function invalidateIssueCache(): void {
    issueCache.clear();
    spellCheckToken++;
  }

  function invalidateIssueCacheKey(key: string): void {
    issueCache.delete(key);
    spellCheckToken++;
  }
  // Dictionaries resolved so far, and the languages currently loading
  // (AC-17: on demand, not at startup) — module-external cache
  // (lib/spellcheck/dictionaries) does the actual fetch/memoize; these are
  // just this composable's synchronous view of what has resolved. The actual
  // dictionary *scan* (checkText) runs in a worker now (SPEC-047 AC-13,
  // spellCheckClient.ts), which loads and caches its own dictionaries
  // independently — this main-thread copy exists only so suggestionsFor
  // (the right-click menu, SPEC-035 AC-3.1) has a Dictionary object to rank
  // against without a worker round trip.
  const resolvedDictionaries = new Map<string, Dictionary>();
  const pendingLanguages = new Set<string>();

  // Session-only state (SPEC-035 AC-5, AC-6): "Ignore" is one occurrence,
  // "Ignore all" is every occurrence of a word in this document, for as
  // long as the document stays open. Neither persists (see
  // lib/customDictionary.ts for what does).
  const ignoredOccurrences = new Set<string>();
  const ignoredWords = new Set<string>();

  let spellIdleHandle = noIdleHandle();
  const unsubscribeCustomDictionary = subscribeCustomDictionary(() => {
    customDictionaryList.value = [...customDictionaryEntries()];
    invalidateIssueCache();
    scheduleSpellCheck();
  });

  function isIgnored(section: number, block: number, start: number, word: string): boolean {
    return (
      ignoredWords.has(word.toLowerCase()) ||
      ignoredOccurrences.has(occurrenceKey(section, block, start, word))
    );
  }

  function requestDictionary(lang: string): void {
    const normalized = normalizeLanguageTag(lang);
    if (resolvedDictionaries.has(normalized) || pendingLanguages.has(normalized)) {
      return;
    }
    const promise = loadDictionary(normalized);
    if (!promise) {
      return;
    }
    pendingLanguages.add(normalized);
    void promise.then((dictionary) => {
      pendingLanguages.delete(normalized);
      resolvedDictionaries.set(normalized, dictionary);
    });
  }

  /** The spelling issues in one paragraph, resolving each run's language
   * independently (SPEC-035 AC-10, AC-10.1) and skipping noProof/code-styled
   * runs (AC-11, AC-12.1, AC-12.2), tracked deletions (SPEC-039 AC-17) and
   * words already ignored or in the custom dictionary. A run whose language
   * has no bundled dictionary contributes nothing (AC-19). The dictionary
   * scan itself runs in a worker (SPEC-047 AC-13): every checkable run in
   * the paragraph batches into one round trip rather than one per run. */
  async function computeIssuesAsync(
    section: number,
    block: number,
    paragraph: Paragraph,
    model: DocumentModel,
  ): Promise<SpellCheckIssue[]> {
    const runStarts: number[] = [];
    const checkable: Array<{ key: string; lang: string; text: string; runStart: number }> = [];
    let offset = 0;
    for (const [runIndex, run] of paragraph.runs.entries()) {
      runStarts[runIndex] = offset;
      offset += run.text.length;
      if (!run.text || run.formatting.noProof || run.changes?.delete || isCodeStyledFont(run.formatting.fontFamily)) {
        continue;
      }
      const lang = effectiveLanguage(run.formatting.lang, model.defaultStyles.language);
      if (!isDictionaryAvailable(lang)) {
        continue;
      }
      // Warms the main-thread suggestions cache (see the field's own
      // comment) — independent of, and in parallel with, the worker's own
      // dictionary load below.
      requestDictionary(lang);
      checkable.push({ key: String(runIndex), lang: normalizeLanguageTag(lang), text: run.text, runStart: runStarts[runIndex] });
    }
    if (checkable.length === 0) {
      return [];
    }
    const results = await checkTextsInWorker(
      checkable.map(({ key, lang, text }) => ({ key, lang, text })),
    );
    const byKey = new Map(results.map((result) => [result.key, result.matches]));
    const issues: SpellCheckIssue[] = [];
    for (const { key, lang, runStart } of checkable) {
      for (const found of byKey.get(key) ?? []) {
        const start = runStart + found.start;
        if (isIgnored(section, block, start, found.word) || hasCustomWord(found.word, lang)) {
          continue;
        }
        issues.push({ section, block, start, end: runStart + found.end, word: found.word, lang });
      }
    }
    return issues;
  }

  function viewportBounds(): { top: number; bottom: number } {
    const scrollContainer = getScrollContainer();
    if (scrollContainer) {
      const rect = scrollContainer.getBoundingClientRect();
      return { top: rect.top, bottom: rect.bottom };
    }
    return { top: 0, bottom: window.innerHeight };
  }

  /** Body-paragraph blocks currently on screen (plus the margin), skipping
   * table-cell paragraphs (`data-row`) — out of scope, see the file header. */
  function visibleParagraphKeys(
    root: HTMLElement,
    bounds: { top: number; bottom: number },
  ): Array<{ section: number; block: number; key: string }> {
    const found: Array<{ section: number; block: number; key: string }> = [];
    for (const element of Array.from(root.querySelectorAll(BLOCK_SELECTOR))) {
      if (element.hasAttribute("data-row")) {
        continue;
      }
      const rect = element.getBoundingClientRect();
      if (rect.bottom < bounds.top - VIEWPORT_MARGIN_PX || rect.top > bounds.bottom + VIEWPORT_MARGIN_PX) {
        continue;
      }
      const section = Number(element.getAttribute("data-section"));
      const block = Number(element.getAttribute("data-block"));
      if (Number.isNaN(section) || Number.isNaN(block)) {
        continue;
      }
      found.push({ section, block, key: `${section}:${block}` });
    }
    return found;
  }

  function issueClientRects(root: HTMLElement, issue: SpellCheckIssue): DOMRect[] {
    const startPoint = domPointForCaret(root, { section: issue.section, block: issue.block, offset: issue.start });
    const endPoint = domPointForCaret(root, { section: issue.section, block: issue.block, offset: issue.end });
    if (!startPoint || !endPoint) {
      return [];
    }
    try {
      const range = root.ownerDocument.createRange();
      range.setStart(startPoint.node, startPoint.offset);
      range.setEnd(endPoint.node, endPoint.offset);
      return Array.from(range.getClientRects());
    } catch {
      // Defensive against a Range implementation missing getClientRects
      // (some test/embedding environments) — no rects just means nothing
      // paints this pass, not a crash.
      return [];
    }
  }

  // The full glyph-box rect is used both for the click/right-click hit area
  // and as the element the wavy underline paints into — CSS positions the
  // squiggle pattern at the box's own bottom edge (background-position:
  // left bottom) rather than this composable computing a separate thin
  // strip, so the clickable region is the whole word (AC-2: right-clicking
  // *a misspelled word*, not a 3px sliver under it) while the paint stays a
  // thin line.
  function pushIssueRects(
    rects: SpellUnderlineRect[],
    root: HTMLElement,
    pageRect: DOMRect,
    issue: SpellCheckIssue,
  ): void {
    for (const rect of issueClientRects(root, issue)) {
      if (rect.width <= 0 || rect.height <= 0) {
        continue;
      }
      rects.push({
        left: rect.left - pageRect.left,
        top: rect.top - pageRect.top,
        width: rect.width,
        height: rect.height,
        section: issue.section,
        block: issue.block,
        start: issue.start,
        end: issue.end,
        word: issue.word,
        lang: issue.lang,
      });
    }
  }

  /** Recomputes (from cache where possible) and measures the visible
   * region's issues. The expensive halves — dictionary scanning of
   * newly-visible paragraphs and DOM rect measurement — both only ever
   * touch the visible-plus-margin set (AC-16, AC-16.1, AC-16.2). Reads
   * `spellCheckToken` rather than bumping it: only invalidation (the model,
   * the language, ignore/custom-dictionary state changing) marks a pass
   * stale, not the mere act of running another one, so two passes racing
   * over an *unchanged* cache generation both still get to write. */
  async function runSpellCheck(): Promise<void> {
    const root = contentRoot.value as HTMLElement | null;
    const pageEl = pageRef.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !pageEl || !model) {
      spellUnderlineRects.value = [];
      return;
    }
    const token = spellCheckToken;
    const visible = visibleParagraphKeys(root, viewportBounds());
    const misses: Array<{ section: number; block: number; key: string; paragraph: Paragraph }> = [];
    for (const { section, block, key } of visible) {
      if (issueCache.has(key)) {
        continue;
      }
      const target = model.sections[section]?.blocks[block];
      if (target && target.type === "paragraph") {
        misses.push({ section, block, key, paragraph: target });
      }
    }
    if (misses.length > 0) {
      const computed = await Promise.all(
        misses.map(async ({ section, block, key, paragraph }) => ({
          key,
          issues: await computeIssuesAsync(section, block, paragraph, model),
        })),
      );
      if (token !== spellCheckToken) {
        return;
      }
      for (const { key, issues } of computed) {
        issueCache.set(key, issues);
      }
    }
    const pageRect = pageEl.getBoundingClientRect();
    const rects: SpellUnderlineRect[] = [];
    for (const { key } of visible) {
      for (const issue of issueCache.get(key) ?? []) {
        pushIssueRects(rects, root, pageRect, issue);
      }
    }
    spellUnderlineRects.value = rects;
  }

  /** Schedules (idle, off the typing path — AC-1.1) a recheck of the
   * visible region. Safe to call as often as needed; overlapping calls
   * coalesce into the next idle slot. */
  function scheduleSpellCheck(): void {
    if (spellIdleHandle) {
      return;
    }
    spellIdleHandle = scheduleIdleWork(
      () => {
        spellIdleHandle = null;
        void runSpellCheck();
      },
      { timeoutMs: 150, fallbackBudgetMs: 8, maxBudgetMs: 16 },
    );
  }

  function cancelSpellCheckScheduled(): void {
    if (spellIdleHandle) {
      spellIdleHandle.cancel();
      spellIdleHandle = null;
    }
  }

  /** "Ignore" (AC-5): clears the underline for this one occurrence, for the
   * session. */
  function ignoreOccurrence(issue: SpellCheckIssue): void {
    ignoredOccurrences.add(occurrenceKey(issue.section, issue.block, issue.start, issue.word));
    invalidateIssueCacheKey(`${issue.section}:${issue.block}`);
    scheduleSpellCheck();
  }

  /** "Ignore all" (AC-6): clears every occurrence of this word in the
   * document, for the session. */
  function ignoreAllOccurrences(word: string): void {
    ignoredWords.add(word.toLowerCase());
    invalidateIssueCache();
    scheduleSpellCheck();
  }

  /** "Add to dictionary" (AC-7, AC-7.1, AC-7.2): persists the word for its
   * language and clears its underline everywhere immediately. */
  function addWordToDictionary(word: string, lang: string): void {
    // addCustomWord's persist() calls subscribeCustomDictionary's listener
    // synchronously, which already clears the issue cache and reschedules.
    addCustomWord(word, lang);
  }

  /** Removes a word from the custom dictionary (SPEC-035 AC-8.2's settings
   * dialog) — re-flags it immediately via the same subscription. */
  function removeWordFromDictionary(word: string, lang: string): void {
    removeCustomWord(word, lang);
  }

  /** Up to five ranked suggestions for a flagged word (AC-2, AC-3.1). Empty
   * when the dictionary has not (or will never) resolve for its language —
   * callers only ever pass an `issue.lang` that was already resolved to
   * produce the issue in the first place, so this is effectively always
   * populated in practice. */
  function suggestionsFor(word: string, lang: string): string[] {
    const dictionary = resolvedDictionaries.get(normalizeLanguageTag(lang));
    return dictionary ? getSuggestions(word, dictionary) : [];
  }

  function resetSpellCheckSession(): void {
    ignoredOccurrences.clear();
    ignoredWords.clear();
    issueCache = new Map();
    spellCheckToken++;
    spellUnderlineRects.value = [];
    cancelSpellCheckScheduled();
  }

  watch(doc, () => {
    invalidateIssueCache();
    scheduleSpellCheck();
  });

  onScroll(scheduleSpellCheck);
  onSessionClose(resetSpellCheckSession);
  onUnmount(() => {
    cancelSpellCheckScheduled();
    unsubscribeCustomDictionary();
  });

  return {
    spellUnderlineRects,
    customDictionaryList,
    scheduleSpellCheck,
    ignoreOccurrence,
    ignoreAllOccurrences,
    addWordToDictionary,
    removeWordFromDictionary,
    suggestionsFor,
  };
}
