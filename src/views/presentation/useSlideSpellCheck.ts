// Spell check on slides (SPEC-098 "View wiring"): SPEC-035's checker,
// dictionaries, suggestion ranking and language model, reused — not hoisted —
// from this composition point. Only the current slide's text frames, shape
// text and table cells plus its speaker notes are ever checked, on idle and
// off the main thread (checkTextsInWorker); each run resolves its language
// through core's effectiveLanguage chain with `PresentationModel.language`
// standing in for the text model's document default, and a noProof run, a
// code-styled run or a language with no bundled dictionary is skipped.
// Dictionaries load lazily per language (loadDictionary caches them).
//
// Misspelled words on the stage get a red wavy underline in
// SlideProofingLayer.vue; a right-click on one opens SpellContextMenu.vue
// (≤ 5 suggestions, Ignore, Ignore all, Add to dictionary). A suggestion
// keeps the word's capitalisation pattern and is ONE undo entry, "Correct
// spelling". Ignore / Ignore all last for the view session; Add to
// dictionary persists (lib/customDictionary.ts).
//
// The language control (LanguageStatusControl.vue, unchanged) sets the deck
// language with nothing selected and the selection's run `lang` when text is
// selected (a text session, or a selected text frame — whole frame); "Do not
// check spelling" sets `noProof` on the selection. Labels: "Language", "Do
// not check spelling".
import { computed, nextTick, onScopeDispose, ref, shallowRef, watch, type ComputedRef, type Ref } from "vue";
import { DEFAULT_LANGUAGE, isDictionaryAvailable, languageLabel, normalizeLanguageTag, SUPPORTED_LANGUAGES } from "../../lib/core/language";
import { scheduleIdleWork, type IdleHandle } from "../../lib/core/idleScheduler";
import { resolveFont } from "../../lib/core/theme";
import { loadDictionary } from "../../lib/spellcheck/dictionaries";
import type { Dictionary } from "../../lib/spellcheck/dictionary";
import { isCodeStyledFont } from "../../lib/spellcheck/skip";
import { applyCasePattern, detectCasePattern, getSuggestions } from "../../lib/spellcheck/suggestions";
import { checkTextsInWorker, type SpellCheckItem } from "../../lib/workers/spellCheckClient";
import { addCustomWord, customDictionaryEntries, hasCustomWord, removeCustomWord, subscribeCustomDictionary, type CustomDictionaryEntry } from "../../lib/customDictionary";
import { usePreferences } from "../../composables/usePreferences";
import type { PresentationModel, Slide, SlideParagraph } from "../../lib/presentation/model";
import { paragraphText } from "../../lib/presentation/slideModel";
import { runLanguage } from "../../lib/presentation/slideLanguage";
import { deckThemeOf } from "../../lib/presentation/theme";
import { applyRunFormat, markSlideEdited, replaceParagraphText } from "../../lib/presentation/text";
import { measureStageRange, type SpellUnderlineRect, type StageTextRange } from "../../components/presentation/proofingRects";
import type { FrameSelection } from "./useFrameSelection";
import type { SlideTextEditing } from "./useSlideTextEditing";

export const SPELL_UNDO_LABELS = { correct: "Correct spelling", language: "Language", noProof: "Do not check spelling" } as const;

/** A flagged word on the current slide: in a frame/cell paragraph, or (with
 * `notes`) in speaker-notes line `para`. */
export interface SlideSpellIssue extends StageTextRange {
  notes?: true;
  word: string;
  lang: string;
}

export interface SlideSpellMenu {
  left: number;
  top: number;
  word: string;
  suggestions: string[];
  issue: SlideSpellIssue;
}

export interface SlideSpellCheckDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
  selection: FrameSelection;
  /** The text session: its selection's formatting state and its one-entry
   * command runner, for the language control. */
  editing: SlideTextEditing;
  /** The current slide's notes as the panel's newline-joined text. */
  notesText: ComputedRef<string>;
  /** The stage surface the underlines are measured in. */
  surface: () => HTMLElement | null;
  /** Test seam: the idle scheduler. */
  schedule?: (callback: () => void) => IdleHandle;
}

/** The paragraphs a frame/cell range points into, or null when gone. */
function rangeParagraphs(slide: Slide | null, range: StageTextRange): SlideParagraph[] | null {
  const frame = slide?.frames.find((each) => each.id === range.frameId);
  if (frame?.kind === "text" || frame?.kind === "shape") return frame.paragraphs ?? null;
  if (frame?.kind === "table") {
    const cell = frame.rows[range.row ?? -1]?.cells[range.column ?? -1];
    return cell && !cell.covered ? cell.paragraphs : null;
  }
  return null;
}

/** Every checkable paragraph on a slide, addressed as stage ranges. */
function slideParagraphs(slide: Slide): { at: Omit<StageTextRange, "start" | "end">; paragraph: SlideParagraph }[] {
  const out: { at: Omit<StageTextRange, "start" | "end">; paragraph: SlideParagraph }[] = [];
  for (const frame of slide.frames) {
    if (frame.kind === "text" || frame.kind === "shape") {
      (frame.paragraphs ?? []).forEach((paragraph, para) => out.push({ at: { frameId: frame.id, para }, paragraph }));
    } else if (frame.kind === "table") {
      frame.rows.forEach((row, r) =>
        row.cells.forEach((cell, c) => {
          if (!cell.covered) cell.paragraphs.forEach((paragraph, para) => out.push({ at: { frameId: frame.id, row: r, column: c, para }, paragraph }));
        }),
      );
    }
  }
  return out;
}

const occurrenceKey = (slide: Slide, issue: SlideSpellIssue): string =>
  `${slide.id}|${issue.notes ? "notes" : issue.frameId}|${issue.row ?? ""}|${issue.column ?? ""}|${issue.para}|${issue.start}|${issue.word.toLowerCase()}`;

export function useSlideSpellCheck(deps: SlideSpellCheckDeps) {
  const { preferences } = usePreferences();
  const schedule = deps.schedule ?? ((callback: () => void): IdleHandle => scheduleIdleWork(callback, { timeoutMs: 150, fallbackBudgetMs: 8, maxBudgetMs: 16 }));

  const issues = shallowRef<SlideSpellIssue[]>([]);
  const underlineRects = shallowRef<SpellUnderlineRect[]>([]);
  const contextMenu = ref<SlideSpellMenu | null>(null);
  const dictionaryOpen = ref(false);
  const customDictionaryList = shallowRef<CustomDictionaryEntry[]>([...customDictionaryEntries()]);

  const ignoredOccurrences = new Set<string>();
  const ignoredWords = new Set<string>();
  const dictionaries = new Map<string, Dictionary>();
  const deckLanguageTag = (): string => deps.model.value?.language ?? DEFAULT_LANGUAGE;

  function requestDictionary(lang: string): void {
    const tag = normalizeLanguageTag(lang);
    if (dictionaries.has(tag)) return;
    void loadDictionary(tag)?.then((dictionary) => dictionaries.set(tag, dictionary));
  }

  // --- checking ----------------------------------------------------------------
  let handle: IdleHandle | null = null;
  let token = 0;

  /** The current slide's checkable runs and notes lines, one worker item each. */
  function collect(model: PresentationModel, slide: Slide): { items: SpellCheckItem[]; places: Omit<SlideSpellIssue, "start" | "end" | "word">[]; offsets: number[] } {
    const items: SpellCheckItem[] = [];
    const places: Omit<SlideSpellIssue, "start" | "end" | "word">[] = [];
    const offsets: number[] = [];
    const add = (text: string, lang: string, at: Omit<StageTextRange, "start" | "end"> & { notes?: true }, offset: number): void => {
      if (!text || !isDictionaryAvailable(lang)) return;
      requestDictionary(lang);
      items.push({ key: String(items.length), lang: normalizeLanguageTag(lang), text });
      places.push({ ...at, lang: normalizeLanguageTag(lang) });
      offsets.push(offset);
    };
    const theme = deckThemeOf(model);
    for (const { at, paragraph } of slideParagraphs(slide)) {
      let offset = 0;
      for (const run of paragraph.runs) {
        const code = run.fontFamily !== undefined && isCodeStyledFont(resolveFont(run.fontFamily, theme));
        if (!run.noProof && !code) add(run.text, runLanguage(run, model), at, offset);
        offset += run.text.length;
      }
    }
    deps.notesText.value.split("\n").forEach((line, para) => add(line, deckLanguageTag(), { notes: true, para }, 0));
    return { items, places, offsets };
  }

  async function check(): Promise<void> {
    handle = null;
    const model = deps.model.value;
    const slide = deps.currentSlide.value;
    const mine = ++token;
    if (!model || !slide || !preferences.spellCheck) {
      issues.value = [];
      measure();
      return;
    }
    const { items, places, offsets } = collect(model, slide);
    const results = await checkTextsInWorker(items);
    if (mine !== token) return;
    const found: SlideSpellIssue[] = [];
    for (const result of results) {
      const index = Number(result.key);
      for (const match of result.matches) {
        const issue: SlideSpellIssue = { ...places[index], start: offsets[index] + match.start, end: offsets[index] + match.end, word: match.word };
        if (ignoredWords.has(issue.word.toLowerCase()) || ignoredOccurrences.has(occurrenceKey(slide, issue)) || hasCustomWord(issue.word, issue.lang)) continue;
        found.push(issue);
      }
    }
    issues.value = found;
    remeasure();
  }

  /** Re-checks the current slide on idle (coalescing repeated requests). */
  function scheduleCheck(): void {
    token += 1; // any in-flight answer is now stale
    if (!handle) handle = schedule(() => void check());
  }

  // --- the stage underlines --------------------------------------------------------
  function measure(): void {
    const surface = deps.surface();
    const rects: SpellUnderlineRect[] = [];
    if (surface) {
      issues.value.forEach((issue, index) => {
        if (issue.notes) return;
        for (const rect of measureStageRange(surface, issue)) rects.push({ ...rect, issue: index, word: issue.word });
      });
    }
    underlineRects.value = rects;
  }
  /** Re-measures after the next render (a slide change, a stage rescale). */
  const remeasure = (): void => void nextTick(measure);

  watch([deps.currentSlide, deps.notesText, () => deps.model.value?.language, () => preferences.spellCheck], scheduleCheck, { immediate: true });
  watch([() => deps.editing.editingFrameId.value, () => deps.editing.editingCell.value], remeasure);
  // Another deck in this view is a new session: Ignore / Ignore all start over.
  watch(() => deps.model.value?.id, (id, previous) => void (previous !== undefined && id !== previous && reset()));
  const unsubscribe = subscribeCustomDictionary(() => {
    customDictionaryList.value = [...customDictionaryEntries()];
    scheduleCheck();
  });

  // --- the suggestions menu ------------------------------------------------------------
  /** Opens the menu for underline `index` (SlideProofingLayer's spell-context). */
  function openContextMenuAt(index: number, clientX: number, clientY: number): void {
    const issue = issues.value[index];
    if (!issue) return;
    const dictionary = dictionaries.get(normalizeLanguageTag(issue.lang));
    contextMenu.value = { left: clientX, top: clientY, word: issue.word, suggestions: dictionary ? getSuggestions(issue.word, dictionary) : [], issue };
  }
  const closeContextMenu = (): void => void (contextMenu.value = null);

  function applySuggestion(suggestion: string): void {
    const menu = contextMenu.value;
    contextMenu.value = null;
    const model = deps.model.value;
    const slide = deps.currentSlide.value;
    if (!menu || !model || !slide) return;
    const { issue } = menu;
    const shaped = applyCasePattern(suggestion, detectCasePattern(issue.word));
    const edit = deps.editing;
    const inSession = edit.editingFrameId.value === issue.frameId && (edit.editingCell.value?.row ?? undefined) === issue.row && (edit.editingCell.value?.col ?? undefined) === issue.column;
    const range = { start: { para: issue.para, offset: issue.start }, end: { para: issue.para, offset: issue.end } };
    const stillThere = (paragraphs: SlideParagraph[] | null): boolean => {
      const paragraph = paragraphs?.[issue.para];
      return !!paragraph && paragraphText(paragraph).slice(issue.start, issue.end) === issue.word;
    };
    if (inSession) {
      // The session's own one-entry command keeps it open and re-renders it.
      edit.runTextCommand(SPELL_UNDO_LABELS.correct, (owner) => {
        if (stillThere(owner.paragraphs)) replaceParagraphText(owner.paragraphs[issue.para], issue.start, issue.end, shaped);
      }, range);
      return;
    }
    if (edit.editingFrameId.value) edit.stopEditing();
    const current = deps.currentSlide.value;
    const paragraphs = current ? rangeParagraphs(current, issue) : null;
    if (!current || !stillThere(paragraphs)) {
      scheduleCheck();
      return;
    }
    deps.record(SPELL_UNDO_LABELS.correct);
    replaceParagraphText(paragraphs![issue.para], issue.start, issue.end, shaped);
    markSlideEdited(current, model.format);
    deps.commit();
  }

  function ignore(): void {
    const menu = contextMenu.value;
    const slide = deps.currentSlide.value;
    contextMenu.value = null;
    if (menu && slide) ignoredOccurrences.add(occurrenceKey(slide, menu.issue));
    scheduleCheck();
  }

  function ignoreAll(): void {
    const menu = contextMenu.value;
    contextMenu.value = null;
    if (menu) ignoredWords.add(menu.word.toLowerCase());
    scheduleCheck();
  }

  function addToDictionary(): void {
    const menu = contextMenu.value;
    contextMenu.value = null;
    if (menu) addCustomWord(menu.word, menu.issue.lang); // the subscription re-checks
  }

  // --- the language control -------------------------------------------------------------
  const deckLanguage = computed(() => deckLanguageTag());
  const languageOptions = SUPPORTED_LANGUAGES;
  /** Text is selected: a session, or a selected text frame (whole frame). */
  const selectionActive = computed(() => deps.editing.toolbarEnabled.value);
  /** The selection's language tag: its runs' own tag, the deck's where they
   * inherit, "mixed" across several. */
  const selectionLanguage = computed(() => {
    const state = deps.editing.toolbarState.value;
    if (!selectionActive.value || !state) return deckLanguage.value;
    return state.lang === "mixed" ? "mixed" : state.lang || deckLanguage.value;
  });
  const selectionNoProof = computed(() => selectionActive.value && deps.editing.toolbarState.value?.noProof === true);
  const currentLanguageLabel = computed(() => (selectionLanguage.value === "mixed" ? "Multiple languages" : languageLabel(selectionLanguage.value)));

  function setDeckLanguage(tag: string): void {
    const model = deps.model.value;
    if (!model || model.language === tag) return;
    deps.record(SPELL_UNDO_LABELS.language);
    model.language = tag;
    deps.commit();
  }
  function setSelectionLanguage(tag: string): void {
    deps.editing.runTextCommand(SPELL_UNDO_LABELS.language, (owner, range) => void applyRunFormat(owner, range, { lang: tag }));
  }
  /** The control's pick: the selection's language, else the deck's. */
  const setLanguage = (tag: string): void => (selectionActive.value ? setSelectionLanguage(tag) : setDeckLanguage(tag));
  function setNoProof(on: boolean): void {
    deps.editing.runTextCommand(SPELL_UNDO_LABELS.noProof, (owner, range) => void applyRunFormat(owner, range, { noProof: on }));
  }

  const openDictionary = (): void => void (dictionaryOpen.value = true);
  const closeDictionary = (): void => void (dictionaryOpen.value = false);

  /** Forgets the session's ignores and the current results (a new deck). */
  function reset(): void {
    ignoredOccurrences.clear();
    ignoredWords.clear();
    issues.value = [];
    underlineRects.value = [];
    contextMenu.value = null;
    scheduleCheck();
  }

  onScopeDispose(() => {
    token += 1;
    handle?.cancel();
    unsubscribe();
  });

  return {
    issues,
    underlineRects,
    contextMenu,
    deckLanguage,
    languageOptions,
    selectionActive,
    selectionLanguage,
    selectionNoProof,
    currentLanguageLabel,
    customDictionaryList,
    openContextMenuAt,
    closeContextMenu,
    applySuggestion,
    ignore,
    ignoreAll,
    addToDictionary,
    removeFromDictionary: (word: string, lang: string): void => removeCustomWord(word, lang),
    addDictionaryWord: (word: string, lang: string): void => void addCustomWord(word, lang),
    setDeckLanguage,
    setSelectionLanguage,
    setLanguage,
    setNoProof,
    dictionaryOpen,
    openDictionary,
    closeDictionary,
    remeasure,
    reset,
  };
}

export type SlideSpellCheck = ReturnType<typeof useSlideSpellCheck>;

/** Narrows a component's runtime-typed controller prop (the .vue parser bars casts). */
export function asSlideSpellCheck(value: unknown): SlideSpellCheck {
  return value as SlideSpellCheck;
}
