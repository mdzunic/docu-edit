// @vitest-environment jsdom
// SPEC-098 — spell check on slides: only the current slide's frames, cells and
// notes go to the worker, each run in its effective language (noProof runs and
// languages without a dictionary skipped); the suggestions menu, a correction
// that keeps the word's case as ONE "Correct spelling" entry, Ignore / Ignore
// all / Add to dictionary; and the language control's deck-vs-selection rule.
import { computed, effectScope, nextTick, ref, shallowRef, triggerRef } from "vue";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { IdleHandle } from "../../lib/core/idleScheduler";
import type { PresentationModel, Slide, SlideFrame, SlideParagraph } from "../../lib/presentation/model";
import type { SpellCheckItem } from "../../lib/workers/spellCheckClient";
import { removeCustomWord } from "../../lib/customDictionary";
import { useFrameSelection } from "./useFrameSelection";
import type { SlideTextEditing } from "./useSlideTextEditing";
import { useSlideSpellCheck } from "./useSlideSpellCheck";

const { WORDS, sent } = vi.hoisted(() => ({
  WORDS: new Set(["the", "a", "slide", "world", "hello", "guten", "tag", "colour", "notes", "speaker", "tea"]),
  sent: [] as SpellCheckItem[][],
}));

vi.mock("../../lib/spellcheck/dictionaries", () => ({
  loadDictionary: (tag: string) =>
    ["en-US", "en-GB", "de-DE"].includes(tag) ? Promise.resolve({ has: (word: string) => WORDS.has(word.toLowerCase()), words: [...WORDS] }) : null,
}));

vi.mock("../../lib/workers/spellCheckClient", async () => {
  const { checkText } = await vi.importActual<typeof import("../../lib/spellcheck/checker")>("../../lib/spellcheck/checker");
  const dictionary = { has: (word: string) => WORDS.has(word.toLowerCase()), words: [...WORDS] };
  return {
    checkTextsInWorker: async (items: SpellCheckItem[]) => {
      sent.push(items);
      return items.map((item) => ({ key: item.key, matches: checkText(item.text, dictionary) }));
    },
  };
});

const box = { xCm: 0, yCm: 0, widthCm: 10, heightCm: 2 };
const para = (...runs: SlideParagraph["runs"]): SlideParagraph => ({ runs });

function deckWith(slides: Slide[], language?: string): PresentationModel {
  return { kind: "presentation", id: "d", title: "Deck", format: "pptx", slides, slideWidthCm: 33.867, slideHeightCm: 19.05, ...(language ? { language } : {}) };
}

function setup(model: PresentationModel, opts: { selectionActive?: boolean; editingFrameId?: string } = {}) {
  const deck = shallowRef<PresentationModel | null>(model);
  const currentIndex = ref(0);
  const currentSlide = computed(() => deck.value?.slides[currentIndex.value] ?? null);
  const calls: string[] = [];
  const commands: { label: string; range?: unknown }[] = [];
  const selection = useFrameSelection({ currentSlide });
  const editing = {
    editingFrameId: ref<string | null>(opts.editingFrameId ?? null),
    editingCell: ref(null),
    toolbarEnabled: computed(() => !!opts.selectionActive),
    toolbarState: ref(opts.selectionActive ? { lang: "", noProof: false } : null),
    stopEditing: () => calls.push("stopEditing"),
    runTextCommand: (label: string, apply: (owner: { paragraphs: SlideParagraph[] }, range: unknown, frame: SlideFrame) => void, range?: unknown) => {
      commands.push({ label, range });
      const frame = currentSlide.value!.frames[0];
      if (frame.kind === "text") apply(frame, range ?? { start: { para: 0, offset: 0 }, end: { para: 0, offset: 99 } }, frame);
      return true;
    },
  } as unknown as SlideTextEditing;
  const notesText = computed(() => (currentSlide.value?.notes ?? []).join("\n"));
  const queue: (() => void)[] = [];
  const schedule = (callback: () => void): IdleHandle => {
    queue.push(callback);
    return { cancel: () => void queue.splice(queue.indexOf(callback), 1) };
  };
  const scope = effectScope();
  const spell = scope.run(() =>
    useSlideSpellCheck({
      model: deck,
      currentSlide,
      currentIndex,
      record: (label) => calls.push(`record:${label}`),
      // The view's commit re-triggers the deck's shallowRef (useSlideRender).
      commit: () => {
        calls.push("commit");
        triggerRef(deck);
      },
      selection,
      editing,
      notesText,
      surface: () => null,
      schedule,
    }),
  )!;
  /** Runs the scheduled idle check and lets the (mocked) worker answer. */
  async function flush(): Promise<void> {
    await nextTick();
    while (queue.length > 0) queue.shift()!();
    for (let i = 0; i < 5; i++) await Promise.resolve();
    await nextTick();
  }
  return { spell, deck, currentIndex, calls, commands, flush, scope };
}

afterEach(() => {
  sent.length = 0;
  removeCustomWord("wrld", "en-US");
});

describe("useSlideSpellCheck — what is checked", () => {
  it("checks only the current slide's frames, shape text, table cells and notes, each run in its effective language", async () => {
    const model = deckWith([
      {
        id: "s1",
        frames: [
          { kind: "text", id: "t", ...box, paragraphs: [para({ text: "hello " }, { text: "guten tag", lang: "de-DE" }, { text: " wrld", noProof: true })] },
          { kind: "shape", id: "sh", ...box, geometry: "rect", paragraphs: [para({ text: "colour", lang: "en-GB" }, { text: " olá", lang: "pt-BR" })] },
          { kind: "table", id: "tb", ...box, columnWidthsCm: [2], rows: [{ cells: [{ paragraphs: [para({ text: "tea" })] }] }] },
          { kind: "group", id: "g", ...box, children: [{ kind: "text", id: "c", ...box, paragraphs: [para({ text: "grouped" })] }] },
        ],
        notes: ["speaker notes"],
      },
      { id: "s2", frames: [{ kind: "text", id: "other", ...box, paragraphs: [para({ text: "elsewhere" })] }] },
    ]);
    const { flush } = setup(model);
    await flush();
    expect(sent.at(-1)!.map(({ lang, text }) => [lang, text])).toEqual([
      ["en-US", "hello "],
      ["de-DE", "guten tag"],
      ["en-GB", "colour"],
      ["en-US", "tea"],
      ["en-US", "speaker notes"],
    ]);
  });

  it("falls back to the deck language for runs without their own, and reports issues at paragraph offsets", async () => {
    const model = deckWith([{ id: "s", frames: [{ kind: "text", id: "t", ...box, paragraphs: [para({ text: "Hello " }, { text: "wrld teh", bold: true })] }] }], "en-GB");
    const { spell, flush } = setup(model);
    await flush();
    expect(sent.at(-1)!.every((item) => item.lang === "en-GB")).toBe(true);
    expect(spell.issues.value.map(({ frameId, para: p, start, end, word }) => [frameId, p, start, end, word])).toEqual([
      ["t", 0, 6, 10, "wrld"],
      ["t", 0, 11, 14, "teh"],
    ]);
  });

  it("re-checks when the slide changes, checking the new slide only", async () => {
    const model = deckWith([
      { id: "a", frames: [{ kind: "text", id: "t", ...box, paragraphs: [para({ text: "first" })] }] },
      { id: "b", frames: [{ kind: "text", id: "u", ...box, paragraphs: [para({ text: "second" })] }] },
    ]);
    const { currentIndex, flush } = setup(model);
    await flush();
    currentIndex.value = 1;
    await flush();
    expect(sent.at(-1)!.map((item) => item.text)).toEqual(["second"]);
  });
});

describe("useSlideSpellCheck — the suggestions menu", () => {
  function teh(): PresentationModel {
    return deckWith([{ id: "s", frames: [{ kind: "text", id: "t", ...box, paragraphs: [para({ text: "Teh slide " }, { text: "teh", italic: true })] }] }]);
  }

  it("offers ≤ 5 suggestions and applies one keeping the word's case, as ONE \"Correct spelling\" entry", async () => {
    const model = teh();
    const { spell, calls, flush } = setup(model);
    await flush();
    await flush(); // the dictionary for suggestions resolves
    spell.openContextMenuAt(0, 40, 50);
    expect(spell.contextMenu.value).toMatchObject({ left: 40, top: 50, word: "Teh" });
    expect(spell.contextMenu.value!.suggestions.length).toBeLessThanOrEqual(5);
    expect(spell.contextMenu.value!.suggestions).toContain("the");
    spell.applySuggestion("the");
    expect(calls).toEqual(["record:Correct spelling", "commit"]);
    const frame = model.slides[0].frames[0];
    expect(frame.kind === "text" && frame.paragraphs[0].runs).toEqual([{ text: "The slide " }, { text: "teh", italic: true }]);
    expect(model.slides[0].dirty).toBe(true);
    expect(spell.contextMenu.value).toBeNull();
  });

  it("corrects inside a running session through its own one-entry command", async () => {
    const model = teh();
    const { spell, commands, calls, flush } = setup(model, { editingFrameId: "t" });
    await flush();
    spell.openContextMenuAt(1, 0, 0);
    spell.applySuggestion("the");
    expect(commands).toEqual([{ label: "Correct spelling", range: { start: { para: 0, offset: 10 }, end: { para: 0, offset: 13 } } }]);
    expect(calls).not.toContain("stopEditing");
    const frame = model.slides[0].frames[0];
    expect(frame.kind === "text" && frame.paragraphs[0].runs.map((run) => run.text).join("")).toBe("Teh slide the");
  });

  it("Ignore clears one occurrence, Ignore all every one, Add to dictionary persists the word", async () => {
    const model = deckWith([{ id: "s", frames: [{ kind: "text", id: "t", ...box, paragraphs: [para({ text: "teh teh wrld" })] }] }]);
    const { spell, flush } = setup(model);
    await flush();
    expect(spell.issues.value.map((issue) => issue.start)).toEqual([0, 4, 8]);
    spell.openContextMenuAt(0, 0, 0);
    spell.ignore();
    await flush();
    expect(spell.issues.value.map((issue) => issue.start)).toEqual([4, 8]);
    spell.openContextMenuAt(0, 0, 0);
    spell.ignoreAll();
    await flush();
    expect(spell.issues.value.map((issue) => issue.word)).toEqual(["wrld"]);
    spell.openContextMenuAt(0, 0, 0);
    spell.addToDictionary();
    await flush();
    expect(spell.issues.value).toEqual([]);
    expect(spell.customDictionaryList.value).toContainEqual({ word: "wrld", lang: "en-US" });
  });
});

describe("useSlideSpellCheck — the language control", () => {
  const plain = () => deckWith([{ id: "s", frames: [{ kind: "text", id: "t", ...box, paragraphs: [para({ text: "hello" })] }] }]);

  it("with nothing selected it sets the deck language, as one \"Language\" entry", () => {
    const model = plain();
    const { spell, calls } = setup(model);
    expect([spell.deckLanguage.value, spell.currentLanguageLabel.value, spell.selectionActive.value]).toEqual(["en-US", "English (US)", false]);
    spell.setLanguage("de-DE");
    expect(model.language).toBe("de-DE");
    expect(calls).toEqual(["record:Language", "commit"]);
    expect(spell.currentLanguageLabel.value).toBe("German");
  });

  it("with text selected it sets the selection's run lang, and Do not check spelling sets noProof", () => {
    const model = plain();
    const { spell, commands } = setup(model, { selectionActive: true });
    spell.setLanguage("en-GB");
    spell.setNoProof(true);
    expect(commands.map((command) => command.label)).toEqual(["Language", "Do not check spelling"]);
    const frame = model.slides[0].frames[0];
    expect(frame.kind === "text" && frame.paragraphs[0].runs).toEqual([{ text: "hello", lang: "en-GB", noProof: true }]);
    expect(model.language).toBeUndefined();
  });

  it("opens and closes the custom dictionary dialog", () => {
    const { spell } = setup(plain());
    spell.openDictionary();
    expect(spell.dictionaryOpen.value).toBe(true);
    spell.closeDictionary();
    expect(spell.dictionaryOpen.value).toBe(false);
  });
});
