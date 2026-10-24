// @vitest-environment jsdom
// SPEC-098 — the deck search composable: the debounced walk and its clock,
// the idle-sliced walk over a big deck (and its measured budgets), options and
// the invalid-regex state, navigation into the stage hooks with the wrap
// hints, Replace / Replace all as single undo entries, the keys and the stage
// highlights.
import { effectScope, nextTick, ref, shallowRef } from "vue";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { INVALID_REGEX_MESSAGE } from "../../lib/core/findQuery";
import type { IdleHandle } from "../../lib/core/idleScheduler";
import type { PresentationModel, Slide, SlideFrame } from "../../lib/presentation/model";
import type { DeckHit } from "../../lib/presentation/deckSearch";
import { FIND_DEBOUNCE_MS, useDeckSearch, type DeckSearchDeps } from "./useDeckSearch";
import { WALK_BATCH_SLIDES } from "./deckSearchWalk";

const box = { xCm: 0, yCm: 0, widthCm: 10, heightCm: 2 };
const text = (id: string, ...lines: string[]): SlideFrame => ({ kind: "text", id, ...box, paragraphs: lines.map((line) => ({ runs: [{ text: line }] })) });
const slide = (id: string, frames: SlideFrame[], notes?: string[]): Slide => ({ id, frames, ...(notes ? { notes } : {}) });

function deckOf(slides: Slide[]): PresentationModel {
  return { kind: "presentation", id: "d", title: "Deck", format: "pptx", slides, slideWidthCm: 33.867, slideHeightCm: 19.05 };
}

/** The fixture's shape: slide 1 notes "Open with…", slide 2 body "Open Berlin office". */
const MINI = (): PresentationModel =>
  deckOf([
    slide("s1", [text("title1", "Quarterly Review")], ["Open with the Q3 highlight"]),
    slide("s2", [text("title2", "Roadmap"), text("body2", "Hire two engineers", "Open Berlin office")]),
    slide("s3", [text("t3", "Thank you")]),
  ]);

function setup(model: PresentationModel, extra: Partial<DeckSearchDeps> = {}) {
  const deck = shallowRef<PresentationModel | null>(model);
  const currentIndex = ref(0);
  const calls: string[] = [];
  const deps: DeckSearchDeps = {
    model: deck,
    currentIndex,
    record: (label) => calls.push(`record:${label}`),
    commit: () => calls.push("commit"),
    selectFrame: (id) => calls.push(`select:${id}`),
    startEditing: (hit: DeckHit) => calls.push(`edit:${hit.frameId}:${hit.para}:${hit.start}-${hit.end}`),
    openNotes: (hit: DeckHit) => void calls.push(`notes:${hit.para}:${hit.start}-${hit.end}`),
    stopEditing: () => calls.push("stop"),
    editingFrameId: () => null,
    surface: () => null,
    focusStage: () => calls.push("focusStage"),
    ...extra,
  };
  const scope = effectScope();
  const search = scope.run(() => useDeckSearch(deps))!;
  return { search, deck, currentIndex, calls, scope };
}

/** Lets activate()'s nextTicks and its macrotask settle. */
async function settle(): Promise<void> {
  await vi.runAllTimersAsync();
  await nextTick();
}

describe("useDeckSearch — the walk", () => {
  beforeEach(() => void vi.useFakeTimers());
  afterEach(() => void vi.useRealTimers());

  it("debounces the query FIND_DEBOUNCE_MS and starts the clock when the debounce fires", async () => {
    let clock = 0;
    const { search } = setup(MINI(), { now: () => clock });
    search.openFind();
    search.query.value = "Open";
    await nextTick();
    clock = 10;
    vi.advanceTimersByTime(FIND_DEBOUNCE_MS - 1);
    expect(search.hitCount.value).toBe(0);
    clock = 160;
    vi.advanceTimersByTime(1);
    expect(search.hitCount.value).toBe(2);
    expect(search.walkTiming.value).toEqual({ startedAt: 160, firstHitsAt: 160, settledAt: 160 });
  });

  it("groups hits by slide with snippet, frame name and scope tag", async () => {
    const { search } = setup(MINI());
    search.openFind();
    search.query.value = "Open";
    await nextTick();
    vi.advanceTimersByTime(FIND_DEBOUNCE_MS);
    expect(search.groups.value).toEqual([
      { slideIndex: 0, slideNumber: 1, hits: [{ index: 0, before: "", match: "Open", after: " with the Q3 highlight", frameName: "Speaker notes", scopeTag: "Notes" }] },
      { slideIndex: 1, slideNumber: 2, hits: [{ index: 1, before: "", match: "Open", after: " Berlin office", frameName: "Text box", scopeTag: "" }] },
    ]);
    expect(search.bar.value).toMatchObject({ matchCount: 2, currentIndex: -1, scopeLabel: "Slides, shapes, tables and notes" });
    expect(search.bar.value.scopeTitle).toContain("Not searched: master and layout text.");
  });

  it("shows INVALID_REGEX_MESSAGE and runs nothing for an invalid expression; options persist across close", async () => {
    const { search } = setup(MINI());
    search.openFind();
    search.useRegex.value = true;
    search.query.value = "(";
    await nextTick();
    vi.advanceTimersByTime(FIND_DEBOUNCE_MS);
    expect(search.error.value).toBe(INVALID_REGEX_MESSAGE);
    expect(search.bar.value.error).toBe(INVALID_REGEX_MESSAGE);
    expect(search.hits.value).toEqual([]);
    search.closeFind();
    search.openFind();
    expect(search.useRegex.value).toBe(true);
    expect(search.query.value).toBe("(");
  });

  it("walks a deck over 40 slides WALK_BATCH_SLIDES at a time through the idle scheduler, with progress until settled", async () => {
    const slides = Array.from({ length: 60 }, (_, i) => slide(`s${i}`, [text(`t${i}`, `needle ${i}`)]));
    const queue: ((budget: number) => void)[] = [];
    const schedule = (callback: (budget: number) => void): IdleHandle => {
      queue.push(callback);
      return { cancel: () => void queue.splice(queue.indexOf(callback), 1) };
    };
    let clock = 0;
    const { search } = setup(deckOf(slides), { schedule, now: () => clock });
    search.query.value = "needle";
    search.openFind();
    expect(search.indexing.value).toBe(true);
    expect(search.hitCount.value).toBe(0);
    // A zero-budget slot still walks one whole batch, appending its hits.
    queue.shift()!(0);
    expect(search.hitCount.value).toBe(WALK_BATCH_SLIDES);
    expect(search.groups.value[0].slideNumber).toBe(1);
    expect(search.indexing.value).toBe(true);
    while (queue.length > 0) queue.shift()!(0);
    expect(search.hitCount.value).toBe(60);
    expect(search.indexing.value).toBe(false);
    expect(search.hint.value).toBe("");
  });
});

/** Real timers and a real macrotask scheduler (jsdom has no idle callback). */
function macrotaskSlots(callback: (budget: number) => void): IdleHandle {
  const timer = setTimeout(() => callback(12), 0);
  return { cancel: () => clearTimeout(timer) };
}

describe("useDeckSearch — measured budgets (SPEC-098 performance rules)", () => {
  it("first hits within 150 ms and settled within 500 ms of the debounce firing, on a 60-slide deck", async () => {
    // A heavy deck: 60 slides × 6 frames × 6 paragraphs of prose, plus notes.
    const prose = "The quarterly results were strong across every region and the client was pleased with the outcome";
    const slides = Array.from({ length: 60 }, (_, i) =>
      slide(
        `s${i}`,
        Array.from({ length: 6 }, (_frame, f) => text(`t${i}-${f}`, ...Array.from({ length: 6 }, (_para, p) => `${prose} ${i}.${f}.${p}`))),
        [`${prose} notes ${i}`],
      ),
    );
    const { search } = setup(deckOf(slides), { schedule: macrotaskSlots });
    search.openFind();
    search.query.value = "client";
    await nextTick();
    await vi.waitFor(() => expect(search.indexing.value === false && search.hitCount.value > 0).toBe(true), { timeout: 2000 });
    const timing = search.walkTiming.value!;
    expect(search.hitCount.value).toBe(60 * 37);
    expect(timing.firstHitsAt! - timing.startedAt).toBeLessThan(150);
    expect(timing.settledAt! - timing.startedAt).toBeLessThan(500);
  });
});

describe("useDeckSearch — navigation, replace and keys", () => {
  beforeEach(() => void vi.useFakeTimers());
  afterEach(() => void vi.useRealTimers());

  async function searched(model = MINI(), query = "Open") {
    const ctx = setup(model);
    ctx.search.openFind();
    ctx.search.query.value = query;
    await nextTick();
    vi.advanceTimersByTime(FIND_DEBOUNCE_MS);
    return ctx;
  }

  it("activating a hit sets the slide, selects the frame and opens a session on the range — no mutation", async () => {
    const { search, currentIndex, calls } = await searched();
    search.goToHit(1);
    await settle();
    expect(currentIndex.value).toBe(1);
    expect(search.currentHitIndex.value).toBe(1);
    expect(calls).toEqual(["stop", "select:body2", "edit:body2:1:0-4"]);
  });

  it("a notes hit opens the notes instead", async () => {
    const { search, calls } = await searched();
    search.goToHit(0);
    await settle();
    expect(calls).toEqual(["stop", "notes:0:0-4"]);
  });

  it("steps from the current slide, wraps with the hints, and finishes a pending walk first", async () => {
    const { search, currentIndex } = await searched();
    currentIndex.value = 1;
    search.next();
    await settle();
    expect(search.currentHitIndex.value).toBe(1);
    search.next();
    await settle();
    expect([search.currentHitIndex.value, search.hint.value]).toEqual([0, "Wrapped to top"]);
    search.prev();
    await settle();
    expect([search.currentHitIndex.value, search.hint.value]).toEqual([1, "Wrapped to bottom"]);
    // Typing then stepping at once: the step walks the new query itself.
    search.query.value = "Roadmap";
    await nextTick();
    search.next();
    await settle();
    expect(search.hitCount.value).toBe(1);
    expect(search.currentHitIndex.value).toBe(0);
  });

  it("Replace rewrites the current hit as one undo entry and steps to the next", async () => {
    const model = deckOf([slide("a", [text("t", "cat and cat")]), slide("b", [text("u", "cat")])]);
    const { search, calls } = await searched(model, "cat");
    search.goToHit(0);
    await settle();
    calls.length = 0;
    search.replaceText.value = "dog";
    search.replaceCurrent();
    await settle();
    expect(calls.filter((call) => call.startsWith("record") || call === "commit")).toEqual(["record:Replace", "commit"]);
    expect(model.slides[0].frames[0].kind === "text" && model.slides[0].frames[0].paragraphs[0].runs[0].text).toBe("dog and cat");
    expect(search.hitCount.value).toBe(2);
    expect(search.currentHitIndex.value).toBe(0); // the next "cat", now first
    expect(model.slides[0].dirty).toBe(true);
  });

  it("Replace all is ONE undo entry — one record, every mutation, one commit — and reports the count", async () => {
    const model = MINI();
    const { search, calls } = await searched(model);
    calls.length = 0;
    search.replaceText.value = "Launch";
    search.replaceAll();
    expect(calls).toEqual(["stop", "record:Replace all", "commit"]);
    expect(search.hint.value).toBe("Replaced 2 occurrences");
    expect(model.slides[0].notes).toEqual(["Launch with the Q3 highlight"]);
    expect(model.slides[0].notesDirty).toBe(true);
    expect(model.slides[1].dirty).toBe(true);
    expect([model.slides[2].dirty, model.slides[2].notesDirty]).toEqual([undefined, undefined]);
    expect(search.hitCount.value).toBe(0);
  });

  it("Cmd/Ctrl+F opens the bar, Cmd/Ctrl+H with the replace row; never inside a dialog", () => {
    const { search } = setup(MINI());
    const press = (key: string, target?: EventTarget) => {
      const event = new KeyboardEvent("keydown", { key, ctrlKey: true, metaKey: true, cancelable: true });
      if (target) Object.defineProperty(event, "target", { value: target });
      search.onGlobalKeydown(event);
      return event;
    };
    const dialog = document.createElement("div");
    dialog.setAttribute("role", "dialog");
    const inside = document.createElement("input");
    dialog.appendChild(inside);
    expect(press("f", inside).defaultPrevented).toBe(false);
    expect(search.open.value).toBe(false);
    expect(press("f").defaultPrevented).toBe(true);
    expect([search.open.value, search.expanded.value]).toEqual([true, false]);
    press("h");
    expect(search.expanded.value).toBe(true);
  });

  it("closing clears the hits and hands focus to the stage; the toggle opens and closes", async () => {
    const { search, calls } = await searched();
    search.toggle();
    expect([search.open.value, search.hitCount.value]).toEqual([false, 0]);
    expect(calls).toContain("focusStage");
    search.toggle();
    expect(search.open.value).toBe(true);
    expect(search.hitCount.value).toBe(2);
  });
});

describe("useDeckSearch — stage highlights", () => {
  // jsdom's Range has no layout; a stub stands in for the browser's rects.
  afterEach(() => void Reflect.deleteProperty(Range.prototype, "getClientRects"));

  it("paints the current slide's frame hits (never notes), with .current on the current one", async () => {
    const surface = document.createElement("div");
    surface.innerHTML =
      `<div data-frame-id="title2"><div class="paragraph"><span>Roadmap</span></div></div>` +
      `<div data-frame-id="body2"><div class="paragraph"><span class="bullet-marker">•</span><span>Hire two engineers</span></div>` +
      `<div class="paragraph"><span class="bullet-marker">•</span><span>Open </span><span>Berlin office</span></div></div>`;
    document.body.appendChild(surface);
    const ranges: string[] = [];
    Object.defineProperty(Range.prototype, "getClientRects", {
      configurable: true,
      value(this: Range) {
        ranges.push(this.toString());
        return [{ left: 10, top: 20, width: 30, height: 10 }];
      },
    });
    const { search, currentIndex } = setup(MINI(), { surface: () => surface });
    search.query.value = "Open";
    search.openFind();
    currentIndex.value = 1;
    await nextTick();
    await nextTick();
    expect(ranges).toEqual(["Open"]);
    expect(search.highlightRects.value).toEqual([{ left: 10, top: 20, width: 30, height: 10, current: false }]);
    search.currentHitIndex.value = 1;
    await nextTick();
    await nextTick();
    expect(search.highlightRects.value[0].current).toBe(true);
    currentIndex.value = 0;
    await nextTick();
    await nextTick();
    expect(search.highlightRects.value).toEqual([]); // slide 1's only hit is in its notes
  });
});
