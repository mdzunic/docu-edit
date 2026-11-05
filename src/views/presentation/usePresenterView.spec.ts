// SPEC-100 — the presenter window (PresenterView over usePresenterView),
// mounted on the real route table with the idb seam mocked so the library runs
// in memory. window.open is stubbed (jsdom has no second window); the channel
// is Node's real BroadcastChannel, so the last test can pair this window with
// a mounted audience play view (?follow=) and watch the two stay in step.
// Covers the load guards, the surface (current/next/notes/counters/builds),
// keys and blanking, the grid, the notes scale and timer wiring, popup-blocked
// and no-transport single-window mode, a lost audience, the exit handoff, and
// that presenting writes nothing (download parts byte-identical).
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defineComponent, nextTick } from "vue";
import { flushPromises, mount, type VueWrapper } from "@vue/test-utils";
import { createMemoryHistory, createRouter, type Router } from "vue-router";
import { routes } from "../../router";
import type { PresentationModel } from "../../lib/presentation/model";

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  vi.doMock("../../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    getRecord: vi.fn().mockResolvedValue(undefined),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
    loadAllTemplateRecords: vi.fn().mockResolvedValue([]),
    putTemplateRecord: vi.fn().mockResolvedValue(undefined),
    deleteTemplateRecord: vi.fn().mockResolvedValue(undefined),
  }));
});

// Navigations fired and forgotten (exit/redirect) lazy-load other route
// components; await them before teardown (see PresentationPlayView.spec.ts).
// The redirect to /doc/:id lazy-loads the whole editor, whose first transform
// on a cold worker under the full suite runs past Vitest's 10 s default hook
// timeout — so the hook gets the room vite.config.ts gives every test (a
// timeout is a hang detector, not a performance budget).
const pendingNavigations: Promise<unknown>[] = [];
const mounted: VueWrapper[] = [];

afterEach(async () => {
  for (const wrapper of mounted.splice(0)) wrapper.unmount();
  await Promise.allSettled(pendingNavigations);
  pendingNavigations.length = 0;
  vi.doUnmock("../../lib/idb");
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
  vi.useRealTimers();
}, 30_000);

/** A stand-in for the audience window's WindowProxy. */
function fakeWindow() {
  const win = { closed: false, close: vi.fn(() => (win.closed = true)) };
  return win;
}

async function miniDeck(): Promise<{ deck: PresentationModel; bytes: Uint8Array }> {
  const { buildPptxFixtureBytes } = await import("../../lib/presentation/fixtures/build");
  const { readPresentationBytes } = await import("../../lib/presentation/open");
  const bytes = await buildPptxFixtureBytes();
  return { deck: await readPresentationBytes(bytes, "pptx", "mini-deck"), bytes };
}

async function animationsDeck(): Promise<PresentationModel> {
  const { buildPptxAnimationsFixtureBytes } = await import("../../lib/presentation/fixtures/build");
  const { readPresentationBytes } = await import("../../lib/presentation/open");
  return readPresentationBytes(await buildPptxAnimationsFixtureBytes(), "pptx", "animations");
}

function spyRouter(router: Router) {
  const realReplace = router.replace.bind(router);
  const replace = vi.spyOn(router, "replace").mockImplementation((to) => {
    const nav = realReplace(to);
    pendingNavigations.push(nav.catch(() => undefined));
    return nav;
  });
  const realPush = router.push.bind(router);
  const push = vi.spyOn(router, "push").mockImplementation((to) => {
    const nav = realPush(to);
    pendingNavigations.push(nav.catch(() => undefined));
    return nav;
  });
  return { replace, push };
}

async function mountView(file: string, path: string) {
  const { default: View } = await import(`../${file}.vue`);
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(path);
  await router.isReady();
  const spies = spyRouter(router);
  const wrapper = mount(View, { global: { plugins: [router] }, attachTo: document.body });
  mounted.push(wrapper);
  await flushPromises();
  await flushPromises();
  return { wrapper, router, ...spies };
}

async function mountPresenter(path: string, opened: unknown = fakeWindow(), model?: PresentationModel) {
  const store = await import("../../lib/documentStore");
  if (model) store.storePresentation(model);
  const open = vi.spyOn(window, "open").mockReturnValue(opened as Window);
  const mountedView = await mountView("PresenterView", path);
  return { ...mountedView, open, store };
}

/** The persisted notes scale. */
const stored = () => JSON.parse(localStorage.getItem("docu-edit:preferences") ?? "{}").presentationPresenterNotesScale;
/** Every persisted preference (SPEC-146's timer mode and length among them). */
const prefs = () => JSON.parse(localStorage.getItem("docu-edit:preferences") ?? "{}");
const text = (wrapper: VueWrapper, testid: string) => wrapper.get(`[data-testid="${testid}"]`).text();
function key(name: string, init: KeyboardEventInit = {}) {
  window.dispatchEvent(new KeyboardEvent("keydown", { key: name, ...init }));
}
/** Real BroadcastChannel messages cross on the event loop. */
const settle = () => new Promise((resolve) => setTimeout(resolve, 30));

describe("PresenterView — load and surface", () => {
  it("shows the current slide, the next slide, the notes and the counter, and opens the audience once", async () => {
    const { deck } = await miniDeck();
    const { wrapper, open } = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow(), deck);
    expect(wrapper.find('[data-testid="presenter-view"]').exists()).toBe(true);
    expect(text(wrapper, "presenter-current-slide")).toContain("Quarterly Review");
    expect(text(wrapper, "presenter-next-slide")).toContain("Ship presentations MVP");
    expect(wrapper.findAll('[data-testid="presenter-notes"] p').map((p) => p.text())).toEqual([
      "Open with the Q3 highlight",
      "Pause before the revenue chart",
    ]);
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 1 of 4");
    expect(wrapper.find('[data-testid="presenter-build-indicator"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="presenter-live"]').attributes("aria-live")).toBe("polite");
    expect(open).toHaveBeenCalledTimes(1);
    const [url, name, features] = open.mock.calls[0];
    expect(url).toMatch(new RegExp(`^/slides/${deck.id}/present/1\\?follow=[\\w-]+$`));
    expect(name).toBe(`docu-edit-audience-${deck.id}`);
    expect(features).toMatch(/popup=yes,width=\d+,height=\d+,left=-?\d+,top=-?\d+/);
  });

  it("clamps an out-of-range n and replaces the URL", async () => {
    const { deck } = await miniDeck();
    const { wrapper, replace } = await mountPresenter(`/slides/${deck.id}/presenter/99`, fakeWindow(), deck);
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 4 of 4");
    expect(replace).toHaveBeenCalledWith(`/slides/${deck.id}/presenter/4`);
    expect(text(wrapper, "presenter-next-end")).toBe("End of show");
  });

  it("shows presenter-not-found with Back to library for an unknown id", async () => {
    const { wrapper, push, open } = await mountPresenter("/slides/missing/presenter/1");
    expect(wrapper.find('[data-testid="presenter-not-found"]').exists()).toBe(true);
    expect(wrapper.text()).toContain("Back to library");
    await wrapper.get('[data-testid="presentation-back"]').trigger("click");
    expect(push).toHaveBeenCalledWith("/");
    expect(open).not.toHaveBeenCalled();
  });

  it("redirects a text-document id to /doc/:id", async () => {
    const store = await import("../../lib/documentStore");
    const { createDocument } = await import("../../lib/documentModel");
    const doc = createDocument();
    store.storeDocument(doc);
    const { replace } = await mountPresenter(`/slides/${doc.id}/presenter/1`);
    expect(replace).toHaveBeenCalledWith(`/doc/${doc.id}`);
  });

  it("counts builds from slideBuildCount, hidden on a slide with none", async () => {
    const deck = await animationsDeck();
    const { wrapper } = await mountPresenter(`/slides/${deck.id}/presenter/2`, fakeWindow(), deck);
    expect(text(wrapper, "presenter-build-indicator")).toBe("Build 0 of 1");
    key("ArrowRight");
    await flushPromises();
    expect(text(wrapper, "presenter-build-indicator")).toBe("Build 1 of 1");
    // The pane mirrors the build: a group is animating in it now.
    expect(wrapper.get('[data-testid="presenter-current-slide"]').attributes("data-building")).toBe("true");
    key("ArrowRight"); // mid-build: completes it and holds
    await flushPromises();
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 2 of 3");
    expect(wrapper.get('[data-testid="presenter-current-slide"]').attributes("data-building")).toBe("false");
    key("ArrowRight");
    await flushPromises();
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 3 of 3");
    expect(wrapper.find('[data-testid="presenter-build-indicator"]').exists()).toBe(false);
  });
});

describe("PresenterView — input", () => {
  it("navigates with keys and buttons, blanks with b/w (badge) and clears the blank on navigation", async () => {
    const { deck } = await miniDeck();
    const { wrapper } = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow(), deck);
    key("ArrowRight");
    await flushPromises();
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 2 of 4");
    expect(text(wrapper, "presenter-live")).toBe("Slide 2 of 4 — Roadmap");
    await wrapper.get('[data-testid="presenter-next-button"]').trigger("click");
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 3 of 4");
    await wrapper.get('[data-testid="presenter-prev-button"]').trigger("click");
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 2 of 4");
    key("b");
    await flushPromises();
    expect(wrapper.find('[data-testid="presenter-blank-badge"]').exists()).toBe(true);
    expect(text(wrapper, "presenter-live")).toBe("Screen blanked");
    key("w");
    await flushPromises();
    expect(text(wrapper, "presenter-blank-badge")).toContain("white");
    key("ArrowRight"); // clears the blank, stays on the slide
    await flushPromises();
    expect(wrapper.find('[data-testid="presenter-blank-badge"]').exists()).toBe(false);
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 2 of 4");
    key("End");
    await flushPromises();
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 4 of 4");
    key("Home");
    key("3");
    key("Enter");
    await flushPromises();
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 3 of 4");
  });

  it("ignores modifier combos and keys typed into an input", async () => {
    const { deck } = await miniDeck();
    const { wrapper } = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow(), deck);
    key("ArrowRight", { ctrlKey: true });
    key("b", { metaKey: true });
    const input = document.createElement("input");
    document.body.append(input);
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
    input.remove();
    await flushPromises();
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 1 of 4");
    expect(wrapper.find('[data-testid="presenter-blank-badge"]').exists()).toBe(false);
  });

  it("opens the grid with g or its toggle, jumps by Enter/click, and Escape closes it before exiting", async () => {
    const { deck } = await miniDeck();
    const { wrapper, replace } = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow(), deck);
    key("g");
    await flushPromises();
    const grid = wrapper.get('[data-testid="presenter-grid"]');
    expect(grid.attributes("role")).toBe("listbox");
    const options = wrapper.findAll('[data-testid="presenter-grid-slide"]');
    expect(options).toHaveLength(4);
    expect(options[1].attributes("aria-label")).toBe("Slide 2 of 4 — Roadmap");
    expect(options[0].attributes("tabindex")).toBe("0");
    // Roving focus: ArrowRight moves the tabbable option; Enter jumps there.
    await grid.trigger("keydown", { key: "ArrowRight" });
    await grid.trigger("keydown", { key: "ArrowRight" });
    await nextTick();
    expect(wrapper.findAll('[data-testid="presenter-grid-slide"]')[2].attributes("tabindex")).toBe("0");
    await grid.trigger("keydown", { key: "Enter" });
    await flushPromises();
    expect(wrapper.find('[data-testid="presenter-grid"]').exists()).toBe(false);
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 3 of 4");
    // Toggle + click.
    await wrapper.get('[data-testid="presenter-grid-toggle"]').trigger("click");
    await wrapper.findAll('[data-testid="presenter-grid-slide"]')[3].trigger("click");
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 4 of 4");
    // Escape closes an open grid first, and only then exits.
    key("g");
    await flushPromises();
    replace.mockClear();
    key("Escape");
    await flushPromises();
    expect(wrapper.find('[data-testid="presenter-grid"]').exists()).toBe(false);
    expect(replace).not.toHaveBeenCalled();
  });

  it("Escape exits to the editor with the showing slide selected, saying bye and closing the audience", async () => {
    const { deck } = await miniDeck();
    const audience = fakeWindow();
    const { replace, open } = await mountPresenter(`/slides/${deck.id}/presenter/2`, audience, deck);
    const session = new URL(String(open.mock.calls[0][0]), "http://x").searchParams.get("follow");
    const listener = new BroadcastChannel(`docu-edit:present:${deck.id}`);
    const heard: unknown[] = [];
    listener.addEventListener("message", (event) => heard.push(event.data));
    key("Escape");
    await flushPromises();
    await settle();
    listener.close();
    expect(replace).toHaveBeenCalledWith({ path: `/slides/${deck.id}`, query: { return: "2" } });
    expect(audience.close).toHaveBeenCalled();
    expect(heard).toContainEqual({ kind: "bye", session, from: "presenter" });
  });

  it("presenter-exit ends the show the same way", async () => {
    const { deck } = await miniDeck();
    const { wrapper, replace } = await mountPresenter(`/slides/${deck.id}/presenter/3`, fakeWindow(), deck);
    await wrapper.get('[data-testid="presenter-exit"]').trigger("click");
    expect(replace).toHaveBeenCalledWith({ path: `/slides/${deck.id}`, query: { return: "3" } });
  });
});

describe("PresenterView — notes scale and timer", () => {
  it("steps the notes scale through PRESENTER_NOTES_SCALES with the buttons and -/+, persisted per device", async () => {
    const { PRESENTER_NOTES_SCALES, PREFERENCE_DEFAULTS } = await import("../../lib/preferences");
    expect(PRESENTER_NOTES_SCALES).toEqual([1, 1.25, 1.5, 2, 2.5]);
    expect(PREFERENCE_DEFAULTS.presentationPresenterNotesScale).toBe(1.25);
    const { deck } = await miniDeck();
    const { wrapper } = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow(), deck);
    const scale = () => wrapper.get('[data-testid="presenter-notes"]').attributes("data-scale");
    expect(scale()).toBe("1.25");
    await wrapper.get('[data-testid="presenter-notes-larger"]').trigger("click");
    expect(scale()).toBe("1.5");
    expect(stored()).toBe(1.5);
    key("+");
    key("+");
    key("+"); // past the top: stays at 2.5
    await flushPromises();
    expect(scale()).toBe("2.5");
    expect(wrapper.get('[data-testid="presenter-notes-larger"]').attributes("disabled")).toBeDefined();
    key("-");
    await flushPromises();
    expect(scale()).toBe("2");
    for (let i = 0; i < 5; i += 1) await wrapper.get('[data-testid="presenter-notes-smaller"]').trigger("click");
    expect(scale()).toBe("1");
    expect(wrapper.get('[data-testid="presenter-notes-smaller"]').attributes("disabled")).toBeDefined();
    expect(stored()).toBe(1);
  });

  // SPEC-146 rewrote SPEC-100's timer test for the show timer: the toggle
  // names its next action instead of carrying aria-pressed, and Stop (s, or
  // SPEC-100's r) replaces Reset — it zeroes AND halts the timer.
  it("starts the timer with the show; t/the toggle pause and resume it; r and s stop it", async () => {
    const { setClock } = await import("../../lib/core/clock");
    const t0 = Date.UTC(2026, 9, 23, 9, 0, 0);
    setClock(new Date(t0));
    try {
      const { deck } = await miniDeck();
      const { wrapper } = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow(), deck);
      const toggle = () => wrapper.get('[data-testid="presenter-timer-toggle"]');
      const stop = () => wrapper.get('[data-testid="presenter-timer-stop"]');
      const state = () => wrapper.get('[data-testid="presenter-timer"]').attributes("data-state");
      expect(text(wrapper, "presenter-timer")).toBe("0:00");
      expect(state()).toBe("running");
      expect(toggle().text()).toBe("Pause (T)");
      expect(toggle().attributes("aria-label")).toBe("Pause timer");
      expect(toggle().attributes("aria-pressed")).toBeUndefined();
      expect(text(wrapper, "presenter-timer-status")).toBe(""); // begin() never announces
      setClock(new Date(t0 + 75_000));
      await toggle().trigger("click"); // pause at 1:15
      expect(toggle().text()).toBe("Resume (T)");
      expect(toggle().attributes("aria-label")).toBe("Resume timer");
      expect(text(wrapper, "presenter-timer")).toBe("1:15");
      expect(state()).toBe("paused");
      expect(text(wrapper, "presenter-timer-status")).toBe("Timer paused");
      setClock(new Date(t0 + 600_000));
      key("t"); // resume: the paused stretch does not count
      await flushPromises();
      expect(toggle().text()).toBe("Pause (T)");
      expect(text(wrapper, "presenter-timer")).toBe("1:15");
      expect(text(wrapper, "presenter-timer-status")).toBe("Timer resumed");
      setClock(new Date(t0 + 605_000));
      key("r"); // SPEC-100's reset key now stops: zero, and halted
      await flushPromises();
      expect(text(wrapper, "presenter-timer")).toBe("0:00");
      expect(state()).toBe("stopped");
      expect(toggle().text()).toBe("Start (T)");
      expect(stop().attributes("disabled")).toBeDefined();
      expect(stop().text()).toBe("Stop (S)");
      expect(stop().attributes("aria-label")).toBe("Stop timer");
      expect(text(wrapper, "presenter-timer-status")).toBe("Timer stopped");
      expect(wrapper.find('[data-testid="presenter-timer-reset"]').exists()).toBe(false);
      setClock(new Date(t0 + 700_000));
      expect(text(wrapper, "presenter-timer")).toBe("0:00"); // stopped stays at the start
      key("t"); // Start begins a fresh count
      await flushPromises();
      expect(state()).toBe("running");
      expect(stop().attributes("disabled")).toBeUndefined();
      expect(text(wrapper, "presenter-timer-status")).toBe("Timer started");
      setClock(new Date(t0 + 710_000));
      await stop().trigger("click"); // the button, from running
      expect(state()).toBe("stopped");
      key("t");
      await flushPromises();
      setClock(new Date(t0 + 720_000));
      key("s"); // s, from running
      await flushPromises();
      expect(state()).toBe("stopped");
      expect(text(wrapper, "presenter-timer")).toBe("0:00");
    } finally {
      setClock(null);
    }
  });

  it("with auto-start off the show begins stopped, showing Start (T)", async () => {
    localStorage.setItem("docu-edit:preferences", JSON.stringify({ presentationTimerAutoStart: false }));
    const { deck } = await miniDeck();
    const { wrapper } = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow(), deck);
    const readout = wrapper.get('[data-testid="presenter-timer"]');
    expect(readout.text()).toBe("0:00");
    expect(readout.attributes("data-state")).toBe("stopped");
    expect(text(wrapper, "presenter-timer-toggle")).toBe("Start (T)");
    expect(wrapper.get('[data-testid="presenter-timer-stop"]').attributes("disabled")).toBeDefined();
    expect(text(wrapper, "presenter-timer-status")).toBe("");
    key("t");
    await flushPromises();
    expect(wrapper.get('[data-testid="presenter-timer"]').attributes("data-state")).toBe("running");
    expect(text(wrapper, "presenter-timer-status")).toBe("Timer started");
  });

  it("an audience timerToggle intent toggles the presenter's timer, which posts timer and no state", async () => {
    const { deck } = await miniDeck();
    const { wrapper, open } = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow(), deck);
    const session = new URL(String(open.mock.calls[0][0]), "http://x").searchParams.get("follow");
    const port = new BroadcastChannel(`docu-edit:present:${deck.id}`);
    const heard: Array<{ kind: string; timer?: unknown; state?: { seq: number } }> = [];
    port.addEventListener("message", (event) => heard.push(event.data));
    port.postMessage({ kind: "hello", session, from: "audience" });
    await settle();
    // The hello is answered with the state, then the timer.
    expect(heard.map((message) => message.kind).slice(0, 2)).toEqual(["state", "timer"]);
    const seq = Math.max(...heard.flatMap((message) => (message.state ? [message.state.seq] : [])));
    port.postMessage({ kind: "intent", session, action: "timerToggle" });
    await settle();
    await flushPromises();
    expect(wrapper.get('[data-testid="presenter-timer"]').attributes("data-state")).toBe("paused");
    expect(heard.at(-1)).toMatchObject({ kind: "timer", timer: { startedAtMs: null } });
    expect(Math.max(...heard.flatMap((message) => (message.state ? [message.state.seq] : [])))).toBe(seq);
    port.postMessage({ kind: "intent", session, action: "timerStop" });
    await settle();
    await flushPromises();
    port.close();
    expect(wrapper.get('[data-testid="presenter-timer"]').attributes("data-state")).toBe("stopped");
  });

  it("switches mode with aria-pressed buttons and sets the talk length, keeping the count", async () => {
    const { setClock } = await import("../../lib/core/clock");
    const t0 = Date.UTC(2026, 9, 23, 9, 0, 0);
    setClock(new Date(t0));
    try {
      const { deck } = await miniDeck();
      const { wrapper } = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow(), deck);
      const readout = () => wrapper.get('[data-testid="presenter-timer"]');
      const stopwatch = () => wrapper.get('[data-testid="presenter-timer-mode-stopwatch"]');
      const countdown = () => wrapper.get('[data-testid="presenter-timer-mode-countdown"]');
      const group = wrapper.get('[role="group"][aria-label="Timer mode"]');
      expect(group.find('[data-testid="presenter-timer-mode-stopwatch"]').text()).toBe("Stopwatch");
      expect(group.find('[data-testid="presenter-timer-mode-countdown"]').text()).toBe("Countdown");
      expect(stopwatch().attributes("aria-pressed")).toBe("true");
      expect(countdown().attributes("aria-pressed")).toBe("false");
      expect(readout().attributes("aria-label")).toBe("Elapsed time");
      expect(wrapper.find('[data-testid="presenter-timer-length"]').exists()).toBe(false);

      setClock(new Date(t0 + 75_000));
      await countdown().trigger("click");
      expect(prefs().presentationTimerMode).toBe("countdown");
      expect(countdown().attributes("aria-pressed")).toBe("true");
      expect(stopwatch().attributes("aria-pressed")).toBe("false");
      expect(readout().text()).toBe("18:45"); // 20 minutes less the 1:15 already elapsed
      expect(readout().attributes("data-mode")).toBe("countdown");
      expect(readout().attributes("data-state")).toBe("running");
      expect(readout().attributes("aria-label")).toBe("Time remaining");

      const length = wrapper.get('[data-testid="presenter-timer-length"]');
      const input = length.element as HTMLInputElement;
      expect(length.attributes("type")).toBe("text");
      expect(length.attributes("inputmode")).toBe("numeric");
      expect(input.value).toBe("20");
      expect(wrapper.get('label[for="presenter-timer-length"]').text()).toBe("Talk length (minutes)");
      input.value = "abc"; // not a number: the field restores
      await length.trigger("change");
      expect(input.value).toBe("20");
      expect(prefs().presentationTimerCountdownMinutes ?? 20).toBe(20);
      input.value = "1";
      await length.trigger("change");
      expect(prefs().presentationTimerCountdownMinutes).toBe(1);
      expect(readout().text()).toBe("\u22120:15"); // straight into overtime, the count kept
      expect(readout().attributes("data-overtime")).toBe("true");
      expect(readout().attributes("aria-label")).toBe("Time over");
      expect(text(wrapper, "presenter-timer-status")).toBe("Time is up");
      input.value = "2.6"; // rounded
      await length.trigger("change");
      expect(prefs().presentationTimerCountdownMinutes).toBe(3);
      expect(input.value).toBe("3");

      // Typing t into the length field neither toggles the timer nor advances.
      length.element.dispatchEvent(new KeyboardEvent("keydown", { key: "t", bubbles: true }));
      length.element.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
      await flushPromises();
      expect(readout().attributes("data-state")).toBe("running");
      expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 1 of 4");

      // ArrowRight on a focused mode button advances exactly once and leaves the mode.
      (countdown().element as HTMLElement).focus();
      countdown().element.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
      await flushPromises();
      expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 2 of 4");
      expect(countdown().attributes("aria-pressed")).toBe("true");
      // Enter on it is the button's own (it does not advance).
      countdown().element.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
      await flushPromises();
      expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 2 of 4");

      await stopwatch().trigger("click");
      expect(prefs().presentationTimerMode).toBe("stopwatch");
      expect(readout().text()).toBe("1:15");
      expect(wrapper.find('[data-testid="presenter-timer-length"]').exists()).toBe(false);
    } finally {
      setClock(null);
    }
  });
});

describe("PresenterView — degraded modes", () => {
  it("a blocked pop-up falls back to single-window mode: banner, retry, Present here, solo note, the pane is the show", async () => {
    const { deck } = await miniDeck();
    const { wrapper, open } = await mountPresenter(`/slides/${deck.id}/presenter/1`, null, deck);
    const banner = wrapper.get('[data-testid="presenter-audience-blocked"]');
    expect(banner.find('[data-testid="presenter-audience-open"]').exists()).toBe(true);
    expect(banner.text()).toContain("Present here");
    expect(text(wrapper, "presenter-solo-note")).toContain("the current slide is the show");
    expect(text(wrapper, "presenter-live")).toContain("Single-window mode");
    // The current pane is the show: a click advances it, and B paints it.
    await wrapper.get('[data-testid="presenter-current-slide"]').trigger("click");
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 2 of 4");
    key("b");
    await flushPromises();
    expect(wrapper.find('[data-testid="presenter-current-slide"] .pane-blank').exists()).toBe(true);
    // Present here dismisses the banner; the note keeps a way to the audience window.
    await wrapper.get('[data-testid="presenter-present-here"]').trigger("click");
    expect(wrapper.find('[data-testid="presenter-audience-blocked"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="presenter-solo-note"] [data-testid="presenter-audience-open"]').exists()).toBe(true);
    // A fresh click retries the pop-up; this time it opens.
    open.mockReturnValue(fakeWindow() as unknown as Window);
    await wrapper.get('[data-testid="presenter-audience-open"]').trigger("click");
    await flushPromises();
    expect(open).toHaveBeenCalledTimes(2);
    expect(String(open.mock.calls[1][0])).toMatch(new RegExp(`/slides/${deck.id}/present/2\\?follow=`));
    expect(wrapper.find('[data-testid="presenter-solo-note"]').exists()).toBe(false);
    // Off single-window mode the pane is a preview again: a click does not advance.
    await wrapper.get('[data-testid="presenter-current-slide"]').trigger("click");
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 2 of 4");
  });

  it("with no channel transport it presents here without trying a window it could not drive", async () => {
    vi.stubGlobal("BroadcastChannel", undefined);
    vi.spyOn(Storage.prototype, "setItem").mockImplementation(() => {
      throw new Error("blocked");
    });
    const { deck } = await miniDeck();
    const { wrapper, open } = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow(), deck);
    expect(open).not.toHaveBeenCalled();
    expect(text(wrapper, "presenter-audience-blocked")).toContain("can't link a second window");
    expect(wrapper.find('[data-testid="presenter-solo-note"]').exists()).toBe(true);
  });

  it("shows presenter-audience-lost when the audience says bye, keeping its state, and re-opens it", async () => {
    const { deck } = await miniDeck();
    const { wrapper, open } = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow(), deck);
    const session = new URL(String(open.mock.calls[0][0]), "http://x").searchParams.get("follow");
    key("ArrowRight");
    await flushPromises();
    const port = new BroadcastChannel(`docu-edit:present:${deck.id}`);
    port.postMessage({ kind: "hello", session, from: "audience" });
    await settle();
    expect(wrapper.find('[data-testid="presenter-audience-lost"]').exists()).toBe(false);
    port.postMessage({ kind: "bye", session, from: "audience" });
    await settle();
    port.close();
    expect(wrapper.find('[data-testid="presenter-audience-lost"]').exists()).toBe(true);
    expect(text(wrapper, "presenter-slide-counter")).toBe("Slide 2 of 4");
    await wrapper.get('[data-testid="presenter-audience-lost"] [data-testid="presenter-audience-open"]').trigger("click");
    await flushPromises();
    expect(open).toHaveBeenCalledTimes(2);
    expect(String(open.mock.calls[1][0])).toContain(`/present/2?follow=${session}`);
  });

  it("places the audience on the first non-primary screen when the permission is already granted", async () => {
    const { useAudienceWindow } = await import("./useAudienceWindow");
    const { ref } = await import("vue");
    const open = vi.fn(() => fakeWindow() as unknown as Window);
    let audience!: ReturnType<typeof useAudienceWindow>;
    mounted.push(
      mount(
        defineComponent({
          setup() {
            audience = useAudienceWindow({
              deckId: ref("d1"),
              session: "s",
              slideNumber: () => 3,
              transport: "broadcast",
              openWindow: open,
              screenAccess: async () => ({ state: "granted", placement: { left: 1920, top: 0, width: 1280, height: 720 } }),
            });
            return () => null;
          },
        }),
      ),
    );
    await audience.open();
    expect(open).toHaveBeenCalledWith(
      "/slides/d1/present/3?follow=s&placed=1",
      "docu-edit-audience-d1",
      "popup=yes,left=1920,top=0,width=1280,height=720",
    );
  });

  it("asks for the screens permission at most once per session, never blocking the open, silent on refusal", async () => {
    const { useAudienceWindow } = await import("./useAudienceWindow");
    const { ref } = await import("vue");
    sessionStorage.clear();
    const open = vi.fn((_url: string, _name: string, _features: string) => fakeWindow() as unknown as Window);
    const ask = vi.fn(() => Promise.reject(new Error("NotAllowedError")));
    let audience!: ReturnType<typeof useAudienceWindow>;
    mounted.push(
      mount(
        defineComponent({
          setup() {
            audience = useAudienceWindow({
              deckId: ref("d1"),
              session: "s",
              slideNumber: () => 1,
              transport: "broadcast",
              openWindow: open,
              screenAccess: async () => ({ state: "prompt", placement: null }),
              askForScreens: ask,
            });
            return () => null;
          },
        }),
      ),
    );
    await audience.open();
    await audience.open();
    await flushPromises();
    expect(open).toHaveBeenCalledTimes(2);
    expect(open.mock.calls[0][0]).toBe("/slides/d1/present/1?follow=s"); // an ordinary pop-up, with the drag hint
    expect(ask).toHaveBeenCalledTimes(1);
    expect(audience.status.value).toBe("opening");
  });
});

describe("PresenterView — presenting writes nothing", () => {
  it("never stores, and a presented-and-exited deck downloads byte-identical part for part", async () => {
    const { deck, bytes } = await miniDeck();
    const store = await import("../../lib/documentStore");
    store.storePresentation(deck);
    const storeSpy = vi.spyOn(store, "storePresentation");
    const { wrapper } = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow());
    for (const name of ["ArrowRight", "b", "ArrowRight", "End", "w", "Home", "g", "Escape"]) {
      key(name);
      await flushPromises();
    }
    await wrapper.get('[data-testid="presenter-exit"]').trigger("click");
    await flushPromises();
    expect(storeSpy).not.toHaveBeenCalled();
    const { writePptx } = await import("../../lib/presentation/pptx/writer");
    const { unzip } = await import("../../lib/core/zip");
    const { equalBytes } = await import("../../lib/core/equality");
    const written = await unzip((await writePptx(store.getPresentation(deck.id)!)).bytes);
    const source = await unzip(bytes);
    expect([...written.keys()]).toEqual([...source.keys()]);
    for (const [name, data] of source) {
      expect(equalBytes(written.get(name)!, data), `part ${name}`).toBe(true);
    }
  });
});

describe("PresenterView with a following audience play view", () => {
  // Both views share this one jsdom window, so a key would reach both their
  // window-level handlers; the pair is driven through its buttons instead.
  it("the audience follows the presenter, forwards its own input, and is promoted when the presenter goes", async () => {
    const { deck } = await miniDeck();
    const presenter = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow(), deck);
    const session = new URL(String(presenter.open.mock.calls[0][0]), "http://x").searchParams.get("follow");
    const audience = await mountView("PresentationPlayView", `/slides/${deck.id}/present/1?follow=${session}`);
    await settle();
    const audienceCounter = () => text(audience.wrapper, "play-slide-counter");
    expect(text(presenter.wrapper, "presenter-live")).toBe("Audience window connected");

    // Presenter input moves both; the audience keeps ?follow= as it tracks.
    await presenter.wrapper.get('[data-testid="presenter-next-button"]').trigger("click");
    await settle();
    expect(audienceCounter()).toBe("2 / 4");
    expect(audience.replace).toHaveBeenCalledWith({ path: `/slides/${deck.id}/present/2`, query: { follow: session } });

    // Audience input is forwarded to the presenter, and both move.
    await audience.wrapper.get('[data-testid="play-next"]').trigger("click");
    await settle();
    await flushPromises();
    expect(text(presenter.wrapper, "presenter-slide-counter")).toBe("Slide 3 of 4");
    expect(audienceCounter()).toBe("3 / 4");

    // The presenter ends the show: it says bye, and the audience is promoted
    // to plain playback on the same slide — ?follow= drops, input runs locally.
    await presenter.wrapper.get('[data-testid="presenter-exit"]').trigger("click");
    await settle();
    await flushPromises();
    expect(presenter.replace).toHaveBeenCalledWith({ path: `/slides/${deck.id}`, query: { return: "3" } });
    expect(audience.replace).toHaveBeenCalledWith(`/slides/${deck.id}/present/3`);
    await audience.wrapper.get('[data-testid="play-next"]').trigger("click");
    await flushPromises();
    expect(audienceCounter()).toBe("4 / 4");
  });
});

describe("PresenterView's show timer across the window pair (SPEC-146)", () => {
  it("the following audience shows no timer, and on promotion adopts the presenter's as it was", async () => {
    const { setClock } = await import("../../lib/core/clock");
    const t0 = Date.UTC(2026, 9, 23, 9, 0, 0);
    setClock(new Date(t0));
    try {
      const { deck } = await miniDeck();
      const presenter = await mountPresenter(`/slides/${deck.id}/presenter/1`, fakeWindow(), deck);
      const session = new URL(String(presenter.open.mock.calls[0][0]), "http://x").searchParams.get("follow");
      const audience = await mountView("PresentationPlayView", `/slides/${deck.id}/present/1?follow=${session}`);
      await settle();
      expect(audience.wrapper.findAll('[data-testid="play-timer"]')).toHaveLength(0);
      setClock(new Date(t0 + 75_000));
      await presenter.wrapper.get('[data-testid="presenter-timer-toggle"]').trigger("click"); // paused at 1:15
      await settle();
      setClock(new Date(t0 + 600_000)); // the paused value must not move
      await presenter.wrapper.get('[data-testid="presenter-exit"]').trigger("click");
      await settle();
      await flushPromises();
      const timer = audience.wrapper.get('[data-testid="play-timer"]');
      expect(timer.text()).toBe("1:15");
      expect(timer.attributes("data-state")).toBe("paused");
      expect(audience.wrapper.find('[data-testid="play-timer-mode-stopwatch"]').exists()).toBe(false);
      await audience.wrapper.get('[data-testid="play-timer-toggle"]').trigger("click"); // it owns it now
      expect(audience.wrapper.get('[data-testid="play-timer"]').attributes("data-state")).toBe("running");
      expect(text(audience.wrapper, "play-slide-counter")).toBe("1 / 4"); // the click did not advance
    } finally {
      setClock(null);
    }
  });
});
