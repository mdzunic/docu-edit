// SPEC-084 — the PresentationPlayView route: guards (text-id redirect, unknown
// id, out-of-range clamp), history discipline (replace not push), the announce
// string, chrome-free render, and the exit selection handoff. The idb seam is
// mocked so the store runs in-memory; jsdom has no fullscreen/matchMedia APIs,
// which the view already guards, so mounting exercises the tolerated path.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { flushPromises, mount } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";
import type { PresentationModel } from "../lib/presentation/model";

beforeEach(() => {
  vi.resetModules();
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
    loadAllTemplateRecords: vi.fn().mockResolvedValue([]),
    putTemplateRecord: vi.fn().mockResolvedValue(undefined),
    deleteTemplateRecord: vi.fn().mockResolvedValue(undefined),
  }));
});

// Navigations the view fires and forgets (exit/redirect replaces) lazy-load
// other route components; vitest 4 fails the run if those imports settle after
// the file's environment is torn down, so every navigation is collected here
// and awaited before teardown.
const pendingNavigations: Promise<unknown>[] = [];

afterEach(async () => {
  await Promise.allSettled(pendingNavigations);
  pendingNavigations.length = 0;
  vi.doUnmock("../lib/idb");
  vi.restoreAllMocks();
});

async function freshDeck(): Promise<PresentationModel> {
  const { buildPptxFixtureBytes } = await import("../lib/presentation/fixtures/build");
  const { readPresentationBytes } = await import("../lib/presentation/open");
  return readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "mini-deck");
}

async function mountAt(path: string, model?: PresentationModel) {
  const store = await import("../lib/documentStore");
  if (model) store.storePresentation(model);
  const { default: PresentationPlayView } = await import("./PresentationPlayView.vue");
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(path);
  await router.isReady();
  // Spy before mount so the load-time URL sync (clamp reflection) is captured;
  // cross-route navigations resolve lazily in this env, so the spec asserts the
  // router calls rather than the resulting currentRoute.
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
  const wrapper = mount(PresentationPlayView, { global: { plugins: [router] } });
  await flushPromises();
  await flushPromises();
  return { wrapper, router, store, replace, push };
}

function key(name: string, init: KeyboardEventInit = {}) {
  window.dispatchEvent(new KeyboardEvent("keydown", { key: name, ...init }));
}

describe("PresentationPlayView", () => {
  it("renders the current slide with no app chrome (AC-3/11/12)", async () => {
    const deck = await freshDeck();
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    const region = wrapper.get('[data-testid="presentation-play"]');
    expect(region.attributes("role")).toBe("region");
    expect(region.attributes("aria-roledescription")).toBe("slide show");
    expect(region.attributes("aria-label")).toBe(deck.title);
    // The SPEC-079 stage renders (one renderer), but no toolbar/filmstrip.
    expect(wrapper.find('[data-testid="slide-stage"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="slide-filmstrip"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="presentation-toolbar"]').exists()).toBe(false);
  });

  it("announces 'Slide N of M — title' via a polite live region (AC-50)", async () => {
    const deck = await freshDeck();
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    const live = wrapper.get('[data-testid="play-live"]');
    expect(live.attributes("aria-live")).toBe("polite");
    expect(live.text()).toBe("Slide 1 of 4 — Quarterly Review");
  });

  it("clamps an out-of-range n and replaces the URL (AC-5)", async () => {
    const deck = await freshDeck();
    const { wrapper, replace } = await mountAt(`/slides/${deck.id}/present/99`, deck);
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("4 / 4");
    expect(replace).toHaveBeenCalledWith(`/slides/${deck.id}/present/4`);
  });

  it("advances with ArrowRight and tracks the slide by replace, never push (AC-14/37)", async () => {
    const deck = await freshDeck();
    const { wrapper, replace, push } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    replace.mockClear();
    key("ArrowRight");
    await flushPromises();
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("2 / 4");
    expect(replace).toHaveBeenCalledWith(`/slides/${deck.id}/present/2`);
    expect(push).not.toHaveBeenCalled();
  });

  it("shows the end screen after the last slide, then exits (AC-30/31/34/35)", async () => {
    const deck = await freshDeck();
    const { wrapper, replace } = await mountAt(`/slides/${deck.id}/present/4`, deck);
    replace.mockClear();
    key("ArrowRight"); // past the last slide → end screen
    await flushPromises();
    const end = wrapper.get('[data-testid="play-end-screen"]');
    expect(end.attributes("role")).toBe("alert");
    expect(end.text()).toContain("End of slide show");
    key("ArrowRight"); // one more → exit to the editor, last slide selected
    await flushPromises();
    expect(replace).toHaveBeenCalledWith({ path: `/slides/${deck.id}`, query: { return: "4" } });
  });

  it("Esc exits playback to the editor with the showing slide selected (AC-33/34/35)", async () => {
    const deck = await freshDeck();
    const { replace } = await mountAt(`/slides/${deck.id}/present/2`, deck);
    replace.mockClear();
    key("Escape");
    await flushPromises();
    expect(replace).toHaveBeenCalledWith({ path: `/slides/${deck.id}`, query: { return: "2" } });
  });

  it("shows a GIF from a fresh URL on each visit to its slide, so it plays from the start", async () => {
    // Chrome keeps one playback timeline per image URL: from the shared data:
    // URI, a non-looping GIF that already played in the editor came up frozen.
    let minted = 0;
    vi.spyOn(URL, "createObjectURL").mockImplementation(() => `blob:showing/${++minted}`);
    vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => undefined);
    const deck = await freshDeck();
    const mediaId = deck.slides[3].frames.flatMap((f) => (f.kind === "image" ? [f.mediaId] : []))[0];
    const part = deck.media?.find((p) => p.mediaId === mediaId);
    expect(part).toBeDefined();
    if (part) part.contentType = "image/gif";

    const { wrapper } = await mountAt(`/slides/${deck.id}/present/4`, deck);
    const gifSrc = () => wrapper.findAll("img").map((img) => img.attributes("src") ?? "").find((src) => src.startsWith("blob:"));
    const first = gifSrc();
    expect(first).toMatch(/^blob:showing\//);

    key("ArrowLeft");
    await flushPromises();
    expect(gifSrc()).toBeUndefined(); // slide 3 has no GIF
    key("ArrowRight");
    await flushPromises();
    const second = gifSrc();
    expect(second).toMatch(/^blob:showing\//);
    expect(second).not.toBe(first);
  });

  it("number-then-Enter jumps to that slide (AC-27)", async () => {
    const deck = await freshDeck();
    const { wrapper, replace } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    replace.mockClear();
    key("3");
    key("Enter");
    await flushPromises();
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("3 / 4");
    expect(replace).toHaveBeenCalledWith(`/slides/${deck.id}/present/3`);
  });

  it("redirects a text-document id to /doc/:id (AC-6)", async () => {
    const store = await import("../lib/documentStore");
    const { createDocument } = await import("../lib/documentModel");
    const doc = createDocument();
    store.storeDocument(doc);
    const { replace } = await mountAt(`/slides/${doc.id}/present/1`);
    expect(replace).toHaveBeenCalledWith(`/doc/${doc.id}`);
  });

  it("shows the not-found state with a Back to library action for an unknown id (AC-7)", async () => {
    const { wrapper, push } = await mountAt("/slides/missing/present/1");
    expect(wrapper.find('[data-testid="presentation-play-not-found"]').exists()).toBe(true);
    await wrapper.get('[data-testid="presentation-play-back"]').trigger("click");
    expect(push).toHaveBeenCalledWith("/");
  });
});

// SPEC-097 AC-49/50 — a click on a rendered link navigates and never also
// advances a build: in-deck through jumpTo, external through window.open.
describe("PresentationPlayView links (SPEC-097)", () => {
  /** The mini-deck with slide 1's first text frame given a linked run (to
   * `link`) and a second frame entering on click, so a stray advance would
   * visibly play a build. */
  async function linkedDeck(link: { linkSlide?: number; linkHref?: string }): Promise<PresentationModel> {
    const deck = await freshDeck();
    const first = deck.slides[0];
    const textFrames = first.frames.filter((frame) => frame.kind === "text");
    const host = textFrames[0];
    if (host.kind !== "text") throw new Error("fixture slide 1 has no text frame");
    host.paragraphs = [{ runs: [{ text: "Jump " }, { text: "here", ...link }] }];
    const target = textFrames[1] ?? host;
    first.animations = [{ frameId: target.id, effect: "fadeIn", trigger: "click" }];
    return deck;
  }

  it("an in-deck link jumps to its slide without also advancing the build", async () => {
    const deck = await linkedDeck({ linkSlide: 3 });
    const { wrapper, replace } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    replace.mockClear();
    const link = wrapper.get('[data-testid="presentation-play"] [data-testid="slide-link"]');
    expect(link.text()).toBe("here");
    await link.trigger("click");
    await flushPromises();
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("3 / 4");
    expect(replace).toHaveBeenCalledWith(`/slides/${deck.id}/present/3`);
    expect(wrapper.get('[data-testid="presentation-play"]').attributes("data-building")).toBe("false");
  });

  it("an external link opens a new tab and the show stays put (no advance)", async () => {
    const open = vi.spyOn(window, "open").mockReturnValue(null);
    const deck = await linkedDeck({ linkHref: "https://example.com/deck" });
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    await wrapper.get('[data-testid="slide-link"]').trigger("click");
    await flushPromises();
    expect(open).toHaveBeenCalledWith("https://example.com/deck", "_blank", "noopener,noreferrer");
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("1 / 4");
    expect(wrapper.get('[data-testid="presentation-play"]').attributes("data-building")).toBe("false");
  });

  it("a click anywhere else still advances (the build plays)", async () => {
    const deck = await linkedDeck({ linkSlide: 3 });
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    await wrapper.get('[data-testid="presentation-play"]').trigger("click");
    await flushPromises();
    // The click-triggered build began on slide 1 (not a slide change).
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("1 / 4");
    expect(wrapper.get('[data-testid="presentation-play"]').attributes("data-building")).toBe("true");
  });
});

// SPEC-104 — hidden slides: a deep link to one still shows it, and only
// advancing/going back steps over one.
async function hiddenDeck(): Promise<PresentationModel> {
  const { buildPptxHiddenFixtureBytes } = await import("../lib/presentation/fixtures/build");
  const { readPresentationBytes } = await import("../lib/presentation/open");
  return readPresentationBytes(await buildPptxHiddenFixtureBytes(), "pptx", "hidden");
}

describe("PresentationPlayView hidden slides (SPEC-104)", () => {
  it("advancing from slide 1 steps over the hidden slide 2 onto slide 3, and back again", async () => {
    const deck = await hiddenDeck();
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    const counter = () => wrapper.get('[data-testid="play-slide-counter"]').text();
    expect(counter()).toBe("1 / 3");
    key("ArrowRight");
    await flushPromises();
    expect(counter()).toBe("3 / 3");
    key("ArrowLeft");
    await flushPromises();
    expect(counter()).toBe("1 / 3");
  });

  it("a deep link to the hidden slide still shows it", async () => {
    const deck = await hiddenDeck();
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/2`, deck);
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("2 / 3");
    key("ArrowRight");
    await flushPromises();
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("3 / 3");
  });
});

// SPEC-103 AC-25 — playing a deck is not editing it: the play view never
// writes a draft, a commit or a version, whatever it is driven through.
describe("PresentationPlayView writes nothing (SPEC-103)", () => {
  it("never writes a draft, a commit or a version", async () => {
    const putRecord = vi.fn().mockResolvedValue(undefined);
    const putVersionsRecord = vi.fn().mockResolvedValue(undefined);
    vi.doMock("../lib/idb", () => ({
      loadAllRecords: vi.fn().mockResolvedValue([]),
      putRecord,
      getRecord: vi.fn().mockResolvedValue(undefined),
      deleteRecord: vi.fn().mockResolvedValue(undefined),
      getVersionsRecord: vi.fn().mockResolvedValue(undefined),
      putVersionsRecord,
      deleteVersionsRecord: vi.fn().mockResolvedValue(undefined),
      loadAllTemplateRecords: vi.fn().mockResolvedValue([]),
      putTemplateRecord: vi.fn().mockResolvedValue(undefined),
      deleteTemplateRecord: vi.fn().mockResolvedValue(undefined),
    }));
    const deck = await freshDeck();
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    key("ArrowRight");
    key("ArrowRight");
    key("ArrowLeft");
    await flushPromises();
    key("Escape");
    await flushPromises();
    await new Promise((resolve) => setTimeout(resolve, 20));
    const written = putRecord.mock.calls.map(([record]) => record as { id: string });
    expect(written.filter((record) => record.id.startsWith("draft"))).toEqual([]);
    // The import's own idle write (storePresentation) is the only committed
    // write; a commit would carry the editor state beside the model.
    expect(written.filter((record) => "caret" in record || "slideIndex" in record)).toEqual([]);
    expect(putVersionsRecord).not.toHaveBeenCalled();
    wrapper.unmount();
  });
});

/** Real BroadcastChannel messages cross on the event loop. */
async function settle() {
  await new Promise((resolve) => setTimeout(resolve, 30));
  await flushPromises();
}

// SPEC-146 — the show timer in the full-screen show: a window that owns the
// show has it in the HUD (read-only), a following one has none; t/s/r drive it
// through the session and reveal the HUD; a toggle click never advances; the
// end screen's any-key exit keeps t; a follower forwards t as timerToggle; a
// promoted window adopts the presenter's timer; mode and length come from the
// preferences.
describe("PresentationPlayView show timer (SPEC-146)", () => {
  const SESSION = "timer-test";
  const T0 = Date.UTC(2026, 10, 5, 9, 0, 0);
  const cleanups: Array<() => void> = [];

  beforeEach(async () => {
    localStorage.clear();
    sessionStorage.clear();
    const { setClock } = await import("../lib/core/clock");
    setClock(new Date(T0));
    cleanups.push(() => setClock(null));
  });

  afterEach(() => {
    for (const cleanup of cleanups.splice(0)) cleanup();
    vi.useRealTimers();
  });

  async function owned(path: string, deck: PresentationModel) {
    const mounted = await mountAt(path, deck);
    cleanups.unshift(() => mounted.wrapper.unmount());
    return mounted;
  }
  async function at(ms: number) {
    const { setClock } = await import("../lib/core/clock");
    setClock(new Date(T0 + ms));
  }
  /** A presenter window, as the audience sees it: a channel end that records. */
  function fakeOwner(deckId: string) {
    const port = new BroadcastChannel(`docu-edit:present:${deckId}`);
    const heard: Array<{ kind: string; action?: string }> = [];
    port.addEventListener("message", (event) => heard.push(event.data));
    cleanups.push(() => port.close());
    return { heard, post: (message: unknown) => port.postMessage(message) };
  }
  it("shows a read-only HUD timer after play-next when the window owns the show, started with it", async () => {
    const deck = await freshDeck();
    const { wrapper } = await owned(`/slides/${deck.id}/present/1`, deck);
    const hud = wrapper.get(".play-hud");
    const timer = hud.get('[data-testid="play-timer"]');
    expect(timer.text()).toBe("0:00");
    expect(timer.attributes("data-state")).toBe("running");
    expect(timer.attributes("data-mode")).toBe("stopwatch");
    expect(hud.get('[data-testid="play-timer-toggle"]').text()).toBe("Pause (T)");
    // Placed after play-next, inside the fading HUD.
    const order = hud.findAll("[data-testid]").map((node) => node.attributes("data-testid"));
    expect(order.indexOf("play-timer")).toBeGreaterThan(order.indexOf("play-next"));
    expect(wrapper.find('[data-testid="play-timer-mode-stopwatch"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="play-timer-length"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="play-timer-status"]').text()).toBe("");
    // The record this tab keeps for a reload.
    expect(JSON.parse(sessionStorage.getItem(`docu-edit:show-timer:${deck.id}`) ?? "null")).toEqual({ startedAtMs: T0, accumulatedMs: 0 });
  });

  it("t, s and r drive it, and a timer key reveals the HUD", async () => {
    const deck = await freshDeck();
    const { wrapper } = await owned(`/slides/${deck.id}/present/1`, deck);
    const state = () => wrapper.get('[data-testid="play-timer"]').attributes("data-state");
    expect(wrapper.get(".play-hud").classes()).not.toContain("is-visible");
    await at(75_000);
    key("t");
    await flushPromises();
    expect(state()).toBe("paused");
    expect(wrapper.get('[data-testid="play-timer"]').text()).toBe("1:15");
    expect(wrapper.get('[data-testid="play-timer-status"]').text()).toBe("Timer paused");
    expect(wrapper.get(".play-hud").classes()).toContain("is-visible");
    key("t");
    key("s");
    await flushPromises();
    expect(state()).toBe("stopped");
    expect(wrapper.get('[data-testid="play-timer-stop"]').attributes("disabled")).toBeDefined();
    key("T"); // Shift+t too
    await flushPromises();
    expect(state()).toBe("running");
    key("r");
    await flushPromises();
    expect(state()).toBe("stopped");
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("1 / 4"); // none of them navigated
    key("t", { ctrlKey: true }); // a modifier combo is not a timer key
    await flushPromises();
    expect(state()).toBe("stopped");
  });

  it("a click on the toggle or Stop does not advance", async () => {
    const deck = await freshDeck();
    const { wrapper } = await owned(`/slides/${deck.id}/present/1`, deck);
    await wrapper.get('[data-testid="play-timer-toggle"]').trigger("click");
    expect(wrapper.get('[data-testid="play-timer"]').attributes("data-state")).toBe("stopped"); // a 0 ms pause
    await wrapper.get('[data-testid="play-timer-toggle"]').trigger("click");
    await wrapper.get('[data-testid="play-timer-stop"]').trigger("click");
    await flushPromises();
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("1 / 4");
    expect(wrapper.get('[data-testid="presentation-play"]').attributes("data-building")).toBe("false");
  });

  it("on the end screen t is not a timer key: it exits, as any key does", async () => {
    const deck = await freshDeck();
    const { wrapper, replace } = await owned(`/slides/${deck.id}/present/4`, deck);
    key("ArrowRight");
    await flushPromises();
    expect(wrapper.find('[data-testid="play-end-screen"]').exists()).toBe(true);
    replace.mockClear();
    key("t");
    await flushPromises();
    expect(replace).toHaveBeenCalledWith({ path: `/slides/${deck.id}`, query: { return: "4" } });
  });

  it("renders Countdown and overtime from the preferences", async () => {
    vi.useFakeTimers({ toFake: ["setInterval", "clearInterval"] });
    localStorage.setItem("docu-edit:preferences", JSON.stringify({ presentationTimerMode: "countdown", presentationTimerCountdownMinutes: 2 }));
    const deck = await freshDeck();
    const { wrapper } = await owned(`/slides/${deck.id}/present/1`, deck);
    const timer = () => wrapper.get('[data-testid="play-timer"]');
    expect(timer().text()).toBe("2:00");
    expect(timer().attributes("data-mode")).toBe("countdown");
    expect(timer().attributes("aria-label")).toBe("Time remaining");
    await at(75_000);
    vi.advanceTimersByTime(1000);
    await flushPromises();
    expect(timer().text()).toBe("0:45");
    expect(timer().attributes("data-overtime")).toBe("false");
    await at(162_000);
    vi.advanceTimersByTime(1000);
    await flushPromises();
    expect(timer().text()).toBe("−0:42");
    expect(timer().attributes("data-overtime")).toBe("true");
    expect(timer().attributes("aria-label")).toBe("Time over");
    expect(wrapper.get('[data-testid="play-timer-status"]').text()).toBe("Time is up");
    key("s");
    await flushPromises();
    expect(timer().text()).toBe("2:00");
    expect(timer().attributes("data-state")).toBe("stopped");
  });

  it("a following window has no timer; its t, s and r go to the presenter as intents", async () => {
    const deck = await freshDeck();
    const owner = fakeOwner(deck.id);
    const { wrapper } = await owned(`/slides/${deck.id}/present/1?follow=${SESSION}`, deck);
    owner.post({ kind: "state", session: SESSION, state: { seq: 1, slideIndex: 0, buildIndex: 0, atEnd: false, blank: "none" } });
    await settle();
    expect(wrapper.findAll('[data-testid="play-timer"]')).toHaveLength(0);
    key("t");
    key("s");
    key("r");
    await settle();
    expect(owner.heard.filter((message) => message.kind === "intent")).toEqual([
      { kind: "intent", session: SESSION, action: "timerToggle" },
      { kind: "intent", session: SESSION, action: "timerStop" },
      { kind: "intent", session: SESSION, action: "timerStop" },
    ]);
    expect(wrapper.get(".play-hud").classes()).not.toContain("is-visible"); // nothing of its own to show
    expect(sessionStorage.getItem(`docu-edit:show-timer:${deck.id}`)).toBeNull(); // never touches the record
  });

  it("on promotion the HUD timer appears with the presenter's timer, adopted as it was", async () => {
    const deck = await freshDeck();
    const owner = fakeOwner(deck.id);
    const { wrapper, replace } = await owned(`/slides/${deck.id}/present/1?follow=${SESSION}`, deck);
    owner.post({ kind: "state", session: SESSION, state: { seq: 1, slideIndex: 1, buildIndex: 0, atEnd: false, blank: "none" } });
    owner.post({ kind: "timer", session: SESSION, timer: { startedAtMs: null, accumulatedMs: 75_000 } });
    await settle();
    expect(wrapper.findAll('[data-testid="play-timer"]')).toHaveLength(0);
    await at(600_000);
    owner.post({ kind: "bye", session: SESSION, from: "presenter" });
    await settle();
    expect(replace).toHaveBeenLastCalledWith(`/slides/${deck.id}/present/2`);
    const timer = wrapper.get('[data-testid="play-timer"]');
    expect(timer.text()).toBe("1:15");
    expect(timer.attributes("data-state")).toBe("paused");
    expect(wrapper.get('[data-testid="play-timer-status"]').text()).toBe(""); // adopting announces nothing
    key("t"); // its own now
    await flushPromises();
    expect(wrapper.get('[data-testid="play-timer"]').attributes("data-state")).toBe("running");
  });

  it("a promoted window that never heard a timer begins one, like a fresh show", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    const deck = await freshDeck();
    const { PRESENTER_LOST_MS } = await import("../lib/presentation/presenterSession");
    const { wrapper } = await owned(`/slides/${deck.id}/present/1?follow=${SESSION}`, deck);
    expect(wrapper.findAll('[data-testid="play-timer"]')).toHaveLength(0);
    vi.advanceTimersByTime(PRESENTER_LOST_MS);
    await flushPromises();
    const timer = wrapper.get('[data-testid="play-timer"]');
    expect(timer.attributes("data-state")).toBe("running");
    expect(timer.text()).toBe("0:00");
  });

  it("unmounting removes the record; a reload (no unmount) would have found it", async () => {
    const deck = await freshDeck();
    const { wrapper } = await owned(`/slides/${deck.id}/present/1`, deck);
    const keyName = `docu-edit:show-timer:${deck.id}`;
    expect(sessionStorage.getItem(keyName)).not.toBeNull();
    wrapper.unmount();
    expect(sessionStorage.getItem(keyName)).toBeNull();
  });

  it("restores this tab's record on mount: a running timer counts the reload, auto-start never overrides a paused one", async () => {
    const deck = await freshDeck();
    sessionStorage.setItem(`docu-edit:show-timer:${deck.id}`, JSON.stringify({ startedAtMs: T0 - 30_000, accumulatedMs: 0 }));
    const running = await owned(`/slides/${deck.id}/present/1`, deck);
    expect(running.wrapper.get('[data-testid="play-timer"]').text()).toBe("0:30");
    expect(running.wrapper.get('[data-testid="play-timer"]').attributes("data-state")).toBe("running");
    running.wrapper.unmount();
    const other = await freshDeck();
    sessionStorage.setItem(`docu-edit:show-timer:${other.id}`, JSON.stringify({ startedAtMs: null, accumulatedMs: 90_000 }));
    const paused = await owned(`/slides/${other.id}/present/1`, other);
    expect(paused.wrapper.get('[data-testid="play-timer"]').text()).toBe("1:30");
    expect(paused.wrapper.get('[data-testid="play-timer"]').attributes("data-state")).toBe("paused");
  });
});
