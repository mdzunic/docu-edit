// SPEC-100 — PresentationPlayView as the presenter view's AUDIENCE window
// (?follow=<session>), driven by a fake owner on Node's real BroadcastChannel.
// The audience says hello, renders only the states that arrive (setStep), turns
// every local key/click/link into an intent, paints play-blank, keeps the
// full-screen prompt (and `f`) for the gesture a popup never had, re-reads the
// store once on the first state, plays the incoming transition for a one-step
// advance, and promotes itself to plain playback when the owner goes quiet.
// The plain (unfollowing) behaviour stays covered, unchanged, by
// PresentationPlayView.spec.ts and PresentationPlayView.transitions.spec.ts.
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { flushPromises, mount, type VueWrapper } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { routes } from "../router";
import type { PresentationModel } from "../lib/presentation/model";
import type { PresenterMessage, PresenterState } from "../lib/presentation/presenterSession";

beforeEach(() => {
  vi.resetModules();
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    getRecord: vi.fn().mockResolvedValue(undefined),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
    loadAllTemplateRecords: vi.fn().mockResolvedValue([]),
    putTemplateRecord: vi.fn().mockResolvedValue(undefined),
    deleteTemplateRecord: vi.fn().mockResolvedValue(undefined),
  }));
});

const pendingNavigations: Promise<unknown>[] = [];
const cleanups: Array<() => void> = [];

afterEach(async () => {
  for (const cleanup of cleanups.splice(0)) cleanup();
  await Promise.allSettled(pendingNavigations);
  pendingNavigations.length = 0;
  vi.doUnmock("../lib/idb");
  vi.restoreAllMocks();
  vi.useRealTimers();
});

const SESSION = "follow-test";

async function freshDeck(): Promise<PresentationModel> {
  const { buildPptxFixtureBytes } = await import("../lib/presentation/fixtures/build");
  const { readPresentationBytes } = await import("../lib/presentation/open");
  return readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "mini-deck");
}

/** The presenter, as seen from the audience: a channel end that records what
 * the audience posts and sends it states. */
function fakeOwner(deckId: string) {
  const port = new BroadcastChannel(`docu-edit:present:${deckId}`);
  const heard: PresenterMessage[] = [];
  port.addEventListener("message", (event) => heard.push(event.data as PresenterMessage));
  let seq = 0;
  cleanups.push(() => port.close());
  return {
    heard,
    state(over: Partial<PresenterState> = {}) {
      seq += 1;
      port.postMessage({
        kind: "state",
        session: SESSION,
        state: { seq, slideIndex: 0, buildIndex: 0, atEnd: false, blank: "none", ...over },
      });
    },
    post(message: unknown) {
      port.postMessage(message);
    },
  };
}

async function mountAudience(path: string, model?: PresentationModel) {
  const store = await import("../lib/documentStore");
  if (model) store.storePresentation(model);
  const { default: PresentationPlayView } = await import("./PresentationPlayView.vue");
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(path);
  await router.isReady();
  const realReplace = router.replace.bind(router);
  const replace = vi.spyOn(router, "replace").mockImplementation((to) => {
    const nav = realReplace(to);
    pendingNavigations.push(nav.catch(() => undefined));
    return nav;
  });
  const wrapper: VueWrapper = mount(PresentationPlayView, { global: { plugins: [router] } });
  cleanups.unshift(() => wrapper.unmount());
  await flushPromises();
  await flushPromises();
  return { wrapper, router, replace, store };
}

/** Real BroadcastChannel messages cross on the event loop. */
const settle = async () => {
  await new Promise((resolve) => setTimeout(resolve, 30));
  await flushPromises();
};
const counter = (wrapper: VueWrapper) => wrapper.get('[data-testid="play-slide-counter"]').text();
function key(name: string) {
  window.dispatchEvent(new KeyboardEvent("keydown", { key: name }));
}

describe("PresentationPlayView as the audience window (?follow=)", () => {
  it("says hello and renders the owner's state through setStep, keeping ?follow= as it tracks", async () => {
    const deck = await freshDeck();
    const owner = fakeOwner(deck.id);
    const { wrapper, replace } = await mountAudience(`/slides/${deck.id}/present/1?follow=${SESSION}`, deck);
    await settle();
    expect(owner.heard).toContainEqual({ kind: "hello", session: SESSION, from: "audience" });
    owner.state({ slideIndex: 2 });
    await settle();
    expect(counter(wrapper)).toBe("3 / 4");
    expect(replace).toHaveBeenCalledWith({ path: `/slides/${deck.id}/present/3`, query: { follow: SESSION } });
    // A foreign session's state is ignored.
    owner.post({ kind: "state", session: "other", state: { seq: 99, slideIndex: 0, buildIndex: 0, atEnd: false, blank: "none" } });
    await settle();
    expect(counter(wrapper)).toBe("3 / 4");
  });

  it("forwards keys, clicks and slide links as intents and renders only what comes back", async () => {
    const deck = await freshDeck();
    const first = deck.slides[0].frames.find((frame) => frame.kind === "text");
    if (first?.kind === "text") first.paragraphs = [{ runs: [{ text: "Jump " }, { text: "here", linkSlide: 4 }] }];
    const owner = fakeOwner(deck.id);
    const { wrapper, replace } = await mountAudience(`/slides/${deck.id}/present/1?follow=${SESSION}`, deck);
    owner.state();
    await settle();
    replace.mockClear();
    key("ArrowRight");
    key("b");
    key("4");
    key("Enter");
    await wrapper.get('[data-testid="presentation-play"]').trigger("click");
    await wrapper.get('[data-testid="slide-link"]').trigger("click");
    await wrapper.get('[data-testid="play-prev"]').trigger("click");
    key("Escape");
    await settle();
    const intents = owner.heard.filter((message) => message.kind === "intent");
    expect(intents).toEqual([
      { kind: "intent", session: SESSION, action: "advance" },
      { kind: "intent", session: SESSION, action: "blackout" },
      { kind: "intent", session: SESSION, action: "goTo", slideIndex: 3 },
      { kind: "intent", session: SESSION, action: "advance" },
      { kind: "intent", session: SESSION, action: "goTo", slideIndex: 3 },
      { kind: "intent", session: SESSION, action: "back" },
      { kind: "intent", session: SESSION, action: "exit" },
    ]);
    // Nothing moved locally, and Escape did not leave the show on its own.
    expect(counter(wrapper)).toBe("1 / 4");
    expect(replace).not.toHaveBeenCalled();
  });

  it("paints play-blank over the stage, aria-hidden, and announces it", async () => {
    const deck = await freshDeck();
    const owner = fakeOwner(deck.id);
    const { wrapper } = await mountAudience(`/slides/${deck.id}/present/1?follow=${SESSION}`, deck);
    owner.state({ blank: "black" });
    await settle();
    const blank = wrapper.get('[data-testid="play-blank"]');
    expect(blank.attributes("aria-hidden")).toBe("true");
    expect(blank.attributes("data-blank")).toBe("black");
    expect(wrapper.get('[data-testid="play-live"]').text()).toBe("Screen blanked");
    owner.state({ blank: "white" });
    await settle();
    expect(wrapper.get('[data-testid="play-blank"]').classes()).toContain("is-white");
    owner.state({ blank: "none", slideIndex: 1 });
    await settle();
    expect(wrapper.find('[data-testid="play-blank"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="play-live"]').text()).toBe("Slide 2 of 4 — Roadmap");
  });

  it("keeps the full-screen prompt and the drag hint up, and f / the prompt request full screen", async () => {
    const deck = await freshDeck();
    fakeOwner(deck.id);
    const request = vi.fn(() => Promise.resolve());
    Object.defineProperty(HTMLElement.prototype, "requestFullscreen", { value: request, configurable: true });
    cleanups.push(() => {
      delete (HTMLElement.prototype as { requestFullscreen?: unknown }).requestFullscreen;
    });
    const { wrapper } = await mountAudience(`/slides/${deck.id}/present/1?follow=${SESSION}`, deck);
    const prompt = wrapper.get('[data-testid="play-fullscreen-prompt"]');
    expect(prompt.element.tagName).toBe("BUTTON");
    expect(wrapper.get('[data-testid="play-screen-hint"]').text()).toBe(
      "Drag this window to your projector, then press F for full screen",
    );
    const onLoad = request.mock.calls.length; // the load-time request, as SPEC-084 makes it
    expect(onLoad).toBe(1);
    await prompt.trigger("click");
    key("f");
    expect(request).toHaveBeenCalledTimes(onLoad + 2);
    expect(counter(wrapper)).toBe("1 / 4"); // the prompt click did not advance
  });

  it("drops the drag hint for a window the presenter already placed on another screen", async () => {
    const deck = await freshDeck();
    fakeOwner(deck.id);
    const { wrapper } = await mountAudience(`/slides/${deck.id}/present/1?follow=${SESSION}&placed=1`, deck);
    expect(wrapper.find('[data-testid="play-fullscreen-prompt"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="play-screen-hint"]').exists()).toBe(false);
  });

  it("plays the incoming transition for a one-step advance, and none for a jump", async () => {
    const deck = await freshDeck();
    deck.slides[1].transition = { type: "push", direction: "left", durationMs: 1000 };
    deck.slides[3].transition = { type: "fade", durationMs: 1000 };
    const owner = fakeOwner(deck.id);
    const { wrapper } = await mountAudience(`/slides/${deck.id}/present/1?follow=${SESSION}`, deck);
    owner.state({ slideIndex: 0 });
    await settle();
    owner.state({ slideIndex: 1 });
    await settle();
    expect(wrapper.get('[data-testid="slide-transition-layer"]').attributes("data-effect")).toBe("push");
    owner.state({ slideIndex: 3 }); // a jump: instant, and the in-flight one is collapsed
    await settle();
    expect(wrapper.find('[data-testid="slide-transition-layer"]').exists()).toBe(false);
    expect(counter(wrapper)).toBe("4 / 4");
  });

  it("re-reads the store once, on the owner's first state, and never stores", async () => {
    const deck = await freshDeck();
    const owner = fakeOwner(deck.id);
    const store = await import("../lib/documentStore");
    const refresh = vi.spyOn(store, "refreshDocumentFromStorage");
    const storeSpy = vi.spyOn(store, "storePresentation");
    store.storePresentation(deck);
    storeSpy.mockClear();
    await mountAudience(`/slides/${deck.id}/present/1?follow=${SESSION}`);
    expect(refresh).not.toHaveBeenCalled(); // it waits for the owner
    owner.state({ slideIndex: 1 });
    await settle();
    owner.state({ slideIndex: 2 });
    await settle();
    expect(refresh).toHaveBeenCalledTimes(1);
    expect(refresh).toHaveBeenCalledWith(deck.id);
    expect(storeSpy).not.toHaveBeenCalled();
  });

  it("promotes itself to plain playback when the presenter says bye", async () => {
    const deck = await freshDeck();
    const owner = fakeOwner(deck.id);
    const { wrapper, replace } = await mountAudience(`/slides/${deck.id}/present/1?follow=${SESSION}`, deck);
    owner.state({ slideIndex: 1 });
    await settle();
    owner.post({ kind: "bye", session: SESSION, from: "presenter" });
    await settle();
    expect(replace).toHaveBeenLastCalledWith(`/slides/${deck.id}/present/2`); // ?follow= drops
    expect(owner.heard.at(-1)).toEqual({ kind: "bye", session: SESSION, from: "audience" });
    key("ArrowRight"); // plain SPEC-084 playback now
    await flushPromises();
    expect(counter(wrapper)).toBe("3 / 4");
    key("Escape");
    await flushPromises();
    expect(replace).toHaveBeenLastCalledWith({ path: `/slides/${deck.id}`, query: { return: "3" } });
  });

  it("a stale ?follow= URL opened alone promotes after PRESENTER_LOST_MS and plays on its own", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    const deck = await freshDeck();
    const { PRESENTER_LOST_MS } = await import("../lib/presentation/presenterSession");
    const { wrapper, replace } = await mountAudience(`/slides/${deck.id}/present/2?follow=${SESSION}`, deck);
    expect(counter(wrapper)).toBe("2 / 4"); // it shows its route's slide meanwhile
    key("ArrowRight"); // forwarded into the void: nothing moves
    await flushPromises();
    expect(counter(wrapper)).toBe("2 / 4");
    vi.advanceTimersByTime(PRESENTER_LOST_MS);
    await flushPromises();
    expect(replace).toHaveBeenLastCalledWith(`/slides/${deck.id}/present/2`);
    key("ArrowRight");
    await flushPromises();
    expect(counter(wrapper)).toBe("3 / 4");
  });
});

describe("presenting is inert (SPEC-100)", () => {
  it("no presenter or audience module dirties, edits, locks or casts", () => {
    const here = dirname(fileURLToPath(import.meta.url));
    const files = [
      "PresenterView.vue",
      "PresentationPlayView.vue",
      "presentation/usePresenterView.ts",
      "presentation/usePresenterSession.ts",
      "presentation/usePresenterTools.ts",
      "presentation/useAudienceWindow.ts",
      "presentation/usePlayInput.ts",
      "../components/presentation/PresenterSlidePane.vue",
      "../components/presentation/PresenterNotes.vue",
      "../components/presentation/PresenterGrid.vue",
    ];
    for (const file of files) {
      const source = readFileSync(join(here, file), "utf8").replace(/\/\/.*$/gm, "").replace(/\/\*[\s\S]*?\*\//g, "");
      for (const forbidden of ["acquireEditLock", "markSlideEdited", "notesDirty", "structureDirty", "dirty.value", "navigator.presentation"]) {
        expect(source.includes(forbidden), `${file} must not use ${forbidden}`).toBe(false);
      }
    }
  });
});
