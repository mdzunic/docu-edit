// SPEC-085 — playback transition rendering layered onto SPEC-084. A forward
// step between slides plays the incoming slide's transition (two stacked
// stages); backward steps are instant; an input mid-transition completes it
// then applies as the next advance; auto-advance schedules and cleans up. The idb seam
// is mocked so the store runs in-memory (mirrors PresentationPlayView.spec.ts).
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

afterEach(() => {
  vi.doUnmock("../lib/idb");
  vi.restoreAllMocks();
});

async function deckWithTransitions(): Promise<PresentationModel> {
  const { buildPptxFixtureBytes } = await import("../lib/presentation/fixtures/build");
  const { readPresentationBytes } = await import("../lib/presentation/open");
  const model = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "mini-deck");
  model.slides[1].transition = { type: "push", direction: "left", durationMs: 1000 };
  model.slides[2].transition = { type: "fade", durationMs: 1000 };
  return model;
}

async function mountAt(path: string, model?: PresentationModel) {
  const store = await import("../lib/documentStore");
  if (model) store.storePresentation(model);
  const { default: PresentationPlayView } = await import("./PresentationPlayView.vue");
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push(path);
  await router.isReady();
  const wrapper = mount(PresentationPlayView, { global: { plugins: [router] } });
  await flushPromises();
  await flushPromises();
  return { wrapper, router, store };
}

function key(name: string) {
  window.dispatchEvent(new KeyboardEvent("keydown", { key: name }));
}

const LAYER = '[data-testid="slide-transition-layer"]';

describe("PresentationPlayView transitions", () => {
  it("plays the incoming slide's transition on a forward step (AC-13)", async () => {
    const deck = await deckWithTransitions();
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    expect(wrapper.find(LAYER).exists()).toBe(false);
    key("ArrowRight"); // → slide 2, which has a push transition
    await flushPromises();
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("2 / 4");
    const layer = wrapper.get(LAYER);
    expect(layer.attributes("data-effect")).toBe("push");
    // Two stacked stages render inside the letterbox.
    expect(wrapper.findAll('[data-testid="slide-stage"]').length).toBe(2);
  });

  it("completes the in-flight transition then applies the input as the next advance (AC-15)", async () => {
    const deck = await deckWithTransitions();
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    key("ArrowRight"); // → slide 2, push transition begins
    await flushPromises();
    expect(wrapper.find(LAYER).exists()).toBe(true);
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("2 / 4");
    key("ArrowRight"); // completes the push instantly, then advances to slide 3
    await flushPromises();
    // The input is honoured (not dropped): we reach slide 3, exactly one step on
    // from the settled slide 2 — no queued double-advance skips a slide.
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("3 / 4");
    // The push is gone; slide 3's fade is the only live animation.
    const layer = wrapper.get(LAYER);
    expect(layer.attributes("data-effect")).toBe("fade");
  });

  it("makes backward steps instant, with no transition (AC-14)", async () => {
    const deck = await deckWithTransitions();
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/2`, deck);
    key("ArrowLeft");
    await flushPromises();
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("1 / 4");
    expect(wrapper.find(LAYER).exists()).toBe(false);
  });

  it("carries build state into the transition: unbuilt frames stay hidden, exited frames stay gone (SPEC-086)", async () => {
    const deck = await deckWithTransitions();
    const exited = deck.slides[0].frames[0];
    const unbuilt = deck.slides[1].frames[0];
    deck.slides[0].animations = [{ frameId: exited.id, effect: "fadeOut", trigger: "click" }];
    deck.slides[1].animations = [{ frameId: unbuilt.id, effect: "fadeIn", trigger: "click" }];
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/2`, deck);
    key("ArrowLeft"); // back lands on slide 1 fully built — its exit has played
    await flushPromises();
    key("ArrowRight"); // → slide 2 (push), none of its builds played yet
    await flushPromises();
    const frameStyle = (stage: string, index: number): string =>
      wrapper.get(`[data-testid="${stage}"]`).findAll(".stage-surface > *")[index].attributes("style") ?? "";
    expect(frameStyle("transition-outgoing", 0)).toContain("visibility: hidden");
    expect(frameStyle("transition-outgoing", 1)).not.toContain("visibility: hidden");
    expect(frameStyle("transition-incoming", 0)).toContain("visibility: hidden");
    expect(frameStyle("transition-incoming", 1)).not.toContain("visibility: hidden");
  });

  it("keeps a number-jump instant (AC-14)", async () => {
    const deck = await deckWithTransitions();
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    key("3");
    key("Enter");
    await flushPromises();
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("3 / 4");
    expect(wrapper.find(LAYER).exists()).toBe(false);
  });
});

describe("PresentationPlayView auto-advance", () => {
  it("schedules an advance for a slide with advanceAfterMs (AC-17)", async () => {
    const { buildPptxFixtureBytes } = await import("../lib/presentation/fixtures/build");
    const { readPresentationBytes } = await import("../lib/presentation/open");
    const deck = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "mini-deck");
    deck.slides[0].transition = { type: "none", advanceAfterMs: 40 };
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("1 / 4");
    // Slide 1's 40ms auto-advance fires → slide 2 (real timer).
    await vi.waitFor(
      () => expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("2 / 4"),
      { timeout: 1000, interval: 20 },
    );
  });

  it("cleans up the auto-advance timer on manual navigation (AC-17)", async () => {
    const { buildPptxFixtureBytes } = await import("../lib/presentation/fixtures/build");
    const { readPresentationBytes } = await import("../lib/presentation/open");
    const deck = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "mini-deck");
    deck.slides[0].transition = { type: "none", advanceAfterMs: 10000 }; // long
    const { wrapper } = await mountAt(`/slides/${deck.id}/present/1`, deck);
    key("End"); // manual jump reschedules for the new slide, cancelling slide 1's
    await flushPromises();
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("4 / 4");
    // If slide 1's timer had leaked it would fire; it was cleared, so we stay.
    await new Promise((r) => setTimeout(r, 60));
    await flushPromises();
    expect(wrapper.get('[data-testid="play-slide-counter"]').text()).toBe("4 / 4");
  });
});
