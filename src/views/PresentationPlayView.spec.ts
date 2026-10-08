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
