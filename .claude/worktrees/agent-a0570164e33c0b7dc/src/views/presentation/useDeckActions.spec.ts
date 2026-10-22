// SPEC-084 — the deck header's playback-entry actions (the download pair is
// covered by SPEC-080's view spec). Present starts at the current slide, is
// gated on a non-empty deck, has the Cmd/Ctrl+Shift+Return shortcut, and the
// `?return` handoff reselects the slide the play view was showing. SPEC-088
// adds the HTML export action (dialog gate + serialise-and-download).
import { describe, expect, it, vi } from "vitest";
import { ref } from "vue";
import type { Router, RouteLocationNormalizedLoaded } from "vue-router";
import { useDeckActions } from "./useDeckActions";
import type { PresentationModel } from "../../lib/presentation/model";

const triggerBlobDownload = vi.fn();
vi.mock("../../lib/core/triggerDownload", () => ({
  triggerBlobDownload: (...args: unknown[]) => triggerBlobDownload(...args),
}));

function harness(opts: { slides?: number; current?: number; query?: Record<string, string> } = {}) {
  const push = vi.fn();
  const router = { push } as unknown as Router;
  const route = { query: opts.query ?? {} } as unknown as RouteLocationNormalizedLoaded;
  const model = ref<PresentationModel | null>({ format: "pptx" } as PresentationModel);
  const actions = useDeckActions({
    model,
    presentationId: ref("deck-1"),
    currentIndex: ref(opts.current ?? 0),
    slideTotal: ref(opts.slides ?? 3),
    router,
    route,
    reportError: vi.fn(),
  });
  return { actions, push };
}

describe("useDeckActions playback entry", () => {
  it("canPresent is false only for an empty deck (AC-2)", () => {
    expect(harness({ slides: 0 }).actions.canPresent.value).toBe(false);
    expect(harness({ slides: 1 }).actions.canPresent.value).toBe(true);
  });

  it("present() starts playback at the current 1-based slide (AC-1)", () => {
    const { actions, push } = harness({ current: 2 });
    actions.present();
    expect(push).toHaveBeenCalledWith("/slides/deck-1/present/3");
  });

  it("present() from a hidden current slide pushes that very slide's route (SPEC-104 AC-11)", () => {
    const { actions, push } = harness({ current: 1 });
    // The model's hidden flags are the play view's concern: the deep link
    // always names the current slide, so presenting from a hidden one opens on it.
    actions.present();
    expect(push).toHaveBeenCalledWith("/slides/deck-1/present/2");
  });

  it("present() is a no-op on an empty deck", () => {
    const { actions, push } = harness({ slides: 0 });
    actions.present();
    expect(push).not.toHaveBeenCalled();
  });

  it("Cmd/Ctrl+Shift+Return triggers present, other combos do not (AC-1)", () => {
    const { actions, push } = harness({ current: 0 });
    const evt = new KeyboardEvent("keydown", { key: "Enter", ctrlKey: true, shiftKey: true });
    const prevent = vi.spyOn(evt, "preventDefault");
    actions.onPresentShortcut(evt);
    expect(push).toHaveBeenCalledWith("/slides/deck-1/present/1");
    expect(prevent).toHaveBeenCalled();

    push.mockClear();
    actions.onPresentShortcut(new KeyboardEvent("keydown", { key: "Enter", shiftKey: true })); // no modifier
    actions.onPresentShortcut(new KeyboardEvent("keydown", { key: "Enter", ctrlKey: true })); // no shift
    expect(push).not.toHaveBeenCalled();
  });

  it("returnIndex reads and clamps the ?return handoff (AC-35)", () => {
    expect(harness({ query: {} }).actions.returnIndex()).toBe(0); // no handoff → first
    expect(harness({ query: { return: "3" } }).actions.returnIndex()).toBe(2); // 1-based → 0-based
    expect(harness({ slides: 3, query: { return: "99" } }).actions.returnIndex()).toBe(2); // clamp high
    expect(harness({ query: { return: "0" } }).actions.returnIndex()).toBe(0); // clamp low
  });
});

function deck(): PresentationModel {
  return {
    kind: "presentation",
    id: "deck-1",
    title: "My Deck",
    format: "pptx",
    slideWidthCm: 25.4,
    slideHeightCm: 14.29,
    slides: [{ id: "s1", frames: [], notes: ["a note"] }],
  };
}

describe("useDeckActions HTML export (SPEC-088)", () => {
  function exportHarness(model: PresentationModel | null) {
    triggerBlobDownload.mockClear();
    return useDeckActions({
      model: ref(model),
      presentationId: ref("deck-1"),
      currentIndex: ref(0),
      slideTotal: ref(model ? model.slides.length : 0),
      router: { push: vi.fn() } as unknown as Router,
      route: { query: {} } as unknown as RouteLocationNormalizedLoaded,
      reportError: vi.fn(),
    });
  }

  it("opens the options dialog only for a non-empty deck", () => {
    const actions = exportHarness(deck());
    expect(actions.htmlDialogOpen.value).toBe(false);
    actions.openHtmlExport();
    expect(actions.htmlDialogOpen.value).toBe(true);

    const empty = exportHarness({ ...deck(), slides: [] });
    empty.openHtmlExport();
    expect(empty.htmlDialogOpen.value).toBe(false);
  });

  it("cancel closes the dialog without downloading", () => {
    const actions = exportHarness(deck());
    actions.openHtmlExport();
    actions.cancelHtmlExport();
    expect(actions.htmlDialogOpen.value).toBe(false);
    expect(triggerBlobDownload).not.toHaveBeenCalled();
  });

  it("confirm downloads <title>.html and honours the notes option", () => {
    const actions = exportHarness(deck());
    actions.openHtmlExport();
    actions.confirmHtmlExport({ includeNotes: true });
    expect(actions.htmlDialogOpen.value).toBe(false);
    expect(triggerBlobDownload).toHaveBeenCalledTimes(1);
    const [bytes, fileName, mime] = triggerBlobDownload.mock.calls[0];
    expect(fileName).toBe("My Deck.html");
    expect(mime).toContain("text/html");
    const html = new TextDecoder().decode(bytes as Uint8Array);
    expect(html).toContain("a note");
  });

  it("confirm with notes off omits the note text", () => {
    const actions = exportHarness(deck());
    actions.confirmHtmlExport({ includeNotes: false });
    const [bytes] = triggerBlobDownload.mock.calls[0];
    expect(new TextDecoder().decode(bytes as Uint8Array)).not.toContain("a note");
  });
});
