// SPEC-103 — the deck's version history panel: SPEC-085/086's panel contract
// (a controller in, nothing emitted), every testid the spec names, the
// newest-first list with author, slide count and slide delta, the actions
// routed to the controller, the paused notice, and the read-only preview on
// its own stage. Its typed transforms are pinned beside it.
import { afterEach, describe, expect, it, vi } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import { computed, ref, shallowRef } from "vue";
import DeckHistoryPanel from "./DeckHistoryPanel.vue";
import { downloadLabel, RESTORE_UNDO_NOTE, slideCountLabel, slideDeltaLabel, triggerLabel } from "./deckHistoryRender";
import type { DeckVersionListEntry } from "../../lib/presentationVersions";
import type { PresentationModel } from "../../lib/presentation/model";

const wrappers: VueWrapper[] = [];
afterEach(() => {
  for (const wrapper of wrappers.splice(0)) wrapper.unmount();
});

function entry(id: string, extra: Partial<DeckVersionListEntry> = {}): DeckVersionListEntry {
  return { id, timestamp: Date.UTC(2026, 9, 23, 9, 30), author: "Ada", words: 0, slides: 4, slideDelta: 0, kind: "delta", parentId: null, size: 10, ...extra };
}

function controller(entries: DeckVersionListEntry[] = []) {
  const versions = ref(entries);
  const namedOnly = ref(false);
  return {
    open: ref(true),
    format: ref("pptx"),
    versions,
    visibleVersions: computed(() => (namedOnly.value ? versions.value.filter((e) => e.name) : versions.value)),
    namedOnly,
    pausedMessage: ref(""),
    previewModel: shallowRef<PresentationModel | null>(null),
    previewVersionId: ref<string | null>(null),
    previewIndex: ref(0),
    preview: vi.fn(),
    exitPreview: vi.fn(),
    selectPreviewSlide: vi.fn(),
    restore: vi.fn(),
    nameVersion: vi.fn(),
    removeName: vi.fn(),
    deleteVersion: vi.fn(),
    downloadVersion: vi.fn(),
  };
}

function mountPanel(c: ReturnType<typeof controller>) {
  const wrapper = mount(DeckHistoryPanel, { props: { controller: c, slideNumber: 2 } });
  wrappers.push(wrapper);
  return wrapper;
}

describe("deckHistoryRender", () => {
  it("labels slide counts, deltas, triggers and the download format", () => {
    expect(slideCountLabel(1)).toBe("1 slide");
    expect(slideCountLabel(12)).toBe("12 slides");
    expect(slideDeltaLabel(0)).toBe("no slide change");
    expect(slideDeltaLabel(2)).toBe("+2 slides");
    expect(slideDeltaLabel(-1)).toBe("−1 slide");
    expect(triggerLabel(entry("a", { trigger: "pre-restore" }))).toBe("Before restore");
    expect(triggerLabel(entry("a", { trigger: "post-restore" }))).toBe("After restore");
    expect(triggerLabel(entry("a", { trigger: "cadence" }))).toBe("");
    expect(downloadLabel("pptx")).toBe("Download .pptx");
    expect(downloadLabel("odp")).toBe("Download .odp");
  });
});

describe("DeckHistoryPanel", () => {
  it("shows the empty state, the filter and the restore copy with no versions", () => {
    const wrapper = mountPanel(controller());
    expect(wrapper.find('[data-testid="slides-history-panel"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="history-empty"]').text()).toBe("No versions yet");
    expect(wrapper.find('[data-testid="history-list"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="history-named-only"]').exists()).toBe(true);
    expect(wrapper.text()).toContain(RESTORE_UNDO_NOTE);
    expect(wrapper.emitted()).toEqual({});
  });

  it("lists versions as given (newest first) with time, author, slide count and delta", () => {
    const wrapper = mountPanel(
      controller([entry("b", { slides: 6, slideDelta: 2, name: "Board review" }), entry("a", { slides: 4, slideDelta: 4, trigger: "pre-restore" })]),
    );
    const items = wrapper.findAll('[data-testid="history-version"]');
    expect(items).toHaveLength(2);
    expect(items[0].get('[data-testid="history-name"]').text()).toBe("Board review");
    expect(items[0].text()).toContain("Ada");
    expect(items[0].text()).toContain("6 slides · +2 slides");
    expect(items[1].get('[data-testid="history-name"]').text()).toBe("Unnamed version");
    expect(items[1].get('[data-testid="history-trigger"]').text()).toBe("Before restore");
    expect(items[0].get('[data-testid="history-download"]').text()).toBe("Download .pptx");
    // Remove name only where there is one.
    expect(items[0].find('[data-testid="history-remove-name"]').exists()).toBe(true);
    expect(items[1].find('[data-testid="history-remove-name"]').exists()).toBe(false);
  });

  it("routes every action to the controller", async () => {
    const c = controller([entry("v1", { name: "Kept" })]);
    const wrapper = mountPanel(c);
    await wrapper.get('[data-testid="history-preview"]').trigger("click");
    await wrapper.get('[data-testid="history-restore"]').trigger("click");
    await wrapper.get('[data-testid="history-download"]').trigger("click");
    await wrapper.get('[data-testid="history-remove-name"]').trigger("click");
    await wrapper.get('[data-testid="history-delete"]').trigger("click");
    await wrapper.get('[data-testid="history-name-input"]').setValue("Final cut");
    await wrapper.get(".history-name-form").trigger("submit");
    expect(c.preview).toHaveBeenCalledWith("v1");
    expect(c.restore).toHaveBeenCalledWith("v1");
    expect(c.downloadVersion).toHaveBeenCalledWith("v1");
    expect(c.removeName).toHaveBeenCalledWith("v1");
    expect(c.deleteVersion).toHaveBeenCalledWith("v1");
    expect(c.nameVersion).toHaveBeenCalledWith("v1", "Final cut");
  });

  it("filters to named versions", async () => {
    const c = controller([entry("b", { name: "Named" }), entry("a")]);
    const wrapper = mountPanel(c);
    await wrapper.get('[data-testid="history-named-only"]').setValue(true);
    expect(c.namedOnly.value).toBe(true);
    expect(wrapper.findAll('[data-testid="history-version"]')).toHaveLength(1);
  });

  it("shows the paused notice", () => {
    const c = controller();
    c.pausedMessage.value = "Version history off — this deck's package is larger than the 50 MB history budget";
    const wrapper = mountPanel(c);
    expect(wrapper.get('[data-testid="history-paused-notice"]').text()).toBe(c.pausedMessage.value);
  });

  it("previews on its own read-only stage with its own slide strip, and exits", async () => {
    const c = controller([entry("v1", { name: "Pitch" })]);
    c.previewModel.value = {
      kind: "presentation",
      id: "deck",
      title: "Then",
      format: "pptx",
      slideWidthCm: 25.4,
      slideHeightCm: 19.05,
      slides: [
        { id: "s1", frames: [] },
        { id: "s2", frames: [] },
      ],
    };
    c.previewVersionId.value = "v1";
    c.previewIndex.value = 1;
    const wrapper = mountPanel(c);
    const banner = wrapper.get('[data-testid="history-preview-banner"]');
    expect(banner.text()).toContain("Previewing Pitch — read-only");
    expect(banner.find('[data-testid="slide-stage"]').exists()).toBe(true);
    const strip = banner.findAll(".history-preview-slide");
    expect(strip.map((b) => b.attributes("aria-pressed"))).toEqual(["false", "true"]);
    await strip[0].trigger("click");
    expect(c.selectPreviewSlide).toHaveBeenCalledWith(0);
    await wrapper.get('[data-testid="history-exit-preview"]').trigger("click");
    expect(c.exitPreview).toHaveBeenCalled();
  });
});
