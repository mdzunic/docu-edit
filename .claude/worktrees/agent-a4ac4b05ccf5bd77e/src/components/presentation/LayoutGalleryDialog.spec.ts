// SPEC-089 layout gallery dialog: the card list (real thumbnails + the "new"
// badge), selection, the apply/close/reset controls and their disabled state,
// and the dismissible kept-in-place report.
import { mount } from "@vue/test-utils";
import { ref } from "vue";
import { describe, expect, it, vi } from "vitest";
import LayoutGalleryDialog from "./LayoutGalleryDialog.vue";
import { instantiateStandardLayout, STANDARD_LAYOUTS } from "../../lib/presentation/layouts";

function makeController(overrides = {}) {
  const twoContent = instantiateStandardLayout(STANDARD_LAYOUTS.find((d) => d.kind === "twoContent")!, 33.867, 19.05);
  const deckLayout = { key: "ppt/slideLayouts/slideLayout1.xml", kind: "title", name: "Deck Title", slots: [] };
  return {
    open: ref(true),
    mode: ref("change"),
    cards: ref([deckLayout, twoContent]),
    currentKey: ref(deckLayout.key),
    canReset: ref(true),
    report: ref(null),
    widthCm: ref(33.867),
    heightCm: ref(19.05),
    currentSlide: ref({ id: "s", frames: [] }),
    choose: vi.fn(),
    apply: vi.fn(),
    resetSlide: vi.fn(),
    close: vi.fn(),
    dismissReport: vi.fn(),
    ...overrides,
  };
}

function mountDialog(controller: ReturnType<typeof makeController>) {
  return mount(LayoutGalleryDialog, { props: { controller, slideNumber: 2 } });
}

describe("LayoutGalleryDialog", () => {
  it("renders one option card per layout with the 'new' badge on synthesised fallbacks", () => {
    const w = mountDialog(makeController());
    expect(w.find('[data-testid="slides-layout-gallery"]').exists()).toBe(true);
    const options = w.findAll('[role="option"]');
    expect(options).toHaveLength(2);
    expect(w.find('[data-testid="layout-card-0"]').exists()).toBe(true);
    expect(w.find('[data-testid="layout-card-1"]').exists()).toBe(true);
    // The deck's own layout carries no badge; the synthesised fallback does.
    const badges = w.findAll('[data-testid="layout-card-new-badge"]');
    expect(badges).toHaveLength(1);
    expect(w.find('[data-testid="layout-card-1"]').text()).toContain("will be added to this deck");
  });

  it("marks the current layout aria-selected and chooses on click", async () => {
    const c = makeController();
    const w = mountDialog(c);
    expect(w.find('[data-testid="layout-card-0"]').attributes("aria-selected")).toBe("true");
    await w.find('[data-testid="layout-card-1"]').trigger("click");
    expect(c.choose).toHaveBeenCalledWith((c.cards.value[1] as { key: string }).key);
  });

  it("wires apply / close / reset controls", async () => {
    const c = makeController();
    const w = mountDialog(c);
    await w.find('[data-testid="layout-apply"]').trigger("click");
    expect(c.apply).toHaveBeenCalled();
    await w.find('[data-testid="layout-gallery-close"]').trigger("click");
    expect(c.close).toHaveBeenCalled();
    await w.find('[data-testid="layout-reset-slide"]').trigger("click");
    expect(c.resetSlide).toHaveBeenCalled();
  });

  it("disables Reset slide when the slide has no resolvable layout", () => {
    const c = makeController({ canReset: ref(false) });
    const w = mountDialog(c);
    expect(w.find('[data-testid="layout-reset-slide"]').attributes("aria-disabled")).toBe("true");
  });

  it("shows the kept-in-place report and dismisses it", async () => {
    const c = makeController({
      report: ref({ applied: true, placed: 1, created: 0, keptInPlace: ["img", "extra"] }),
      currentSlide: ref({
        id: "s",
        frames: [
          { id: "img", kind: "image", mediaId: "m", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1 },
          { id: "extra", kind: "text", role: "body", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, paragraphs: [] },
        ],
      }),
    });
    const w = mountDialog(c);
    const report = w.find('[data-testid="layout-report"]');
    expect(report.exists()).toBe(true);
    expect(report.text()).toContain("2 items kept in place");
    await w.find('[data-testid="layout-report-dismiss"]').trigger("click");
    expect(c.dismissReport).toHaveBeenCalled();
  });
});
