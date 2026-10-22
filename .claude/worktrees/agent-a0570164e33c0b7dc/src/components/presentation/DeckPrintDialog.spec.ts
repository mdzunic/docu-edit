// SPEC-104 — the print-and-export dialog's control-enablement matrix, its
// tabs, the Download menu's two new entries and the off-screen sheet's
// layout (print vs export magnification). The dialog is mounted against a
// real useDeckExport controller over a deep-frozen deck.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import { effectScope, nextTick, ref, shallowRef } from "vue";
import DeckPrintDialog from "./DeckPrintDialog.vue";
import PresentationDownloadButton from "./PresentationDownloadButton.vue";
import { sheetPageViews } from "./deckPrintSheet";
import { useDeckExport } from "../../views/presentation/useDeckExport";
import { DEFAULT_DECK_PRINT_OPTIONS, planSheet } from "../../lib/presentation/printLayout";
import type { PresentationModel, Slide } from "../../lib/presentation/model";

function deck(count: number, hidden: number[] = []): PresentationModel {
  const slides: Slide[] = Array.from({ length: count }, (_, index) =>
    Object.freeze(hidden.includes(index) ? { id: `s${index}`, frames: [], hidden: true as const } : { id: `s${index}`, frames: [] }),
  );
  return Object.freeze({
    kind: "presentation",
    id: "d",
    title: "Pitch",
    format: "pptx",
    slides: Object.freeze(slides),
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
  }) as unknown as PresentationModel;
}

function mountDialog(model: PresentationModel) {
  const controller = effectScope().run(() =>
    useDeckExport({
      model: shallowRef<PresentationModel | null>(model),
      currentIndex: ref(0),
      slideTotal: ref(model.slides.length),
      reportError: () => undefined,
    }),
  )!;
  controller.layout.value = "slides";
  controller.range.value = "all";
  controller.rangeText.value = "";
  controller.includeHidden.value = false;
  controller.noteLines.value = false;
  controller.scale.value = 2;
  controller.open("pdf");
  const wrapper = mount(DeckPrintDialog, { props: { controller } });
  return { wrapper, controller };
}

const byId = (wrapper: ReturnType<typeof mountDialog>["wrapper"], id: string) => wrapper.get(`[data-testid="${id}"]`);
const has = (wrapper: ReturnType<typeof mountDialog>["wrapper"], id: string) => wrapper.find(`[data-testid="${id}"]`).exists();
const input = (wrapper: ReturnType<typeof mountDialog>["wrapper"], id: string) =>
  byId(wrapper, id).element as HTMLInputElement | HTMLSelectElement;

describe("DeckPrintDialog", () => {
  it("renders every control natively, with the defaults: Slides, A4 portrait, frame off, hidden off, 2×", () => {
    const { wrapper } = mountDialog(deck(3, [1]));
    expect(has(wrapper, "deck-print-dialog")).toBe(true);
    for (const id of ["deck-print-range-all", "deck-print-range-current", "deck-print-range-custom"]) {
      expect(input(wrapper, id).type).toBe("radio");
    }
    expect((input(wrapper, "deck-print-range-all") as HTMLInputElement).checked).toBe(true);
    expect(input(wrapper, "deck-print-range-input").tagName).toBe("INPUT");
    expect(input(wrapper, "deck-print-layout").tagName).toBe("SELECT");
    expect(
      Array.from((input(wrapper, "deck-print-layout") as HTMLSelectElement).options).map((option) => option.text),
    ).toEqual([
      "Slides",
      "Notes pages",
      "Handout 1 per page",
      "Handout 2 per page",
      "Handout 3 per page",
      "Handout 4 per page",
      "Handout 6 per page",
      "Handout 9 per page",
    ]);
    expect(input(wrapper, "deck-print-paper").value).toBe("A4");
    expect((input(wrapper, "deck-print-portrait") as HTMLInputElement).checked).toBe(true);
    expect((input(wrapper, "deck-print-frame") as HTMLInputElement).checked).toBe(false);
    expect((input(wrapper, "deck-print-hidden") as HTMLInputElement).checked).toBe(false);
    expect(input(wrapper, "deck-print-scale").value).toBe("2");
    expect(Array.from((input(wrapper, "deck-print-scale") as HTMLSelectElement).options).map((o) => o.text)).toEqual(["1×", "2×", "4×"]);
    expect(has(wrapper, "deck-print-print")).toBe(true);
    expect(has(wrapper, "deck-print-pdf")).toBe(true);
    expect(has(wrapper, "deck-print-cancel")).toBe(true);
  });

  it("the note-lines / frame matrix follows the layout", async () => {
    const { wrapper } = mountDialog(deck(3));
    const rows: [string, boolean, boolean, boolean][] = [];
    for (const layout of ["slides", "notes", "handout-1", "handout-2", "handout-3", "handout-4", "handout-6", "handout-9"]) {
      await byId(wrapper, "deck-print-layout").setValue(layout);
      const lines = input(wrapper, "deck-print-note-lines") as HTMLInputElement;
      const frame = input(wrapper, "deck-print-frame") as HTMLInputElement;
      rows.push([layout, lines.disabled, lines.checked, frame.checked]);
    }
    expect(rows).toEqual([
      ["slides", true, false, false],
      ["notes", true, false, true],
      ["handout-1", false, false, true],
      ["handout-2", false, false, true],
      ["handout-3", true, true, true],
      ["handout-4", false, false, true],
      ["handout-6", true, false, true],
      ["handout-9", true, false, true],
    ]);
  });

  it("paper and orientation apply to every layout but Slides", async () => {
    const { wrapper } = mountDialog(deck(3));
    expect(input(wrapper, "deck-print-paper").disabled).toBe(true);
    expect(input(wrapper, "deck-print-landscape").disabled).toBe(true);
    await byId(wrapper, "deck-print-layout").setValue("handout-6");
    expect(input(wrapper, "deck-print-paper").disabled).toBe(false);
    expect(input(wrapper, "deck-print-landscape").disabled).toBe(false);
  });

  it("Include hidden slides is disabled when the deck has none", () => {
    expect(input(mountDialog(deck(3)).wrapper, "deck-print-hidden").disabled).toBe(true);
    expect(input(mountDialog(deck(3, [2])).wrapper, "deck-print-hidden").disabled).toBe(false);
  });

  it("a bad or empty range shows deck-print-range-error and disables every action", async () => {
    const { wrapper, controller } = mountDialog(deck(3, [1]));
    expect(wrapper.find('[data-testid="deck-print-range-error"]').exists()).toBe(false);
    await byId(wrapper, "deck-print-range-input").setValue("7");
    expect(controller.range.value).toBe("custom");
    expect(byId(wrapper, "deck-print-range-error").text()).toBe("Slides go from 1 to 3");
    expect(byId(wrapper, "deck-print-range-error").attributes("role")).toBe("alert");
    expect(byId(wrapper, "deck-print-pdf").attributes("disabled")).toBeDefined();
    expect(byId(wrapper, "deck-print-print").attributes("disabled")).toBeDefined();
    await byId(wrapper, "deck-print-range-input").setValue("2"); // only the hidden slide
    expect(byId(wrapper, "deck-print-range-error").text()).toContain("hidden");
    await byId(wrapper, "deck-print-hidden").setValue(true);
    expect(wrapper.find('[data-testid="deck-print-range-error"]').exists()).toBe(false);
    expect(byId(wrapper, "deck-print-pdf").attributes("disabled")).toBeUndefined();
  });

  it("switches tabs: Images shows the scale, its estimate and the images action", async () => {
    const { wrapper, controller } = mountDialog(deck(3));
    await byId(wrapper, "deck-print-tab-images").trigger("click");
    expect(controller.tab.value).toBe("images");
    expect(byId(wrapper, "deck-print-tab-images").attributes("aria-selected")).toBe("true");
    expect(byId(wrapper, "deck-print-estimate").text()).toBe("3 images, 2560 × 1440 px each");
    await byId(wrapper, "deck-print-scale").setValue("4");
    expect(controller.scale.value).toBe(4);
    expect(has(wrapper, "deck-print-images")).toBe(true);
    expect(wrapper.find('[data-testid="deck-print-pdf"]').exists()).toBe(false);
    // Arrow keys move between the tabs.
    await byId(wrapper, "deck-print-tab-images").trigger("keydown", { key: "ArrowLeft" });
    expect(controller.tab.value).toBe("pdf");
  });

  it("shows the rotation note when the range holds a rotated frame", async () => {
    const rotated = deck(2);
    const model = Object.freeze({
      ...rotated,
      slides: [
        rotated.slides[0],
        { id: "r", frames: [{ id: "f", kind: "text", xCm: 1, yCm: 1, widthCm: 4, heightCm: 2, rotationDeg: 15, paragraphs: [] }] },
      ],
    }) as unknown as PresentationModel;
    const { wrapper, controller } = mountDialog(model);
    expect(has(wrapper, "deck-print-rotation-note")).toBe(true);
    controller.range.value = "current";
    await nextTick();
    expect(wrapper.find('[data-testid="deck-print-rotation-note"]').exists()).toBe(false);
  });

  it("shows progress with role=status and an abort control while a run is busy", async () => {
    const { wrapper, controller } = mountDialog(deck(3));
    controller.busy.value = true;
    controller.progress.value = { done: 2, total: 5 };
    await nextTick();
    expect(byId(wrapper, "deck-print-progress").text()).toBe("Page 2 of 5");
    expect(byId(wrapper, "deck-print-progress").attributes("role")).toBe("status");
    expect(has(wrapper, "deck-print-abort")).toBe(true);
    controller.busy.value = false;
  });

  it("Close closes the dialog", async () => {
    const { wrapper, controller } = mountDialog(deck(3));
    await byId(wrapper, "deck-print-cancel").trigger("click");
    expect(controller.isOpen.value).toBe(false);
  });
});

describe("PresentationDownloadButton (SPEC-104 entries)", () => {
  it("offers PDF (.pdf)… and Images (.png)… beside the copy and HTML entries", async () => {
    const wrapper = mount(PresentationDownloadButton, { props: { otherFormatLabel: "ODF Presentation (.odp)" } });
    await wrapper.get('[data-testid="presentation-download-menu"]').trigger("click");
    const items = wrapper.findAll('[role="menuitem"]').map((item) => item.attributes("data-testid"));
    expect(items).toEqual(["download-copy-as", "presentation-export-html", "presentation-export-pdf", "presentation-export-images"]);
    expect(wrapper.get('[data-testid="presentation-export-pdf"]').text()).toBe("PDF (.pdf)…");
    expect(wrapper.get('[data-testid="presentation-export-images"]').text()).toBe("Images (.png)…");
    await wrapper.get('[data-testid="presentation-export-pdf"]').trigger("click");
    expect(wrapper.emitted("export-pdf")).toHaveLength(1);
    await wrapper.get('[data-testid="presentation-download-menu"]').trigger("click");
    await wrapper.get('[data-testid="presentation-export-images"]').trigger("click");
    expect(wrapper.emitted("export-images")).toHaveLength(1);
  });
});

describe("sheetPageViews", () => {
  const model = deck(4);
  const pages = planSheet(model, { ...DEFAULT_DECK_PRINT_OPTIONS, layout: "handout-4", frameSlides: true });

  it("print: real paper size, each slide stage scaled into its slot", () => {
    const [page] = sheetPageViews(model, pages, "print", new Date(0));
    expect(page.style.width).toBe(`${21 * (96 / 2.54)}px`);
    expect(page.slots).toHaveLength(4);
    expect(page.slots[0].stageScale).toBeCloseTo(pages[0].slots[0].scale, 9);
    expect(page.frames).toHaveLength(4);
    expect(page.pageNumber?.text).toBe("1");
    expect(page.pageNumber?.style.fontSize).toBe(`${9 * (96 / 72)}px`);
  });

  it("export: the page magnified so every slide renders at scale 1, every length (fonts included) alike", () => {
    const [print] = sheetPageViews(model, pages, "print", new Date(0));
    const [exported] = sheetPageViews(model, pages, "export", new Date(0));
    const magnify = 1 / pages[0].slots[0].scale;
    expect(exported.slots[0].stageScale).toBeCloseTo(1, 9);
    expect(parseFloat(exported.style.width)).toBeCloseTo(parseFloat(print.style.width) * magnify, 6);
    expect(parseFloat(exported.pageNumber!.style.fontSize)).toBeCloseTo(9 * (96 / 72) * magnify, 6);
    // The slot box is exactly the slide's own CSS size.
    expect(parseFloat(exported.slots[0].style.width)).toBeCloseTo(33.867 * (96 / 2.54), 6);
  });
});
