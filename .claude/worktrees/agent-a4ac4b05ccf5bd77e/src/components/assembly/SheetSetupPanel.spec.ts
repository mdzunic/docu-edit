// SPEC-108's shared sheet setup panel over a real useSheetSetup controller:
// unit conversion in and out, the margin link toggle and presets, the
// transparent-background refusals, error states, control enablement and the
// live estimate. (VTU's setValue on a text input fires input AND change, as
// committing a typed value does.)
import { beforeEach, describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import { defineComponent, h, ref } from "vue";
import SheetSetupPanel from "./SheetSetupPanel.vue";
import { resetSheetSetupSessions, useSheetSetup, type SheetSetupController } from "../../composables/useSheetSetup";
import { resetPreferences, savePreference } from "../../lib/preferences";
import type { SheetOutput } from "../../lib/core/sheetLayout";

function mountPanel(output: SheetOutput = "pdf", itemPx = { widthPx: 64, heightPx: 40 }) {
  let controller!: SheetSetupController;
  const Host = defineComponent({
    setup() {
      controller = useSheetSetup({ surface: `panel-spec-${output}`, output, itemPx: ref(itemPx) });
      return () => h(SheetSetupPanel, { controller });
    },
  });
  const wrapper = mount(Host);
  return { wrapper, controller: () => controller };
}

const byId = (id: string) => `[data-testid="${id}"]`;

beforeEach(() => {
  localStorage.clear();
  resetPreferences();
  resetSheetSetupSessions();
});

describe("SheetSetupPanel", () => {
  it("exposes every control by test id, each a native element", () => {
    const { wrapper } = mountPanel();
    const expected: Record<string, string> = {
      "sheet-paper": "SELECT",
      "sheet-portrait": "BUTTON",
      "sheet-landscape": "BUTTON",
      "sheet-width": "INPUT",
      "sheet-height": "INPUT",
      "sheet-margin-preset": "SELECT",
      "sheet-margin-top": "INPUT",
      "sheet-margin-right": "INPUT",
      "sheet-margin-bottom": "INPUT",
      "sheet-margin-left": "INPUT",
      "sheet-margin-link": "INPUT",
      "sheet-fit": "SELECT",
      "sheet-align": "SELECT",
      "sheet-background": "SELECT",
      "sheet-dpi": "SELECT",
      "sheet-estimate": "P",
    };
    for (const [id, tag] of Object.entries(expected)) {
      expect(wrapper.get(byId(id)).element.tagName, id).toBe(tag);
    }
    expect(wrapper.find(byId("sheet-error")).exists()).toBe(false);
  });

  it("shows lengths in the user's unit and reads typed values back into centimetres", async () => {
    savePreference("measurementUnit", "in");
    const { wrapper, controller } = mountPanel();
    const top = wrapper.get(byId("sheet-margin-top"));
    expect((top.element as HTMLInputElement).value).toBe("1");
    expect((wrapper.get(byId("sheet-width")).element as HTMLInputElement).value).toBe("8.27");
    await top.setValue("0.5");
    // Linked (Normal is equal all round): all four follow.
    expect(controller().setup.value.margins).toEqual({ topCm: 1.27, rightCm: 1.27, bottomCm: 1.27, leftCm: 1.27 });
    // An explicit suffix wins over the display unit.
    await top.setValue("2cm");
    expect(controller().setup.value.margins.topCm).toBe(2);
    expect((wrapper.get(byId("sheet-margin-left")).element as HTMLInputElement).value).toBe("0.79");
  });

  it("the link toggle keeps all four margins equal; unlinked, each edge is its own", async () => {
    const { wrapper, controller } = mountPanel();
    const link = wrapper.get(byId("sheet-margin-link"));
    expect((link.element as HTMLInputElement).checked).toBe(true);
    await link.setValue(false);
    const right = wrapper.get(byId("sheet-margin-right"));
    await right.setValue("3");
    expect(controller().setup.value.margins).toEqual({ topCm: 2.54, rightCm: 3, bottomCm: 2.54, leftCm: 2.54 });
    expect((wrapper.get(byId("sheet-margin-preset")).element as HTMLSelectElement).value).toBe("Custom");
    // Re-linking makes all four equal to the top margin.
    await link.setValue(true);
    expect(controller().setup.value.margins).toEqual({ topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 });
  });

  it("applies a margin preset; an unequal one unlinks", async () => {
    const { wrapper, controller } = mountPanel();
    await wrapper.get(byId("sheet-margin-preset")).setValue("Moderate");
    expect(controller().setup.value.margins).toEqual({ topCm: 2.54, rightCm: 1.91, bottomCm: 2.54, leftCm: 1.91 });
    expect((wrapper.get(byId("sheet-margin-link")).element as HTMLInputElement).checked).toBe(false);
    await wrapper.get(byId("sheet-margin-preset")).setValue("Narrow");
    expect(controller().setup.value.margins.leftCm).toBe(1.27);
  });

  it("rejects margins that leave less than 1 cm, with the reason, and never clamps", async () => {
    const { wrapper, controller } = mountPanel();
    const top = wrapper.get(byId("sheet-margin-top"));
    // A4 is 21 cm across: 10 + 10 leaves exactly 1 cm — accepted.
    await top.setValue("10");
    expect(controller().setup.value.margins.leftCm).toBe(10);
    await top.setValue("10.01");
    expect(controller().setup.value.margins.leftCm).toBe(10);
    const error = wrapper.get(byId("sheet-error"));
    expect(error.attributes("role")).toBe("alert");
    expect(error.text()).toContain("These margins leave 0.98 × 9.68 cm for the content on a 21 × 29.7 cm sheet");
    expect(top.attributes("aria-invalid")).toBe("true");
    // The typed value stays visible beside its reason until it is fixed.
    expect((top.element as HTMLInputElement).value).toBe("10.01");
  });

  it("rejects a negative or unreadable margin", async () => {
    const { wrapper, controller } = mountPanel();
    const bottom = wrapper.get(byId("sheet-margin-bottom"));
    await bottom.setValue("-1");
    expect(wrapper.get(byId("sheet-error")).text()).toBe("Margins cannot be negative.");
    await bottom.setValue("lots");
    expect(wrapper.get(byId("sheet-error")).text()).toContain('"lots" is not a length');
    expect(controller().setup.value.margins.bottomCm).toBe(2.54);
  });

  it("refuses a paper the margins cannot fit, and the select snaps back", async () => {
    const { wrapper, controller } = mountPanel();
    const top = wrapper.get(byId("sheet-margin-top"));
    await top.setValue("7");
    // A4 holds 7 cm all round (7 cm across is left); A5's 14.8 cm does not.
    const paper = wrapper.get(byId("sheet-paper"));
    await paper.setValue("A5");
    expect(controller().setup.value.paper).toBe("A4");
    expect((paper.element as HTMLSelectElement).value).toBe("A4");
    expect(wrapper.get(byId("sheet-error")).text()).toContain("on a 14.8 × 21 cm sheet");
    expect(paper.attributes("aria-invalid")).toBe("true");
    // An accepted change clears the reason.
    await paper.setValue("A3");
    expect(controller().setup.value.paper).toBe("A3");
    expect(wrapper.find(byId("sheet-error")).exists()).toBe(false);
  });

  it("enables width/height for Custom only, and orientation not for Fit to content", async () => {
    const { wrapper, controller } = mountPanel();
    const width = wrapper.get(byId("sheet-width"));
    expect(width.attributes("disabled")).toBeDefined();
    await wrapper.get(byId("sheet-paper")).setValue("Custom");
    expect(width.attributes("disabled")).toBeUndefined();
    await width.setValue("30");
    expect(controller().setup.value).toMatchObject({ paper: "Custom", widthCm: 30, heightCm: 29.7, orientation: "landscape" });
    expect(wrapper.get(byId("sheet-landscape")).attributes("aria-pressed")).toBe("true");
    await wrapper.get(byId("sheet-paper")).setValue("Fit to content");
    expect(wrapper.get(byId("sheet-portrait")).attributes("disabled")).toBeDefined();
    expect(width.attributes("disabled")).toBeDefined();
  });

  it("disables alignment for stretch, which fills the box exactly", async () => {
    const { wrapper } = mountPanel();
    expect(wrapper.get(byId("sheet-align")).attributes("disabled")).toBeUndefined();
    await wrapper.get(byId("sheet-fit")).setValue("stretch");
    expect(wrapper.get(byId("sheet-align")).attributes("disabled")).toBeDefined();
  });

  it("PDF: transparent is accepted as a choice but the panel says it will be white", async () => {
    const { wrapper, controller } = mountPanel("pdf");
    expect(controller().setup.value.background).toBe("#ffffff");
    expect(wrapper.find(byId("sheet-background-notice")).exists()).toBe(false);
    await wrapper.get(byId("sheet-background")).setValue("transparent");
    expect(controller().setup.value.background).toBe("transparent");
    expect(wrapper.get(byId("sheet-background-notice")).text()).toBe(
      "PDF pages have no transparency, so the background will be white.",
    );
  });

  it("JPEG: transparent is refused outright with the reason; PNG defaults to it", async () => {
    const jpeg = mountPanel("jpeg");
    await jpeg.wrapper.get(byId("sheet-background")).setValue("transparent");
    expect(jpeg.controller().setup.value.background).toBe("#ffffff");
    expect(jpeg.wrapper.get(byId("sheet-error")).text()).toContain("JPEG images have no transparency");
    expect((jpeg.wrapper.get(byId("sheet-background")).element as HTMLSelectElement).value).toBe("#ffffff");
    const png = mountPanel("png");
    expect(png.controller().setup.value.background).toBe("transparent");
  });

  it("a custom colour is edited through the colour input", async () => {
    const { wrapper, controller } = mountPanel();
    await wrapper.get(byId("sheet-background")).setValue("custom");
    const colour = wrapper.get(byId("sheet-background-color"));
    await colour.setValue("#336699");
    expect(controller().setup.value.background).toBe("#336699");
  });

  it("shows the estimate live: a 600 dpi A3 is a number before anything is written", async () => {
    const { wrapper } = mountPanel("png");
    await wrapper.get(byId("sheet-paper")).setValue("A3");
    await wrapper.get(byId("sheet-dpi")).setValue("600");
    expect(wrapper.get(byId("sheet-estimate")).text()).toContain("7016 × 9921 px at 600 dpi");
    expect(wrapper.get(byId("sheet-estimate")).text()).toContain("265.5 MB");
  });

  it("persists paper, margins and dpi as preferences", async () => {
    const { wrapper } = mountPanel();
    await wrapper.get(byId("sheet-paper")).setValue("Letter");
    await wrapper.get(byId("sheet-margin-preset")).setValue("Narrow");
    await wrapper.get(byId("sheet-dpi")).setValue("300");
    const stored = JSON.parse(localStorage.getItem("docu-edit:preferences") ?? "{}");
    expect(stored).toMatchObject({
      assemblyPaper: "Letter",
      assemblyMargins: { topCm: 1.27, rightCm: 1.27, bottomCm: 1.27, leftCm: 1.27 },
      assemblyDpi: 300,
    });
    resetSheetSetupSessions();
    const again = mountPanel();
    expect(again.controller().setup.value).toMatchObject({ paper: "Letter", dpi: 300, margins: { topCm: 1.27 } });
  });
});
