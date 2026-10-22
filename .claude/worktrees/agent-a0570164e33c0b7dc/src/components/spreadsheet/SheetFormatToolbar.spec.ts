// SPEC-123 SheetFormatToolbar: role="toolbar" named "Sheet toolbar", the
// eight bands in order with 1px rules between them, ONE roving tab stop,
// aria-pressed from the active cell's state, the commands it runs, the
// painter's click/double-click, the number-format items with their previews
// and band 6's slot for SPEC-124's merge.
import { mount } from "@vue/test-utils";
import { describe, expect, it, vi } from "vitest";
import { computed, ref } from "vue";
import { DEFAULT_THEME } from "../../lib/core/theme";
import type { SheetToolbarState } from "../../views/spreadsheet/sheetToolbar";
import SheetFormatToolbar from "./SheetFormatToolbar.vue";

function state(partial: Partial<SheetToolbarState> = {}): SheetToolbarState {
  return {
    disabled: false,
    fontFamily: "Calibri",
    fontSizePt: 11,
    documentFonts: ["Calibri"],
    bold: true,
    italic: false,
    underline: false,
    strikethrough: false,
    textColor: "",
    theme: DEFAULT_THEME,
    recentColors: [],
    fillColor: "#ffff00",
    borderLineStyle: "thin",
    borderLineColor: "",
    horizontal: null,
    vertical: "bottom",
    wrap: false,
    textRotation: 0,
    indent: 0,
    numberFormatLabel: "General",
    numberFormats: [{ id: "currency", label: "Currency", preview: "$1,234.50" }],
    cellStyles: [{ label: "Good, Bad and Neutral", styles: [{ name: "Good", box: {}, text: {} }] }],
    ...partial,
  };
}

function mountToolbar() {
  const run = vi.fn();
  const painter = { mode: ref(""), disabled: computed(() => false), click: vi.fn(), dblclick: vi.fn() };
  const wrapper = mount(SheetFormatToolbar, { props: { formatting: { toolbar: computed(() => state()), run }, painter }, attachTo: document.body });
  return { wrapper, run, painter };
}

describe("SPEC-123 SheetFormatToolbar", () => {
  it("is the named toolbar with the eight bands in order and one tab stop", async () => {
    // jsdom lays nothing out: every element's offsetParent is null, which the
    // roving tab stop reads as hidden. Lay the controls out for this test.
    const offsetParent = Object.getOwnPropertyDescriptor(HTMLElement.prototype, "offsetParent")!;
    Object.defineProperty(HTMLElement.prototype, "offsetParent", { configurable: true, get: () => document.body });
    const { wrapper } = mountToolbar();
    const strip = wrapper.get('[data-testid="sheet-format-toolbar"]');
    expect(strip.attributes("role")).toBe("toolbar");
    expect(strip.attributes("aria-label")).toBe("Sheet toolbar");
    expect(strip.findAll("[data-band]").map((band) => band.attributes("aria-label"))).toEqual([
      "History",
      "Type",
      "Character",
      "Colour",
      "Number",
      "Alignment",
      "Indent & orientation",
      "Styles",
    ]);
    expect(strip.find('[data-band="6"] [data-slot="sheet-merge"]').exists()).toBe(true);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(strip.findAll("button, input, select").filter((control) => control.attributes("tabindex") === "0")).toHaveLength(1);
    wrapper.unmount();
    Object.defineProperty(HTMLElement.prototype, "offsetParent", offsetParent);
  });

  it("shows the active cell's state and runs the commands", async () => {
    const { wrapper, run, painter } = mountToolbar();
    expect(wrapper.get('[data-testid="sheet-bold"]').attributes("aria-pressed")).toBe("true");
    expect(wrapper.get('[data-testid="sheet-bold"]').attributes("title")).toBe("Bold (Ctrl+B)");
    expect(wrapper.get('[data-testid="sheet-italic"]').attributes("aria-pressed")).toBe("false");
    await wrapper.get('[data-testid="sheet-italic"]').trigger("click");
    await wrapper.get('[data-testid="sheet-align-center"]').trigger("click");
    await wrapper.get('[data-testid="sheet-decimals-increase"]').trigger("click");
    await wrapper.get('[data-testid="sheet-fill-button"]').trigger("click");
    expect(run.mock.calls.map((call) => call[0])).toEqual([
      { kind: "toggle", format: "italic" },
      { kind: "align", horizontal: "center" },
      { kind: "decimals", delta: 1 },
      { kind: "fill-last" },
    ]);
    await wrapper.get('[data-testid="sheet-format-painter"]').trigger("click");
    await wrapper.get('[data-testid="sheet-format-painter"]').trigger("dblclick");
    expect(painter.click).toHaveBeenCalledOnce();
    expect(painter.dblclick).toHaveBeenCalledOnce();
    await wrapper.get('[data-testid="sheet-number-format"]').trigger("click");
    expect(document.querySelector('[data-testid="sheet-number-format-currency"]')?.textContent).toContain("$1,234.50");
    wrapper.unmount();
  });
});
