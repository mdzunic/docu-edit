// SPEC-123 Format Cells: the six tabs; it opens on the active cell's
// effective style; the live preview and the Number tab's sample; a Custom
// code parseFormatCode rejects shows the error and disables OK; OK applies
// only the fields the user changed (a mixed selection is never flattened).
import { mount } from "@vue/test-utils";
import { describe, expect, it, vi } from "vitest";
import { DEFAULT_THEME } from "../../lib/core/theme";
import type { StylePatchEntry } from "../../lib/spreadsheet/formatting";
import type { CellStyle } from "../../lib/spreadsheet/model";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { useFormatCellsDialog } from "../../views/spreadsheet/useFormatCellsDialog";
import FormatCellsDialog from "./FormatCellsDialog.vue";

const EN_US = SPREADSHEET_LOCALES["en-US"];

function open(style: CellStyle, value: number | string = 0.125) {
  const apply = vi.fn<(entries: StylePatchEntry[]) => void>();
  const dialog = useFormatCellsDialog({
    target: () => ({
      style,
      value,
      locale: EN_US,
      format: { locale: EN_US, dateSystem: "1900" },
      theme: DEFAULT_THEME,
      documentFonts: ["Calibri"],
      styles: [{}, style],
      ranges: [{ top: 0, left: 0, bottom: 2, right: 1 }],
      recentColors: [],
    }),
    apply,
  });
  dialog.show();
  const wrapper = mount(FormatCellsDialog, { props: { dialog }, attachTo: document.body });
  return { dialog, wrapper, apply };
}

describe("SPEC-123 FormatCellsDialog", () => {
  it("has the six tabs and opens on the active cell's effective style", async () => {
    const { wrapper } = open({ numberFormat: "0.00%", font: { bold: true, sizePt: 14 } });
    expect(wrapper.findAll('[role="tab"]').map((tab) => tab.attributes("data-testid"))).toEqual(
      ["number", "alignment", "font", "border", "fill", "protection"].map((tab) => `format-cells-tab-${tab}`),
    );
    expect((wrapper.get('[data-testid="format-cells-category"]').element as HTMLSelectElement).value).toBe("percentage");
    expect(wrapper.get('[data-testid="format-cells-sample"]').text()).toBe("12.50%");
    await wrapper.get('[data-testid="format-cells-tab-font"]').trigger("click");
    expect((wrapper.get('[data-testid="format-cells-font-style"]').element as HTMLSelectElement).value).toBe("bold");
    expect((wrapper.get('[data-testid="format-cells-font-size"]').element as HTMLInputElement).value).toBe("14");
    wrapper.unmount();
  });

  it("previews a Custom code as typed and refuses one parseFormatCode rejects", async () => {
    const { wrapper } = open({});
    await wrapper.get('[data-testid="format-cells-category"]').setValue("custom");
    await wrapper.get('[data-testid="format-cells-custom-code"]').setValue("0.0%");
    expect(wrapper.get('[data-testid="format-cells-sample"]').text()).toBe("12.5%");
    expect(wrapper.get('[data-testid="format-cells-preview"]').text()).toBe("12.5%");
    await wrapper.get('[data-testid="format-cells-custom-code"]').setValue('0.0"');
    expect(wrapper.get('[data-testid="format-cells-custom-error"]').text()).toBe("This number format isn't valid");
    expect(wrapper.get('[data-testid="format-cells-ok"]').attributes("disabled")).toBeDefined();
    wrapper.unmount();
  });

  it("applies only the changed fields as one patch per range, plus border entries", async () => {
    const { wrapper, apply } = open({ font: { bold: true }, fill: { kind: "solid", color: "#ff0000" } });
    await wrapper.get('[data-testid="format-cells-tab-alignment"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-wrap"]').setValue(true);
    await wrapper.get('[data-testid="format-cells-tab-border"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-border-preset-inside"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-ok"]').trigger("click");
    const entries = apply.mock.calls[0]![0];
    expect(entries[0]).toEqual({ range: { top: 0, left: 0, bottom: 2, right: 1 }, patch: { alignment: { wrap: true } } });
    expect(entries.slice(1).map((entry) => Object.keys(entry.patch.borders ?? {})[0])).toEqual(["bottom", "top", "right", "left"]);
    wrapper.unmount();
  });

  // A colour or border edge the dialog holds must stay the plain object the
  // picker emitted: held in a deep `ref` it came back out as a Vue reactive
  // Proxy, which patchStyle's clone (and IndexedDB's) refuses — the preview
  // threw instead of showing the colour, and OK applied nothing.
  it("previews and applies theme colours and a coloured outline as plain data", async () => {
    const { wrapper, apply } = open({ font: { bold: true } }, "Total");
    const accent = { theme: "accent1", lumMod: 0.6, lumOff: 0.4 };
    await wrapper.get('[data-testid="format-cells-tab-fill"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-fill-theme-swatch"][aria-label="Accent 1, lighter 40%"]').trigger("click");
    const preview = wrapper.get('[data-testid="format-cells-preview"]');
    expect(preview.attributes("style")).toContain("background");
    await wrapper.get('[data-testid="format-cells-tab-font"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-font-color"] [data-testid="text-color-button"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-font-color"] [data-testid="text-color-theme-swatch"][aria-label="Accent 2"]').trigger("click");
    expect(preview.get("span").attributes("style")).toContain("color");
    await wrapper.get('[data-testid="format-cells-tab-border"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-border-color"] [data-testid="text-color-button"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-border-color"] [data-testid="text-color-theme-swatch"][aria-label="Accent 6"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-border-preset-outline"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-ok"]').trigger("click");
    const entries = apply.mock.calls[0]![0];
    // What the model receives must survive the structured clone IndexedDB
    // persistence runs on it.
    expect(() => structuredClone(entries)).not.toThrow();
    expect(entries[0]).toEqual({ range: { top: 0, left: 0, bottom: 2, right: 1 }, patch: { font: { color: { theme: "accent2" } }, fill: { kind: "solid", color: accent } } });
    expect(entries.find((entry) => entry.patch.borders?.top)?.patch.borders).toEqual({ top: { style: "thin", color: { theme: "accent6" } } });
    wrapper.unmount();
  });
});
