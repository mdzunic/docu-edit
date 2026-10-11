// SPEC-123 Format Cells: the six tabs; it opens on the active cell's
// effective style; the live preview (the pending edges included) and the
// Number tab's sample; a Custom code parseFormatCode rejects shows the error
// and disables OK; OK applies only the fields the user changed (a mixed
// selection is never flattened), a picked colour as plain data.
import { mount } from "@vue/test-utils";
import { describe, expect, it, vi } from "vitest";
import { TEXT_COLOR_PALETTE } from "../../lib/core/colors";
import { DEFAULT_THEME, resolveColor, type ThemeColorRef } from "../../lib/core/theme";
import { LINE_STYLES } from "../../lib/spreadsheet/formatCells";
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
    const { wrapper, dialog, apply } = open({ font: { bold: true } }, "Total");
    const accent: ThemeColorRef = { theme: "accent1", lumMod: 0.6, lumOff: 0.4 };
    await wrapper.get('[data-testid="format-cells-tab-fill"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-fill-theme-swatch"][aria-label="Accent 1, lighter 40%"]').trigger("click");
    expect(dialog.preview.value.box["background-color"]).toBe(resolveColor(accent, DEFAULT_THEME));
    expect(wrapper.get('[data-testid="format-cells-preview"]').attributes("style")).toContain("background-color");
    await wrapper.get('[data-testid="format-cells-tab-font"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-font-color"] [data-testid="text-color-button"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-font-color"] [data-testid="text-color-theme-swatch"][aria-label="Accent 2"]').trigger("click");
    expect(dialog.preview.value.textCss.color).toBe(resolveColor({ theme: "accent2" }, DEFAULT_THEME));
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

  it("shows the active cell's colours and fills from the standard palette too", async () => {
    const red = TEXT_COLOR_PALETTE[1]!;
    const { wrapper, dialog, apply } = open({ font: { color: { theme: "accent2" } }, fill: { kind: "pattern", pattern: "gray125", fg: "#00ff00" } });
    await wrapper.get('[data-testid="format-cells-tab-font"]').trigger("click");
    expect(wrapper.get('[data-testid="format-cells-font-color"] [data-testid="text-color-current"]').attributes("data-color")).toBe(resolveColor({ theme: "accent2" }, DEFAULT_THEME));
    await wrapper.get('[data-testid="format-cells-tab-fill"]').trigger("click");
    expect(wrapper.get('[data-testid="format-cells-pattern-color"] [data-testid="text-color-current"]').attributes("data-color")).toBe("#00ff00");
    const swatches = wrapper.findAll('[data-testid="format-cells-fill-swatch"]');
    expect(swatches.map((swatch) => swatch.attributes("data-color"))).toEqual([...TEXT_COLOR_PALETTE]);
    await wrapper.get('[data-testid="format-cells-pattern-style"]').setValue("solid");
    await swatches[1]!.trigger("click");
    expect(dialog.preview.value.box["background-color"]).toBe(resolveColor(red, DEFAULT_THEME));
    await wrapper.get('[data-testid="format-cells-ok"]').trigger("click");
    expect(apply.mock.calls[0]![0]).toEqual([{ range: { top: 0, left: 0, bottom: 2, right: 1 }, patch: { fill: { kind: "solid", color: red } } }]);
    wrapper.unmount();
  });

  it("draws the pending outer edges in the preview box; the None line style removes edges", async () => {
    const { wrapper, dialog, apply } = open({ borders: { top: { style: "thin" } } });
    await wrapper.get('[data-testid="format-cells-tab-border"]').trigger("click");
    const style = wrapper.get('[data-testid="format-cells-border-style"]');
    expect(style.findAll("option").map((option) => option.attributes("value"))).toEqual(["none", ...LINE_STYLES]);
    expect(dialog.preview.value.box["border-top"]).toBe("1px solid #000000");
    await style.setValue("double");
    await wrapper.get('[data-testid="format-cells-border-edge-bottom"]').trigger("click");
    expect(dialog.preview.value.box["border-bottom"]).toBe("3px double #000000");
    await style.setValue("none");
    await wrapper.get('[data-testid="format-cells-border-edge-top"]').trigger("click");
    expect(wrapper.get('[data-testid="format-cells-border-edge-top"]').attributes("aria-pressed")).toBe("false");
    expect(dialog.preview.value.box["border-top"]).toBeUndefined();
    await wrapper.get('[data-testid="format-cells-ok"]').trigger("click");
    const entries = apply.mock.calls[0]![0].filter((entry) => !entry.storedOnly);
    expect(entries.map((entry) => entry.patch.borders)).toEqual([{ top: null }, { bottom: { style: "double" } }]);
    wrapper.unmount();
  });
});

// SPEC-129 "The rule editor": the same dialog, opened over a Dxf instead of a
// cell's style.
describe("SPEC-129 FormatCellsDialog differential mode", () => {
  function openDifferential(initial: Record<string, unknown> = {}) {
    const apply = vi.fn<(entries: StylePatchEntry[]) => void>();
    const onSave = vi.fn<(dxf: unknown) => void>();
    const dialog = useFormatCellsDialog({ target: () => null, apply });
    dialog.showDifferential(initial as never, { theme: DEFAULT_THEME, locale: EN_US, documentFonts: ["Calibri"] }, onSave);
    const wrapper = mount(FormatCellsDialog, { props: { dialog }, attachTo: document.body });
    return { dialog, wrapper, apply, onSave };
  }

  it("shows only Number, Font, Border and Fill, and hides font family/size", async () => {
    const { wrapper } = openDifferential();
    expect(wrapper.findAll('[role="tab"]').map((tab) => tab.attributes("data-testid"))).toEqual(
      ["number", "font", "border", "fill"].map((tab) => `format-cells-tab-${tab}`),
    );
    await wrapper.get('[data-testid="format-cells-tab-font"]').trigger("click");
    expect(wrapper.find('[data-testid="format-cells-font-family"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="format-cells-font-size"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="format-cells-font-style"]').exists()).toBe(true);
    await wrapper.get('[data-testid="format-cells-tab-border"]').trigger("click");
    expect(wrapper.findAll('[data-testid^="format-cells-border-edge-"]').map((edge) => edge.attributes("data-testid"))).toEqual([
      "format-cells-border-edge-top",
      "format-cells-border-edge-bottom",
      "format-cells-border-edge-left",
      "format-cells-border-edge-right",
    ]);
    wrapper.unmount();
  });

  it("seeds the Font tab from an already-set Dxf and preserves it when OK is pressed untouched", async () => {
    const { wrapper, onSave } = openDifferential({ font: { color: "#9C0006", bold: true } });
    await wrapper.get('[data-testid="format-cells-tab-font"]').trigger("click");
    expect((wrapper.get('[data-testid="format-cells-font-style"]').element as HTMLSelectElement).value).toBe("bold");
    await wrapper.get('[data-testid="format-cells-ok"]').trigger("click");
    expect(onSave).toHaveBeenCalledWith({ font: { color: "#9C0006", bold: true, italic: false } });
    wrapper.unmount();
  });

  it("OK hands the rebuilt Dxf to onSave instead of calling apply", async () => {
    const { wrapper, apply, onSave } = openDifferential();
    await wrapper.get('[data-testid="format-cells-tab-fill"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-fill-swatch"]').trigger("click");
    await wrapper.get('[data-testid="format-cells-ok"]').trigger("click");
    expect(apply).not.toHaveBeenCalled();
    expect(onSave).toHaveBeenCalledTimes(1);
    expect(onSave.mock.calls[0]![0]).toMatchObject({ fill: { kind: "solid" } });
    wrapper.unmount();
  });

  it("Cancel closes without saving", async () => {
    const { wrapper, onSave } = openDifferential();
    await wrapper.get('[data-testid="format-cells-cancel"]').trigger("click");
    expect(onSave).not.toHaveBeenCalled();
    wrapper.unmount();
  });
});
