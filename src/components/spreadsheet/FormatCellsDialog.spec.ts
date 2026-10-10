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
});
