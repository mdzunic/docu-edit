import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import PrintPreviewDialog from "./PrintPreviewDialog.vue";

function mountPanel(props: Record<string, unknown> = {}) {
  return mount(PrintPreviewDialog, {
    props: {
      pageCount: 3,
      rangeMode: "all",
      rangeText: "",
      error: "",
      canPrint: true,
      ...props,
    },
  });
}

describe("PrintPreviewDialog", () => {
  it("renders a dialog with the page count and range controls (AC-2, AC-4)", () => {
    const wrapper = mountPanel();
    const panel = wrapper.get('[data-testid="print-preview"]');
    expect(panel.attributes("role")).toBe("dialog");
    expect(wrapper.get('[data-testid="print-page-count"]').text()).toBe("3 pages");
    expect(wrapper.get('[data-testid="print-range-all"]').attributes("type")).toBe("radio");
    expect(wrapper.get('[data-testid="print-range-custom"]').attributes("type")).toBe("radio");
  });

  it("uses the singular for one page", () => {
    expect(mountPanel({ pageCount: 1 }).get('[data-testid="print-page-count"]').text()).toBe(
      "1 page",
    );
  });

  it("shows the custom range input only in custom mode, emitting edits", async () => {
    const wrapper = mountPanel();
    expect(wrapper.find('[data-testid="print-range-input"]').exists()).toBe(false);
    await wrapper.get('[data-testid="print-range-custom"]').trigger("change");
    expect(wrapper.emitted("update:range-mode")).toEqual([["custom"]]);
    await wrapper.setProps({ rangeMode: "custom" });
    const input = wrapper.get('[data-testid="print-range-input"]');
    await input.setValue("1-2");
    expect(wrapper.emitted("update:range-text")).toEqual([["1-2"]]);
  });

  it("shows the validation error and disables Print when printing is not allowed (AC-28)", async () => {
    const wrapper = mountPanel({
      rangeMode: "custom",
      rangeText: "9",
      error: "Pages go from 1 to 3",
      canPrint: false,
    });
    const error = wrapper.get('[data-testid="print-range-error"]');
    expect(error.attributes("role")).toBe("alert");
    expect(error.text()).toBe("Pages go from 1 to 3");
    expect(
      (wrapper.get('[data-testid="print-confirm"]').element as HTMLButtonElement).disabled,
    ).toBe(true);
    await wrapper.get('[data-testid="print-confirm"]').trigger("click");
    expect(wrapper.emitted("print")).toBeUndefined();
  });

  it("emits print from the Print button when allowed (AC-5)", async () => {
    const wrapper = mountPanel();
    await wrapper.get('[data-testid="print-confirm"]').trigger("click");
    expect(wrapper.emitted("print")).toHaveLength(1);
  });

  it("emits close from the close button, Cancel and Escape", async () => {
    const wrapper = mountPanel();
    await wrapper.get('[data-testid="print-preview-close"]').trigger("click");
    await wrapper.get('[data-testid="print-cancel"]').trigger("click");
    await wrapper.get('[data-testid="print-preview"]').trigger("keydown.escape");
    expect(wrapper.emitted("close")).toHaveLength(3);
  });

  it("reports no pages for an empty document and disables the range control (AC-27)", () => {
    const wrapper = mountPanel({ pageCount: 0, canPrint: false });
    expect(wrapper.get('[data-testid="print-page-count"]').text()).toBe("0 pages");
    expect(wrapper.get('[data-testid="print-no-pages"]').text()).toContain("no content to print");
    expect(
      (wrapper.get(".panel-range").element as HTMLFieldSetElement).disabled,
    ).toBe(true);
  });
});
