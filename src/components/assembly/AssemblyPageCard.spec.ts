// SPEC-111: a page card — a paper-shaped placeholder at the page's real
// aspect with its number, source, size and rotation, saying it has no
// preview, its selected state, and its actions on the tray.
import { describe, expect, it, vi } from "vitest";
import { mount } from "@vue/test-utils";
import { ref } from "vue";
import AssemblyPageCard from "./AssemblyPageCard.vue";

function tray() {
  return {
    rotate: vi.fn(),
    duplicatePages: vi.fn(),
    removePages: vi.fn(),
    selectSource: vi.fn(),
    isPagesFull: ref(false),
  };
}

const page = {
  key: "tray-7",
  id: "pdf-1",
  title: "contract",
  page: { index: 1, count: 3, widthPt: 595.28, heightPt: 841.89, rotate: 90 },
  rotateBy: 90,
};

describe("AssemblyPageCard", () => {
  it("states what it knows of the page — and that it has no preview", () => {
    const wrapper = mount(AssemblyPageCard, { props: { entry: page, position: 4, selected: true, tray: tray() } });
    expect(wrapper.find('[data-testid="page-number"]').text()).toBe("Page 4");
    expect(wrapper.find('[data-testid="page-source"]').text()).toBe("contract · p. 2 of 3");
    expect(wrapper.find('[data-testid="page-size"]').text()).toBe("595 × 842 pt");
    expect(wrapper.find('[data-testid="page-rotation"]').text()).toBe("rotate 180");
    expect(wrapper.text()).toContain("No preview");
    expect(wrapper.find('[data-testid="page-selected"]').exists()).toBe(true);
    expect(wrapper.find("img").exists()).toBe(false);
    expect(wrapper.find("canvas").exists()).toBe(false);
    // Upright again after its two quarter turns: portrait paper.
    const paper = wrapper.find('[data-testid="page-paper"]');
    expect(Number(paper.attributes("data-aspect"))).toBeCloseTo(595.28 / 841.89, 3);
    expect(paper.text()).toBe("4");
  });

  it("turns, duplicates, deletes and selects its source through the tray", async () => {
    const controller = tray();
    const wrapper = mount(AssemblyPageCard, { props: { entry: { ...page, rotateBy: 0 }, position: 1, tray: controller } });
    expect(wrapper.find('[data-testid="page-selected"]').exists()).toBe(false);
    expect(Number(wrapper.find('[data-testid="page-paper"]').attributes("data-aspect"))).toBeCloseTo(841.89 / 595.28, 3);
    await wrapper.find('[data-testid="page-rotate-left"]').trigger("click");
    await wrapper.find('[data-testid="page-rotate-right"]').trigger("click");
    await wrapper.find('[data-testid="page-duplicate"]').trigger("click");
    await wrapper.find('[data-testid="page-delete"]').trigger("click");
    await wrapper.find('[data-testid="page-select-source"]').trigger("click");
    expect(controller.rotate.mock.calls).toEqual([
      ["tray-7", -90],
      ["tray-7", 90],
    ]);
    expect(controller.duplicatePages).toHaveBeenCalledWith("tray-7");
    expect(controller.removePages).toHaveBeenCalledWith("tray-7");
    expect(controller.selectSource).toHaveBeenCalledWith("tray-7");
  });

  it("reads a blank page at the sheet's size", () => {
    const wrapper = mount(AssemblyPageCard, {
      props: { entry: { key: "tray-9", id: "", title: "Blank page", blank: true }, position: 2, tray: tray(), sheetWidthPt: 595.28, sheetHeightPt: 841.89 },
    });
    expect(wrapper.find('[data-testid="page-source"]').text()).toBe("Blank page");
    expect(wrapper.find('[data-testid="page-size"]').text()).toBe("595 × 842 pt");
    expect(wrapper.find('[data-testid="page-select-source"]').exists()).toBe(false);
  });
});
