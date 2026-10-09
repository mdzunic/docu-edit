// SPEC-117 "Aggregates" and "Zoom" in the workbook's status bar: the Average,
// Count and Sum texts it is handed (each hidden when empty — a single cell or
// a selection of text shows less), and SPEC-063's ZoomControl at its right
// with its fit entries hidden.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import SpreadsheetStatusBar from "./SpreadsheetStatusBar.vue";
import ZoomControl from "../ZoomControl.vue";

describe("SpreadsheetStatusBar", () => {
  it("shows Average, Count and Sum for a multi-cell numeric selection", () => {
    const wrapper = mount(SpreadsheetStatusBar, {
      props: { average: "Average: 20", count: "Count: 3", sum: "Sum: 60", zoomPercent: 85 },
    });
    expect(wrapper.find('[data-testid="spreadsheet-status-bar"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="sheet-status-average"]').text()).toBe("Average: 20");
    expect(wrapper.get('[data-testid="sheet-status-count"]').text()).toBe("Count: 3");
    expect(wrapper.get('[data-testid="sheet-status-sum"]').text()).toBe("Sum: 60");
  });

  it("shows only Count when nothing numeric is selected, and nothing for a single cell", async () => {
    const wrapper = mount(SpreadsheetStatusBar, { props: { count: "Count: 2" } });
    expect(wrapper.find('[data-testid="sheet-status-average"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="sheet-status-sum"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="sheet-status-count"]').text()).toBe("Count: 2");
    await wrapper.setProps({ count: "" });
    expect(wrapper.find('[data-testid="sheet-status-count"]').exists()).toBe(false);
  });

  it("puts the zoom control at its right, at the sheet's zoom, without the fit entries", () => {
    const wrapper = mount(SpreadsheetStatusBar, { props: { count: "Count: 3", zoomPercent: 85 }, attachTo: document.body });
    const bar = wrapper.get('[data-testid="spreadsheet-status-bar"]').element;
    const zoom = wrapper.get('[data-testid="zoom-control"]').element;
    // The zoom control is the bar's last child: aggregates left, zoom right.
    expect(bar.lastElementChild).toBe(zoom);
    expect((wrapper.get('[data-testid="zoom-percent"]').element as HTMLInputElement).value).toBe("85%");
    expect(wrapper.findComponent(ZoomControl).props("showFit")).toBe(false);
    expect(wrapper.find('[data-testid="zoom-fit-width"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="zoom-fit-page"]').exists()).toBe(false);
    wrapper.unmount();
  });

  it("forwards the zoom control's steps, presets and typed value", async () => {
    const wrapper = mount(SpreadsheetStatusBar, { props: { zoomPercent: 100 } });
    await wrapper.get('[data-testid="zoom-in"]').trigger("click");
    await wrapper.get('[data-testid="zoom-out"]').trigger("click");
    const zoom = wrapper.findComponent(ZoomControl);
    zoom.vm.$emit("set-preset", 150);
    zoom.vm.$emit("apply-input", "85%");
    expect(wrapper.emitted("zoom-in")).toHaveLength(1);
    expect(wrapper.emitted("zoom-out")).toHaveLength(1);
    expect(wrapper.emitted("set-zoom")).toEqual([[150]]);
    expect(wrapper.emitted("apply-zoom")).toEqual([["85%"]]);
  });
});
