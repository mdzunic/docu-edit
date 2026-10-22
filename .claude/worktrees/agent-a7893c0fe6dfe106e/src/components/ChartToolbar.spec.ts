import { describe, it, expect, afterEach } from "vitest";
import { mount } from "@vue/test-utils";
import ChartToolbar from "./ChartToolbar.vue";
import { setTierWidthForTest, FULL_MIN_WIDTH } from "../composables/useResponsiveTier";

// SPEC-068 AC-12/13 — the chart data sheet is disabled (with a reason) at the
// compact tier, since a spreadsheet grid at 390px is not usable. The chart
// still renders and can be moved/deleted, so those controls stay enabled.

describe("ChartToolbar — compact-tier data sheet gating (SPEC-068 AC-12/13)", () => {
  afterEach(() => setTierWidthForTest(FULL_MIN_WIDTH));

  it('disables "Edit data" with a stated reason at the compact tier', () => {
    setTierWidthForTest(390);
    const wrapper = mount(ChartToolbar);
    const editData = wrapper.get('[data-testid="chart-edit-data-button"]');
    expect(editData.attributes("disabled")).toBeDefined();
    expect(editData.attributes("title")).toBe("not available on this screen size");
  });

  it('keeps "Edit data" enabled at the full tier', () => {
    setTierWidthForTest(FULL_MIN_WIDTH);
    const wrapper = mount(ChartToolbar);
    const editData = wrapper.get('[data-testid="chart-edit-data-button"]');
    expect(editData.attributes("disabled")).toBeUndefined();
  });

  it("never disables move/delete/duplicate — the chart stays reachable at compact", () => {
    setTierWidthForTest(390);
    const wrapper = mount(ChartToolbar);
    expect(wrapper.get('[data-testid="chart-delete-button"]').attributes("disabled")).toBeUndefined();
    expect(wrapper.get('[data-testid="chart-duplicate-button"]').attributes("disabled")).toBeUndefined();
  });
});

describe("ChartToolbar — the unsupported notice's reason (SPEC-132)", () => {
  it("shows the default text for an unsupported chart without a reason, and the given reason instead of it", () => {
    const plain = mount(ChartToolbar, { props: { unsupported: true } });
    expect(plain.get('[data-testid="chart-unsupported-reason"]').text()).toBe("We can't edit this chart type yet — it's kept exactly as it was.");
    expect(plain.find('[data-testid="chart-edit-data-button"]').exists()).toBe(false);
    expect(plain.find('[data-testid="chart-options-button"]').exists()).toBe(false);
    const reasoned = mount(ChartToolbar, { props: { unsupported: true, reason: "Pivot charts can't be edited yet" } });
    expect(reasoned.get('[data-testid="chart-unsupported-reason"]').text()).toBe("Pivot charts can't be edited yet");
    expect(reasoned.find('[data-testid="chart-edit-data-button"]').exists()).toBe(false);
  });

  it("ignores a reason on a supported chart: the edit controls show and no notice does", () => {
    const wrapper = mount(ChartToolbar, { props: { reason: "Pivot charts can't be edited yet" } });
    expect(wrapper.find('[data-testid="chart-unsupported-reason"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="chart-edit-data-button"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="chart-options-button"]').exists()).toBe(true);
  });
});
