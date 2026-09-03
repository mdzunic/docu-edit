// SPEC-077 AC-22/AC-23 — the chart options dialog's per-series theme swatch
// rows. The dialog's SPEC-052 behaviour is covered end-to-end by
// e2e/SPEC-052.spec.ts; these tests pin the theme contract: a swatch emits
// `series-color` with the ThemeColorRef itself (the owner writes only that
// series, so every untouched series keeps its stored value — see
// charts.spec.ts for the model-level pin), while the colour inputs display
// resolved hexes.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import ChartOptionsDialog from "./ChartOptionsDialog.vue";
import { DEFAULT_THEME } from "../lib/core/theme";

const INITIAL = {
  legend: "none",
  dataLabels: false,
  gridlines: false,
  hasAxes: false,
  series: [
    { name: "A", values: [1], color: { theme: "accent1" } },
    { name: "B", values: [2], color: "#222222" },
  ],
};

function open(theme: object | undefined = DEFAULT_THEME) {
  return mount(ChartOptionsDialog, { props: { initial: INITIAL, theme } });
}

describe("ChartOptionsDialog — series theme swatches (SPEC-077 AC-22/AC-23)", () => {
  it("renders no theme rows without a theme (AC-21: existing mounts unchanged)", () => {
    const wrapper = mount(ChartOptionsDialog, { props: { initial: INITIAL } });
    expect(wrapper.find('[data-testid="chart-series-theme-0-row"]').exists()).toBe(false);
  });

  it("displays a theme-referenced series colour as its resolved hex", () => {
    const wrapper = open();
    expect((wrapper.get('[data-testid="chart-options-series-color-0"]').element as HTMLInputElement).value).toBe("#4472c4");
    expect((wrapper.get('[data-testid="chart-options-series-color-1"]').element as HTMLInputElement).value).toBe("#222222");
  });

  it("a swatch emits series-color with the ThemeColorRef, for that series only (AC-23)", async () => {
    const wrapper = open();
    const swatch = wrapper
      .findAll('[data-testid="chart-series-theme-1-swatch"]')
      .find((s) => s.attributes("data-theme-slot") === "accent2")!;
    await swatch.trigger("click");
    const emitted = wrapper.emitted("series-color")!;
    // Only the changed series is written — untouched series are never emitted.
    expect(emitted).toHaveLength(1);
    expect(emitted[0]).toEqual([1, { theme: "accent2" }]);
  });

  it("hand-editing a series colour input still emits a literal hex (AC-21)", async () => {
    const wrapper = open();
    await wrapper.get('[data-testid="chart-options-series-color-1"]').setValue("#abcdef");
    expect(wrapper.emitted("series-color")![0]).toEqual([1, "#abcdef"]);
  });
});
