// SPEC-077 AC-22/AC-23 — the shape style panel's theme swatch rows for the
// solid fill and the outline. The panel's own SPEC-073 behaviour is covered
// end-to-end by e2e/SPEC-073.spec.ts; these tests pin the theme contract: a
// swatch emits the whole fill/outline with the ThemeColorRef as its colour
// (the owner applies it to the model verbatim), while the colour inputs only
// ever display resolved hexes.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import ShapeStylePanel from "./ShapeStylePanel.vue";
import { DEFAULT_THEME } from "../lib/core/theme";

function open(styling: Record<string, unknown>, theme: object | undefined = DEFAULT_THEME) {
  return mount(ShapeStylePanel, { props: { styling, theme } });
}

function swatchFor(wrapper: ReturnType<typeof open>, testid: string, slot: string) {
  // The first matching swatch is the base row's (no lumMod/lumOff variant).
  return wrapper.findAll(`[data-testid="${testid}"]`).find((s) => s.attributes("data-theme-slot") === slot)!;
}

describe("ShapeStylePanel — theme swatch rows (SPEC-077 AC-22/AC-23)", () => {
  it("renders no theme rows without a theme (AC-21: existing mounts unchanged)", () => {
    const wrapper = mount(ShapeStylePanel, { props: { styling: { fill: { kind: "solid", color: "#ff0000" } } } });
    expect(wrapper.find('[data-testid="shape-fill-theme-row"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="shape-outline-theme-row"]').exists()).toBe(false);
  });

  it("a fill swatch emits a solid fill carrying the ThemeColorRef (AC-23)", async () => {
    const wrapper = open({ fill: { kind: "solid", color: "#ff0000" } });
    await swatchFor(wrapper, "shape-fill-theme-swatch", "accent2").trigger("click");
    expect(wrapper.emitted("fill")![0][0]).toEqual({ kind: "solid", color: { theme: "accent2" } });
  });

  it("a fill swatch preserves the fill's alpha, like the colour input does", async () => {
    const wrapper = open({ fill: { kind: "solid", color: "#ff0000", alpha: 0.5 } });
    await swatchFor(wrapper, "shape-fill-theme-swatch", "accent3").trigger("click");
    expect(wrapper.emitted("fill")![0][0]).toEqual({ kind: "solid", color: { theme: "accent3" }, alpha: 0.5 });
  });

  it("an outline swatch emits the outline carrying the ThemeColorRef (AC-23)", async () => {
    const wrapper = open({ outline: { color: "#2e5395", widthPt: 2, dash: "dash" } });
    await swatchFor(wrapper, "shape-outline-theme-swatch", "accent6").trigger("click");
    expect(wrapper.emitted("outline")![0][0]).toEqual({ color: { theme: "accent6" }, widthPt: 2, dash: "dash" });
  });

  it("theme-referenced fill and outline colours display as resolved hexes", () => {
    const wrapper = open({
      fill: { kind: "solid", color: { theme: "accent1" } },
      outline: { color: { theme: "accent5" }, widthPt: 1, dash: "solid" },
    });
    expect((wrapper.get('[data-testid="shape-fill-solid-color"]').element as HTMLInputElement).value).toBe("#4472c4");
    expect((wrapper.get('[data-testid="shape-style-outline-color"]').element as HTMLInputElement).value).toBe("#5b9bd5");
  });

  it("hand-editing the colour inputs still emits literal hexes (AC-21)", async () => {
    const wrapper = open({
      fill: { kind: "solid", color: { theme: "accent1" } },
      outline: { color: { theme: "accent5" }, widthPt: 1, dash: "solid" },
    });
    await wrapper.get('[data-testid="shape-fill-solid-color"]').setValue("#123456");
    await wrapper.get('[data-testid="shape-style-outline-color"]').setValue("#654321");
    expect(wrapper.emitted("fill")![0][0]).toEqual({ kind: "solid", color: "#123456" });
    expect(wrapper.emitted("outline")![0][0]).toEqual({ color: "#654321", widthPt: 1, dash: "solid" });
  });
});
