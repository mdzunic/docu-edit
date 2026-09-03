import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import BordersShadingDialog from "./BordersShadingDialog.vue";
import { DEFAULT_THEME } from "../lib/core/theme";

const LINE = { style: "single", widthPt: 1, color: "#000000" };

function open(initial: Record<string, unknown> = {}, unit = "cm") {
  return mount(BordersShadingDialog, { props: { initial, unit } });
}

function applied(wrapper: ReturnType<typeof open>) {
  return wrapper.emitted("apply")![0][0] as {
    borders: Record<string, unknown> | null;
    shadingFill: string | null;
  };
}

describe("BordersShadingDialog (SPEC-074 AC-6..AC-10)", () => {
  it("offers all five edges independently (AC-6)", () => {
    const wrapper = open();
    for (const edge of ["top", "bottom", "left", "right", "between"]) {
      expect(wrapper.find(`[data-testid="borders-edge-${edge}"]`).exists()).toBe(true);
    }
  });

  it("gives each edge its own style, width and colour (AC-7)", async () => {
    const wrapper = open();
    await wrapper.get('[data-testid="borders-edge-top"]').setValue(true);
    await wrapper.get('[data-testid="borders-style-top"]').setValue("double");
    await wrapper.get('[data-testid="borders-width-top"]').setValue("2.5");
    await wrapper.get('[data-testid="borders-edge-bottom"]').setValue(true);
    await wrapper.get('[data-testid="borders-style-bottom"]').setValue("dotted");
    await wrapper.get('[data-testid="borders-apply"]').trigger("click");
    const result = applied(wrapper);
    expect(result.borders?.top).toEqual({ style: "double", widthPt: 2.5, color: "#000000" });
    expect(result.borders?.bottom).toEqual({ style: "dotted", widthPt: 1, color: "#000000" });
    expect(result.borders?.left).toBeUndefined();
  });

  it("offers the five presets (AC-8)", () => {
    const wrapper = open();
    for (const preset of ["none", "box", "all", "topAndBottom", "custom"]) {
      expect(wrapper.find(`[data-testid="borders-preset-${preset}"]`).exists()).toBe(true);
    }
  });

  it("box sets the four outer edges and no between edge (AC-8)", async () => {
    const wrapper = open();
    await wrapper.get('[data-testid="borders-preset-box"]').trigger("click");
    await wrapper.get('[data-testid="borders-apply"]').trigger("click");
    const result = applied(wrapper);
    expect(result.borders?.top).toEqual(LINE);
    expect(result.borders?.right).toEqual(LINE);
    expect(result.borders?.between).toBeUndefined();
  });

  it("all adds the between edge (AC-8, AC-12)", async () => {
    const wrapper = open();
    await wrapper.get('[data-testid="borders-preset-all"]').trigger("click");
    await wrapper.get('[data-testid="borders-apply"]').trigger("click");
    expect(applied(wrapper).borders?.between).toEqual(LINE);
  });

  it("top and bottom sets exactly those two edges (AC-8)", async () => {
    const wrapper = open();
    await wrapper.get('[data-testid="borders-preset-topAndBottom"]').trigger("click");
    await wrapper.get('[data-testid="borders-apply"]').trigger("click");
    const result = applied(wrapper);
    expect(result.borders?.top).toEqual(LINE);
    expect(result.borders?.bottom).toEqual(LINE);
    expect(result.borders?.left).toBeUndefined();
    expect(result.borders?.right).toBeUndefined();
  });

  it("none clears every edge (AC-8)", async () => {
    const wrapper = open({ borders: { top: LINE, bottom: LINE, left: LINE, right: LINE } });
    await wrapper.get('[data-testid="borders-preset-none"]').trigger("click");
    await wrapper.get('[data-testid="borders-apply"]').trigger("click");
    expect(applied(wrapper).borders).toBeNull();
  });

  it("moves the preset row to Custom once one edge is hand-edited (AC-8)", async () => {
    const wrapper = open();
    await wrapper.get('[data-testid="borders-preset-box"]').trigger("click");
    expect(wrapper.get('[data-testid="borders-preset-box"]').attributes("aria-pressed")).toBe("true");
    await wrapper.get('[data-testid="borders-edge-left"]').setValue(false);
    expect(wrapper.get('[data-testid="borders-preset-custom"]').attributes("aria-pressed")).toBe("true");
  });

  it("defaults every padding field to 4pt in the document's unit (AC-9)", () => {
    const inches = open({}, "in");
    // 4pt is 1/18 inch; formatLength rounds for display.
    expect((inches.get('[data-testid="borders-padding-top"]').element as HTMLInputElement).value).toBe("0.06");
    const points = open({}, "pt");
    expect((points.get('[data-testid="borders-padding-left"]').element as HTMLInputElement).value).toBe("4");
  });

  it("emits an edited padding in points, whatever unit it was typed in (AC-9)", async () => {
    const wrapper = open({}, "pt");
    await wrapper.get('[data-testid="borders-preset-box"]').trigger("click");
    await wrapper.get('[data-testid="borders-padding-top"]').setValue("10");
    await wrapper.get('[data-testid="borders-apply"]').trigger("click");
    expect(applied(wrapper).borders?.paddingPt).toEqual({ top: 10, bottom: 4, left: 4, right: 4 });
  });

  it("omits padding entirely when every edge is left at the 4pt default (AC-9)", async () => {
    const wrapper = open();
    await wrapper.get('[data-testid="borders-preset-box"]').trigger("click");
    await wrapper.get('[data-testid="borders-apply"]').trigger("click");
    expect(applied(wrapper).borders?.paddingPt).toBeUndefined();
  });

  it("treats no fill as a distinct value from white (AC-10)", async () => {
    const off = open();
    await off.get('[data-testid="borders-apply"]').trigger("click");
    expect(applied(off).shadingFill).toBeNull();

    const white = open();
    await white.get('[data-testid="borders-shading-on"]').setValue(true);
    await white.get('[data-testid="borders-shading-color"]').setValue("#ffffff");
    await white.get('[data-testid="borders-apply"]').trigger("click");
    expect(applied(white).shadingFill).toBe("#ffffff");
  });

  it("clearing the fill on a shaded paragraph emits no fill, not white (AC-10)", async () => {
    const wrapper = open({ shadingFill: "#fff2cc" });
    expect((wrapper.get('[data-testid="borders-shading-on"]').element as HTMLInputElement).checked).toBe(true);
    await wrapper.get('[data-testid="borders-shading-on"]').setValue(false);
    await wrapper.get('[data-testid="borders-apply"]').trigger("click");
    expect(applied(wrapper).shadingFill).toBeNull();
  });

  it("prefills from the paragraph it opened on", () => {
    const wrapper = open({
      borders: { bottom: { style: "dashed", widthPt: 2, color: "#c00000" } },
      shadingFill: "#eeeeee",
    });
    expect((wrapper.get('[data-testid="borders-edge-bottom"]').element as HTMLInputElement).checked).toBe(true);
    expect((wrapper.get('[data-testid="borders-style-bottom"]').element as HTMLSelectElement).value).toBe("dashed");
    expect((wrapper.get('[data-testid="borders-width-bottom"]').element as HTMLInputElement).value).toBe("2");
    expect((wrapper.get('[data-testid="borders-edge-top"]').element as HTMLInputElement).checked).toBe(false);
    expect((wrapper.get('[data-testid="borders-shading-color"]').element as HTMLInputElement).value).toBe("#eeeeee");
  });

  it("cancels without emitting an apply", async () => {
    const wrapper = open();
    await wrapper.get('[data-testid="borders-cancel"]').trigger("click");
    expect(wrapper.emitted("apply")).toBeUndefined();
    expect(wrapper.emitted("cancel")).toHaveLength(1);
  });
});

describe("BordersShadingDialog — theme swatch row (SPEC-077 AC-22/AC-23)", () => {
  function openThemed(initial: Record<string, unknown> = {}) {
    return mount(BordersShadingDialog, { props: { initial, unit: "cm", theme: DEFAULT_THEME } });
  }

  function swatchFor(wrapper: ReturnType<typeof openThemed>, slot: string) {
    // The first matching swatch is the base row's (no lumMod/lumOff variant).
    return wrapper.findAll('[data-testid="borders-shading-theme-swatch"]').find(
      (s) => s.attributes("data-theme-slot") === slot,
    )!;
  }

  it("renders no theme row without a theme (AC-21: existing mounts unchanged)", () => {
    expect(open().find('[data-testid="borders-shading-theme-row"]').exists()).toBe(false);
  });

  it("a swatch stores the ThemeColorRef on apply, not the hex it displays (AC-23)", async () => {
    const wrapper = openThemed();
    await swatchFor(wrapper, "accent1").trigger("click");
    // The swatch turns the Fill on and the colour input shows the resolved hex...
    expect((wrapper.get('[data-testid="borders-shading-on"]').element as HTMLInputElement).checked).toBe(true);
    expect((wrapper.get('[data-testid="borders-shading-color"]').element as HTMLInputElement).value).toBe("#4472c4");
    await wrapper.get('[data-testid="borders-apply"]').trigger("click");
    // ...but the payload carries the reference itself.
    expect(applied(wrapper).shadingFill).toEqual({ theme: "accent1" });
  });

  it("hand-editing the colour input after a swatch reverts to a literal", async () => {
    const wrapper = openThemed();
    await swatchFor(wrapper, "accent2").trigger("click");
    await wrapper.get('[data-testid="borders-shading-color"]').setValue("#123456");
    await wrapper.get('[data-testid="borders-apply"]').trigger("click");
    expect(applied(wrapper).shadingFill).toBe("#123456");
  });

  it("a seeded ThemeColorRef displays resolved but survives an untouched OK as the reference", async () => {
    const wrapper = openThemed({ shadingFill: { theme: "accent3" } });
    expect((wrapper.get('[data-testid="borders-shading-on"]').element as HTMLInputElement).checked).toBe(true);
    expect((wrapper.get('[data-testid="borders-shading-color"]').element as HTMLInputElement).value).toBe("#a5a5a5");
    await wrapper.get('[data-testid="borders-apply"]').trigger("click");
    expect(applied(wrapper).shadingFill).toEqual({ theme: "accent3" });
  });
});
