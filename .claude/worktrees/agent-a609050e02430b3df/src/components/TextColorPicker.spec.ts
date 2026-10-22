// SPEC-094 AC: the theme-aware colour picker — the theme row above the
// standard palette, "Accent 1, lighter 40%"-form accessible names, and
// ref-not-hex storage; recents and custom unchanged below.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import TextColorPicker from "./TextColorPicker.vue";
import { DEFAULT_THEME, resolveColor } from "../lib/core/theme";

async function openPicker(theme?: object) {
  const wrapper = mount(TextColorPicker, {
    props: theme ? { theme } : {},
    attachTo: document.body,
  });
  await wrapper.get('[data-testid="text-color-button"]').trigger("click");
  return wrapper;
}

describe("without a theme (the word processor's picker)", () => {
  it("renders no theme row; palette, recent and custom stay as they were", async () => {
    const wrapper = await openPicker();
    expect(wrapper.findAll('[data-testid="text-color-theme-swatch"]')).toHaveLength(0);
    expect(wrapper.findAll('[data-testid="text-color-swatch"]').length).toBeGreaterThan(0);
    expect(wrapper.find('[data-testid="text-color-custom-input"]').exists()).toBe(true);
    wrapper.unmount();
  });
});

describe("with a theme (SPEC-094 AC-45..47/49)", () => {
  it("renders ten swatch columns with five variants each, above the palette", async () => {
    const wrapper = await openPicker(DEFAULT_THEME);
    const swatches = wrapper.findAll('[data-testid="text-color-theme-swatch"]');
    expect(swatches).toHaveLength(60); // 10 base + 10 × 5 variants
    // The theme rows precede the standard palette in DOM order.
    const menu = wrapper.get('[data-testid="text-color-menu"]').element as HTMLElement;
    const firstTheme = menu.querySelector('[data-testid="text-color-theme-swatch"]')!;
    const palette = menu.querySelector('[data-testid="text-color-palette"]')!;
    expect(firstTheme.compareDocumentPosition(palette) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    wrapper.unmount();
  });

  it('names swatches in the "Accent 1, lighter 40%" form', async () => {
    const wrapper = await openPicker(DEFAULT_THEME);
    const names = wrapper
      .findAll('[data-testid="text-color-theme-swatch"]')
      .map((swatch) => swatch.attributes("aria-label"));
    expect(names).toContain("Accent 1");
    expect(names).toContain("Accent 1, lighter 40%");
    expect(names).toContain("Dark 2, darker 25%");
    wrapper.unmount();
  });

  it("stores the ThemeColorRef, not the hex it displays (AC-49)", async () => {
    const wrapper = await openPicker(DEFAULT_THEME);
    const lighter40 = wrapper
      .findAll('[data-testid="text-color-theme-swatch"]')
      .find((swatch) => swatch.attributes("aria-label") === "Accent 1, lighter 40%")!;
    // The swatch DISPLAYS the resolved value…
    expect(lighter40.attributes("data-color")).toBe(
      resolveColor({ theme: "accent1", lumMod: 0.6, lumOff: 0.4 }, DEFAULT_THEME),
    );
    await lighter40.trigger("click");
    // …but STORES the reference.
    expect(wrapper.emitted("select")).toEqual([[{ theme: "accent1", lumMod: 0.6, lumOff: 0.4 }]]);
    wrapper.unmount();
  });

  it("a standard palette swatch still stores a literal", async () => {
    const wrapper = await openPicker(DEFAULT_THEME);
    await wrapper.findAll('[data-testid="text-color-swatch"]')[0].trigger("click");
    expect(wrapper.emitted("select")).toEqual([["#000000"]]);
    wrapper.unmount();
  });
});
