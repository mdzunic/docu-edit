// SPEC-077 AC-20/21/22/23 for the highlight surface — the theme row above
// the 15-colour palette, and the stored value pinned at model level: a
// theme swatch emits its ThemeColorRef, a palette swatch the OOXML name,
// "No colour" the explicit false (AC-5's three forms).
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import HighlightPicker from "./HighlightPicker.vue";
import { DEFAULT_THEME } from "../lib/core/theme";

async function openPicker(theme?: object) {
  const wrapper = mount(HighlightPicker, {
    props: theme ? { theme } : {},
    attachTo: document.body,
  });
  await wrapper.get('[data-testid="highlight-button"]').trigger("click");
  return wrapper;
}

describe("without a theme (AC-21: existing mounts unchanged)", () => {
  it("renders no theme row; the 15 swatches and No colour stay as they were", async () => {
    const wrapper = await openPicker();
    expect(wrapper.findAll('[data-testid="highlight-theme-swatch"]')).toHaveLength(0);
    expect(wrapper.findAll('[data-testid="highlight-swatch"]')).toHaveLength(15);
    expect(wrapper.find('[data-testid="highlight-none"]').exists()).toBe(true);
    wrapper.unmount();
  });
});

describe("with a theme (SPEC-077 AC-20/23)", () => {
  it("renders the theme row above the standard palette", async () => {
    const wrapper = await openPicker(DEFAULT_THEME);
    expect(wrapper.findAll('[data-testid="highlight-theme-swatch"]')).toHaveLength(60); // 10 base + 10 × 5 variants
    const menu = wrapper.get('[data-testid="highlight-menu"]').element as HTMLElement;
    const firstTheme = menu.querySelector('[data-testid="highlight-theme-swatch"]')!;
    const firstPalette = menu.querySelector('[data-testid="highlight-swatch"]')!;
    expect(firstTheme.compareDocumentPosition(firstPalette) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    wrapper.unmount();
  });

  it("a theme swatch emits the ThemeColorRef; palette and No colour keep today's string/false forms (AC-5)", async () => {
    const wrapper = await openPicker(DEFAULT_THEME);
    const accent5 = wrapper
      .findAll('[data-testid="highlight-theme-swatch"]')
      .find((swatch) => swatch.attributes("aria-label") === "Accent 5")!;
    await accent5.trigger("click");
    expect(wrapper.emitted("select")![0]).toEqual([{ theme: "accent5" }]);

    await wrapper.get('[data-testid="highlight-button"]').trigger("click");
    await wrapper.get('[data-testid="highlight-swatch"][data-color="yellow"]').trigger("click");
    expect(wrapper.emitted("select")![1]).toEqual(["yellow"]);

    await wrapper.get('[data-testid="highlight-button"]').trigger("click");
    await wrapper.get('[data-testid="highlight-none"]').trigger("click");
    expect(wrapper.emitted("select")![2]).toEqual([false]);
    wrapper.unmount();
  });

  it("the emitted reference is a plain object that survives structured clone (AC-22)", async () => {
    const wrapper = await openPicker(DEFAULT_THEME);
    const swatch = wrapper
      .findAll('[data-testid="highlight-theme-swatch"]')
      .find((s) => s.attributes("aria-label") === "Accent 5, lighter 40%")!;
    await swatch.trigger("click");
    expect(() => structuredClone(wrapper.emitted("select")![0][0])).not.toThrow();
    wrapper.unmount();
  });
});
