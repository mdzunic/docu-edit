import { beforeEach, describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import PreferencesDialog from "./PreferencesDialog.vue";
import { PALETTE_OPTIONS, loadPreferences, resetPreferences } from "../lib/preferences";

// The redesigned Appearance controls: the segmented icon theme switch
// (system / sun / moon) and the palette chips with their dominant-colour
// swatches. The stored-preferences behaviour behind them is covered by
// preferences.spec.ts and useTheme.spec.ts; this exercises the dialog's
// wiring and the rendered control inventory.
describe("PreferencesDialog appearance controls", () => {
  beforeEach(() => {
    localStorage.clear();
    resetPreferences();
  });

  it("renders the three-way theme segment and applies a click straight through", async () => {
    const wrapper = mount(PreferencesDialog);
    for (const which of ["system", "light", "dark"]) {
      expect(wrapper.find(`[data-testid="preferences-theme-${which}"]`).exists()).toBe(true);
    }
    expect(
      wrapper.get('[data-testid="preferences-theme-system"]').attributes("aria-pressed"),
    ).toBe("true");

    await wrapper.get('[data-testid="preferences-theme-dark"]').trigger("click");
    expect(loadPreferences().theme).toBe("dark");
    expect(
      wrapper.get('[data-testid="preferences-theme-dark"]').attributes("aria-pressed"),
    ).toBe("true");
  });

  it("renders one swatch chip per palette, in order, with its dominant colour", () => {
    const wrapper = mount(PreferencesDialog);
    const chips = wrapper.findAll('[data-testid="preferences-palette"] .palette-chip');
    expect(chips.map((chip) => chip.attributes("data-testid"))).toEqual(
      PALETTE_OPTIONS.map((option) => `preferences-palette-${option.value}`),
    );
    // jsdom serialises the inline style to rgb(); compare numerically.
    const toRgb = (hex: string) => {
      const n = parseInt(hex.slice(1), 16);
      return `rgb(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255})`;
    };
    for (const [index, option] of PALETTE_OPTIONS.entries()) {
      const swatch = chips[index]!.get(".palette-swatch");
      expect(swatch.attributes("style")).toContain(`background: ${toRgb(option.swatch)}`);
      expect(chips[index]!.text()).toContain(option.label);
    }
  });

  it("clicking a palette chip applies it and marks the chip active", async () => {
    const wrapper = mount(PreferencesDialog);
    await wrapper.get('[data-testid="preferences-palette-nocturne"]').trigger("click");
    expect(loadPreferences().palette).toBe("nocturne");
    expect(
      wrapper.get('[data-testid="preferences-palette-nocturne"]').attributes("aria-pressed"),
    ).toBe("true");
    expect(
      wrapper.get('[data-testid="preferences-palette-modernist"]').attributes("aria-pressed"),
    ).toBe("false");
  });
});
