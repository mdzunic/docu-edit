import { beforeEach, describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import PreferencesDialog from "./PreferencesDialog.vue";
import { PALETTE_OPTIONS, loadPreferences, resetPreferences } from "../lib/preferences";

// jsdom serialises the inline style to rgb(); compare numerically.
const toRgb = (hex: string) => {
  const n = parseInt(hex.slice(1), 16);
  return `rgb(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255})`;
};

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

// SPEC-118 "The preference": a "Spreadsheets" section right after "Editing"
// with "Number and date format" — each locale labelled with its name and a
// live sample the format engine writes — applied straight through.
describe("PreferencesDialog spreadsheet locale", () => {
  beforeEach(() => {
    localStorage.clear();
    resetPreferences();
  });

  it("puts a Spreadsheets section right after Editing", () => {
    const wrapper = mount(PreferencesDialog);
    const headings = wrapper.findAll("h3.section-title").map((heading) => heading.text());
    expect(headings[headings.indexOf("Editing") + 1]).toBe("Spreadsheets");
    const section = wrapper.get('[aria-labelledby="preferences-spreadsheets-heading"]');
    expect(section.find('[data-testid="preferences-spreadsheet-locale"]').exists()).toBe(true);
    expect(section.text()).toContain("Number and date format");
  });

  it("labels each of the six locales with its name and a live sample, en-US chosen by default", () => {
    const wrapper = mount(PreferencesDialog);
    const select = wrapper.get('[data-testid="preferences-spreadsheet-locale"]');
    const options = select.findAll("option");
    expect(options.map((option) => option.attributes("value"))).toEqual(["en-US", "en-GB", "de-DE", "fr-FR", "es-ES", "cs-CZ"]);
    expect(options[0]!.text()).toBe("English (United States) — 1,234.56 · 10/8/2026");
    expect(options[2]!.text()).toBe("Deutsch (Deutschland) — 1.234,56 · 08.10.2026");
    expect(options[5]!.text()).toBe("Čeština (Česko) — 1\u00a0234,56 · 08.10.2026");
    expect((select.element as HTMLSelectElement).value).toBe("en-US");
  });

  it("applies a choice straight through to the preference", async () => {
    const wrapper = mount(PreferencesDialog);
    await wrapper.get('[data-testid="preferences-spreadsheet-locale"]').setValue("de-DE");
    expect(loadPreferences().spreadsheetLocale).toBe("de-DE");
  });
});
