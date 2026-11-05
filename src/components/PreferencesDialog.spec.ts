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
// live sample the format engine writes — applied straight through. SPEC-146
// places its "Presentations" section before "Spreadsheets", so it now sits
// between the two.
describe("PreferencesDialog spreadsheet locale", () => {
  beforeEach(() => {
    localStorage.clear();
    resetPreferences();
  });

  it("puts a Spreadsheets section after Editing, with only SPEC-146's Presentations between", () => {
    const wrapper = mount(PreferencesDialog);
    const headings = wrapper.findAll("h3.section-title").map((heading) => heading.text());
    expect(headings.slice(headings.indexOf("Editing"), headings.indexOf("Editing") + 3)).toEqual([
      "Editing",
      "Presentations",
      "Spreadsheets",
    ]);
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

// SPEC-113: a "Labs" section (prefs-labs) joins the dialog just before
// Advanced — so "Reset all preferences" still restores "every setting above"
// — and every other section stays exactly where it was (SPEC-146 has since
// added "Presentations", before "Spreadsheets").
describe("PreferencesDialog Labs section", () => {
  beforeEach(() => {
    localStorage.clear();
    resetPreferences();
  });

  it("sits between Drawing and Advanced, leaving the other sections in order", () => {
    const wrapper = mount(PreferencesDialog);
    const sections = wrapper.findAll("section").map((section) => section.get("h3").text());
    expect(sections).toEqual([
      "Appearance",
      "Editing",
      "Presentations",
      "Spreadsheets",
      "New documents",
      "Saving",
      "Drawing",
      "Labs",
      "Advanced",
    ]);
    const labs = wrapper.get('[data-testid="prefs-labs"]');
    expect(labs.find('[data-testid="experiment-library-search"]').exists()).toBe(true);
    expect(labs.find('[data-testid="experiment-deck-designer"]').exists()).toBe(true);
  });
});

// SPEC-146: a "Presentations" section, before "Spreadsheets", holding the show
// timer's three per-device preferences — auto-start, mode and talk length —
// each reflecting and setting its key; the length commits through
// parseCountdownMinutes; "Reset to defaults" covers all three.
describe("PreferencesDialog Presentations section", () => {
  beforeEach(() => {
    localStorage.clear();
    resetPreferences();
  });

  it("sits before Spreadsheets, labelled by its heading, with an icon", () => {
    const wrapper = mount(PreferencesDialog);
    const headings = wrapper.findAll("h3.section-title").map((heading) => heading.text());
    expect(headings.indexOf("Presentations")).toBe(headings.indexOf("Spreadsheets") - 1);
    const section = wrapper.get('[aria-labelledby="preferences-presentations-heading"]');
    expect(section.get("#preferences-presentations-heading").text()).toBe("Presentations");
    expect(section.find("#preferences-presentations-heading svg").exists()).toBe(true);
    for (const id of ["autostart", "mode", "length"]) {
      expect(section.find(`[data-testid="preferences-presentation-timer-${id}"]`).exists(), id).toBe(true);
    }
  });

  it("the auto-start checkbox reflects and sets presentationTimerAutoStart", async () => {
    const wrapper = mount(PreferencesDialog);
    const box = wrapper.get('[data-testid="preferences-presentation-timer-autostart"]');
    expect(box.attributes("type")).toBe("checkbox");
    expect((box.element as HTMLInputElement).checked).toBe(true);
    expect(wrapper.get('label[for="preferences-presentation-timer-autostart"]').text()).toBe("Start the timer when a show starts");
    expect(wrapper.get("#preferences-presentation-timer-autostart-hint").text()).toBe(
      "Presenter view and the full-screen show. Off: the timer waits until you press Start (T). Default: on",
    );
    await box.setValue(false);
    expect(loadPreferences().presentationTimerAutoStart).toBe(false);
    await box.setValue(true);
    expect(loadPreferences().presentationTimerAutoStart).toBe(true);
  });

  it("the Timer select offers the two modes and sets presentationTimerMode", async () => {
    const wrapper = mount(PreferencesDialog);
    const select = wrapper.get('[data-testid="preferences-presentation-timer-mode"]');
    expect(wrapper.get('label[for="preferences-presentation-timer-mode"]').text()).toBe("Timer");
    expect(select.findAll("option").map((option) => [option.attributes("value"), option.text()])).toEqual([
      ["stopwatch", "Stopwatch — counts up from 0:00"],
      ["countdown", "Countdown — counts down from the talk length"],
    ]);
    expect((select.element as HTMLSelectElement).value).toBe("stopwatch");
    expect(wrapper.get("#preferences-presentation-timer-mode-hint").text()).toBe("Default: Stopwatch");
    await select.setValue("countdown");
    expect(loadPreferences().presentationTimerMode).toBe("countdown");
  });

  it("the talk length reflects its key and commits through parseCountdownMinutes", async () => {
    const wrapper = mount(PreferencesDialog);
    const field = wrapper.get('[data-testid="preferences-presentation-timer-length"]');
    const input = field.element as HTMLInputElement;
    expect(field.attributes()).toMatchObject({ type: "text", inputmode: "numeric" });
    expect(wrapper.get('label[for="preferences-presentation-timer-length"]').text()).toBe("Talk length (minutes)");
    expect(wrapper.get("#preferences-presentation-timer-length-hint").text()).toBe("1–240. Used by Countdown. Default: 20");
    expect(input.value).toBe("20");
    for (const [typed, minutes] of [["45", 45], ["2.6", 3], ["0", 1], ["999", 240]] as const) {
      input.value = typed;
      await field.trigger("change");
      expect(loadPreferences().presentationTimerCountdownMinutes, typed).toBe(minutes);
      expect(input.value).toBe(String(minutes));
    }
    input.value = "abc"; // not a number: the field restores, the key keeps its value
    await field.trigger("change");
    expect(input.value).toBe("240");
    expect(loadPreferences().presentationTimerCountdownMinutes).toBe(240);
  });

  it("Reset to defaults restores all three", async () => {
    const wrapper = mount(PreferencesDialog);
    await wrapper.get('[data-testid="preferences-presentation-timer-autostart"]').setValue(false);
    await wrapper.get('[data-testid="preferences-presentation-timer-mode"]').setValue("countdown");
    const length = wrapper.get('[data-testid="preferences-presentation-timer-length"]');
    (length.element as HTMLInputElement).value = "5";
    await length.trigger("change");
    await wrapper.get('[data-testid="preferences-reset"]').trigger("click");
    await wrapper.get('[data-testid="preferences-reset-confirm"]').trigger("click");
    const prefs = loadPreferences();
    expect([prefs.presentationTimerAutoStart, prefs.presentationTimerMode, prefs.presentationTimerCountdownMinutes]).toEqual([
      true,
      "stopwatch",
      20,
    ]);
    expect((length.element as HTMLInputElement).value).toBe("20");
  });
});
