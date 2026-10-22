// SPEC-077 AC-10..AC-13, AC-19: the theme gallery dialog — card order and
// the current marker, the preview/specimen fallback, the Customise pane's
// twelve slots + two fonts + reset, and that every path only ever emits.
import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import { BUILT_IN_THEMES, DEFAULT_THEME, THEME_COLOR_SLOTS, themeSlotLabel } from "../lib/core/theme";
import ThemeGalleryDialog from "./ThemeGalleryDialog.vue";

function mountDialog(props: Record<string, unknown> = {}) {
  return mount(ThemeGalleryDialog, {
    props: { currentTheme: DEFAULT_THEME, previewParagraphs: [], ...props },
    attachTo: document.body,
  });
}

describe("ThemeGalleryDialog (SPEC-077)", () => {
  it("shows the current theme first, marked Current, then the six built-ins Office first (AC-10)", () => {
    const wrapper = mountDialog({ currentTheme: { ...DEFAULT_THEME, name: "My scheme" } });
    const grid = wrapper.get('[data-testid="theme-gallery-grid"]');
    const cards = grid.findAll("button");
    expect(cards).toHaveLength(1 + BUILT_IN_THEMES.length);
    // First card is the document's theme, visibly badged and named as current.
    expect(cards[0].attributes("data-testid")).toBe("theme-card-current");
    expect(cards[0].get('[data-testid="theme-current-badge"]').text()).toBe("Current");
    expect(cards[0].attributes("aria-label")).toBe("My scheme (current theme)");
    // Then the built-ins in catalogue order, Office first.
    expect(cards.slice(1).map((card) => card.attributes("data-theme-name"))).toEqual(
      BUILT_IN_THEMES.map((theme) => theme.name),
    );
  });

  it("emits apply with the card's theme on click (AC-14 upstream)", async () => {
    const wrapper = mountDialog();
    await wrapper.get('[data-testid="theme-card"][data-theme-name="Facet"]').trigger("click");
    const emitted = wrapper.emitted("apply");
    expect(emitted).toHaveLength(1);
    expect(emitted?.[0]?.[0]).toMatchObject({ name: "Facet", colors: BUILT_IN_THEMES[1].colors });
  });

  it("falls back to a specimen sheet with six accent swatches when no preview exists (AC-11)", () => {
    const wrapper = mountDialog();
    const card = wrapper.get('[data-testid="theme-card-current"]');
    const preview = card.get(".card-preview");
    expect(preview.attributes("aria-hidden")).toBe("true");
    expect(preview.text()).toContain("Heading");
    expect(preview.findAll('[data-testid="theme-accent-swatch"]')).toHaveLength(6);
  });

  it("renders the document miniature in the card's theme when a preview exists (AC-11)", () => {
    const wrapper = mountDialog({
      previewParagraphs: [
        { text: "Annual report", heading: true },
        { text: "Body copy line", heading: false },
      ],
    });
    const facet = wrapper.get('[data-theme-name="Facet"]');
    const lines = facet.findAll(".preview-line");
    expect(lines.map((line) => line.text())).toEqual(["Annual report", "Body copy line"]);
    // Heading styled with the card's major font/accent, body with minor/dk1.
    expect(lines[0].attributes("style")).toContain("Trebuchet MS");
    expect(facet.findAll('[data-testid="theme-accent-swatch"]')).toHaveLength(0);
  });

  it("Customise lists the twelve slots in scheme order, each labelled (AC-12, AC-13)", async () => {
    const wrapper = mountDialog();
    await wrapper.get('[data-testid="theme-customise"]').trigger("click");
    const pane = wrapper.get('[data-testid="theme-customise-pane"]');
    const inputs = pane.findAll('input[type="color"]');
    expect(inputs.map((input) => input.attributes("data-testid"))).toEqual(
      THEME_COLOR_SLOTS.map((slot) => `theme-colour-${slot}`),
    );
    expect(pane.findAll(".colour-field .field-label").map((label) => label.text())).toEqual(
      THEME_COLOR_SLOTS.map((slot) => themeSlotLabel(slot)),
    );
    expect(pane.find('[data-testid="theme-font-major"]').exists()).toBe(true);
    expect(pane.find('[data-testid="theme-font-minor"]').exists()).toBe(true);
  });

  it("applies the customised draft without ever touching the input theme (AC-12, AC-19)", async () => {
    const current = { ...DEFAULT_THEME, colors: { ...DEFAULT_THEME.colors }, fonts: { ...DEFAULT_THEME.fonts } };
    const wrapper = mountDialog({ currentTheme: current });
    await wrapper.get('[data-testid="theme-customise"]').trigger("click");
    await wrapper.get('[data-testid="theme-colour-accent1"]').setValue("#123456");
    await wrapper.get('[data-testid="theme-font-major"]').setValue("Georgia");
    await wrapper.get('[data-testid="theme-apply-custom"]').trigger("click");
    const applied = wrapper.emitted("apply")?.[0]?.[0] as {
      colors: Record<string, string>;
      fonts: { major: string; minor: string };
    };
    expect(applied.colors.accent1).toBe("#123456");
    expect(applied.colors.accent2).toBe(DEFAULT_THEME.colors.accent2);
    expect(applied.fonts.major).toBe("Georgia");
    // The prop object was never written to — previews and drafts are local.
    expect(current.colors.accent1).toBe(DEFAULT_THEME.colors.accent1);
    expect(current.fonts.major).toBe(DEFAULT_THEME.fonts.major);
  });

  it("Reset returns the draft to the theme as it was when the dialog opened (AC-12)", async () => {
    const wrapper = mountDialog();
    await wrapper.get('[data-testid="theme-customise"]').trigger("click");
    await wrapper.get('[data-testid="theme-colour-dk2"]').setValue("#ff0000");
    await wrapper.get('[data-testid="theme-font-minor"]').setValue("Courier New");
    await wrapper.get('[data-testid="theme-reset"]').trigger("click");
    await wrapper.get('[data-testid="theme-apply-custom"]').trigger("click");
    const applied = wrapper.emitted("apply")?.[0]?.[0] as {
      colors: Record<string, string>;
      fonts: { major: string; minor: string };
    };
    expect(applied.colors.dk2).toBe(DEFAULT_THEME.colors.dk2);
    expect(applied.fonts.minor).toBe(DEFAULT_THEME.fonts.minor);
  });

  it("Cancel emits cancel and never apply (AC-19)", async () => {
    const wrapper = mountDialog();
    await wrapper.get('[data-testid="theme-gallery-cancel"]').trigger("click");
    expect(wrapper.emitted("cancel")).toHaveLength(1);
    expect(wrapper.emitted("apply")).toBeUndefined();
  });
});
