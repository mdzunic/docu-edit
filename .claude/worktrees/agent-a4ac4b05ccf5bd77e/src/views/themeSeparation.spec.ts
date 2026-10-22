import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { nextTick } from "vue";
import { useTheme } from "../composables/useTheme";
import { BUILT_IN_THEMES, DEFAULT_THEME, resolveColor } from "../lib/core/theme";
import { createDefaultStyles, documentThemeOf, type DocumentModel, type Paragraph } from "../lib/documentModel";

// SPEC-077 AC-49: the APPLICATION theme (SPEC-059's data-theme + tokens.css
// custom properties) and the DOCUMENT theme (the model's DocumentTheme,
// resolved by resolveColor) are two systems that never touch. One test does
// both directions: switching the app to dark mode changes no document colour
// in the model or as painted, and applying a document theme changes no app
// CSS custom property.

/** A minimal document with one literal and one theme-referencing run colour. */
function modelWith(): DocumentModel {
  return {
    id: "theme-separation",
    title: "Separation",
    defaultStyles: createDefaultStyles(),
    sections: [
      {
        blocks: [
          {
            type: "paragraph",
            runs: [
              { text: "literal", formatting: { color: "#c0392b" } },
              { text: "themed", formatting: { color: { theme: "accent1", lumMod: 0.75 } } },
            ],
          },
        ],
      },
    ],
  };
}

/** What the page paints for each run: the model colour resolved through the
 * single resolution point against the document's own theme. */
function paintedColors(model: DocumentModel): string[] {
  const theme = documentThemeOf(model);
  const paragraph = model.sections[0].blocks[0] as Paragraph;
  return paragraph.runs.map((run) => resolveColor(run.formatting.color!, theme));
}

// A representative slice of the app palette, set inline so jsdom (which does
// not cascade tokens.css) can observe the custom properties the way
// resolveThemeColor does — off documentElement.
const APP_TOKENS: Record<string, string> = {
  "--surface-base": "#fbfbfa",
  "--surface-raised": "#ffffff",
  "--text-primary": "#1c1c1e",
  "--accent": "#3a6ea5",
  "--focus-ring": "#3a6ea5",
};

function appTokenSnapshot(): Record<string, string> {
  const style = getComputedStyle(document.documentElement);
  return Object.fromEntries(Object.keys(APP_TOKENS).map((name) => [name, style.getPropertyValue(name).trim()]));
}

describe("app dark mode and the document theme are independent (SPEC-077 AC-49)", () => {
  beforeEach(() => {
    localStorage.clear();
    document.documentElement.removeAttribute("data-theme");
    document.documentElement.removeAttribute("data-palette");
    for (const [name, value] of Object.entries(APP_TOKENS)) {
      document.documentElement.style.setProperty(name, value);
    }
  });
  afterEach(() => {
    // Leave the useTheme singleton in `system` so later files start clean.
    useTheme().setTheme("system");
    for (const name of Object.keys(APP_TOKENS)) {
      document.documentElement.style.removeProperty(name);
    }
  });

  it("dark mode changes no document colour; a document theme changes no app custom property", async () => {
    const { setTheme } = useTheme();
    const model = modelWith();

    const modelBefore = JSON.stringify(model);
    const paintedBefore = paintedColors(model);
    // Sanity: the literal passes through; the reference resolves through the
    // default scheme with its lumMod applied (so it is a real derived colour).
    expect(paintedBefore[0]).toBe("#c0392b");
    expect(paintedBefore[1]).not.toBe(DEFAULT_THEME.colors.accent1);

    // Direction 1 — app → document: switch the app to dark mode (SPEC-059).
    setTheme("dark");
    await nextTick();
    expect(document.documentElement.getAttribute("data-theme")).toBe("dark");
    // The model is byte-identical and every painted colour is unchanged.
    expect(JSON.stringify(model)).toBe(modelBefore);
    expect(paintedColors(model)).toEqual(paintedBefore);

    // Direction 2 — document → app: apply a different document theme.
    const appPropsBefore = appTokenSnapshot();
    expect(appPropsBefore).toEqual(APP_TOKENS); // the snapshot really observes them
    model.theme = { ...BUILT_IN_THEMES[1] }; // Facet
    expect(documentThemeOf(model).name).toBe("Facet");
    const paintedAfter = paintedColors(model);
    // The document theme really took effect on the page…
    expect(paintedAfter[1]).not.toBe(paintedBefore[1]);
    expect(paintedAfter[0]).toBe("#c0392b"); // literals stay literal
    // …while the app's custom properties and appearance are untouched.
    expect(appTokenSnapshot()).toEqual(appPropsBefore);
    expect(document.documentElement.getAttribute("data-theme")).toBe("dark");
  });
});

// SPEC-077 AC-3: documentThemeOf is the single way a document surface obtains
// its theme — the document's own theme when set, else the product default.
describe("documentThemeOf (SPEC-077 AC-3)", () => {
  it("returns the model's own theme when set", () => {
    const theme = { ...BUILT_IN_THEMES[2] };
    expect(documentThemeOf({ theme })).toBe(theme);
    expect(documentThemeOf({ theme }).name).toBe("Ion");
  });

  it("falls back to DEFAULT_THEME when the model carries none (or no model)", () => {
    expect(documentThemeOf({})).toBe(DEFAULT_THEME);
    expect(documentThemeOf(null)).toBe(DEFAULT_THEME);
    expect(documentThemeOf(undefined)).toBe(DEFAULT_THEME);
  });
});
