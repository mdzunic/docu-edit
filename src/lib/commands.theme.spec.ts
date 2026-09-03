// SPEC-077 AC-14/AC-18: applying a theme changes `model.theme` and nothing
// else — commands are pure, blocks keep their identity, partName survives.
import { describe, expect, it } from "vitest";
import { caretSelection, documentStart } from "./commands.editing";
import { applyDocumentTheme } from "./commands.theme";
import { BUILT_IN_THEMES } from "./core/theme";
import { createDocument, type DocTheme } from "./documentModel";

const FACET = BUILT_IN_THEMES[1];

describe("applyDocumentTheme (SPEC-077)", () => {
  it("sets model.theme to the chosen theme with dirty: true (AC-18)", () => {
    const model = createDocument("doc-1");
    const selection = caretSelection(documentStart());

    const result = applyDocumentTheme(model, FACET, selection);

    expect(result.model.theme).toEqual({
      name: "Facet",
      colors: FACET.colors,
      fonts: FACET.fonts,
      dirty: true,
    });
    // A built-in applied to a partless document gains no partName.
    expect(result.model.theme && "partName" in result.model.theme).toBe(false);
  });

  it("keeps the existing partName so the part is spliced on save (AC-18)", () => {
    const base = createDocument("doc-2");
    const existing: DocTheme = {
      name: "Office",
      colors: { ...BUILT_IN_THEMES[0].colors },
      fonts: { ...BUILT_IN_THEMES[0].fonts },
      partName: "word/theme/theme1.xml",
    };
    const model = { ...base, theme: existing };

    const result = applyDocumentTheme(model, FACET, caretSelection(documentStart()));

    expect(result.model.theme?.partName).toBe("word/theme/theme1.xml");
    expect(result.model.theme?.dirty).toBe(true);
    expect(result.model.theme?.name).toBe("Facet");
  });

  it("is pure: new model object, untouched blocks, unmutated input", () => {
    const model = createDocument("doc-3");
    const selection = caretSelection(documentStart());

    const result = applyDocumentTheme(model, FACET, selection);

    expect(result.model).not.toBe(model);
    // AC-18: no reference changed, no literal converted — the content is the
    // very same object graph, not a rewritten copy.
    expect(result.model.sections).toBe(model.sections);
    expect(result.model.defaultStyles).toBe(model.defaultStyles);
    expect(model.theme).toBeUndefined();
  });

  it("clones the theme rather than aliasing the caller's object", () => {
    const model = createDocument("doc-4");

    const result = applyDocumentTheme(model, FACET, caretSelection(documentStart()));

    expect(result.model.theme?.colors).not.toBe(FACET.colors);
    expect(result.model.theme?.fonts).not.toBe(FACET.fonts);
    expect(result.model.theme?.colors).toEqual(FACET.colors);
  });

  it("returns the selection it was given, normalised", () => {
    const model = createDocument("doc-5");
    const selection = caretSelection(documentStart());

    const result = applyDocumentTheme(model, FACET, selection);

    expect(result.selection).toEqual(selection);
  });
});
