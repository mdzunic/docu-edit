import { describe, expect, it } from "vitest";
import { createDocument, type DocumentModel, type NamedStyle } from "./documentModel";
import {
  builtInStyles,
  diffParagraphFormatting,
  diffRunFormatting,
  headingStyleId,
  inheritedParagraphPaint,
  namedStyleIdFor,
  paintedParagraphFormatting,
  paragraphPaintResolver,
  NORMAL_STYLE_ID,
  paragraphStyleBase,
  resolveNamedStyle,
  styleById,
  styleChainIncludes,
  styleDisplayName,
  stylePickerEntries,
  tabContextFor,
} from "./styles";
import type { Paragraph, TabStop } from "./documentModel";

function modelWith(styles?: NamedStyle[]): DocumentModel {
  const model = createDocument("styles-fixture");
  if (styles) {
    model.styles = styles;
  }
  return model;
}

describe("SPEC-019 — built-in styles", () => {
  it("offers Normal, Title, Subtitle, Heading 1-6, Quote and Code (AC-2)", () => {
    expect(builtInStyles().map((style) => style.name)).toEqual([
      "Normal",
      "Title",
      "Subtitle",
      "Heading 1",
      "Heading 2",
      "Heading 3",
      "Heading 4",
      "Heading 5",
      "Heading 6",
      "Quote",
      "Code",
    ]);
  });

  it("heading styles carry Word's built-in ids and their level (AC-15)", () => {
    for (let level = 1; level <= 6; level++) {
      expect(headingStyleId(level)).toBe(`Heading${level}`);
      expect(resolveNamedStyle(modelWith(), `Heading${level}`).headingLevel).toBe(level);
    }
  });

  it("resolves a built-in for a document without a catalogue of its own", () => {
    const resolved = resolveNamedStyle(modelWith(), "Heading1");
    expect(resolved.runFormatting.bold).toBe(true);
    expect(resolved.runFormatting.fontSizePt).toBe(20);
  });
});

describe("SPEC-019 — resolution through basedOn chains (AC-7)", () => {
  it("merges the chain root-first, the child overriding the parent", () => {
    const model = modelWith([
      { id: "Base", name: "Base", runFormatting: { bold: true, fontSizePt: 14 } },
      { id: "Child", name: "Child", basedOn: "Base", runFormatting: { italic: true, fontSizePt: 18 } },
    ]);
    expect(resolveNamedStyle(model, "Child").runFormatting).toEqual({
      bold: true,
      italic: true,
      fontSizePt: 18,
    });
  });

  it("a document's own definition shadows the built-in of the same id", () => {
    const model = modelWith([
      { id: "Heading1", name: "Heading 1", runFormatting: { fontSizePt: 30 } },
    ]);
    const resolved = resolveNamedStyle(model, "Heading1");
    expect(resolved.runFormatting).toEqual({ fontSizePt: 30 });
  });

  it("survives a basedOn cycle", () => {
    const model = modelWith([
      { id: "A", name: "A", basedOn: "B", runFormatting: { bold: true } },
      { id: "B", name: "B", basedOn: "A", runFormatting: { italic: true } },
    ]);
    const resolved = resolveNamedStyle(model, "A");
    expect(resolved.runFormatting.bold).toBe(true);
    expect(resolved.runFormatting.italic).toBe(true);
  });

  it("resolves a percentage font size against the inherited, then default size", () => {
    const model = modelWith([
      { id: "Big", name: "Big", fontSizePercent: 200 },
      { id: "Bigger", name: "Bigger", basedOn: "Sized", fontSizePercent: 150 },
      { id: "Sized", name: "Sized", runFormatting: { fontSizePt: 12 } },
    ]);
    // 11 pt document default * 200%.
    expect(resolveNamedStyle(model, "Big").runFormatting.fontSizePt).toBe(22);
    // 12 pt inherited * 150%.
    expect(resolveNamedStyle(model, "Bigger").runFormatting.fontSizePt).toBe(18);
  });

  it("drops an ltr direction a chain resolves to", () => {
    const model = modelWith([
      { id: "Rtl", name: "Rtl", paragraphFormatting: { direction: "rtl" } },
      { id: "BackLtr", name: "BackLtr", basedOn: "Rtl", paragraphFormatting: { direction: "ltr" } },
    ]);
    expect(resolveNamedStyle(model, "BackLtr").paragraphFormatting.direction).toBeUndefined();
    expect(resolveNamedStyle(model, "Rtl").paragraphFormatting.direction).toBe("rtl");
  });
});

describe("SPEC-054 — tabContextFor (AC-27, AC-28)", () => {
  const styleStops: TabStop[] = [{ positionCm: 4, alignment: "right", leader: "dot" }];

  function modelWithTabStyle(): DocumentModel {
    return modelWith([
      { id: "Figures", name: "Figures", paragraphFormatting: { tabStops: styleStops } },
    ]);
  }

  it("inherits the paragraph style's tab stops when the paragraph sets none (AC-27)", () => {
    const model = modelWithTabStyle();
    const paragraph: Paragraph = { type: "paragraph", runs: [], styleId: "Figures" };
    expect(tabContextFor(model, paragraph)).toEqual(styleStops);
  });

  it("returns the paragraph's own stops, replacing (not merging) the inherited list (AC-28)", () => {
    const model = modelWithTabStyle();
    const ownStops: TabStop[] = [{ positionCm: 2, alignment: "left" }];
    const paragraph: Paragraph = {
      type: "paragraph",
      runs: [],
      styleId: "Figures",
      formatting: { tabStops: ownStops },
    };
    expect(tabContextFor(model, paragraph)).toEqual(ownStops);
  });

  it("resolves to no stops for an unstyled paragraph with none of its own", () => {
    const model = modelWith();
    expect(tabContextFor(model, { type: "paragraph", runs: [] })).toEqual([]);
  });
});

describe("SPEC-019 — named style lookup and display names (AC-1)", () => {
  it("walks nameless entries to the nearest named ancestor", () => {
    const model = modelWith([
      { id: "P1", basedOn: "Quote Style" },
      { id: "Quote Style", name: "Quote Style", runFormatting: { italic: true } },
    ]);
    expect(namedStyleIdFor(model, "P1")).toBe("Quote Style");
    expect(styleDisplayName(model, "P1")).toBe("Quote Style");
  });

  it("shows Normal for unstyled paragraphs and unknown references", () => {
    expect(styleDisplayName(modelWith(), undefined)).toBe("Normal");
    expect(styleDisplayName(modelWith(), "NoSuchStyle")).toBe("Normal");
  });

  it("falls back to the built-in name for an id the document does not define", () => {
    expect(styleDisplayName(modelWith(), "Heading2")).toBe("Heading 2");
    expect(styleById(modelWith(), "Quote")?.name).toBe("Quote");
  });
});

describe("SPEC-019 — picker entries (AC-2, AC-8)", () => {
  it("lists the built-in set for a blank document, previews in each style's formatting", () => {
    const entries = stylePickerEntries(modelWith());
    expect(entries.map((entry) => entry.name)).toEqual([
      "Normal",
      "Title",
      "Subtitle",
      "Heading 1",
      "Heading 2",
      "Heading 3",
      "Heading 4",
      "Heading 5",
      "Heading 6",
      "Quote",
      "Code",
    ]);
    const heading1 = entries.find((entry) => entry.id === "Heading1")!;
    expect(heading1.previewStyle.fontWeight).toBe("bold");
    expect(heading1.previewStyle.fontSize).toBe("20pt");
    expect(heading1.headingLevel).toBe(1);
  });

  it("lists the document's own styles instead of the hard-coded set", () => {
    const model = modelWith([
      { id: "Normal", name: "Normal" },
      { id: "Fancy", name: "Fancy", runFormatting: { italic: true } },
    ]);
    expect(stylePickerEntries(model).map((entry) => entry.id)).toEqual(["Normal", "Fancy"]);
  });

  it("always offers Normal even when the document does not define one", () => {
    const model = modelWith([{ id: "OnlyStyle", name: "Only Style" }]);
    expect(stylePickerEntries(model).map((entry) => entry.id)).toEqual([
      NORMAL_STYLE_ID,
      "OnlyStyle",
    ]);
  });

  it("skips nameless resolution-only entries", () => {
    const model = modelWith([
      { id: "Normal", name: "Normal" },
      { id: "P7", basedOn: "Normal" },
    ]);
    expect(stylePickerEntries(model).map((entry) => entry.id)).toEqual(["Normal"]);
  });
});

describe("SPEC-019 — direct-formatting diffs (AC-9, AC-14)", () => {
  it("keeps only what the effective formatting adds or changes", () => {
    expect(
      diffRunFormatting(
        { bold: true, italic: true, fontSizePt: 20 },
        { bold: true, fontSizePt: 16 },
      ),
    ).toEqual({ italic: true, fontSizePt: 20 });
    expect(diffRunFormatting({ bold: true }, { bold: true })).toEqual({});
  });

  it("records an explicit off where the base sets a toggle the effective lacks", () => {
    expect(diffRunFormatting({}, { bold: true, highlight: "yellow" })).toEqual({
      bold: false,
      highlight: false,
    });
  });

  it("diffs paragraph formatting per property", () => {
    expect(
      diffParagraphFormatting(
        { alignment: "center", spaceBeforePt: 20 },
        { spaceBeforePt: 20, spaceAfterPt: 6 },
      ),
    ).toEqual({ alignment: "center" });
  });
});

describe("SPEC-019 — chain membership and style bases (AC-13)", () => {
  it("styleChainIncludes covers the style itself and its descendants", () => {
    const model = modelWith([
      { id: "Base", name: "Base" },
      { id: "Child", name: "Child", basedOn: "Base" },
      { id: "Other", name: "Other" },
    ]);
    expect(styleChainIncludes(model, "Base", "Base")).toBe(true);
    expect(styleChainIncludes(model, "Child", "Base")).toBe(true);
    expect(styleChainIncludes(model, "Other", "Base")).toBe(false);
    expect(styleChainIncludes(model, undefined, "Base")).toBe(false);
  });

  it("paragraphStyleBase resolves the nearest named style, or nothing", () => {
    const model = modelWith([
      { id: "P1", basedOn: "Quote Style" },
      { id: "Quote Style", name: "Quote Style", runFormatting: { italic: true } },
    ]);
    expect(
      paragraphStyleBase(model, { type: "paragraph", runs: [], styleId: "P1" }).runFormatting,
    ).toEqual({ italic: true });
    expect(paragraphStyleBase(model, { type: "paragraph", runs: [] }).runFormatting).toEqual({});
  });
});

// SPEC-074 AC-17 — borders and shading are inherited from the paragraph style
// until the paragraph sets its own. Unlike the rest of ParagraphFormatting,
// they are not baked onto the paragraph: they resolve here, which is the only
// way a paragraph that names NO style (the implicit Normal every blank
// document's paragraphs have) can inherit anything at all.

const STYLE_BOX = { top: { style: "single" as const, widthPt: 1, color: "#000000" } };
const OWN_BOX = { bottom: { style: "thick" as const, widthPt: 3, color: "#c00000" } };

function normalWith(paragraphFormatting: Record<string, unknown>): DocumentModel {
  return modelWith([
    { id: NORMAL_STYLE_ID, name: "Normal", paragraphFormatting },
    { id: "Callout", name: "Callout", basedOn: NORMAL_STYLE_ID },
  ] as NamedStyle[]);
}

function plainParagraph(styleId?: string): Paragraph {
  return { type: "paragraph", runs: [{ text: "x", formatting: {} }], ...(styleId ? { styleId } : {}) };
}

describe("SPEC-074 AC-17 — borders and shading inherited from the style", () => {
  it("resolves an absent styleId through Normal", () => {
    const model = normalWith({ borders: STYLE_BOX, shadingFill: "#eeeeee" });
    expect(inheritedParagraphPaint(model, undefined)).toEqual({
      borders: STYLE_BOX,
      shadingFill: "#eeeeee",
    });
  });

  it("carries only borders and shading over — the rest is already baked", () => {
    const model = normalWith({ borders: STYLE_BOX, spaceBeforePt: 18, alignment: "center" });
    expect(inheritedParagraphPaint(model, NORMAL_STYLE_ID)).toEqual({ borders: STYLE_BOX });
  });

  it("follows the basedOn chain", () => {
    const model = normalWith({ shadingFill: "#fff2cc" });
    expect(inheritedParagraphPaint(model, "Callout").shadingFill).toBe("#fff2cc");
  });

  it("paints a paragraph that names no style with Normal's box — the QA case", () => {
    const model = normalWith({ borders: STYLE_BOX });
    expect(paintedParagraphFormatting(model, plainParagraph())?.borders).toEqual(STYLE_BOX);
  });

  it("lets the paragraph's own borders win over the style's", () => {
    const model = normalWith({ borders: STYLE_BOX, shadingFill: "#eeeeee" });
    const paragraph = plainParagraph();
    paragraph.formatting = { borders: OWN_BOX };
    const painted = paintedParagraphFormatting(model, paragraph);
    expect(painted?.borders).toEqual(OWN_BOX);
    // Shading is a separate property: the paragraph set no fill, so it still
    // takes the style's.
    expect(painted?.shadingFill).toBe("#eeeeee");
  });

  it("says nothing about a document whose styles carry no borders", () => {
    const model = modelWith();
    const paragraph = plainParagraph();
    paragraph.formatting = { alignment: "center" };
    expect(paintedParagraphFormatting(model, paragraph)).toBe(paragraph.formatting);
  });

  it("memoises the resolver per style id and answers each paragraph in kind", () => {
    const model = normalWith({ borders: STYLE_BOX });
    (model.styles as NamedStyle[]).push({
      id: "Bare",
      name: "Bare",
      paragraphFormatting: { spaceBeforePt: 12 },
    });
    const resolve = paragraphPaintResolver(model);
    expect(resolve(plainParagraph())).toEqual({ borders: STYLE_BOX });
    expect(resolve(plainParagraph("Callout"))).toEqual({ borders: STYLE_BOX });
    // A style with a parent of its own, not based on Normal, inherits nothing.
    expect(resolve(plainParagraph("Bare"))).toEqual({});
    // Same id, same object: the chain is walked once.
    expect(resolve(plainParagraph())).toBe(resolve(plainParagraph()));
  });
});
