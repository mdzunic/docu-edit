// SPEC-094 AC: the deck-theme commands — apply/reset, which slides dirty per
// format, and refs untouched by apply.
import { describe, expect, it } from "vitest";
import { BUILT_IN_THEMES, DEFAULT_THEME } from "../core/theme";
import type { PresentationModel, Slide } from "./model";
import { applyTheme, cloneDeckTheme, deckThemeOf, setThemeColor, setThemeFont, slideUsesTheme } from "./theme";

function slideWithRef(): Slide {
  return {
    id: "s1",
    frames: [
      {
        id: "f1",
        kind: "shape",
        geometry: "rect",
        xCm: 1,
        yCm: 1,
        widthCm: 4,
        heightCm: 3,
        fill: { type: "solid", color: { theme: "accent1" } },
      },
    ],
    partName: "ppt/slides/slide1.xml",
    sourceXml: "<draw:page/>",
  };
}

function literalSlide(): Slide {
  return {
    id: "s2",
    frames: [
      {
        id: "f2",
        kind: "text",
        xCm: 1,
        yCm: 1,
        widthCm: 5,
        heightCm: 2,
        paragraphs: [{ runs: [{ text: "hi", color: "#ff0000" }] }],
      },
    ],
    partName: "ppt/slides/slide2.xml",
    sourceXml: "<draw:page/>",
  };
}

function deck(format: "pptx" | "odp"): PresentationModel {
  return {
    kind: "presentation",
    id: "d",
    title: "Deck",
    format,
    slides: [slideWithRef(), literalSlide()],
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
    theme: { ...DEFAULT_THEME, colorMap: { tx1: "dk1" }, partName: "ppt/theme/theme1.xml" },
  };
}

describe("deckThemeOf", () => {
  it("falls back to Office for an unthemed deck (AC-7)", () => {
    expect(deckThemeOf({ theme: undefined })).toBe(DEFAULT_THEME);
    expect(deckThemeOf(null)).toBe(DEFAULT_THEME);
  });
});

describe("applyTheme", () => {
  it("replaces the scheme, keeps the package wiring, marks the theme dirty", () => {
    const model = deck("pptx");
    applyTheme(model, BUILT_IN_THEMES[1]);
    expect(model.theme?.name).toBe("Facet");
    expect(model.theme?.colors.accent1).toBe(BUILT_IN_THEMES[1].colors.accent1);
    expect(model.theme?.dirty).toBe(true);
    expect(model.theme?.colorMap).toEqual({ tx1: "dk1" });
    expect(model.theme?.partName).toBe("ppt/theme/theme1.xml");
  });

  it("dirties NO slide on pptx — refs do not change (AC-28)", () => {
    const model = deck("pptx");
    applyTheme(model, BUILT_IN_THEMES[1]);
    expect(model.slides.every((slide) => slide.dirty === undefined)).toBe(true);
  });

  it("marks only ref-carrying slides edited on odp (AC-33)", () => {
    const model = deck("odp");
    applyTheme(model, BUILT_IN_THEMES[1]);
    expect(model.slides[0].dirty).toBe(true);
    expect(model.slides[0].sourceXml).toBeUndefined(); // markSlideEdited drops it
    expect(model.slides[1].dirty).toBeUndefined();
    expect(model.slides[1].sourceXml).toBe("<draw:page/>");
  });

  it("leaves every ref and literal exactly as it was (AC-38)", () => {
    const model = deck("pptx");
    const before = JSON.parse(JSON.stringify(model.slides));
    applyTheme(model, BUILT_IN_THEMES[2]);
    expect(model.slides).toEqual(before);
  });

  it("reset = re-applying the stashed pristine theme restores it", () => {
    const model = deck("pptx");
    const pristine = cloneDeckTheme(model.theme!);
    setThemeColor(model, "accent1", "#123456");
    expect(model.theme?.colors.accent1).toBe("#123456");
    applyTheme(model, pristine);
    expect(model.theme?.colors).toEqual(pristine.colors);
    expect(model.theme?.fonts).toEqual(pristine.fonts);
  });
});

describe("setThemeColor / setThemeFont", () => {
  it("edits one slot / one font through the same dirty rules", () => {
    const model = deck("pptx");
    setThemeColor(model, "accent2", "#0000aa");
    expect(model.theme?.colors.accent2).toBe("#0000aa");
    expect(model.theme?.dirty).toBe(true);
    setThemeFont(model, "major", "Georgia");
    expect(model.theme?.fonts.major).toBe("Georgia");
  });
});

describe("slideUsesTheme", () => {
  it("detects refs in fills, strokes, runs and nested groups", () => {
    expect(slideUsesTheme(slideWithRef())).toBe(true);
    expect(slideUsesTheme(literalSlide())).toBe(false);
    const grouped: Slide = {
      id: "g",
      frames: [
        {
          id: "grp",
          kind: "group",
          xCm: 0,
          yCm: 0,
          widthCm: 5,
          heightCm: 5,
          xml: "<p:grpSp/>",
          children: [
            {
              id: "c",
              kind: "text",
              xCm: 0,
              yCm: 0,
              widthCm: 2,
              heightCm: 2,
              paragraphs: [{ runs: [{ text: "t", fontFamily: { theme: "minor" } }] }],
            },
          ],
        },
      ],
    };
    expect(slideUsesTheme(grouped)).toBe(true);
  });
});
