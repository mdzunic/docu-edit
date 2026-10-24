// SPEC-098 — the deck's checking default derived at read, and a run's
// effective checking language through core's chain.
import { describe, expect, it } from "vitest";
import { DEFAULT_LANGUAGE } from "../core/language";
import type { Slide, SlideFrame, SlideParagraph } from "./model";
import { deckLanguageOf, runLanguage } from "./slideLanguage";

const box = { xCm: 0, yCm: 0, widthCm: 1, heightCm: 1 };
const para = (...langs: (string | undefined)[]): SlideParagraph => ({ runs: langs.map((lang) => ({ text: "x", ...(lang ? { lang } : {}) })) });
const text = (id: string, ...paragraphs: SlideParagraph[]): SlideFrame => ({ kind: "text", id, ...box, paragraphs });

describe("deckLanguageOf", () => {
  it("is the most frequent run tag across frames, shapes, table cells and group children", () => {
    const slides: Slide[] = [
      { id: "a", frames: [text("t", para("en-US", "de-DE")), { kind: "shape", id: "s", ...box, geometry: "rect", paragraphs: [para("de-DE")] }] },
      {
        id: "b",
        frames: [
          { kind: "table", id: "tb", ...box, columnWidthsCm: [1], rows: [{ cells: [{ paragraphs: [para("de-DE")] }] }] },
          { kind: "group", id: "g", ...box, children: [text("c", para("en-US"))] },
        ],
      },
    ];
    expect(deckLanguageOf(slides)).toBe("de-DE");
  });

  it("breaks a tie in favour of the tag met first, and is undefined when no run declares one", () => {
    expect(deckLanguageOf([{ id: "a", frames: [text("t", para("fr-FR", "es-ES"))] }])).toBe("fr-FR");
    expect(deckLanguageOf([{ id: "a", frames: [text("t", para(undefined, undefined))] }])).toBeUndefined();
    expect(deckLanguageOf([])).toBeUndefined();
  });

  it("keeps tags verbatim — never re-normalised", () => {
    expect(deckLanguageOf([{ id: "a", frames: [text("t", para("en-gb"))] }])).toBe("en-gb");
  });
});

describe("runLanguage", () => {
  it("is the run's tag, else the deck's, else DEFAULT_LANGUAGE", () => {
    expect(runLanguage({ text: "x", lang: "cs-CZ" }, { language: "de-DE" })).toBe("cs-CZ");
    expect(runLanguage({ text: "x" }, { language: "de-DE" })).toBe("de-DE");
    expect(runLanguage({ text: "x" }, {})).toBe(DEFAULT_LANGUAGE);
  });
});
