import { describe, expect, it } from "vitest";
import { compileQuery, type CompiledQuery } from "../core/findQuery";
import { buildPptxFixtureBytes } from "./fixtures/build";
import { readPresentationBytes } from "./open";
import { findInDeck } from "./deckSearch";
import { deckText, deckTextSlide } from "./deckText";
import type { PresentationModel, Slide, SlideFrame, SlideParagraph } from "./model";

const para = (text: string): SlideParagraph => ({ runs: [{ text }] });
const box = { xCm: 0, yCm: 0, widthCm: 10, heightCm: 2 };
const textFrame = (id: string, lines: string[], role?: "title" | "body"): SlideFrame => ({
  kind: "text",
  id,
  ...box,
  paragraphs: lines.map(para),
  ...(role ? { role } : {}),
});
const slide = (id: string, frames: SlideFrame[], notes?: string[]): Slide => ({ id, frames, ...(notes ? { notes } : {}) });
const deck = (slides: Slide[]): PresentationModel => ({
  kind: "presentation",
  id: "deck",
  title: "Deck",
  format: "pptx",
  slides,
  slideWidthCm: 33.867,
  slideHeightCm: 19.05,
});

describe("deckText", () => {
  it("takes each slide's title first, then its paragraphs in z-order, then its notes", () => {
    const model = deck([
      slide("s1", [textFrame("body", ["First point", "Second point"], "body"), textFrame("title", ["Roadmap"], "title")], ["Speaker note"]),
      slide("s2", [
        { kind: "shape", id: "sh", ...box, geometry: "rect", paragraphs: [para("In a shape")] },
        {
          kind: "table",
          id: "tb",
          ...box,
          columnWidthsCm: [2, 2],
          rows: [
            { cells: [{ paragraphs: [para("A1")] }, { paragraphs: [para("B1")] }] },
            { cells: [{ paragraphs: [para("A2")] }, { paragraphs: [], covered: true }] },
          ],
        },
        { kind: "group", id: "g", ...box, xml: "", children: [textFrame("inner", ["Grouped words"])] } as unknown as SlideFrame,
      ]),
    ]);
    const extracted = deckText(model);
    expect(extracted.coverage).toBe("full");
    expect(extracted.text.split("\n")).toEqual([
      "Roadmap",
      "First point",
      "Second point",
      "Speaker note",
      "In a shape",
      "A1",
      "B1",
      "A2",
      "Grouped words",
    ]);
  });

  it("names the slide an offset came from — the slide a deck result opens at", () => {
    const model = deck([
      slide("s1", [textFrame("t", ["Opening"], "title")]),
      slide("s2", []),
      slide("s3", [textFrame("t", ["Berlin office"], "title")], ["Mention the lease"]),
    ]);
    const extracted = deckText(model);
    expect(deckTextSlide(extracted, extracted.text.indexOf("Berlin"))).toBe(2);
    expect(deckTextSlide(extracted, extracted.text.indexOf("lease"))).toBe(2);
    expect(deckTextSlide(extracted, 0)).toBe(0);
    expect(deckTextSlide(extracted, 999)).toBe(-1);
    // The same slide the deck's own find (SPEC-098) lands on.
    const hits = findInDeck(model, compileQuery("Berlin", { matchCase: false, wholeWords: false, regex: false }) as CompiledQuery);
    expect(hits[0].slideIndex).toBe(2);
  });

  it("reads an imported .pptx through the import pipeline's converted model", async () => {
    const model = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "mini-deck");
    const extracted = deckText(model);
    expect(extracted.coverage).toBe("full");
    expect(extracted.text).toContain("Quarterly Review");
    expect(extracted.text).toContain("Open Berlin office");
    expect(extracted.text).toContain("Open with the Q3 highlight");
    expect(deckTextSlide(extracted, extracted.text.indexOf("Berlin"))).toBe(1);
  });
});
