// SPEC-115 — the app layer builds lib/design's product-neutral OutlineSource:
// from a text document (SPEC-037's outline headings and the paragraphs
// between them), from a deck (its title, theme and slides — never its notes),
// and from pasted text (blank-line-separated blocks); and renders one as the
// plain text a design request carries.
import { describe, expect, it } from "vitest";
import { createDocument, type Block, type Paragraph } from "../../lib/documentModel";
import { synthesiseStandardLayouts } from "../../lib/presentation/layouts";
import type { TextFrame } from "../../lib/presentation/model";
import { createBlankPresentation, titleBodyFrames } from "../../lib/presentation/slides";
import { BUILT_IN_THEMES } from "../../lib/core/theme";
import { outlineFromDeck, outlineFromDocument, outlineFromText, outlineText } from "./outlineSource";

function paragraph(text: string, extra: Partial<Paragraph> = {}): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }], ...extra };
}

describe("outlineFromDocument", () => {
  it("takes the outline's headings and the paragraphs between them, with list depth; empty paragraphs and tables are skipped", () => {
    const doc = createDocument("d");
    doc.title = "Plan";
    const table = { type: "table", rows: [] } as unknown as Block;
    doc.sections = [
      {
        blocks: [
          paragraph("Intro text"),
          paragraph("Goals", { headingLevel: 1 }),
          paragraph(""),
          paragraph("Ship it", { list: { type: "bullet", level: 1 } }),
          paragraph("Then iterate", { list: { type: "bullet", level: 2 } }),
          table,
          paragraph("Detail", { headingLevel: 3 }),
        ],
      },
      { blocks: [paragraph("Later", { headingLevel: 2 })] },
    ];
    expect(outlineFromDocument(doc)).toEqual({
      title: "Plan",
      blocks: [
        { kind: "paragraph", text: "Intro text" },
        { kind: "heading", level: 1, text: "Goals" },
        { kind: "paragraph", text: "Ship it", listLevel: 0 },
        { kind: "paragraph", text: "Then iterate", listLevel: 1 },
        { kind: "heading", level: 3, text: "Detail" },
        { kind: "heading", level: 2, text: "Later" },
      ],
    });
  });
});

describe("outlineFromDeck", () => {
  it("reads each slide's title (a section header's as a Heading 1) and its other text — never its notes", () => {
    const deck = createBlankPresentation();
    deck.title = "Q3";
    deck.theme = { ...BUILT_IN_THEMES[2] };
    deck.layouts = synthesiseStandardLayouts(deck.slideWidthCm, deck.slideHeightCm);
    const [title, body] = titleBodyFrames(deck.slideWidthCm, deck.slideHeightCm);
    title.paragraphs = [{ runs: [{ text: "Part one" }] }];
    body.paragraphs = [{ runs: [{ text: "Why" }] }, { runs: [{ text: "Because" }], level: 1 }];
    const [untitled] = titleBodyFrames(deck.slideWidthCm, deck.slideHeightCm);
    deck.slides = [
      { id: "a", frames: [title, body], layoutKey: "standard:sectionHeader", notes: ["secret"] },
      { id: "b", frames: [untitled as TextFrame] },
    ];
    expect(outlineFromDeck(deck)).toEqual({
      title: "Q3",
      theme: BUILT_IN_THEMES[2].name,
      blocks: [
        { kind: "heading", level: 1, text: "Part one" },
        { kind: "paragraph", text: "Why", listLevel: 0 },
        { kind: "paragraph", text: "Because", listLevel: 1 },
        { kind: "heading", level: 2, text: "Slide 2" },
      ],
    });
  });
});

describe("outlineFromText", () => {
  it("reads blank-line-separated blocks: a short first line is the title, each block's first line its heading", () => {
    const text = [
      "Rotterdam migration",
      "",
      "Status",
      "Phase one done",
      "- Network cut-over",
      "  - booked for May",
      "",
      "# Risks",
      "",
      "This single line is a long paragraph rather than a heading because it has far more than twelve words in it.",
      "",
      "Questions",
    ].join("\r\n");
    expect(outlineFromText(text)).toEqual({
      title: "Rotterdam migration",
      blocks: [
        { kind: "heading", level: 2, text: "Status" },
        { kind: "paragraph", text: "Phase one done" },
        { kind: "paragraph", text: "Network cut-over", listLevel: 0 },
        { kind: "paragraph", text: "booked for May", listLevel: 1 },
        { kind: "heading", level: 1, text: "Risks" },
        {
          kind: "paragraph",
          text: "This single line is a long paragraph rather than a heading because it has far more than twelve words in it.",
        },
        { kind: "heading", level: 2, text: "Questions" },
      ],
    });
  });

  it("has no title when the text does not open with a short line on its own", () => {
    expect(outlineFromText("Status\nOn track").title).toBe("");
    expect(outlineFromText("").blocks).toEqual([]);
  });
});

describe("outlineText", () => {
  it("is the plain text a request carries: headings marked with #, list items with - at their depth", () => {
    expect(
      outlineText({
        title: "T",
        blocks: [
          { kind: "heading", level: 1, text: "A" },
          { kind: "paragraph", text: "plain" },
          { kind: "paragraph", text: "item", listLevel: 1 },
          { kind: "heading", level: 9, text: "deep" },
        ],
      }),
    ).toBe("# A\nplain\n  - item\n###### deep");
  });
});
