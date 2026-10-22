import { describe, it, expect } from "vitest";
import {
  availableTofLabels,
  defaultTofOptions,
  insertTableOfFigures,
  tofEntriesFromCaptions,
  tofEntriesStale,
  tofPageNumbersStale,
  updateTofEntirely,
  updateTofPageNumbersOnly,
} from "./tof";
import { insertCaption, defaultCaptionOptions } from "./captions";
import { CAPTION_STYLE_ID } from "./styles";
import { isTof, createDocument, type DocumentModel, type Image, type Paragraph, type TableOfFigures } from "./documentModel";
import { caretSelection } from "./commands";

function image(mediaId: string): Image {
  return { type: "image", mediaId, widthCm: 5, heightCm: 5 } as Image;
}

function modelWith(blocks: DocumentModel["sections"][0]["blocks"]): DocumentModel {
  const doc = createDocument("ToF");
  doc.sections[0].blocks = blocks;
  return doc;
}

function withThreeFigures(): DocumentModel {
  let model = modelWith([image("a"), image("b"), image("c")]);
  // caption each image (indices shift; find images fresh each time)
  for (const id of ["a", "b", "c"]) {
    const at = findImage(model, id);
    model = insertCaption(model, at, defaultCaptionOptions("Figure")).model;
  }
  return model;
}

const surfaceCaption = (label: string, cached: string, text: string): Paragraph => ({
  type: "paragraph",
  styleId: CAPTION_STYLE_ID,
  runs: [
    { text: `${label} `, formatting: {} },
    { text: "1", formatting: {}, field: { kind: "seq", label, format: "decimal", cached } },
    { text: `: ${text}`, formatting: {} },
  ],
});

const pageFor = () => 2;

describe("table of figures (SPEC-055)", () => {
  it("lists a label's captions in document order with derived anchors (AC-29)", () => {
    const model = withThreeFigures();
    const entries = tofEntriesFromCaptions(model, "Figure");
    expect(entries).toHaveLength(3);
    expect(entries.map((e) => e.text)).toEqual(["Figure 1: ", "Figure 2: ", "Figure 3: "]);
    expect(new Set(entries.map((e) => e.targetBookmark)).size).toBe(3);
  });

  it("excludes header, footer and footnote captions from entries and labels (AC-39)", () => {
    const model = withThreeFigures();
    model.sections[0].headers = { default: { blocks: [surfaceCaption("Figure", "7", "Header chart")] } };
    model.sections[0].footers = { default: { blocks: [surfaceCaption("Chart", "1", "Footer chart")] } };
    model.notes = [{ noteId: "n1", kind: "footnote", blocks: [surfaceCaption("Figure", "9", "An archival chart")] }];
    // Only the three body figures are listed (§Numbering) — never the header
    // or footnote Figures, whatever their cached numbers say.
    const entries = tofEntriesFromCaptions(model, "Figure");
    expect(entries).toHaveLength(3);
    const texts = entries.map((entry) => entry.text).join("|");
    expect(texts).not.toContain("Header chart");
    expect(texts).not.toContain("archival");
    // The dialog's label list is body-derived too: the footer-only "Chart"
    // label is not offered.
    expect(availableTofLabels(model)).toEqual(["Figure"]);
  });

  it("offers the document's caption labels (AC-23)", () => {
    const model = withThreeFigures();
    expect(availableTofLabels(model)).toEqual(["Figure"]);
    expect(availableTofLabels(modelWith([image("a")]))).toEqual(["Figure"]);
  });

  it("a freshly generated table of figures is not stale (AC-32)", () => {
    const model = withThreeFigures();
    const { model: withTof } = insertTableOfFigures(
      model,
      caretSelection({ section: 0, block: model.sections[0].blocks.length - 1, offset: 0 }),
      "Figure",
      defaultTofOptions(),
      pageFor,
    );
    const tof = withTof.sections[0].blocks.find(isTof) as TableOfFigures;
    expect(tofEntriesStale(withTof, tof)).toBe(false);
    expect(tofPageNumbersStale(tof, pageFor)).toBe(false);
  });

  it("goes stale when a caption is added, and update entire refreshes it (AC-33, AC-36)", () => {
    const model = withThreeFigures();
    const { model: withTof } = insertTableOfFigures(
      model,
      caretSelection({ section: 0, block: model.sections[0].blocks.length - 1, offset: 0 }),
      "Figure",
      defaultTofOptions(),
      () => 1,
    );
    const tofIndex = withTof.sections[0].blocks.findIndex(isTof);
    // add another figure before the ToF
    let grown = withTof;
    grown.sections[0].blocks.splice(0, 0, image("d"));
    grown = insertCaption(grown, findImage(grown, "d"), defaultCaptionOptions("Figure")).model;
    const newIndex = grown.sections[0].blocks.findIndex(isTof);
    const staleTof = grown.sections[0].blocks[newIndex] as TableOfFigures;
    expect(tofEntriesStale(grown, staleTof)).toBe(true);
    const updated = updateTofEntirely(grown, caretSelection({ section: 0, block: newIndex, offset: 0 }), 0, newIndex, () => 1).model;
    const freshTof = updated.sections[0].blocks[newIndex] as TableOfFigures;
    expect(freshTof.entries).toHaveLength(4);
    expect(tofIndex).toBeGreaterThanOrEqual(0);
  });

  it("update page numbers only keeps text but refreshes pages", () => {
    const model = withThreeFigures();
    const { model: withTof } = insertTableOfFigures(
      model,
      caretSelection({ section: 0, block: model.sections[0].blocks.length - 1, offset: 0 }),
      "Figure",
      defaultTofOptions(),
      () => 1,
    );
    const index = withTof.sections[0].blocks.findIndex(isTof);
    const updated = updateTofPageNumbersOnly(withTof, caretSelection({ section: 0, block: index, offset: 0 }), 0, index, () => 5).model;
    const tof = updated.sections[0].blocks[index] as TableOfFigures;
    expect(tof.entries.every((e) => e.page === 5)).toBe(true);
    expect(tof.entries).toHaveLength(3);
  });
});

function findImage(model: DocumentModel, mediaId: string): { section: number; block: number } {
  for (let s = 0; s < model.sections.length; s++) {
    const blocks = model.sections[s].blocks;
    for (let b = 0; b < blocks.length; b++) {
      const block = blocks[b];
      if (block.type === "image" && block.mediaId === mediaId) {
        return { section: s, block: b };
      }
    }
  }
  throw new Error(`image ${mediaId} not found`);
}
