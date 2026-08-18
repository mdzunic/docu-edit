import { describe, it, expect } from "vitest";
import {
  BUILT_IN_CAPTION_LABELS,
  captionDisplayText,
  captionForObject,
  captionRefText,
  isCaptionParagraph,
  isValidNewLabel,
  insertCaption,
  listCaptions,
  objectForCaption,
  renumberCaptions,
  defaultCaptionOptions,
} from "./captions";
import { CAPTION_STYLE_ID } from "./styles";
import type { DocumentModel, Image, Paragraph } from "./documentModel";
import { createDocument } from "./documentModel";

function image(mediaId: string): Image {
  return { type: "image", mediaId, widthCm: 5, heightCm: 5 } as Image;
}

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text ? [{ text, formatting: {} }] : [] };
}

function modelWith(blocks: DocumentModel["sections"][0]["blocks"]): DocumentModel {
  const doc = createDocument("Captions");
  doc.sections[0].blocks = blocks;
  return doc;
}

describe("captions (SPEC-055)", () => {
  it("offers Figure, Table and Equation as built-in labels (AC-5)", () => {
    expect(BUILT_IN_CAPTION_LABELS).toEqual(["Figure", "Table", "Equation"]);
  });

  it("rejects empty and case-insensitive duplicate new labels (AC-11, AC-12)", () => {
    expect(isValidNewLabel("", ["Figure"])).toBe(false);
    expect(isValidNewLabel("  ", ["Figure"])).toBe(false);
    expect(isValidNewLabel("figure", ["Figure"])).toBe(false);
    expect(isValidNewLabel("Diagram", ["Figure", "Table"])).toBe(true);
  });

  it("inserts a caption below an image with label/seq/separator/text runs (AC-3, AC-4)", () => {
    const model = modelWith([image("m1"), paragraph("after")]);
    const { model: next, selection } = insertCaption(model, { section: 0, block: 0 }, defaultCaptionOptions("Figure"));
    const caption = next.sections[0].blocks[1] as Paragraph;
    expect(caption.styleId).toBe(CAPTION_STYLE_ID);
    expect(caption.runs[0].text).toBe("Figure ");
    expect(caption.runs[1].field).toEqual({ kind: "seq", label: "Figure", format: "decimal", cached: "1" });
    expect(caption.runs[2].text).toBe(": ");
    expect(caption.runs[3].text).toBe("");
    // Caret placed in the editable text run.
    expect(selection.focus.offset).toBe("Figure ".length + 1 + ": ".length);
  });

  it("omits the label run when exclude-label is set (AC-4, AC-9)", () => {
    const model = modelWith([image("m1")]);
    const { model: next } = insertCaption(model, { section: 0, block: 0 }, {
      ...defaultCaptionOptions("Figure"),
      excludeLabel: true,
    });
    const caption = next.sections[0].blocks[1] as Paragraph;
    expect(caption.runs[0].field?.kind).toBe("seq");
  });

  it("inserts a caption above the object when position is above (AC-3, AC-7)", () => {
    const model = modelWith([paragraph("before"), image("m1")]);
    const { model: next } = insertCaption(model, { section: 0, block: 1 }, {
      ...defaultCaptionOptions("Figure"),
      position: "above",
    });
    expect(isCaptionParagraph(next.sections[0].blocks[1])).toBe(true);
    expect(next.sections[0].blocks[2]?.type).toBe("image");
  });

  it("numbers two labels independently in document order (AC-15, AC-17)", () => {
    let model = modelWith([image("a"), image("b"), image("c")]);
    // caption each image with Figure, then renumber to check sequence.
    model = insertCaption(model, { section: 0, block: 0 }, defaultCaptionOptions("Figure")).model;
    // block indices shift as captions insert below; re-find the images.
    const withTables = insertCaption(model, findImage(model, "b"), defaultCaptionOptions("Table")).model;
    const captions = listCaptions(withTables);
    const figure = captions.find((c) => c.label === "Figure");
    const table = captions.find((c) => c.label === "Table");
    expect(figure?.ordinal).toBe(1);
    expect(table?.ordinal).toBe(1);
  });

  it("renumbers all captions of a label in document order (AC-15, AC-43)", () => {
    const seq = (label: string, cached: string): Paragraph => ({
      type: "paragraph",
      styleId: CAPTION_STYLE_ID,
      runs: [
        { text: `${label} `, formatting: {} },
        { text: "1", formatting: {}, field: { kind: "seq", label, format: "decimal", cached } },
        { text: ": x", formatting: {} },
      ],
    });
    const model = modelWith([seq("Figure", "9"), seq("Figure", "9"), seq("Figure", "9")]);
    const next = renumberCaptions(model);
    const numbers = (next.sections[0].blocks as Paragraph[]).map(
      (p) => (p.runs[1].field as { cached: string }).cached,
    );
    expect(numbers).toEqual(["1", "2", "3"]);
  });

  it("returns the same model reference when nothing renumbers (AC-44)", () => {
    const already: Paragraph = {
      type: "paragraph",
      styleId: CAPTION_STYLE_ID,
      runs: [{ text: "1", formatting: {}, field: { kind: "seq", label: "Figure", format: "decimal", cached: "1" } }],
    };
    const model = modelWith([already]);
    expect(renumberCaptions(model)).toBe(model);
  });

  it("binds a caption to the adjacent object both ways (AC-20, §binding)", () => {
    const model = modelWith([image("m1"), paragraph("x")]);
    const inserted = insertCaption(model, { section: 0, block: 0 }, defaultCaptionOptions("Figure")).model;
    // image at 0, caption at 1
    expect(captionForObject(inserted, { section: 0, block: 0 })).toEqual({ section: 0, block: 1 });
    expect(objectForCaption(inserted, { section: 0, block: 1 })).toEqual({ section: 0, block: 0 });
  });

  it("breaks the binding when a paragraph separates caption and object (AC-40, AC-42)", () => {
    const seq: Paragraph = {
      type: "paragraph",
      styleId: CAPTION_STYLE_ID,
      runs: [
        { text: "Figure ", formatting: {} },
        { text: "1", formatting: {}, field: { kind: "seq", label: "Figure", format: "decimal", cached: "1" } },
      ],
    };
    const model = modelWith([image("m1"), paragraph("interrupting"), seq]);
    expect(captionForObject(model, { section: 0, block: 0 })).toBeNull();
    expect(objectForCaption(model, { section: 0, block: 2 })).toBeNull();
    // Still numbered and displayed correctly.
    expect(listCaptions(model)).toHaveLength(1);
  });

  it("renders each caption reference format (AC-24, Reference formats)", () => {
    const caption: Paragraph = {
      type: "paragraph",
      styleId: CAPTION_STYLE_ID,
      runs: [
        { text: "Figure ", formatting: {} },
        { text: "1", formatting: {}, field: { kind: "seq", label: "Figure", format: "decimal", cached: "3" } },
        { text: ": ", formatting: {} },
        { text: "Revenue by quarter", formatting: {} },
      ],
    };
    expect(captionDisplayText(caption)).toBe("Figure 3: Revenue by quarter");
    expect(captionRefText(caption, "full")).toBe("Figure 3: Revenue by quarter");
    expect(captionRefText(caption, "labelAndNumber")).toBe("Figure 3");
    expect(captionRefText(caption, "numberOnly")).toBe("3");
    expect(captionRefText(caption, "textOnly")).toBe("Revenue by quarter");
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
