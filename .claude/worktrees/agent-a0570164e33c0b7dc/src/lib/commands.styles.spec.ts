import { describe, expect, it } from "vitest";
import {
  applyParagraphStyle,
  caretSelection,
  resetParagraphStyle,
  toolbarStyleState,
  updateStyleFromSelection,
  type ModelSelection,
} from "./commands";
import { asParagraph, createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import type { CaretPosition } from "./textEditing";

// SPEC-019 — the paragraph-style commands. The model keeps *effective*
// formatting baked on paragraphs and runs; these tests pin the layer
// separation: applying a style keeps direct formatting on top (AC-9),
// resetting removes exactly that layer (AC-14), and updating a style
// re-bakes every user of it (AC-13).

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

function caret(block: number, offset: number): CaretPosition {
  return { section: 0, block, offset };
}

function range(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

function blocks(model: DocumentModel): Paragraph[] {
  return model.sections[0].blocks.map((block) => asParagraph(block));
}

function fixture(): DocumentModel {
  const model = createDocument("styles-commands");
  model.sections = [
    {
      blocks: [
        paragraph(["First paragraph", {}]),
        paragraph(["Second ", {}], ["bold", { bold: true }]),
        paragraph(["Third paragraph", {}]),
      ],
    },
  ];
  return model;
}

describe("applyParagraphStyle (AC-3, AC-5, AC-6, AC-9)", () => {
  it("applies the style to every paragraph in the selection", () => {
    const model = fixture();
    const { model: next } = applyParagraphStyle(model, range(caret(0, 2), caret(1, 3)), "Heading2");
    const [first, second, third] = blocks(next);
    expect(first.styleId).toBe("Heading2");
    expect(first.headingLevel).toBe(2);
    expect(first.runs[0].formatting).toEqual({ bold: true, fontSizePt: 16 });
    expect(first.formatting).toEqual({ spaceBeforePt: 18, spaceAfterPt: 6 });
    expect(second.styleId).toBe("Heading2");
    // The untouched paragraph keeps its identity (v-memo relies on it).
    expect(third).toBe(asParagraph(model.sections[0].blocks[2]));
  });

  it("keeps direct formatting layered on top of the new style", () => {
    const model = fixture();
    const { model: next } = applyParagraphStyle(model, caretSelection(caret(1, 0)), "Heading1");
    const second = blocks(next)[1];
    // The bold run was bold before the style: still bold, now heading-sized.
    expect(second.runs[1].formatting).toEqual({ bold: true, fontSizePt: 20 });
  });

  it("keeps direct formatting the style would otherwise impose off", () => {
    const model = fixture();
    // Apply Heading1 (bold), then explicitly un-bold via a later style
    // change: direct formatting recorded relative to the old style carries
    // the un-bold across.
    const first = applyParagraphStyle(model, caretSelection(caret(0, 0)), "Heading1").model;
    const unbolded = structuredClone(first);
    delete asParagraph(unbolded.sections[0].blocks[0]).runs[0].formatting.bold;
    const { model: next } = applyParagraphStyle(unbolded, caretSelection(caret(0, 0)), "Heading2");
    expect(blocks(next)[0].runs[0].formatting.bold).toBe(false);
  });

  it("applying Normal removes the heading level again", () => {
    const model = fixture();
    const styled = applyParagraphStyle(model, caretSelection(caret(0, 0)), "Heading1").model;
    const { model: next } = applyParagraphStyle(styled, caretSelection(caret(0, 0)), "Normal");
    const first = blocks(next)[0];
    expect(first.styleId).toBe("Normal");
    expect(first.headingLevel).toBeUndefined();
    expect(first.runs[0].formatting).toEqual({});
  });

  it("materializes the applied style into the catalogue", () => {
    const model = fixture();
    // A blank document gets the whole built-in set...
    const { model: next } = applyParagraphStyle(model, caretSelection(caret(0, 0)), "Heading1");
    expect(next.styles?.map((style) => style.id)).toContain("Heading1");
    expect(next.styles?.map((style) => style.id)).toContain("Quote");
    // ...a document with its own catalogue gets just the applied built-in.
    const withCatalogue = fixture();
    withCatalogue.styles = [{ id: "Normal", name: "Normal" }];
    const applied = applyParagraphStyle(withCatalogue, caretSelection(caret(0, 0)), "Heading3").model;
    expect(applied.styles?.map((style) => style.id)).toEqual(["Normal", "Heading3"]);
    // The input model is never mutated (the undo contract).
    expect(withCatalogue.styles.map((style) => style.id)).toEqual(["Normal"]);
  });

  it("resolves the document's own definition over the built-in", () => {
    const model = fixture();
    model.styles = [
      { id: "Normal", name: "Normal" },
      { id: "Heading1", name: "Heading 1", runFormatting: { fontSizePt: 30 }, headingLevel: 1 },
    ];
    const { model: next } = applyParagraphStyle(model, caretSelection(caret(0, 0)), "Heading1");
    expect(blocks(next)[0].runs[0].formatting).toEqual({ fontSizePt: 30 });
  });
});

describe("resetParagraphStyle (AC-14)", () => {
  it("removes direct formatting, leaving the style's own", () => {
    const model = fixture();
    const styled = applyParagraphStyle(model, caretSelection(caret(1, 0)), "Quote").model;
    const withDirect = structuredClone(styled);
    const second = asParagraph(withDirect.sections[0].blocks[1]);
    second.formatting = { ...second.formatting, alignment: "center" };
    second.runs[0].formatting.fontSizePt = 24;

    const { model: next } = resetParagraphStyle(withDirect, caretSelection(caret(1, 0)));
    const reset = blocks(next)[1];
    expect(reset.styleId).toBe("Quote");
    expect(reset.formatting).toEqual({ indentStartCm: 1.27, spaceBeforePt: 10, spaceAfterPt: 10 });
    expect(reset.runs.map((run) => run.formatting)).toEqual([
      { italic: true, color: "#666666" },
      { italic: true, color: "#666666" },
    ]);
  });

  it("clears everything on an unstyled paragraph", () => {
    const model = fixture();
    const withDirect = structuredClone(model);
    asParagraph(withDirect.sections[0].blocks[0]).formatting = { alignment: "right" };
    const { model: next } = resetParagraphStyle(withDirect, caretSelection(caret(0, 0)));
    const first = blocks(next)[0];
    expect(first.formatting).toBeUndefined();
    expect(first.styleId).toBeUndefined();
  });

  it("keeps a heading level the style chain does not set", () => {
    const model = fixture();
    const withHeading = structuredClone(model);
    asParagraph(withHeading.sections[0].blocks[0]).headingLevel = 3;
    const { model: next } = resetParagraphStyle(withHeading, caretSelection(caret(0, 0)));
    expect(blocks(next)[0].headingLevel).toBe(3);
  });
});

describe("updateStyleFromSelection (AC-12, AC-13)", () => {
  it("redefines the style from the current paragraph and re-bakes every user", () => {
    const model = fixture();
    model.styles = [
      { id: "Normal", name: "Normal" },
      { id: "Fancy", name: "Fancy", basedOn: "Normal", runFormatting: { italic: true } },
    ];
    const first = asParagraph(model.sections[0].blocks[0]);
    first.styleId = "Fancy";
    first.runs[0].formatting = { italic: true };
    const third = asParagraph(model.sections[0].blocks[2]);
    third.styleId = "Fancy";
    third.runs[0].formatting = { italic: true, color: "#ff0000" };

    // Make the current (first) paragraph's look the new definition.
    const restyled = structuredClone(model);
    asParagraph(restyled.sections[0].blocks[0]).runs[0].formatting = {
      italic: true,
      bold: true,
      fontSizePt: 16,
    };

    const { model: next } = updateStyleFromSelection(restyled, caretSelection(caret(0, 0)), "Fancy");
    const fancy = next.styles?.find((style) => style.id === "Fancy");
    expect(fancy?.runFormatting).toEqual({ italic: true, bold: true, fontSizePt: 16 });
    // Every paragraph using the style is re-baked, direct formatting kept.
    expect(blocks(next)[0].runs[0].formatting).toEqual({
      italic: true,
      bold: true,
      fontSizePt: 16,
    });
    expect(blocks(next)[2].runs[0].formatting).toEqual({
      italic: true,
      bold: true,
      fontSizePt: 16,
      color: "#ff0000",
    });
    // A paragraph not using the style keeps its object identity.
    expect(next.sections[0].blocks[1]).toBe(restyled.sections[0].blocks[1]);
  });

  it("re-bakes paragraphs styled with a descendant of the updated style", () => {
    const model = fixture();
    model.styles = [
      { id: "Base", name: "Base", runFormatting: { fontSizePt: 14 } },
      { id: "Child", name: "Child", basedOn: "Base", runFormatting: { italic: true } },
    ];
    const first = asParagraph(model.sections[0].blocks[0]);
    first.styleId = "Base";
    first.runs[0].formatting = { fontSizePt: 14, bold: true };
    const second = asParagraph(model.sections[0].blocks[1]);
    second.styleId = "Child";
    second.runs = [{ text: "Second", formatting: { fontSizePt: 14, italic: true } }];

    // Update Base from the first paragraph (bold 14pt becomes the style).
    const { model: next } = updateStyleFromSelection(model, caretSelection(caret(0, 0)), "Base");
    expect(next.styles?.find((style) => style.id === "Base")?.runFormatting).toEqual({
      fontSizePt: 14,
      bold: true,
    });
    // The Child-styled paragraph inherits the new Base.
    expect(blocks(next)[1].runs[0].formatting).toEqual({
      fontSizePt: 14,
      bold: true,
      italic: true,
    });
  });

  it("seeds a blank document's catalogue from the built-ins", () => {
    const model = fixture();
    const styled = applyParagraphStyle(model, caretSelection(caret(0, 0)), "Heading1").model;
    const recolored = structuredClone(styled);
    const first = asParagraph(recolored.sections[0].blocks[0]);
    first.runs[0].formatting = { ...first.runs[0].formatting, color: "#ff0000" };
    const { model: next } = updateStyleFromSelection(recolored, caretSelection(caret(0, 0)), "Heading1");
    expect(next.styles?.find((style) => style.id === "Heading1")?.runFormatting).toEqual({
      bold: true,
      fontSizePt: 20,
      color: "#ff0000",
    });
    expect(next.styles?.some((style) => style.id === "Quote")).toBe(true);
  });
});

describe("toolbarStyleState (AC-1)", () => {
  it("reports the first touched paragraph's named style", () => {
    const model = fixture();
    expect(toolbarStyleState(model, caretSelection(caret(0, 0)))).toEqual({
      id: "",
      name: "Normal",
    });
    const styled = applyParagraphStyle(model, caretSelection(caret(1, 0)), "Heading2").model;
    expect(toolbarStyleState(styled, caretSelection(caret(1, 2)))).toEqual({
      id: "Heading2",
      name: "Heading 2",
    });
  });
});
