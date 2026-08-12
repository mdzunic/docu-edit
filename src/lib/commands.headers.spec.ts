// SPEC-029 — headers, footers and page numbers: the per-section options
// command, the header/footer content addressing helpers (reusing the body
// editing primitives through a throwaway single-section model), and the
// "Insert page number" dialog's command.
import { describe, it, expect } from "vitest";
import {
  createDocument,
  type DocumentModel,
  type Paragraph,
} from "./documentModel";
import type { CaretPosition } from "./textEditing";
import {
  caretSelection,
  commitHeaderFooterModel,
  insertPageNumber,
  insertPageNumberCommand,
  insertText,
  setHeaderFooterOptions,
  withHeaderFooterModel,
  type ModelSelection,
} from "./commands";

function paragraph(text = ""): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function doc(): DocumentModel {
  const model = createDocument("headers-test");
  model.sections = [{ blocks: [paragraph("Hello")] }, { blocks: [paragraph("Second section")] }];
  return model;
}

function caret(): CaretPosition {
  return { section: 0, block: 0, offset: 0 };
}

function selection(): ModelSelection {
  return caretSelection(caret());
}

describe("setHeaderFooterOptions (SPEC-029 AC-11, AC-12, AC-22, AC-23)", () => {
  it("sets titlePage, evenAndOddHeaders and pageNumberStart on the target section only", () => {
    const result = setHeaderFooterOptions(doc(), selection(), 0, {
      titlePage: true,
      evenAndOddHeaders: true,
      pageNumberStart: 5,
    });
    expect(result.model.sections[0].titlePage).toBe(true);
    expect(result.model.sections[0].evenAndOddHeaders).toBe(true);
    expect(result.model.sections[0].pageNumberStart).toBe(5);
    expect(result.model.sections[1].titlePage).toBeUndefined();
  });

  it("clears pageNumberStart with null (continue from previous section)", () => {
    let model = setHeaderFooterOptions(doc(), selection(), 0, { pageNumberStart: 5 }).model;
    model = setHeaderFooterOptions(model, selection(), 0, { pageNumberStart: null }).model;
    expect(model.sections[0].pageNumberStart).toBeUndefined();
  });

  it("clears titlePage/evenAndOddHeaders with false", () => {
    let model = setHeaderFooterOptions(doc(), selection(), 0, {
      titlePage: true,
      evenAndOddHeaders: true,
    }).model;
    model = setHeaderFooterOptions(model, selection(), 0, {
      titlePage: false,
      evenAndOddHeaders: false,
    }).model;
    expect(model.sections[0].titlePage).toBeUndefined();
    expect(model.sections[0].evenAndOddHeaders).toBeUndefined();
  });

  it("does not touch a field left undefined in the patch", () => {
    let model = setHeaderFooterOptions(doc(), selection(), 0, { titlePage: true }).model;
    model = setHeaderFooterOptions(model, selection(), 0, { evenAndOddHeaders: true }).model;
    expect(model.sections[0].titlePage).toBe(true);
    expect(model.sections[0].evenAndOddHeaders).toBe(true);
  });

  it("leaves the input model untouched (undo safety)", () => {
    const model = doc();
    const before = structuredClone(model);
    setHeaderFooterOptions(model, selection(), 0, { titlePage: true });
    expect(model).toEqual(before);
  });
});

describe("withHeaderFooterModel / commitHeaderFooterModel (SPEC-029 AC-3, AC-10)", () => {
  it("addresses missing content as one empty paragraph", () => {
    const synthetic = withHeaderFooterModel(doc(), 0, "header", "default");
    expect(synthetic.sections).toHaveLength(1);
    expect(synthetic.sections[0].blocks).toEqual([{ type: "paragraph", runs: [] }]);
  });

  it("addresses existing content by kind", () => {
    const withHeader: DocumentModel = {
      ...doc(),
      sections: [
        { blocks: [paragraph("Hello")], headers: { default: { blocks: [paragraph("Existing")] } } },
        { blocks: [paragraph("Second section")] },
      ],
    };
    const synthetic = withHeaderFooterModel(withHeader, 0, "header", "default");
    expect(synthetic.sections[0].blocks).toEqual([paragraph("Existing")]);
  });

  it("round-trips a body-editing primitive (insertText) through the synthetic model", () => {
    const model = doc();
    const synthetic = withHeaderFooterModel(model, 0, "footer", "default");
    const edited = insertText(synthetic, caretSelection({ section: 0, block: 0, offset: 0 }), "Confidential");
    const next = commitHeaderFooterModel(model, 0, "footer", "default", edited.model);
    expect(next.sections[0].footers?.default?.blocks).toEqual([paragraph("Confidential")]);
    // Every other section, and the section's own body blocks, are untouched.
    expect(next.sections[0].blocks).toBe(model.sections[0].blocks);
    expect(next.sections[1]).toBe(model.sections[1]);
  });

  it("keeps model-level side effects (e.g. media) made by the inner command", () => {
    const model = doc();
    const synthetic = withHeaderFooterModel(model, 0, "header", "default");
    const withMedia: DocumentModel = { ...synthetic, media: [{ mediaId: "m1", contentType: "image/png", data: "" }] };
    const next = commitHeaderFooterModel(model, 0, "header", "default", withMedia);
    expect(next.media).toEqual([{ mediaId: "m1", contentType: "image/png", data: "" }]);
  });

  it("does not mutate the real section's other header/footer kinds", () => {
    const withFirst: DocumentModel = {
      ...doc(),
      sections: [
        { blocks: [paragraph("Hello")], headers: { first: { blocks: [paragraph("First page")] } } },
        { blocks: [paragraph("Second section")] },
      ],
    };
    const synthetic = withHeaderFooterModel(withFirst, 0, "header", "default");
    const edited = insertText(synthetic, caretSelection({ section: 0, block: 0, offset: 0 }), "X");
    const next = commitHeaderFooterModel(withFirst, 0, "header", "default", edited.model);
    expect(next.sections[0].headers?.first?.blocks).toEqual([paragraph("First page")]);
    expect(next.sections[0].headers?.default?.blocks).toEqual([paragraph("X")]);
  });
});

describe("insertPageNumber / insertPageNumberCommand (SPEC-029 AC-14..20)", () => {
  it("creates the footer with a single PAGE field when none exists", () => {
    const model = insertPageNumber(doc(), 0, {
      part: "footer",
      position: "center",
      format: "decimal",
      pageXofY: false,
    });
    const blocks = model.sections[0].footers?.default?.blocks;
    expect(blocks).toHaveLength(1);
    const p = blocks?.[0] as Paragraph;
    expect(p.formatting?.alignment).toBe("center");
    expect(p.runs).toHaveLength(1);
    expect(p.runs[0].field).toEqual({ kind: "page", format: "decimal" });
  });

  it("builds 'Page X of Y' as a PAGE field, literal ' of ', NUMPAGES field", () => {
    const model = insertPageNumber(doc(), 0, {
      part: "header",
      position: "right",
      format: "lowerRoman",
      pageXofY: true,
    });
    const p = model.sections[0].headers?.default?.blocks[0] as Paragraph;
    expect(p.runs.map((r) => r.field?.kind ?? r.text)).toEqual(["page", " of ", "numpages"]);
    expect((p.runs[0].field as { format?: string } | undefined)?.format).toBe("lowerRoman");
    expect((p.runs[2].field as { format?: string } | undefined)?.format).toBe("lowerRoman");
  });

  it("appends after existing content rather than replacing it", () => {
    const withHeader: DocumentModel = {
      ...doc(),
      sections: [
        { blocks: [paragraph("Hello")], footers: { default: { blocks: [paragraph("Confidential")] } } },
        { blocks: [paragraph("Second section")] },
      ],
    };
    const model = insertPageNumber(withHeader, 0, {
      part: "footer",
      position: "left",
      format: "decimal",
      pageXofY: false,
    });
    expect(model.sections[0].footers?.default?.blocks).toHaveLength(2);
    expect(model.sections[0].footers?.default?.blocks[0]).toEqual(paragraph("Confidential"));
  });

  it("writes the field into every applicable kind (default, first, even) so it shows on every page (AC-20)", () => {
    let model = doc();
    model = setHeaderFooterOptions(model, selection(), 0, {
      titlePage: true,
      evenAndOddHeaders: true,
    }).model;
    model = insertPageNumber(model, 0, {
      part: "footer",
      position: "center",
      format: "decimal",
      pageXofY: false,
    });
    expect(model.sections[0].footers?.default?.blocks).toHaveLength(1);
    expect(model.sections[0].footers?.first?.blocks).toHaveLength(1);
    expect(model.sections[0].footers?.even?.blocks).toHaveLength(1);
  });

  it("does not touch other sections", () => {
    const model = insertPageNumber(doc(), 0, {
      part: "footer",
      position: "center",
      format: "decimal",
      pageXofY: false,
    });
    expect(model.sections[1].footers).toBeUndefined();
  });

  it("insertPageNumberCommand wraps it as a CommandResult with the passed-through selection", () => {
    const model = doc();
    const before = selection();
    const result = insertPageNumberCommand(model, before, 0, {
      part: "footer",
      position: "center",
      format: "decimal",
      pageXofY: false,
    });
    expect(result.selection).toEqual(before);
    expect(result.model.sections[0].footers?.default?.blocks).toHaveLength(1);
  });

  it("leaves the input model untouched (undo safety)", () => {
    const model = doc();
    const before = structuredClone(model);
    insertPageNumber(model, 0, { part: "footer", position: "center", format: "decimal", pageXofY: false });
    expect(model).toEqual(before);
  });
});
