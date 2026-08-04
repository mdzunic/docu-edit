import { describe, it, expect } from "vitest";
import { createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import {
  blockLength,
  blockText,
  clampCaret,
  compareCaret,
  deleteBackward,
  deleteForward,
  deleteRange,
  insertText,
  noCaret,
  splitParagraph,
} from "./textEditing";

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

/** A single-section document built from paragraphs. */
function doc(...blocks: Paragraph[]): DocumentModel {
  const model = createDocument("text-editing-test");
  model.sections = [{ blocks }];
  return model;
}

describe("blockText / blockLength", () => {
  it("concatenates run text", () => {
    const block = paragraph(["Hello "], ["world"]);
    expect(blockText(block)).toBe("Hello world");
    expect(blockLength(block)).toBe(11);
  });

  it("is empty for a paragraph with no runs", () => {
    expect(blockText(paragraph())).toBe("");
    expect(blockLength(paragraph())).toBe(0);
  });
});

describe("noCaret", () => {
  it("returns null", () => {
    expect(noCaret()).toBeNull();
  });
});

describe("compareCaret", () => {
  it("orders by section, then block, then offset", () => {
    expect(compareCaret({ section: 0, block: 0, offset: 5 }, { section: 1, block: 0, offset: 0 })).toBeLessThan(0);
    expect(compareCaret({ section: 0, block: 2, offset: 0 }, { section: 0, block: 1, offset: 9 })).toBeGreaterThan(0);
    expect(compareCaret({ section: 0, block: 1, offset: 3 }, { section: 0, block: 1, offset: 3 })).toBe(0);
  });
});

describe("clampCaret", () => {
  it("clamps out-of-range addresses into the model", () => {
    const model = doc(paragraph(["abc"]));
    expect(clampCaret(model, { section: 5, block: 9, offset: 99 })).toEqual({
      section: 0,
      block: 0,
      offset: 3,
    });
    expect(clampCaret(model, { section: -1, block: -1, offset: -1 })).toEqual({
      section: 0,
      block: 0,
      offset: 0,
    });
  });
});

describe("insertText", () => {
  it("creates a run when typing into an empty paragraph", () => {
    const model = doc(paragraph());
    const caret = insertText(model, { section: 0, block: 0, offset: 0 }, "Hi");
    expect(blockText(model.sections[0].blocks[0])).toBe("Hi");
    expect(caret).toEqual({ section: 0, block: 0, offset: 2 });
  });

  it("inserts inside a run and keeps its formatting", () => {
    const model = doc(paragraph(["bold text", { bold: true }]));
    insertText(model, { section: 0, block: 0, offset: 4 }, "er");
    const runs = model.sections[0].blocks[0].runs;
    expect(runs).toHaveLength(1);
    expect(runs[0].text).toBe("bolder text");
    expect(runs[0].formatting).toEqual({ bold: true });
  });

  it("inherits the left run's formatting at a run boundary", () => {
    const model = doc(paragraph(["plain "], ["bold", { bold: true }]));
    insertText(model, { section: 0, block: 0, offset: 6 }, "X");
    const runs = model.sections[0].blocks[0].runs;
    expect(runs.map((r) => r.text)).toEqual(["plain X", "bold"]);
    expect(runs[0].formatting).toEqual({});
    expect(runs[1].formatting).toEqual({ bold: true });
  });

  it("inherits the right run's formatting at the paragraph start", () => {
    const model = doc(paragraph(["bold", { bold: true }]));
    insertText(model, { section: 0, block: 0, offset: 0 }, "X");
    const runs = model.sections[0].blocks[0].runs;
    expect(runs).toHaveLength(1);
    expect(runs[0].text).toBe("Xbold");
    expect(runs[0].formatting).toEqual({ bold: true });
  });

  it("replaces the block object rather than mutating it (v-memo contract)", () => {
    const model = doc(paragraph(["abc"]));
    const before = model.sections[0].blocks[0];
    insertText(model, { section: 0, block: 0, offset: 1 }, "x");
    expect(model.sections[0].blocks[0]).not.toBe(before);
    expect(before.runs[0].text).toBe("abc");
  });
});

describe("splitParagraph", () => {
  it("splits runs across the two halves and keeps formatting on both", () => {
    const model = doc(paragraph(["Hello world", { italic: true }]));
    model.sections[0].blocks[0].formatting = { alignment: "center" };
    const caret = splitParagraph(model, { section: 0, block: 0, offset: 5 });
    const blocks = model.sections[0].blocks;
    expect(blocks).toHaveLength(2);
    expect(blockText(blocks[0])).toBe("Hello");
    expect(blockText(blocks[1])).toBe(" world");
    expect(blocks[0].formatting).toEqual({ alignment: "center" });
    expect(blocks[1].formatting).toEqual({ alignment: "center" });
    expect(blocks[1].runs[0].formatting).toEqual({ italic: true });
    expect(caret).toEqual({ section: 0, block: 1, offset: 0 });
  });

  it("splitting at the end produces an empty second paragraph", () => {
    const model = doc(paragraph(["abc"]));
    splitParagraph(model, { section: 0, block: 0, offset: 3 });
    const blocks = model.sections[0].blocks;
    expect(blockText(blocks[0])).toBe("abc");
    expect(blocks[1].runs).toHaveLength(0);
  });

  it("splitting at the start produces an empty first paragraph", () => {
    const model = doc(paragraph(["abc"]));
    const caret = splitParagraph(model, { section: 0, block: 0, offset: 0 });
    const blocks = model.sections[0].blocks;
    expect(blocks[0].runs).toHaveLength(0);
    expect(blockText(blocks[1])).toBe("abc");
    expect(caret).toEqual({ section: 0, block: 1, offset: 0 });
  });
});

describe("deleteRange", () => {
  it("deletes within a single run", () => {
    const model = doc(paragraph(["Hello world"]));
    const caret = deleteRange(
      model,
      { section: 0, block: 0, offset: 5 },
      { section: 0, block: 0, offset: 11 },
    );
    expect(blockText(model.sections[0].blocks[0])).toBe("Hello");
    expect(caret).toEqual({ section: 0, block: 0, offset: 5 });
  });

  it("accepts its carets in either order", () => {
    const model = doc(paragraph(["Hello world"]));
    deleteRange(
      model,
      { section: 0, block: 0, offset: 11 },
      { section: 0, block: 0, offset: 5 },
    );
    expect(blockText(model.sections[0].blocks[0])).toBe("Hello");
  });

  it("preserves formatting of runs outside the range", () => {
    const model = doc(paragraph(["aa"], ["bb", { bold: true }], ["cc", { italic: true }]));
    deleteRange(
      model,
      { section: 0, block: 0, offset: 1 },
      { section: 0, block: 0, offset: 5 },
    );
    const runs = model.sections[0].blocks[0].runs;
    expect(runs.map((r) => r.text)).toEqual(["a", "c"]);
    expect(runs[0].formatting).toEqual({});
    expect(runs[1].formatting).toEqual({ italic: true });
  });

  it("merges paragraphs when the range spans blocks", () => {
    const model = doc(paragraph(["Hello"]), paragraph(["skip me"]), paragraph(["world"]));
    const caret = deleteRange(
      model,
      { section: 0, block: 0, offset: 5 },
      { section: 0, block: 2, offset: 0 },
    );
    const blocks = model.sections[0].blocks;
    expect(blocks).toHaveLength(1);
    expect(blockText(blocks[0])).toBe("Helloworld");
    expect(caret).toEqual({ section: 0, block: 0, offset: 5 });
  });

  it("merges sections when the range spans them", () => {
    const model = createDocument("cross-section");
    model.sections = [
      { blocks: [paragraph(["one"]), paragraph(["two"])] },
      { blocks: [paragraph(["three"]), paragraph(["four"])] },
    ];
    deleteRange(
      model,
      { section: 0, block: 1, offset: 1 },
      { section: 1, block: 0, offset: 3 },
    );
    expect(model.sections).toHaveLength(1);
    const blocks = model.sections[0].blocks;
    expect(blocks.map(blockText)).toEqual(["one", "tee", "four"]);
  });

  it("is a no-op for a collapsed range", () => {
    const model = doc(paragraph(["abc"]));
    const before = model.sections[0].blocks[0];
    deleteRange(
      model,
      { section: 0, block: 0, offset: 1 },
      { section: 0, block: 0, offset: 1 },
    );
    expect(model.sections[0].blocks[0]).toBe(before);
  });
});

describe("deleteBackward", () => {
  it("deletes the character before the caret", () => {
    const model = doc(paragraph(["abc"]));
    const caret = deleteBackward(model, { section: 0, block: 0, offset: 2 });
    expect(blockText(model.sections[0].blocks[0])).toBe("ac");
    expect(caret).toEqual({ section: 0, block: 0, offset: 1 });
  });

  it("merges with the previous paragraph at a paragraph start", () => {
    const model = doc(paragraph(["Hello"]), paragraph(["world"]));
    const caret = deleteBackward(model, { section: 0, block: 1, offset: 0 });
    expect(model.sections[0].blocks).toHaveLength(1);
    expect(blockText(model.sections[0].blocks[0])).toBe("Helloworld");
    expect(caret).toEqual({ section: 0, block: 0, offset: 5 });
  });

  it("is a no-op at the start of the document", () => {
    const model = doc(paragraph(["abc"]));
    const caret = deleteBackward(model, { section: 0, block: 0, offset: 0 });
    expect(blockText(model.sections[0].blocks[0])).toBe("abc");
    expect(caret).toEqual({ section: 0, block: 0, offset: 0 });
  });
});

describe("deleteForward", () => {
  it("deletes the character after the caret", () => {
    const model = doc(paragraph(["abc"]));
    const caret = deleteForward(model, { section: 0, block: 0, offset: 1 });
    expect(blockText(model.sections[0].blocks[0])).toBe("ac");
    expect(caret).toEqual({ section: 0, block: 0, offset: 1 });
  });

  it("merges with the next paragraph at a paragraph end", () => {
    const model = doc(paragraph(["Hello"]), paragraph(["world"]));
    deleteForward(model, { section: 0, block: 0, offset: 5 });
    expect(model.sections[0].blocks).toHaveLength(1);
    expect(blockText(model.sections[0].blocks[0])).toBe("Helloworld");
  });

  it("is a no-op at the end of the document", () => {
    const model = doc(paragraph(["abc"]));
    const caret = deleteForward(model, { section: 0, block: 0, offset: 3 });
    expect(blockText(model.sections[0].blocks[0])).toBe("abc");
    expect(caret).toEqual({ section: 0, block: 0, offset: 3 });
  });
});
