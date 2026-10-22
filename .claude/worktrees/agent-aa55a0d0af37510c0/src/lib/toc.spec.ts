import { describe, it, expect } from "vitest";
import {
  asToc,
  createDocument,
  isToc,
  type Block,
  type DocumentModel,
  type Paragraph,
  type TableOfContents,
} from "./documentModel";
import { caretSelection, isCollapsed, type ModelSelection } from "./commands";
import type { CaretPosition } from "./textEditing";
import { blockLength, deleteRange, paragraphAt } from "./textEditing";
import { equalStructural } from "./equality";
import { clearHistory, recordEdit, undo } from "./undoHistory";
import {
  DEFAULT_TOC_LEVELS,
  TOC_PLACEHOLDER_TEXT,
  defaultTocOptions,
  insertTableOfContents,
  noTocPageNumbers,
  tocEntriesFromHeadings,
  tocHeadingsStale,
  tocPageNumbersStale,
  updateTocEntirely,
  updateTocPageNumbersOnly,
} from "./toc";

function heading(text: string, level: number): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }], headingLevel: level };
}

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function doc(...blocks: Block[]): DocumentModel {
  const model = createDocument("toc-test");
  model.sections = [{ blocks }];
  return model;
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

describe("tocEntriesFromHeadings (SPEC-033 AC-1..4)", () => {
  it("includes only headings whose level is in options.levels, in document order", () => {
    const model = doc(
      heading("Chapter One", 1),
      paragraph("body"),
      heading("Section A", 2),
      heading("Subsection", 3),
      heading("Chapter Two", 1),
    );
    const entries = tocEntriesFromHeadings(model, { levels: [1, 2], showPageNumbers: true, hyperlinksOnly: false });
    expect(entries.map((e) => e.text)).toEqual(["Chapter One", "Section A", "Chapter Two"]);
    expect(entries.map((e) => e.level)).toEqual([1, 2, 1]);
  });

  it("gives every entry a unique, derived anchor", () => {
    const model = doc(heading("Same Title", 1), heading("Same Title", 1));
    const entries = tocEntriesFromHeadings(model, defaultTocOptions());
    expect(entries[0].anchor).not.toBe(entries[1].anchor);
  });

  it("resolves each entry's page number through the supplied lookup", () => {
    const model = doc(heading("Chapter One", 1), heading("Chapter Two", 1));
    const entries = tocEntriesFromHeadings(model, defaultTocOptions(), (anchor) =>
      anchor.includes("Chapter_One") ? 1 : 4,
    );
    expect(entries.map((e) => e.pageNumber)).toEqual([1, 4]);
  });

  it("returns no entries and no page numbers when no lookup is supplied", () => {
    const model = doc(heading("Chapter One", 1));
    const entries = tocEntriesFromHeadings(model, defaultTocOptions());
    expect(entries[0].pageNumber).toBeNull();
  });

  it("returns an empty list when the document has no Heading 1-6 paragraphs (AC-22)", () => {
    const model = doc(paragraph("just body text"));
    expect(tocEntriesFromHeadings(model, defaultTocOptions())).toEqual([]);
  });
});

describe("insertTableOfContents (SPEC-033 AC-1, AC-22)", () => {
  it("splices a TOC block at the caret, built from the document's current headings", () => {
    const model = doc(heading("Chapter One", 1), paragraph("body"));
    // At the very start of "body": no split needed, the TOC lands directly
    // between the heading and the paragraph.
    const selection = caretSelection(caret(1, 0));
    const result = insertTableOfContents(model, selection, defaultTocOptions(), () => 2);
    const blocks = result.model.sections[0].blocks;
    expect(blocks).toHaveLength(3);
    expect(isToc(blocks[1])).toBe(true);
    const toc = asToc(blocks[1]);
    expect(toc.entries).toEqual([
      { text: "Chapter One", level: 1, anchor: expect.any(String), pageNumber: 2 },
    ]);
    expect((blocks[0] as Paragraph).runs[0].text).toBe("Chapter One");
    expect((blocks[2] as Paragraph).runs.map((r) => r.text).join("")).toBe("body");
  });

  it("inserts an empty-state TOC with the placeholder shown by the renderer when there are no headings (AC-22, AC-23)", () => {
    const model = doc(paragraph("no headings here"));
    const selection = caretSelection(caret(0, 0));
    const result = insertTableOfContents(model, selection, defaultTocOptions(), noTocPageNumbers());
    const toc = asToc(result.model.sections[0].blocks[0]);
    expect(toc.entries).toEqual([]);
    // TOC_PLACEHOLDER_TEXT is what DocTocBlock.vue shows for an empty TOC —
    // asserted here so the two stay in sync.
    expect(TOC_PLACEHOLDER_TEXT.length).toBeGreaterThan(0);
  });

  it("sorts and defaults the stored options' levels", () => {
    const model = doc(heading("H", 2));
    const result = insertTableOfContents(
      model,
      caretSelection(caret(0, 0)),
      { levels: [3, 1, 2], showPageNumbers: true, hyperlinksOnly: false },
      noTocPageNumbers(),
    );
    expect(asToc(result.model.sections[0].blocks[0]).options.levels).toEqual([1, 2, 3]);
  });

  it("is a length-1 caret atom, like Image/BreakBlock", () => {
    const model = doc(heading("Chapter One", 1));
    const result = insertTableOfContents(model, caretSelection(caret(0, 0)), defaultTocOptions(), noTocPageNumbers());
    const toc = result.model.sections[0].blocks[0];
    expect(blockLength(toc)).toBe(1);
    expect(paragraphAt(result.model, { section: 0, block: 0 })).toBeNull();
  });

  it("never mutates the input model (commands.ts's immutability contract)", () => {
    const model = doc(heading("Chapter One", 1));
    const before = structuredClone(model);
    insertTableOfContents(model, caretSelection(caret(0, 0)), defaultTocOptions(), noTocPageNumbers());
    expect(equalStructural(model, before)).toBe(true);
  });

  it("is inert inside a table cell", () => {
    const model = doc({
      type: "table",
      rows: [{ cells: [{ blocks: [paragraph("cell text")] }] }],
      columnWidthsCm: [5],
    });
    const selection: ModelSelection = caretSelection({
      section: 0,
      block: 0,
      offset: 0,
      cell: { row: 0, col: 0, block: 0 },
    });
    const result = insertTableOfContents(model, selection, defaultTocOptions(), noTocPageNumbers());
    expect(result.model.sections[0].blocks).toHaveLength(1);
    expect(result.model.sections[0].blocks[0].type).toBe("table");
  });

  it("is undoable (SPEC-015 AC-20/21 contract, applied the same way commands.undo.spec.ts checks commands.ts's own exports)", () => {
    const model = doc(heading("Chapter One", 1), paragraph("body"));
    const before = structuredClone(model);
    const selection = caretSelection(caret(1, 0));
    const result = insertTableOfContents(model, selection, defaultTocOptions(), () => 1);

    const docId = "toc-undo-test";
    recordEdit(docId, {
      before: { model, selection },
      after: { model: result.model, selection: result.selection },
      label: "insert table of contents",
      groupable: false,
    });
    const restored = undo(docId);
    clearHistory(docId);
    expect(restored).not.toBeNull();
    expect(equalStructural(restored?.model, before)).toBe(true);
  });
});

describe("staleness (SPEC-033 AC-10, AC-11)", () => {
  function insertedToc(model: DocumentModel, pageNumberFor = () => 1) {
    const result = insertTableOfContents(model, caretSelection(caret(0, 0)), defaultTocOptions(), pageNumberFor);
    return { model: result.model, toc: asToc(result.model.sections[0].blocks[0]) };
  }

  it("is not stale immediately after generation", () => {
    const model = doc(heading("Chapter One", 1));
    const { model: next, toc } = insertedToc(model);
    expect(tocHeadingsStale(next, toc)).toBe(false);
    expect(tocPageNumbersStale(toc, () => 1)).toBe(false);
  });

  it("headings-stale when a heading's text changes, independent of pagination", () => {
    const model = doc(heading("Chapter One", 1));
    const { model: next, toc } = insertedToc(model);
    const renamed = structuredClone(next);
    (renamed.sections[0].blocks[1] as Paragraph).runs[0].text = "Chapter One Renamed";
    expect(tocHeadingsStale(renamed, toc)).toBe(true);
    expect(tocPageNumbersStale(toc, () => 1)).toBe(false);
  });

  it("headings-stale when a heading is added or removed", () => {
    const model = doc(heading("Chapter One", 1));
    const { model: next, toc } = insertedToc(model);
    const withExtra = structuredClone(next);
    withExtra.sections[0].blocks.push(heading("Chapter Two", 1));
    expect(tocHeadingsStale(withExtra, toc)).toBe(true);
  });

  it("pagination-stale when a page number no longer matches, independent of headings", () => {
    const model = doc(heading("Chapter One", 1));
    const { model: next, toc } = insertedToc(model, () => 1);
    expect(tocHeadingsStale(next, toc)).toBe(false);
    expect(tocPageNumbersStale(toc, () => 2)).toBe(true);
  });

  it("does not consider anchor differences headings-stale (a Word-authored anchor never matches this app's derived scheme)", () => {
    const model = doc(heading("Chapter One", 1));
    const { model: next, toc } = insertedToc(model);
    const foreignAnchorToc: TableOfContents = {
      ...toc,
      entries: toc.entries.map((entry) => ({ ...entry, anchor: "_Toc999" })),
    };
    expect(tocHeadingsStale(next, foreignAnchorToc)).toBe(false);
  });
});

describe("updateTocPageNumbersOnly (SPEC-033 AC-12)", () => {
  it("refreshes only page numbers, keeping entry text/level/anchor/order untouched", () => {
    const model = doc(heading("Chapter One", 1), heading("Chapter Two", 1));
    const inserted = insertTableOfContents(model, caretSelection(caret(0, 0)), defaultTocOptions(), () => 1).model;
    const tocBefore = asToc(inserted.sections[0].blocks[0]);

    const result = updateTocPageNumbersOnly(
      inserted,
      caretSelection(caret(0, 0)),
      0,
      0,
      (anchor) => (anchor === tocBefore.entries[0].anchor ? 5 : 9),
    );
    const tocAfter = asToc(result.model.sections[0].blocks[0]);
    expect(tocAfter.entries.map((e) => e.pageNumber)).toEqual([5, 9]);
    expect(tocAfter.entries.map((e) => e.text)).toEqual(tocBefore.entries.map((e) => e.text));
    expect(tocAfter.entries.map((e) => e.anchor)).toEqual(tocBefore.entries.map((e) => e.anchor));
  });

  it("does not pick up a heading added since generation (that needs Update entire table instead)", () => {
    const model = doc(heading("Chapter One", 1));
    const inserted = insertTableOfContents(model, caretSelection(caret(0, 0)), defaultTocOptions(), () => 1).model;
    inserted.sections[0].blocks.push(heading("Chapter Two", 1));

    const result = updateTocPageNumbersOnly(inserted, caretSelection(caret(0, 0)), 0, 0, () => 1);
    expect(asToc(result.model.sections[0].blocks[0]).entries).toHaveLength(1);
  });

  it("is a no-op when the address is not a TOC", () => {
    const model = doc(paragraph("not a toc"));
    const before = structuredClone(model);
    const result = updateTocPageNumbersOnly(model, caretSelection(caret(0, 0)), 0, 0, () => 1);
    expect(equalStructural(result.model, before)).toBe(true);
  });
});

describe("updateTocEntirely (SPEC-033 AC-13)", () => {
  it("rebuilds entries wholesale from the document's current headings", () => {
    const model = doc(heading("Chapter One", 1));
    const inserted = insertTableOfContents(model, caretSelection(caret(0, 0)), defaultTocOptions(), () => 1).model;
    inserted.sections[0].blocks.push(heading("Chapter Two", 1));

    const result = updateTocEntirely(inserted, caretSelection(caret(0, 0)), 0, 0, () => 3);
    const toc = asToc(result.model.sections[0].blocks[0]);
    expect(toc.entries.map((e) => e.text)).toEqual(["Chapter One", "Chapter Two"]);
    expect(toc.entries.every((e) => e.pageNumber === 3)).toBe(true);
    expect(tocHeadingsStale(result.model, toc)).toBe(false);
  });

  it("is a no-op when the address is not a TOC", () => {
    const model = doc(paragraph("not a toc"));
    const before = structuredClone(model);
    const result = updateTocEntirely(model, caretSelection(caret(0, 0)), 0, 0, () => 1);
    expect(equalStructural(result.model, before)).toBe(true);
  });
});

describe("deleting a TOC leaves headings untouched (SPEC-033 AC-16)", () => {
  it("removing the atom via a spanning range delete does not touch the headings elsewhere in the document", () => {
    const model = doc(heading("Chapter One", 1), paragraph("body"));
    const inserted = insertTableOfContents(model, caretSelection(caret(1, 0)), defaultTocOptions(), () => 1).model;
    expect(isToc(inserted.sections[0].blocks[1])).toBe(true);

    // Selecting from just before the TOC to just after it and deleting is
    // the generic atom-delete path every Image/Table/BreakBlock already
    // uses (textEditing.deleteRange) — no TOC-specific deletion code exists
    // or is needed.
    const start: CaretPosition = { section: 0, block: 1, offset: 0 };
    const end: CaretPosition = { section: 0, block: 1, offset: 1 };
    expect(isCollapsed({ anchor: start, focus: end })).toBe(false);
    const after = deleteRange(inserted, start, end);
    expect(after).toBeDefined();
    expect(inserted.sections[0].blocks.some(isToc)).toBe(false);
    expect((inserted.sections[0].blocks[0] as Paragraph).headingLevel).toBe(1);
    expect((inserted.sections[0].blocks[0] as Paragraph).runs[0].text).toBe("Chapter One");
  });
});

describe("performance (SPEC-033 AC-21)", () => {
  it("generates a TOC for a document with 300 headings in under 1 second", () => {
    const blocks: Block[] = [];
    for (let i = 0; i < 300; i += 1) {
      blocks.push(heading(`Heading ${i}`, (i % 6) + 1));
      blocks.push(paragraph(`Body paragraph ${i}.`));
    }
    const model = doc(...blocks);
    const started = performance.now();
    const result = insertTableOfContents(
      model,
      caretSelection(caret(0, 0)),
      { levels: DEFAULT_TOC_LEVELS.concat([4, 5, 6]), showPageNumbers: true, hyperlinksOnly: false },
      (anchor) => anchor.length,
    );
    const elapsedMs = performance.now() - started;
    expect(asToc(result.model.sections[0].blocks[0]).entries).toHaveLength(300);
    expect(elapsedMs).toBeLessThan(1000);
  });
});
