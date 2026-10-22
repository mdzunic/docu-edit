// SPEC-036 — word counting rules, body/header/footer/notes scoping, the
// incremental cache and the 200-page performance budget.
import { describe, expect, it } from "vitest";
import {
  createDocument,
  createEmptyHeaderFooterContent,
  createEmptyNoteContent,
  createNoteRun,
  type Block,
  type DocumentModel,
  type NoteKind,
  type Paragraph,
  type Table,
} from "./documentModel";
import { caretSelection, type ModelSelection } from "./commands";
import type { CaretPosition } from "./textEditing";
import {
  addCountStats,
  countBlock,
  countText,
  documentBodyStats,
  documentNotesStats,
  documentStats,
  emptyCountStats,
  selectionText,
  selectionWordCount,
  type CountStats,
} from "./wordCount";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function doc(...blocks: Block[]): DocumentModel {
  const model = createDocument("word-count-test");
  model.sections = [{ blocks }];
  return model;
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

function selectionOf(a: CaretPosition, b: CaretPosition): ModelSelection {
  return { anchor: a, focus: b };
}

describe("countText — word counting rules (SPEC-036)", () => {
  it("counts a plain sentence's words (AC-14: numbers count as words)", () => {
    expect(countText("There are 42 cats here").words).toBe(5);
  });

  it("counts a hyphenated word as one word (AC-13, AC-17)", () => {
    expect(countText("a well-known state-of-the-art result").words).toBe(4);
  });

  it("does not merge a dash used between separate words", () => {
    expect(countText("word - word").words).toBe(2);
  });

  it("counts a bare https URL as one word regardless of internal punctuation (AC-15, AC-19)", () => {
    // "Visit" + the URL (one word despite its /?=& punctuation) + "for" + "info".
    expect(countText("Visit https://example.com/page?x=1&y=2 for info").words).toBe(4);
  });

  it("counts a www. URL as one word and strips trailing sentence punctuation", () => {
    const stats = countText("See www.example.com/page, please.");
    // "See" + URL + "please" = 3 words.
    expect(stats.words).toBe(3);
  });

  it("counts each CJK character as its own word and its own CJK character (AC-16, AC-20)", () => {
    const stats = countText("你好世界");
    expect(stats.words).toBe(4);
    expect(stats.cjkCharacters).toBe(4);
  });

  it("counts mixed CJK and Latin text separately", () => {
    const stats = countText("hello 你好 world");
    expect(stats.words).toBe(4); // hello, 你, 好, world
    expect(stats.cjkCharacters).toBe(2);
  });

  it("a non-CJK sample has zero CJK characters", () => {
    expect(countText("hello world").cjkCharacters).toBe(0);
  });

  it("counts characters with and without spaces", () => {
    const stats = countText("ab cd");
    expect(stats.charactersWithSpaces).toBe(5);
    expect(stats.charactersWithoutSpaces).toBe(4);
  });

  it("counts sentences by terminal punctuation, not decimal points", () => {
    expect(countText("Pi is 3.14. It is irrational.").sentences).toBe(2);
  });

  it("counts a run of punctuation (ellipsis, interrobang) as one sentence end", () => {
    expect(countText("Wait... Really?! Yes.").sentences).toBe(3);
  });

  it("counts trailing text with no terminal punctuation as one more sentence", () => {
    expect(countText("Hello world").sentences).toBe(1);
    expect(countText("One. Two").sentences).toBe(2);
  });

  it("an empty string has zero of everything", () => {
    const stats = countText("");
    expect(stats.words).toBe(0);
    expect(stats.sentences).toBe(0);
    expect(stats.charactersWithSpaces).toBe(0);
  });
});

describe("countBlock / addCountStats", () => {
  it("counts a paragraph as one paragraph plus its text's counts", () => {
    const stats = countBlock(paragraph("two words"));
    expect(stats.paragraphs).toBe(1);
    expect(stats.words).toBe(2);
  });

  it("counts an empty paragraph as a paragraph with zero words", () => {
    const stats = countBlock(paragraph(""));
    expect(stats.paragraphs).toBe(1);
    expect(stats.words).toBe(0);
  });

  it("excludes a tracked deletion's words (SPEC-039 AC-17)", () => {
    const p: Paragraph = {
      type: "paragraph",
      runs: [
        { text: "two words ", formatting: {} },
        {
          text: "three deleted words",
          formatting: {},
          changes: { delete: { changeId: "c1", author: "Ada", date: "2026-01-01T00:00:00.000Z" } },
        },
      ],
    };
    const stats = countBlock(p);
    expect(stats.words).toBe(2);
  });

  it("sums a table's cell paragraphs, one per cell paragraph", () => {
    const table: Table = {
      type: "table",
      columnWidthsCm: [5, 5],
      rows: [
        {
          cells: [
            { blocks: [paragraph("cell one")] },
            { blocks: [paragraph("cell two words")] },
          ],
        },
      ],
    };
    const stats = countBlock(table);
    expect(stats.paragraphs).toBe(2);
    expect(stats.words).toBe(5);
  });

  it("images, breaks and a table of contents contribute nothing", () => {
    const image: Block = {
      type: "image",
      mediaId: "m1",
      naturalWidthPx: 10,
      naturalHeightPx: 10,
      widthCm: 1,
      heightCm: 1,
      wrap: "inline",
    };
    const brk: Block = { type: "break", kind: "page" };
    const toc: Block = { type: "toc", options: { levels: [1], showPageNumbers: true, hyperlinksOnly: false }, entries: [] };
    expect(countBlock(image)).toEqual(emptyCountStats());
    expect(countBlock(brk)).toEqual(emptyCountStats());
    expect(countBlock(toc)).toEqual(emptyCountStats());
  });

  it("addCountStats sums every field", () => {
    const a: CountStats = { words: 1, charactersWithSpaces: 2, charactersWithoutSpaces: 2, sentences: 1, cjkCharacters: 0, paragraphs: 1 };
    const b: CountStats = { words: 3, charactersWithSpaces: 4, charactersWithoutSpaces: 3, sentences: 0, cjkCharacters: 2, paragraphs: 1 };
    expect(addCountStats(a, b)).toEqual({
      words: 4,
      charactersWithSpaces: 6,
      charactersWithoutSpaces: 5,
      sentences: 1,
      cjkCharacters: 2,
      paragraphs: 2,
    });
  });
});

describe("documentBodyStats — scoping (SPEC-036 AC-21..24)", () => {
  it("counts only body paragraphs, excluding header/footer content", () => {
    const model = doc(paragraph("body text here"));
    model.sections[0].headers = { default: createEmptyHeaderFooterContent() };
    model.sections[0].headers.default!.blocks = [paragraph("secret header words that must not count")];
    model.sections[0].footers = { default: createEmptyHeaderFooterContent() };
    model.sections[0].footers.default!.blocks = [paragraph("footer words too")];
    const stats = documentBodyStats(model);
    expect(stats.words).toBe(3);
    expect(stats.paragraphs).toBe(1);
  });

  it("sums across every section", () => {
    const model = doc(paragraph("one two"));
    model.sections.push({ blocks: [paragraph("three four five")] });
    expect(documentBodyStats(model).words).toBe(5);
  });

  it("a block's cached count is reused until the block object is replaced", () => {
    const p = paragraph("two words");
    const model = doc(p);
    const cache = new WeakMap();
    const first = documentBodyStats(model, cache);
    expect(cache.has(p)).toBe(true);
    // Mutating the cached entry directly proves the second call reused the
    // cache instead of recomputing from the (still "two words") block.
    cache.set(p, { ...first, words: 999 });
    const second = documentBodyStats(model, cache);
    expect(second.words).toBe(999);
    // Replacing the block object (the mutation contract every command
    // follows) is what invalidates the cache — a fresh object always misses.
    model.sections[0].blocks[0] = paragraph("three different words now");
    const third = documentBodyStats(model, cache);
    expect(third.words).toBe(4);
  });
});

function docWithNote(bodyText: string, noteText: string, kind: NoteKind = "footnote"): DocumentModel {
  const model = doc(paragraph(bodyText));
  (model.sections[0].blocks[0] as Paragraph).runs.push(createNoteRun("n1", kind));
  const note = createEmptyNoteContent("n1", kind);
  note.blocks = [paragraph(noteText)];
  model.notes = [note];
  return model;
}

describe("documentStats — notes toggle (SPEC-036 AC-11, AC-12)", () => {
  it("excludes notes from the body count by default", () => {
    const model = docWithNote("body words here", "footnote content words");
    const bodyOnly = documentStats(model, { includeNotes: false });
    expect(bodyOnly.words).toBe(3);
  });

  it("includes referenced footnotes and endnotes when the toggle is on", () => {
    const model = docWithNote("body words here", "footnote content words");
    const withNotes = documentStats(model, { includeNotes: true });
    expect(withNotes.words).toBe(6);
  });

  it("excludes a note whose reference mark was deleted, even with the toggle on", () => {
    const model = docWithNote("body words here", "footnote content words");
    // Deleting the reference mark run (SPEC-031's "dangling" case).
    (model.sections[0].blocks[0] as Paragraph).runs = [{ text: "body words here", formatting: {} }];
    const withNotes = documentStats(model, { includeNotes: true });
    expect(withNotes.words).toBe(3);
  });

  it("documentNotesStats alone counts every referenced note stream", () => {
    const footnoteModel = docWithNote("body", "foot note text", "footnote");
    expect(documentNotesStats(footnoteModel).words).toBe(3);
    const endnoteModel = docWithNote("body", "end note text", "endnote");
    expect(documentNotesStats(endnoteModel).words).toBe(3);
  });
});

describe("selectionText / selectionWordCount (SPEC-036 AC-3)", () => {
  it("is empty and zero for a collapsed selection", () => {
    const model = doc(paragraph("hello world"));
    const sel = caretSelection(caret(0, 3));
    expect(selectionText(model, sel)).toBe("");
    expect(selectionWordCount(model, sel)).toBe(0);
  });

  it("counts only the words inside the selection, not the whole document", () => {
    const model = doc(paragraph("one two three four"));
    // Select "two three" (offsets 4..13).
    const sel = selectionOf(caret(0, 4), caret(0, 13));
    expect(selectionText(model, sel)).toBe("two three");
    expect(selectionWordCount(model, sel)).toBe(2);
    expect(documentBodyStats(model).words).toBe(4);
  });

  it("spans multiple paragraphs, joined by newline", () => {
    const model = doc(paragraph("first paragraph"), paragraph("second paragraph"));
    const sel = selectionOf(caret(0, 0), caret(1, 6));
    expect(selectionText(model, sel)).toBe("first paragraph\nsecond");
    expect(selectionWordCount(model, sel)).toBe(3);
  });
});

describe("performance (SPEC-036 AC-25, AC-26)", () => {
  it("counts a 200-page document in under 200ms", () => {
    // Mirrors lib/pagination.spec.ts's own 200-page perf fixture: 10
    // paragraphs per page.
    const blocks = Array.from({ length: 10 * 200 }, (_, i) =>
      paragraph(`Paragraph number ${i} has a few words and a sentence. It continues here.`),
    );
    const model = doc(...blocks);
    const started = performance.now();
    const stats = documentBodyStats(model);
    const elapsed = performance.now() - started;
    expect(stats.paragraphs).toBe(2000);
    expect(elapsed).toBeLessThan(200);
  });

  it("recomputing after a single-paragraph edit only re-counts that paragraph", () => {
    const blocks = Array.from({ length: 10 * 200 }, (_, i) => paragraph(`Paragraph ${i} has words.`));
    const model = doc(...blocks);
    const cache = new WeakMap();
    documentBodyStats(model, cache); // warm the cache
    model.sections[0].blocks[5] = paragraph("A single replaced paragraph.");
    const started = performance.now();
    const stats = documentBodyStats(model, cache);
    const elapsed = performance.now() - started;
    expect(stats.paragraphs).toBe(2000);
    // Every block except the one replaced above hit the warm cache, so the
    // incremental recompute is far cheaper than the cold pass above.
    expect(elapsed).toBeLessThan(50);
  });
});
