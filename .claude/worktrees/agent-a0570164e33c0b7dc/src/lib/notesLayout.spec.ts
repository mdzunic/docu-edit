// SPEC-031 — footnote/endnote placement over the pagination engine, driven
// by fake metrics like pagination.spec: 1 character = 1 line = 10px, page
// content exactly 10 lines tall.
import { describe, expect, it } from "vitest";
import { paginate, type MetricsProvider } from "./pagination";
import { type BlockMetrics } from "./core/pageMetrics";
import { PX_PER_CM } from "./core/units";
import {
  layoutWithNotes,
  NOTE_GAP_PX,
  NOTE_SEPARATOR_HEIGHT_PX,
  type NoteMetricsProvider,
} from "./notesLayout";
import {
  createDocument,
  createNoteRun,
  type Block,
  type DocumentModel,
  type NoteContent,
  type NoteKind,
  type Paragraph,
} from "./documentModel";

const LINE = 10;
const PAGE_LINES = 10;

function paragraph(lineCount: number): Paragraph {
  return { type: "paragraph", runs: [{ text: "x".repeat(lineCount), formatting: {} }] };
}

/** A paragraph ending in a note reference mark: `lineCount - 1` filler
 * characters plus the one-character mark. */
function refParagraph(lineCount: number, noteId: string, kind: NoteKind = "footnote"): Paragraph {
  return {
    type: "paragraph",
    runs: [
      { text: "x".repeat(Math.max(0, lineCount - 1)), formatting: {} },
      createNoteRun(noteId, kind),
    ],
  };
}

function metricsOf(block: Block, contentWidthPx: number): BlockMetrics {
  if (block.type !== "paragraph") {
    return { contentWidthPx, heightPx: block.type === "break" ? 0 : LINE, marginTopPx: 0, marginBottomPx: 0, lines: [] };
  }
  const count = Math.max(1, block.runs.reduce((n, run) => n + run.text.length, 0));
  return {
    contentWidthPx,
    heightPx: count * LINE,
    marginTopPx: 0,
    marginBottomPx: 0,
    lines: Array.from({ length: count }, (_, i) => ({ top: i * LINE, height: LINE })),
  };
}

const fakeMetrics: MetricsProvider = {
  metricsFor(_s, _b, block, contentWidthPx) {
    return metricsOf(block, contentWidthPx);
  },
};

const fakeNoteMetrics: NoteMetricsProvider = {
  noteMetricsFor(_note, _index, block, contentWidthPx) {
    return metricsOf(block, contentWidthPx);
  },
};

/** Line offsets for the fake convention: one character per line. */
function fakeLineOffsets(block: Block): number[] | null {
  if (block.type !== "paragraph") {
    return null;
  }
  const count = Math.max(1, block.runs.reduce((n, run) => n + run.text.length, 0));
  return Array.from({ length: count }, (_, i) => i);
}

function fixture(blocks: Block[], notes: NoteContent[] = []): DocumentModel {
  const model = createDocument("notes-layout-fixture");
  model.defaultStyles.pageSize = {
    widthCm: 100 / PX_PER_CM,
    heightCm: (PAGE_LINES * LINE) / PX_PER_CM,
  };
  model.defaultStyles.margins = { topCm: 0, rightCm: 0, bottomCm: 0, leftCm: 0 };
  model.sections = [{ blocks }];
  if (notes.length > 0) {
    model.notes = notes;
  }
  return model;
}

function note(noteId: string, lineCount: number, kind: NoteKind = "footnote"): NoteContent {
  return { noteId, kind, blocks: [paragraph(lineCount)] };
}

function layout(model: DocumentModel) {
  return layoutWithNotes(model, fakeMetrics, fakeNoteMetrics, fakeLineOffsets);
}

describe("layoutWithNotes (SPEC-031)", () => {
  it("matches plain pagination when there are no notes", () => {
    const model = fixture(Array.from({ length: 15 }, () => paragraph(1)));
    const result = layout(model);
    expect(result.pages).toEqual(paginate(model, fakeMetrics));
    expect(result.footnoteAreas.every((area) => area === null)).toBe(true);
    expect(result.endnoteAreas.every((area) => area === null)).toBe(true);
  });

  it("places a footnote on the page of its reference (AC-2, AC-25)", () => {
    const blocks: Block[] = Array.from({ length: 15 }, () => paragraph(1));
    blocks[12] = refParagraph(1, "n1");
    const model = fixture(blocks, [note("n1", 1)]);
    const result = layout(model);
    expect(result.pages).toHaveLength(2);
    expect(result.footnoteAreas[0]).toBeNull();
    expect(result.footnoteAreas[1]).not.toBeNull();
    expect(result.footnoteAreas[1]?.items).toEqual([{ noteId: "n1", blockIndex: 0 }]);
    expect(result.footnotePageByNote.get("n1")).toBe(1);
  });

  it("reserves page space for the note area, moving displaced body content on (AC-5)", () => {
    const blocks: Block[] = Array.from({ length: 10 }, () => paragraph(1));
    blocks[0] = refParagraph(1, "n1");
    const model = fixture(blocks, [note("n1", 1)]);
    const result = layout(model);
    // The area (1 line + separator = 26px) shrinks page 1 to 74px = 7 lines.
    expect(result.pages).toHaveLength(2);
    expect(result.pages[0].items).toHaveLength(7);
    expect(result.pages[1].items).toHaveLength(3);
    expect(result.footnoteAreas[0]?.heightPx).toBe(LINE + NOTE_SEPARATOR_HEIGHT_PX);
  });

  it("moves a note with its reference when the reference re-paginates to another page (AC-12)", () => {
    const short = fixture([paragraph(5), refParagraph(1, "n1")], [note("n1", 1)]);
    const before = layout(short);
    expect(before.footnotePageByNote.get("n1")).toBe(0);

    const long = fixture([paragraph(15), refParagraph(1, "n1")], [note("n1", 1)]);
    const after = layout(long);
    expect(after.footnotePageByNote.get("n1")).toBe(1);
    expect(after.footnoteAreas[0]).toBeNull();
    expect(after.footnoteAreas[1]?.items[0]?.noteId).toBe("n1");
  });

  it("splits an over-long note across pages at line boundaries (AC-13)", () => {
    const model = fixture([refParagraph(1, "n1")], [note("n1", 8)]);
    const result = layout(model);
    // Area capacity is 70% of 100px minus the separator = 54px -> 5 lines;
    // the remaining 3 lines continue on an appended note-only page.
    expect(result.pages).toHaveLength(2);
    expect(result.pages[1].items).toEqual([]);
    expect(result.footnoteAreas[0]?.items).toEqual([
      { noteId: "n1", blockIndex: 0, startLine: 0, endLine: 5 },
    ]);
    expect(result.footnoteAreas[1]?.items).toEqual([
      { noteId: "n1", blockIndex: 0, startLine: 5, endLine: 8 },
    ]);
  });

  it("stacks several notes of one page in reference order with gaps", () => {
    const blocks: Block[] = [refParagraph(1, "a"), refParagraph(1, "b")];
    const model = fixture(blocks, [note("a", 1), note("b", 2)]);
    const result = layout(model);
    expect(result.footnoteAreas[0]?.items.map((item) => item.noteId)).toEqual(["a", "b"]);
    expect(result.footnoteAreas[0]?.heightPx).toBe(
      3 * LINE + NOTE_GAP_PX + NOTE_SEPARATOR_HEIGHT_PX,
    );
  });

  it("reports the page of every footnote reference for per-page restarts (AC-19)", () => {
    const blocks: Block[] = Array.from({ length: 15 }, () => paragraph(1));
    blocks[2] = refParagraph(1, "a");
    blocks[12] = refParagraph(1, "b");
    const model = fixture(blocks, [note("a", 1), note("b", 1)]);
    const result = layout(model);
    expect(result.footnoteRefPages).toEqual([0, 1]);
  });

  it("lays the endnote area after the last content of the document (AC-4)", () => {
    const blocks: Block[] = Array.from({ length: 12 }, () => paragraph(1));
    blocks[0] = refParagraph(1, "e1", "endnote");
    const model = fixture(blocks, [note("e1", 2, "endnote")]);
    const result = layout(model);
    expect(result.pages).toHaveLength(2);
    expect(result.endnoteAreas[0]).toBeNull();
    expect(result.endnoteAreas[1]?.items).toEqual([{ noteId: "e1", blockIndex: 0 }]);
  });

  it("continues an endnote area onto an appended page when the last page is full", () => {
    const blocks: Block[] = Array.from({ length: 20 }, () => paragraph(1));
    blocks[0] = refParagraph(1, "e1", "endnote");
    const model = fixture(blocks, [note("e1", 2, "endnote")]);
    const result = layout(model);
    expect(result.pages).toHaveLength(3);
    expect(result.pages[2].items).toEqual([]);
    expect(result.endnoteAreas[2]?.items[0]?.noteId).toBe("e1");
  });

  it("keeps footnotes and endnotes in separate areas", () => {
    const model = fixture(
      [refParagraph(1, "f1"), refParagraph(1, "e1", "endnote")],
      [note("f1", 1), note("e1", 1, "endnote")],
    );
    const result = layout(model);
    expect(result.footnoteAreas[0]?.kind).toBe("footnote");
    expect(result.footnoteAreas[0]?.items[0]?.noteId).toBe("f1");
    expect(result.endnoteAreas[0]?.kind).toBe("endnote");
    expect(result.endnoteAreas[0]?.items[0]?.noteId).toBe("e1");
  });

  it("ignores dangling references without stored note content", () => {
    const model = fixture([refParagraph(1, "ghost")]);
    const result = layout(model);
    expect(result.footnoteAreas.every((area) => area === null)).toBe(true);
  });
});
