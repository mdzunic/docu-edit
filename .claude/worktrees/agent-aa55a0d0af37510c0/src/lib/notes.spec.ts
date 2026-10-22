// SPEC-031 — footnote/endnote reference collection, numbering and content
// addressing.
import { describe, expect, it } from "vitest";
import {
  createDocument,
  createEmptyNoteContent,
  createNoteRun,
  type DocumentModel,
  type NoteKind,
  type Paragraph,
  type Table,
} from "./documentModel";
import {
  collectNoteRefs,
  effectiveNoteSettings,
  formatNoteNumber,
  noteById,
  noteLabels,
  noteText,
  referencedNotes,
  withNoteBlocks,
} from "./notes";

function paragraph(...texts: string[]): Paragraph {
  return { type: "paragraph", runs: texts.map((text) => ({ text, formatting: {} })) };
}

function paragraphWithRef(before: string, noteId: string, kind: NoteKind, after = ""): Paragraph {
  return {
    type: "paragraph",
    runs: [
      { text: before, formatting: {} },
      createNoteRun(noteId, kind),
      { text: after, formatting: {} },
    ],
  };
}

function docWithNotes(blocks: Paragraph[], noteIds: string[], kind: NoteKind = "footnote"): DocumentModel {
  const model = createDocument("notes-test");
  model.sections = [{ blocks }];
  model.notes = noteIds.map((noteId) => createEmptyNoteContent(noteId, kind));
  return model;
}

describe("collectNoteRefs (SPEC-031 AC-7)", () => {
  it("collects reference marks in document order with their offsets", () => {
    const model = docWithNotes(
      [paragraphWithRef("ab", "n1", "footnote", "cd"), paragraphWithRef("", "n2", "footnote")],
      ["n1", "n2"],
    );
    const refs = collectNoteRefs(model, "footnote");
    expect(refs).toEqual([
      { noteId: "n1", kind: "footnote", section: 0, block: 0, offset: 2 },
      { noteId: "n2", kind: "footnote", section: 0, block: 1, offset: 0 },
    ]);
  });

  it("separates footnote and endnote streams", () => {
    const model = docWithNotes(
      [paragraphWithRef("a", "f1", "footnote"), paragraphWithRef("b", "e1", "endnote")],
      ["f1"],
    );
    expect(collectNoteRefs(model, "footnote").map((ref) => ref.noteId)).toEqual(["f1"]);
    expect(collectNoteRefs(model, "endnote").map((ref) => ref.noteId)).toEqual(["e1"]);
  });

  it("finds references inside table cells with their cell path", () => {
    const table: Table = {
      type: "table",
      columnWidthsCm: [5, 5],
      rows: [
        {
          cells: [
            { blocks: [paragraph("plain")] },
            { blocks: [paragraphWithRef("x", "n1", "footnote")] },
          ],
        },
      ],
    };
    const model = createDocument("notes-table-test");
    model.sections = [{ blocks: [table] }];
    const refs = collectNoteRefs(model, "footnote");
    expect(refs).toEqual([
      {
        noteId: "n1",
        kind: "footnote",
        section: 0,
        block: 0,
        offset: 1,
        cell: { row: 0, col: 1, block: 0 },
      },
    ]);
  });
});

describe("formatNoteNumber (SPEC-031 AC-17)", () => {
  it("formats decimal, roman and alphabetic ordinals", () => {
    expect(formatNoteNumber("decimal", 3)).toBe("3");
    expect(formatNoteNumber("lowerRoman", 4)).toBe("iv");
    expect(formatNoteNumber("lowerAlpha", 2)).toBe("b");
  });

  it("cycles the footnote symbols, doubling each full cycle", () => {
    expect(formatNoteNumber("symbol", 1)).toBe("*");
    expect(formatNoteNumber("symbol", 2)).toBe("†");
    expect(formatNoteNumber("symbol", 4)).toBe("§");
    expect(formatNoteNumber("symbol", 5)).toBe("**");
    expect(formatNoteNumber("symbol", 8)).toBe("§§");
    expect(formatNoteNumber("symbol", 9)).toBe("***");
  });
});

describe("effectiveNoteSettings (SPEC-031 AC-17..20)", () => {
  it("defaults footnotes to decimal and endnotes to lowerRoman, both continuous from 1", () => {
    const model = createDocument("settings-test");
    expect(effectiveNoteSettings(model, "footnote")).toEqual({
      numberFormat: "decimal",
      startAt: 1,
      restart: "continuous",
    });
    expect(effectiveNoteSettings(model, "endnote")).toEqual({
      numberFormat: "lowerRoman",
      startAt: 1,
      restart: "continuous",
    });
  });

  it("reads stored settings and clamps the start to at least 1", () => {
    const model = createDocument("settings-test");
    model.footnoteSettings = { numberFormat: "symbol", startAt: 0, restart: "page" };
    expect(effectiveNoteSettings(model, "footnote")).toEqual({
      numberFormat: "symbol",
      startAt: 1,
      restart: "page",
    });
  });
});

describe("noteLabels (SPEC-031 AC-7, AC-17..20)", () => {
  it("numbers notes in document order from the start value", () => {
    const model = docWithNotes(
      [paragraphWithRef("a", "n1", "footnote"), paragraphWithRef("b", "n2", "footnote")],
      ["n1", "n2"],
    );
    model.footnoteSettings = { startAt: 5 };
    const labels = noteLabels(model, "footnote", collectNoteRefs(model, "footnote"));
    expect(labels.get("n1")).toBe("5");
    expect(labels.get("n2")).toBe("6");
  });

  it("restarts numbering per section (AC-20)", () => {
    const model = createDocument("restart-section");
    model.sections = [
      { blocks: [paragraphWithRef("a", "n1", "footnote")] },
      { blocks: [paragraphWithRef("b", "n2", "footnote"), paragraphWithRef("c", "n3", "footnote")] },
    ];
    model.notes = ["n1", "n2", "n3"].map((noteId) => createEmptyNoteContent(noteId, "footnote"));
    model.footnoteSettings = { restart: "section" };
    const labels = noteLabels(model, "footnote", collectNoteRefs(model, "footnote"));
    expect(labels.get("n1")).toBe("1");
    expect(labels.get("n2")).toBe("1");
    expect(labels.get("n3")).toBe("2");
  });

  it("restarts numbering per page when a page lookup is supplied (AC-19)", () => {
    const model = docWithNotes(
      [
        paragraphWithRef("a", "n1", "footnote"),
        paragraphWithRef("b", "n2", "footnote"),
        paragraphWithRef("c", "n3", "footnote"),
      ],
      ["n1", "n2", "n3"],
    );
    model.footnoteSettings = { restart: "page" };
    const refs = collectNoteRefs(model, "footnote");
    const pageByBlock = [0, 1, 1];
    const labels = noteLabels(model, "footnote", refs, (ref) => pageByBlock[ref.block]);
    expect(labels.get("n1")).toBe("1");
    expect(labels.get("n2")).toBe("1");
    expect(labels.get("n3")).toBe("2");
  });

  it("keeps the first number for a note referenced twice", () => {
    const model = docWithNotes(
      [
        paragraphWithRef("a", "n1", "footnote"),
        paragraphWithRef("b", "n2", "footnote"),
        paragraphWithRef("c", "n1", "footnote"),
      ],
      ["n1", "n2"],
    );
    const labels = noteLabels(model, "footnote", collectNoteRefs(model, "footnote"));
    expect(labels.get("n1")).toBe("1");
    expect(labels.get("n2")).toBe("2");
    expect(labels.size).toBe(2);
  });
});

describe("referencedNotes and note content (SPEC-031 AC-10, AC-11)", () => {
  it("lists only referenced notes, in first-reference order", () => {
    const model = docWithNotes(
      [paragraphWithRef("a", "n2", "footnote"), paragraphWithRef("b", "n1", "footnote")],
      ["n1", "n2", "orphan"],
    );
    expect(referencedNotes(model, "footnote").map((note) => note.noteId)).toEqual(["n2", "n1"]);
  });

  it("skips dangling references without stored content", () => {
    const model = docWithNotes([paragraphWithRef("a", "missing", "footnote")], []);
    expect(referencedNotes(model, "footnote")).toEqual([]);
  });

  it("withNoteBlocks replaces one note's blocks copy-on-write", () => {
    const model = docWithNotes([paragraphWithRef("a", "n1", "footnote")], ["n1"]);
    const next = withNoteBlocks(model, "n1", [paragraph("note body")]);
    expect(noteText(noteById(next, "n1")!)).toBe("note body");
    expect(noteById(model, "n1")!.blocks[0]).toEqual({ type: "paragraph", runs: [] });
    expect(next.notes).not.toBe(model.notes);
  });

  it("noteText joins block texts for the tooltip", () => {
    const note = createEmptyNoteContent("n1", "footnote");
    note.blocks = [paragraph("first"), paragraph("second")];
    expect(noteText(note)).toBe("first\nsecond");
  });
});
