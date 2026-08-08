// SPEC-031 — footnote/endnote commands: inserting reference marks,
// reference/note integrity under deletion, note-content editing through the
// synthetic-model seam, and the numbering settings command.
import { describe, expect, it } from "vitest";
import {
  caretSelection,
  commitNoteModel,
  deleteBackward,
  insertNote,
  insertText,
  setNoteSettings,
  withNoteModel,
  type ModelSelection,
} from "./commands";
import { createDocument, type DocumentModel, type Paragraph, type Run } from "./documentModel";
import { collectNoteRefs, noteById, noteLabels, referencedNotes } from "./notes";
import type { CaretPosition } from "./textEditing";

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} }) as Run),
  };
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

function range(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

function fixture(): DocumentModel {
  const model = createDocument("notes-commands-test");
  model.sections = [{ blocks: [paragraph(["Hello world"]), paragraph(["Second"])] }];
  return model;
}

describe("insertNote (SPEC-031 AC-1..4, AC-7, AC-8)", () => {
  it("inserts a superscript reference mark at the caret and stores an empty note", () => {
    const model = fixture();
    const { model: next, selection } = insertNote(model, caretSelection(caret(0, 5)), "footnote", "n1");
    const block = next.sections[0].blocks[0] as Paragraph;
    expect(block.runs.map((run) => run.text)).toEqual(["Hello", "1", " world"]);
    expect(block.runs[1].note).toEqual({ noteId: "n1", kind: "footnote" });
    expect(block.runs[1].formatting.vertAlign).toBe("superscript");
    expect(noteById(next, "n1")?.blocks).toEqual([{ type: "paragraph", runs: [] }]);
    expect(selection.focus).toEqual({ section: 0, block: 0, offset: 6 });
    // The input model is untouched (undo contract).
    expect((model.sections[0].blocks[0] as Paragraph).runs).toHaveLength(1);
    expect(model.notes).toBeUndefined();
  });

  it("generates a fresh note id when none is passed", () => {
    const { model: next } = insertNote(fixture(), caretSelection(caret(0, 0)), "endnote");
    const refs = collectNoteRefs(next, "endnote");
    expect(refs).toHaveLength(1);
    expect(noteById(next, refs[0].noteId)?.kind).toBe("endnote");
  });

  it("replaces a non-collapsed selection with the mark", () => {
    const { model: next } = insertNote(fixture(), range(caret(0, 5), caret(0, 11)), "footnote", "n1");
    const block = next.sections[0].blocks[0] as Paragraph;
    expect(block.runs.map((run) => run.text)).toEqual(["Hello", "1"]);
  });

  it("renumbers existing notes when one is inserted before them (AC-8)", () => {
    let state = insertNote(fixture(), caretSelection(caret(0, 11)), "footnote", "late").model;
    state = insertNote(state, caretSelection(caret(0, 5)), "footnote", "early").model;
    const labels = noteLabels(state, "footnote", collectNoteRefs(state, "footnote"));
    expect(labels.get("early")).toBe("1");
    expect(labels.get("late")).toBe("2");
  });

  it("does not join a link when inserted at a link's edge", () => {
    const model = createDocument("notes-link-test");
    model.sections = [
      { blocks: [paragraph(["visit ", {}], ["here", { linkHref: "https://example.com" }])] },
    ];
    const { model: next } = insertNote(model, caretSelection(caret(0, 10)), "footnote", "n1");
    const block = next.sections[0].blocks[0] as Paragraph;
    const mark = block.runs.find((run) => run.note);
    expect(mark?.formatting.linkHref).toBeUndefined();
  });

  it("is inert on a table-atom caret", () => {
    const model = createDocument("notes-atom-test");
    model.sections = [
      {
        blocks: [
          { type: "table", columnWidthsCm: [5], rows: [{ cells: [{ blocks: [paragraph(["x"])] }] }] },
        ],
      },
    ];
    const { model: next } = insertNote(model, caretSelection(caret(0, 0)), "footnote", "n1");
    expect(collectNoteRefs(next, "footnote")).toEqual([]);
    expect(next.notes).toBeUndefined();
  });
});

describe("reference/note integrity (SPEC-031 AC-9..11)", () => {
  it("deleting the reference mark removes the note from every derived view (AC-9, AC-10)", () => {
    const inserted = insertNote(fixture(), caretSelection(caret(0, 5)), "footnote", "n1");
    const { model: next } = deleteBackward(inserted.model, inserted.selection);
    expect(collectNoteRefs(next, "footnote")).toEqual([]);
    expect(referencedNotes(next, "footnote")).toEqual([]);
    expect((next.sections[0].blocks[0] as Paragraph).runs.map((run) => run.text)).toEqual([
      "Hello world",
    ]);
  });

  it("renumbers the remaining notes after a deletion (AC-9)", () => {
    let state = insertNote(fixture(), caretSelection(caret(0, 2)), "footnote", "a").model;
    state = insertNote(state, caretSelection(caret(0, 6)), "footnote", "b").model;
    state = insertNote(state, caretSelection(caret(0, 10)), "footnote", "c").model;
    // Delete the first mark (offset 2..3).
    state = deleteBackward(state, caretSelection(caret(0, 3))).model;
    const labels = noteLabels(state, "footnote", collectNoteRefs(state, "footnote"));
    expect(labels.get("b")).toBe("1");
    expect(labels.get("c")).toBe("2");
    expect(labels.has("a")).toBe(false);
  });

  it("deleting all of the note's text keeps the reference mark (AC-11)", () => {
    const inserted = insertNote(fixture(), caretSelection(caret(0, 5)), "footnote", "n1");
    let model = inserted.model;
    // Type into the note, then delete it all through the synthetic seam.
    let synthetic = withNoteModel(model, "n1");
    const typed = insertText(synthetic, caretSelection(caret(0, 0)), "gone soon");
    model = commitNoteModel(model, "n1", typed.model);
    synthetic = withNoteModel(model, "n1");
    const cleared = deleteBackward(
      { ...synthetic },
      range(caret(0, 0), caret(0, 9)),
    );
    model = commitNoteModel(model, "n1", cleared.model);
    expect(collectNoteRefs(model, "footnote")).toHaveLength(1);
    expect(noteById(model, "n1")?.blocks).toEqual([{ type: "paragraph", runs: [] }]);
  });
});

describe("withNoteModel/commitNoteModel (SPEC-031 AC-6)", () => {
  it("edits note content through the body-editing primitives", () => {
    const inserted = insertNote(fixture(), caretSelection(caret(0, 5)), "footnote", "n1");
    const synthetic = withNoteModel(inserted.model, "n1");
    expect(synthetic.sections).toHaveLength(1);
    const typed = insertText(synthetic, caretSelection(caret(0, 0)), "Note body");
    const committed = commitNoteModel(inserted.model, "n1", typed.model);
    expect((noteById(committed, "n1")!.blocks[0] as Paragraph).runs[0].text).toBe("Note body");
    // The body sections come back untouched.
    expect(committed.sections).toBe(inserted.model.sections);
  });

  it("starts a missing note from one empty paragraph", () => {
    const synthetic = withNoteModel(fixture(), "missing");
    expect(synthetic.sections[0].blocks).toEqual([{ type: "paragraph", runs: [] }]);
  });
});

describe("setNoteSettings (SPEC-031 AC-17..20)", () => {
  it("stores the patch on the addressed stream only", () => {
    const { model: next } = setNoteSettings(fixture(), caretSelection(caret(0, 0)), "footnote", {
      numberFormat: "symbol",
      startAt: 3,
      restart: "page",
    });
    expect(next.footnoteSettings).toEqual({ numberFormat: "symbol", startAt: 3, restart: "page" });
    expect(next.endnoteSettings).toBeUndefined();
  });

  it("leaves omitted fields unchanged and clears a continuous restart", () => {
    let state = setNoteSettings(fixture(), caretSelection(caret(0, 0)), "footnote", {
      numberFormat: "lowerAlpha",
      restart: "section",
    }).model;
    state = setNoteSettings(state, caretSelection(caret(0, 0)), "footnote", {
      restart: "continuous",
    }).model;
    expect(state.footnoteSettings).toEqual({ numberFormat: "lowerAlpha" });
  });
});
