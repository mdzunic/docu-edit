// SPEC-039 — selection remapping through the track-changes diff engine.
//
// applyTrackChanges resurrects deleted text into a command's result, which
// shifts offsets relative to the untracked model the command computed its
// selection against. These tests drive the same command → diff → remap
// pipeline DocumentView's applyEdit runs, keystroke by keystroke — the
// regression they pin: typing a replacement over a selection used to leave
// the caret inside the resurrected strike, scattering later keystrokes into
// it ("Oound-tripspensR"), and accept-all then committed the scramble.
import { describe, expect, it } from "vitest";
import { asCharacterFormatKey, createDocument, type DocumentModel, type Paragraph, type Run, type Table } from "./documentModel";
import {
  caretSelection,
  deleteBackward,
  insertText,
  setFormat,
  splitBlock,
  type CommandResult,
  type ModelSelection,
} from "./commands";
import type { CaretPosition } from "./textEditing";
import { applyTrackChanges } from "./trackChangesDiff";
import { acceptAllChanges } from "./trackChanges";
import { remapSelectionThroughTracking } from "./trackChangesRemap";

const AUTHOR = "Ada";
const DATE = "2026-01-01T00:00:00.000Z";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function docWithText(...texts: string[]): DocumentModel {
  const model = createDocument("track-changes-remap-test");
  model.sections = [{ blocks: texts.map(paragraph) }];
  return model;
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

function selection(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

function runsOf(model: DocumentModel, block = 0): Run[] {
  return (model.sections[0].blocks[block] as Paragraph).runs;
}

function textOf(model: DocumentModel, block = 0): string {
  return runsOf(model, block)
    .map((run) => run.text)
    .join("");
}

interface EditorState {
  model: DocumentModel;
  selection: ModelSelection;
}

/** One edit through the exact pipeline applyEdit runs while recording:
 * ordinary command → applyTrackChanges → selection remap. */
function applyTracked(
  state: EditorState,
  command: (model: DocumentModel, selection: ModelSelection) => CommandResult,
): EditorState {
  const result = command(state.model, state.selection);
  const tracked = applyTrackChanges(state.model, result.model, AUTHOR, DATE);
  return { model: tracked, selection: remapSelectionThroughTracking(result.model, tracked, result.selection) };
}

/** Types `text` one keystroke at a time — each character its own edit
 * event, like a person typing (the shape that used to scramble). */
function typeTracked(state: EditorState, text: string): EditorState {
  let current = state;
  for (const ch of text) {
    current = applyTracked(current, (model, sel) => insertText(model, sel, ch));
  }
  return current;
}

describe("remapSelectionThroughTracking — typing over a selection", () => {
  it("keeps the caret after the insertion, not inside the resurrected strike (single keystroke)", () => {
    const before = docWithText("Opens .docx, .odt and Markdown");
    const sel = selection(caret(0, 0), caret(0, 5));
    const result = insertText(before, sel, "R");
    const tracked = applyTrackChanges(before, result.model, AUTHOR, DATE);
    // Tracked text is "OpensR …": the struck "Opens" precedes the typed "R".
    expect(textOf(tracked)).toBe("OpensR .docx, .odt and Markdown");
    const remapped = remapSelectionThroughTracking(result.model, tracked, result.selection);
    // The untracked caret (offset 1, after "R") must land after "R" in the
    // tracked text — offset 6 — not between the strike's "O" and "p".
    expect(remapped.focus).toEqual(caret(0, 6));
    expect(remapped.anchor).toEqual(caret(0, 6));
  });

  it("keeps per-keystroke typing over a selection as one contiguous deletion + insertion pair", () => {
    const before = docWithText("Opens .docx, .odt and Markdown");
    let state: EditorState = { model: before, selection: selection(caret(0, 0), caret(0, 5)) };
    state = typeTracked(state, "Round-trips");

    expect(textOf(state.model)).toBe("OpensRound-trips .docx, .odt and Markdown");
    const runs = runsOf(state.model);
    expect(runs.map((run) => run.text)).toEqual(["Opens", "Round-trips", " .docx, .odt and Markdown"]);
    expect(runs[0].changes?.delete).toMatchObject({ author: AUTHOR });
    expect(runs[0].changes?.insert).toBeUndefined();
    expect(runs[1].changes?.insert).toMatchObject({ author: AUTHOR });
    expect(runs[1].changes?.delete).toBeUndefined();
    expect(runs[2].changes).toBeUndefined();
    // One review-pane change per side: the per-keystroke marks merged.
    const insertIds = new Set(runs.filter((r) => r.changes?.insert).map((r) => r.changes!.insert!.changeId));
    const deleteIds = new Set(runs.filter((r) => r.changes?.delete).map((r) => r.changes!.delete!.changeId));
    expect(insertIds.size).toBe(1);
    expect(deleteIds.size).toBe(1);
  });

  it("accept-all after per-keystroke replace-typing commits the replacement, not a scramble", () => {
    const before = docWithText("Opens .docx, .odt and Markdown");
    let state: EditorState = { model: before, selection: selection(caret(0, 0), caret(0, 5)) };
    state = typeTracked(state, "Round-trips");
    const accepted = acceptAllChanges(state.model);
    expect(textOf(accepted)).toBe("Round-trips .docx, .odt and Markdown");
    expect(runsOf(accepted).every((run) => run.changes === undefined)).toBe(true);
  });
});

describe("remapSelectionThroughTracking — shapes that must keep today's caret", () => {
  it("leaves an Enter-split caret at the start of the new paragraph", () => {
    const before = docWithText("hello world");
    let state: EditorState = { model: before, selection: caretSelection(caret(0, 5)) };
    state = applyTracked(state, (model, sel) => splitBlock(model, sel));
    expect(state.selection.focus).toEqual(caret(1, 0));
    // Typing continues in the tail paragraph, marked inserted there.
    state = typeTracked(state, "X");
    expect(textOf(state.model, 1)).toBe("X world");
    expect(runsOf(state.model, 1)[0].changes?.insert).toMatchObject({ author: AUTHOR });
  });

  it("keeps a backspace caret before the strike, so repeated backspaces extend one deletion", () => {
    const before = docWithText("hello world");
    let state: EditorState = { model: before, selection: caretSelection(caret(0, 11)) };
    state = applyTracked(state, (model, sel) => deleteBackward(model, sel));
    // The struck "d" is retained; the caret sits before it.
    expect(textOf(state.model)).toBe("hello world");
    expect(state.selection.focus).toEqual(caret(0, 10));
    state = applyTracked(state, (model, sel) => deleteBackward(model, sel));
    const deleted = runsOf(state.model).filter((run) => run.changes?.delete);
    expect(deleted.map((run) => run.text).join("")).toBe("ld");
    expect(new Set(deleted.map((run) => run.changes!.delete!.changeId)).size).toBe(1);
  });

  it("returns the selection unchanged (same object) for a formatting-only edit", () => {
    const before = docWithText("hello world");
    const sel = selection(caret(0, 0), caret(0, 5));
    const result = setFormat(before, sel, asCharacterFormatKey("bold"), true);
    const tracked = applyTrackChanges(before, result.model, AUTHOR, DATE);
    expect(remapSelectionThroughTracking(result.model, tracked, result.selection)).toBe(result.selection);
  });
});

describe("remapSelectionThroughTracking — multi-paragraph replacement", () => {
  it("shifts the caret past the resurrected paragraphs, and later keystrokes stay contiguous", () => {
    const before = docWithText("hello world", "foo bar");
    let state: EditorState = { model: before, selection: selection(caret(0, 6), caret(1, 3)) };
    state = typeTracked(state, "X");
    // Both source paragraphs are retained (struck) ahead of the kept result.
    expect(state.model.sections[0].blocks).toHaveLength(3);
    expect(runsOf(state.model, 0).every((run) => run.changes?.delete)).toBe(true);
    expect(runsOf(state.model, 1).every((run) => run.changes?.delete)).toBe(true);
    expect(textOf(state.model, 2)).toBe("hello X bar");
    expect(state.selection.focus).toEqual(caret(2, 7));
    state = typeTracked(state, "Y");
    expect(textOf(state.model, 2)).toBe("hello XY bar");
  });
});

describe("remapSelectionThroughTracking — table cells", () => {
  function tableDoc(): DocumentModel {
    const model = createDocument("track-changes-remap-table-test");
    const table: Table = {
      type: "table",
      columnWidthsCm: [5, 5],
      rows: [
        { cells: [{ blocks: [paragraph("a1")] }, { blocks: [paragraph("a2")] }] },
        { cells: [{ blocks: [paragraph("b1")] }, { blocks: [paragraph("b2")] }] },
      ],
    };
    model.sections = [{ blocks: [table] }];
    return model;
  }

  it("keeps per-keystroke replace-typing inside a cell contiguous", () => {
    const before = tableDoc();
    const cellCaret = (offset: number): CaretPosition => ({
      section: 0,
      block: 0,
      offset,
      cell: { row: 0, col: 0, block: 0 },
    });
    let state: EditorState = { model: before, selection: selection(cellCaret(0), cellCaret(2)) };
    state = typeTracked(state, "XY");
    const cellRuns = ((state.model.sections[0].blocks[0] as Table).rows[0].cells[0].blocks[0] as Paragraph).runs;
    expect(cellRuns.map((run) => run.text)).toEqual(["a1", "XY"]);
    expect(cellRuns[0].changes?.delete).toMatchObject({ author: AUTHOR });
    expect(cellRuns[1].changes?.insert).toMatchObject({ author: AUTHOR });
    expect(state.selection.focus).toEqual(cellCaret(4));
  });
});
