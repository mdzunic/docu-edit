// SPEC-039 — the track-changes diff engine: converting an ordinary
// command's before/after model pair into tracked-change markup.
import { describe, expect, it } from "vitest";
import { asCharacterFormatKey, createDocument, type DocumentModel, type Paragraph, type Run } from "./documentModel";
import {
  caretSelection,
  deleteBackward as deleteBackwardCommand,
  deleteRange as deleteRangeCommand,
  insertText as insertTextCommand,
  setFormat,
  splitBlock,
  type ModelSelection,
} from "./commands";
import type { CaretPosition } from "./textEditing";
import { applyTrackChanges } from "./trackChangesDiff";

const AUTHOR = "Ada";
const OTHER_AUTHOR = "Grace";
const DATE = "2026-01-01T00:00:00.000Z";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function docWithText(...texts: string[]): DocumentModel {
  const model = createDocument("track-changes-test");
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

describe("applyTrackChanges — insertion (AC-4, AC-13)", () => {
  it("marks freshly typed text as an insertion by the given author, leaving the rest untouched", () => {
    const before = docWithText("hello world");
    // "XYZ" shares no boundary character with its neighbours, so the
    // inserted span is unambiguous.
    const after = insertTextCommand(before, caretSelection(caret(0, 5)), "XYZ").model;
    const tracked = applyTrackChanges(before, after, AUTHOR, DATE);
    const runs = runsOf(tracked);
    expect(runs.map((r) => r.text).join("")).toBe("helloXYZ world");
    const inserted = runs.find((r) => r.text === "XYZ");
    expect(inserted?.changes?.insert).toMatchObject({ author: AUTHOR, date: DATE });
    expect(runs[0].changes).toBeUndefined();
  });
});

describe("applyTrackChanges — deletion (AC-4, AC-11, AC-17)", () => {
  it("retains deleted text, marked struck-through, instead of removing it", () => {
    const before = docWithText("hello world");
    const after = deleteRangeCommand(before, selection(caret(0, 0), caret(0, 6))).model;
    const tracked = applyTrackChanges(before, after, AUTHOR, DATE);
    const runs = runsOf(tracked);
    // The full original text survives — nothing was actually removed.
    expect(runs.map((r) => r.text).join("")).toBe("hello world");
    const deleted = runs.find((r) => r.text === "hello ");
    expect(deleted?.changes?.delete).toMatchObject({ author: AUTHOR, date: DATE });
  });

  it("rejecting the tracked deletion (dropping the marker) restores the exact original formatting", () => {
    const before = docWithText("hello world");
    (before.sections[0].blocks[0] as Paragraph).runs = [
      { text: "hello ", formatting: { bold: true } },
      { text: "world", formatting: {} },
    ];
    const after = deleteRangeCommand(before, selection(caret(0, 0), caret(0, 6))).model;
    const tracked = applyTrackChanges(before, after, AUTHOR, DATE);
    const deletedRun = runsOf(tracked).find((r) => r.text === "hello ");
    // The pre-deletion formatting (bold) is exactly what was there before —
    // rejecting just needs to drop the `delete` marker to restore it.
    expect(deletedRun?.formatting).toEqual({ bold: true });
  });
});

describe("applyTrackChanges — same-author collapse vs nested change (AC-12)", () => {
  it("deleting your own still-pending insertion removes it outright", () => {
    const before = docWithText("hello world");
    const typed = insertTextCommand(before, caretSelection(caret(0, 5)), "XYZ").model;
    const withInsert = applyTrackChanges(before, typed, AUTHOR, DATE);
    // "XYZ" now lives at offsets [5, 8) — delete it (same author).
    const deletedAgain = deleteRangeCommand(withInsert, selection(caret(0, 5), caret(0, 8))).model;
    const finalModel = applyTrackChanges(withInsert, deletedAgain, AUTHOR, DATE);
    expect(runsOf(finalModel).map((r) => r.text).join("")).toBe("hello world");
    expect(runsOf(finalModel).some((r) => r.changes)).toBe(false);
  });

  it("another author deleting your still-pending insertion nests the deletion instead of losing it", () => {
    const before = docWithText("hello world");
    const typed = insertTextCommand(before, caretSelection(caret(0, 5)), "XYZ").model;
    const withInsert = applyTrackChanges(before, typed, AUTHOR, DATE);
    const deletedByOther = deleteRangeCommand(withInsert, selection(caret(0, 5), caret(0, 8))).model;
    const finalModel = applyTrackChanges(withInsert, deletedByOther, OTHER_AUTHOR, DATE);
    const runs = runsOf(finalModel);
    // Text is retained (not lost) and carries BOTH the original insert and
    // the new delete mark.
    expect(runs.map((r) => r.text).join("")).toBe("helloXYZ world");
    const nested = runs.find((r) => r.text === "XYZ");
    expect(nested?.changes?.insert?.author).toBe(AUTHOR);
    expect(nested?.changes?.delete?.author).toBe(OTHER_AUTHOR);
  });
});

describe("applyTrackChanges — formatting change (AC-5)", () => {
  it("marks a same-text formatting change with the prior formatting for rollback", () => {
    const before = docWithText("hello world");
    const after = setFormat(before, selection(caret(0, 0), caret(0, 5)), asCharacterFormatKey("bold"), true).model;
    const tracked = applyTrackChanges(before, after, AUTHOR, DATE);
    const runs = runsOf(tracked);
    expect(runs.map((r) => r.text).join("")).toBe("hello world");
    const bolded = runs.find((r) => r.text === "hello");
    expect(bolded?.formatting.bold).toBe(true);
    expect(bolded?.changes?.formatChange).toMatchObject({ author: AUTHOR, date: DATE, previous: {} });
  });
});

describe("applyTrackChanges — paragraph split/merge (AC-15)", () => {
  it("marks a split paragraph's new mark as inserted, keeping the tail's own state", () => {
    const before = docWithText("hello world");
    const after = splitBlock(before, caretSelection(caret(0, 5))).model;
    const tracked = applyTrackChanges(before, after, AUTHOR, DATE);
    expect(tracked.sections[0].blocks).toHaveLength(2);
    const first = tracked.sections[0].blocks[0] as Paragraph;
    const second = tracked.sections[0].blocks[1] as Paragraph;
    expect(first.runs.map((r) => r.text).join("")).toBe("hello");
    expect(second.runs.map((r) => r.text).join("")).toBe(" world");
    expect(first.changes?.markInsert).toMatchObject({ author: AUTHOR, date: DATE });
  });

  it("marks a merged paragraph's mark as pending-deleted without physically merging it", () => {
    const before = docWithText("hello", "world");
    const after = deleteBackwardCommand(before, selection(caret(1, 0), caret(1, 0))).model;
    // The raw (untracked) command really does merge the two paragraphs.
    expect(after.sections[0].blocks).toHaveLength(1);
    const tracked = applyTrackChanges(before, after, AUTHOR, DATE);
    // Tracking keeps both paragraphs, the first's own mark pending deletion.
    expect(tracked.sections[0].blocks).toHaveLength(2);
    const first = tracked.sections[0].blocks[0] as Paragraph;
    const second = tracked.sections[0].blocks[1] as Paragraph;
    expect(first.runs.map((r) => r.text).join("")).toBe("hello");
    expect(second.runs.map((r) => r.text).join("")).toBe("world");
    expect(first.changes?.markDelete).toMatchObject({ author: AUTHOR, date: DATE });
  });
});

describe("applyTrackChanges — no-op paths", () => {
  it("returns the same model reference when nothing changed", () => {
    const before = docWithText("hello world");
    const after = before;
    expect(applyTrackChanges(before, after, AUTHOR, DATE)).toBe(after);
  });

  it("leaves untouched paragraphs completely alone", () => {
    const before = docWithText("first", "second", "third");
    const after = insertTextCommand(before, caretSelection(caret(1, 6)), "!").model;
    const tracked = applyTrackChanges(before, after, AUTHOR, DATE);
    expect(tracked.sections[0].blocks[0]).toBe(before.sections[0].blocks[0]);
    expect(tracked.sections[0].blocks[2]).toBe(before.sections[0].blocks[2]);
  });
});
