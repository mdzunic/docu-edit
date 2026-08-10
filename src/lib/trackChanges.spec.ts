// SPEC-039 — review-pane listing, navigation and accept/reject.
import { describe, expect, it } from "vitest";
import { asCharacterFormatKey, createDocument, type DocumentModel, type Paragraph } from "./documentModel";
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
import {
  acceptAllChanges,
  acceptChange,
  collectChanges,
  nextChangeFrom,
  previousChangeFrom,
  rejectAllChanges,
  rejectChange,
} from "./trackChanges";

const AUTHOR = "Ada";
const DATE = "2026-01-01T00:00:00.000Z";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function docWithText(...texts: string[]): DocumentModel {
  const model = createDocument("track-changes-cmd-test");
  model.sections = [{ blocks: texts.map(paragraph) }];
  return model;
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

function selection(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

function track(before: DocumentModel, after: DocumentModel): DocumentModel {
  return applyTrackChanges(before, after, AUTHOR, DATE);
}

function runsOf(model: DocumentModel, block = 0): Paragraph["runs"] {
  return (model.sections[0].blocks[block] as Paragraph).runs;
}

describe("collectChanges", () => {
  it("lists insertions and deletions in document order with author/date/description", () => {
    const before = docWithText("hello world");
    const inserted = track(before, insertTextCommand(before, caretSelection(caret(0, 5)), "XYZ").model);
    const entries = collectChanges(inserted);
    expect(entries).toHaveLength(1);
    expect(entries[0]).toMatchObject({ kind: "insert", author: AUTHOR, date: DATE, description: "Inserted", preview: "XYZ" });
  });

  it("describes a formatting change using the changed property name", () => {
    const before = docWithText("hello world");
    const after = setFormat(before, selection(caret(0, 0), caret(0, 5)), asCharacterFormatKey("bold"), true).model;
    const tracked = track(before, after);
    const entries = collectChanges(tracked);
    expect(entries).toHaveLength(1);
    expect(entries[0].description).toBe("Formatted: Bold");
  });
});

describe("nextChangeFrom / previousChangeFrom (AC-8)", () => {
  it("navigates forward and backward through changes in document order", () => {
    const before = docWithText("aaaa", "bbbb", "cccc");
    const afterFirst = track(before, insertTextCommand(before, caretSelection(caret(0, 2)), "X").model);
    const model = track(afterFirst, insertTextCommand(afterFirst, caretSelection(caret(2, 2)), "Y").model);

    // The caret starts before both changes.
    const first = nextChangeFrom(model, caret(0, 0));
    expect(first?.at).toEqual({ section: 0, block: 0, offset: 2 });
    const second = nextChangeFrom(model, first!.at);
    expect(second?.at).toEqual({ section: 0, block: 2, offset: 2 });
    expect(nextChangeFrom(model, second!.at)).toBeNull();

    expect(previousChangeFrom(model, second!.at)).toMatchObject({ at: first!.at });
    expect(previousChangeFrom(model, first!.at)).toBeNull();
  });
});

describe("acceptChange / rejectChange — insertion", () => {
  it("accepting an insertion commits the text, clearing the marker", () => {
    const before = docWithText("hello world");
    const tracked = track(before, insertTextCommand(before, caretSelection(caret(0, 5)), "XYZ").model);
    const changeId = collectChanges(tracked)[0].changeId;
    const accepted = acceptChange(tracked, changeId);
    expect(runsOf(accepted).map((r) => r.text).join("")).toBe("helloXYZ world");
    expect(collectChanges(accepted)).toHaveLength(0);
  });

  it("rejecting an insertion removes the inserted text (AC-10, AC-11)", () => {
    const before = docWithText("hello world");
    const tracked = track(before, insertTextCommand(before, caretSelection(caret(0, 5)), "XYZ").model);
    const changeId = collectChanges(tracked)[0].changeId;
    const rejected = rejectChange(tracked, changeId);
    expect(runsOf(rejected).map((r) => r.text).join("")).toBe("hello world");
    expect(collectChanges(rejected)).toHaveLength(0);
  });
});

describe("acceptChange / rejectChange — deletion (AC-11)", () => {
  it("accepting a deletion actually removes the struck-through text", () => {
    const before = docWithText("hello world");
    const tracked = track(before, deleteRangeCommand(before, selection(caret(0, 0), caret(0, 6))).model);
    const changeId = collectChanges(tracked)[0].changeId;
    const accepted = acceptChange(tracked, changeId);
    expect(runsOf(accepted).map((r) => r.text).join("")).toBe("world");
  });

  it("rejecting a deletion restores the exact original text and formatting", () => {
    const before = docWithText("hello world");
    (before.sections[0].blocks[0] as Paragraph).runs = [
      { text: "hello ", formatting: { italic: true } },
      { text: "world", formatting: {} },
    ];
    const tracked = track(before, deleteRangeCommand(before, selection(caret(0, 0), caret(0, 6))).model);
    const changeId = collectChanges(tracked)[0].changeId;
    const rejected = rejectChange(tracked, changeId);
    expect(runsOf(rejected)).toEqual([
      { text: "hello ", formatting: { italic: true } },
      { text: "world", formatting: {} },
    ]);
  });
});

describe("acceptChange / rejectChange — formatting change", () => {
  it("rejecting a formatting change restores the prior formatting", () => {
    const before = docWithText("hello world");
    const after = setFormat(before, selection(caret(0, 0), caret(0, 5)), asCharacterFormatKey("bold"), true).model;
    const tracked = track(before, after);
    const changeId = collectChanges(tracked)[0].changeId;
    const rejected = rejectChange(tracked, changeId);
    expect(runsOf(rejected)[0].formatting.bold).toBeUndefined();
    expect(runsOf(rejected).map((r) => r.text).join("")).toBe("hello world");
  });

  it("accepting a formatting change keeps the new formatting and clears the marker", () => {
    const before = docWithText("hello world");
    const after = setFormat(before, selection(caret(0, 0), caret(0, 5)), asCharacterFormatKey("bold"), true).model;
    const tracked = track(before, after);
    const changeId = collectChanges(tracked)[0].changeId;
    const accepted = acceptChange(tracked, changeId);
    expect(runsOf(accepted)[0].formatting.bold).toBe(true);
    expect(collectChanges(accepted)).toHaveLength(0);
  });
});

describe("acceptChange / rejectChange — paragraph split/merge", () => {
  it("rejecting a split's inserted mark re-merges the two paragraphs", () => {
    const before = docWithText("hello world");
    const tracked = track(before, splitBlock(before, caretSelection(caret(0, 5))).model);
    expect(tracked.sections[0].blocks).toHaveLength(2);
    const changeId = collectChanges(tracked).find((e) => e.kind === "paragraphMark")!.changeId;
    const rejected = rejectChange(tracked, changeId);
    expect(rejected.sections[0].blocks).toHaveLength(1);
    expect(runsOf(rejected).map((r) => r.text).join("")).toBe("hello world");
  });

  it("accepting a merge's pending-delete mark commits the physical merge", () => {
    const before = docWithText("hello", "world");
    const tracked = track(before, deleteBackwardCommand(before, selection(caret(1, 0), caret(1, 0))).model);
    expect(tracked.sections[0].blocks).toHaveLength(2);
    const changeId = collectChanges(tracked).find((e) => e.kind === "paragraphMark")!.changeId;
    const accepted = acceptChange(tracked, changeId);
    expect(accepted.sections[0].blocks).toHaveLength(1);
    expect(runsOf(accepted).map((r) => r.text).join("")).toBe("helloworld");
  });

  it("rejecting a merge's pending-delete mark just clears the marker, keeping paragraphs apart", () => {
    const before = docWithText("hello", "world");
    const tracked = track(before, deleteBackwardCommand(before, selection(caret(1, 0), caret(1, 0))).model);
    const changeId = collectChanges(tracked).find((e) => e.kind === "paragraphMark")!.changeId;
    const rejected = rejectChange(tracked, changeId);
    expect(rejected.sections[0].blocks).toHaveLength(2);
    expect(collectChanges(rejected)).toHaveLength(0);
  });
});

describe("acceptAllChanges / rejectAllChanges", () => {
  it("accepts every change in the document", () => {
    const before = docWithText("hello world");
    const afterInsert = track(before, insertTextCommand(before, caretSelection(caret(0, 5)), "XYZ").model);
    const afterDelete = track(afterInsert, deleteRangeCommand(afterInsert, selection(caret(0, 0), caret(0, 5))).model);
    expect(collectChanges(afterDelete).length).toBeGreaterThanOrEqual(2);
    const accepted = acceptAllChanges(afterDelete);
    expect(collectChanges(accepted)).toHaveLength(0);
  });

  it("rejects every change in the document, restoring the original text", () => {
    const before = docWithText("hello world");
    const inserted = insertTextCommand(before, caretSelection(caret(0, 5)), "XYZ").model;
    const tracked = track(before, inserted);
    const rejected = rejectAllChanges(tracked);
    expect(runsOf(rejected).map((r) => r.text).join("")).toBe("hello world");
    expect(collectChanges(rejected)).toHaveLength(0);
  });
});
