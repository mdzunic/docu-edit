// SPEC-039 — the track-changes diff engine: converting an ordinary
// command's before/after model pair into tracked-change markup.
import { describe, expect, it } from "vitest";
import { asCharacterFormatKey, createDocument, type DocumentModel, type Paragraph, type Run, type Table } from "./documentModel";
import {
  caretSelection,
  deleteBackward as deleteBackwardCommand,
  deleteRange as deleteRangeCommand,
  insertText as insertTextCommand,
  setFormat,
  splitBlock,
  type ModelSelection,
} from "./commands";
import {
  deleteTableColumn,
  insertTableColumn,
  mergeCells,
  setCellShading,
  setRowMinHeight,
  setTableAlignment,
  type TableLocation,
} from "./tables";
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

describe("applyTrackChanges — merging consecutive same-author edits (real typing/backspacing)", () => {
  it("merges three separate keystrokes into one insert run and one change, not three", () => {
    let model = docWithText("hello world");
    const original = model;
    let offset = 0;
    for (const ch of ["X", "Y", "Z"]) {
      const before = model;
      const after = insertTextCommand(before, caretSelection(caret(0, offset)), ch).model;
      model = applyTrackChanges(before, after, AUTHOR, DATE);
      offset += 1;
    }
    const runs = runsOf(model);
    expect(runs.map((r) => r.text)).toEqual(["XYZ", "hello world"]);
    expect(runs[0].changes?.insert).toBeDefined();
    expect(model.sections[0].blocks[0]).not.toBe(original.sections[0].blocks[0]);
  });

  it("merges three separate backspaces into one delete run, not three", () => {
    let model = docWithText("hello");
    for (let i = 0; i < 3; i++) {
      const before = model;
      const after = deleteRangeCommand(before, selection(caret(0, 5 - i - 1), caret(0, 5 - i))).model;
      model = applyTrackChanges(before, after, AUTHOR, DATE);
    }
    const runs = runsOf(model);
    // "hello" -> deleting the last 3 chars one at a time ("o", "l", "l") should
    // leave "he" untouched and "llo" as one merged deleted run.
    expect(runs.map((r) => r.text)).toEqual(["he", "llo"]);
    expect(runs[1].changes?.delete).toBeDefined();
  });

  it("does not merge insertions from two different authors", () => {
    const before = docWithText("hello world");
    const afterA = insertTextCommand(before, caretSelection(caret(0, 0)), "A").model;
    const trackedA = applyTrackChanges(before, afterA, "Ada", DATE);
    const afterB = insertTextCommand(trackedA, caretSelection(caret(0, 1)), "B").model;
    const trackedB = applyTrackChanges(trackedA, afterB, "Grace", DATE);
    const runs = runsOf(trackedB);
    expect(runs.map((r) => r.text)).toEqual(["A", "B", "hello world"]);
    expect(runs[0].changes?.insert?.author).toBe("Ada");
    expect(runs[1].changes?.insert?.author).toBe("Grace");
  });
});

function tableDoc(): DocumentModel {
  const model = createDocument("track-changes-table-test");
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

function tableOf(model: DocumentModel): Table {
  return model.sections[0].blocks[0] as Table;
}

const LOCATION: TableLocation = { section: 0, block: 0, row: 0, col: 0 };

describe("applyTrackChanges — table structural changes (AC-15)", () => {
  it("marks an inserted column's new cells as tracked-inserted", () => {
    const before = tableDoc();
    const after = insertTableColumn(before, caretSelection(caret(0, 0)), LOCATION, "right").model;
    const tracked = applyTrackChanges(before, after, AUTHOR, DATE);
    const table = tableOf(tracked);
    expect(table.rows[0].cells[1].changes?.insert).toMatchObject({ author: AUTHOR });
    expect(table.rows[1].cells[1].changes?.insert).toMatchObject({ author: AUTHOR });
    // The untouched original cells are not marked.
    expect(table.rows[0].cells[0].changes).toBeUndefined();
    expect(table.rows[0].cells[2].changes).toBeUndefined();
    // Both new cells share the one changeId (one edit event, AC-7).
    expect(table.rows[0].cells[1].changes?.insert?.changeId).toBe(table.rows[1].cells[1].changes?.insert?.changeId);
  });

  it("marks a deleted column's cells as tracked-deleted instead of removing them", () => {
    const before = tableDoc();
    const after = deleteTableColumn(before, caretSelection(caret(0, 0)), LOCATION).model;
    const tracked = applyTrackChanges(before, after, AUTHOR, DATE);
    const table = tableOf(tracked);
    // Retained (not removed) — same shape as a tracked row/text deletion.
    expect(table.rows).toHaveLength(2);
    expect(table.rows[0].cells).toHaveLength(2);
    expect(table.rows[0].cells[0].changes?.delete).toMatchObject({ author: AUTHOR });
    expect(table.rows[1].cells[0].changes?.delete).toMatchObject({ author: AUTHOR });
  });

  it("marks a cell merge as a pending property change recording the previous span", () => {
    const before = tableDoc();
    const range = { section: 0, block: 0, anchor: { row: 0, col: 0 }, focus: { row: 0, col: 1 } };
    const after = mergeCells(before, caretSelection(caret(0, 0)), range).model;
    const tracked = applyTrackChanges(before, after, AUTHOR, DATE);
    const cell = tableOf(tracked).rows[0].cells[0];
    expect(cell.colSpan).toBe(2);
    expect(cell.changes?.propertyChange).toMatchObject({
      author: AUTHOR,
      previous: { colSpan: undefined, rowSpan: undefined },
    });
  });

  it("marks a cell shading edit as a pending property change, restorable on reject", () => {
    const before = tableDoc();
    const range = { section: 0, block: 0, anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } };
    const after = setCellShading(before, caretSelection(caret(0, 0)), range, "#ff0000").model;
    const tracked = applyTrackChanges(before, after, AUTHOR, DATE);
    const cell = tableOf(tracked).rows[0].cells[0];
    expect(cell.shadingFill).toBe("#ff0000");
    expect(cell.changes?.propertyChange).toMatchObject({ author: AUTHOR, previous: { shadingFill: undefined } });
  });

  it("marks a row height edit as a pending row property change", () => {
    const before = tableDoc();
    const after = setRowMinHeight(before, caretSelection(caret(0, 0)), LOCATION, 0, 3).model;
    const tracked = applyTrackChanges(before, after, AUTHOR, DATE);
    const row = tableOf(tracked).rows[0];
    expect(row.minHeightCm).toBeCloseTo(3);
    expect(row.changes?.propertyChange).toMatchObject({ author: AUTHOR, previous: { minHeightCm: undefined } });
  });

  it("marks a table alignment edit as a pending table property change", () => {
    const before = tableDoc();
    const after = setTableAlignment(before, caretSelection(caret(0, 0)), LOCATION, "center").model;
    const tracked = applyTrackChanges(before, after, AUTHOR, DATE);
    const table = tableOf(tracked);
    expect(table.alignment).toBe("center");
    expect(table.changes?.propertyChange).toMatchObject({ author: AUTHOR, previous: { alignment: undefined } });
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
