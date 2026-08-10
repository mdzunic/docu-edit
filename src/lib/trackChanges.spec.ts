// SPEC-039 — review-pane listing, navigation and accept/reject.
import { describe, expect, it } from "vitest";
import { asCharacterFormatKey, createDocument, type DocumentModel, type Paragraph, type Table } from "./documentModel";
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

function tableDoc(): DocumentModel {
  const model = createDocument("track-changes-table-cmd-test");
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

describe("collectChanges / accept / reject — table structural changes (AC-5, AC-7, AC-15)", () => {
  it("lists a column insert once, described by column, and accept clears the marker", () => {
    const before = tableDoc();
    const inserted = track(before, insertTableColumn(before, caretSelection(caret(0, 0)), LOCATION, "right").model);
    const entries = collectChanges(inserted);
    expect(entries).toHaveLength(1);
    expect(entries[0]).toMatchObject({ kind: "cell", description: "Inserted table cell" });
    const accepted = acceptChange(inserted, entries[0].changeId);
    expect(collectChanges(accepted)).toHaveLength(0);
    expect(tableOf(accepted).rows[0].cells).toHaveLength(3);
  });

  it("rejecting an inserted column removes the whole column and its width entry", () => {
    const before = tableDoc();
    const inserted = track(before, insertTableColumn(before, caretSelection(caret(0, 0)), LOCATION, "right").model);
    const entries = collectChanges(inserted);
    const rejected = rejectChange(inserted, entries[0].changeId);
    const table = tableOf(rejected);
    expect(table.rows[0].cells).toHaveLength(2);
    expect(table.rows[1].cells).toHaveLength(2);
    expect(table.columnWidthsCm).toHaveLength(2);
    expect(collectChanges(rejected)).toHaveLength(0);
  });

  it("rejecting a deleted column restores it (accept drops it)", () => {
    const before = tableDoc();
    const deleted = track(before, deleteTableColumn(before, caretSelection(caret(0, 0)), LOCATION).model);
    const entries = collectChanges(deleted);
    expect(entries[0]).toMatchObject({ kind: "cell", description: "Deleted table cell" });

    const rejected = rejectChange(deleted, entries[0].changeId);
    expect(tableOf(rejected).rows[0].cells).toHaveLength(2);
    expect(collectChanges(rejected)).toHaveLength(0);

    const accepted = acceptChange(deleted, entries[0].changeId);
    expect(tableOf(accepted).rows[0].cells).toHaveLength(1);
  });

  it("tracks a merge's span as a cell property change; rejecting restores the original span", () => {
    // A merge also moves the covered cell's content into the origin cell,
    // which the generic diff engine tracks as its own insertion/deletion
    // pair (same "no per-command special-casing" design as everything else
    // in trackChangesDiff.ts) — the span change is one further, separate
    // entry among those.
    const before = tableDoc();
    const range = { section: 0, block: 0, anchor: { row: 0, col: 0 }, focus: { row: 0, col: 1 } };
    const merged = track(before, mergeCells(before, caretSelection(caret(0, 0)), range).model);
    const entries = collectChanges(merged);
    const spanChange = entries.find((entry) => entry.kind === "tableProperty");
    expect(spanChange).toMatchObject({ description: "Formatted: Cell merge" });

    const rejected = rejectChange(merged, spanChange!.changeId);
    expect(tableOf(rejected).rows[0].cells[0].colSpan).toBeUndefined();

    const accepted = acceptChange(merged, spanChange!.changeId);
    expect(tableOf(accepted).rows[0].cells[0].colSpan).toBe(2);
  });

  it("tracks a cell shading edit; rejecting restores the previous fill", () => {
    const before = tableDoc();
    const range = { section: 0, block: 0, anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } };
    const shaded = track(before, setCellShading(before, caretSelection(caret(0, 0)), range, "#ff0000").model);
    const entries = collectChanges(shaded);
    expect(entries[0]).toMatchObject({ kind: "tableProperty", description: "Formatted: Cell shading" });
    const rejected = rejectChange(shaded, entries[0].changeId);
    expect(tableOf(rejected).rows[0].cells[0].shadingFill).toBeUndefined();
  });

  it("tracks a row-height edit; rejecting restores the previous height", () => {
    const before = tableDoc();
    const resized = track(before, setRowMinHeight(before, caretSelection(caret(0, 0)), LOCATION, 0, 3).model);
    const entries = collectChanges(resized);
    expect(entries[0]).toMatchObject({ kind: "tableProperty", description: "Formatted: Row height" });
    const rejected = rejectChange(resized, entries[0].changeId);
    expect(tableOf(rejected).rows[0].minHeightCm).toBeUndefined();
  });

  it("tracks a table alignment edit; rejecting restores the previous alignment", () => {
    const before = tableDoc();
    const aligned = track(before, setTableAlignment(before, caretSelection(caret(0, 0)), LOCATION, "center").model);
    const entries = collectChanges(aligned);
    expect(entries[0]).toMatchObject({ kind: "tableProperty", description: "Formatted: Table alignment" });
    const rejected = rejectChange(aligned, entries[0].changeId);
    expect(tableOf(rejected).alignment).toBeUndefined();
  });
});
