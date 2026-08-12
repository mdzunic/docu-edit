// SPEC-046 — mail-merge editing commands: inserting a chip at the caret,
// and toggling "Table -> Repeat row for each..." on a table row.
import { describe, expect, it } from "vitest";
import { caretSelection, insertMergeChip, type ModelSelection } from "./commands";
import { createDocument, type DocumentModel, type Paragraph, type Run, type Table } from "./documentModel";
import { toggleRowRepeat, type TableLocation } from "./tables";
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
  const model = createDocument("merge-commands-test");
  model.sections = [{ blocks: [paragraph(["Hello world"])] }];
  return model;
}

describe("insertMergeChip (AC-12, AC-13, AC-14)", () => {
  it("inserts a one-character field run at the caret", () => {
    const model = fixture();
    const { model: next, selection } = insertMergeChip(model, caretSelection(caret(0, 5)), { kind: "merge", name: "client_name" });
    const block = next.sections[0].blocks[0] as Paragraph;
    expect(block.runs.map((run) => run.text)).toEqual(["Hello", "1", " world"]);
    expect(block.runs[1].field).toEqual({ kind: "merge", name: "client_name" });
    expect(selection.focus).toEqual(caret(0, 6));
  });

  it("replaces a non-collapsed selection (typing-over rule)", () => {
    const model = fixture();
    const { model: next } = insertMergeChip(model, range(caret(0, 0), caret(0, 5)), { kind: "if", name: "paid" });
    const block = next.sections[0].blocks[0] as Paragraph;
    expect(block.runs.map((run) => run.text)).toEqual(["1", " world"]);
  });

  it("inherits surrounding formatting but drops link fields", () => {
    const model = fixture();
    model.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "Hello", formatting: { bold: true, linkHref: "https://example.com" } }] },
    ];
    const { model: next } = insertMergeChip(model, caretSelection(caret(0, 5)), { kind: "merge", name: "x" });
    const block = next.sections[0].blocks[0] as Paragraph;
    expect(block.runs[1].formatting.bold).toBe(true);
    expect(block.runs[1].formatting.linkHref).toBeUndefined();
  });

  it("does not mutate the input model", () => {
    const model = fixture();
    insertMergeChip(model, caretSelection(caret(0, 5)), { kind: "merge", name: "x" });
    expect((model.sections[0].blocks[0] as Paragraph).runs).toHaveLength(1);
  });
});

function tableFixture(): { model: DocumentModel; location: TableLocation } {
  const model = createDocument("merge-table-commands-test");
  const table: Table = {
    type: "table",
    columnWidthsCm: [5, 5],
    rows: [{ cells: [{ blocks: [paragraph(["SKU"])] }, { blocks: [paragraph(["Qty"])] }] }],
  };
  model.sections = [{ blocks: [table] }];
  return { model, location: { section: 0, block: 0, row: 0, col: 0 } };
}

describe("toggleRowRepeat (AC-29, AC-30)", () => {
  it("inserts an {{#each name}} chip at the start of the row's first cell", () => {
    const { model, location } = tableFixture();
    const { model: next } = toggleRowRepeat(model, caretSelection(caret(0, 0)), location, "line_items");
    const table = next.sections[0].blocks[0] as Table;
    const firstCellRuns = table.rows[0].cells[0].blocks[0].runs;
    expect(firstCellRuns[0].field).toEqual({ kind: "each", name: "line_items" });
    expect(firstCellRuns[0].text).toBe("1");
    expect(firstCellRuns.map((r) => r.text).slice(1)).toEqual(["SKU"]);
  });

  it("removes the chip when called again on an already-repeating row", () => {
    const { model, location } = tableFixture();
    const once = toggleRowRepeat(model, caretSelection(caret(0, 0)), location, "line_items").model;
    const { model: twice } = toggleRowRepeat(once, caretSelection(caret(0, 0)), location, "line_items");
    const table = twice.sections[0].blocks[0] as Table;
    expect(table.rows[0].cells[0].blocks[0].runs.every((r) => !r.field)).toBe(true);
    expect(table.rows[0].cells[0].blocks[0].runs.map((r) => r.text)).toEqual(["SKU"]);
  });

  it("is a no-op on a header row", () => {
    const { model, location } = tableFixture();
    model.sections[0].blocks[0] = { ...(model.sections[0].blocks[0] as Table) };
    (model.sections[0].blocks[0] as Table).rows[0].isHeader = true;
    const { model: next } = toggleRowRepeat(model, caretSelection(caret(0, 0)), location, "line_items");
    const table = next.sections[0].blocks[0] as Table;
    expect(table.rows[0].cells[0].blocks[0].runs.every((r) => !r.field)).toBe(true);
  });

  it("does not mutate the input model", () => {
    const { model, location } = tableFixture();
    toggleRowRepeat(model, caretSelection(caret(0, 0)), location, "line_items");
    const table = model.sections[0].blocks[0] as Table;
    expect(table.rows[0].cells[0].blocks[0].runs.every((r) => !r.field)).toBe(true);
  });
});
