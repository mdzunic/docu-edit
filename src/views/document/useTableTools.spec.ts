// SPEC-076 — the Table menu's composable: what each command applies to, why
// it is disabled when it is, and the dialog state each opens with. Driven
// with a hand-built EditorContext, so no component is mounted.
import { describe, expect, it } from "vitest";
import { ref, shallowRef, type Ref } from "vue";
import { useTableTools, NOT_IN_TABLE_REASON, NO_PARAGRAPHS_REASON } from "./useTableTools";
import { caretSelection, type CommandResult, type ModelSelection } from "../../lib/commands";
import { createDocument, type DocumentModel, type Table } from "../../lib/documentModel";
import type { CaretPosition } from "../../lib/textEditing";
import type { TableCellTarget } from "../../lib/domCaret";

function cell(text: string) {
  return { blocks: [{ type: "paragraph" as const, runs: [{ text, formatting: {} }] }] };
}

function fixture(): DocumentModel {
  const table: Table = {
    type: "table",
    columnWidthsCm: [4, 4, 4],
    rows: [
      { isHeader: true, cells: [cell("Name"), cell("Qty"), cell("City")] },
      { cells: [cell("a"), cell("2"), cell("Rome")] },
      { cells: [cell("b"), cell("1"), cell("Oslo")] },
    ],
  };
  const model = createDocument("table-tools");
  model.sections = [{ blocks: [{ type: "paragraph", runs: [{ text: "a\tb", formatting: {} }] }, table] }];
  return model;
}

interface Harness {
  tools: ReturnType<typeof useTableTools>;
  doc: { value: DocumentModel | null };
  caret: Ref<CaretPosition | null>;
  edits: { label: string; result: CommandResult | null }[];
  toasts: string[];
  contextCell: { value: TableCellTarget | null };
}

/** The composable wired to a minimal context: enough of EditorContext to
 * exercise every command, with each applyEdit recorded rather than applied. */
function harness(selection?: ModelSelection): Harness {
  const doc = shallowRef<DocumentModel | null>(fixture());
  const caret = ref<CaretPosition | null>(null);
  const edits: { label: string; result: CommandResult | null }[] = [];
  const toasts: string[] = [];
  const contextCell = { value: null as TableCellTarget | null };
  const noop = () => {};
  const tools = useTableTools({
    doc,
    documentId: ref("d"),
    contentRoot: ref(null),
    pageRef: ref(null),
    editable: ref(true),
    currentSelection: () => null,
    applyEdit: (result, _before, label) => {
      edits.push({ label, result });
      if (result) {
        doc.value = result.model;
      }
    },
    applyCaretMove: noop,
    scrollModelCaretIntoView: async () => {},
    getCaret: () => caret.value,
    getScrollContainer: () => null,
    clearCellSelection: noop,
    addToast: (message) => toasts.push(message),
    markDirty: noop,
    refreshUndoRedoState: noop,
    onSessionClose: noop,
    onUnmount: noop,
    onScroll: noop,
    zoom: ref(1),
    clientToContentPx: () => ({ x: 0, y: 0 }),
    contentPxToClient: () => ({ x: 0, y: 0 }),
    caretPosition: caret,
    paragraphSelection: () =>
      selection ?? caretSelection({ section: 0, block: 0, offset: 0 }),
    getContextMenuTableCell: () => contextCell.value,
    closeContextMenu: noop,
  });
  return { tools, doc, caret, edits, toasts, contextCell };
}

/** Puts the caret in a cell of the fixture's table (block 1). */
function inCell(h: Harness, row = 0, col = 0): void {
  h.caret.value = { section: 0, block: 1, offset: 0, cell: { row, col, block: 0 } };
}

describe("the Table menu's state (AC-17, AC-18, AC-52)", () => {
  it("reports every command inapplicable, with a reason, outside a table", () => {
    const h = harness();
    expect(h.tools.menuState.value.inTable).toBe(false);
    expect(h.tools.menuState.value.reason).toBe(NOT_IN_TABLE_REASON);
    expect(h.tools.menuState.value.convertTextReason).toBe(NO_PARAGRAPHS_REASON);
  });

  it("reports the caret row's own header state (AC-17)", () => {
    const h = harness();
    inCell(h, 0);
    expect(h.tools.menuState.value.inTable).toBe(true);
    expect(h.tools.menuState.value.rowIsHeader).toBe(true);
    inCell(h, 1);
    expect(h.tools.menuState.value.rowIsHeader).toBe(false);
  });

  it("refuses Convert text to table inside a table, with its own reason (AC-52)", () => {
    const h = harness();
    inCell(h, 1);
    expect(h.tools.menuState.value.canConvertText).toBe(false);
    expect(h.tools.menuState.value.convertTextReason).toBe("the selection is inside a table");
  });

  it("offers Convert text to table for a paragraph selection (AC-52)", () => {
    expect(harness().tools.menuState.value.canConvertText).toBe(true);
  });
});

describe("Repeat header row (AC-17, AC-19)", () => {
  it("toggles the caret row and records one undo entry", () => {
    const h = harness();
    inCell(h, 1);
    h.tools.toggleRepeatHeaderRow();
    expect(h.edits).toHaveLength(1);
    expect(h.edits[0].label).toBe("repeat header row");
    expect((h.doc.value!.sections[0].blocks[1] as Table).rows[1].isHeader).toBe(true);
  });

  it("acts on the RIGHT-CLICKED row in the context menu, not the caret's (AC-19)", () => {
    const h = harness();
    inCell(h, 0);
    h.contextCell.value = { section: 0, block: 1, row: 2, col: 0 };
    expect(h.tools.contextRowIsHeader()).toBe(false);
    h.tools.toggleContextRepeatHeaderRow();
    const table = h.doc.value!.sections[0].blocks[1] as Table;
    expect(table.rows[2].isHeader).toBe(true);
    expect(table.rows[0].isHeader).toBe(true); // the caret's row is untouched
  });

  it("says so rather than acting when the caret is not in a table", () => {
    const h = harness();
    h.tools.toggleRepeatHeaderRow();
    expect(h.edits).toHaveLength(0);
    expect(h.toasts).toEqual(["Put the caret in a table first"]);
  });
});

describe("the style gallery (AC-33, AC-34, AC-35, AC-36)", () => {
  it("applies a built-in look and materialises its style in one undo entry (AC-33)", () => {
    const h = harness();
    inCell(h, 1);
    h.tools.applyGalleryStyle("TableStyleAccent");
    expect(h.edits).toHaveLength(1);
    const table = h.doc.value!.sections[0].blocks[1] as Table;
    expect(table.styleId).toBe("TableStyleAccent");
    expect(table.look).toEqual({ firstRow: true, bandedRows: true });
    expect(h.doc.value!.styles?.some((style) => style.id === "TableStyleAccent")).toBe(true);
  });

  it("the None entry clears both fields in one undo entry (AC-34)", () => {
    const h = harness();
    inCell(h, 1);
    h.tools.applyGalleryStyle("TableStyleGrid");
    h.tools.applyGalleryStyle(null);
    expect(h.edits).toHaveLength(2);
    const table = h.doc.value!.sections[0].blocks[1] as Table;
    expect(table.styleId).toBeUndefined();
    expect(table.look).toBeUndefined();
  });

  it("each look checkbox is one undo entry, and toggles back off (AC-35)", () => {
    const h = harness();
    inCell(h, 1);
    h.tools.toggleLookFlag("bandedRows");
    expect((h.doc.value!.sections[0].blocks[1] as Table).look).toEqual({ bandedRows: true });
    h.tools.toggleLookFlag("bandedRows");
    expect((h.doc.value!.sections[0].blocks[1] as Table).look).toEqual({});
    expect(h.edits).toHaveLength(2);
  });

  it("does not open outside a table (AC-36)", () => {
    const h = harness();
    h.tools.openStyleGallery();
    expect(h.tools.dialog.value).toBe("");
  });
});

describe("the sort dialog's defaults and refusals (AC-39, AC-40, AC-49, AC-50)", () => {
  it("seeds key one from the caret's column, ascending, text (AC-39)", () => {
    const h = harness();
    inCell(h, 1, 2);
    h.tools.openSortDialog();
    expect(h.tools.dialog.value).toBe("sort");
    expect(h.tools.sortKeys.value).toEqual([{ columnIndex: 2, direction: "asc", type: "text" }]);
  });

  it("defaults the header toggle to whether the first row is one (AC-40)", () => {
    const h = harness();
    inCell(h, 1);
    h.tools.openSortDialog();
    expect(h.tools.sortHasHeader.value).toBe(true);

    const plain = harness();
    delete (plain.doc.value!.sections[0].blocks[1] as Table).rows[0].isHeader;
    inCell(plain, 1);
    plain.tools.openSortDialog();
    expect(plain.tools.sortHasHeader.value).toBe(false);
  });

  it("refuses a range of fewer than two rows (AC-50)", () => {
    const h = harness();
    const table = h.doc.value!.sections[0].blocks[1] as Table;
    table.rows = [table.rows[0], table.rows[1]];
    inCell(h, 1);
    expect(h.tools.sortBlockedReason(true)).toBe("there is nothing to sort");
    // …and the same table sorts fine when the header row joins the range.
    expect(h.tools.sortBlockedReason(false)).toBe("");
  });

  it("refuses a vertically merged cell but not a horizontally merged one (AC-49)", () => {
    const h = harness();
    const table = h.doc.value!.sections[0].blocks[1] as Table;
    inCell(h, 1);
    table.rows[1].cells[0].colSpan = 2;
    expect(h.tools.sortBlockedReason(true)).toBe("");
    table.rows[1].cells[0].rowSpan = 2;
    expect(h.tools.sortBlockedReason(true)).toBe(
      "sorting is not available for tables with merged rows",
    );
  });

  it("applies the sort as one undo entry and closes (AC-51)", () => {
    const h = harness();
    inCell(h, 1);
    h.tools.openSortDialog();
    h.tools.applySort([{ columnIndex: 1, direction: "asc", type: "number" }], true);
    expect(h.tools.dialog.value).toBe("");
    expect(h.edits).toHaveLength(1);
    expect(h.edits[0].label).toBe("sort table");
    const rows = (h.doc.value!.sections[0].blocks[1] as Table).rows;
    expect(rows.map((row) => row.cells[1].blocks[0].runs[0].text)).toEqual(["Qty", "1", "2"]);
  });
});

describe("the conversions (AC-52, AC-60, AC-61, AC-64)", () => {
  it("opens the text→table dialog for a paragraph selection and applies it", () => {
    const h = harness();
    h.tools.openTextToTable();
    expect(h.tools.dialog.value).toBe("text-to-table");
    expect(h.tools.convertSourceParagraphs.value).toHaveLength(1);
    h.tools.applyTextToTable("tab", "");
    expect(h.edits).toHaveLength(1);
    expect(h.edits[0].label).toBe("convert text to table");
    expect(h.doc.value!.sections[0].blocks[0].type).toBe("table");
  });

  it("opens the table→text dialog only inside a table, and applies it (AC-61, AC-64)", () => {
    const h = harness();
    h.tools.openTableToText();
    expect(h.tools.dialog.value).toBe("");
    inCell(h, 1);
    h.tools.openTableToText();
    expect(h.tools.dialog.value).toBe("table-to-text");
    h.tools.applyTableToText("tab", "");
    expect(h.edits[0].label).toBe("convert table to text");
    expect(h.doc.value!.sections[0].blocks.every((block) => block.type === "paragraph")).toBe(true);
  });
});

describe("the command registry (AC-80)", () => {
  it("registers the five Table ids", () => {
    expect(harness().tools.commands.map(([id]) => id)).toEqual([
      "table.repeat-header-row",
      "table.style-gallery",
      "table.sort",
      "table.convert-text-to-table",
      "table.convert-table-to-text",
    ]);
  });
});
