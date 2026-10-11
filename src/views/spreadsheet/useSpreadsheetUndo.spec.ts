// SPEC-121 "The undo stack" and "The edit pipeline": snapshots by reference
// (a 200,000-cell workbook's snapshot allocates no Row, Cell or OpaquePart),
// the 100-step cap, labels, the flags undo restores, the pipeline's keys and
// dirty marking, and the shortcuts.
import { effectScope, shallowRef } from "vue";
import { afterEach, describe, expect, it, vi } from "vitest";
import { setRowsHidden } from "../../lib/spreadsheet/bands";
import { cellKey, type CellKey } from "../../lib/spreadsheet/formula/dependencies";
import type { Row, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { addSheet, createBlankWorkbook, deleteSheet, renameSheet, setTabColor } from "../../lib/spreadsheet/sheets";
import { MAX_UNDO_STEPS, SHEET_BY_REFERENCE_FIELDS, takeSnapshot, useSpreadsheetUndo, WORKBOOK_SNAPSHOT_FIELDS } from "./useSpreadsheetUndo";

const scopes: Array<() => void> = [];
afterEach(() => {
  for (const stop of scopes.splice(0)) {
    stop();
  }
  vi.restoreAllMocks();
});

function rowsOf(cells: number, width = 10): Row[] {
  const rows: Row[] = [];
  for (let index = 0; index * width < cells; index += 1) {
    rows.push({ index, cells: Array.from({ length: Math.min(width, cells - index * width) }, (_, col) => ({ col, value: index * width + col })) });
  }
  return rows;
}

/** A workbook whose first sheet holds a formula naming "Data". */
function book(): SpreadsheetModel {
  const model = createBlankWorkbook();
  addSheet(model, 0);
  model.sheets = model.sheets.map((sheet, index) =>
    index === 0 ? { ...sheet, rows: [{ index: 0, cells: [{ col: 0, value: 4, formula: "Sheet2!A1*2" }] }], dirty: undefined } : sheet,
  ) as Sheet[];
  delete model.sheets[0]!.dirty;
  delete model.structureDirty;
  model.activeSheetIndex = 0;
  return model;
}

function mount(model: SpreadsheetModel) {
  const scope = effectScope();
  const recalls: CellKey[][] = [];
  const applied = vi.fn();
  const activated: string[] = [];
  let pendingDone: ((ids: string[]) => void) | null = null;
  const ref = shallowRef<SpreadsheetModel | null>(model);
  const undo = scope.run(() =>
    useSpreadsheetUndo({
      model: ref,
      recalc: (changed, done) => {
        recalls.push(changed);
        pendingDone = done;
      },
      onApplied: applied,
      activateSheet: (id) => activated.push(id),
    }),
  )!;
  scopes.push(() => scope.stop());
  return { undo, recalls, applied, activated, finishPass: (ids: string[] = []) => pendingDone?.(ids) };
}

describe("snapshots are by reference", () => {
  it("a snapshot of a 200,000-cell workbook allocates no Row, Cell or OpaquePart", () => {
    const model = createBlankWorkbook();
    model.sheets = [{ ...model.sheets[0]!, rows: rowsOf(200_000) }];
    model.opaque = [{ name: "a.xml", data: "AAAA" }];
    const clone = vi.spyOn(globalThis, "structuredClone");
    const snapshot = takeSnapshot(model);
    // Only small fields were cloned: never a rows array, a row, a cell or a part.
    for (const [value] of clone.mock.calls) {
      expect(Array.isArray(value) && (value as unknown[]).some((item) => typeof item === "object" && item !== null && ("cells" in item || "col" in item || "data" in item))).toBe(false);
    }
    expect(snapshot.sheets[0]!.rows).toBe(model.sheets[0]!.rows);
    expect(snapshot.workbook.opaque).toBe(model.opaque);
    expect(snapshot.workbook.styles).toBe(model.styles);
    // SPEC-124: columns, merges and frozen are copy-on-write, kept by reference.
    expect(snapshot.sheets[0]!.columns).toBe(model.sheets[0]!.columns);
    expect(snapshot.sheets[0]!.merges).toBe(model.sheets[0]!.merges);
    // Small fields are copies, not the live objects.
    expect(snapshot.sheets[0]).not.toBe(model.sheets[0]);
  });

  it("holds SPEC-128's autofilter and sort record by reference, and an undo restores them by identity", () => {
    const model = createBlankWorkbook();
    const autoFilter = { range: { top: 0, left: 0, bottom: 3, right: 1 }, columns: [{ colId: 0, criteria: { kind: "values" as const, values: ["West"], blanks: false, dateGroups: [] } }] };
    const sortState = { ref: { top: 1, left: 3, bottom: 9, right: 4 }, keys: [{ col: 4, descending: true }] };
    model.sheets = [{ ...model.sheets[0]!, autoFilter, sortState }];
    const snapshot = takeSnapshot(model);
    expect(snapshot.sheets[0]!.autoFilter).toBe(autoFilter);
    expect(snapshot.sheets[0]!.sortState).toBe(sortState);
    const { undo } = mount(model);
    undo.record("Clear filters", { sheetId: model.sheets[0]!.id });
    model.sheets = [{ ...model.sheets[0]!, autoFilter: { ...autoFilter, columns: [] } }];
    undo.commit();
    undo.undo();
    expect(model.sheets[0]!.autoFilter).toBe(autoFilter);
    expect(model.sheets[0]!.sortState).toBe(sortState);
  });

  it("names the by-reference fields the spec lists", () => {
    expect(SHEET_BY_REFERENCE_FIELDS).toEqual(["rows", "sourceXml", "columns", "merges", "frozen", "autoFilter", "sortState", "conditionalFormats"]);
    expect(WORKBOOK_SNAPSHOT_FIELDS).toEqual(["activeSheetIndex", "structureDirty", "styles", "opaque", "odsContentEnvelope", "dxfs"]);
  });
});

describe("the stack", () => {
  it("one step per command, labelled; undo restores the sheets and both flags, redo replays", () => {
    const model = book();
    const { undo, activated } = mount(model);
    const before = model.sheets;
    undo.record("Rename sheet", { sheetId: model.sheets[1]!.id });
    renameSheet(model, 1, "Data");
    undo.commit();
    expect(undo.canUndo.value).toBe(true);
    expect(undo.undoLabel.value).toBe("Rename sheet");
    expect(model.structureDirty).toBe(true);
    expect(model.sheets[0]!.rows[0]!.cells[0]!.formula).toBe("Data!A1*2");
    expect(model.sheets[0]!.dirty).toBe(true);

    undo.undo();
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(["Sheet1", "Sheet2"]);
    expect(model.structureDirty).toBeUndefined();
    expect(model.sheets[0]!.dirty).toBeUndefined();
    expect(model.sheets[0]!.rows).toBe(before[0]!.rows);
    expect(activated.at(-1)).toBe(model.sheets[1]!.id);
    expect(undo.canRedo.value).toBe(true);
    expect(undo.redoLabel.value).toBe("Rename sheet");

    undo.redo();
    expect(model.sheets[1]!.name).toBe("Data");
    expect(model.structureDirty).toBe(true);
  });

  it("a command that changed nothing records no step", () => {
    const model = book();
    const { undo } = mount(model);
    undo.record("Tab colour", "workbook");
    setTabColor(model, 0, undefined);
    undo.commit();
    expect(undo.canUndo.value).toBe(false);
  });

  it(`keeps at most ${MAX_UNDO_STEPS} steps, dropping the oldest`, () => {
    const model = book();
    const { undo } = mount(model);
    for (let i = 0; i < MAX_UNDO_STEPS + 5; i += 1) {
      undo.record("Add sheet", "workbook");
      addSheet(model, 0);
      undo.commit();
    }
    let undone = 0;
    while (undo.canUndo.value) {
      undo.undo();
      undone += 1;
    }
    expect(undone).toBe(MAX_UNDO_STEPS);
    expect(model.sheets).toHaveLength(7);
  });
});

describe("the edit pipeline", () => {
  it("recalculates the changed cells plus alsoRecalc, marks the pass's sheets dirty and writes through twice", () => {
    const model = book();
    const { undo, recalls, applied, finishPass } = mount(model);
    undo.record("Delete sheet", { sheetId: model.sheets[1]!.id });
    const result = deleteSheet(model, 1);
    undo.commit(["extra-key"]);
    expect(result.activeIndex).toBe(0);
    expect(model.sheets[0]!.rows[0]!.cells[0]!.formula).toBe("#REF!*2");
    expect(recalls).toHaveLength(1);
    expect(recalls[0]).toContain("extra-key");
    expect(recalls[0]!.some((key) => key.startsWith(model.sheets[0]!.id))).toBe(true);
    expect(applied).toHaveBeenCalledTimes(1);
    // A sliced pass rewrote a value: that sheet is dirty, and it is written again.
    const sheet = model.sheets[0]!;
    model.sheets = [{ ...sheet, dirty: undefined, rows: [{ index: 0, cells: [{ col: 0, value: { error: "#REF!" }, formula: "#REF!*2" }] }] } as Sheet];
    delete model.sheets[0]!.dirty;
    finishPass([sheet.id]);
    expect(model.sheets[0]!.dirty).toBe(true);
    expect(applied).toHaveBeenCalledTimes(2);
  });

  it("hiding or unhiding rows also recalculates every cell calling SUBTOTAL or AGGREGATE (SPEC-127)", () => {
    const model = createBlankWorkbook();
    const sheet = model.sheets[0]!;
    sheet.rows = [
      { index: 0, cells: [{ col: 0, value: 1 }] },
      { index: 1, cells: [{ col: 0, value: 2 }] },
      { index: 2, cells: [{ col: 0, value: 3, formula: "SUBTOTAL(109,A1:A2)" }, { col: 1, value: 3, formula: "SUM(A1:A2)" }] },
      { index: 5, cells: [{ col: 3, value: 3, formula: "1+AGGREGATE(9,5,A1:A2)" }] },
    ];
    const { undo, recalls } = mount(model);
    const subtotal = cellKey(sheet.id, 2, 0);
    const aggregate = cellKey(sheet.id, 5, 3);
    const sum = cellKey(sheet.id, 2, 1);
    undo.record("Hide rows", { sheetId: sheet.id });
    setRowsHidden(model, sheet.id, [{ start: 1, end: 1 }], true);
    undo.commit();
    expect(recalls[0]).toEqual(expect.arrayContaining([subtotal, aggregate]));
    expect(recalls[0]).not.toContain(sum);
    undo.record("Unhide rows", { sheetId: model.sheets[0]!.id });
    setRowsHidden(model, model.sheets[0]!.id, [{ start: 1, end: 1 }], false);
    undo.commit();
    expect(recalls[1]).toEqual(expect.arrayContaining([subtotal, aggregate]));
    // Undo puts the row back as hidden: the same cells again.
    undo.undo();
    expect(recalls[2]).toEqual(expect.arrayContaining([subtotal, aggregate]));
    // An edit that hides nothing leaves them alone.
    undo.record("Edit", { sheetId: model.sheets[0]!.id });
    const now = model.sheets[0]!;
    model.sheets = [{ ...now, rows: now.rows.map((row) => (row.index === 0 ? { ...row, cells: [{ col: 0, value: 7 }] } : row)) }];
    undo.commit();
    expect(recalls[3]).toEqual([cellKey(sheet.id, 0, 0)]);
  });

  it("a step that changes filterHidden (a filter command) recalculates every cell calling SUBTOTAL or AGGREGATE too (SPEC-128)", () => {
    const model = createBlankWorkbook();
    const sheet = model.sheets[0]!;
    sheet.rows = [
      { index: 0, cells: [{ col: 0, value: "Region" }] },
      { index: 1, cells: [{ col: 0, value: 2 }] },
      { index: 2, cells: [{ col: 0, value: 3, formula: "SUBTOTAL(9,A2:A2)" }] },
    ];
    const { undo, recalls } = mount(model);
    const subtotal = cellKey(sheet.id, 2, 0);
    undo.record("Filter", { sheetId: sheet.id });
    model.sheets = [{ ...sheet, rows: sheet.rows.map((row) => (row.index === 1 ? { ...row, filterHidden: true as const } : row)) }];
    undo.commit();
    expect(recalls[0]).toContain(subtotal);
    undo.undo();
    expect(recalls[1]).toContain(subtotal);
  });

  it("undoing everything leaves no sheet dirty and no structure change", () => {
    const model = book();
    const { undo } = mount(model);
    undo.record("Rename sheet", "workbook");
    renameSheet(model, 1, "Data");
    undo.commit();
    undo.record("Add sheet", "workbook");
    addSheet(model, 1);
    undo.commit();
    undo.undo();
    undo.undo();
    expect(model.structureDirty).toBeUndefined();
    expect(model.sheets.some((sheet) => sheet.dirty && sheet.name !== "Sheet2")).toBe(false);
  });
});

describe("shortcuts", () => {
  it("Ctrl+Z undoes and Ctrl+Y redoes, but not from an input", () => {
    const model = book();
    const { undo } = mount(model);
    undo.record("Add sheet", "workbook");
    addSheet(model, 0);
    undo.commit();
    const input = document.createElement("input");
    document.body.append(input);
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "z", ctrlKey: true, bubbles: true }));
    expect(model.sheets).toHaveLength(3);
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "z", ctrlKey: true }));
    expect(model.sheets).toHaveLength(2);
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "y", ctrlKey: true }));
    expect(model.sheets).toHaveLength(3);
    input.remove();
  });
});
