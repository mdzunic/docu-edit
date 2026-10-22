// SPEC-133's useSheetTables: Insert ▸ Table's Create Table dialog (the range
// defaulting to the selection or the active cell's current region, headers to
// detectHeaderRow, one "Create table" step, the overlap refusal inline, cancel),
// the once-per-session ODS notice, the design bar following the active cell,
// the name field (the table it was focused on is the one renamed; invalid names
// toast and revert), the options, the totals-cell menu's SUBTOTAL codes, the
// style gallery, the Resize Table dialog, Convert to Range behind its
// confirmation — each one labelled undo step — and the typed-entry hooks it
// installs on the cell editor (header rename inside the entry's step, "Expand
// table" and "Fill table column" as steps of their own after it, never for a
// locked table). The hooks are driven the way cellCommit.ts's typed commit
// drives them: prepare → writeEntry (within inside its step) → after.
import { beforeEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef, triggerRef } from "vue";
import { getCell, parseA1 } from "../../lib/spreadsheet/address";
import { fixedMeasurer } from "../../lib/spreadsheet/autofit";
import type { CellEntry } from "../../lib/spreadsheet/cells";
import type { Cell, CellValue, Row, Sheet, SheetTable, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { NAME_PROBLEM_MESSAGES } from "../../lib/spreadsheet/names";
import { SPREADSHEET_LOCALES } from "../../lib/spreadsheet/numberFormat/locale";
import { workbookFormatContext } from "../../lib/spreadsheet/numberFormat/render";
import { singleCell, type SheetSelection } from "../../lib/spreadsheet/selection";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import type { EntryHooks } from "./cellCommit";
import { typingLabel, writeEntry } from "./cellEntry";
import { TABLE_RANGE_MESSAGE } from "./tableDesign";
import { useSheetMetrics } from "./useSheetMetrics";
import { forgetOdsTableNotice, ODS_TABLE_NOTICE, TABLE_UNDO_LABELS, useSheetTables, type TableOption, type TotalsChoice } from "./useSheetTables";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

type CellSpec = CellValue | Omit<Cell, "col">;

/** Rows from A1 addresses: a value, or a cell without its column. */
function rowsOf(cells: Record<string, CellSpec>): Row[] {
  const byRow = new Map<number, Cell[]>();
  for (const [address, spec] of Object.entries(cells)) {
    const at = parseA1(address)!;
    const cell: Cell = spec !== null && typeof spec === "object" && "value" in spec ? { col: at.col, ...spec } : { col: at.col, value: spec };
    byRow.set(at.row, [...(byRow.get(at.row) ?? []), cell]);
  }
  return [...byRow]
    .toSorted((a, b) => a[0] - b[0])
    .map(([index, rowCells]) => ({ index, cells: rowCells.toSorted((a, b) => a.col - b.col) }));
}

const STYLE = { name: "TableStyleMedium2", showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false };

/** Sales A1:C4 (Item, Price, Qty), Stock I1:J3 (Name, Score), a plain block
 * F1:G3 (Code, Count) and E8 = SUM(Sales[Price]). */
function workbook(options: { format?: "xlsx" | "ods"; locked?: boolean } = {}): SpreadsheetModel {
  const model = createBlankWorkbook();
  model.format = options.format ?? "xlsx";
  const sales: SheetTable = {
    id: "sales",
    name: "Sales",
    ref: { top: 0, left: 0, bottom: 3, right: 2 },
    headerRow: true,
    totalsRow: false,
    columns: [
      { name: "Item", xlsxId: 1 },
      { name: "Price", xlsxId: 2 },
      { name: "Qty", xlsxId: 3 },
    ],
    style: { ...STYLE },
    autoFilter: {},
    ...(options.locked ? { locked: true as const } : {}),
  };
  const stock: SheetTable = {
    id: "stock",
    name: "Stock",
    ref: { top: 0, left: 8, bottom: 2, right: 9 },
    headerRow: true,
    totalsRow: false,
    columns: [
      { name: "Name", xlsxId: 1 },
      { name: "Score", xlsxId: 2 },
    ],
    style: { ...STYLE },
    autoFilter: {},
  };
  const sheet: Sheet = {
    ...model.sheets[0]!,
    id: "s1",
    rows: rowsOf({
      A1: "Item",
      B1: "Price",
      C1: "Qty",
      F1: "Code",
      G1: "Count",
      I1: "Name",
      J1: "Score",
      A2: "Pen",
      B2: 2,
      C2: 10,
      F2: "X1",
      G2: 5,
      I2: "Ann",
      J2: 1,
      A3: "Ink",
      B3: 5,
      C3: 4,
      F3: "X2",
      G3: 6,
      I3: "Bob",
      J3: 2,
      A4: "Pad",
      B4: 3,
      C4: 7,
      E8: { value: 10, formula: "SUM(Sales[Price])" },
    }),
    tables: [sales, stock],
  };
  delete sheet.dirty;
  model.sheets = [sheet];
  return model;
}

function setup(model: SpreadsheetModel = workbook()) {
  const scope = effectScope();
  const wb = shallowRef<SpreadsheetModel | null>(model);
  const activeIndex = ref(0);
  const activeSheet = computed(() => wb.value?.sheets[activeIndex.value] ?? null);
  const current = shallowRef<SheetSelection>(singleCell({ row: 0, col: 0 }));
  const selection = computed(() => current.value);
  /** Selects `address` (A1 or A1:B2); its top-left is the active cell. */
  const select = (address: string): void => {
    const [from, to = from] = address.split(":");
    const a = parseA1(from!)!;
    const b = parseA1(to!)!;
    current.value = { active: { row: a.row, col: a.col }, ranges: [{ top: a.row, left: a.col, bottom: b.row, right: b.col }], activeRangeIndex: 0 };
  };
  const editing = ref(false);
  const toast = vi.fn();
  const focusCell = vi.fn();
  const installed: Array<EntryHooks | null> = [];
  const steps: string[] = [];
  const result = scope.run(() => {
    const undo = useSpreadsheetUndo({ model: wb, recalc: (_changed, done) => done([]), onApplied: () => triggerRef(wb), activateSheet: () => {} });
    undo.onStep((kind, label) => {
      if (kind === "commit") {
        steps.push(label);
      }
    });
    const sheetMetrics = useSheetMetrics(() => activeSheet.value);
    const tables = useSheetTables({
      book: { workbook: wb, activeIndex, activeSheet },
      viewState: { syncInto: () => {} },
      selection: { selection },
      sheetMetrics,
      editing: { undo, cell: { editing: computed(() => editing.value), setEntryHooks: (hooks) => void installed.push(hooks), focusCell } },
      locale: () => SPREADSHEET_LOCALES["en-US"],
      toast,
    });
    return { undo, tables };
  })!;
  const sheet = (): Sheet => wb.value!.sheets[0]!;
  const table = (id: string): SheetTable | undefined => sheet().tables?.find((each) => each.id === id);
  const cell = (address: string): Cell | undefined => {
    const at = parseA1(address)!;
    return getCell(sheet(), at.row, at.col);
  };
  /** A typed entry's commit at `address`, as cellCommit.ts runs it. */
  const type = (address: string, entry: CellEntry, text: string): void => {
    const hooks = installed.at(-1) ?? null;
    const m = wb.value!;
    const at = parseA1(address)!;
    const typed = { model: m, sheetId: sheet().id, at: { row: at.row, col: at.col }, entry };
    const prepared = hooks?.prepare?.(typed) ?? entry;
    const done = { ...typed, entry: prepared };
    const ctx = { model: m, sheetIndex: 0, undo: result.undo, measurer: fixedMeasurer(7, 20), format: workbookFormatContext(m, SPREADSHEET_LOCALES["en-US"]) };
    writeEntry(ctx, typed.at, prepared, text, hooks?.within ? () => hooks.within?.(done) : undefined);
    hooks?.after?.(done);
  };
  return { ...result, model, wb, activeIndex, select, editing, toast, focusCell, installed, steps, sheet, table, cell, type, stop: () => scope.stop() };
}

const flush = async () => {
  await nextTick();
  await nextTick();
};

beforeEach(() => {
  forgetOdsTableNotice();
});

describe("useSheetTables — Insert ▸ Table", () => {
  it("opens on the selection, headers detected; a single cell takes its current region", () => {
    const t = setup();
    t.select("F1:G3");
    t.tables.openCreate();
    expect(t.tables.dialog.value).toEqual({ kind: "create", sheetId: "s1", range: "$F$1:$G$3", headers: true });
    t.tables.closeDialog();
    t.select("G2");
    t.tables.openCreate();
    expect(t.tables.dialog.value).toEqual({ kind: "create", sheetId: "s1", range: "$F$1:$G$3", headers: true });
    t.tables.closeDialog();
    // A first row of numbers is no header row.
    t.select("B2:C4");
    t.tables.openCreate();
    expect(t.tables.dialog.value).toMatchObject({ range: "$B$2:$C$4", headers: false });
    t.stop();
  });

  it("OK creates Table1 — TableStyleMedium2, banded rows, the filter button — as one \"Create table\" step that undo removes", () => {
    const t = setup();
    t.select("G2");
    t.tables.openCreate();
    const dialog = t.tables.dialog.value!;
    expect(dialog.kind).toBe("create");
    expect(t.tables.confirmCreate(dialog.range, dialog.kind === "create" && dialog.headers)).toBe(true);
    expect(t.tables.dialog.value).toBeNull();
    const created = t.sheet().tables!.find((each) => each.name === "Table1");
    expect(created).toMatchObject({
      name: "Table1",
      ref: { top: 0, left: 5, bottom: 2, right: 6 },
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "Code" }, { name: "Count" }],
      style: { name: "TableStyleMedium2", showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
      autoFilter: {},
    });
    expect(t.steps).toEqual([TABLE_UNDO_LABELS.create]);
    expect(TABLE_UNDO_LABELS.create).toBe("Create table");
    expect(t.undo.undoLabel.value).toBe("Create table");
    t.undo.undo();
    expect(t.sheet().tables!.some((each) => each.name === "Table1")).toBe(false);
    expect(t.undo.canUndo.value).toBe(false);
    t.stop();
  });

  it("refuses an overlap with another table inline (\"Tables can't overlap\"), changing nothing", () => {
    const t = setup();
    const before = t.sheet();
    t.select("E5");
    t.tables.openCreate();
    expect(t.tables.confirmCreate("$C$3:$D$6", true)).toBe(false);
    expect(t.tables.dialogError.value).toBe("Tables can't overlap");
    expect(t.tables.dialog.value?.kind).toBe("create");
    expect(t.sheet().tables).toBe(before.tables);
    expect(t.sheet().rows).toBe(before.rows);
    expect(t.undo.canUndo.value).toBe(false);
    expect(t.steps).toEqual([]);
    t.stop();
  });

  it("refuses an unreadable range inline", () => {
    const t = setup();
    t.select("F1:G3");
    t.tables.openCreate();
    expect(t.tables.confirmCreate("Elsewhere!A1:B2", true)).toBe(false);
    expect(t.tables.dialogError.value).toBe(TABLE_RANGE_MESSAGE);
    expect(t.undo.canUndo.value).toBe(false);
    t.stop();
  });

  it("Cancel changes nothing", () => {
    const t = setup();
    const before = t.sheet();
    t.select("F1:G3");
    t.tables.openCreate();
    t.tables.closeDialog();
    expect(t.tables.dialog.value).toBeNull();
    expect(t.sheet()).toBe(before);
    expect(t.undo.canUndo.value).toBe(false);
    t.stop();
  });

  it("nothing opens or runs while a cell is being edited", () => {
    const t = setup();
    t.editing.value = true;
    t.select("F1:G3");
    t.tables.openCreate();
    expect(t.tables.dialog.value).toBeNull();
    t.select("B2");
    t.tables.pickStyle("TableStyleLight1");
    expect(t.table("sales")!.style.name).toBe("TableStyleMedium2");
    expect(t.undo.canUndo.value).toBe(false);
    t.stop();
  });
});

/** Insert ▸ Table on `address` with the dialog's defaults, then OK. */
function create(t: ReturnType<typeof setup>, address: string): void {
  t.select(address);
  t.tables.openCreate();
  const dialog = t.tables.dialog.value!;
  expect(t.tables.confirmCreate(dialog.range, dialog.kind === "create" && dialog.headers)).toBe(true);
}

describe("useSheetTables — the ODS notice", () => {
  it("shows once per session when a table goes into an .ods workbook; forgetOdsTableNotice starts a new session", () => {
    expect(ODS_TABLE_NOTICE).toBe("OpenDocument has no tables — this saves as a named range with the table's formatting applied to its cells");
    const t = setup(workbook({ format: "ods" }));
    expect(t.tables.odsNotice.value).toBe(false);
    create(t, "F1:G3");
    expect(t.tables.odsNotice.value).toBe(true);
    t.tables.dismissOdsNotice();
    expect(t.tables.odsNotice.value).toBe(false);
    create(t, "A12:B14");
    expect(t.tables.odsNotice.value).toBe(false);
    t.stop();
    // Another view in the same session: still shown already.
    const again = setup(workbook({ format: "ods" }));
    create(again, "F1:G3");
    expect(again.tables.odsNotice.value).toBe(false);
    forgetOdsTableNotice();
    create(again, "A12:B14");
    expect(again.tables.odsNotice.value).toBe(true);
    again.stop();
  });

  it("never shows for an .xlsx workbook", () => {
    const t = setup();
    create(t, "F1:G3");
    expect(t.tables.odsNotice.value).toBe(false);
    t.stop();
  });
});

describe("useSheetTables — the design bar", () => {
  it("follows the active cell: the table it is in, none outside every table", () => {
    const t = setup();
    t.select("B2");
    expect(t.tables.design.value).toEqual({
      id: "sales",
      name: "Sales",
      styleName: "TableStyleMedium2",
      locked: false,
      options: { headerRow: true, totalsRow: false, bandedRows: true, bandedColumns: false, firstColumn: false, lastColumn: false, filterButton: true },
    });
    expect(t.tables.nameText.value).toBe("Sales");
    expect(t.tables.optionRows.value.map((row) => [row.testid, row.checked, row.disabled])).toEqual([
      ["sheet-table-option-header-row", true, false],
      ["sheet-table-option-totals-row", false, false],
      ["sheet-table-option-banded-rows", true, false],
      ["sheet-table-option-banded-columns", false, false],
      ["sheet-table-option-first-column", false, false],
      ["sheet-table-option-last-column", false, false],
      ["sheet-table-option-filter-button", true, false],
    ]);
    t.select("J3");
    expect(t.tables.design.value?.name).toBe("Stock");
    t.select("E2");
    expect(t.tables.design.value).toBeNull();
    expect(t.tables.active.value).toBeNull();
    expect(t.tables.nameText.value).toBe("");
    t.stop();
  });

  it("a locked table keeps the style options only", () => {
    const t = setup(workbook({ locked: true }));
    t.select("B2");
    expect(t.tables.design.value?.locked).toBe(true);
    expect(t.tables.optionRows.value.filter((row) => !row.disabled).map((row) => row.option)).toEqual(["bandedRows", "bandedColumns", "firstColumn", "lastColumn"]);
    t.tables.setOption("totalsRow", true);
    expect(t.table("sales")!.totalsRow).toBe(false);
    t.tables.setOption("bandedColumns", true);
    expect(t.table("sales")!.style.showColumnStripes).toBe(true);
    expect(t.steps).toEqual([TABLE_UNDO_LABELS.options]);
    t.stop();
  });
});

describe("useSheetTables — the name field", () => {
  it("renames the table it was focused on even when the active cell moved into another table first", () => {
    const t = setup();
    t.select("B2");
    t.tables.beginRename();
    t.tables.editName("Orders");
    expect(t.tables.nameText.value).toBe("Orders");
    // A click on the grid moves the active cell before the field blurs.
    t.select("I2");
    t.tables.commitRename();
    expect(t.table("sales")!.name).toBe("Orders");
    expect(t.table("stock")!.name).toBe("Stock");
    expect(t.steps).toEqual([TABLE_UNDO_LABELS.rename]);
    t.stop();
  });

  it("a pending rename commits as soon as the active cell leaves its table", async () => {
    const t = setup();
    t.select("B2");
    t.tables.beginRename();
    t.tables.editName("Orders");
    t.select("J3");
    await flush();
    expect(t.table("sales")!.name).toBe("Orders");
    expect(t.table("stock")!.name).toBe("Stock");
    t.tables.commitRename(); // the blur that follows: nothing left to rename
    expect(t.steps).toEqual([TABLE_UNDO_LABELS.rename]);
    t.stop();
  });

  it("an invalid name toasts its NAME_PROBLEM_MESSAGES text, reverts and changes nothing", () => {
    const t = setup();
    t.select("B2");
    for (const [text, problem] of [
      ["Stock", "tableConflict"],
      ["B12", "looksLikeReference"],
      ["two words", "invalidCharacters"],
      ["", "empty"],
      ["_xlnm.Sales", "reserved"],
    ] as const) {
      t.tables.beginRename();
      t.tables.editName(text);
      t.tables.commitRename();
      expect(t.toast).toHaveBeenLastCalledWith(NAME_PROBLEM_MESSAGES[problem]);
      expect(t.tables.nameText.value).toBe("Sales");
    }
    expect(t.toast).toHaveBeenCalledTimes(5);
    expect(t.table("sales")!.name).toBe("Sales");
    expect(t.cell("E8")!.formula).toBe("SUM(Sales[Price])");
    expect(t.undo.canUndo.value).toBe(false);
    t.stop();
  });

  it("a valid rename is one \"Rename table\" step rewriting every formula that used the old name", () => {
    const t = setup();
    t.select("B2");
    t.tables.beginRename();
    t.tables.editName("Orders");
    t.tables.commitRename();
    expect(t.table("sales")!.name).toBe("Orders");
    expect(t.cell("E8")!.formula).toBe("SUM(Orders[Price])");
    expect(t.tables.design.value?.name).toBe("Orders");
    expect(t.steps).toEqual(["Rename table"]);
    t.undo.undo();
    expect(t.table("sales")!.name).toBe("Sales");
    expect(t.cell("E8")!.formula).toBe("SUM(Sales[Price])");
    expect(t.undo.canUndo.value).toBe(false);
    t.stop();
  });

  it("Escape (cancelRename) and an unchanged name change nothing", () => {
    const t = setup();
    t.select("B2");
    t.tables.beginRename();
    t.tables.editName("Orders");
    t.tables.cancelRename();
    expect(t.tables.nameText.value).toBe("Sales");
    t.tables.beginRename();
    t.tables.commitRename();
    expect(t.table("sales")!.name).toBe("Sales");
    expect(t.toast).not.toHaveBeenCalled();
    expect(t.undo.canUndo.value).toBe(false);
    t.stop();
  });
});

describe("useSheetTables — options, totals and style", () => {
  it("toggling each option is one \"Table options\" step", () => {
    const t = setup();
    t.select("B3");
    const options: TableOption[] = ["headerRow", "totalsRow", "bandedRows", "bandedColumns", "firstColumn", "lastColumn", "filterButton"];
    options.forEach((option, i) => {
      const was = t.tables.design.value!.options[option];
      t.tables.setOption(option, !was);
      expect(t.tables.design.value!.options[option]).toBe(!was);
      expect(t.steps).toHaveLength(i + 1);
      expect(t.steps.at(-1)).toBe("Table options");
    });
    const sales = t.table("sales")!;
    expect(sales.headerRow).toBe(false);
    expect(sales.totalsRow).toBe(true);
    expect(sales.style).toEqual({ name: "TableStyleMedium2", showRowStripes: false, showColumnStripes: true, showFirstColumn: true, showLastColumn: true });
    expect(sales.autoFilter).toBeUndefined();
    t.stop();
  });

  it("the totals row writes \"Total\" in the first column and a SUBTOTAL in the last; undo takes it away", () => {
    const t = setup();
    t.select("B2");
    t.tables.setOption("totalsRow", true);
    expect(t.table("sales")).toMatchObject({ totalsRow: true, ref: { top: 0, left: 0, bottom: 4, right: 2 } });
    expect(t.cell("A5")?.value).toBe("Total");
    expect(t.cell("B5")).toBeUndefined();
    expect(t.cell("C5")?.formula).toBe("SUBTOTAL(109,Sales[Qty])");
    expect(t.steps).toEqual(["Table options"]);
    t.undo.undo();
    expect(t.table("sales")).toMatchObject({ totalsRow: false, ref: { bottom: 3 } });
    expect(t.cell("A5")).toBeUndefined();
    expect(t.cell("C5")).toBeUndefined();
    t.stop();
  });

  it("the totals-cell menu writes SUBTOTAL(code,T[Col]) with Excel's 101–110 codes, each a \"Table options\" step", () => {
    const t = setup();
    t.select("B2");
    t.tables.setOption("totalsRow", true);
    expect(t.tables.totalsCell.value).toBeNull();
    t.select("B5");
    expect(t.tables.totalsCell.value).toEqual({ columnIndex: 1, fn: "none", label: "None" });
    const codes: [TotalsChoice, number][] = [
      ["average", 101],
      ["countNums", 102],
      ["count", 103],
      ["max", 104],
      ["min", 105],
      ["stdDev", 107],
      ["sum", 109],
      ["var", 110],
    ];
    for (const [fn, code] of codes) {
      t.tables.setTotalsFunction(fn);
      expect(t.cell("B5")?.formula).toBe(`SUBTOTAL(${code},Sales[Price])`);
      expect(t.tables.totalsCell.value?.fn).toBe(fn);
    }
    t.tables.setTotalsFunction("average");
    expect(t.tables.totalsCell.value?.label).toBe("Average");
    expect(t.steps).toEqual(Array.from({ length: codes.length + 2 }, () => "Table options"));
    t.tables.setTotalsFunction("none");
    expect(t.cell("B5")?.formula).toBeUndefined();
    expect(t.cell("B5")?.value ?? null).toBeNull();
    t.stop();
  });

  it("turning the totals row off clears it and keeps the functions for next time", () => {
    const t = setup();
    t.select("B2");
    t.tables.setOption("totalsRow", true);
    t.select("B5");
    t.tables.setTotalsFunction("average");
    t.tables.setOption("totalsRow", false);
    expect(t.table("sales")).toMatchObject({ totalsRow: false, ref: { bottom: 3 } });
    expect(t.cell("A5")?.value ?? null).toBeNull();
    expect(t.cell("B5")?.formula).toBeUndefined();
    t.select("B2");
    t.tables.setOption("totalsRow", true);
    expect(t.cell("A5")?.value).toBe("Total");
    expect(t.cell("B5")?.formula).toBe("SUBTOTAL(101,Sales[Price])");
    expect(t.cell("C5")?.formula).toBe("SUBTOTAL(109,Sales[Qty])");
    t.stop();
  });

  it("picking a style is one \"Table style\" step; None removes the name", () => {
    const t = setup();
    t.select("C3");
    t.tables.pickStyle("TableStyleLight9");
    expect(t.table("sales")!.style.name).toBe("TableStyleLight9");
    expect(t.tables.design.value?.styleName).toBe("TableStyleLight9");
    expect(t.steps).toEqual(["Table style"]);
    t.tables.pickStyle(undefined);
    expect(t.table("sales")!.style.name).toBeUndefined();
    expect(t.steps).toEqual(["Table style", "Table style"]);
    t.undo.undo();
    expect(t.table("sales")!.style.name).toBe("TableStyleLight9");
    t.stop();
  });
});

describe("useSheetTables — Resize table", () => {
  const RESIZE_MESSAGE = "The header row must stay in the same row and the new range must overlap the table";

  it("opens on the table's range", () => {
    const t = setup();
    t.select("B2");
    t.tables.openResize();
    expect(t.tables.dialog.value).toEqual({ kind: "resize", sheetId: "s1", tableId: "sales", range: "$A$1:$C$4" });
    t.stop();
  });

  it("refuses a range moving the header row or not overlapping the table, changing nothing", () => {
    const t = setup();
    const before = t.sheet();
    t.select("B2");
    t.tables.openResize();
    expect(t.tables.confirmResize("$A$2:$C$6")).toBe(false);
    expect(t.tables.dialogError.value).toBe(RESIZE_MESSAGE);
    expect(t.tables.confirmResize("$E$1:$E$5")).toBe(false);
    expect(t.tables.dialogError.value).toBe(RESIZE_MESSAGE);
    expect(t.tables.dialog.value?.kind).toBe("resize");
    expect(t.sheet().tables).toBe(before.tables);
    expect(t.undo.canUndo.value).toBe(false);
    t.stop();
  });

  it("a valid range is one \"Resize table\" step", () => {
    const t = setup();
    t.select("B2");
    t.tables.openResize();
    expect(t.tables.confirmResize("$A$1:$D$5")).toBe(true);
    expect(t.tables.dialog.value).toBeNull();
    const sales = t.table("sales")!;
    expect(sales.ref).toEqual({ top: 0, left: 0, bottom: 4, right: 3 });
    expect(sales.columns.map((column) => column.name)).toEqual(["Item", "Price", "Qty", "Column1"]);
    expect(t.cell("D1")?.value).toBe("Column1");
    expect(t.steps).toEqual(["Resize table"]);
    t.undo.undo();
    expect(t.table("sales")!.ref).toEqual({ top: 0, left: 0, bottom: 3, right: 2 });
    t.stop();
  });
});

describe("useSheetTables — Convert to range", () => {
  it("asks first: nothing changes until the confirmation, and cancelling keeps the table", () => {
    const t = setup();
    t.select("B2");
    t.tables.askConvert();
    expect(t.tables.converting.value).toEqual({ sheetId: "s1", tableId: "sales", name: "Sales" });
    expect(t.table("sales")).toBeDefined();
    t.tables.cancelConvert();
    expect(t.tables.converting.value).toBeNull();
    t.tables.confirmConvert();
    expect(t.table("sales")).toBeDefined();
    expect(t.undo.canUndo.value).toBe(false);
    t.stop();
  });

  it("confirmed, it is one \"Convert to range\" step removing the table and rewriting its references to A1", () => {
    const t = setup();
    t.select("B2");
    t.tables.askConvert();
    t.tables.confirmConvert();
    expect(t.table("sales")).toBeUndefined();
    expect(t.table("stock")).toBeDefined();
    expect(t.cell("E8")!.formula).toBe("SUM($B$2:$B$4)");
    expect(t.cell("B1")?.value).toBe("Price");
    expect(t.tables.design.value).toBeNull();
    expect(t.steps).toEqual(["Convert to range"]);
    t.undo.undo();
    expect(t.table("sales")?.name).toBe("Sales");
    expect(t.cell("E8")!.formula).toBe("SUM(Sales[Price])");
    t.stop();
  });
});

describe("useSheetTables — the typed-entry hooks", () => {
  it("installs its hooks on the cell editor and removes them when the view goes", () => {
    const t = setup();
    expect(t.installed).toHaveLength(1);
    expect(t.installed[0]).not.toBeNull();
    t.stop();
    expect(t.installed.at(-1)).toBeNull();
  });

  it("an entry directly below the last data row expands the table as its own \"Expand table\" step after the entry's", () => {
    const t = setup();
    t.type("A5", { value: "Eraser" }, "Eraser");
    expect(t.table("sales")!.ref).toEqual({ top: 0, left: 0, bottom: 4, right: 2 });
    expect(t.steps).toEqual([typingLabel("Eraser", { row: 4, col: 0 }), "Expand table"]);
    t.undo.undo();
    expect(t.table("sales")!.ref.bottom).toBe(3);
    expect(t.cell("A5")?.value).toBe("Eraser");
    expect(t.undo.undoLabel.value).toBe('Typing "Eraser" in A5');
    t.stop();
  });

  it("an entry right of the last column adds a column named from its header cell", () => {
    const t = setup();
    t.type("D1", { value: "Total" }, "Total");
    const sales = t.table("sales")!;
    expect(sales.ref).toEqual({ top: 0, left: 0, bottom: 3, right: 3 });
    expect(sales.columns.at(-1)?.name).toBe("Total");
    expect(t.steps.at(-1)).toBe("Expand table");
    t.stop();
  });

  it("a formula committed into an empty column is qualified and fills it (\"Fill table column\"); one undo removes only the fill", () => {
    const t = setup();
    t.type("D1", { value: "Total" }, "Total");
    t.type("D2", { value: null, formula: "[@Price]*[@Qty]" }, "=[@Price]*[@Qty]");
    expect(t.cell("D2")?.formula).toBe("Sales[@Price]*Sales[@Qty]");
    expect(t.cell("D3")?.formula).toBe("Sales[@Price]*Sales[@Qty]");
    expect(t.cell("D4")?.formula).toBe("Sales[@Price]*Sales[@Qty]");
    expect(t.table("sales")!.columns[3]?.calculatedFormula).toBe("Sales[@Price]*Sales[@Qty]");
    expect(t.steps.slice(-2)).toEqual([typingLabel("=[@Price]*[@Qty]", { row: 1, col: 3 }), "Fill table column"]);
    // A new row takes the calculated column.
    t.type("A5", { value: "Eraser" }, "Eraser");
    expect(t.cell("D5")?.formula).toBe("Sales[@Price]*Sales[@Qty]");
    t.undo.undo(); // Expand table
    t.undo.undo(); // the entry
    t.undo.undo(); // Fill table column
    expect(t.cell("D2")?.formula).toBe("Sales[@Price]*Sales[@Qty]");
    expect(t.cell("D3")).toBeUndefined();
    expect(t.table("sales")!.columns[3]?.calculatedFormula).toBeUndefined();
    t.stop();
  });

  it("a header cell's entry renames its column inside the entry's own step", () => {
    const t = setup();
    t.type("B1", { value: "Cost" }, "Cost");
    expect(t.table("sales")!.columns[1]?.name).toBe("Cost");
    expect(t.cell("E8")!.formula).toBe("SUM(Sales[Cost])");
    expect(t.steps).toEqual([typingLabel("Cost", { row: 0, col: 1 })]);
    t.undo.undo();
    expect(t.table("sales")!.columns[1]?.name).toBe("Price");
    expect(t.cell("E8")!.formula).toBe("SUM(Sales[Price])");
    t.stop();
  });

  it("a locked table never expands", () => {
    const t = setup(workbook({ locked: true }));
    t.type("A5", { value: "Eraser" }, "Eraser");
    t.type("D2", { value: 1 }, "1");
    expect(t.table("sales")!.ref).toEqual({ top: 0, left: 0, bottom: 3, right: 2 });
    expect(t.steps).toEqual([typingLabel("Eraser", { row: 4, col: 0 }), typingLabel("1", { row: 1, col: 3 })]);
    t.stop();
  });

  it("an empty entry or one away from every table runs no table step", () => {
    const t = setup();
    t.type("A5", { value: null }, "");
    t.type("A6", { value: "Far" }, "Far");
    t.type("E2", { value: 3 }, "3");
    expect(t.steps).toEqual([typingLabel("Far", { row: 5, col: 0 }), typingLabel("3", { row: 1, col: 4 })]);
    expect(t.table("sales")!.ref).toEqual({ top: 0, left: 0, bottom: 3, right: 2 });
    t.stop();
  });
});
