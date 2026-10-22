// The tables' view wiring (SPEC-133 "View wiring", "Creating a table",
// "Auto-expansion and calculated columns", "Header and totals rows",
// "Convert to range"): Insert ▸ Table's Create Table dialog, the design bar's
// state and commands, the totals-cell menu, the Resize Table dialog, the
// convert confirmation, the ODS notice and the typed-entry hooks;
// tableCommands.ts owns every command.
// - Every command is ONE SPEC-121 undo step on the table's sheet (record →
//   command → commit; a refusal changes nothing, so no step is kept):
//   "Create table", "Rename table", "Table options" (header row, totals row,
//   the stripes, first/last column, filter button — and a totals function),
//   "Table style", "Resize table", "Convert to range". Nothing runs while a
//   cell is being edited. A locked table (queryTable/xml) takes the style
//   options only.
// - The dialogs' range fields point the way SPEC-132's Select Data does:
//   "Select cells" collapses the dialog (`picking`) while the grid's
//   selection fills the field; Done brings it back. A range must be on the
//   dialog's sheet.
// - The name field renames the table it was focused on: `beginRename`
//   captures that table's id and `commitRename` renames THAT table — a click
//   on the grid moves the active cell before the field blurs, so the commit
//   never re-reads "the table at the active cell"; a pending rename also
//   commits as soon as the active cell leaves its table. An invalid name
//   toasts its NAME_PROBLEM_MESSAGES text and reverts.
// - The typed-entry hooks (cellCommit.ts; a typed commit only — never a
//   Fill, paste, fill handle, undo or dropdown pick): `prepare` qualifies a
//   formula typed into (or growing) a table, so `=[@Price]*2` is stored
//   `Sales[@Price]*2`; `within` renames a header cell's column INSIDE the
//   entry's own step (the header text and the column name never disagree);
//   `after` records "Expand table" when the entry grows a table (never a
//   locked one), then "Fill table column" when its formula makes a
//   calculated column — each a step of its own after the entry's, so one
//   undo removes only the last.
// - The ODS notice shows the first time in a session (a module-level flag,
//   kept in memory only) that a table goes into an .ods workbook, until
//   dismissed.
import { computed, onScopeDispose, ref, shallowRef, watch, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { DEFAULT_THEME } from "../../lib/core/theme";
import type { GridRange } from "../../lib/spreadsheet/address";
import type { Sheet, SheetTable, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { NAME_PROBLEM_MESSAGES, validateTableName } from "../../lib/spreadsheet/names";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES, type SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { currentRegion } from "../../lib/spreadsheet/selection";
import * as commands from "../../lib/spreadsheet/tableCommands";
import { tableAt } from "../../lib/spreadsheet/tables";
import { OPTION_ROWS, parseTableRange, STYLE_OPTIONS, TABLE_RANGE_MESSAGE, TABLE_UNDO_LABELS, tableRangeText, type TableOption, type TotalsChoice } from "./tableDesign";
import { tableEntryHooks } from "./tableEntryHooks";
import type { CellEditing } from "./useCellEditing";
import type { SheetMetrics } from "./useSheetMetrics";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export { ODS_TABLE_NOTICE, TABLE_UNDO_LABELS, type TableOption, type TotalsChoice } from "./tableDesign";

let odsNoticeShown = false;
/** Specs only: a fresh session for the ODS notice. */
export function forgetOdsTableNotice(): void {
  odsNoticeShown = false;
}

export interface SheetTablesDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeIndex: Ref<number>; activeSheet: Readonly<Ref<Sheet | null>> };
  viewState: Pick<SheetViewState, "syncInto">;
  selection: Pick<SheetSelectionController, "selection">;
  sheetMetrics: Pick<SheetMetrics, "nav">;
  editing: { undo: Pick<SpreadsheetUndo, "record" | "commit">; cell: Pick<CellEditing, "editing" | "setEntryHooks" | "focusCell"> };
  /** The locale typed headers and new column names are read in. */
  locale?: () => SpreadsheetLocale;
  /** Shows a toast — default: useToasts' addToast. */
  toast?: (message: string) => void;
}

/** The open dialog, on the sheet (and table) it was opened for. */
export type TableDialog = { kind: "create"; sheetId: string; range: string; headers: boolean } | { kind: "resize"; sheetId: string; tableId: string; range: string };

/** A command's refusal as text ("" when it went through or did nothing). */
const refusalOf = (result: { ok: boolean; message?: string } | undefined): string => (result && !result.ok ? (result.message ?? "") : "");

export function useSheetTables(deps: SheetTablesDeps) {
  const toast = deps.toast ?? useToasts().addToast;
  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const busy = (): boolean => deps.editing.cell.editing.value;
  const locale = (): SpreadsheetLocale => deps.locale?.() ?? SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE];
  const theme = computed(() => model()?.theme ?? DEFAULT_THEME);
  const sheetOf = (sheetId: string): Sheet | undefined => model()?.sheets.find((sheet) => sheet.id === sheetId);

  /** One command as ONE undo step on `sheetId`; its refusal ("" for none). */
  function run(label: string, sheetId: string, command: (m: SpreadsheetModel) => { ok: boolean; message?: string }): string {
    const m = model();
    if (!m || busy()) {
      return "";
    }
    deps.viewState.syncInto(m);
    deps.editing.undo.record(label, { sheetId });
    const result = command(m);
    deps.editing.undo.commit();
    return refusalOf(result);
  }

  // --- the table at the active cell ----------------------------------------------------------
  const active = computed((): { sheetId: string; table: SheetTable } | null => {
    const sheet = deps.book.activeSheet.value;
    if (sheet?.kind !== "grid") {
      return null;
    }
    const at = deps.selection.selection.value.active;
    const table = tableAt(sheet, at.row, at.col);
    return table ? { sheetId: sheet.id, table } : null;
  });

  /** What the design bar shows (null: the active cell is in no table). */
  const design = computed(() => {
    const found = active.value;
    if (!found) {
      return null;
    }
    const { table } = found;
    const options: Record<TableOption, boolean> = {
      headerRow: table.headerRow,
      totalsRow: table.totalsRow,
      bandedRows: table.style.showRowStripes,
      bandedColumns: table.style.showColumnStripes,
      firstColumn: table.style.showFirstColumn,
      lastColumn: table.style.showLastColumn,
      filterButton: table.autoFilter !== undefined,
    };
    return { id: table.id, name: table.name, styleName: table.style.name, locked: table.locked === true, options };
  });

  /** The design bar's checkboxes, in order (a locked table: the style ones only). */
  const optionRows = computed(() =>
    OPTION_ROWS.map((row) => ({ ...row, checked: design.value?.options[row.option] === true, disabled: !design.value || (design.value.locked && !STYLE_OPTIONS[row.option]) })),
  );

  /** The active cell's totals cell: its column, function and the menu's label (null: not in a totals row). */
  const totalsCell = computed((): { columnIndex: number; fn: string; label: string } | null => {
    const found = active.value;
    const at = deps.selection.selection.value.active;
    if (!found || !found.table.totalsRow || at.row !== found.table.ref.bottom) {
      return null;
    }
    const columnIndex = at.col - found.table.ref.left;
    const fn = found.table.columns[columnIndex]?.totalsFunction ?? "none";
    return { columnIndex, fn, label: fn === "custom" ? "Custom" : commands.TOTALS_FUNCTION_LABELS[fn] };
  });

  /** A design command on the active table — style commands only when it is locked. */
  function onActive(label: string, styleOnly: boolean, command: (m: SpreadsheetModel, sheetId: string, table: SheetTable) => commands.TableResult): void {
    const found = active.value;
    if (!found || (found.table.locked && !styleOnly)) {
      return;
    }
    const refusal = run(label, found.sheetId, (m) => command(m, found.sheetId, found.table));
    if (refusal) {
      toast(refusal);
    }
  }

  /** A design-bar checkbox: "Table options". */
  function setOption(option: TableOption, on: boolean): void {
    const styleKey = STYLE_OPTIONS[option];
    onActive(TABLE_UNDO_LABELS.options, styleKey !== undefined, (m, sheetId, table) => {
      if (styleKey) {
        return commands.setTableOptions(m, sheetId, table.id, { [styleKey]: on });
      }
      if (option === "headerRow") {
        return commands.setHeaderRow(m, sheetId, table.id, on);
      }
      if (option === "totalsRow") {
        return commands.setTotalsRow(m, sheetId, table.id, on);
      }
      return commands.setTableOptions(m, sheetId, table.id, { filterButton: on });
    });
  }

  /** The gallery's pick: "Table style" (undefined: None). */
  const pickStyle = (name: string | undefined): void =>
    onActive(TABLE_UNDO_LABELS.style, true, (m, sheetId, table) => commands.setTableStyle(m, sheetId, table.id, name));

  /** The totals-cell menu: SUBTOTAL(code, T[Col]) in the active totals cell ("Table options"). */
  function setTotalsFunction(fn: TotalsChoice): void {
    const cell = totalsCell.value;
    if (cell) {
      onActive(TABLE_UNDO_LABELS.options, false, (m, sheetId, table) => commands.setTotalsFunction(m, sheetId, table.id, cell.columnIndex, fn));
    }
  }

  // --- the name field ------------------------------------------------------------------------
  /** The rename in progress: the table the field was focused on, by id. */
  const renaming = shallowRef<{ tableId: string; sheetId: string; original: string; text: string } | null>(null);
  const nameText = computed(() => renaming.value?.text ?? design.value?.name ?? "");

  /** The field gained focus: capture the active table. */
  function beginRename(): void {
    const found = active.value;
    if (!renaming.value && found && !found.table.locked) {
      renaming.value = { tableId: found.table.id, sheetId: found.sheetId, original: found.table.name, text: found.table.name };
    }
  }

  function editName(text: string): void {
    beginRename();
    if (renaming.value) {
      renaming.value = { ...renaming.value, text };
    }
  }

  /** Enter or blur: renames the captured table ("Rename table"), or toasts why not and reverts. */
  function commitRename(): void {
    const pending = renaming.value;
    renaming.value = null;
    const m = model();
    if (!pending || !m || pending.text === pending.original || !sheetOf(pending.sheetId)?.tables?.some((table) => table.id === pending.tableId)) {
      return;
    }
    const problem = validateTableName(m, pending.text, pending.tableId);
    const refusal = problem ? NAME_PROBLEM_MESSAGES[problem] : run(TABLE_UNDO_LABELS.rename, pending.sheetId, (current) => commands.renameTable(current, pending.tableId, pending.text));
    if (refusal) {
      toast(refusal);
    }
  }

  const cancelRename = (): void => void (renaming.value = null);

  // The active cell left the table being renamed (a click moves it before the field blurs).
  watch(
    () => active.value?.table.id,
    (id) => {
      if (renaming.value && renaming.value.tableId !== id) {
        commitRename();
      }
    },
  );

  // --- the dialogs ---------------------------------------------------------------------------
  const dialog = shallowRef<TableDialog | null>(null);
  /** The open dialog's inline error ("" for none). */
  const dialogError = ref("");
  /** The dialog is collapsed while the grid's selection fills its range. */
  const picking = ref(false);
  const odsNotice = ref(false);

  function closeDialog(): void {
    dialog.value = null;
    dialogError.value = "";
    picking.value = false;
  }

  /** Insert ▸ Table: the selection when it is more than one cell, else the active cell's current region. */
  function openCreate(): void {
    const sheet = deps.book.activeSheet.value;
    if (sheet?.kind !== "grid" || busy()) {
      return;
    }
    const sel = deps.selection.selection.value;
    const chosen = sel.ranges[sel.activeRangeIndex] ?? sel.ranges[0];
    const range = chosen && (chosen.top !== chosen.bottom || chosen.left !== chosen.right) ? chosen : currentRegion(sel.active, deps.sheetMetrics.nav());
    dialogError.value = "";
    picking.value = false;
    dialog.value = { kind: "create", sheetId: sheet.id, range: tableRangeText(range), headers: commands.detectHeaderRow(sheet, range, model() ?? undefined) };
  }

  /** The design bar's Resize table. */
  function openResize(): void {
    const found = active.value;
    if (!found || found.table.locked || busy()) {
      return;
    }
    dialogError.value = "";
    picking.value = false;
    dialog.value = { kind: "resize", sheetId: found.sheetId, tableId: found.table.id, range: tableRangeText(found.table.ref) };
  }

  /** The dialog's range on its sheet, or the inline error. */
  function dialogRange(text: string): { sheetId: string; range: GridRange } | null {
    const open = dialog.value;
    const sheet = open ? sheetOf(open.sheetId) : undefined;
    const range = open && sheet ? parseTableRange(text, sheet.name) : null;
    if (!open || !range) {
      dialogError.value = TABLE_RANGE_MESSAGE;
      return null;
    }
    return { sheetId: open.sheetId, range };
  }

  /** OK: "Create table" — false (the inline error) when refused. */
  function confirmCreate(text: string, headers: boolean): boolean {
    const target = dialog.value?.kind === "create" ? dialogRange(text) : null;
    if (!target) {
      return false;
    }
    const refusal = run(TABLE_UNDO_LABELS.create, target.sheetId, (m) => commands.createTable(m, target.sheetId, target.range, headers, locale()));
    if (refusal) {
      dialogError.value = refusal;
      return false;
    }
    closeDialog();
    if (model()?.format === "ods" && !odsNoticeShown) {
      odsNoticeShown = true;
      odsNotice.value = true;
    }
    return true;
  }

  /** OK: "Resize table" — false (the inline error) when refused. */
  function confirmResize(text: string): boolean {
    const open = dialog.value;
    const target = open?.kind === "resize" ? dialogRange(text) : null;
    if (open?.kind !== "resize" || !target) {
      return false;
    }
    const refusal = run(TABLE_UNDO_LABELS.resize, target.sheetId, (m) => commands.resizeTable(m, target.sheetId, open.tableId, target.range, locale()));
    if (refusal) {
      dialogError.value = refusal;
      return false;
    }
    closeDialog();
    return true;
  }

  /** The grid selection on show as a range field's text ("" off a grid). */
  const selectionText = computed(() => {
    const sheet = deps.book.activeSheet.value;
    const sel = deps.selection.selection.value;
    const range = sel.ranges[sel.activeRangeIndex] ?? sel.ranges[0];
    return sheet?.kind === "grid" && range ? tableRangeText(range) : "";
  });

  const startPick = (): void => void (picking.value = dialog.value !== null);

  /** Picking done (the dialog took the range): back to the dialog's sheet. */
  function endPick(): void {
    picking.value = false;
    const index = dialog.value ? (model()?.sheets.findIndex((sheet) => sheet.id === dialog.value?.sheetId) ?? -1) : -1;
    if (index >= 0) {
      deps.book.activeIndex.value = index;
    }
  }

  // --- convert to range ----------------------------------------------------------------------
  /** The table the convert confirmation asks about. */
  const converting = shallowRef<{ sheetId: string; tableId: string; name: string } | null>(null);

  function askConvert(): void {
    const found = active.value;
    if (found && !found.table.locked && !busy()) {
      converting.value = { sheetId: found.sheetId, tableId: found.table.id, name: found.table.name };
    }
  }

  function confirmConvert(): void {
    const target = converting.value;
    converting.value = null;
    if (target) {
      const refusal = run(TABLE_UNDO_LABELS.convert, target.sheetId, (m) => commands.convertTableToRange(m, target.sheetId, target.tableId, theme.value));
      if (refusal) {
        toast(refusal);
      }
    }
  }

  // --- the typed-entry hooks (tableEntryHooks.ts) ----------------------------------------------
  const hooks = tableEntryHooks({ model, undo: deps.editing.undo, locale });
  deps.editing.cell.setEntryHooks(hooks);
  onScopeDispose(() => deps.editing.cell.setEntryHooks(null));

  // prettier-ignore
  return {
    active, design, optionRows, totalsCell, theme, nameText, dialog, dialogError, picking, selectionText, converting, odsNotice,
    setOption, pickStyle, setTotalsFunction, beginRename, editName, commitRename, cancelRename, openCreate, openResize, confirmCreate,
    confirmResize, closeDialog, startPick, endPick, askConvert, confirmConvert,
    cancelConvert: (): void => void (converting.value = null),
    dismissOdsNotice: (): void => void (odsNotice.value = false),
    focusGrid: (): void => deps.editing.cell.focusCell(),
  };
}

export type SheetTables = ReturnType<typeof useSheetTables>;

/** Narrows a component's runtime-typed `tables` prop (the .vue parser bars casts). */
export function asSheetTables(value: unknown): SheetTables {
  return value as SheetTables;
}

/** The grid context's tables (null: not wired). */
export function tablesOf(grid: unknown): SheetTables | null {
  const value = (grid as { tables?: unknown } | null)?.tables;
  return value ? asSheetTables(value) : null;
}
