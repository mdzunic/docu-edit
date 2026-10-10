// Rows and columns in the workbook view (SPEC-124): the header context menu
// (SheetHeaderContextMenu), the insert/delete dialog (SheetInsertDeleteDialog),
// the row height and column width dialogs (SheetSizeDialog), View ▸ Freeze
// and the Format menu's sizes — each command ONE undo step through SPEC-121's
// record/commit over lib/spreadsheet/structure.ts and bands.ts, written
// through by the edit pipeline. A refused command records nothing and shows
// the refusal as a toast.
//
// - Right-clicking a row or column header opens the header menu; a header
//   outside the selection is selected first. Shift+F10 or the Menu key opens
//   it instead of SPEC-123's cell menu when the selection is whole rows or
//   columns (useFormatting yields those keys then). Its entries and labels
//   are headerMenu.ts's; N is the selected rows/columns in the band holding
//   the clicked header; delete and hide act on every band of the selection,
//   as one step (bands deleted bottom-up).
// - Ctrl+Alt+= / Cmd+Option+= (sheet.insertCells) and Ctrl+Alt+- /
//   Cmd+Option+- (sheet.deleteCells), bound by KeyboardEvent.code through
//   SHEET_SHORTCUTS: on whole rows or columns they insert as many above/left
//   (the active range's band) or delete the selection's bands; on any other
//   selection they open the insert/delete dialog ("Entire row", "Entire
//   column", over the active range's rows or columns). The keys act only in
//   the grid while no cell is being edited.
// - Labels: "Insert rows|columns", "Delete rows|columns", "Hide rows|columns",
//   "Unhide rows|columns", "Resize rows|columns", "Autofit", "Freeze panes".
// - After a structural edit the live SheetSelection is the model's adjusted
//   one (structure.ts's selection adjuster, adjustRange). Freezing scrolls
//   the grid to its top-left first, so Sheet.frozen counts from row and
//   column 0; "Freeze at current cell" freezes above and left of it.
import { computed, onScopeDispose, ref, shallowRef, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { isMacPlatform } from "../../lib/core/platform";
import { DEFAULT_THEME } from "../../lib/core/theme";
import { findShortcut } from "../../lib/shortcuts";
import type { CellTextMeasurer } from "../../lib/spreadsheet/autofit";
import {
  autofitColumns,
  autofitRows,
  charsToXlsxWidth,
  freezePanes,
  resizeColumns,
  resizeRows,
  setColumnsHidden,
  setRowsHidden,
  type Band,
  type BandAxis,
} from "../../lib/spreadsheet/bands";
import type { FormatResult } from "../../lib/spreadsheet/formatting";
import { xlsxWidthToPx } from "../../lib/spreadsheet/geometry";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { headerSelection } from "../../lib/spreadsheet/selection";
import { adjustSelection, applyStructuralEdit, type StructuralEdit } from "../../lib/spreadsheet/structure";
import { fontMaxDigitWidth } from "../../lib/spreadsheet/styles";
import type { CellEditing } from "./useCellEditing";
import { createCanvasCellMeasurer } from "./cellMeasurer";
import { bandContaining, bandsHoldHidden, headerAt, headerMenuItems, inGrid, selectionAxis, selectionBands, type HeaderMenuCommand } from "./headerMenu";
import { SHEET_SHORTCUTS } from "./sheetShortcuts";
import { initialSize } from "./sheetSizes";
import type { SheetMetrics } from "./useSheetMetrics";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export interface StructureDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeSheet: Readonly<Ref<Sheet | null>> };
  viewState: Pick<SheetViewState, "syncInto" | "update">;
  selection: Pick<SheetSelectionController, "selection" | "set" | "formatContext">;
  sheetMetrics: Pick<SheetMetrics, "nav">;
  editing: { undo: Pick<SpreadsheetUndo, "record" | "commit">; cell: Pick<CellEditing, "editing"> };
  /** Autofit's measurer — default: the canvas one. */
  measurer?: CellTextMeasurer;
  /** Shows a toast — default: useToasts' addToast. */
  toast?: (message: string) => void;
}

/** The header menu's anchor: the band's axis, the clicked index, a point. */
export interface HeaderMenuAnchor {
  axis: BandAxis;
  index: number;
  x: number;
  y: number;
}

const AXIS_WORD = { row: "rows", col: "columns" } as const;

/** The workbook's maximum digit width (the column-width dialog's unit). */
const mdw = (m: SpreadsheetModel): number => fontMaxDigitWidth(m.styles[0]?.font, m.theme ?? DEFAULT_THEME);

export function useStructure(deps: StructureDeps) {
  const toast = deps.toast ?? useToasts().addToast;
  let measurer = deps.measurer;
  const headerMenu = ref<HeaderMenuAnchor | null>(null);
  /** The insert/delete dialog (null: closed). */
  const insertDelete = ref<"insert" | "delete" | null>(null);
  /** The size dialog: its axis, bands and the value it opens with. */
  const sizeDialog = shallowRef<{ axis: BandAxis; bands: Band[]; initial: string; mdw: number } | null>(null);

  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const gridSheet = (): Sheet | null => {
    const sheet = deps.book.activeSheet.value;
    return sheet && sheet.kind === "grid" ? sheet : null;
  };
  /** Runs one command as ONE undo step; refusals toast. */
  function command(label: string, mutate: (m: SpreadsheetModel, sheet: Sheet) => FormatResult): boolean {
    const m = model();
    const sheet = gridSheet();
    if (!m || !sheet || deps.editing.cell.editing.value) {
      return false;
    }
    deps.viewState.syncInto(m);
    deps.editing.undo.record(label, { sheetId: sheet.id });
    const result = mutate(m, m.sheets.find((each) => each.id === sheet.id)!);
    deps.editing.undo.commit();
    if (!result.ok) {
      toast(result.message);
      return false;
    }
    return true;
  }

  /** Structural edits as one step: all made, or none (a refusal restores
   * what the model held). The live selection follows the model's. */
  function structural(label: string, edits: (sheet: Sheet) => StructuralEdit[]): void {
    let made: StructuralEdit[] = [];
    const live = deps.selection.selection.value;
    const done = command(label, (m, sheet) => {
      const before = { sheets: m.sheets, opaque: m.opaque, envelope: m.odsContentEnvelope };
      made = edits(sheet);
      for (const edit of made) {
        const result = applyStructuralEdit(m, edit);
        if (!result.ok) {
          m.sheets = before.sheets;
          m.opaque = before.opaque;
          m.odsContentEnvelope = before.envelope;
          return result;
        }
      }
      return { ok: true };
    });
    if (done && made.length > 0) {
      deps.viewState.update(made[0]!.sheetId, { selection: made.reduce(adjustSelection, live) });
    }
  }

  const kinds = { row: { insert: "insertRows", delete: "deleteRows" }, col: { insert: "insertCols", delete: "deleteCols" } } as const;

  /** Inserts `band`'s size of rows/columns before or after it. */
  function insertBand(axis: BandAxis, band: Band, where: "before" | "after"): void {
    const count = band.end - band.start + 1;
    const at = where === "before" ? band.start : band.end + 1;
    structural(`Insert ${AXIS_WORD[axis]}`, (sheet) => [{ kind: kinds[axis].insert, sheetId: sheet.id, at, count }]);
  }

  /** Deletes `bands`, bottom-up so each index still means what it did. */
  function deleteBands(axis: BandAxis, bands: readonly Band[]): void {
    const ordered = [...bands].toSorted((a, b) => b.start - a.start);
    structural(`Delete ${AXIS_WORD[axis]}`, (sheet) => ordered.map((band) => ({ kind: kinds[axis].delete, sheetId: sheet.id, at: band.start, count: band.end - band.start + 1 })));
  }

  function setHidden(axis: BandAxis, bands: readonly Band[], hidden: boolean): void {
    command(`${hidden ? "Hide" : "Unhide"} ${AXIS_WORD[axis]}`, (m, sheet) => (axis === "row" ? setRowsHidden : setColumnsHidden)(m, sheet.id, bands, hidden));
  }

  /** Resizes `bands` to `px` (a header drag; under 4 px hides). */
  function resizeBands(axis: BandAxis, bands: readonly Band[], px: number): void {
    command(`Resize ${AXIS_WORD[axis]}`, (m, sheet) => (axis === "row" ? resizeRows(m, sheet.id, bands, px * 0.75) : resizeColumns(m, sheet.id, bands, px)));
  }

  function autofitBands(axis: BandAxis, bands: readonly Band[]): void {
    measurer ??= createCanvasCellMeasurer();
    const used = measurer;
    command("Autofit", (m, sheet) => (axis === "row" ? autofitRows : autofitColumns)(m, sheet.id, bands, used, deps.selection.formatContext.value));
  }

  /** Opens the size dialog over `bands` (blank when their sizes differ). */
  function openSize(axis: BandAxis, bands: Band[]): void {
    const m = model();
    const sheet = gridSheet();
    if (m && sheet) {
      sizeDialog.value = { axis, bands, initial: initialSize(sheet, axis, bands, mdw(m)), mdw: mdw(m) };
    }
  }

  /** The size dialog's OK: points for rows, characters for columns (0 hides). */
  function applySize(amount: number): void {
    const dialog = sizeDialog.value;
    sizeDialog.value = null;
    if (!dialog) {
      return;
    }
    if (dialog.axis === "row") {
      command("Resize rows", (m, sheet) => resizeRows(m, sheet.id, dialog.bands, amount));
      return;
    }
    command("Resize columns", (m, sheet) => {
      if (amount === 0) {
        return setColumnsHidden(m, sheet.id, dialog.bands, true);
      }
      const width = charsToXlsxWidth(amount, dialog.mdw);
      return resizeColumns(m, sheet.id, dialog.bands, xlsxWidthToPx(width, dialog.mdw), m.format === "xlsx" ? width : undefined);
    });
  }

  // --- the header menu --------------------------------------------------------------------

  /** Opens the header menu for a header — selecting it first when it lies
   * outside the selection. */
  function openHeaderMenu(axis: BandAxis, index: number, x: number, y: number): void {
    const current = deps.selection.selection.value;
    const inside = selectionAxis(current) === axis && selectionBands(current, axis).some((band) => index >= band.start && index <= band.end);
    if (!inside) {
      deps.selection.set(headerSelection(axis === "row" ? "row" : "column", index, index, deps.sheetMetrics.nav()));
    }
    headerMenu.value = { axis, index, x, y };
  }

  const headerItems = computed(() => {
    const anchor = headerMenu.value;
    const sheet = gridSheet();
    if (!anchor || !sheet) {
      return [];
    }
    const selection = deps.selection.selection.value;
    const bands = selectionBands(selection, anchor.axis);
    return headerMenuItems(anchor.axis, bandContaining(selection, anchor.axis, anchor.index), bands, bandsHoldHidden(sheet, anchor.axis, bands));
  });

  function runHeaderMenu(entry: HeaderMenuCommand): void {
    const anchor = headerMenu.value;
    headerMenu.value = null;
    if (!anchor) {
      return;
    }
    const selection = deps.selection.selection.value;
    const bands = selectionBands(selection, anchor.axis);
    const clicked = bandContaining(selection, anchor.axis, anchor.index);
    switch (entry) {
      case "insert-before":
      case "insert-after":
        insertBand(anchor.axis, clicked, entry === "insert-before" ? "before" : "after");
        break;
      case "delete":
        deleteBands(anchor.axis, bands);
        break;
      case "hide":
      case "unhide":
        setHidden(anchor.axis, bands, entry === "hide");
        break;
      case "size":
        openSize(anchor.axis, bands);
        break;
      case "autofit":
        autofitBands(anchor.axis, bands);
        break;
    }
  }

  // --- the insert/delete shortcuts and dialog -------------------------------------------------

  /** The active range's band along `axis`. */
  function activeBand(axis: BandAxis): Band {
    const selection = deps.selection.selection.value;
    const range = selection.ranges[selection.activeRangeIndex] ?? selection.ranges[0]!;
    return axis === "row" ? { start: range.top, end: range.bottom } : { start: range.left, end: range.right };
  }

  /** sheet.insertCells / sheet.deleteCells (see the module header). */
  function insertOrDelete(mode: "insert" | "delete"): void {
    const axis = selectionAxis(deps.selection.selection.value);
    if (axis === null) {
      insertDelete.value = mode;
    } else if (mode === "insert") {
      insertBand(axis, activeBand(axis), "before");
    } else {
      deleteBands(axis, selectionBands(deps.selection.selection.value, axis));
    }
  }

  /** The dialog's choice: entire rows or columns of the active range. */
  function chooseInsertDelete(axis: BandAxis): void {
    const mode = insertDelete.value;
    insertDelete.value = null;
    if (mode === "insert") {
      insertBand(axis, activeBand(axis), "before");
    } else if (mode === "delete") {
      deleteBands(axis, [activeBand(axis)]);
    }
  }

  // --- freeze ---------------------------------------------------------------------------------

  const freezeState = computed(() => {
    const sheet = gridSheet();
    const active = deps.selection.selection.value.active;
    return {
      frozen: (sheet?.frozen?.rows ?? 0) > 0 || (sheet?.frozen?.cols ?? 0) > 0,
      atA1: active.row === 0 && active.col === 0,
      hasGrid: sheet !== null,
    };
  });

  function freeze(kind: "top-row" | "first-column" | "at-cell" | "unfreeze"): void {
    const active = deps.selection.selection.value.active;
    const [rows, cols] = { "top-row": [1, 0], "first-column": [0, 1], "at-cell": [active.row, active.col], unfreeze: [0, 0] }[kind];
    if (kind === "at-cell" && rows === 0 && cols === 0) {
      return;
    }
    const sheet = gridSheet();
    if (sheet) {
      // Scroll to the top-left first, so the panes count from row and column 0.
      deps.viewState.update(sheet.id, { topLeftCell: undefined });
      document.querySelector('[data-testid="spreadsheet-grid"]')?.scrollTo?.(0, 0);
    }
    command("Freeze panes", (m, s) => freezePanes(m, s.id, rows!, cols!));
  }

  // --- keys and the pointer -------------------------------------------------------------------

  function onKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented || !inGrid(event.target) || deps.editing.cell.editing.value || !gridSheet()) {
      return;
    }
    const axis = selectionAxis(deps.selection.selection.value);
    if (axis !== null && (event.key === "ContextMenu" || (event.key === "F10" && event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey))) {
      event.preventDefault();
      const band = activeBand(axis);
      const header = document.querySelector(`[data-testid="${axis === "row" ? "sheet-row-header" : "sheet-col-header"}"][data-index="${band.start}"]`);
      const box = header?.getBoundingClientRect();
      openHeaderMenu(axis, band.start, box ? box.left + box.width / 2 : 0, box ? box.bottom : 0);
      return;
    }
    const id = findShortcut(event, isMacPlatform(), SHEET_SHORTCUTS)?.id;
    if (id === "sheet.insertCells" || id === "sheet.deleteCells") {
      event.preventDefault();
      insertOrDelete(id === "sheet.insertCells" ? "insert" : "delete");
    }
  }

  function onContextMenu(event: MouseEvent): void {
    const header = headerAt(event.target);
    if (!header || !inGrid(event.target) || !gridSheet()) {
      return;
    }
    event.preventDefault();
    openHeaderMenu(header.axis, header.index, event.clientX, event.clientY);
  }

  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown, true);
    window.addEventListener("contextmenu", onContextMenu);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown, true);
      window.removeEventListener("contextmenu", onContextMenu);
    }
  });

  return {
    headerMenu,
    headerItems,
    openHeaderMenu,
    runHeaderMenu,
    closeHeaderMenu: (): null => (headerMenu.value = null),
    insertDelete,
    insertOrDelete,
    chooseInsertDelete,
    closeInsertDelete: (): null => (insertDelete.value = null),
    sizeDialog,
    openSize,
    applySize,
    closeSize: (): null => (sizeDialog.value = null),
    /** The Format menu's sizes and autofits, over the selection's bands. */
    formatBands(axis: BandAxis, action: "size" | "autofit"): void {
      const bands = selectionBands(deps.selection.selection.value, axis);
      if (action === "size") {
        openSize(axis, bands);
      } else {
        autofitBands(axis, bands);
      }
    },
    resizeBands,
    autofitBands,
    freezeState,
    freeze,
    onKeydown,
  };
}

export type Structure = ReturnType<typeof useStructure>;

/** Narrows a component's runtime-typed `structure` prop (the .vue parser
 * bars casts). */
export function asStructure(value: unknown): Structure {
  return value as Structure;
}
