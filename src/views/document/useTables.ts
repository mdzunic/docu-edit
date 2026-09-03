// Table editing (SPEC-022, SPEC-023): insert/delete rows and columns, merge and
// split cells, the table properties dialog, and live column/row resizing with
// the hover handles. The cell-range drag *selection* itself stays in the canvas
// event hub in useCanvasGestures (it shares mousedown/mousemove/mouseup with
// image and TOC deselect); this composable owns the actions those ranges and
// the context menu trigger.
import type { Ref } from "vue";
import {
  applyBorderPreset,
  deleteTable,
  deleteTableColumn,
  deleteTableRow,
  fullTableRange,
  insertTable,
  insertTableColumn,
  insertTableRow,
  mergeCells,
  setCellPadding,
  setCellShading,
  setCellVerticalAlign,
  setColumnWidths,
  setRowMinHeight,
  setTableAlignment,
  setTableWidthMode,
  splitCell,
  tableAt,
  toggleRowRepeat,
} from "../../lib/tables";
import {
  cellRangeBounds,
  MIN_COLUMN_WIDTH_CM,
  MIN_ROW_HEIGHT_CM,
  type CellRange,
} from "../../lib/core/tables";
import {
  findTableElement,
  noColumnResizeDrag,
  noRowResizeDrag,
  type ColumnResizeDrag,
  type RowResizeDrag,
} from "../../lib/tableResize";
import { previewColumnWidths, previewRowHeight, tableRowElement } from "../../lib/core/tables";
import { CM_PER_PX } from "../../lib/core/units";
import { caretSelection, noCommandResult } from "../../lib/commands";
import { withHeaderFooterModel } from "../../lib/commands";
import {
  asBorderLineStyle,
  asCellVerticalAlign,
  asTableAlignment,
  asTableWidthMode,
} from "../../lib/core/tables";
import { asBorderPreset } from "../../lib/core/tables";
import type {
  TableHandleGeometry,
} from "../../lib/caretGeometry";
import type { TableCellTarget } from "../../lib/domCaret";
import type { EditorContext, HeaderFooterBridge } from "./context";
import type { ModelSelection } from "../../lib/commands";
import type { DocumentModel } from "../../lib/documentModel";

/** The cell/table a right-click or dialog targets. */
export interface TableLoc {
  section: number;
  block: number;
  row: number;
  col: number;
}

/** The table properties dialog's prefill (SPEC-023) — every field optional
 * so the view can seed an empty ref before any table was targeted. */
export interface TablePropertiesInitial {
  borderStyle?: string;
  borderWidthPt?: number;
  borderColor?: string;
  shadingFill?: string;
  paddingCm?: number;
  verticalAlign?: string;
  alignment?: string;
  widthMode?: string;
  widthPercent?: number;
  scope?: string;
}

/** Typed empty seed for the view's ref (espree bars annotations there). */
export function noTablePropertiesInitial(): TablePropertiesInitial {
  return {};
}

export interface TablesDeps extends EditorContext, HeaderFooterBridge {
  /** DocumentView's paragraphSelection always resolves to a selection. */
  paragraphSelection: (root?: HTMLElement | null) => ModelSelection;
  hfRootEl: () => HTMLElement | null;
  // --- shared selection / context-menu state (owned by the view) ---------
  closeContextMenu: () => void;
  tableHover: Ref<TableHandleGeometry | null>;
  cellSelection: Ref<CellRange | null>;
  insertTableDialogOpen: Ref<boolean>;
  splitCellDialogOpen: Ref<boolean>;
  tablePropertiesOpen: Ref<boolean>;
  /** "Table -> Repeat row for each..." (SPEC-046 AC-29, AC-30) — opened only
   * when the right-clicked row does not already repeat; an already-repeating
   * row toggles off immediately with no dialog. */
  repeatRowDialogOpen: Ref<boolean>;
  tablePropertiesInitial: Ref<TablePropertiesInitial>;
  getTableFormattingLocation: () => TableLoc;
  setTableFormattingLocation: (location: TableLoc) => void;
  getContextMenuTableCell: () => TableCellTarget | null;
  /** Routes a "delete table" through the caption-delete prompt (SPEC-055
   * AC-20): given the table's block address and a `perform` that does the
   * plain delete, it prompts first when the table has a bound caption, else
   * calls `perform` immediately. Forward-referenced from useCaptions, which is
   * wired after this composable — only ever invoked at user-action time. */
  requestObjectDelete: (
    object: { section: number; block: number },
    objectKind: string,
    perform: () => void,
  ) => void;
}

export function useTables(deps: TablesDeps) {
  const {
    doc,
    contentRoot,
    editable,
    currentSelection,
    paragraphSelection,
    applyEdit,
    addToast,
    activeHeaderFooter,
    currentHfSelection,
    applyHeaderFooterEdit,
    hfRootEl,
    closeContextMenu,
    tableHover,
    cellSelection,
    insertTableDialogOpen,
    splitCellDialogOpen,
    tablePropertiesOpen,
    repeatRowDialogOpen,
    tablePropertiesInitial,
    getTableFormattingLocation,
    setTableFormattingLocation,
    getContextMenuTableCell,
    requestObjectDelete,
    onUnmount,
    zoom,
    clientToContentPx,
  } = deps;

  // Live column/row resize drag state — internal to this composable (nothing
  // else in the view renders from it); cleared on release or unmount.
  let columnResizeDrag: ColumnResizeDrag | null = noColumnResizeDrag();
  let rowResizeDrag: RowResizeDrag | null = noRowResizeDrag();

  /** One handle/menu table action against an explicit target cell. */
  function applyTableAction(action = "", target: TableLoc = { section: 0, block: 0, row: 0, col: 0 }) {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return;
    }
    const location = target;
    const fallbackCaret = {
      section: location.section,
      block: location.block,
      offset: 0,
      cell: { row: location.row, col: location.col, block: 0 },
    };
    const selection = currentSelection(root) ?? caretSelection(fallbackCaret);
    let result = noCommandResult();
    let label = "";
    if (action === "insert-row-above") {
      result = insertTableRow(model, selection, location, "above");
      label = "insert row";
    } else if (action === "insert-row-below") {
      result = insertTableRow(model, selection, location, "below");
      label = "insert row";
    } else if (action === "insert-col-left") {
      result = insertTableColumn(model, selection, location, "left");
      label = "insert column";
    } else if (action === "insert-col-right") {
      result = insertTableColumn(model, selection, location, "right");
      label = "insert column";
    } else if (action === "delete-row") {
      result = deleteTableRow(model, selection, location);
      label = "delete row";
    } else if (action === "delete-col") {
      result = deleteTableColumn(model, selection, location);
      label = "delete column";
    } else if (action === "delete-table") {
      // SPEC-055 AC-20: prompt whether to delete a bound caption too before
      // removing the table (falls through to a plain delete when uncaptioned).
      requestObjectDelete({ section: location.section, block: location.block }, "table", () => {
        const current = doc.value;
        const currentRoot = contentRoot.value as HTMLElement | null;
        if (!current || !currentRoot) {
          return;
        }
        applyEdit(deleteTable(current, selection, location), selection, "delete table");
        currentRoot.focus();
      });
      return;
    }
    if (result) {
      applyEdit(result, selection, label);
      root.focus();
    }
  }

  /** A hover handle's click (AC-20, AC-21): boundary index `at` for inserts
   * (0 = before the first row/column), row/column index for deletes. */
  function handleTableHandle(kind = "", at = 0) {
    const hover = tableHover.value;
    if (!hover) {
      return;
    }
    const base = { section: hover.section, block: hover.block, row: 0, col: 0 };
    if (kind === "insert-row") {
      applyTableAction(at === 0 ? "insert-row-above" : "insert-row-below", {
        ...base,
        row: at === 0 ? 0 : at - 1,
      });
    } else if (kind === "delete-row") {
      applyTableAction("delete-row", { ...base, row: at });
    } else if (kind === "insert-col") {
      applyTableAction(at === 0 ? "insert-col-left" : "insert-col-right", {
        ...base,
        col: at === 0 ? 0 : at - 1,
      });
    } else if (kind === "delete-col") {
      applyTableAction("delete-col", { ...base, col: at });
    }
  }

  // --- column/row resize (SPEC-023 AC-1..5) -------------------------------

  /** Dragging a column boundary handle (AC-1..3): live-resizes the column on
   * either side of the boundary directly in the DOM; the model updates once,
   * on release. */
  function startColumnResize(event: MouseEvent, colIndex = 0) {
    const hover = tableHover.value;
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!hover || !model || !root || !editable.value) {
      return;
    }
    const table = tableAt(model, hover.section, hover.block);
    const tableEl = findTableElement(root, hover.section, hover.block);
    if (!table || !tableEl) {
      return;
    }
    event.preventDefault();
    columnResizeDrag = {
      section: hover.section,
      block: hover.block,
      colIndex,
      startX: event.clientX,
      startWidthsCm: [...table.columnWidthsCm],
      tableEl,
      liveWidthsCm: null,
    };
    window.addEventListener("mousemove", handleColumnResizeMove);
    window.addEventListener("mouseup", finishColumnResize);
  }

  function handleColumnResizeMove(event: MouseEvent) {
    const drag = columnResizeDrag;
    if (!drag) {
      return;
    }
    // The boundary must track the pointer 1:1 at any zoom: the drag delta is
    // in unscaled content px, not raw (scaled) screen px (SPEC-063 AC-24).
    const deltaCm = (clientToContentPx(event.clientX, 0).x - clientToContentPx(drag.startX, 0).x) * CM_PER_PX;
    const widths = [...drag.startWidthsCm];
    const i = drag.colIndex;
    if (i < widths.length - 1) {
      // An internal boundary: the neighbour absorbs exactly what this column
      // gains, so the table's overall width is unchanged (AC-1, AC-2).
      const total = widths[i] + widths[i + 1];
      const next = Math.min(
        Math.max(widths[i] + deltaCm, MIN_COLUMN_WIDTH_CM),
        total - MIN_COLUMN_WIDTH_CM,
      );
      widths[i] = next;
      widths[i + 1] = total - next;
    } else {
      // The table's own right edge: only the last column changes.
      widths[i] = Math.max(MIN_COLUMN_WIDTH_CM, widths[i] + deltaCm);
    }
    drag.liveWidthsCm = widths;
    previewColumnWidths(drag.tableEl, widths);
  }

  /** Release (AC-3): the live DOM widths become the one model update. */
  function finishColumnResize() {
    window.removeEventListener("mousemove", handleColumnResizeMove);
    window.removeEventListener("mouseup", finishColumnResize);
    const drag = columnResizeDrag;
    columnResizeDrag = null;
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!drag || !drag.liveWidthsCm || !root || !model) {
      return;
    }
    const location = { section: drag.section, block: drag.block, row: 0, col: drag.colIndex };
    const fallbackCaret = {
      section: drag.section,
      block: drag.block,
      offset: 0,
      cell: { row: 0, col: drag.colIndex, block: 0 },
    };
    const selection = currentSelection(root) ?? caretSelection(fallbackCaret);
    applyEdit(setColumnWidths(model, selection, location, drag.liveWidthsCm), selection, "resize column");
  }

  /** Dragging a row boundary handle (AC-4, AC-5): live-resizes just that row;
   * the model's minimum height updates once, on release. */
  function startRowResize(event: MouseEvent, rowIndex = 0) {
    const hover = tableHover.value;
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!hover || !model || !root || !editable.value) {
      return;
    }
    const table = tableAt(model, hover.section, hover.block);
    const tableEl = findTableElement(root, hover.section, hover.block);
    const trEl = tableEl ? tableRowElement(tableEl, rowIndex) : null;
    if (!table || !trEl) {
      return;
    }
    event.preventDefault();
    const row = table.rows[rowIndex];
    // The measured row height is scaled by the zoom transform; divide it back
    // out so the model cm is zoom-independent (SPEC-063 AC-22).
    const startHeightCm = row.minHeightCm ?? (trEl.getBoundingClientRect().height / zoom.value) * CM_PER_PX;
    rowResizeDrag = {
      section: hover.section,
      block: hover.block,
      rowIndex,
      startY: event.clientY,
      startHeightCm,
      trEl,
      liveHeightCm: null,
    };
    window.addEventListener("mousemove", handleRowResizeMove);
    window.addEventListener("mouseup", finishRowResize);
  }

  function handleRowResizeMove(event: MouseEvent) {
    const drag = rowResizeDrag;
    if (!drag) {
      return;
    }
    // Unscaled content-px delta so the row tracks the pointer at any zoom
    // (SPEC-063 AC-24).
    const deltaCm = (clientToContentPx(0, event.clientY).y - clientToContentPx(0, drag.startY).y) * CM_PER_PX;
    const heightCm = Math.max(MIN_ROW_HEIGHT_CM, drag.startHeightCm + deltaCm);
    drag.liveHeightCm = heightCm;
    previewRowHeight(drag.trEl, heightCm);
  }

  /** Release (AC-4): the live height becomes the row's minimum. */
  function finishRowResize() {
    window.removeEventListener("mousemove", handleRowResizeMove);
    window.removeEventListener("mouseup", finishRowResize);
    const drag = rowResizeDrag;
    rowResizeDrag = null;
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!drag || drag.liveHeightCm === null || !root || !model) {
      return;
    }
    const location = { section: drag.section, block: drag.block, row: drag.rowIndex, col: 0 };
    const fallbackCaret = {
      section: drag.section,
      block: drag.block,
      offset: 0,
      cell: { row: drag.rowIndex, col: 0, block: 0 },
    };
    const selection = currentSelection(root) ?? caretSelection(fallbackCaret);
    applyEdit(setRowMinHeight(model, selection, location, drag.rowIndex, drag.liveHeightCm), selection, "resize row");
  }

  onUnmount(() => {
    window.removeEventListener("mousemove", handleColumnResizeMove);
    window.removeEventListener("mouseup", finishColumnResize);
    window.removeEventListener("mousemove", handleRowResizeMove);
    window.removeEventListener("mouseup", finishRowResize);
  });

  /** The context menu's table action (AC-17..19), against the right-clicked
   * cell — or, for SPEC-023's formatting actions, routed to their handlers. */
  function handleContextMenuTableAction(action = "") {
    const target = getContextMenuTableCell();
    closeContextMenu();
    if (!target) {
      return;
    }
    if (action === "merge-cells") {
      handleMergeCells();
    } else if (action === "split-cell") {
      openSplitCellDialog(target);
    } else if (action === "table-properties") {
      openTableProperties(target);
    } else if (action === "repeat-row") {
      handleRepeatRowAction(target);
    } else {
      applyTableAction(action, target);
    }
  }

  /** True when the row at `target` already carries an `{{#each}}` chip in
   * one of its cells (SPEC-046 "Repeating table rows" — the chip is the
   * entire repeat state). */
  function rowAlreadyRepeats(model: DocumentModel, target: TableLoc): boolean {
    const table = tableAt(model, target.section, target.block);
    const row = table?.rows[target.row];
    return !!row?.cells.some((cell) =>
      cell.blocks.some((paragraph) => paragraph.runs.some((run) => run.field?.kind === "each")),
    );
  }

  /** "Repeat row for each..." (SPEC-046 AC-29, AC-30): an already-repeating
   * row toggles off immediately (no field to choose for removal); otherwise
   * remembers the target and opens the array-field picker dialog the view
   * renders, whose choice `confirmRepeatRow` below applies. */
  function handleRepeatRowAction(target: TableLoc) {
    const model = doc.value;
    if (!model || !editable.value) {
      return;
    }
    setTableFormattingLocation(target);
    if (rowAlreadyRepeats(model, target)) {
      const root = contentRoot.value as HTMLElement | null;
      const selection = paragraphSelection(root);
      applyEdit(toggleRowRepeat(model, selection, target, ""), selection, "stop repeating row");
      return;
    }
    repeatRowDialogOpen.value = true;
  }

  /** The array-field picker's choice (empty = cancelled). */
  function confirmRepeatRow(name = "") {
    repeatRowDialogOpen.value = false;
    const model = doc.value;
    if (!model || !name) {
      return;
    }
    const target = getTableFormattingLocation();
    const root = contentRoot.value as HTMLElement | null;
    const selection = paragraphSelection(root);
    applyEdit(toggleRowRepeat(model, selection, target, name), selection, "repeat row for each");
  }

  /** "Merge cells" (AC-6, AC-7): the active cell-range selection. */
  function handleMergeCells() {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    const range = cellSelection.value;
    if (!root || !model || !editable.value || !range) {
      return;
    }
    const selection =
      currentSelection(root) ?? caretSelection({ section: range.section, block: range.block, offset: 0 });
    applyEdit(mergeCells(model, selection, range), selection, "merge cells");
    root.focus();
  }

  /** "Split cell…" (AC-8, AC-9): opens the dialog against the right-clicked
   * cell (or the active range's anchor cell, if one covers it). */
  function openSplitCellDialog(target = { section: 0, block: 0, row: 0, col: 0 }) {
    setTableFormattingLocation(target);
    splitCellDialogOpen.value = true;
  }

  /** OK in the split-cell dialog: one combined edit (AC-9, AC-34). */
  function handleSplitCellApply(size = { rows: 1, cols: 1 }) {
    splitCellDialogOpen.value = false;
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return;
    }
    const location = getTableFormattingLocation();
    const fallbackCaret = {
      section: location.section,
      block: location.block,
      offset: 0,
      cell: { row: location.row, col: location.col, block: 0 },
    };
    const selection = currentSelection(root) ?? caretSelection(fallbackCaret);
    applyEdit(splitCell(model, selection, location, size.rows, size.cols), selection, "split cell");
    root.focus();
  }

  /** The table properties dialog (SPEC-023 AC-10..23), prefilled from the
   * right-clicked cell's table. */
  function openTableProperties(target = { section: 0, block: 0, row: 0, col: 0 }) {
    const model = doc.value;
    if (!model) {
      return;
    }
    const table = tableAt(model, target.section, target.block);
    if (!table) {
      return;
    }
    setTableFormattingLocation(target);
    const range = cellSelection.value ?? fullTableRange(model, target);
    const bounds = range
      ? cellRangeBounds(range)
      : { top: target.row, left: target.col, bottom: target.row, right: target.col };
    const cell = table.rows[bounds.top]?.cells[bounds.left];
    const scope =
      cellSelection.value && (bounds.top !== bounds.bottom || bounds.left !== bounds.right)
        ? "the selected cells"
        : "the whole table";
    const firstBorder =
      cell?.borders?.top ?? cell?.borders?.bottom ?? cell?.borders?.left ?? cell?.borders?.right;
    tablePropertiesInitial.value = {
      borderStyle: firstBorder?.style ?? "single",
      borderWidthPt: firstBorder?.widthPt ?? 1,
      borderColor: firstBorder && firstBorder.color !== "auto" ? firstBorder.color : "#000000",
      shadingFill: cell?.shadingFill ?? "",
      paddingCm: cell?.paddingCm ?? 0.19,
      verticalAlign: cell?.verticalAlign ?? "top",
      alignment: table.alignment ?? "left",
      widthMode: table.widthMode ?? "fixed",
      widthPercent: table.widthPercent ?? 100,
      scope,
    };
    tablePropertiesOpen.value = true;
  }

  /** OK in the table properties dialog: every field chained through the model
   * as one combined edit — a single undo step (AC-34). */
  function handleTablePropertiesApply(
    payload = {
      borderPreset: "",
      borderStyle: "single",
      borderWidthPt: 1,
      borderColor: "#000000",
      shadingFill: "",
      paddingCm: 0.19,
      verticalAlign: "top",
      alignment: "left",
      widthMode: "fixed",
      widthPercent: 100,
    },
  ) {
    tablePropertiesOpen.value = false;
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    const location = getTableFormattingLocation();
    if (!root || !model || !editable.value) {
      return;
    }
    const range = cellSelection.value ?? fullTableRange(model, location);
    if (!range) {
      return;
    }
    const fallbackCaret = {
      section: location.section,
      block: location.block,
      offset: 0,
      cell: { row: location.row, col: location.col, block: 0 },
    };
    const selection = currentSelection(root) ?? caretSelection(fallbackCaret);
    let current = model;
    if (payload.borderPreset) {
      current = applyBorderPreset(current, selection, range, asBorderPreset(payload.borderPreset), {
        style: asBorderLineStyle(payload.borderStyle),
        widthPt: payload.borderWidthPt,
        color: payload.borderColor,
      }).model;
    }
    current = setCellShading(current, selection, range, payload.shadingFill || undefined).model;
    current = setCellPadding(current, selection, range, payload.paddingCm).model;
    current = setCellVerticalAlign(current, selection, range, asCellVerticalAlign(payload.verticalAlign)).model;
    current = setTableAlignment(current, selection, location, asTableAlignment(payload.alignment)).model;
    current = setTableWidthMode(
      current,
      selection,
      location,
      asTableWidthMode(payload.widthMode),
      payload.widthPercent,
    ).model;
    applyEdit({ model: current, selection }, selection, "table properties");
    root.focus();
  }

  /** The toolbar's grid picker or the dialog (AC-1..4): inserts at the caret
   * and focuses the first cell. Inside a table the insert is refused. */
  function handleInsertTable(
    size: { rows: number; cols: number; headerRow?: boolean } = { rows: 0, cols: 0 },
  ) {
    insertTableDialogOpen.value = false;
    const active = activeHeaderFooter.value;
    const model = doc.value;
    if (active) {
      // Header/footer tables (SPEC-029 AC-5): run against the band's synthetic
      // model.
      if (!model || !editable.value) {
        return;
      }
      const selection = currentHfSelection() ?? caretSelection({ section: 0, block: 0, offset: 0 });
      if (selection.focus.cell) {
        addToast("Tables can't be inserted inside a table");
        hfRootEl()?.focus();
        return;
      }
      const synthetic = withHeaderFooterModel(model, active.sectionIndex, active.part, active.kind);
      const result = insertTable(synthetic, selection, size.rows, size.cols, size.headerRow);
      applyHeaderFooterEdit(result.model, result.selection, "insert table");
      hfRootEl()?.focus();
      return;
    }
    const root = contentRoot.value as HTMLElement | null;
    if (!root || !model || !editable.value) {
      return;
    }
    const selection = paragraphSelection(root);
    if (selection.focus.cell) {
      addToast("Tables can't be inserted inside a table");
      root.focus();
      return;
    }
    applyEdit(insertTable(model, selection, size.rows, size.cols, size.headerRow), selection, "insert table");
    root.focus();
  }

  return {
    applyTableAction,
    handleTableHandle,
    startColumnResize,
    handleColumnResizeMove,
    finishColumnResize,
    startRowResize,
    handleRowResizeMove,
    finishRowResize,
    handleContextMenuTableAction,
    handleMergeCells,
    openSplitCellDialog,
    handleSplitCellApply,
    openTableProperties,
    handleTablePropertiesApply,
    handleInsertTable,
    confirmRepeatRow,
  };
}
