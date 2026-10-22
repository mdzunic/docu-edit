// The table editing controller (SPEC-091): the presentation view's wiring for
// the Table properties panel and the row/column/merge/split commands. Backed by
// the persisted `presentationTableOpen` preference (SPEC-066) like the other
// three panels. Every command records one undo entry, mutates the selected
// TableFrame in place through the framework-free command layer
// (lib/presentation/tables.ts), marks the slide edited, and commits.
import { computed, ref, watch, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, SlideFill, SlideFrame, Slide, TableFrame } from "../../lib/presentation/model";
import type { BorderPreset, BorderSide, CellRange, CellVerticalAlign } from "../../lib/core/tables";
import { computeOwnerGrid } from "../../lib/core/tables";
import { markSlideEdited } from "../../lib/presentation/text";
import { deleteFrame } from "../../lib/presentation/frames";
import {
  applyCellBorderPreset,
  deleteTableColumn,
  deleteTableRow,
  insertTableColumn,
  insertTableRow,
  mergeCells,
  setCellFill,
  setCellPadding,
  setCellVerticalAlign,
  setTableBandRows,
  setTableHeaderRow,
  splitCell,
} from "../../lib/presentation/tables";
import { usePreferences } from "../../composables/usePreferences";

export interface TablesDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  /** The selected frame, or null (stageEdit.selection.selectedFrame). */
  selectedFrame: ComputedRef<SlideFrame | null>;
  /** Re-select a frame id after a structural change (or null when deleted). */
  selectFrame: (id: string | null) => void;
  record: (label?: string) => void;
  commit: () => void;
}

export function useTables(deps: TablesDeps) {
  const { preferences, setPreference } = usePreferences();

  const open = computed(() => preferences.presentationTableOpen);
  function toggle(): void {
    setPreference("presentationTableOpen", !open.value);
  }

  /** The selected table frame, or null. */
  const table = computed<TableFrame | null>(() => {
    const frame = deps.selectedFrame.value;
    return frame && frame.kind === "table" ? frame : null;
  });
  const hasTable = computed(() => table.value !== null);

  /** The cell the panel's commands anchor to. Clamped to the grid; reset when
   * the selected table changes. */
  const currentCell = ref({ row: 0, col: 0 });
  watch(table, (t) => {
    if (!t) {
      currentCell.value = { row: 0, col: 0 };
      return;
    }
    currentCell.value = {
      row: Math.min(currentCell.value.row, t.rows.length - 1),
      col: Math.min(currentCell.value.col, t.columnWidthsCm.length - 1),
    };
  });
  function setCurrentCell(row: number, col: number): void {
    currentCell.value = { row, col };
  }

  /** The single-cell range the formatting commands act on (expanded to the
   * whole owner cell by the command). */
  function cellRange(): CellRange {
    const { row, col } = currentCell.value;
    return { section: 0, block: 0, anchor: { row, col }, focus: { row, col } };
  }

  /** Runs `fn` against the selected table inside one undo entry, then commits.
   * `fn` returns `false` when the whole frame should be deleted. */
  function edit(label: string, fn: (frame: TableFrame) => boolean | void): void {
    const t = table.value;
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!t || !slide || !model) {
      return;
    }
    deps.record(label);
    const survived = fn(t);
    if (survived === false) {
      const next = deleteFrame(model, deps.currentIndex.value, t.id);
      deps.selectFrame(next);
    } else {
      markSlideEdited(slide, model.format);
    }
    deps.commit();
  }

  // --- row / column (AC-7) ---------------------------------------------------
  const insertRowAbove = () => edit("Insert row", (t) => void insertTableRow(t, currentCell.value.row, "above"));
  const insertRowBelow = () => edit("Insert row", (t) => void insertTableRow(t, currentCell.value.row, "below"));
  const insertColumnLeft = () => edit("Insert column", (t) => void insertTableColumn(t, currentCell.value.col, "left"));
  const insertColumnRight = () => edit("Insert column", (t) => void insertTableColumn(t, currentCell.value.col, "right"));
  const deleteRow = () => edit("Delete row", (t) => deleteTableRow(t, currentCell.value.row));
  const deleteColumn = () => edit("Delete column", (t) => deleteTableColumn(t, currentCell.value.col));
  const deleteTable = () =>
    edit("Delete table", () => false);

  // --- merge / split (AC-9) --------------------------------------------------
  /** Merges the current cell with its right neighbour (a one-click merge that
   * needs no drag-selection). */
  const mergeRight = () =>
    edit("Merge cells", (t) => {
      const { row, col } = currentCell.value;
      mergeCells(t, { section: 0, block: 0, anchor: { row, col }, focus: { row, col: Math.min(col + 1, t.columnWidthsCm.length - 1) } });
    });
  const mergeDown = () =>
    edit("Merge cells", (t) => {
      const { row, col } = currentCell.value;
      mergeCells(t, { section: 0, block: 0, anchor: { row, col }, focus: { row: Math.min(row + 1, t.rows.length - 1), col } });
    });
  /** Merges an explicit range (drag-selected cells). */
  const mergeRange = (range: CellRange) => edit("Merge cells", (t) => mergeCells(t, range));
  /** Splits the current cell back into its spanned rows×columns (un-merge), or
   * into two columns when it is not spanning. */
  const splitCurrent = () =>
    edit("Split cell", (t) => {
      const { row, col } = currentCell.value;
      const owner = computeOwnerGrid(t)[row]?.[col];
      const rows = owner ? owner.rowSpan : 1;
      const cols = owner ? owner.colSpan : 1;
      if (rows > 1 || cols > 1) {
        splitCell(t, row, col, rows, cols);
      } else {
        splitCell(t, row, col, 1, 2);
      }
    });

  // --- per-cell formatting (AC-10) -------------------------------------------
  const setFill = (fill: SlideFill) => edit("Cell fill", (t) => setCellFill(t, cellRange(), fill));
  const clearFill = () => edit("Cell fill", (t) => setCellFill(t, cellRange(), undefined));
  const applyBorder = (preset: BorderPreset, side: BorderSide) =>
    edit("Cell border", (t) => applyCellBorderPreset(t, cellRange(), preset, side));
  const setPadding = (paddingCm: number) => edit("Cell padding", (t) => setCellPadding(t, cellRange(), paddingCm));
  const setValign = (align: CellVerticalAlign) => edit("Cell alignment", (t) => setCellVerticalAlign(t, cellRange(), align));

  // --- table look (AC-10d) ---------------------------------------------------
  const setHeaderRow = (on: boolean) => edit("Header row", (t) => setTableHeaderRow(t, on));
  const setBandRows = (on: boolean) => edit("Banded rows", (t) => setTableBandRows(t, on));

  return {
    open,
    toggle,
    table,
    hasTable,
    currentCell,
    setCurrentCell,
    insertRowAbove,
    insertRowBelow,
    insertColumnLeft,
    insertColumnRight,
    deleteRow,
    deleteColumn,
    deleteTable,
    mergeRight,
    mergeDown,
    mergeRange,
    splitCurrent,
    setFill,
    clearFill,
    applyBorder,
    setPadding,
    setValign,
    setHeaderRow,
    setBandRows,
  };
}

export type TablesController = ReturnType<typeof useTables>;
