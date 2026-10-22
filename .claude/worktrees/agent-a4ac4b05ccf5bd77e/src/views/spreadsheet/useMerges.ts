// Merged cells in the workbook view (SPEC-124 "Merges"): the toolbar's split
// button in SPEC-123's band-6 slot (SheetMergeButton — `sheet-merge`,
// pressed when the active cell is merged, its primary action then Unmerge;
// `sheet-merge-menu` with Merge & Centre, Merge Across, Merge Cells and
// Unmerge) and the warning (SheetMergeWarningDialog, `sheet-merge-warning`:
// "Merging keeps only the upper-left value", Merge / Cancel) shown before a
// merge — or the format painter's merges — would clear a value or formula.
// Each command is ONE undo step ("Merge cells", "Unmerge cells") over
// lib/spreadsheet/mergeCells.ts; a refusal records nothing and toasts.
import { computed, shallowRef, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import type { GridRange } from "../../lib/spreadsheet/address";
import type { FormatResult } from "../../lib/spreadsheet/formatting";
import { mergeAt } from "../../lib/spreadsheet/merges";
import { mergeCells, mergeDiscardsValues, unmergeCells, type MergeMode } from "../../lib/spreadsheet/mergeCells";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { CellEditing } from "./useCellEditing";
import type { SheetMetrics } from "./useSheetMetrics";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export interface MergesDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeSheet: Readonly<Ref<Sheet | null>> };
  viewState: Pick<SheetViewState, "syncInto">;
  selection: Pick<SheetSelectionController, "selection">;
  sheetMetrics: Pick<SheetMetrics, "merges">;
  editing: { undo: Pick<SpreadsheetUndo, "record" | "commit">; cell: Pick<CellEditing, "editing"> };
  toast?: (message: string) => void;
}

export function useMerges(deps: MergesDeps) {
  const toast = deps.toast ?? useToasts().addToast;
  /** The pending warning: what Merge does (null: none showing). */
  const warning = shallowRef<{ proceed: () => void } | null>(null);

  const gridSheet = (): Sheet | null => {
    const sheet = deps.book.activeSheet.value;
    return sheet && sheet.kind === "grid" ? sheet : null;
  };

  /** Whether the active cell is merged — the button shows pressed. */
  const merged = computed(() => {
    const active = deps.selection.selection.value.active;
    return gridSheet() !== null && mergeAt(deps.sheetMetrics.merges.value, active.row, active.col) !== undefined;
  });
  const disabled = computed(() => gridSheet() === null || deps.editing.cell.editing.value);

  /** One command as ONE undo step on the sheet on show. */
  function command(label: string, mutate: (m: SpreadsheetModel, sheetId: string, ranges: GridRange[]) => FormatResult): void {
    const m = deps.book.workbook.value;
    const sheet = gridSheet();
    if (!m || !sheet || deps.editing.cell.editing.value) {
      return;
    }
    deps.viewState.syncInto(m);
    deps.editing.undo.record(label, { sheetId: sheet.id });
    const result = mutate(m, sheet.id, deps.selection.selection.value.ranges);
    deps.editing.undo.commit();
    if (!result.ok) {
      toast(result.message);
    }
  }

  /** Runs `proceed` now, or after the warning's Merge when `discards`. */
  function confirmDiscard(discards: boolean, proceed: () => void): void {
    if (discards) {
      warning.value = { proceed };
    } else {
      proceed();
    }
  }

  function merge(mode: MergeMode): void {
    const m = deps.book.workbook.value;
    const sheet = gridSheet();
    if (!m || !sheet || disabled.value) {
      return;
    }
    const discards = mergeDiscardsValues(m, sheet.id, deps.selection.selection.value.ranges, mode);
    confirmDiscard(discards, () => command("Merge cells", (model, sheetId, ranges) => mergeCells(model, sheetId, ranges, mode)));
  }

  function unmerge(): void {
    command("Unmerge cells", (model, sheetId, ranges) => unmergeCells(model, sheetId, ranges));
  }

  return {
    merged,
    disabled,
    merge,
    unmerge,
    /** The split button's main half: Unmerge on a merged cell, else Merge & Centre. */
    primary: (): void => (merged.value ? unmerge() : merge("centre")),
    warning,
    confirmDiscard,
    confirm(): void {
      const pending = warning.value;
      warning.value = null;
      pending?.proceed();
    },
    cancel(): void {
      warning.value = null;
    },
  };
}

export type Merges = ReturnType<typeof useMerges>;

/** Narrows a component's runtime-typed `merges` prop (the .vue parser bars
 * casts). */
export function asMerges(value: unknown): Merges {
  return value as Merges;
}
