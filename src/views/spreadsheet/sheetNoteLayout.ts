// Where SPEC-135's note chrome sits on the grid: a shown note's box and the
// note editor at Excel's default placement, and the hover popover beside its
// cell. Plain geometry in the paper's unscaled px (the grid's content box
// scales it with the zoom), from the cell's rect — a merge's when the cell
// is in one — and the height of the row above it, which useSheetGrid's
// `noteAnchor` measures (noteGridHooks, with the pointer's hover). No Vue,
// so the placement is unit-tested directly.
import { parseA1, type CellRef, type GridRange } from "../../lib/spreadsheet/address";
import type { GridMetrics } from "../../lib/spreadsheet/geometry";
import { mergeAt, type MergeIndex } from "../../lib/spreadsheet/merges";

export interface NoteRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** A cell's rect and the height of the row above it (0 on the first row). */
export interface NoteAnchor {
  cell: NoteRect;
  above: number;
}

export type NoteAnchorOf = (ref: CellRef) => NoteAnchor | null;

/** The popover's width (paper px; the zoom scales it). */
export const NOTE_POPOVER_WIDTH_PX = 240;
/** The gap between a cell and its popover. */
const POPOVER_GAP_PX = 4;

/** Excel's default note placement (the VML anchor `{c+1}, 15, {max(r−1,0)},
 * 10`): 15 px right of the cell's right edge, 10 px below the top of the
 * row above it (of the cell's own row on the first row), at the note's size. */
export function noteBoxPlacement(anchor: NoteAnchor, box: { widthPx: number; heightPx: number }): NoteRect {
  return {
    left: anchor.cell.left + anchor.cell.width + 15,
    top: anchor.cell.top - anchor.above + 10,
    width: box.widthPx,
    height: box.heightPx,
  };
}

/** The popover's top-left: right of the cell, or left of it when it would
 * leave the cell area on the right (and there is room on the left). */
export function notePopoverPlacement(cell: NoteRect, bounds: { width: number }): { left: number; top: number } {
  const right = cell.left + cell.width + POPOVER_GAP_PX;
  const left = cell.left - POPOVER_GAP_PX - NOTE_POPOVER_WIDTH_PX;
  return { left: right + NOTE_POPOVER_WIDTH_PX > bounds.width && left >= 0 ? left : right, top: cell.top };
}

/** SheetNoteLayer's `anchor` prop (useSheetGrid's noteAnchor). */
export function asNoteAnchorOf(value: unknown): NoteAnchorOf {
  return typeof value === "function" ? (value as NoteAnchorOf) : () => null;
}

/** What noteGridHooks reads from the grid on show. */
export interface NoteGridDeps {
  /** The notes (useCellNotes) — absent: none wired. */
  notes: () => { hoverCell(ref: CellRef | null): void; hoverPopover(inside: boolean): void } | undefined;
  /** useSheetOverlays' rectOf (paper px; null under a frozen pane). */
  rectOf: (range: GridRange, key: string) => NoteRect | null;
  metrics: () => Pick<GridMetrics, "rowTop">;
  merges: () => MergeIndex;
}

/** The grid's side of the notes: `noteAnchor` — a cell's rect (its merge's)
 * and the height of the row above it, for SheetNoteLayer — and
 * `notePointer`, the grid element's pointerover/pointerleave: the pointer
 * over a cell (by its element's data-ref), over a popover
 * (`data-note-popover`), or gone. */
export function noteGridHooks(deps: NoteGridDeps) {
  function noteAnchor(ref: CellRef): NoteAnchor | null {
    const merge = mergeAt(deps.merges(), ref.row, ref.col);
    const cell = deps.rectOf(merge ?? { top: ref.row, left: ref.col, bottom: ref.row, right: ref.col }, "note");
    const m = deps.metrics();
    return cell ? { cell, above: ref.row > 0 ? m.rowTop(ref.row) - m.rowTop(ref.row - 1) : 0 } : null;
  }
  const notePointer = {
    over(event: Event): void {
      const notes = deps.notes();
      const target = event.target instanceof Element ? event.target : null;
      if (!notes || !target) {
        return;
      }
      if (target.closest("[data-note-popover]")) {
        notes.hoverPopover(true);
        return;
      }
      const ref = target.closest<HTMLElement>(".sheet-cell[data-ref]")?.dataset.ref;
      notes.hoverCell(ref ? parseA1(ref) : null);
    },
    leave(): void {
      deps.notes()?.hoverCell(null);
    },
  };
  return { noteAnchor, notePointer };
}
